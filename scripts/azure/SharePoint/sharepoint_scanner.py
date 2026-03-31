import sys
import json
import csv
import os
import time
import traceback
import logging
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass, asdict
from enum import Enum
from typing import Optional, Dict, List, Tuple
from urllib.parse import urlparse

import requests
import msal

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QFormLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTextEdit,
    QProgressBar,
    QCheckBox,
    QFileDialog,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QGroupBox,
    QStatusBar,
    QComboBox,
    QHeaderView,
    QFrame,
    QDialog,
)
from PySide6.QtCore import QThread, Signal
from PySide6.QtGui import QFont, QColor


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = [
    "User.Read",
    "Sites.Read.All",
    "Files.Read.All",
]

MAX_RETRIES = 4
INITIAL_BACKOFF = 1.5
BACKOFF_FACTOR = 2.0
REQUEST_TIMEOUT = 30
DEFAULT_PAGE_SIZE = 200


class ScanPhase(Enum):
    IDLE = "Idle"
    AUTHENTICATING = "Authenticating..."
    RESOLVING_SITE = "Resolving site..."
    RESOLVING_LIBRARY = "Resolving library..."
    COUNTING = "Counting items..."
    ANALYZING = "Analyzing permissions..."
    EXPORTING = "Exporting results..."
    COMPLETED = "Completed"
    CANCELLED = "Cancelled"
    ERROR = "Error"


class PermissionSignal(Enum):
    INHERITED_OR_UNKNOWN = "InheritedOrUnknown"
    EXPLICIT_GRANTS = "ExplicitGrantsDetected"
    SHARED_VIA_LINK = "SharedViaLink"
    UNKNOWN = "Unknown"


@dataclass
class FlaggedItem:
    site_url: str
    library_title: str
    top_level_folder: str
    item_type: str
    name: str
    web_url: str
    drive_item_id: str
    has_sharing_link: bool = False
    has_anonymous_link: bool = False
    has_guest_or_external: bool = False
    shared_with: str = ""
    access_details: str = ""
    permission_signal: str = PermissionSignal.UNKNOWN.value
    notes: str = ""

    def to_dict(self):
        return asdict(self)


@dataclass
class FolderCount:
    site_url: str
    library_title: str
    top_level_folder: str
    item_count: int


class GuiLogHandler(logging.Handler):
    def __init__(self, emit_fn):
        super().__init__()
        self.emit_fn = emit_fn

    def emit(self, record):
        try:
            self.emit_fn(self.format(record))
        except Exception:
            pass


def build_logger(gui_emit=None):
    logger = logging.getLogger("sp_preflight")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")

    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG)
    console.setFormatter(formatter)
    logger.addHandler(console)

    if gui_emit:
        gui = GuiLogHandler(gui_emit)
        gui.setLevel(logging.DEBUG)
        gui.setFormatter(formatter)
        logger.addHandler(gui)

    return logger


logger = build_logger()


def build_impact_summary(flagged_items: List[FlaggedItem]) -> Dict[str, int]:
    files_affected = 0
    folders_affected = 0
    unique_permissions = 0
    sharing_exposure = 0
    anonymous_sharing = 0
    external_guest = 0

    for item in flagged_items:
        if item.item_type.lower() == "file":
            files_affected += 1
        elif item.item_type.lower() == "folder":
            folders_affected += 1

        if item.permission_signal in (
            PermissionSignal.EXPLICIT_GRANTS.value,
            PermissionSignal.SHARED_VIA_LINK.value,
        ):
            unique_permissions += 1

        if item.has_sharing_link or item.has_guest_or_external or item.has_anonymous_link:
            sharing_exposure += 1

        if item.has_anonymous_link:
            anonymous_sharing += 1

        if item.has_guest_or_external:
            external_guest += 1

    return {
        "affected_files": files_affected,
        "affected_folders": folders_affected,
        "unique_permissions": unique_permissions,
        "sharing_exposure": sharing_exposure,
        "anonymous_sharing": anonymous_sharing,
        "external_guest": external_guest,
        "total_flagged": len(flagged_items),
    }


class AuthManager:
    def __init__(self, client_id: str, tenant_id: str):
        self.client_id = client_id.strip()
        self.tenant_id = tenant_id.strip()
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.app: Optional[msal.PublicClientApplication] = None
        self.token_result: Optional[Dict] = None
        self.username: Optional[str] = None
        self.token_expires_at: int = 0

    def initialize(self):
        self.app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
        )

    def _get_cached_accounts(self):
        if not self.app:
            self.initialize()
        return self.app.get_accounts()

    def refresh_access_token(self, force: bool = False) -> bool:
        if not self.app:
            self.initialize()

        accounts = self._get_cached_accounts()
        if not accounts:
            logger.warning("AUTH   | REFRESH | No cached accounts available for token refresh")
            return False

        for account in accounts:
            try:
                result = self.app.acquire_token_silent(
                    GRAPH_SCOPES,
                    account=account,
                    force_refresh=force,
                )
                if result and "access_token" in result:
                    self._store_result(result)
                    logger.info("AUTH   | REFRESH | Access token refreshed silently")
                    return True
            except Exception as e:
                logger.debug(f"AUTH   | REFRESH | Silent refresh failed for cached account: {e}")

        return False

    def authenticate(self) -> bool:
        if not self.client_id or not self.tenant_id:
            raise RuntimeError("Client ID and Tenant ID are required.")

        if not self.app:
            self.initialize()

        if self.refresh_access_token(force=False):
            return True

        result = self.app.acquire_token_interactive(
            scopes=GRAPH_SCOPES,
            prompt="select_account",
        )

        if result and "access_token" in result:
            self._store_result(result)
            return True

        if isinstance(result, dict):
            err = result.get("error_description") or result.get("error") or "Authentication failed."
            raise RuntimeError(err)

        raise RuntimeError("Authentication failed.")

    def _store_result(self, result: Dict):
        self.token_result = result
        account = result.get("account") or {}
        self.username = account.get("username")
        self.token_expires_at = int(result.get("expires_on", 0) or 0)

        if not self.username:
            claims = result.get("id_token_claims") or {}
            self.username = claims.get("preferred_username") or claims.get("upn") or "Authenticated user"

    def get_access_token(self) -> str:
        now = int(time.time())

        if self.token_result and "access_token" in self.token_result:
            # Refresh a little before expiry to avoid failures mid-scan
            if self.token_expires_at and now >= (self.token_expires_at - 120):
                logger.info("AUTH   | REFRESH | Token nearing expiry, refreshing silently")
                if self.refresh_access_token(force=True):
                    return self.token_result["access_token"]
            else:
                return self.token_result["access_token"]

        if self.refresh_access_token(force=False):
            return self.token_result["access_token"]

        self.authenticate()
        if self.token_result and "access_token" in self.token_result:
            return self.token_result["access_token"]

        raise RuntimeError("Unable to acquire access token.")

    def invalidate_token(self):
        self.token_result = None
        self.token_expires_at = 0

    def is_authenticated(self) -> bool:
        return bool(self.token_result and self.token_result.get("access_token"))


class GraphClient:
    def __init__(self, auth: AuthManager):
        self.auth = auth
        self.session = requests.Session()

    def _headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self.auth.get_access_token()}",
            "Content-Type": "application/json",
        }

    def request(self, method: str, url: str, **kwargs) -> Dict:
        retry = 0
        while retry < MAX_RETRIES:
            try:
                kwargs["headers"] = self._headers()
                kwargs.setdefault("timeout", REQUEST_TIMEOUT)

                response = self.session.request(method, url, **kwargs)

                if response.status_code in (200, 201):
                    return response.json() if response.content else {}

                if response.status_code == 401:
                    try:
                        data = response.json()
                        msg = data.get("error", {}).get("message", response.text)
                    except Exception:
                        msg = response.text

                    logger.warning(f"GRAPH  | REFRESH | HTTP 401 | {msg}")
                    self.auth.invalidate_token()
                    if self.auth.refresh_access_token(force=True):
                        retry += 1
                        continue

                    raise RuntimeError(
                        "Graph authentication expired and silent refresh failed. Please reconnect from the setup screen."
                    )

                if response.status_code in (429, 500, 502, 503, 504):
                    wait = INITIAL_BACKOFF * (BACKOFF_FACTOR ** retry)
                    logger.warning(f"GRAPH  | RETRY   | HTTP {response.status_code} | Waiting {wait:.1f}s")
                    time.sleep(wait)
                    retry += 1
                    continue

                try:
                    data = response.json()
                    msg = data.get("error", {}).get("message", response.text)
                except Exception:
                    msg = response.text

                raise RuntimeError(f"Graph error {response.status_code}: {msg}")

            except requests.RequestException as e:
                if retry < MAX_RETRIES - 1:
                    wait = INITIAL_BACKOFF * (BACKOFF_FACTOR ** retry)
                    logger.warning(f"GRAPH  | RETRY   | Network error {e} | Waiting {wait:.1f}s")
                    time.sleep(wait)
                    retry += 1
                    continue
                raise RuntimeError(f"Request failed: {e}")

        raise RuntimeError("Max retries exceeded.")

    def get_me(self) -> Dict:
        return self.request("GET", f"{GRAPH_BASE_URL}/me")

    def resolve_site(self, site_url: str) -> Dict:
        parsed = urlparse(site_url)
        if parsed.scheme != "https":
            raise RuntimeError("Site URL must start with https://")
        if "sharepoint.com" not in parsed.netloc.lower():
            raise RuntimeError("Site URL must be a SharePoint Online URL.")

        hostname = parsed.netloc
        path = parsed.path.strip("/")

        if path:
            url = f"{GRAPH_BASE_URL}/sites/{hostname}:/{path}"
        else:
            url = f"{GRAPH_BASE_URL}/sites/{hostname}"

        logger.debug(f"SITE   | RESOLVE | {url}")
        return self.request("GET", url)

    def list_drives(self, site_id: str) -> List[Dict]:
        result = self.request("GET", f"{GRAPH_BASE_URL}/sites/{site_id}/drives")
        return result.get("value", [])

    def list_document_libraries(self, site_id: str) -> List[Dict]:
        libraries = []
        for drive in self.list_drives(site_id):
            drive_type = (drive.get("driveType") or "").lower()
            if drive_type in ("documentlibrary", "business", ""):
                libraries.append(drive)
        return libraries

    def list_drive_items(self, drive_id: str, folder_id: str = "root", top: int = DEFAULT_PAGE_SIZE) -> Dict:
        if folder_id == "root":
            url = f"{GRAPH_BASE_URL}/drives/{drive_id}/root/children"
        else:
            url = f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{folder_id}/children"

        params = {
            "$top": top,
            "$select": "id,name,webUrl,file,folder,parentReference",
        }
        return self.request("GET", url, params=params)

    def get_item_permissions(self, drive_id: str, item_id: str) -> List[Dict]:
        url = f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}/permissions"
        try:
            result = self.request("GET", url)
            return result.get("value", [])
        except Exception as e:
            logger.debug(f"PERM   | SKIP    | {item_id} | {e}")
            return []


class SharePointResolver:
    def __init__(self, graph: GraphClient):
        self.graph = graph

    def list_libraries_for_site(self, site_url: str) -> Tuple[str, List[Dict]]:
        site = self.graph.resolve_site(site_url)
        site_id = site.get("id")
        if not site_id:
            raise RuntimeError("Could not resolve site ID.")
        libraries = self.graph.list_document_libraries(site_id)
        return site_id, libraries


class LibraryScanner:
    def __init__(self, graph: GraphClient, page_size: int = DEFAULT_PAGE_SIZE):
        self.graph = graph
        self.page_size = page_size
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def pre_count(self, drive_id: str, progress_callback=None, log_callback=None) -> Tuple[int, Dict[str, int]]:
        total_items = 0
        folder_counts: Dict[str, int] = {"(root)": 0}

        def walk(folder_id: str, depth: int, top_level_folder: str, next_link: Optional[str] = None):
            nonlocal total_items
            if self._cancel:
                return

            response = self.graph.request("GET", next_link) if next_link else self.graph.list_drive_items(
                drive_id, folder_id, self.page_size
            )

            items = response.get("value", [])

            for item in items:
                if self._cancel:
                    return

                item_id = item.get("id", "")
                name = item.get("name", "")
                is_folder = "folder" in item and item.get("folder") is not None
                item_type = "FOLDER" if is_folder else "FILE"

                current_top = top_level_folder
                if depth == 0 and is_folder:
                    current_top = name
                    folder_counts.setdefault(current_top, 0)

                folder_counts.setdefault(current_top, 0)
                folder_counts[current_top] += 1
                total_items += 1

                if log_callback:
                    log_callback(f"{item_type:<6} | COUNTED | {name} | TopFolder={current_top}")

                if progress_callback and total_items % 25 == 0:
                    progress_callback(ScanPhase.COUNTING.value, total_items, 0)

                if is_folder:
                    walk(item_id, depth + 1, current_top)

            next_page = response.get("@odata.nextLink")
            if next_page and not self._cancel:
                walk(folder_id, depth, top_level_folder, next_page)

        walk("root", 0, "(root)")
        return total_items, folder_counts

    def analyze(
        self,
        site_url: str,
        library_title: str,
        drive_id: str,
        total_items: int,
        folder_counts: Dict[str, int],
        scan_sharing: bool,
        scan_permissions: bool,
        progress_callback=None,
        log_callback=None,
        flagged_callback=None,
    ) -> Tuple[List[FlaggedItem], int]:
        flagged_items: List[FlaggedItem] = []
        processed = 0

        def walk(folder_id: str, depth: int, top_level_folder: str, next_link: Optional[str] = None):
            nonlocal processed
            if self._cancel:
                return

            response = self.graph.request("GET", next_link) if next_link else self.graph.list_drive_items(
                drive_id, folder_id, self.page_size
            )

            items = response.get("value", [])

            for item in items:
                if self._cancel:
                    return

                item_id = item.get("id", "")
                name = item.get("name", "")
                web_url = item.get("webUrl", "")
                is_folder = "folder" in item and item.get("folder") is not None
                item_type = "Folder" if is_folder else "File"
                log_type = "FOLDER" if is_folder else "FILE"

                current_top = top_level_folder
                if depth == 0 and is_folder:
                    current_top = name

                folder_counts.setdefault(current_top, 0)
                folder_counts[current_top] += 1

                if log_callback:
                    log_callback(f"{log_type:<6} | CHECK   | {name} | TopFolder={current_top}")

                flags = self._analyze_item(
                    drive_id=drive_id,
                    item_id=item_id,
                    scan_sharing=scan_sharing,
                    scan_permissions=scan_permissions,
                )

                processed += 1
                if progress_callback:
                    progress_callback(ScanPhase.ANALYZING.value, processed, total_items)

                if flags["flagged"]:
                    flagged_item = FlaggedItem(
                        site_url=site_url,
                        library_title=library_title,
                        top_level_folder=current_top,
                        item_type=item_type,
                        name=name,
                        web_url=web_url,
                        drive_item_id=item_id,
                        has_sharing_link=flags["has_sharing_link"],
                        has_anonymous_link=flags["has_anonymous_link"],
                        has_guest_or_external=flags["has_guest_or_external"],
                        shared_with=flags["shared_with"],
                        access_details=flags["access_details"],
                        permission_signal=flags["permission_signal"],
                        notes=flags["notes"],
                    )
                    flagged_items.append(flagged_item)

                    if log_callback:
                        log_callback(
                            f"{log_type:<6} | FLAGGED | {name} | "
                            f"Signal={flags['permission_signal']} | "
                            f"Link={flags['has_sharing_link']} | "
                            f"Anonymous={flags['has_anonymous_link']} | "
                            f"External={flags['has_guest_or_external']}"
                        )

                    if flagged_callback:
                        flagged_callback(flagged_item)
                else:
                    if log_callback:
                        log_callback(f"{log_type:<6} | CLEAR   | {name}")

                if is_folder:
                    walk(item_id, depth + 1, current_top)

            next_page = response.get("@odata.nextLink")
            if next_page and not self._cancel:
                walk(folder_id, depth, top_level_folder, next_page)

        walk("root", 0, "(root)")
        return flagged_items, processed

    def _extract_shared_with_identities(self, perm: Dict) -> List[str]:
        identities = []

        def add_identity(name: str):
            if name and name not in identities:
                identities.append(name)

        def consume_user_obj(user_obj: Dict):
            if not isinstance(user_obj, dict):
                return
            email = (user_obj.get("email") or user_obj.get("userPrincipalName") or "").strip()
            display = (user_obj.get("displayName") or "").strip()
            email_lower = email.lower()
            is_external = "#ext#" in email_lower or "guest" in email_lower
            if is_external:
                if email and display and display.lower() != email.lower():
                    add_identity(f"{display} <{email}>")
                elif email:
                    add_identity(email)
                elif display:
                    add_identity(display)
            else:
                if display:
                    add_identity(display)
                elif email:
                    add_identity(email)

        def consume_identity(identity: Dict):
            if not isinstance(identity, dict):
                return
            if "user" in identity:
                consume_user_obj(identity.get("user") or {})
            if "siteUser" in identity:
                consume_user_obj(identity.get("siteUser") or {})
            if "group" in identity and isinstance(identity.get("group"), dict):
                grp = identity["group"]
                display = (grp.get("displayName") or grp.get("email") or grp.get("id") or "").strip()
                if display:
                    add_identity(display)
            if "siteGroup" in identity and isinstance(identity.get("siteGroup"), dict):
                grp = identity["siteGroup"]
                display = (grp.get("displayName") or grp.get("id") or "").strip()
                if display:
                    add_identity(display)
            # fallback direct fields
            email = (identity.get("email") or identity.get("userPrincipalName") or "").strip()
            display = (identity.get("displayName") or "").strip()
            email_lower = email.lower()
            is_external = "#ext#" in email_lower or "guest" in email_lower
            if is_external:
                if email and display and display.lower() != email.lower():
                    add_identity(f"{display} <{email}>")
                elif email:
                    add_identity(email)
                elif display:
                    add_identity(display)
            else:
                if display:
                    add_identity(display)
                elif email:
                    add_identity(email)

        # Common Graph shapes
        for key in ("grantedTo", "grantedToV2", "invitation"):
            val = perm.get(key)
            if isinstance(val, dict):
                consume_identity(val)

        for list_key in ("grantedToIdentities", "grantedToIdentitiesV2", "grantedToV2List"):
            vals = perm.get(list_key)
            if isinstance(vals, list):
                for val in vals:
                    if isinstance(val, dict):
                        consume_identity(val)

        # Link recipients can show up under link.webHtml or recipients in some APIs, but permissions usually don't include them.
        return identities

    def _analyze_item(self, drive_id: str, item_id: str, scan_sharing: bool, scan_permissions: bool) -> Dict:
        result = {
            "flagged": False,
            "has_sharing_link": False,
            "has_anonymous_link": False,
            "has_guest_or_external": False,
            "shared_with": "",
            "access_details": "",
            "permission_signal": PermissionSignal.UNKNOWN.value,
            "notes": "",
        }

        permissions = []
        if scan_sharing or scan_permissions:
            permissions = self.graph.get_item_permissions(drive_id, item_id)

        if scan_sharing and permissions:
            shared_with_values = []

            for perm in permissions:
                link = perm.get("link")
                if link:
                    result["has_sharing_link"] = True
                    result["permission_signal"] = PermissionSignal.SHARED_VIA_LINK.value
                    result["flagged"] = True

                    scope = (link.get("scope") or "").lower()
                    if scope in ("anonymous", "everyone"):
                        result["has_anonymous_link"] = True

                identities = self._extract_shared_with_identities(perm)
                for identity in identities:
                    if identity not in shared_with_values:
                        shared_with_values.append(identity)

                    low = identity.lower()
                    if "#ext#" in low or "guest" in low:
                        result["has_guest_or_external"] = True
                        result["flagged"] = True

                # Heuristic: if explicit user/group principals exist, flag as shared
                if identities:
                    result["flagged"] = True

            if shared_with_values:
                result["shared_with"] = "; ".join(shared_with_values)
                result["access_details"] = "; ".join(shared_with_values)

        if scan_permissions and permissions:
            explicit = any(p.get("grantedTo") or p.get("grantedToV2") or p.get("link") for p in permissions)
            if explicit:
                if result["permission_signal"] == PermissionSignal.UNKNOWN.value:
                    result["permission_signal"] = PermissionSignal.EXPLICIT_GRANTS.value
                result["flagged"] = True
                summary_parts = []
                if result.get("shared_with"):
                    shared_entries = [s.strip() for s in result["shared_with"].split(";") if s.strip()]
                    if shared_entries:
                        preview = ", ".join(shared_entries[:5])
                        if len(shared_entries) > 5:
                            preview += f" (+{len(shared_entries) - 5} more)"
                        summary_parts.append(f"Shared with: {preview}")
                if result.get("has_anonymous_link"):
                    summary_parts.append("Anonymous link present")
                elif result.get("has_sharing_link"):
                    summary_parts.append("Sharing link present")
                if result.get("has_guest_or_external"):
                    summary_parts.append("Includes guest/external access")

                if summary_parts:
                    result["notes"] = " | ".join(summary_parts)
                else:
                    result["notes"] = f"{len(permissions)} permission object(s) detected"

        return result



class ExportManager:
    @staticmethod
    def export(
        output_folder: str,
        site_url: str,
        library_title: str,
        site_id: str,
        drive_id: str,
        flagged_items: List[FlaggedItem],
        folder_counts: List[FolderCount],
        total_items: int,
        settings: Dict,
        impact_summary: Optional[Dict] = None,
        duration_seconds: Optional[float] = None,
        generate_pdf: bool = True,
    ) -> Dict[str, str]:
        out = Path(output_folder)
        out.mkdir(parents=True, exist_ok=True)

        ts = datetime.now().strftime("%Y%m%d-%H%M%S")

        flagged_csv = out / f"preflight-flagged-items-{ts}.csv"
        counts_csv = out / f"preflight-topfolder-counts-{ts}.csv"
        summary_json = out / f"preflight-summary-{ts}.json"
        pdf_report = out / f"preflight-report-{ts}.pdf"

        with open(flagged_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=[
                    "SiteUrl", "LibraryTitle", "TopLevelFolder", "ItemType", "Name",
                    "WebUrl", "DriveItemId", "HasSharingLink", "HasAnonymousLink",
                    "HasGuestOrExternal", "SharedWith", "PermissionSignal", "Notes",
                ],
            )
            writer.writeheader()
            for item in flagged_items:
                writer.writerow({
                    "SiteUrl": item.site_url,
                    "LibraryTitle": item.library_title,
                    "TopLevelFolder": item.top_level_folder,
                    "ItemType": item.item_type,
                    "Name": item.name,
                    "WebUrl": item.web_url,
                    "DriveItemId": item.drive_item_id,
                    "HasSharingLink": item.has_sharing_link,
                    "HasAnonymousLink": item.has_anonymous_link,
                    "HasGuestOrExternal": item.has_guest_or_external,
                    "SharedWith": item.shared_with,
                    "PermissionSignal": item.permission_signal,
                    "Notes": item.notes,
                })

        with open(counts_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["SiteUrl", "LibraryTitle", "TopLevelFolder", "ItemCount"],
            )
            writer.writeheader()
            for count in folder_counts:
                writer.writerow({
                    "SiteUrl": count.site_url,
                    "LibraryTitle": count.library_title,
                    "TopLevelFolder": count.top_level_folder,
                    "ItemCount": count.item_count,
                })

        output_paths = {
            "flaggedItemsCsv": str(flagged_csv),
            "topFolderCountsCsv": str(counts_csv),
            "summaryJson": str(summary_json),
        }

        pdf_note = ""
        if generate_pdf:
            if REPORTLAB_AVAILABLE:
                try:
                    ExportManager._write_pdf_report(
                        pdf_report,
                        site_url=site_url,
                        library_title=library_title,
                        site_id=site_id,
                        drive_id=drive_id,
                        flagged_items=flagged_items,
                        folder_counts=folder_counts,
                        total_items=total_items,
                        settings=settings,
                        impact_summary=impact_summary or {},
                        duration_seconds=duration_seconds or 0.0,
                    )
                    output_paths["pdfReport"] = str(pdf_report)
                    logger.info(f"EXPORT | PDF     | {pdf_report}")
                except Exception as e:
                    pdf_note = f"PDF generation failed: {e}"
                    logger.error(f"EXPORT | PDFERR  | {e}")
            else:
                pdf_note = "ReportLab not installed. Install with: pip install reportlab"
                logger.warning("EXPORT | PDFSKIP | ReportLab not installed; skipping PDF report")

        summary = {
            "timestamp": datetime.now().isoformat(),
            "siteUrl": site_url,
            "libraryTitle": library_title,
            "siteId": site_id,
            "driveId": drive_id,
            "totalItemsScanned": total_items,
            "flaggedItemCount": len(flagged_items),
            "settingsUsed": settings,
            "impactSummary": impact_summary or {},
            "outputFilePaths": output_paths,
            "limitations": (
                "Microsoft Graph does not expose classic SharePoint HasUniqueRoleAssignments. "
                "PermissionSignal is best-effort only."
            ),
        }
        if duration_seconds is not None:
            summary["durationSeconds"] = duration_seconds
        if pdf_note:
            summary["pdfNote"] = pdf_note

        with open(summary_json, "w", encoding="utf-8") as f:
            json.dump(summary, f, indent=2)

        return output_paths

    @staticmethod
    def _format_duration(seconds: float) -> str:
        seconds = int(seconds)
        hours = seconds // 3600
        minutes = (seconds % 3600) // 60
        secs = seconds % 60
        if hours > 0:
            return f"{hours:02}:{minutes:02}:{secs:02}"
        return f"{minutes:02}:{secs:02}"

    @staticmethod
    def _bool_text(value: bool) -> str:
        return "Yes" if value else "No"

    @staticmethod
    def _folder_flag_counts(flagged_items: List[FlaggedItem]) -> Dict[str, int]:
        counts: Dict[str, int] = {}
        for item in flagged_items:
            counts[item.top_level_folder] = counts.get(item.top_level_folder, 0) + 1
        return counts

    @staticmethod
    def _write_pdf_report(
        filepath: Path,
        site_url: str,
        library_title: str,
        site_id: str,
        drive_id: str,
        flagged_items: List[FlaggedItem],
        folder_counts: List[FolderCount],
        total_items: int,
        settings: Dict,
        impact_summary: Dict,
        duration_seconds: float,
    ):
        styles = getSampleStyleSheet()
        doc = SimpleDocTemplate(
            str(filepath),
            pagesize=A4,
            leftMargin=15 * mm,
            rightMargin=15 * mm,
            topMargin=15 * mm,
            bottomMargin=15 * mm,
        )

        story = []
        folder_flag_counts = ExportManager._folder_flag_counts(flagged_items)

        story.append(Paragraph("SharePoint Online Preflight Report", styles["Title"]))
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles["Normal"]))
        story.append(Paragraph(f"Duration: {ExportManager._format_duration(duration_seconds)}", styles["Normal"]))
        story.append(Spacer(1, 10))

        meta_data = [
            ["Site URL", site_url],
            ["Library", library_title],
            ["Site ID", site_id],
            ["Drive ID", drive_id],
            ["Total Items Scanned", str(total_items)],
            ["Flagged Items", str(len(flagged_items))],
        ]
        meta_table = Table(meta_data, colWidths=[45 * mm, 130 * mm])
        meta_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EAF2FF")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#C7D2E5")),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("PADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(meta_table)
        story.append(Spacer(1, 12))

        story.append(Paragraph("Impact Summary", styles["Heading2"]))
        impact_data = [
            ["Affected Files", str(impact_summary.get("affected_files", 0))],
            ["Affected Folders", str(impact_summary.get("affected_folders", 0))],
            ["Unique Permission Signal", str(impact_summary.get("unique_permissions", 0))],
            ["Sharing Exposure", str(impact_summary.get("sharing_exposure", 0))],
            ["Anonymous Sharing", str(impact_summary.get("anonymous_sharing", 0))],
            ["External / Guest Access", str(impact_summary.get("external_guest", 0))],
        ]
        impact_table = Table(impact_data, colWidths=[70 * mm, 40 * mm])
        impact_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#FFF5E6")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E1C699")),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("PADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(impact_table)
        story.append(Spacer(1, 12))

        story.append(Paragraph("Top-level Folder Breakdown", styles["Heading2"]))
        folder_rows = [["Folder", "Item Count", "Flagged Count"]]
        for count in folder_counts:
            folder_rows.append([
                count.top_level_folder,
                str(count.item_count),
                str(folder_flag_counts.get(count.top_level_folder, 0)),
            ])
        folder_table = Table(folder_rows, colWidths=[95 * mm, 30 * mm, 30 * mm], repeatRows=1)
        folder_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAD3")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B7C9A8")),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("PADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(folder_table)
        story.append(PageBreak())

        story.append(Paragraph("Detailed Findings", styles["Heading2"]))
        if flagged_items:
            detail_rows = [[
                "Name", "Type", "Folder", "Link", "Anon", "External", "Signal", "Notes"
            ]]
            for item in flagged_items:
                # Create style for header row
                name_style = styles["Normal"]
                name_style.fontSize = 8
                
                detail_rows.append([
                    Paragraph(item.name[:32] if item.name else "", name_style),
                    Paragraph(item.item_type, name_style),
                    Paragraph(item.top_level_folder[:20] if item.top_level_folder else "", name_style),
                    Paragraph(ExportManager._bool_text(item.has_sharing_link), name_style),
                    Paragraph(ExportManager._bool_text(item.has_anonymous_link), name_style),
                    Paragraph(ExportManager._bool_text(item.has_guest_or_external), name_style),
                    Paragraph(item.permission_signal or "", name_style),
                    Paragraph(item.notes or "", name_style),
                ])
            detail_table = Table(
                detail_rows,
                colWidths=[28 * mm, 14 * mm, 20 * mm, 11 * mm, 11 * mm, 13 * mm, 20 * mm, 38 * mm],
                repeatRows=1,
            )
            detail_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9D9D9")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                ("TOPPADDING", (0, 0), (-1, 0), 6),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#999999")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F5F5")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 1), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
            ]))
            story.append(detail_table)
        else:
            story.append(Paragraph("No flagged items were found.", styles["Normal"]))

        story.append(Spacer(1, 12))
        story.append(Paragraph("Scan Settings", styles["Heading2"]))
        settings_rows = [["Setting", "Value"]] + [[str(k), str(v)] for k, v in settings.items()]
        settings_table = Table(settings_rows, colWidths=[55 * mm, 120 * mm], repeatRows=1)
        settings_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF2FF")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#C7D2E5")),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("PADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(settings_table)
        story.append(Spacer(1, 10))

        story.append(Paragraph("Limitations", styles["Heading2"]))
        story.append(Paragraph(
            "Microsoft Graph does not expose classic SharePoint HasUniqueRoleAssignments. "
            "PermissionSignal and sharing indicators in this report are best-effort signals based on Graph permissions and links.",
            styles["BodyText"],
        ))

        doc.build(story)


class ScanWorker(QThread):
    progress = Signal(str, int, int)
    log = Signal(str)
    flagged_item = Signal(object)
    completed = Signal(list, list, int, dict, dict, float)
    error = Signal(str)
    cancelled = Signal()

    def __init__(
        self,
        graph_client: GraphClient,
        site_url: str,
        library_title: str,
        drive_id: str,
        output_folder: str,
        page_size: int,
        scan_sharing: bool,
        scan_permissions: bool,
        export_pdf: bool,
        skip_count: bool = False,
    ):
        super().__init__()
        self.graph = graph_client
        self.site_url = site_url
        self.library_title = library_title
        self.drive_id = drive_id
        self.output_folder = output_folder
        self.page_size = page_size
        self.scan_sharing = scan_sharing
        self.scan_permissions = scan_permissions
        self.export_pdf = export_pdf
        self.skip_count = skip_count
        self._cancel_requested = False
        self._scanner: Optional[LibraryScanner] = None

    def cancel(self):
        self._cancel_requested = True
        if self._scanner:
            self._scanner.cancel()

    def run(self):
        start_time = time.time()
        try:
            self.progress.emit(ScanPhase.RESOLVING_SITE.value, 0, 0)
            self.log.emit(f"SITE   | RESOLVE | {self.site_url}")

            site = self.graph.resolve_site(self.site_url)
            site_id = site.get("id")
            if not site_id:
                raise RuntimeError("Could not resolve site ID.")

            if not self.drive_id:
                raise RuntimeError("No document library drive ID was supplied.")

            if self._cancel_requested:
                self.cancelled.emit()
                return

            self.progress.emit(ScanPhase.RESOLVING_LIBRARY.value, 0, 0)
            self.log.emit(f"LIB    | TARGET  | {self.library_title} | DriveId={self.drive_id}")

            self._scanner = LibraryScanner(self.graph, self.page_size)

            if self.skip_count:
                self.progress.emit(ScanPhase.ANALYZING.value, 0, 0)
                self.log.emit("SCAN   | MODE    | Fast mode enabled - skipping file/folder count")
                total_items = 0
                folder_counts_map = {"(root)": 0}
            else:
                self.progress.emit(ScanPhase.COUNTING.value, 0, 0)
                self.log.emit("SCAN   | PHASE   | Pre-counting all files and folders")

                total_items, folder_counts_map = self._scanner.pre_count(
                    drive_id=self.drive_id,
                    progress_callback=self._progress_callback,
                    log_callback=self._log_callback,
                )

                if self._cancel_requested:
                    self.cancelled.emit()
                    return

                self.log.emit(f"SCAN   | COUNT   | TotalItems={total_items}")

            flagged_items, processed = self._scanner.analyze(
                site_url=self.site_url,
                library_title=self.library_title,
                drive_id=self.drive_id,
                total_items=total_items,
                folder_counts=folder_counts_map,
                scan_sharing=self.scan_sharing,
                scan_permissions=self.scan_permissions,
                progress_callback=self._progress_callback,
                log_callback=self._log_callback,
                flagged_callback=self._flagged_callback,
            )

            if self._cancel_requested:
                self.cancelled.emit()
                return

            folder_counts = [
                FolderCount(
                    site_url=self.site_url,
                    library_title=self.library_title,
                    top_level_folder=name,
                    item_count=count,
                )
                for name, count in sorted(folder_counts_map.items(), key=lambda x: x[0].lower())
            ]

            self.progress.emit(ScanPhase.EXPORTING.value, processed, total_items)
            self.log.emit("EXPORT | WRITE   | CSV/JSON outputs")

            settings = {
                "siteUrl": self.site_url,
                "libraryTitle": self.library_title,
                "driveId": self.drive_id,
                "pageSize": self.page_size,
                "scanSharing": self.scan_sharing,
                "scanPermissions": self.scan_permissions,
            }

            impact_summary = build_impact_summary(flagged_items)
            duration_seconds = time.time() - start_time
            export_paths = ExportManager.export(
                output_folder=self.output_folder,
                site_url=self.site_url,
                library_title=self.library_title,
                site_id=site_id,
                drive_id=self.drive_id,
                flagged_items=flagged_items,
                folder_counts=folder_counts,
                total_items=total_items,
                settings=settings,
                impact_summary=impact_summary,
                duration_seconds=duration_seconds,
                generate_pdf=self.export_pdf,
            )

            self.progress.emit(ScanPhase.COMPLETED.value, total_items, total_items)
            self.log.emit(f"SCAN   | DONE    | Total={total_items} | Flagged={len(flagged_items)}")
            self.completed.emit(flagged_items, folder_counts, total_items, export_paths, impact_summary, duration_seconds)

        except Exception as e:
            logger.error(f"WORKER | ERROR   | {e}")
            logger.debug(traceback.format_exc())
            self.error.emit(str(e))

    def _progress_callback(self, phase: str, current: int, total: int):
        if not self._cancel_requested:
            self.progress.emit(phase, current, total)

    def _log_callback(self, message: str):
        if not self._cancel_requested:
            self.log.emit(message)

    def _flagged_callback(self, item: FlaggedItem):
        if not self._cancel_requested:
            self.flagged_item.emit(item)


class InfoCard(QFrame):
    def __init__(self, title: str, value: str = "-"):
        super().__init__()
        self.setObjectName("InfoCard")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("CardTitle")

        self.value_label = QLabel(value)
        self.value_label.setObjectName("CardValue")
        self.value_label.setWordWrap(True)

        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)

    def set_value(self, value: str):
        self.value_label.setText(value)


class ConnectionSetupDialog(QDialog):
    def __init__(self, parent=None, auth_manager=None, graph_client=None):
        super().__init__(parent)
        self.setWindowTitle("Connection Setup")
        self.resize(620, 360)
        self.auth_manager: Optional[AuthManager] = auth_manager
        self.graph_client: Optional[GraphClient] = graph_client
        self._build_ui()
        self._apply_styles()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        header = QLabel("Connect to Microsoft 365 before opening the preflight scanner")
        header.setObjectName("SetupHeader")
        sub = QLabel("Enter the Entra app details, authenticate interactively, then continue into the scanner.")
        sub.setWordWrap(True)
        sub.setObjectName("SetupSubheader")
        layout.addWidget(header)
        layout.addWidget(sub)

        form_group = QGroupBox("Connection Details")
        form = QFormLayout(form_group)
        self.client_id_input = QLineEdit()
        self.client_id_input.setPlaceholderText("Application (client) ID")
        self.tenant_id_input = QLineEdit()
        self.tenant_id_input.setPlaceholderText("Tenant ID or tenant domain")
        form.addRow("Client ID", self.client_id_input)
        form.addRow("Tenant ID", self.tenant_id_input)
        layout.addWidget(form_group)

        status_group = QGroupBox("Status")
        status_layout = QVBoxLayout(status_group)
        self.status_label = QLabel("Not authenticated")
        self.status_label.setObjectName("SetupStatus")
        self.detail_label = QLabel("Authenticate to continue")
        self.detail_label.setWordWrap(True)
        status_layout.addWidget(self.status_label)
        status_layout.addWidget(self.detail_label)
        layout.addWidget(status_group)

        buttons = QHBoxLayout()
        self.authenticate_button = QPushButton("Authenticate")
        self.authenticate_button.setObjectName("SecondaryButton")
        self.authenticate_button.clicked.connect(self._authenticate)
        self.continue_button = QPushButton("Open Preflight Scanner")
        self.continue_button.setObjectName("PrimaryButton")
        self.continue_button.setEnabled(False)
        self.continue_button.clicked.connect(self.accept)
        self.cancel_button = QPushButton("Exit")
        self.cancel_button.clicked.connect(self.reject)
        buttons.addWidget(self.authenticate_button)
        buttons.addStretch(1)
        buttons.addWidget(self.cancel_button)
        buttons.addWidget(self.continue_button)
        layout.addLayout(buttons)

    def _apply_styles(self):
        self.setStyleSheet("""
            QMainWindow, QWidget { font-size: 11px; }
            QGroupBox {
                font-weight: 600;
                border: 1px solid #cfcfcf;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                background: #fafafa;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
            QLineEdit { border: 1px solid #c8c8c8; border-radius: 6px; padding: 6px; background: white; }
            QPushButton { min-height: 24px; max-height: 24px; padding: 2px 10px; border-radius: 4px; border: 1px solid #bdbdbd; background: #f4f4f4; }
            QPushButton:hover { background: #ececec; }
            QPushButton:disabled { color: #888; background: #f0f0f0; }
            QPushButton#PrimaryButton { background: #1f6feb; color: white; border: 1px solid #1f6feb; font-weight: 600; }
            QPushButton#PrimaryButton:hover { background: #1857b8; }
            QPushButton#SecondaryButton { background: #2da44e; color: white; border: 1px solid #2da44e; font-weight: 600; }
            QPushButton#SecondaryButton:hover { background: #238636; }
            QLabel#SetupHeader { font-size: 18px; font-weight: 700; }
            QLabel#SetupSubheader { color: #555; }
            QLabel#SetupStatus { font-size: 14px; font-weight: 700; }
        """)

    def _authenticate(self):
        client_id = self.client_id_input.text().strip()
        tenant_id = self.tenant_id_input.text().strip()

        if not client_id:
            QMessageBox.warning(self, "Missing Client ID", "Enter the Entra application client ID.")
            return
        if not tenant_id:
            QMessageBox.warning(self, "Missing Tenant ID", "Enter the Entra tenant ID.")
            return

        self.authenticate_button.setEnabled(False)
        self.status_label.setText("Authenticating...")
        self.detail_label.setText("Waiting for Microsoft sign-in to complete...")
        QApplication.processEvents()

        try:
            auth_manager = AuthManager(client_id, tenant_id)
            auth_manager.authenticate()
            graph_client = GraphClient(auth_manager)
            me = graph_client.get_me()
            name = me.get("displayName") or auth_manager.username or "Authenticated user"

            self.auth_manager = auth_manager
            self.graph_client = graph_client
            self.status_label.setText(f"Connected as {name}")
            self.detail_label.setText("Authentication successful. Opening scanner...")
            # Auto-navigate to scanner after successful authentication
            self.accept()
        except Exception as e:
            self.auth_manager = None
            self.graph_client = None
            self.status_label.setText("Authentication failed")
            self.detail_label.setText(str(e))
            QMessageBox.critical(self, "Authentication Failed", str(e))
        finally:
            self.authenticate_button.setEnabled(True)


class MainWindow(QMainWindow):
    log_signal = Signal(str)

    def __init__(self, initial_client_id="", initial_tenant_id="", auth_manager=None, graph_client=None):
        super().__init__()
        self.setWindowTitle("SharePoint Online Preflight Scanner")
        self.resize(1320, 940)

        global logger
        logger = build_logger(self.log_signal.emit)
        self.log_signal.connect(self._append_log)

        self.auth_manager: Optional[AuthManager] = auth_manager
        self.graph_client: Optional[GraphClient] = graph_client
        self.scan_worker: Optional[ScanWorker] = None
        self.current_site_id: Optional[str] = None
        self.current_libraries: List[Dict] = []
        self.flagged_count = 0
        self.scan_start_time: Optional[float] = None

        self._setup_ui()
        self._apply_styles()
        self._update_auth_status()

    def _setup_ui(self):
        main = QWidget()
        self.setCentralWidget(main)
        root = QVBoxLayout(main)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        self.tabs = QTabWidget()
        root.addWidget(self.tabs)

        self._build_scan_tab()
        self._build_results_tab()

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Ready")

    def _build_scan_tab(self):
        scan_tab = QWidget()
        scan_layout = QVBoxLayout(scan_tab)
        scan_layout.setSpacing(10)

        banner = QFrame()
        banner.setObjectName("Banner")
        banner_layout = QGridLayout(banner)
        banner_layout.setContentsMargins(12, 10, 12, 10)

        self.banner_user = QLabel("User: Not authenticated")
        self.banner_site = QLabel("Site: -")
        self.banner_library = QLabel("Library: -")
        self.banner_status = QLabel("Status: Ready")

        banner_layout.addWidget(self.banner_user, 0, 0)
        banner_layout.addWidget(self.banner_site, 0, 1)
        banner_layout.addWidget(self.banner_library, 1, 0)
        banner_layout.addWidget(self.banner_status, 1, 1)

        scan_layout.addWidget(banner)

        top_row = QHBoxLayout()
        top_row.setSpacing(10)

        target_group = QGroupBox("Target")
        target_form = QFormLayout(target_group)
        self.site_url_input = QLineEdit()
        self.site_url_input.setPlaceholderText("https://tenant.sharepoint.com/sites/SiteName")

        lib_row = QHBoxLayout()
        self.library_combo = QComboBox()
        self.library_combo.setMinimumWidth(280)
        self.load_libraries_button = QPushButton("Load Libraries")
        self.load_libraries_button.setEnabled(False)
        self.load_libraries_button.clicked.connect(self._load_libraries)
        lib_row.addWidget(self.library_combo, 1)
        lib_row.addWidget(self.load_libraries_button)

        out_row = QHBoxLayout()
        self.output_folder_input = QLineEdit(str(Path.home() / "SharePointScanResults"))
        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self._browse_output)
        out_row.addWidget(self.output_folder_input, 1)
        out_row.addWidget(self.browse_button)

        target_form.addRow("Site URL", self.site_url_input)
        target_form.addRow("Document Library", lib_row)
        target_form.addRow("Output Folder", out_row)

        options_group = QGroupBox("Options")
        options_form = QFormLayout(options_group)
        self.page_size_input = QLineEdit(str(DEFAULT_PAGE_SIZE))
        self.scan_sharing_checkbox = QCheckBox("Scan sharing exposure")
        self.scan_sharing_checkbox.setChecked(True)
        self.scan_permissions_checkbox = QCheckBox("Scan permission signals")
        self.scan_permissions_checkbox.setChecked(True)
        self.skip_count_checkbox = QCheckBox("Skip file/folder counting (faster, less accurate progress)")
        self.skip_count_checkbox.setChecked(False)
        self.export_pdf_checkbox = QCheckBox("Generate PDF report after scan")
        self.export_pdf_checkbox.setChecked(True)
        options_form.addRow("Page Size", self.page_size_input)
        options_form.addRow("", self.scan_sharing_checkbox)
        options_form.addRow("", self.scan_permissions_checkbox)
        options_form.addRow("", self.skip_count_checkbox)
        options_form.addRow("", self.export_pdf_checkbox)

        top_row.addWidget(target_group, 3)
        top_row.addWidget(options_group, 2)
        scan_layout.addLayout(top_row)

        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout(progress_group)

        cards_row = QHBoxLayout()
        self.card_phase = InfoCard("Phase", ScanPhase.IDLE.value)
        self.card_progress = InfoCard("Progress", "0 / 0")
        self.card_flagged = InfoCard("Flagged", "0")
        self.card_duration = InfoCard("Duration", "00:00")
        self.card_impacted = InfoCard("Affected Items", "0")

        for card in (
            self.card_phase,
            self.card_progress,
            self.card_flagged,
            self.card_duration,
            self.card_impacted,
        ):
            cards_row.addWidget(card)
        progress_layout.addLayout(cards_row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)

        self.progress_detail = QLabel("Ready")
        progress_layout.addWidget(self.progress_detail)

        scan_layout.addWidget(progress_group)

        log_group = QGroupBox("Live Log")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)
        scan_layout.addWidget(log_group, 1)

        actions = QHBoxLayout()
        self.scan_button = QPushButton("Start Scan")
        self.scan_button.setObjectName("PrimaryButton")
        self.scan_button.setFixedWidth(110)
        self.scan_button.clicked.connect(self._start_scan)
        self.scan_button.setEnabled(False)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setObjectName("DangerButton")
        self.cancel_button.setFixedWidth(90)
        self.cancel_button.clicked.connect(self._cancel_scan)
        self.cancel_button.setEnabled(False)

        self.open_output_button = QPushButton("Open Output Folder")
        self.open_output_button.setFixedWidth(140)
        self.open_output_button.clicked.connect(self._open_output)

        actions.addWidget(self.scan_button)
        actions.addWidget(self.cancel_button)
        actions.addStretch(1)
        actions.addWidget(self.open_output_button)

        scan_layout.addLayout(actions)
        self.tabs.addTab(scan_tab, "Scan")

    def _build_results_tab(self):
        results_tab = QWidget()
        results_layout = QVBoxLayout(results_tab)
        results_layout.setSpacing(10)

        summary_group = QGroupBox("Summary")
        summary_layout = QVBoxLayout(summary_group)
        self.results_summary = QTextEdit()
        self.results_summary.setReadOnly(True)
        self.results_summary.setMaximumHeight(220)
        summary_layout.addWidget(self.results_summary)
        results_layout.addWidget(summary_group)

        table_group = QGroupBox("Flagged Items (Live)")
        table_layout = QVBoxLayout(table_group)
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(8)
        self.results_table.setHorizontalHeaderLabels([
            "Name",
            "Type",
            "Top Folder",
            "Sharing Link",
            "Anonymous",
            "Guest/External",
            "Permission Signal",
            "Notes",
        ])
        self.results_table.setAlternatingRowColors(True)
        self.results_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.results_table.setSelectionMode(QTableWidget.SingleSelection)
        self.results_table.verticalHeader().setVisible(False)
        self.results_table.horizontalHeader().setStretchLastSection(True)
        self.results_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        table_layout.addWidget(self.results_table)
        results_layout.addWidget(table_group, 1)

        self.tabs.addTab(results_tab, "Results")

    def _apply_styles(self):
        self.setStyleSheet("""
            QMainWindow, QWidget {
                font-size: 11px;
            }
            QGroupBox {
                font-weight: 600;
                border: 1px solid #cfcfcf;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                background: #fafafa;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px 0 4px;
            }
            QLineEdit, QComboBox, QTextEdit, QTableWidget {
                border: 1px solid #c8c8c8;
                border-radius: 6px;
                padding: 6px;
                background: white;
            }
            QPushButton {
                padding: 4px 12px;
                border-radius: 6px;
                border: 1px solid #c0c0c0;
                background: #f4f4f4;
            }
            QPushButton:hover {
                background: #ececec;
            }
            QPushButton:disabled {
                color: #888;
                background: #f0f0f0;
            }
            QPushButton#PrimaryButton {
                background: #1f6feb;
                color: white;
                border: 1px solid #1f6feb;
                font-weight: 600;
            }
            QPushButton#PrimaryButton:hover {
                background: #1857b8;
            }
            QPushButton#DangerButton {
                background: #d73a49;
                color: white;
                border: 1px solid #d73a49;
                font-weight: 600;
            }
            QPushButton#DangerButton:hover {
                background: #b92f3d;
            }
            QPushButton#SecondaryButton {
                background: #2da44e;
                color: white;
                border: 1px solid #2da44e;
                font-weight: 600;
            }
            QPushButton#SecondaryButton:hover {
                background: #238636;
            }
            QProgressBar {
                border: 1px solid #c8c8c8;
                border-radius: 6px;
                text-align: center;
                min-height: 20px;
                background: white;
            }
            QProgressBar::chunk {
                background: #1f6feb;
                border-radius: 5px;
            }
            QFrame#Banner {
                border: 1px solid #d8d8d8;
                border-radius: 8px;
                background: #f7faff;
            }
            QFrame#InfoCard {
                border: 1px solid #d8d8d8;
                border-radius: 8px;
                background: white;
            }
            QLabel#CardTitle {
                color: #666;
                font-size: 10px;
                font-weight: 600;
            }
            QLabel#CardValue {
                font-size: 14px;
                font-weight: 600;
            }
        """)

    def _append_log(self, msg: str):
        color = "#222222"
        upper = msg.upper()
        if "FLAGGED" in upper:
            color = "#9a6700"
        elif "ERROR" in upper or "FAILED" in upper:
            color = "#cf222e"
        elif "RETRY" in upper or "WARN" in upper:
            color = "#bf8700"
        elif "DONE" in upper or "CLEAR" in upper or "SUCCESS" in upper:
            color = "#1a7f37"

        self.log_text.append(f'<span style="color:{color}">{msg}</span>')
        sb = self.log_text.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _update_auth_status(self):
        authenticated = self.auth_manager and self.auth_manager.is_authenticated()

        if authenticated:
            user = self.auth_manager.username or "User"
            self.banner_user.setText(f"User: {user}")
            self.scan_button.setEnabled(True)
            self.load_libraries_button.setEnabled(True)
        else:
            self.banner_user.setText("User: Not authenticated")
            self.scan_button.setEnabled(False)
            self.load_libraries_button.setEnabled(False)

    def _browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose Output Folder", self.output_folder_input.text())
        if folder:
            self.output_folder_input.setText(folder)

    def _load_libraries(self):
        if not self.graph_client:
            QMessageBox.warning(self, "Not Authenticated", "Authenticate before loading libraries.")
            return

        site_url = self.site_url_input.text().strip()
        if not site_url:
            QMessageBox.warning(self, "Missing Site URL", "Enter a SharePoint site URL first.")
            return

        try:
            self.card_phase.set_value(ScanPhase.RESOLVING_SITE.value)
            self.banner_status.setText("Status: Loading libraries")
            self.banner_site.setText(f"Site: {site_url}")
            self.statusBar().showMessage("Loading libraries...")

            resolver = SharePointResolver(self.graph_client)
            site_id, libraries = resolver.list_libraries_for_site(site_url)

            self.current_site_id = site_id
            self.current_libraries = libraries
            self.library_combo.clear()

            if not libraries:
                QMessageBox.warning(self, "No Libraries Found", "No document libraries were found for this site.")
                self.statusBar().showMessage("No libraries found")
                self.card_phase.set_value(ScanPhase.IDLE.value)
                self.banner_status.setText("Status: Ready")
                return

            # Sort libraries alphabetically by name
            sorted_libraries = sorted(libraries, key=lambda lib: lib.get("name", "").lower())
            
            for lib in sorted_libraries:
                name = lib.get("name", "Unnamed Library")
                self.library_combo.addItem(name, {
                    "id": lib.get("id", ""),
                    "name": name,
                    "webUrl": lib.get("webUrl", ""),
                })

            logger.info(f"LIB    | LOADED  | {len(libraries)} libraries")
            self.banner_library.setText(f"Library: {self.library_combo.currentText()}")
            self.card_phase.set_value(ScanPhase.IDLE.value)
            self.banner_status.setText("Status: Ready")
            self.statusBar().showMessage(f"Loaded {len(libraries)} libraries")

        except Exception as e:
            logger.error(f"LIB    | FAILED  | {e}")
            QMessageBox.critical(self, "Load Libraries Failed", str(e))
            self.card_phase.set_value(ScanPhase.ERROR.value)
            self.banner_status.setText("Status: Failed loading libraries")
            self.statusBar().showMessage("Failed to load libraries")

    def _start_scan(self):
        if not self.graph_client:
            QMessageBox.warning(self, "Not Authenticated", "Authenticate before starting a scan.")
            return

        site_url = self.site_url_input.text().strip()
        output_folder = self.output_folder_input.text().strip()
        library_title = self.library_combo.currentText().strip()
        selected = self.library_combo.currentData()

        if not site_url:
            QMessageBox.warning(self, "Missing Site URL", "Enter a SharePoint site URL.")
            return
        if not selected or not selected.get("id"):
            QMessageBox.warning(self, "Missing Library", "Load and select a document library first.")
            return
        if not output_folder:
            QMessageBox.warning(self, "Missing Output Folder", "Enter an output folder.")
            return

        try:
            page_size = int(self.page_size_input.text().strip())
            if page_size <= 0:
                raise ValueError
        except Exception:
            QMessageBox.warning(self, "Invalid Page Size", "Page size must be a positive integer.")
            return

        Path(output_folder).mkdir(parents=True, exist_ok=True)
        drive_id = selected["id"]

        self.results_summary.clear()
        self.results_table.setRowCount(0)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.flagged_count = 0
        self.scan_start_time = time.time()

        self.card_flagged.set_value("0")
        self.card_progress.set_value("0 / 0")
        self.card_phase.set_value(ScanPhase.COUNTING.value)
        self.card_duration.set_value("00:00")
        self.card_impacted.set_value("0")
        self.banner_status.setText("Status: Scanning (fast mode)" if self.skip_count_checkbox.isChecked() else "Status: Scanning")
        self.banner_site.setText(f"Site: {site_url}")
        self.banner_library.setText(f"Library: {library_title}")
        self.progress_detail.setText("Starting scan...")

        self.site_url_input.setEnabled(False)
        self.library_combo.setEnabled(False)
        self.load_libraries_button.setEnabled(False)
        self.output_folder_input.setEnabled(False)
        self.skip_count_checkbox.setEnabled(False)
        self.scan_button.setEnabled(False)
        self.cancel_button.setEnabled(True)

        self.scan_worker = ScanWorker(
            graph_client=self.graph_client,
            site_url=site_url,
            library_title=library_title,
            drive_id=drive_id,
            output_folder=output_folder,
            page_size=page_size,
            scan_sharing=self.scan_sharing_checkbox.isChecked(),
            scan_permissions=self.scan_permissions_checkbox.isChecked(),
            export_pdf=self.export_pdf_checkbox.isChecked(),
            skip_count=self.skip_count_checkbox.isChecked(),
        )
        self.scan_worker.progress.connect(self._on_progress)
        self.scan_worker.log.connect(self._append_log)
        self.scan_worker.flagged_item.connect(self._append_flagged_item_live)
        self.scan_worker.completed.connect(self._on_completed)
        self.scan_worker.error.connect(self._on_error)
        self.scan_worker.cancelled.connect(self._on_cancelled)
        self.scan_worker.start()

        self.statusBar().showMessage("Scan in progress...")

    def _cancel_scan(self):
        if self.scan_worker:
            logger.info("SCAN   | CANCEL  | User requested cancellation")
            self.scan_worker.cancel()
            self.cancel_button.setEnabled(False)

    def _format_duration(self, seconds: float) -> str:
        seconds = int(seconds)
        hours = seconds // 3600
        minutes = (seconds % 3600) // 60
        secs = seconds % 60
        if hours > 0:
            return f"{hours:02}:{minutes:02}:{secs:02}"
        return f"{minutes:02}:{secs:02}"

    def _on_progress(self, phase: str, current: int, total: int):
        self.card_phase.set_value(phase)
        self.card_progress.set_value(f"{current} / {total if total else '?'}")

        if self.scan_start_time:
            elapsed = time.time() - self.scan_start_time
            self.card_duration.set_value(self._format_duration(elapsed))

        if total and total > 0:
            self.progress_bar.setRange(0, 100)
            percent = int((current / total) * 100)
            percent = max(0, min(100, percent))
            self.progress_bar.setValue(percent)
            self.statusBar().showMessage(f"{phase} | {current}/{total} items | {percent}%")
            self.progress_detail.setText(f"{phase} | {current}/{total} items | {percent}%")
        else:
            self.progress_bar.setRange(0, 0)
            self.statusBar().showMessage(f"{phase} | {current} items")
            self.progress_detail.setText(f"{phase} | {current} items")

    def _append_flagged_item_live(self, item: FlaggedItem):
        self.flagged_count += 1
        self.card_flagged.set_value(str(self.flagged_count))
        self.card_impacted.set_value(str(self.flagged_count))

        row = self.results_table.rowCount()
        self.results_table.insertRow(row)

        values = [
            item.name,
            item.item_type,
            item.top_level_folder,
            "Yes" if item.has_sharing_link else "",
            "Yes" if item.has_anonymous_link else "",
            "Yes" if item.has_guest_or_external else "",
            item.permission_signal,
            item.notes,
        ]

        for col, value in enumerate(values):
            cell = QTableWidgetItem(value)
            if col in (3, 4, 5) and value == "Yes":
                cell.setForeground(QColor("#9a6700"))
            if col == 6 and item.permission_signal != PermissionSignal.UNKNOWN.value:
                cell.setForeground(QColor("#b54708"))
            self.results_table.setItem(row, col, cell)

    def _on_completed(self, flagged_items: List[FlaggedItem], folder_counts: List[FolderCount], total_items: int, export_paths: Dict, impact_summary: Dict, duration_seconds: float):
        self.card_phase.set_value(ScanPhase.COMPLETED.value)
        self.card_progress.set_value(f"{total_items} / {total_items}")
        self.card_duration.set_value(self._format_duration(duration_seconds))
        self.card_flagged.set_value(str(len(flagged_items)))
        self.card_impacted.set_value(str(impact_summary["affected_files"] + impact_summary["affected_folders"]))

        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100)
        self.banner_status.setText("Status: Completed")
        self.statusBar().showMessage("Scan completed")

        lines = [
            "Scan Summary",
            "------------",
            f"Duration: {self._format_duration(duration_seconds)}",
            f"Total items scanned: {total_items}",
            f"Total flagged items: {impact_summary['total_flagged']}",
            "",
            "Affected Items",
            f"- Affected files: {impact_summary['affected_files']}",
            f"- Affected folders: {impact_summary['affected_folders']}",
            "",
            "Permission / Sharing Breakdown",
            f"- Unique permission signal: {impact_summary['unique_permissions']}",
            f"- Sharing exposure: {impact_summary['sharing_exposure']}",
            f"- Anonymous sharing: {impact_summary['anonymous_sharing']}",
            f"- External / guest access: {impact_summary['external_guest']}",
            "",
            "Top-level folder counts:",
        ]

        for count in folder_counts:
            lines.append(f"- {count.top_level_folder}: {count.item_count}")

        lines.extend([
            "",
            "Exports:",
            f"- Flagged items CSV: {export_paths.get('flaggedItemsCsv', '')}",
            f"- Top folder counts CSV: {export_paths.get('topFolderCountsCsv', '')}",
            f"- Summary JSON: {export_paths.get('summaryJson', '')}",
            f"- PDF Report: {export_paths.get('pdfReport', 'Not generated')}",
        ])

        self.results_summary.setPlainText("\n".join(lines))

        self._reset_after_scan()

        QMessageBox.information(
            self,
            "Scan Complete",
            "\n".join([
                "Scan completed successfully.",
                "",
                f"Duration: {self._format_duration(duration_seconds)}",
                f"Total scanned: {total_items}",
                f"Affected files: {impact_summary['affected_files']}",
                f"Affected folders: {impact_summary['affected_folders']}",
                f"Unique permissions: {impact_summary['unique_permissions']}",
                f"Sharing exposure: {impact_summary['sharing_exposure']}",
                f"PDF report: {'Generated' if export_paths.get('pdfReport') else 'Not generated'}",
            ]),
        )

    def _on_error(self, message: str):
        self.card_phase.set_value(ScanPhase.ERROR.value)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.banner_status.setText("Status: Error")
        self.statusBar().showMessage("Scan failed")
        self._reset_after_scan()
        QMessageBox.critical(self, "Scan Error", message)

    def _on_cancelled(self):
        self.card_phase.set_value(ScanPhase.CANCELLED.value)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.banner_status.setText("Status: Cancelled")
        self.statusBar().showMessage("Scan cancelled")
        self._reset_after_scan()

    def _reset_after_scan(self):
        self.scan_button.setEnabled(True)
        self.cancel_button.setEnabled(False)
        self.library_combo.setEnabled(True)
        self.load_libraries_button.setEnabled(True)
        self.site_url_input.setEnabled(True)
        self.output_folder_input.setEnabled(True)
        self.skip_count_checkbox.setEnabled(True)
        self.scan_worker = None
        self.scan_start_time = None

    def _open_output(self):
        folder = self.output_folder_input.text().strip()
        if folder and Path(folder).exists():
            os.startfile(folder)
        else:
            QMessageBox.warning(self, "Folder Not Found", "The output folder does not exist.")

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self,
            "Confirm Exit",
            "Are you sure you want to close the SharePoint Scanner?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    setup = ConnectionSetupDialog()
    if setup.exec() != QDialog.Accepted:
        sys.exit(0)

    window = MainWindow(
        initial_client_id=setup.client_id_input.text().strip(),
        initial_tenant_id=setup.tenant_id_input.text().strip(),
        auth_manager=setup.auth_manager,
        graph_client=setup.graph_client,
    )
    window.show()

    logger.info("APP    | START   | Application started")
    sys.exit(app.exec())


if __name__ == "__main__":
    main()