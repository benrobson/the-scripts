import os
import csv
import time
import requests
from urllib.parse import quote
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

UNIFI_BASE_URL = "https://api.ui.com/v1"


def log(message):
    print(f"[{time.strftime('%H:%M:%S')}] {message}", flush=True)


def get_env_or_input(env_name: str, prompt_text: str) -> str:
    value = os.environ.get(env_name)
    if not value:
        value = input(prompt_text).strip()
    return value


def normalize_mac(mac: str) -> str:
    if not mac:
        return ""
    return (
        mac.replace(":", "")
           .replace("-", "")
           .replace(".", "")
           .strip()
           .upper()
    )


def build_session() -> requests.Session:
    session = requests.Session()

    retry = Retry(
        total=5,
        connect=3,
        read=3,
        backoff_factor=1.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
        respect_retry_after_header=True,
    )

    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    return session


def get_unifi_headers(api_key: str) -> dict:
    return {
        "X-API-KEY": api_key,
        "Accept": "application/json",
    }


def get_snipeit_headers(api_key: str) -> dict:
    return {
        "Authorization": f"Bearer {api_key}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


def fetch_unifi_devices(session: requests.Session, unifi_api_key: str) -> list[dict]:
    headers = get_unifi_headers(unifi_api_key)

    log("Requesting UniFi sites...")
    sites_resp = session.get(f"{UNIFI_BASE_URL}/sites", headers=headers, timeout=30)
    sites_resp.raise_for_status()
    log("Sites retrieved successfully")

    log("Requesting UniFi devices...")
    devices_resp = session.get(f"{UNIFI_BASE_URL}/devices", headers=headers, timeout=60)
    devices_resp.raise_for_status()
    log("Devices retrieved successfully")

    hosts = devices_resp.json().get("data", [])
    all_devices = []

    for host in hosts:
        site_name = host.get("hostName", "N/A")

        for device in host.get("devices", []):
            if device.get("status") != "online":
                continue

            raw_mac = device.get("mac", "")
            stripped_mac = normalize_mac(raw_mac)

            all_devices.append({
                "Site": site_name,
                "Model": device.get("model", "N/A"),
                "Device Name": device.get("name", "N/A"),
                "MAC Address": raw_mac,
                "MAC Stripped": stripped_mac,
                "Adoption Time": device.get("adoptionTime", "N/A"),
                "Last Seen": device.get("startupTime", "N/A"),
                "Status": device.get("status", "N/A"),
            })

    log(f"Found {len(all_devices)} online UniFi devices")

    return all_devices


def lookup_snipeit_asset_by_serial(session, snipeit_url, api_key, serial):
    base_url = snipeit_url.rstrip("/")
    headers = get_snipeit_headers(api_key)

    url = f"{base_url}/api/v1/hardware/byserial/{quote(serial)}"

    try:
        resp = session.get(url, headers=headers, timeout=30)

        if resp.status_code == 429:
            retry_after = resp.headers.get("Retry-After", "unknown")
            log(f"⚠ Rate limited by SnipeIT (429). Retry-After={retry_after}")
            return {"found": False, "error": "Rate limited"}

        resp.raise_for_status()
        payload = resp.json()

        asset = None

        if isinstance(payload, dict):
            if payload.get("payload"):
                asset = payload["payload"]
            elif payload.get("rows"):
                if payload["rows"]:
                    asset = payload["rows"][0]
            elif payload.get("id"):
                asset = payload

        if not asset:
            return {"found": False, "error": "Not found"}

        return {
            "found": True,
            "id": asset.get("id"),
            "tag": asset.get("asset_tag"),
            "name": asset.get("name"),
            "serial": asset.get("serial"),
        }

    except Exception as e:
        return {"found": False, "error": str(e)}


def export_results(rows, filename):
    fields = [
        "Site",
        "Model",
        "Device Name",
        "MAC Address",
        "MAC Stripped",
        "Adoption Time",
        "Last Seen",
        "Status",
        "SnipeIT Found",
        "SnipeIT Asset Tag",
        "SnipeIT Name",
        "SnipeIT Serial",
        "Error"
    ]

    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def main():
    log("Starting UniFi → SnipeIT comparison script v1.1.2")

    unifi_api_key = get_env_or_input("UNIFI_API_KEY", "Enter UniFi API key: ")
    snipeit_api_key = get_env_or_input("SNIPEIT_API_KEY", "Enter SnipeIT API key: ")
    snipeit_url = get_env_or_input("SNIPEIT_URL", "Enter SnipeIT URL: ")

    session = build_session()

    devices = fetch_unifi_devices(session, unifi_api_key)

    total = len(devices)
    found_count = 0
    missing_count = 0

    results = []

    for i, device in enumerate(devices, start=1):

        mac = device["MAC Stripped"]

        log(f"[{i}/{total}] Checking MAC {mac} ({device['Device Name']})")

        match = lookup_snipeit_asset_by_serial(
            session,
            snipeit_url,
            snipeit_api_key,
            mac
        )

        if match["found"]:
            found_count += 1
        else:
            missing_count += 1

        results.append({
            **device,
            "SnipeIT Found": "Yes" if match["found"] else "No",
            "SnipeIT Asset Tag": match.get("tag", ""),
            "SnipeIT Name": match.get("name", ""),
            "SnipeIT Serial": match.get("serial", ""),
            "Error": match.get("error", ""),
        })

        time.sleep(0.35)

    filename = "unifi_devices_snipeit_comparison_v1_1_2.csv"

    log("Writing CSV export...")
    export_results(results, filename)

    log("Export complete")
    log(f"Devices checked: {total}")
    log(f"Matches found: {found_count}")
    log(f"Missing assets: {missing_count}")
    log(f"CSV written to: {filename}")


if __name__ == "__main__":
    main()