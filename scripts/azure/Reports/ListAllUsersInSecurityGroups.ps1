<#
.SYNOPSIS
Exports Entra ID (Azure AD) user -> group memberships to CSV using Microsoft Graph PowerShell.

.DESCRIPTION
- Retrieves all users of userType 'Member'
- For each user, queries memberOf
- Filters to groups only (excludes directory roles)
- Optional keyword filter on group displayName
- Optional filter to security-enabled groups only
- Exports results to CSV

.REQUIREMENTS
Install-Module Microsoft.Graph -Scope CurrentUser
Connect-MgGraph with scopes: User.Read.All, GroupMember.Read.All
#>

[CmdletBinding()]
param(
    [bool]$UseKeyword = $true,
    [string]$Keyword  = "SEC_",
    [string]$OutFile  = "C:\Data\TenantUserGroups_Filtered.csv",
    [bool]$SecurityEnabledOnly = $false,
    [bool]$IncludeUsersWithNoMatches = $false
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Ensure-OutputDirectory {
    param([Parameter(Mandatory)][string]$Path)
    $dir = Split-Path -Path $Path -Parent
    if ($dir) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

function Ensure-GraphConnection {
    $ctx = $null
    try { $ctx = Get-MgContext } catch {}
    if (-not $ctx -or -not $ctx.Account) {
        Connect-MgGraph -Scopes "User.Read.All","GroupMember.Read.All" | Out-Null
    }
}

function Get-GroupByDirectoryObjectId {
    param([Parameter(Mandatory)][string]$Id)

    # memberOf can contain group OR directoryRole OR other directoryObjects.
    # This tries to fetch it as a Group; if it isn't a group, Graph throws and we return $null.
    try {
        return Get-MgGroup -GroupId $Id -Property "id,displayName,securityEnabled" -ErrorAction Stop
    }
    catch {
        return $null
    }
}

try {
    if (-not $UseKeyword -and $OutFile -eq "C:\Data\TenantUserGroups_Filtered.csv") {
        $OutFile = "C:\Data\TenantUserGroups_All.csv"
    }

    Ensure-OutputDirectory -Path $OutFile
    Ensure-GraphConnection

    $users = Get-MgUser -All -Filter "userType eq 'Member'" -Property "id,displayName,mail,userPrincipalName"

    $rows = foreach ($u in $users) {
        $email = if ($u.Mail) { $u.Mail } else { $u.UserPrincipalName }

        # Get directory objects user is a member of
        $memberOf = Get-MgUserMemberOf -UserId $u.Id -All

        # Convert directory objects -> actual groups (robust)
        $groups = foreach ($obj in $memberOf) {
            $objId = $null

            # Different object shapes expose Id differently; handle both.
            if ($obj.PSObject.Properties.Name -contains "Id") {
                $objId = $obj.Id
            } elseif ($obj.AdditionalProperties -and $obj.AdditionalProperties.ContainsKey("id")) {
                $objId = $obj.AdditionalProperties["id"]
            }

            if ($objId) {
                $g = Get-GroupByDirectoryObjectId -Id $objId
                if ($null -ne $g) { $g }
            }
        }

        if ($SecurityEnabledOnly) {
            $groups = $groups | Where-Object { $_.SecurityEnabled -eq $true }
        }

        if ($UseKeyword) {
            $groups = $groups | Where-Object { $_.DisplayName -like "*$Keyword*" }
        }

        $groupNames = ($groups | Select-Object -ExpandProperty DisplayName) -join "; "

        if ($IncludeUsersWithNoMatches -or $groupNames) {
            [PSCustomObject]@{
                Name           = $u.DisplayName
                Email          = $email
                SecurityGroups = $groupNames
            }
        }
    }

    $rows | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
    Write-Host "Export complete. File saved at $OutFile"
}
catch {
    Write-Error ("Export failed: " + $_.Exception.Message)
    throw
}
