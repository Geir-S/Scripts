
<#
.SYNOPSIS
Exports all published Citrix applications with their command line executable and arguments to CSV.

.DESCRIPTION
- Queries Citrix published applications via Get-BrokerApplication (Citrix Studio SDK).
- Outputs a CSV with core fields: Published Name, Application Name (ID), Enabled, CommandLineExecutable, CommandLineArguments.
- Optionally enriches with additional metadata (ApplicationType, ClientFolder, WorkingDirectory, Delivery Groups, etc.).

.REQUIREMENTS
- Run on a Citrix Delivery Controller or admin VM with the Citrix Studio SDK installed.
- PowerShell 5.1
#>

[CmdletBinding()]
param(
    [string]$CsvPath = "C:\Temp\PublishedApps_Executable_Arguments.csv",

    # If set, restrict to applications that are visible/published to users.
    [switch]$PublishedOnly,

    # If set, outputs only the five core columns (as in your original script).
    [switch]$Minimal
)

$ErrorActionPreference = 'Stop'

Write-Host "Loading Citrix modules/snap-ins..." -ForegroundColor Cyan
try {
    # Try both module and snap-in approaches for different SDK versions
    if (-not (Get-Module -ListAvailable -Name Citrix.Broker.Admin.*)) {
        # Fall back to snap-ins for older SDKs
        asnp Citrix* -ErrorAction SilentlyContinue
    } else {
        Import-Module Citrix.Broker.Admin.* -ErrorAction Stop
    }
} catch {
    Write-Warning "Could not pre-load Citrix modules/snap-ins automatically. Continuing to verify cmdlets..."
}

# Verify cmdlet
if (-not (Get-Command Get-BrokerApplication -ErrorAction SilentlyContinue)) {
    Write-Error "Get-BrokerApplication not found. Install the Citrix Studio SDK and run on a Delivery Controller/admin node."
    return
}

# Ensure output folder exists
$folder = Split-Path -Path $CsvPath -Parent
if ([string]::IsNullOrWhiteSpace($folder)) {
    $folder = "."
}
if (-not (Test-Path -LiteralPath $folder)) {
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
}

Write-Host "Querying published applications..." -ForegroundColor Cyan

# Build base query
# Note: Get-BrokerApplication doesn't expose a simple -Published switch. We can approximate "published"
# as applications that are visible and not hidden; however, some environments rely on 'Visible' and
# 'Enabled' to control availability. We'll provide a post-filter to meet the user's intent.
try {
    $apps = Get-BrokerApplication -MaxRecordCount 100000
} catch {
    Write-Error "Failed to query Broker Applications. $_"
    return
}

if (-not $apps -or $apps.Count -eq 0) {
    Write-Warning "No applications found or visible. Check permissions/site connectivity."
    return
}

# Optionally restrict to 'published-looking' apps
if ($PublishedOnly.IsPresent) {
    # Typical published criteria: Visible = $true (apps appear in Storefront/CVAD) and not a special/hidden object.
    # Enabled indicates launchable, but some orgs keep Enabled=$false for temporary maintenance; we won't force it.
    $apps = $apps | Where-Object { $_.Visible -eq $true }
    if (-not $apps -or $apps.Count -eq 0) {
        Write-Warning "No visible (published) applications found after filtering."
        return
    }
}

# Optionally resolve Desktop Group UIDs to names (best-effort)
$dgByUid = @{}
try {
    if (Get-Command Get-BrokerDesktopGroup -ErrorAction SilentlyContinue) {
        $allDgs = Get-BrokerDesktopGroup -MaxRecordCount 100000
        foreach ($dg in $allDgs) {
            $dgByUid[$dg.Uid] = $dg.Name
        }
    }
} catch {
    Write-Warning "Could not resolve Delivery Group names: $_"
}

# Build result objects (PS 5.1-safe; avoid PS 7 inline ternary where possible)
Write-Host "Building result set..." -ForegroundColor Cyan

# Helper to safely stringify Delivery Group names from UIDs
function Resolve-DesktopGroupNames {
    param([int[]]$Uids)
    if (-not $Uids -or $Uids.Count -eq 0) { return $null }
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($id in $Uids) {
        if ($dgByUid.ContainsKey($id)) {
            $null = $names.Add($dgByUid[$id])
        } else {
            $null = $names.Add([string]$id)
        }
    }
    return ($names -join '; ')
}

# Compose objects
$result = foreach ($app in $apps) {

    $displayName = if ($app.PublishedName) { $app.PublishedName }
                   elseif ($app.ApplicationName) { $app.ApplicationName }
                   else { '<unnamed>' }

    $exe  = if ($app.CommandLineExecutable) { $app.CommandLineExecutable } else { '<none>' }
    $args = if ($app.CommandLineArguments)  { $app.CommandLineArguments }  else { '<none>' }

    if ($Minimal) {
        # Exact core set (matches your original request)
        [PSCustomObject]@{
            'Published Name'        = $displayName
            'Application Name (ID)' = $app.ApplicationName
            'Enabled'               = [bool]$app.Enabled
            'CommandLineExecutable' = $exe
            'CommandLineArguments'  = $args
        }
        continue
    }

    # Enriched set: add commonly useful context if not using -Minimal
    $dgNames = Resolve-DesktopGroupNames -Uids $app.AssociatedDesktopGroupUids

    [PSCustomObject]@{
        'Application Name (ID)'    = $app.ApplicationName
        'Published Name'           = $displayName
        'WorkingDirectory'         = $app.WorkingDirectory
        'CommandLineExecutable'    = $exe
        'CommandLineArguments'     = $args
        
    }
}

# Show a console table for quick verification (sorted by Published Name)
$result | Sort-Object 'Published Name' | Format-Table -AutoSize

# Export CSV (UTF8)
try {
    $result | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $CsvPath
    Write-Host ("`nExported {0} applications to: {1}" -f $result.Count, $CsvPath) -ForegroundColor Green
} catch {
    Write-Error "Failed to export CSV to '$CsvPath'. $_"
}
