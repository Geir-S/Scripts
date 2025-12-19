<#
.SYNOPSIS
    Analyzes Citrix Profile Management shares for stale or legacy profiles.

.DESCRIPTION
    This script traverses the profile share and flags folders based on age, 
    naming conventions (e.g., .old), and legacy OS types (e.g., Win7).

.PARAMETER IncludeSize
    Switch. If present, the script calculates the total size of each profile folder in MB. 
    Note: This significantly increases execution time.

.PARAMETER ProfileRootPath
    The UNC path to the root of the Citrix VDI profile share. 
    Default is \\profileserver\vdiprofiles.

.PARAMETER ExportPath
    The local or network path where the final CSV report will be saved. 
    Default is the current user's desktop.

.EXAMPLE
    .\CitrixCleanup.ps1 -IncludeSize
#>

param (
    [Parameter(Mandatory=$false)]
    [switch]$IncludeSize,

    [Parameter(Mandatory=$false)]
    [string]$ProfileRootPath = "\\profileserver\vdiprofiles",

    [Parameter(Mandatory=$false)]
    [string]$ExportPath = "$env:USERPROFILE\Desktop\Citrix_Profile_Cleanup_Report.csv"
)

# --- Configuration ---
$CutoffDate = (Get-Date).AddMonths(-6)
$LegacyOSKeywords = "Win2012x64", "Win7x64", "Win2008x64"

# 1. Validation
if (!(Test-Path $ProfileRootPath)) {
    Write-Host "ERROR: Cannot reach $ProfileRootPath." -ForegroundColor Red
    return
}

Write-Host "Gathering folder list..." -ForegroundColor Cyan
# Pre-filter to find Depth 1 folders: Root\User\OS_Version
$AllFolders = Get-ChildItem -Path $ProfileRootPath -Directory -Recurse -Depth 1 -ErrorAction SilentlyContinue | 
              Where-Object { ($_.FullName.Replace($ProfileRootPath, "").TrimStart('\').Split('\')).Count -eq 2 }

$totalFolders = $AllFolders.Count
if ($totalFolders -eq 0) {
    Write-Warning "No profile folders found at the expected depth. Check permissions/path."
    return
}

$counter = 0

# 2. Process Folders
$results = foreach ($folder in $AllFolders) {
    $counter++
    $relativePath = $folder.FullName.Replace($ProfileRootPath, "").TrimStart('\')
    $pathParts = $relativePath.Split('\')
    
    $userName = $pathParts[0]
    $osFolderName = $pathParts[1]

    # Update Progress Bar
    Write-Progress -Activity "Analyzing Citrix Profiles" `
                   -Status "Processing: $userName ($osFolderName)" `
                   -PercentComplete (($counter / $totalFolders) * 100)

    $lastUsed = $folder.LastWriteTime
    $daysIdle = [math]::Round(((Get-Date) - $lastUsed).TotalDays)
    
    # --- Optional Size Calculation (MB) ---
    $sizeMB = "Skipped"
    if ($IncludeSize) {
        $sizeInBytes = (Get-ChildItem $folder.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        # Calculate in MB
        $sizeMB = if ($sizeInBytes) { [math]::Round($sizeInBytes / 1MB, 2) } else { 0 }
    }

    # --- Logic Checks ---
    # Check 1: Legacy OS Name
    $isLegacyOS = $false
    foreach ($keyword in $LegacyOSKeywords) {
        if ($osFolderName -like "*$keyword*") { $isLegacyOS = $true; break }
    }

    # Check 2: .old/_old naming convention within last 5 chars of User ID
    $isOldNamingPattern = $false
    if ($userName.Length -ge 5) {
        $lastFive = $userName.Substring($userName.Length - 5).ToLower()
        if ($lastFive -like "*old*" -and ($userName -match '[\._]')) {
            $isOldNamingPattern = $true
        }
    }

    # --- Status Assignment ---
    if ($isLegacyOS) {
        $status = "Old Citrix site"
    }
    elseif ($isOldNamingPattern -or ($lastUsed -lt $CutoffDate)) {
        $status = "Eligible for deletion"
    }
    else {
        $status = "Active"
    }

    [PSCustomObject]@{
        "User ID"           = $userName
        "OS Version Folder" = $osFolderName
        "Last Modified"     = $lastUsed
        "Days Idle"         = $daysIdle
        "Size (MB)"         = $sizeMB
        "Status"            = $status
        "Full Path"         = $folder.FullName
    }
}

# 3. Export and Finalize
if ($results) {
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nReport Success! File saved to: $ExportPath" -ForegroundColor Green
    
    # Quick summary of findings
    $results | Group-Object Status | Select-Object Name, Count | Format-Table -AutoSize
    
    Invoke-Item $ExportPath
}