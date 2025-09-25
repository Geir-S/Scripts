# Version 1.0
# Created on 25.09.25
# Created by Geir Sandstad, Sopra Steria via Google Gemini.

# This script exports Group Policy Objects (GPOs) linked to a specified Organizational Unit (OU).
# It uses native Group Policy and Active Directory cmdlets and offers several powerful options:
#
# OPTIONS:
#
# 1. -OrganizationalUnitDN (Mandatory): The LDAP Distinguished Name of the starting OU.
# 2. -OutputPath (Mandatory): The full path where results will be saved.
# 3. -ExportChildOUs (Switch): If present, recursively exports GPOs linked to all OUs beneath the Start OU.
# 4. -LogSummary (Switch): If present, saves all console output (summary and warnings) to 'GPO_Export_Summary.txt' in the OutputPath.
#
# FOLDER STRUCTURE:
# - Reports linked directly to the Start OU are saved in the main OutputPath.
# - Reports linked to Child OUs are saved in subfolders named after the Child OU.
# - The script checks if the OutputPath exists and asks for confirmation to create or overwrite it.
#
# EXAMPLES:
#
# # 1. Export ONLY GPOs linked to the 'Sales' OU (Non-Recursive, No Log)
# .\Export-OU-GPOs.ps1 -OrganizationalUnitDN "OU=Sales,DC=corp,DC=local" -OutputPath "C:\GPO_Reports\Sales"
#
# # 2. Export GPOs from 'Users' OU and ALL child OUs (Recursive, With Log)
# .\Export-OU-GPOs.ps1 -OrganizationalUnitDN "OU=Users,DC=corp,DC=local" -OutputPath "\\Server\Reports\User_GPO_Audit" -ExportChildOUs -LogSummary

param(
    [Parameter(Mandatory=$true, HelpMessage="The LDAP Distinguished Name (DN) of the Organizational Unit (OU) where the export should start (e.g., 'OU=Users,DC=contoso,DC=com').")]
    [string]$OrganizationalUnitDN,
    
    [Parameter(Mandatory=$true, HelpMessage="The full local or UNC path where the GPO reports and log file should be saved (e.g., 'C:\GPO_Reports' or '\\Server\Share\GPOs').")]
    [string]$OutputPath,
    
    # Optional Parameter for Recursion
    [switch]$ExportChildOUs,
    
    # Optional Parameter for Logging the Summary
    [switch]$LogSummary
)

# --- Configuration ---
$ReportFormat = "html"
$LogFileName = "GPO_Export_Summary.txt"
# ---------------------

# Function to write output to console and optionally to the log file
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    # Write to console
    Write-Host $Message

    # Write to log file if -LogSummary is specified
    if ($LogSummary) {
        $Message | Out-File -FilePath $LogFilePath -Append
    }
}

# 1. --- Path Validation and Creation/Overwrite Check ---
if (-not (Test-Path -Path $OutputPath)) {
    # CASE 1: Path does NOT exist. Ask if it should be created.
    $Confirm = Read-Host "Output directory '$OutputPath' does not exist. Do you want to create it? (Y/N)"
    $UserAnswer = [char]::ToLower($Confirm)

    if ($UserAnswer -ceq 'y') {
        try {
            Write-Host "Creating main output directory: '$OutputPath'"
            New-Item -Path $OutputPath -ItemType Directory | Out-Null
        } catch {
            Write-Host "ERROR: Failed to create directory. Aborting script. Details: $($_.Exception.Message)"
            exit 1
        }
    } elseif ($UserAnswer -ceq 'n') {
        Write-Host "User aborted. Output directory not created. Aborting script."
        exit 1
    } else {
        Write-Host "Invalid input. Aborting script."
        exit 1
    }
} else {
    # CASE 2: Path DOES exist. Ask if it should be replaced/overwritten.
    $Confirm = Read-Host "Output directory '$OutputPath' already exists. Do you want to DELETE its contents and proceed? (Y/N)"
    $UserAnswer = [char]::ToLower($Confirm)

    if ($UserAnswer -ceq 'y') {
        try {
            Write-Host "Deleting existing directory and contents: '$OutputPath'"
            # Use -Force to delete read-only files if necessary
            Remove-Item -Path $OutputPath -Recurse -Force | Out-Null
            Write-Host "Recreating main output directory: '$OutputPath'"
            New-Item -Path $OutputPath -ItemType Directory | Out-Null
        } catch {
            Write-Host "ERROR: Failed to delete/recreate directory. Aborting script. Details: $($_.Exception.Message)"
            exit 1
        }
    } elseif ($UserAnswer -ceq 'n') {
        Write-Host "User aborted. Existing directory preserved. Aborting script."
        exit 1
    } else {
        Write-Host "Invalid input. Aborting script."
        exit 1
    }
}

# Now that the output path is confirmed and prepared, set the log path
$LogFilePath = Join-Path -Path $OutputPath -ChildPath $LogFileName

# 2. --- Prepare the log file ---
if ($LogSummary) {
    # This safely creates/truncates the log file, preventing errors.
    "" | Out-File -FilePath $LogFilePath 
    
    Write-Log "--- Starting GPO Export Log ($(Get-Date)) ---"
}


# 3. --- Import necessary modules ---
if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
    Write-Log "Importing ActiveDirectory module..."
    Import-Module ActiveDirectory -ErrorAction Stop
}

# 4. --- Determine scope and retrieve OUs ---
if ($ExportChildOUs) {
    Write-Log "Scope: Searching recursively in '$OrganizationalUnitDN' and all Child OUs."
    $SearchScope = "Subtree"
} else {
    Write-Log "Scope: Searching only the target OU: '$OrganizationalUnitDN'."
    $SearchScope = "OneLevel"
}

try {
    # Get the target OU and (optionally) all nested OUs
    $OUs = Get-ADOrganizationalUnit -Filter * -SearchBase $OrganizationalUnitDN -SearchScope $SearchScope -Properties DistinguishedName, Name -ErrorAction Stop
    
    # Identify the DistinguishedName of the initial parent OU
    $ParentOUDN = ($OUs | Where-Object {$_.DistinguishedName -eq $OrganizationalUnitDN}).DistinguishedName
    
} catch {
    Write-Log "ERROR: An error occurred while retrieving Organizational Units. Ensure the Group Policy and Active Directory modules are functional. Details: $($_.Exception.Message)"
    exit 1
}

$AllGpoIDs = @{} # Hashtable: Key=GPO_GUID, Value=Array_of_Target_Paths (main folder or subfolder)
$OUCount = $OUs.Count
Write-Log "Found $OUCount OUs in the scope. Processing links..."

# 5. --- Map GPOs to their final target folder paths ---
foreach ($OU in $OUs) {
    
    if ($OU.DistinguishedName -eq $ParentOUDN) {
        # Parent OU reports go directly to the main output path
        $TargetFolderKey = $OutputPath 
        Write-Log "  -> Checking links for PARENT OU: $($OU.DistinguishedName). Reports will go to main folder."
    } else {
        # Child OU reports go into a subfolder named after the child OU
        $SafeOUName = $($OU.Name) -replace '[^a-zA-Z0-9_ -]', '_'
        $TargetFolderKey = Join-Path -Path $OutputPath -ChildPath $SafeOUName
        Write-Log "  -> Checking links for CHILD OU: $($OU.DistinguishedName). Reports will go to subfolder '$SafeOUName'."
    }
    
    try {
        $OUInheritance = Get-GPInheritance -Target $OU.DistinguishedName -ErrorAction Stop
        $GpoLinks = $OUInheritance.GpoLinks | Select-Object -ExpandProperty GpoId

        if ($GpoLinks) {
            foreach ($GpoID in $GpoLinks) {
                # Store the GPO ID and associate it with the determined Target Folder Key
                if ($AllGpoIDs.ContainsKey($GpoID)) {
                    if ($TargetFolderKey -notin $AllGpoIDs[$GpoID]) {
                        $AllGpoIDs[$GpoID] += @($TargetFolderKey)
                    }
                } else {
                    $AllGpoIDs.Add($GpoID, @($TargetFolderKey))
                }
            }
        }

    } catch {
        Write-Log "WARNING: Failed to get GPO links for OU '$($OU.DistinguishedName)'. Skipping. Details: $($_.Exception.Message)"
    }
}

$UniqueGpoIDs = $AllGpoIDs.Keys
if ($UniqueGpoIDs.Count -eq 0) {
    Write-Log "WARNING: No GPOs found linked within the specified scope."
    exit 1
}

Write-Log "`nFound $($UniqueGpoIDs.Count) unique GPOs to export. Starting export..."

$ExportCounter = 0
# 6. --- Export GPO Reports ---
foreach ($GpoID in $UniqueGpoIDs) {
    try {
        $GpoObject = Get-GPO -Guid $GpoID -ErrorAction Stop
        $SafeGpoName = $GpoObject.DisplayName -replace '[^a-zA-Z0-9_ -]', '_'
        $FileName = "$SafeGpoName.$ReportFormat"
        
        # Get the unique list of target folder paths for this GPO
        $TargetPaths = $AllGpoIDs[$GpoID] | Select-Object -Unique

        foreach ($TargetPath in $TargetPaths) {
            
            # Create the subfolder if needed (only for child OUs, not the main OutputPath)
            if ($TargetPath -ne $OutputPath) {
                if (-not (Test-Path -Path $TargetPath)) {
                    New-Item -Path $TargetPath -ItemType Directory | Out-Null
                }
            }

            $FullFilePath = Join-Path -Path $TargetPath -ChildPath $FileName
            
            # Determine folder name for logging
            $FolderNameForLog = if ($TargetPath -eq $OutputPath) {"(Main Output Folder)"} else {(Split-Path -Path $TargetPath -Leaf)}
            
            Write-Log "   Exporting '$($GpoObject.DisplayName)' to folder: '$FolderNameForLog'"
            
            # Generate the GPO report
            Get-GPOReport -Guid $GpoID -ReportType $ReportFormat -Path $FullFilePath -ErrorAction Stop
        }
        $ExportCounter++
        
    } catch {
        Write-Log "WARNING: Failed to export GPO with ID '$GpoID'. Details: $($_.Exception.Message)"
    }
}

# 7. --- Final Summary ---
Write-Log "--------------------------------------------------------"
Write-Log "Export completed."
Write-Log "Total unique GPOs found: $($UniqueGpoIDs.Count)"
Write-Log "Total unique GPOs successfully exported: $ExportCounter"
Write-Log "Reports saved in main folder and/or child OU subfolders under: $OutputPath"
if ($LogSummary) {
    Write-Log "Log file saved to: $LogFilePath"
}
Write-Log "--------------------------------------------------------"