# 1. Configuration
$AdminAddress = "tl00ctxcdc11p.tl.ad"
$FilePath = "$env:USERPROFILE\Desktop\Full_CitrixAppAssignments.csv"

# 2. Load Citrix Snap-ins
if (!(Get-PSSnapin -Name Citrix.* -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Citrix.* -ErrorAction SilentlyContinue
}

Write-Host "Fetching applications and resolving all assignment paths..." -ForegroundColor Cyan

# 3. Fetch all Applications
$results = Get-BrokerApplication -AdminAddress $AdminAddress | ForEach-Object {
    $app = $_
    $dgNames = @()

    # Path A: Direct Delivery Group Association
    if ($app.AssociatedDesktopGroupNames) {
        $dgNames += $app.AssociatedDesktopGroupNames
    } 
    # Fallback if names are missing but UIDs exist
    elseif ($app.AssociatedDesktopGroupUids) {
        foreach ($uid in $app.AssociatedDesktopGroupUids) {
            $dg = Get-BrokerDesktopGroup -Uid $uid -AdminAddress $AdminAddress -ErrorAction SilentlyContinue
            if ($dg) { $dgNames += $dg.Name }
        }
    }

    # Path B: Indirect via Application Groups
    if ($app.AssociatedApplicationGroupUids) {
        foreach ($agUid in $app.AssociatedApplicationGroupUids) {
            $appGroup = Get-BrokerApplicationGroup -Uid $agUid -AdminAddress $AdminAddress -ErrorAction SilentlyContinue
            if ($appGroup.AssociatedDesktopGroupNames) {
                $dgNames += $appGroup.AssociatedDesktopGroupNames
            }
        }
    }

    # Clean up the list (unique names and sorted)
    $cleanList = ($dgNames | Select-Object -Unique | Sort-Object | Where-Object { $_ -ne $null })
    $finalDGs = $cleanList -join ", "

    # Creating the object with Enabled at the end
    [PSCustomObject]@{
        "Application Name" = $app.Name
        "Published Name"   = $app.PublishedName
        "Delivery Groups"  = if ($finalDGs) { $finalDGs } else { "Unassigned" }
        "Enabled"          = $app.Enabled
    }
}

# 4. Output to Console
$results | Sort-Object "Application Name" | Format-Table -AutoSize

# 5. Export to CSV
$results | Export-Csv -Path $FilePath -NoTypeInformation
Write-Host "`nSuccess! Application report saved to: $FilePath" -ForegroundColor Green
