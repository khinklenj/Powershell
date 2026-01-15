# Parameters
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$exportPath = "C:\Temp\SharePointNavigationReport_$($timestamp).csv"
$results = @()
$seenLinks = @{} # Used to track and prevent duplicates

if (!(Test-Path "C:\Temp")) { New-Item -ItemType Directory -Path "C:\Temp" -Force }

# 1. Connect to Admin Service
Write-Host "Connecting to Admin Service..." -ForegroundColor Cyan
$adminConn = Connect-PnPOnline -Url "https://wkx8x-admin.sharepoint.com" -UseWebLogin -ReturnConnection

# Set your email manually here if the auto-detection failed previously
$myUserEmail = "admin@wkx8x.onmicrosoft.com" 
Write-Host "Elevating permissions as: $myUserEmail" -ForegroundColor Gray

# 2. Get ALL Sites
$sites = Get-PnPTenantSite -Connection $adminConn
Write-Host "Total sites found: $($sites.Count)" -ForegroundColor Cyan

foreach ($site in $sites) {
    Write-Host "`n--- Processing: $($site.Url) ---" -ForegroundColor Yellow
    
    # ELEVATION
    Set-PnPTenantSite -Url $site.Url -Owners $myUserEmail -Connection $adminConn -ErrorAction SilentlyContinue

    # Connect to the specific site
    $siteConn = Connect-PnPOnline -Url $site.Url -UseWebLogin -ReturnConnection
    
    $locations = @("QuickLaunch", "TopNavigationBar", "HubWebRelative")
    
    foreach ($location in $locations) {
        try {
            $nodes = Get-PnPNavigationNode -Location $location -Connection $siteConn -ErrorAction Stop
            if ($null -eq $nodes) { continue }

            foreach ($node in $nodes) {
                # Create a unique key based on Site, Title, and URL
                $key = "$($site.Url)_$($node.Title)_$($node.Url)"
                
                if (-not $seenLinks.ContainsKey($key)) {
                    $results += [PSCustomObject]@{
                        SiteUrl    = $site.Url
                        MenuType   = if($location -eq "HubWebRelative") { "Hub" } else { "Local" }
                        Location   = $location
                        Title      = $node.Title
                        Url        = $node.Url
                        Level      = "Parent"
                    }
                    $seenLinks[$key] = $true
                }
                
                # Check for Children
                $children = Get-PnPNavigationNode -Id $node.Id -Connection $siteConn -ErrorAction SilentlyContinue
                foreach ($child in $children) {
                    $childKey = "$($site.Url)_$($child.Title)_$($child.Url)"
                    
                    if (-not $seenLinks.ContainsKey($childKey)) {
                        $results += [PSCustomObject]@{
                            SiteUrl    = $site.Url
                            MenuType   = if($location -eq "HubWebRelative") { "Hub" } else { "Local" }
                            Location   = $location
                            Title      = $child.Title
                            Url        = $child.Url
                            Level      = "Child"
                        }
                        $seenLinks[$childKey] = $true
                    }
                }
            }
        } catch { }
    }
}

# 3. Final Export
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding utf8 -Force
    Write-Host "`nDONE! Unique links captured: $($results.Count)" -ForegroundColor Green
    Write-Host "File saved to: $exportPath" -ForegroundColor Green
    Invoke-Item "C:\Temp"
} else {
    Write-Warning "No unique links were found."
}

Disconnect-PnPOnline -Connection $adminConn