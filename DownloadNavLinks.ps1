# Parameters
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$exportPath = "C:\Temp\SharePointNavigationReport_$($timestamp).csv"
$results = @()
$seenLinks = @{} 

if (!(Test-Path "C:\Temp")) { New-Item -ItemType Directory -Path "C:\Temp" -Force }

# 1. Reset Session completely
Disconnect-PnPOnline -ErrorAction SilentlyContinue

# 2. Connect to Admin Service
# We use -UseWebLogin to support MFA. 
Write-Host "Please log in via the browser window..." -ForegroundColor Cyan
try {
    # If your PnP version supports it, -ClearTokenCache is best here.
    # If it throws an error, I will provide a version without it.
    $adminConn = Connect-PnPOnline -Url "https://wkx8x-admin.sharepoint.com" -UseWebLogin -ReturnConnection
} catch {
    Write-Error "Failed to connect to Admin service."
    return
}

# 3. Get ALL Sites
$sites = Get-PnPTenantSite -Connection $adminConn
Write-Host "Total sites found: $($sites.Count)" -ForegroundColor Cyan

foreach ($site in $sites) {
    Write-Host "`n--- Processing: $($site.Url) ---" -ForegroundColor Yellow
    
    # We reuse the web session so you only log in ONCE for the whole tenant
    $siteConn = Connect-PnPOnline -Url $site.Url -UseWebLogin -ReturnConnection
    
    $locations = @("QuickLaunch", "TopNavigationBar", "HubWebRelative")
    
    foreach ($location in $locations) {
        try {
            $nodes = Get-PnPNavigationNode -Location $location -Connection $siteConn -ErrorAction Stop
            if ($null -eq $nodes) { continue }

            foreach ($node in $nodes) {
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

# 4. Final Export
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding utf8 -Force
    Write-Host "`nDONE! Captured: $($results.Count) unique links." -ForegroundColor Green
    Invoke-Item "C:\Temp"
}

Disconnect-PnPOnline -Connection $adminConn
