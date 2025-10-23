# === SharePoint Online RandomUser -> List importer (site-scoped canonical link, WinPS 5.1 safe) ===
# Works with:
#   - Windows PowerShell 5.1 + SharePointPnPPowerShellOnline  (legacy)
#   - PowerShell 7+         + PnP.PowerShell                  (modern)
# Install if needed:
#   WinPS 5.1: Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser
#   PS7+     : Install-Module PnP.PowerShell -Scope CurrentUser

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,                 # e.g. https://wkx8x.sharepoint.com/sites/wlive
    [string]$ListTitle = "RandomUser7",
    [int]$Results = 5,
    [switch]$OpenInBrowser
)

$ErrorActionPreference = 'Stop'
function Get-String([object]$v) { if ($null -eq $v) { "" } else { [string]$v } }

# --- Detect module/cmdlets and connect ---
$hasAddList = (Get-Command -Name Add-PnPList -ErrorAction SilentlyContinue) -ne $null
$hasNewList = (Get-Command -Name New-PnPList -ErrorAction SilentlyContinue) -ne $null

if ($hasAddList) {
    Write-Host "Detected PnP.PowerShell (modern). Using -Interactive auth." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -Interactive
} elseif ($hasNewList) {
    Write-Host "Detected SharePointPnPPowerShellOnline (legacy). Using -UseWebLogin." -ForegroundColor Yellow
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
} else {
    throw "No PnP module found. Install 'PnP.PowerShell' (preferred) or 'SharePointPnPPowerShellOnline'."
}

# --- Ensure list exists (cmdlet -> fallback to CSOM) ---
function Ensure-List {
    param([string]$Title)

    $existing = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if ($existing) { return $existing }

    Write-Host "Creating list '$Title'..."
    $created = $false
    try {
        if ($hasAddList) { Add-PnPList -Title $Title -Template GenericList -OnQuickLaunch:$true | Out-Null; $created = $true }
        elseif ($hasNewList) { New-PnPList -Title $Title -Template GenericList -OnQuickLaunch | Out-Null; $created = $true }
    } catch { Write-Warning ("Cmdlet create failed: " + $_.Exception.Message) }

    if (-not $created) {
        Write-Warning "Falling back to CSOM list creation."
        $ctx = Get-PnPContext
        $web = $ctx.Web
        $ctx.Load($web); $ctx.ExecuteQuery()
        $lc = New-Object Microsoft.SharePoint.Client.ListCreationInformation
        $lc.Title = $Title
        $lc.TemplateType = 100                       # Generic List
        $lc.QuickLaunchOptions = [Microsoft.SharePoint.Client.QuickLaunchOptions]::On
        $null = $web.Lists.Add($lc); $ctx.ExecuteQuery()
    }

    $attempt = 0
    do {
        Start-Sleep -Milliseconds 800; $attempt++
        $existing = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    } while (-not $existing -and $attempt -lt 20)

    if (-not $existing) { throw "List '$Title' was not provisioned after create attempt." }
    return $existing
}
$list = Ensure-List -Title $ListTitle

# --- Ensure fields (Text) ---
function Ensure-TextField {
    param([string]$InternalName, [string]$DisplayName)
    $f = Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction SilentlyContinue
    if (-not $f) {
        Write-Host "Adding field $DisplayName ($InternalName)"
        Add-PnPField -List $ListTitle -DisplayName $DisplayName -InternalName $InternalName -Type Text | Out-Null
    }
}
$contactFields = @("FirstName","LastName","Email","Phone","Cell","Street","City","State","PostalCode","Country")
@(
    @{n="FirstName";   d="First Name"},
    @{n="LastName";    d="Last Name"},
    @{n="Email";       d="Email"},
    @{n="Phone";       d="Phone"},
    @{n="Cell";        d="Cell"},
    @{n="Street";      d="Street"},
    @{n="City";        d="City"},
    @{n="State";       d="State/Province"},
    @{n="PostalCode";  d="Postal Code"},
    @{n="Country";     d="Country"}
) | ForEach-Object { Ensure-TextField -InternalName $_.n -DisplayName $_.d }

# --- Neutralize Title (hide/optional/remove from view & CT) ---
Set-PnPField -List $ListTitle -Identity "Title" -Values @{ Required = $false } | Out-Null
Set-PnPField -List $ListTitle -Identity "Title" -Values @{ ShowInNewForm=$false; ShowInEditForm=$false; ShowInDisplayForm=$false } | Out-Null
$defaultView = Get-PnPView -List $ListTitle | Where-Object { $_.DefaultView -eq $true } | Select-Object -First 1
if ($defaultView) { Set-PnPView -List $ListTitle -Identity $defaultView.Id -Fields $contactFields | Out-Null }
try {
    $ctx = Get-PnPContext
    $ctx.Load($list.ContentTypes); $ctx.ExecuteQuery()
    $itemCt = $list.ContentTypes | Where-Object { $_.Name -eq "Item" } | Select-Object -First 1
    if ($itemCt) {
        $ctx.Load($itemCt.FieldLinks); $ctx.ExecuteQuery()
        $titleLink = $itemCt.FieldLinks | Where-Object { $_.Name -eq "Title" } | Select-Object -First 1
        if ($titleLink) {
            $titleLink.Required = $false; $titleLink.Hidden = $true
            $itemCt.Update($true); $ctx.ExecuteQuery()
        }
    }
} catch { Write-Warning "Content type tweak failed: $($_.Exception.Message)" }

# --- Fetch RandomUser data ---
$api = "https://randomuser.me/api/?results=$Results&inc=name,email,phone,cell,location"
Write-Host "Fetching $Results contacts from RandomUser..."
$resp = Invoke-RestMethod -Uri $api -Method Get -TimeoutSec 30
$users = $resp.results
if (-not $users) { throw "RandomUser API returned no results." }

# --- Build email index for upsert ---
$existingItems = Get-PnPListItem -List $ListTitle -PageSize 2000 -Fields "ID","Email"
$existingByEmail = @{}
foreach ($it in $existingItems) {
    $email = (Get-String $it.FieldValues["Email"]).Trim().ToLower()
    if ($email) { $existingByEmail[$email] = $it }
}

# --- Upsert contacts ---
$added = 0; $updated = 0
foreach ($u in $users) {
    $first = (Get-String $u.name.first).Trim()
    $last  = (Get-String $u.name.last).Trim()
    $email = (Get-String $u.email).Trim().ToLower()
    $phone = (Get-String $u.phone).Trim()
    $cell  = (Get-String $u.cell).Trim()

    $streetNumber = if ($u.location.street.number) { $u.location.street.number.ToString() } else { "" }
    $streetName   = (Get-String $u.location.street.name).Trim()
    $street       = ($streetNumber, $streetName -join " ").Trim()

    $city    = (Get-String $u.location.city).Trim()
    $state   = (Get-String $u.location.state).Trim()
    $country = (Get-String $u.location.country).Trim()
    $postal  = if ($u.location.postcode) { $u.location.postcode.ToString() } else { "" }

    $values = @{
        "FirstName"  = $first
        "LastName"   = $last
        "Email"      = $email
        "Phone"      = $phone
        "Cell"       = $cell
        "Street"     = $street
        "City"       = $city
        "State"      = $state
        "PostalCode" = $postal
        "Country"    = $country
    }

    if ($email -and $existingByEmail.ContainsKey($email)) {
        Set-PnPListItem -List $ListTitle -Identity $existingByEmail[$email].Id -Values $values | Out-Null
        $updated++
    } else {
        $newItem = Add-PnPListItem -List $ListTitle -Values $values
        if ($email) { $existingByEmail[$email] = $newItem }
        $added++
    }
}

# --- Build site-scoped canonical link robustly (avoid dupes/missing leaf) ---
$web  = Get-PnPWeb
$null = Get-PnPProperty -ClientObject $web  -Property Url, ServerRelativeUrl
$null = Get-PnPProperty -ClientObject $list -Property Id, Title, RootFolder, DefaultView, DefaultViewUrl

# Origin (scheme+host only) to safely join with server-relative paths
$uri         = [Uri]$web.Url
$origin      = '{0}://{1}' -f $uri.Scheme, $uri.Host           # e.g. https://wkx8x.sharepoint.com
$siteBase    = $web.Url.TrimEnd('/')                           # e.g. https://wkx8x.sharepoint.com/sites/wlive
$settingsUrl = "$siteBase/_layouts/15/listedit.aspx?List={$($list.Id.Guid)}"

# Ensure RootFolder fully loaded
try { $ctx = Get-PnPContext; $ctx.Load($list.RootFolder); $ctx.ExecuteQuery() } catch {}

# Build server-relative folder path: /sites/<site>/Lists/<ListUrlName>
$serverRelFolder = $null

if ($list.RootFolder -and $list.RootFolder.ServerRelativeUrl) {
    $serverRelFolder = [string]$list.RootFolder.ServerRelativeUrl
    if (-not $serverRelFolder.StartsWith('/')) { $serverRelFolder = "/$serverRelFolder" }
    $serverRelFolder = $serverRelFolder -replace '/lists/', '/Lists/'

    # If path ends at /Lists/ (missing leaf), append the leaf safely
    if ($serverRelFolder -match '/Lists/?$') {
        $leaf = $null
        try { $leaf = [string]$list.RootFolder.Name } catch {}
        if ([string]::IsNullOrWhiteSpace($leaf)) {
            if ($list.DefaultViewUrl) {
                $dv = [string]$list.DefaultViewUrl
                if (-not $dv.StartsWith('/')) { $dv = "/$dv" }
                $dv = $dv -replace '/lists/', '/Lists/'
                $serverRelFolder = $dv -replace '/[^/]+\.aspx$',''  # parent folder of the view
            } else {
                $leaf = [Regex]::Replace($list.Title, '[^\w\- ]','')
                $leaf = $leaf -replace ' ', ''
                $serverRelFolder = ($serverRelFolder.TrimEnd('/')) + '/' + [System.Uri]::EscapeDataString($leaf)
            }
        } else {
            $serverRelFolder = ($serverRelFolder.TrimEnd('/')) + '/' + [System.Uri]::EscapeDataString($leaf)
        }
    }
}
elseif ($list.DefaultViewUrl) {
    $dv = [string]$list.DefaultViewUrl
    if (-not $dv.StartsWith('/')) { $dv = "/$dv" }
    $dv = $dv -replace '/lists/', '/Lists/'
    $serverRelFolder = $dv -replace '/[^/]+\.aspx$',''
}
else {
    # Absolute fallback (functional but not ideal)
    $encodedTitle = [System.Uri]::EscapeDataString($list.Title)
    $serverRelFolder = "/sites/$($web.ServerRelativeUrl.Trim('/').Split('/')[-1])/Lists/$encodedTitle"
}

# Final canonical link: origin + server-relative folder + /AllItems.aspx
if ($serverRelFolder -match '\.aspx$') {
    $viewUrl = "$origin$serverRelFolder"
} else {
    $viewUrl = "$origin$($serverRelFolder.TrimEnd('/'))/AllItems.aspx"
}

Write-Host ""
Write-Host ("Done. Added: {0}, Updated: {1}" -f $added,$updated) -ForegroundColor Green
Write-Host ("List URL (canonical): {0}" -f $viewUrl) -ForegroundColor Cyan
Write-Host ("Settings URL:         {0}" -f $settingsUrl) -ForegroundColor DarkCyan

if ($OpenInBrowser -and $viewUrl) {
    try { Start-Process $viewUrl } catch { Write-Warning "Could not open browser: $_" }
}
