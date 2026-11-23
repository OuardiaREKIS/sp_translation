# Diagnostic verrous/extractions des pages du site (tenant demo).
# A lancer AU MOMENT ou quelqu'un est bloque, pour voir qui tient le verrou.
#
# Usage :
#   .\check-page-locks.ps1                 -> toutes les copies traduites
#   .\check-page-locks.ps1 -All            -> toutes les pages du site

param(
    [switch]$All,
    [string]$SiteUrl = "https://m365x62999571.sharepoint.com/sites/IntranetMOSAIC",
    [string]$ClientId = $env:SP_CLIENT_ID,
    [string]$Tenant = "m365x62999571.onmicrosoft.com",
    [string]$Thumbprint = $env:SP_THUMBPRINT
)

$env:PSModulePath = "$HOME\Documents\PowerShell\Modules;$HOME\Documents\WindowsPowerShell\Modules;$env:PSModulePath"
Import-Module PnP.PowerShell -ErrorAction Stop
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant -ErrorAction Stop

$list = Get-PnPList | Where-Object { $_.BaseTemplate -eq 119 } | Select-Object -First 1
$items = Get-PnPListItem -List $list -PageSize 200 -Fields "FileRef" |
    Where-Object { $_["FileRef"] -like "*.aspx" }
if (-not $All) {
    $items = $items | Where-Object { $_["FileRef"] -match "/SitePages/[a-z]{2}(-[a-z]+)?/" }
}

foreach ($it in $items) {
    $url = $it["FileRef"]
    try {
        $f = Get-PnPFile -Url $url -AsFileObject -ErrorAction Stop
        Get-PnPProperty -ClientObject $f -Property CheckOutType, Level | Out-Null
        $lock = "libre"
        try {
            $lu = Get-PnPProperty -ClientObject $f -Property LockedByUser
            if ($lu -and -not $lu.ServerObjectIsNull) { $lock = "VERROU : $($lu.LoginName)" }
        } catch {}
        if ($f.CheckOutType -ne "None") { $lock = "EXTRAITE (CheckOut=$($f.CheckOutType)) ; $lock" }
        $color = if ($lock -eq "libre") { "Gray" } else { "Red" }
        Write-Host ("{0,-55} {1,-10} {2}" -f ($url -replace ".*SitePages/", ""), $f.Level, $lock) -ForegroundColor $color
    }
    catch { Write-Host ("{0,-55} ERREUR : {1}" -f ($url -replace ".*SitePages/", ""), $_.Exception.Message) -ForegroundColor Yellow }
}

Disconnect-PnPOnline
