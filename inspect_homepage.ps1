# Recupere tout le contenu d'une page SharePoint (tenant demo m365x62999571).
#
# Usage :
#   .\inspect_homepage.ps1                              -> TopicHome.aspx
#   .\inspect_homepage.ps1 -Page "en/TopicHome.aspx"    -> la copie anglaise
#   .\inspect_homepage.ps1 -Page "autre-page.aspx"

param(
    [string]$Page = "TopicHome.aspx",
    [string]$SiteUrl = "https://m365x62999571.sharepoint.com/sites/IntranetMOSAIC",
    [string]$ClientId = $env:SP_CLIENT_ID,
    [string]$Tenant = "m365x62999571.onmicrosoft.com",
    [string]$Thumbprint = $env:SP_THUMBPRINT
)

# PnP.PowerShell est installe dans le dossier de modules de Windows PowerShell 5.1,
# que pwsh ne scanne pas par defaut : on l'ajoute au chemin de recherche.
$env:PSModulePath = "$HOME\Documents\PowerShell\Modules;$HOME\Documents\WindowsPowerShell\Modules;$env:PSModulePath"
Import-Module PnP.PowerShell -ErrorAction Stop

Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant -ErrorAction Stop

$p = Get-PnPPage -Identity $Page

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host " PAGE : $Page" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host " Titre  : $($p.PageTitle)"
Write-Host " Layout : $($p.LayoutType)"

foreach ($c in $p.Controls) {
    Write-Host "------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " Bloc      : $($c.Type.Name)  $($c.Title)" -ForegroundColor Yellow
    Write-Host " WebPartId : $($c.WebPartId)"
    if ($c.Type.Name -eq "PageText") {
        Write-Host " Texte     : $($c.Text)"
    }
    elseif ($c.Type.Name -eq "PageWebPart") {
        Write-Host " ServerProcessedContent :"
        Write-Host "   $($c.ServerProcessedContent.ToString())"
        Write-Host " PropertiesJson :"
        Write-Host "   $($c.PropertiesJson.ToString())"
    }
}

Disconnect-PnPOnline
