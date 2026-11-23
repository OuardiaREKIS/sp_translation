# Outil de diagnostic : liste tous les composants (webparts) d'une page SharePoint
# et leurs textes, pour voir ce que la fonction de traduction doit couvrir.
#
# Usage :
#   .\inspect-page.ps1                          -> inspecte test-mosaic.aspx
#   .\inspect-page.ps1 -Page "autre-page.aspx"  -> inspecte une autre page
#   .\inspect-page.ps1 -Copies                  -> affiche aussi le titre des copies traduites

param(
    [string]$Page = "test-mosaic.aspx",
    [string]$SiteUrl = "https://3whd6t.sharepoint.com",
    [switch]$Copies
)

# PnP.PowerShell est installe dans le dossier de modules de Windows PowerShell 5.1,
# que pwsh ne scanne pas par defaut : on l'ajoute au chemin de recherche.
$env:PSModulePath = "$HOME\Documents\PowerShell\Modules;$HOME\Documents\WindowsPowerShell\Modules;$env:PSModulePath"
Import-Module PnP.PowerShell -ErrorAction Stop

Connect-PnPOnline -Url $SiteUrl `
    -ClientId $env:SP_CLIENT_ID `
    -Thumbprint $env:SP_THUMBPRINT `
    -Tenant "3whd6t.onmicrosoft.com" -ErrorAction Stop

$p = Get-PnPPage -Identity $Page

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host " PAGE : $Page" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host " Titre        : $($p.PageTitle)"
Write-Host " Topic header : $($p.PageHeader.TopicHeader)"
Write-Host ""

foreach ($c in $p.Controls) {
    Write-Host "------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " Composant : $($c.Type.Name)  $(if ($c.Title) { "($($c.Title))" })" -ForegroundColor Yellow
    if ($c.WebPartId) { Write-Host " WebPartId : $($c.WebPartId)" }

    switch ($c.Type.Name) {
        "PageText" {
            Write-Host " Texte :"
            Write-Host "   $($c.Text)"
        }
        "PageWebPart" {
            Write-Host " searchablePlainTexts :"
            Write-Host "   $($c.ServerProcessedContent)"
            Write-Host " PropertiesJson :"
            Write-Host "   $($c.PropertiesJson)"
        }
        default {
            Write-Host " (type non inspecte)"
        }
    }
}

if ($Copies) {
    Write-Host ""
    Write-Host "=== COPIES TRADUITES ===" -ForegroundColor Cyan
    foreach ($dossier in @("fr", "es", "de", "zh-chs", "tr", "ar", "ja")) {
        try {
            $copie = Get-PnPPage -Identity "$dossier/$Page" -ErrorAction Stop
            Write-Host " [$dossier] titre : $($copie.PageTitle)"
        }
        catch {
            Write-Host " [$dossier] absente" -ForegroundColor DarkGray
        }
    }
}

Disconnect-PnPOnline
