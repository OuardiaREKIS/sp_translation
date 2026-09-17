# Reparation des pages dont le canvas contient des null invalides sur le
# controle "fond de section" (controlType 14), que la co-edition SharePoint
# ne sait pas relire -> 500 startCoAuth -> "Impossible de modifier cette page".
# Deux formes traitees :
#   1) controldata  : "webPartId":null   -> propriete retiree
#   2) webpartdata  : {"id":null, ...    -> propriete retiree (ancree sur le
#      debut de l'attribut data-sp-webpartdata pour ne toucher que ce bloc)
#
# Usage :
#   .\repair-canvas.ps1           -> DIAGNOSTIC SEUL (aucune ecriture)
#   .\repair-canvas.ps1 -Apply    -> repare (backup horodate dans .\canvas-backup\)

param(
    [switch]$Apply,
    [int]$Id = 0,
    [string]$SiteUrl = "https://m365x62999571.sharepoint.com/sites/IntranetMOSAIC",
    [string]$ClientId = "<CLIENT-ID>",
    [string]$Tenant = "m365x62999571.onmicrosoft.com",
    [string]$Thumbprint = "<THUMBPRINT>"
)

$env:PSModulePath = "$HOME\Documents\PowerShell\Modules;$HOME\Documents\WindowsPowerShell\Modules;$env:PSModulePath"
Import-Module PnP.PowerShell -ErrorAction Stop
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant -ErrorAction Stop

$backupDir = Join-Path $PSScriptRoot "canvas-backup"
if ($Apply -and -not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }

# Paires (chercher -> remplacer), avec les variantes d'encodage HTML du canvas
$repairs = @(
    @{ find = '&quot;webPartId&quot;&#58;null,'; repl = '' },
    @{ find = '&quot;webPartId&quot;:null,';     repl = '' },
    @{ find = '"webPartId":null,';               repl = '' },
    @{ find = 'data-sp-webpartdata="&#123;&quot;id&quot;&#58;null,'; repl = 'data-sp-webpartdata="&#123;' },
    @{ find = 'data-sp-webpartdata="&#123;&quot;id&quot;:null,';     repl = 'data-sp-webpartdata="&#123;' },
    @{ find = 'data-sp-webpartdata="{"id":null,';                    repl = 'data-sp-webpartdata="{' }
)

$list = Get-PnPList | Where-Object { $_.BaseTemplate -eq 119 } | Select-Object -First 1
$items = Get-PnPListItem -List $list -PageSize 200 -Fields "ID","FileRef","CanvasContent1" |
    Where-Object { $_["FileRef"] -like "*.aspx" } | Sort-Object Id
if ($Id -gt 0) { $items = $items | Where-Object Id -eq $Id }

$nb = 0
foreach ($it in $items) {
    $name = $it["FileRef"] -replace ".*/SitePages/", ""
    $raw = [string]$it["CanvasContent1"]
    if (-not $raw) { continue }
    $hits = $repairs | Where-Object { $raw.Contains($_.find) }
    if (-not $hits) { continue }
    $nb++

    if (-not $Apply) {
        Write-Host "[A REPARER] id=$($it.Id)  $name  ($($hits.Count) correction(s))" -ForegroundColor Yellow
        continue
    }

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    Set-Content -LiteralPath (Join-Path $backupDir "page-$($it.Id)-$stamp.canvas.bak.html") -Value $raw -NoNewline -Encoding utf8
    $fixed = $raw
    foreach ($r in $repairs) { $fixed = $fixed.Replace($r.find, $r.repl) }
    Set-PnPListItem -List $list -Identity $it.Id -Values @{ CanvasContent1 = $fixed } -UpdateType SystemUpdate | Out-Null

    # Verification SANS startCoAuth (cet appel ouvrait une session de
    # co-edition qui verrouillait la page ~1h et bloquait sa suppression) :
    # on relit l'element et on confirme qu'aucun token invalide ne reste.
    $check = [string](Get-PnPListItem -List $list -Id $it.Id -Fields "CanvasContent1")["CanvasContent1"]
    $reste = $repairs | Where-Object { $check.Contains($_.find) }
    if (-not $reste) { Write-Host "[REPARE OK] id=$($it.Id)  $name" -ForegroundColor Green }
    else { Write-Host "[TOUJOURS KO] id=$($it.Id)  $name (token invalide restant)" -ForegroundColor Red }
}

if ($nb -eq 0) { Write-Host "Aucune page a reparer." -ForegroundColor Green }
elseif (-not $Apply) { Write-Host ""; Write-Host "$nb page(s) a reparer -> relance avec -Apply" -ForegroundColor Cyan }
else { Write-Host ""; Write-Host "$nb page(s) traitees. Backups dans $backupDir" -ForegroundColor Cyan }

Disconnect-PnPOnline
