# Lance tous les tests du dossier tests/ et resume les resultats.
# Usage : .\tests\run-tests.ps1
$ErrorActionPreference = "Continue"
$files = Get-ChildItem -Path $PSScriptRoot -Filter "test-*.ps1" | Sort-Object Name
$totalFail = 0
foreach ($f in $files) {
    Write-Host ""
    Write-Host "=== $($f.Name) ===" -ForegroundColor Cyan
    $fails = & $f.FullName
    $totalFail += [int]$fails
}
Write-Host ""
if ($totalFail -eq 0) { Write-Host "TOUS LES TESTS PASSENT" -ForegroundColor Green; exit 0 }
else { Write-Host "$totalFail ECHEC(S)" -ForegroundColor Red; exit 1 }
