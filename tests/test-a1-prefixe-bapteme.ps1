# A1 - Le motif de detection des titres de bapteme SharePoint, extrait du
# run.ps1 de production (pas de duplication : si le code change, le test suit).
$runPath = Join-Path $PSScriptRoot "..\AzureFunction2\SP-Translator-FunctionApp\HttpTriggerSpTranslator\run.ps1"
$run = (Get-Content -Raw $runPath) -replace "`r`n", "`n"
$line = ($run -split "`n" | Where-Object { $_ -match '^\s*\$prefixPattern = ' }) -join ''
if (-not $line) { Write-Host "[FAIL] prefixPattern introuvable dans run.ps1" -ForegroundColor Red; return 1 }
Invoke-Expression $line.Trim()

$cas = [ordered]@{
    "Translate to French: Winter is coming"        = $true
    "Traduisez en allemand : Bourse du president"  = $true
    "Traduire en anglais : Note de rentree"        = $true
    ([string][char]0x00DC + "bersetzen nach Deutsch: Hallo")        = $true
    ("Almancaya " + [char]0x00E7 + "evir: Merhaba")                 = $true
    ([char]0x7FFB + [char]0x8BD1 + [char]0x6210 + [char]0x4E2D + [char]0x6587 + [char]0xFF1A + "Bonjour") = $true
    "Traduire dans Arabe: lemnaHealthOR2"          = $true
    ("Traduire dans" + [char]0x00A0 + "Anglais: Page") = $true
    "Guide de traduction : outils du service"      = $false
    "Bourse du president"                          = $false
    "Traduction automatique lancee"                = $false
}
$fails = 0
foreach ($k in $cas.Keys) {
    $obtenu = [bool]($k -match $prefixPattern)
    if ($obtenu -eq $cas[$k]) { Write-Host "[OK]   '$k'" }
    else { Write-Host "[FAIL] '$k' attendu=$($cas[$k]) obtenu=$obtenu" -ForegroundColor Red; $fails++ }
}
return $fails
