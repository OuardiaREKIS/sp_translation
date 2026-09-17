# A2 - La derivation LCID -> code culture, extraite du run.ps1 de production.
$runPath = Join-Path $PSScriptRoot "..\AzureFunction2\SP-Translator-FunctionApp\HttpTriggerSpTranslator\run.ps1"
$run = (Get-Content -Raw $runPath) -replace "`r`n", "`n"
$line = ($run -split "`n" | Where-Object { $_ -match 'CultureInfo' }) | Select-Object -First 1
if (-not $line) { Write-Host "[FAIL] ligne CultureInfo introuvable" -ForegroundColor Red; return 1 }

$attendus = [ordered]@{
    1036 = "fr-fr"; 1033 = "en-us"; 1031 = "de-de"; 3082 = "es-es"; 1040 = "it-it"
    1055 = "tr-tr"; 1025 = "ar-sa"; 1049 = "ru-ru"; 2052 = "zh-cn"; 1028 = "zh-tw"
    1041 = "ja-jp"; 1046 = "pt-br"; 2070 = "pt-pt"; 1045 = "pl-pl"; 1043 = "nl-nl"
}
$pageLanguageIds = @($attendus.Keys)
Invoke-Expression $line.Trim()
$obtenus = $codesSP -split "," | ForEach-Object { $_.Trim('"') }
$fails = 0
$i = 0
foreach ($paire in $attendus.GetEnumerator()) {
    if ($obtenus[$i] -eq $paire.Value) { Write-Host "[OK]   $($paire.Key) -> $($obtenus[$i])" }
    else { Write-Host "[FAIL] $($paire.Key) attendu=$($paire.Value) obtenu=$($obtenus[$i])" -ForegroundColor Red; $fails++ }
    $i++
}
return $fails
