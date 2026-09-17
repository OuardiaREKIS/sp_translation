# B1 - Tous les .ps1 du depot doivent parser sans erreur.
$root = Split-Path -Parent $PSScriptRoot
$fails = 0
Get-ChildItem -Path $root -Filter "*.ps1" -Recurse | Where-Object { $_.FullName -notmatch 'node_modules' } | ForEach-Object {
    $t = $null; $e = $null
    [System.Management.Automation.Language.Parser]::ParseFile($_.FullName, [ref]$t, [ref]$e) | Out-Null
    if ($e.Count -eq 0) { Write-Host "[OK]   $($_.FullName.Substring($root.Length + 1))" }
    else { Write-Host "[FAIL] $($_.Name) : $($e.Count) erreur(s)" -ForegroundColor Red; $fails++ }
}
return $fails
