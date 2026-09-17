# B2 - Aucun secret dans les fichiers suivis par git : cle Azure (signature
# JQQJ99), cle de fonction (?code=) hors gabarits/doc historique.
$root = Split-Path -Parent $PSScriptRoot
$fails = 0
$suivis = git -C $root ls-files
foreach ($f in $suivis) {
    $full = Join-Path $root $f
    if (-not (Test-Path $full)) { continue }
    $contenu = Get-Content -Raw $full -ErrorAction SilentlyContinue
    if (-not $contenu) { continue }
    if ($contenu -match '[A-Za-z0-9]{20,}JQQJ99[A-Za-z0-9]+') { Write-Host "[FAIL] cle Translator dans $f" -ForegroundColor Red; $fails++ }
    if ($f -notmatch '\.example$|\.md$' -and $contenu -match '\?code=[A-Za-z0-9_\-]{20,}') { Write-Host "[FAIL] cle de fonction dans $f" -ForegroundColor Red; $fails++ }
}
if ($fails -eq 0) { Write-Host "[OK]   aucun secret dans les $(($suivis | Measure-Object).Count) fichiers suivis" }
return $fails
