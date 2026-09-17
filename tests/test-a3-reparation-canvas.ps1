# A3 - Les regles de reparation, extraites de repair-canvas.ps1, appliquees a
# des fixtures : le canvas casse est repare, le sain reste intact.
$repPath = Join-Path $PSScriptRoot "..\repair-canvas.ps1"
$rep = Get-Content -Raw $repPath
if ($rep -notmatch '(?s)(\$repairs = @\(.*?\n\))') { Write-Host "[FAIL] bloc repairs introuvable" -ForegroundColor Red; return 1 }
Invoke-Expression $Matches[1]

$fixtureCassee = 'data-sp-controldata="&#123;&quot;webPartId&quot;&#58;null,&quot;rteInstanceId&quot;&#58;null,&quot;controlType&quot;&#58;14&#125;"><div data-sp-webpartdata="&#123;&quot;id&quot;&#58;null,&quot;instanceId&quot;&#58;&quot;abc&quot;&#125;">'
$fixtureSaine  = 'data-sp-controldata="&#123;&quot;controlType&quot;&#58;4,&quot;rteInstanceId&quot;&#58;null&#125;"><div data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;d1d91016-032f-456d-98a4-721247c305e8&quot;&#125;">'

$fails = 0
$reparee = $fixtureCassee
foreach ($r in $repairs) { $reparee = $reparee.Replace($r.find, $r.repl) }
$webPartIdRetire = -not $reparee.Contains('webPartId')
$idRetire = -not $reparee.Contains('&quot;id&quot;&#58;null')
if ($webPartIdRetire -and $idRetire) { Write-Host "[OK]   canvas casse -> les 2 nulls retires" }
else { Write-Host "[FAIL] canvas casse : webPartId retire=$webPartIdRetire, id retire=$idRetire" -ForegroundColor Red; $fails++ }
if ($reparee.Contains('&quot;rteInstanceId&quot;&#58;null')) { Write-Host "[OK]   null legitime (rteInstanceId) preserve" }
else { Write-Host "[FAIL] null legitime (rteInstanceId) supprime a tort" -ForegroundColor Red; $fails++ }
if ($reparee.Contains('&quot;instanceId&quot;&#58;&quot;abc&quot;')) { Write-Host "[OK]   instanceId du webpartdata preserve" }
else { Write-Host "[FAIL] instanceId perdu" -ForegroundColor Red; $fails++ }

$saineApres = $fixtureSaine
foreach ($r in $repairs) { $saineApres = $saineApres.Replace($r.find, $r.repl) }
if ($saineApres -eq $fixtureSaine) { Write-Host "[OK]   canvas sain -> inchange octet pour octet" }
else { Write-Host "[FAIL] canvas sain modifie a tort" -ForegroundColor Red; $fails++ }
return $fails
