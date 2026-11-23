using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "PowerShell HTTP trigger function processed a request."
# Marqueur de build : visible dans la reponse 400 (verification de la version deployee)
$buildVersion = "build-2026-09-27-perf-1"
Write-Host $buildVersion

# ============================================================================
# ORCHESTRATEUR : ce point d'entree ne traduit plus lui-meme. Il valide,
# cree les copies manquantes EN UNE FOIS (plus de course entre appels),
# puis enfile UN JOB PAR LANGUE dans la file 'translation-jobs' que le
# worker TranslateWorker consomme en parallele (scale-out par la file).
# La reponse revient en quelques secondes ; le client suit l'avancement
# via la colonne MosaicTranslationStatus des copies.
# ============================================================================

. "$PSScriptRoot/../shared/translate-common.ps1"

$languages = $Request.Body.languages
$pages = $Request.Body.pages
$siteUrl = $Request.Body.siteUrl

Write-Host "Languages are : $languages"
Write-Host "Pages are : $pages"
Write-Host "The site Url is : $siteUrl"

function Send-Response {
    param(
        [Parameter(Mandatory = $true)][string]$code,
        [Parameter(Mandatory = $true)][string]$body
    )
    Write-Host "Response StatusCode is : $([HttpStatusCode]::$code)"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::$code
            Headers    = @{
                "Content-type" = "application/json"
                "Accept"       = "application/json"
            }
            Body       = $body
        })
    exit
}

if (-not $languages -or -not $pages -or -not $siteUrl) {
    Send-Response -code BadRequest -body "Bad Request ($buildVersion)"
}

$languageIds = @()
$languageCodes = @()
foreach ($language in $languages) {
    $languageIds += $language.id
    # ATTENTION au sens des champs : sp = nom COURT du dossier de langue
    # ("fr", "de", "zh-chs"), PAS un code culture. az = code Azure Translator.
    # Les codes culture pour l'API translations/create ("fr-fr", "zh-cn")
    # sont derives du LCID ($language.id), jamais de sp.
    $languageCodes += [PSCustomObject]@{
        sp = $language.code
        az = $language.azCode
        id = $language.id
    }
}

# Connexion app-only par certificat.
try {
    if ($env:SP_CERT_PATH) {
        $certPassword = ConvertTo-SecureString $env:SP_CERT_PASSWORD -AsPlainText -Force
        Connect-PnPOnline -Url $siteUrl -ClientId $env:SP_CLIENT_ID -Tenant $env:SP_TENANT -CertificatePath $env:SP_CERT_PATH -CertificatePassword $certPassword -ErrorAction Stop
    }
    else {
        Connect-PnPOnline -Url $siteUrl -ClientId $env:SP_CLIENT_ID -Tenant $env:SP_TENANT -Thumbprint $env:SP_THUMBPRINT -ErrorAction Stop
    }
}
catch {
    Send-Response -code InternalServerError -body "SharePoint connection failed : $_"
}

$allJobs = @()
$resume = @()
$skipped = @()

foreach ($pagePath in $pages) {
    Write-Host "Source page : $pagePath"
    $sitePath = ([uri]$siteUrl).AbsolutePath.TrimEnd('/')

    # La source doit exister A LA RACINE de la bibliotheque : sinon Get-PnPPage
    # rabat silencieusement sur une copie du meme nom (cas d'une source
    # supprimee) et tout part en vrille -> refus net et explicite.
    try {
        $null = Get-PnPFile -Url "$sitePath/SitePages/$pagePath" -ErrorAction Stop
    }
    catch {
        Disconnect-PnPOnline
        Send-Response -code NotFound -body $(@{ msg = "Page source introuvable a la racine : $pagePath (supprimee ou deplacee ?)" } | ConvertTo-Json)
    }

    $originalPage = Get-PnPPage -Identity $pagePath
    [bool]$commentsDisabled = $originalPage.AreCommentsDisabled()
    $sourceTitle = [string]$originalPage.PageTitle
    $topicHeader = [string]$originalPage.PageHeader.TopicHeader

    $createLanguageIds = @()
    $pageJobLanguages = @()

    foreach ($lc in $languageCodes) {
        $copyPath = "$($lc.sp)/$pagePath"
        $pageCheck = $null
        try { $pageCheck = Get-PnPPage -Identity $copyPath -ErrorAction Stop } catch { $pageCheck = $null }

        if (-not $pageCheck) {
            # Copie inexistante : langue a CREER puis a traduire. Les langues
            # en reprise ne doivent JAMAIS figurer dans translations/create :
            # l'API refuse une traduction existante et ferait echouer tout
            # l'appel, langues neuves comprises.
            $createLanguageIds += $lc.id
            $pageJobLanguages += $lc
            continue
        }

        # Etat de traitement explicite (colonne MosaicTranslationStatus).
        $copyStatus = $null
        try {
            $statusItem = Get-PnPFile -Url "$sitePath/SitePages/$copyPath" -AsListItem -ErrorAction Stop
            $copyStatus = [string]$statusItem["MosaicTranslationStatus"]
        } catch {}
        if ($copyStatus -eq "Pending" -or $copyStatus -eq "Failed") {
            Write-Host "Reprise (statut $copyStatus) : $copyPath"
            $pageJobLanguages += $lc
            $resume += "$($lc.sp)/$pagePath"
            continue
        }
        if ($copyStatus -eq "Completed") {
            $skipped += "$($lc.sp)/$pagePath"
            continue
        }

        # Copie historique sans statut : heuristiques de titre.
        # 1) structurelle (toutes langues) : titre = bapteme se terminant par
        #    le titre SOURCE intact ; 2) gabarits localises.
        $copyTitle = [string]$pageCheck.PageTitle
        $neverTranslated = $false
        if ($copyTitle -and $sourceTitle -and $copyTitle -ne $sourceTitle -and $copyTitle.EndsWith($sourceTitle)) {
            $neverTranslated = $true
        }
        $prefixPattern = '^(Translate to [^:]{1,40}:|(Traduisez|Traduire)[' + [char]0x00A0 + ' ](en|dans)[' + [char]0x00A0 + ' ][^:]{1,40}[' + [char]0x00A0 + ' ]?:|' + [char]0x00DC + 'bersetzen (nach|in|auf) [^:]{1,40}:|In [^:]{1,40} ' + [char]0x00FC + 'bersetzen:|Traducir al? [^:]{1,40}:|(' + [char]0x041F + [char]0x0435 + [char]0x0440 + [char]0x0435 + [char]0x0432 + [char]0x0435 + [char]0x0441 + [char]0x0442 + [char]0x0438 + '|' + [char]0x041F + [char]0x0435 + [char]0x0440 + [char]0x0435 + [char]0x0432 + [char]0x043E + [char]0x0434 + ') ' + [char]0x043D + [char]0x0430 + ' [^:]{1,40}:|[^:]{0,40}' + [char]0x00E7 + 'evir(in)?:|' + [char]0x7FFB + [char]0x8BD1 + '[' + [char]0x6210 + [char]0x4E3A + '][^:' + [char]0xFF1A + ']{1,20}[:' + [char]0xFF1A + ']|' + [char]0x062A + [char]0x0631 + [char]0x062C + [char]0x0645 + [char]0x0629 + ' ' + [char]0x0625 + [char]0x0644 + [char]0x0649 + ' [^:]{1,40}:)'
        if (-not $neverTranslated -and $copyTitle -match $prefixPattern) {
            $neverTranslated = $true
        }
        if ($neverTranslated) {
            Write-Host "Copie jamais traduite ('$copyTitle') : reprise de $copyPath"
            $pageJobLanguages += $lc
            $resume += "$($lc.sp)/$pagePath"
        }
        else {
            $skipped += "$($lc.sp)/$pagePath"
        }
    }

    if ($createLanguageIds.Count -gt 0) {
        # translations/create attend des codes culture complets ("fr-fr",
        # "de-de", "zh-cn"...), derives du LCID. UNE SEULE creation pour
        # toutes les langues manquantes de la page : plus aucune course.
        $codesSP = ($createLanguageIds | ForEach-Object { '"' + ([System.Globalization.CultureInfo]::new([int]$_)).Name.ToLower() + '"' }) -join ","
        $translateBody = '{ "request": { "__metadata": { "type": "SP.TranslationStatusCreationRequest" }, "LanguageCodes": { "results": [ ' + $codesSP + ' ] } } }'
        $created = $false
        for ($essai = 1; $essai -le 5 -and -not $created; $essai++) {
            try {
                Invoke-PnPSPRestMethod -Method Post -Url "/_api/sitepages/pages($($originalPage.PageId))/translations/create" -Content $translateBody -ContentType "application/json;odata=verbose" -ErrorAction Stop | Out-Null
                $created = $true
            }
            catch {
                if ($essai -lt 5) {
                    Write-Warning "translations/create essai $essai rate, nouvelle tentative dans $($essai * 5) s : $_"
                    Start-Sleep -Seconds ($essai * 5)
                }
                else {
                    Disconnect-PnPOnline
                    Send-Response -code InternalServerError -body $(@{ msg = "Creation des copies impossible pour '$pagePath' apres 5 essais (langues activees sur le site ?) : $_" } | ConvertTo-Json)
                }
            }
        }
        # La creation a re-sauve des elements : nettoyage defensif de la source
        Repair-PageCanvas -SiteUrl $siteUrl -PageRelativePath $pagePath
    }

    foreach ($lc in $pageJobLanguages) {
        $allJobs += (@{
                siteUrl          = $siteUrl
                page             = $pagePath
                sourceTitle      = $sourceTitle
                topicHeader      = $topicHeader
                commentsDisabled = $commentsDisabled
                language         = @{ code = $lc.sp; azCode = $lc.az; id = $lc.id }
            } | ConvertTo-Json -Depth 5 -Compress)
    }
}

Disconnect-PnPOnline

if ($allJobs.Count -gt 0) {
    Push-OutputBinding -Name jobs -Value $allJobs
}

$detail = @()
if ($allJobs.Count -gt 0) { $detail += "$($allJobs.Count) traduction(s) en file" }
if ($resume.Count -gt 0) { $detail += "reprises : $($resume -join ', ')" }
if ($skipped.Count -gt 0) { $detail += "deja terminees (ignorees) : $($skipped -join ', ')" }
if ($detail.Count -eq 0) { $detail += "rien a faire" }

Send-Response -code OK -body $(@{
        msg  = ($detail -join ' | ')
        jobs = $allJobs.Count
    } | ConvertTo-Json)
