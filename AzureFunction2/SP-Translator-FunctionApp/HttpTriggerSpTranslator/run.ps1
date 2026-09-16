using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "PowerShell HTTP trigger function processed a request."

$languages = $Request.Body.languages
$pages = $Request.Body.pages
$siteUrl = $Request.Body.siteUrl

Write-Host "Languages are : $languages"
Write-Host "Pages are : $pages"
Write-Host "The site Url is : $siteUrl"

$stopwatch = [system.diagnostics.stopwatch]::StartNew()
$failedValidations = @()

function Send-Response {
    param(
        [Parameter(Mandatory = $true)]
        [string]$code,
        [Parameter(Mandatory = $true)]
        [string]$body
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

# Interact with query parameters or the body of the request.
if (-not $languages -or -not $pages -or -not $siteUrl) {
    Send-Response -code BadRequest -body "Bad Request"
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
    }
}

function Start-Translation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [Parameter(Mandatory = $true)]
        [string]$Language,
        [Parameter(Mandatory = $true)]
        [string]$Service
    )
    
    if ($Service -eq "microsoftTranslator") {
        $baseUri = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0"
        $headers = @{
            'Ocp-Apim-Subscription-Key'    = "$($env:TRANSLATOR_KEY)"
            'Ocp-Apim-Subscription-Region' = "$($env:TRANSLATOR_REGION)"
        }
        # Uri for the request includes language code and text type, which is always html for SharePoint web parts
        $uri = "$baseUri&to=$Language&textType=html"
    }

    # Send request for translation and extract translated text
    $results = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $Text -ContentType "application/json; charset=utf-8"
    return $results
}

# Connexion app-only par certificat.
# En local : SP_CERT_PATH + SP_CERT_PASSWORD (fichier .pfx).
# Sur Azure : SP_THUMBPRINT (et WEBSITE_LOAD_CERTIFICATES = empreinte, pour charger le certificat).
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

Write-Host "After Connection : $($stopwatch.Elapsed.TotalSeconds)"

# Garde anti-corruption : la serialisation PnP recrache le controle
# "fond de section" (controlType 14) avec "webPartId":null, que la
# co-edition SharePoint ne sait pas relire (startCoAuth 500 -> page
# non modifiable). On retire ce null du canvas apres chaque sauvegarde.
function Repair-PageCanvas {
    param([Parameter(Mandatory = $true)][string]$PageRelativePath)
    try {
        $sitePath = ([uri]$siteUrl).AbsolutePath.TrimEnd('/')
        $item = Get-PnPFile -Url "$sitePath/SitePages/$PageRelativePath" -AsListItem -ErrorAction Stop
        $canvas = [string]$item["CanvasContent1"]
        if (-not $canvas) { return }
        $fixed = $canvas.Replace('&quot;webPartId&quot;&#58;null,', '').Replace('&quot;webPartId&quot;:null,', '').Replace('"webPartId":null,', '')
        # 2e forme : le sous-bloc webpartdata du fond de section commence par {"id":null,
        $fixed = $fixed.Replace('data-sp-webpartdata="&#123;&quot;id&quot;&#58;null,', 'data-sp-webpartdata="&#123;').Replace('data-sp-webpartdata="&#123;&quot;id&quot;:null,', 'data-sp-webpartdata="&#123;').Replace('data-sp-webpartdata="{"id":null,', 'data-sp-webpartdata="{')
        if ($fixed -ne $canvas) {
            $item["CanvasContent1"] = $fixed
            $item.SystemUpdate()
            Invoke-PnPQuery
            Write-Host "Canvas repare (webPartId:null retire) : $PageRelativePath"
        }
    }
    catch { Write-Warning "Reparation canvas impossible sur $PageRelativePath : $_" }
}


foreach ($pagePath in $pages) { 
    [System.Collections.ArrayList]$pageLanguageIds = $languageIds
    [System.Collections.ArrayList]$pageLanguageCodes = $languageCodes
    Write-Host "Source page : $pagePath"
    # Check if comments were enabled on the original page
    $originalPage = Get-PnPPage -Identity $pagePath
    [bool]$commentsDisabled = $originalPage.AreCommentsDisabled()

    # Check if a translated page already exist and don't translate it in the language
    for ($i = 0; $i -lt $languageIds.Count; $i++) {
        Write-Host "$($languageCodes[$i].sp)/$pagePath"
        $pageCheck = $null
        try { $pageCheck = Get-PnPPage -Identity "$($languageCodes[$i].sp)/$pagePath" -ErrorAction Stop } catch { $pageCheck = $null }
        if (-not($pageCheck)) {
            continue
        }
        # --- Etat de traitement explicite (colonne MosaicTranslationStatus) ---
        # Pending / Failed = traduction inachevee -> on reprend cette langue.
        # Completed = traduction terminee -> on saute (protege la relecture).
        # Colonne absente ou vide (copies historiques) -> heuristiques de titre.
        $copyStatus = $null
        try {
            $statusItem = Get-PnPFile -Url "$(([uri]$siteUrl).AbsolutePath.TrimEnd('/'))/SitePages/$($languageCodes[$i].sp)/$pagePath" -AsListItem -ErrorAction Stop
            $copyStatus = [string]$statusItem["MosaicTranslationStatus"]
        } catch {}
        if ($copyStatus -eq "Pending" -or $copyStatus -eq "Failed") {
            Write-Warning "Copie existante en statut '$copyStatus' : reprise de la traduction"
            continue
        }
        if ($copyStatus -eq "Completed") {
            Write-Warning "Page Already Exist (statut Completed)"
            $pageLanguageIds.Remove($($languageIds[$i]))
            $pageLanguageCodes.Remove($($languageCodes[$i]))
            continue
        }        # Copie jamais reellement traduite : SharePoint baptise les copies
        # "Translate to X: <titre source>" (prefixe localise dans la langue
        # de l'interface). Si ce titre de bapteme est encore la (execution
        # interrompue, copie heritee, bouton natif), on RETRADUIT au lieu de
        # sauter. Deux detections complementaires :
        #  1) structurelle, independante de la langue : le titre se termine
        #     par le titre SOURCE intact (c'est la forme du bapteme) ;
        #  2) gabarits localises "Traduire en X :" (en, fr, de, es, ru, tr,
        #     zh, ar) pour le cas ou le titre source a change depuis.
        $copyTitle = [string]$pageCheck.PageTitle
        $sourceTitle = [string]$originalPage.PageTitle
        $neverTranslated = $false
        if ($copyTitle -and $sourceTitle -and $copyTitle -ne $sourceTitle -and $copyTitle.EndsWith($sourceTitle)) {
            $neverTranslated = $true
        }
        $prefixPattern = '^(Translate to [^:]{1,40}:|(Traduisez|Traduire) en [^:]{1,40}:|' + [char]0x00DC + 'bersetzen (nach|in|auf) [^:]{1,40}:|In [^:]{1,40} ' + [char]0x00FC + 'bersetzen:|Traducir al? [^:]{1,40}:|(' + [char]0x041F + [char]0x0435 + [char]0x0440 + [char]0x0435 + [char]0x0432 + [char]0x0435 + [char]0x0441 + [char]0x0442 + [char]0x0438 + '|' + [char]0x041F + [char]0x0435 + [char]0x0440 + [char]0x0435 + [char]0x0432 + [char]0x043E + [char]0x0434 + ') ' + [char]0x043D + [char]0x0430 + ' [^:]{1,40}:|[^:]{0,40}' + [char]0x00E7 + 'evir(in)?:|' + [char]0x7FFB + [char]0x8BD1 + '[' + [char]0x6210 + [char]0x4E3A + '][^:' + [char]0xFF1A + ']{1,20}[:' + [char]0xFF1A + ']|' + [char]0x062A + [char]0x0631 + [char]0x062C + [char]0x0645 + [char]0x0629 + ' ' + [char]0x0625 + [char]0x0644 + [char]0x0649 + ' [^:]{1,40}:)'
        if (-not $neverTranslated -and $copyTitle -match $prefixPattern) {
            $neverTranslated = $true
        }
        if ($neverTranslated) {
            Write-Warning "Copie existante jamais traduite ('$copyTitle') : retraduction"
            continue
        }
        Write-Warning "Page Already Exist"
        $pageLanguageIds.Remove($($languageIds[$i]))
        $pageLanguageCodes.Remove($($languageCodes[$i]))
    }
    if ($pageLanguageIds.Count -gt 0) {
        # Creation des copies via l'API REST translations de SharePoint (comme
        # en V1) : contrairement a Set-PnPPage -Translate, cet appel ne re-sauve
        # PAS la page source -> la serialisation PnP ne peut plus corrompre son
        # canvas (cas du fond de section, webPartId/id nulls -> startCoAuth 500).
        # translations/create attend des codes culture complets ("fr-fr",
        # "de-de", "zh-cn"...), PAS les noms courts des dossiers ("fr", "de").
        # On les derive du LCID recu, pour ne pas maintenir de table.
        $codesSP = ($pageLanguageIds | ForEach-Object { '"' + ([System.Globalization.CultureInfo]::new([int]$_)).Name.ToLower() + '"' }) -join ","
        $translateBody = '{ "request": { "__metadata": { "type": "SP.TranslationStatusCreationRequest" }, "LanguageCodes": { "results": [ ' + $codesSP + ' ] } } }'
        try {
            Invoke-PnPSPRestMethod -Method Post -Url "/_api/sitepages/pages($($originalPage.PageId))/translations/create" -Content $translateBody -ContentType "application/json;odata=verbose" -ErrorAction Stop | Out-Null
        }
        catch {
            Send-Response -code InternalServerError -body $(@{ msg = "Creation des copies impossible pour '$pagePath' (langues activees sur le site ?) : $_" } | ConvertTo-Json)
        }
        # Nettoyage defensif de la source (repare l'heritage d'anciennes executions)
        Repair-PageCanvas -PageRelativePath $pagePath
    }
    Write-Host "After pages creation : $($stopwatch.Elapsed.TotalSeconds)"
    foreach ($language in $pageLanguageCodes) {
        $currentPage = "$($language.sp)/$pagePath"
        Write-Host "Current path of the translated page is $currentPage"
        $modernPage = Get-PnPPage -Identity $currentPage
        # Statut Pending des le debut : si la fonction s'interrompt ici ou
        # plus loin, la copie sera automatiquement reprise au prochain passage.
        try {
            $statusItem = Get-PnPFile -Url "$(([uri]$siteUrl).AbsolutePath.TrimEnd('/'))/SitePages/$currentPage" -AsListItem -ErrorAction Stop
            $statusItem["MosaicTranslationStatus"] = "Pending"
            $statusItem.SystemUpdate()
            Invoke-PnPQuery
        } catch { Write-Warning "Statut Pending non pose sur $currentPage (colonne absente ?)" }
        # Traduction du titre depuis la page SOURCE : SharePoint nomme la copie
        # "Translate to <langue>: <titre>", son titre ne doit donc pas servir de source.
        $oldPageTitle = "[$(@{ "Text"=$originalPage.PageTitle } | ConvertTo-Json)]"
        $modernPage.PageTitle = $(Start-Translation -Text $oldPageTitle -Language $language.az -Service microsoftTranslator)[0].translations[0].text
        if ($originalPage.PageHeader.TopicHeader) {
            $oldTopicHeader = "[$(@{ "Text"=$originalPage.PageHeader.TopicHeader } | ConvertTo-Json)]"
            $modernPage.PageHeader.TopicHeader = $(Start-Translation -text $oldTopicHeader -language $language.az -service microsoftTranslator)[0].translations[0].text
        }
        $modernPage.Save()
        $allWebParts = @()
        $allTextsToTranslate = @()
        # Get all webparts content
        foreach ($control in $modernPage.Controls) {
            Write-Host "Current WebPart :"
            Write-Host "Id : $($control.InstanceId)"
            switch ($control.Type.Name) {
                "PageText" {
                    Write-Host "Name : Text"
                    $allWebParts += [PSCustomObject]@{
                        nbTexts        = 1
                        type           = $control.Type.Name
                        id             = $control.InstanceId
                        dataProperties = "<(#Property1#)>"
                    }
                    $allTextsToTranslate += @{ "Text" = $control.text }
                    break
                }
                "PageWebPart" {
                    $newProperties = [PSCustomObject]@{
                        webPartData = [PSCustomObject]@{
                            serverProcessedContent = $(ConvertFrom-Json $control.serverProcessedContent)
                            properties             = $(ConvertFrom-Json $control.propertiesJson)
                        }
                    }
                    $nbPropertiesToTranslate = 0
                    $forceWebPartUpdate = $false
                    $oldProperties = ConvertFrom-Json $control.serverProcessedContent
                    # Passe generique (tous les webparts) : searchablePlainTexts
                    # regroupe les textes visibles/indexables du webpart, on les
                    # traduit tous. Le reste du JSON (links, imageSources,
                    # htmlStrings, IDs, parametres techniques) reste intact.
                    if ($oldProperties.searchablePlainTexts) {
                        foreach ($textProperty in $oldProperties.searchablePlainTexts.PSObject.Properties) {
                            if ($textProperty.Value -is [string] -and $textProperty.Value) {
                                $nbPropertiesToTranslate += 1
                                $newProperties.webPartData.serverProcessedContent.searchablePlainTexts.$($textProperty.Name) = "<(#Property$nbPropertiesToTranslate#)>"
                                $allTextsToTranslate += @{ "Text" = $textProperty.Value }
                            }
                        }
                    }
                    # titleHTML : version HTML riche du titre (celle affichee a
                    # l'ecran), stockee dans properties en DOUBLE du title de
                    # searchablePlainTexts. Le mode HTML du Translator preserve
                    # les balises et classes de mise en forme.
                    $titleHTML = $newProperties.webPartData.properties.titleHTML
                    if ($titleHTML -is [string] -and $titleHTML) {
                        $nbPropertiesToTranslate += 1
                        $newProperties.webPartData.properties.titleHTML = "<(#Property$nbPropertiesToTranslate#)>"
                        $allTextsToTranslate += @{ "Text" = $titleHTML }
                    }
                    # Banniere (Hero) : chaque tuile a aussi son titre riche
                    # dans properties.content[i].titleHTML.
                    if ($newProperties.webPartData.properties.content -is [array]) {
                        foreach ($tile in $newProperties.webPartData.properties.content) {
                            $tileTitleHTML = $tile.titleHTML
                            if ($tileTitleHTML -is [string] -and $tileTitleHTML) {
                                $nbPropertiesToTranslate += 1
                                $tile.titleHTML = "<(#Property$nbPropertiesToTranslate#)>"
                                $allTextsToTranslate += @{ "Text" = $tileTitleHTML }
                            }
                        }
                    }
                    # Extras cibles : certains webparts stockent du texte visible
                    # HORS de searchablePlainTexts, dans webPartData.properties.
                    switch ($control.WebPartId) {
                        "d1d91016-032f-456d-98a4-721247c305e8" {
                            Write-Host "Extras Image : overlayText / overlayTextHTML"
                            foreach ($propName in @("overlayTextHTML", "overlayText")) {
                                $value = $newProperties.webPartData.properties.$propName
                                if ($value -is [string] -and $value) {
                                    $nbPropertiesToTranslate += 1
                                    $newProperties.webPartData.properties.$propName = "<(#Property$nbPropertiesToTranslate#)>"
                                    $allTextsToTranslate += @{ "Text" = $value }
                                }
                            }
                            break
                        }
                        { $_ -in @("62cac389-787f-495d-beca-e11786162ef4", "81b57906-cbed-4bb1-9823-2e3314f46f28") } {
                            Write-Host "Extra description : Countdown / World clock"
                            $description = $newProperties.webPartData.properties.description
                            if ($description -is [string] -and $description) {
                                $nbPropertiesToTranslate += 1
                                $newProperties.webPartData.properties.description = "<(#Property$nbPropertiesToTranslate#)>"
                                $allTextsToTranslate += @{ "Text" = $description }
                            }
                            break
                        }
                        "8c88f208-6c77-4bdb-86a0-0c47b4316588" {
                            # Webpart Actualites : adapter le filtre MosaicLanguage
                            # a la langue de la copie (accueil en -> actualites en).
                            $filters = $newProperties.webPartData.properties.filters
                            if ($filters) {
                                foreach ($filter in $filters) {
                                    if ($filter.fieldname -eq "MosaicLanguage") {
                                        Write-Host "Actualites : filtre MosaicLanguage -> $($language.sp)"
                                        $filter.values = @($language.sp)
                                        $forceWebPartUpdate = $true
                                    }
                                }
                            }
                            if ($forceWebPartUpdate -and $newProperties.webPartData.properties.serializedFilterQuery) {
                                $newProperties.webPartData.properties.serializedFilterQuery = $newProperties.webPartData.properties.serializedFilterQuery -replace '(<FieldRef Name="MosaicLanguage"/><Value Type="Choice">)[^<]*(</Value>)', ('${1}' + $language.sp + '${2}')
                            }
                            break
                        }
                        "2f3b693c-1054-419c-af04-fee2782b414f" {
                            Write-Host "Extra titre : My feed"
                            $title = $newProperties.webPartData.properties.title
                            if ($title -is [string] -and $title) {
                                $nbPropertiesToTranslate += 1
                                $newProperties.webPartData.properties.title = "<(#Property$nbPropertiesToTranslate#)>"
                                $allTextsToTranslate += @{ "Text" = $title }
                            }
                            break
                        }
                    }
                    $allWebParts += [PSCustomObject]@{
                        nbTexts        = $nbPropertiesToTranslate
                        forceUpdate    = $forceWebPartUpdate
                        type           = $control.Type.Name
                        id             = $control.InstanceId
                        dataProperties = $($newProperties | ConvertTo-Json -Depth 15).ToString()
                    }
                    break
                }
                default {
                    Write-Host "Can't translate this type"
                }
            }
        }
        Write-Host "Before translation : $($stopwatch.Elapsed.TotalSeconds)"

        $jsonTextsToTranslate = $allTextsToTranslate | ConvertTo-Json -AsArray -Compress -Depth 7 -EscapeHandling EscapeHtml
        Write-Host "Texts to translate : $jsonTextsToTranslate"

        #Translate all texts for one page once and set them
        $allTranslatedTexts = Start-Translation -Text $jsonTextsToTranslate -Language $language.az -Service microsoftTranslator
        $k = 0
        foreach ($currentWebPart in $allWebParts) {
            for ($i = 1; $i -lt $($currentWebPart.nbTexts + 1); $i++) {
                $translatedText = $allTranslatedTexts[$k].translations[0].text
                if ($currentWebPart.type -eq "PageWebPart") {
                    # Reinsertion dans une chaine JSON : echapper le texte traduit, et couvrir
                    # les deux formes du placeholder (brute, et <... issue de ConvertTo-Json).
                    $escapedText = (ConvertTo-Json "$translatedText" -Compress).Trim('"')
                    $escapedPlaceholder = [string][char]92 + "u003c(#Property$i#)" + [string][char]92 + "u003e"
                    $currentWebPart.dataProperties = $currentWebPart.dataProperties.Replace("<(#Property$i#)>", $escapedText).Replace($escapedPlaceholder, $escapedText)
                }
                else {
                    $currentWebPart.dataProperties = $currentWebPart.dataProperties.Replace("<(#Property$i#)>", $translatedText)
                }
                $k += 1
            }
            if ($currentWebPart.type -eq "PageText" -and $currentWebPart.nbTexts -gt 0) {
                Set-PnPPageTextPart -Page $currentPage -InstanceId $currentWebPart.id -Text $currentWebPart.dataProperties
            }
            elseif ($currentWebPart.type -eq "PageWebPart" -and ($currentWebPart.nbTexts -gt 0 -or $currentWebPart.forceUpdate)) {
                Set-PnPPageWebPart -Page $currentPage -Identity $currentWebPart.id -PropertiesJson $currentWebPart.dataProperties
            }
        }
        Set-PnPPage -Identity $currentPage -CommentsEnabled:$(!($commentsDisabled))
        Repair-PageCanvas -PageRelativePath $currentPage
        # Colonne MosaicLanguage : marquer la copie avec le code de sa langue
        # (fr, en...) pour le filtrage explicite des webparts Actualites.
        # SystemUpdate = pas de nouvelle version, pas de changement d'auteur.
        # Si la colonne n'existe pas sur ce site, on continue sans bloquer.
        try {
            $sitePath = ([uri]$siteUrl).AbsolutePath.TrimEnd('/')
            $pageItem = Get-PnPFile -Url "$sitePath/SitePages/$currentPage" -AsListItem -ErrorAction Stop
            $pageItem["MosaicLanguage"] = $language.sp
            $pageItem.SystemUpdate()
            Invoke-PnPQuery
            Write-Host "MosaicLanguage = $($language.sp) sur $currentPage"
        }
        catch {
            Write-Warning "MosaicLanguage non renseigne sur $currentPage : $_"
        }
        # Canari qualite : verifier que l'editeur (co-edition) accepte la copie
        # finale, AVANT toute publication. Toute future incompatibilite de
        # serialisation est detectee ici et la page reste en brouillon.
        $editorOk = $true
        $checkItem = $null
        try {
            $checkItem = Get-PnPFile -Url "$(([uri]$siteUrl).AbsolutePath.TrimEnd('/'))/SitePages/$currentPage" -AsListItem -ErrorAction Stop
            $null = Invoke-PnPSPRestMethod -Method Post -Url "/_api/sitepages/pages($($checkItem.Id))/startCoAuth" -ErrorAction Stop
            Write-Host "Validation editeur OK : $currentPage"
        }
        catch {
            $editorOk = $false
            $failedValidations += $currentPage
            Write-Warning "VALIDATION EDITEUR KO sur $currentPage : $_"
        }
        # Statut final : Completed si la validation editeur est passee, Failed
        # sinon (une copie Failed est reprise automatiquement au prochain passage).
        if ($checkItem) {
            try {
                $checkItem["MosaicTranslationStatus"] = $(if ($editorOk) { "Completed" } else { "Failed" })
                $checkItem.SystemUpdate()
                Invoke-PnPQuery
            } catch { Write-Warning "Statut final non pose sur $currentPage (colonne absente ?)" }
        }        # Par defaut les copies restent en brouillon ; publier seulement si
        # demande ET si la validation editeur est passee (fail-safe : une page
        # douteuse n'est jamais publiee automatiquement, elle attend la relecture).
        if ($env:PUBLISH_AFTER_TRANSLATION -eq "true" -and $editorOk) {
            $modernPage.Publish()
        }
        Write-Host "Page $currentPage is translated"
    }
    Write-Host "After translate a page : $($stopwatch.Elapsed.TotalSeconds)"

    # Limite d'execution (~10 min max) : apres 500 s, on renvoie 202 avec les pages
    # restantes ; le dialogue SPFx rappelle la function avec ce reliquat.
    $remainingPages = $pages.Length - 1 - $pages.IndexOf($pagePath)
    if ($stopwatch.Elapsed.TotalSeconds -gt 500 -and $remainingPages -gt 0) {
        $body = $(@{
                siteUrl   = $siteUrl
                pages     = $pages[($pages.IndexOf($pagePath) + 1)..($pages.Length - 1)]
                languages = $languages
            } | ConvertTo-Json).ToString()
        Disconnect-PnPOnline
        Send-Response -code Accepted -body $(@{ msg = $body } | ConvertTo-Json)
    }
}

Disconnect-PnPOnline

Write-Host "End : $($stopwatch.Elapsed.TotalSeconds)"

$finalMsg = "Traduction terminee : $($pages -join ', ') -> $(($languageCodes.sp) -join ', ')"
if ($failedValidations.Count -gt 0) {
    $finalMsg += " | ATTENTION validation editeur KO, copies laissees en brouillon : $($failedValidations -join ', ')"
}
Send-Response -code OK -body $(@{ msg = $finalMsg } | ConvertTo-Json)
