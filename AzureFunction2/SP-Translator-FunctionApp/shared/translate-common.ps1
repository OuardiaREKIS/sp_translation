# Module partage entre l'orchestrateur HTTP et le worker de file :
# traduction complete d'UNE copie de langue, et utilitaires communs.

# Chrono de diagnostic : une ligne par etape, requetable dans App Insights
# (table traces, filtre "MOSAIC PERF"). $Last porte le cumul precedent pour
# afficher la duree de l'etape seule.
function Write-MosaicPerf {
    param(
        [Parameter(Mandatory = $true)][string]$Language,
        [Parameter(Mandatory = $true)][string]$Step,
        [Parameter(Mandatory = $true)][System.Diagnostics.Stopwatch]$Watch,
        [Parameter(Mandatory = $true)][ref]$Last
    )
    $elapsed = $Watch.Elapsed
    $delta = $elapsed - $Last.Value
    $Last.Value = $elapsed
    Write-Host ("[MOSAIC PERF] {0} {1} {2:N1} s" -f $Language, $Step.PadRight(18), $delta.TotalSeconds)
}

function Start-Translation {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Language,
        [Parameter(Mandatory = $true)][string]$Service
    )
    if ($Service -eq "microsoftTranslator") {
        $baseUri = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0"
        $headers = @{
            'Ocp-Apim-Subscription-Key'    = "$($env:TRANSLATOR_KEY)"
            'Ocp-Apim-Subscription-Region' = "$($env:TRANSLATOR_REGION)"
        }
        $uri = "$baseUri&to=$Language&textType=html"
    }
    $results = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $Text -ContentType "application/json; charset=utf-8"
    return $results
}

# Garde anti-corruption : la serialisation PnP recrache le controle
# "fond de section" (controlType 14) avec "webPartId":null / {"id":null,
# que la co-edition SharePoint ne sait pas relire. Nettoyage apres sauvegarde.
function Repair-PageCanvas {
    param(
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$PageRelativePath
    )
    try {
        $sitePath = ([uri]$SiteUrl).AbsolutePath.TrimEnd('/')
        $item = Get-PnPFile -Url "$sitePath/SitePages/$PageRelativePath" -AsListItem -ErrorAction Stop
        $canvas = [string]$item["CanvasContent1"]
        if (-not $canvas) { return }
        $fixed = $canvas.Replace('&quot;webPartId&quot;&#58;null,', '').Replace('&quot;webPartId&quot;:null,', '').Replace('"webPartId":null,', '')
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

# Traduction complete d'UNE copie de langue (la copie doit deja exister).
# Retourne $true si tout est passe (statut Completed), $false sinon (Failed).
function Invoke-CopyTranslation {
    param(
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$PagePath,
        [Parameter(Mandatory = $true)][string]$LanguageSp,
        [Parameter(Mandatory = $true)][string]$LanguageAz,
        [Parameter(Mandatory = $true)][string]$SourceTitle,
        [string]$TopicHeader = "",
        [bool]$CommentsDisabled = $false
    )
    $sitePath = ([uri]$SiteUrl).AbsolutePath.TrimEnd('/')
    $currentPage = "$LanguageSp/$PagePath"
    Write-Host "Traduction de $currentPage"
    $perfWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $perfLast = [TimeSpan]::Zero

    # Statut Pending des le debut : si le job s'interrompt, la copie sera
    # reprise (retentative de la file, ou nouveau clic).
    try {
        $statusItem = Get-PnPFile -Url "$sitePath/SitePages/$currentPage" -AsListItem -ErrorAction Stop
        $statusItem["MosaicTranslationStatus"] = "Pending"
        $statusItem.SystemUpdate()
        Invoke-PnPQuery
    } catch { Write-Warning "Statut Pending non pose sur $currentPage (colonne absente ?)" }
    Write-MosaicPerf -Language $LanguageSp -Step "StatutPending" -Watch $perfWatch -Last ([ref]$perfLast)

    $modernPage = Get-PnPPage -Identity $currentPage
    Write-MosaicPerf -Language $LanguageSp -Step "LecturePage" -Watch $perfWatch -Last ([ref]$perfLast)
    # Titre traduit depuis la SOURCE (jamais depuis la copie : SharePoint la
    # baptise "Translate to <langue>: <titre>").
    $oldPageTitle = "[$(@{ "Text" = $SourceTitle } | ConvertTo-Json)]"
    $modernPage.PageTitle = $(Start-Translation -Text $oldPageTitle -Language $LanguageAz -Service microsoftTranslator)[0].translations[0].text
    if ($TopicHeader) {
        $oldTopicHeader = "[$(@{ "Text" = $TopicHeader } | ConvertTo-Json)]"
        $modernPage.PageHeader.TopicHeader = $(Start-Translation -Text $oldTopicHeader -Language $LanguageAz -Service microsoftTranslator)[0].translations[0].text
    }
    $modernPage.Save()
    Write-MosaicPerf -Language $LanguageSp -Step "TitreEtEnTete" -Watch $perfWatch -Last ([ref]$perfLast)

    $allWebParts = @()
    $allTextsToTranslate = @()
    foreach ($control in $modernPage.Controls) {
        switch ($control.Type.Name) {
            "PageText" {
                $allWebParts += [PSCustomObject]@{
                    nbTexts = 1; type = $control.Type.Name; id = $control.InstanceId; dataProperties = "<(#Property1#)>"
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
                # Passe generique : tous les textes visibles (searchablePlainTexts).
                # links/imageSources/htmlStrings/IDs restent intacts.
                if ($oldProperties.searchablePlainTexts) {
                    foreach ($textProperty in $oldProperties.searchablePlainTexts.PSObject.Properties) {
                        if ($textProperty.Value -is [string] -and $textProperty.Value) {
                            $nbPropertiesToTranslate += 1
                            $newProperties.webPartData.serverProcessedContent.searchablePlainTexts.$($textProperty.Name) = "<(#Property$nbPropertiesToTranslate#)>"
                            $allTextsToTranslate += @{ "Text" = $textProperty.Value }
                        }
                    }
                }
                # titleHTML (version riche affichee), webpart + tuiles Hero
                $titleHTML = $newProperties.webPartData.properties.titleHTML
                if ($titleHTML -is [string] -and $titleHTML) {
                    $nbPropertiesToTranslate += 1
                    $newProperties.webPartData.properties.titleHTML = "<(#Property$nbPropertiesToTranslate#)>"
                    $allTextsToTranslate += @{ "Text" = $titleHTML }
                }
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
                # Extras hors searchablePlainTexts
                switch ($control.WebPartId) {
                    "d1d91016-032f-456d-98a4-721247c305e8" {
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
                        $description = $newProperties.webPartData.properties.description
                        if ($description -is [string] -and $description) {
                            $nbPropertiesToTranslate += 1
                            $newProperties.webPartData.properties.description = "<(#Property$nbPropertiesToTranslate#)>"
                            $allTextsToTranslate += @{ "Text" = $description }
                        }
                        break
                    }
                    "8c88f208-6c77-4bdb-86a0-0c47b4316588" {
                        # Actualites : adapter le filtre MosaicLanguage a la langue
                        $filters = $newProperties.webPartData.properties.filters
                        if ($filters) {
                            foreach ($filter in $filters) {
                                if ($filter.fieldname -eq "MosaicLanguage") {
                                    $filter.values = @($LanguageSp)
                                    $forceWebPartUpdate = $true
                                }
                            }
                        }
                        if ($forceWebPartUpdate -and $newProperties.webPartData.properties.serializedFilterQuery) {
                            $newProperties.webPartData.properties.serializedFilterQuery = $newProperties.webPartData.properties.serializedFilterQuery -replace '(<FieldRef Name="MosaicLanguage"/><Value Type="Choice">)[^<]*(</Value>)', ('${1}' + $LanguageSp + '${2}')
                        }
                        break
                    }
                    "2f3b693c-1054-419c-af04-fee2782b414f" {
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
                    nbTexts = $nbPropertiesToTranslate; forceUpdate = $forceWebPartUpdate
                    type = $control.Type.Name; id = $control.InstanceId
                    dataProperties = $($newProperties | ConvertTo-Json -Depth 15).ToString()
                }
                break
            }
            default { }
        }
    }

    Write-MosaicPerf -Language $LanguageSp -Step "AnalyseWebParts" -Watch $perfWatch -Last ([ref]$perfLast)

    if ($allTextsToTranslate.Count -gt 0) {
        $jsonTextsToTranslate = $allTextsToTranslate | ConvertTo-Json -AsArray -Compress -Depth 7 -EscapeHandling EscapeHtml
        $allTranslatedTexts = Start-Translation -Text $jsonTextsToTranslate -Language $LanguageAz -Service microsoftTranslator
    }
    Write-MosaicPerf -Language $LanguageSp -Step "Translator" -Watch $perfWatch -Last ([ref]$perfLast)
    $k = 0
    foreach ($currentWebPart in $allWebParts) {
        for ($i = 1; $i -lt $($currentWebPart.nbTexts + 1); $i++) {
            $translatedText = $allTranslatedTexts[$k].translations[0].text
            if ($currentWebPart.type -eq "PageWebPart") {
                # Reinsertion dans une chaine JSON : echapper, et couvrir les
                # deux formes du placeholder (brute, et issue de ConvertTo-Json).
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
    Write-MosaicPerf -Language $LanguageSp -Step "EcritureWebParts" -Watch $perfWatch -Last ([ref]$perfLast)
    Set-PnPPage -Identity $currentPage -CommentsEnabled:$(!($CommentsDisabled))
    Write-MosaicPerf -Language $LanguageSp -Step "Commentaires" -Watch $perfWatch -Last ([ref]$perfLast)
    Repair-PageCanvas -SiteUrl $SiteUrl -PageRelativePath $currentPage
    Write-MosaicPerf -Language $LanguageSp -Step "ReparationCanvas" -Watch $perfWatch -Last ([ref]$perfLast)

    # MosaicLanguage sur la copie (SystemUpdate : pas de version parasite)
    try {
        $pageItem = Get-PnPFile -Url "$sitePath/SitePages/$currentPage" -AsListItem -ErrorAction Stop
        $pageItem["MosaicLanguage"] = $LanguageSp
        $pageItem.SystemUpdate()
        Invoke-PnPQuery
    } catch { Write-Warning "MosaicLanguage non renseigne sur $currentPage : $_" }
    Write-MosaicPerf -Language $LanguageSp -Step "MetadonneeLangue" -Watch $perfWatch -Last ([ref]$perfLast)

    # Validation locale du canvas (JSON valide, aucun identifiant null)
    $editorOk = $true
    $checkItem = $null
    try {
        $checkItem = Get-PnPFile -Url "$sitePath/SitePages/$currentPage" -AsListItem -ErrorAction Stop
        $rawCanvas = [string]$checkItem["CanvasContent1"]
        if ($rawCanvas) {
            foreach ($m in [regex]::Matches($rawCanvas, 'data-sp-(?:controldata|webpartdata)="([^"]*)"')) {
                $ctrlJson = [System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value)
                $null = [System.Text.Json.JsonDocument]::Parse($ctrlJson)
                if ($ctrlJson -match '"webPartId"\s*:\s*null' -or $ctrlJson -match '^\s*\{\s*"id"\s*:\s*null') {
                    throw "identifiant null dans le canvas (controle fond de section)"
                }
            }
        }
        Write-Host "Validation canvas OK : $currentPage"
    }
    catch {
        $editorOk = $false
        Write-Warning "VALIDATION CANVAS KO sur $currentPage : $_"
    }
    if ($checkItem) {
        try {
            $checkItem["MosaicTranslationStatus"] = $(if ($editorOk) { "Completed" } else { "Failed" })
            $checkItem.SystemUpdate()
            Invoke-PnPQuery
        } catch { Write-Warning "Statut final non pose sur $currentPage" }
    }
    Write-MosaicPerf -Language $LanguageSp -Step "ValidationStatut" -Watch $perfWatch -Last ([ref]$perfLast)
    # Publication uniquement si demandee ET validation passee (fail-safe)
    if ($env:PUBLISH_AFTER_TRANSLATION -eq "true" -and $editorOk) {
        $modernPage.Publish()
    }
    Write-Host ("[MOSAIC PERF] {0} {1} {2:N1} s" -f $LanguageSp, "TOTAL".PadRight(18), $perfWatch.Elapsed.TotalSeconds)
    Write-Host "Page $currentPage is translated (ok=$editorOk)"
    return $editorOk
}
