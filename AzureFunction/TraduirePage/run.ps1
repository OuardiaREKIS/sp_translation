using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "TraduirePage : requete recue."

# Parametres attendus :
#   page    : la page a traduire (ex: myPage.aspx)  [OBLIGATOIRE]
#   langues : dossiers de langue a regenerer (ex: "fr,de")  [OPTIONNEL]
#             Si fourni => regenere SEULEMENT ces langues (choix du contributeur).
#             Si absent => regenere TOUTES les langues gerees (pratique pour tester).
$pageSource   = $Request.Query.page
if (-not $pageSource) { $pageSource = $Request.Body.page }

$languesDemandees = $Request.Query.langues
if (-not $languesDemandees) { $languesDemandees = $Request.Body.langues }

if (-not $pageSource) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = "Parametre 'page' manquant (ex: ?page=myPage.aspx)"
    })
    return
}

# Secrets fournis par les App Settings de la Function Azure
$cleTranslator = $env:TRANSLATOR_KEY
$region        = $env:TRANSLATOR_REGION
$siteUrl       = $env:SP_SITE_URL
$clientId      = $env:SP_CLIENT_ID
$thumbprint    = $env:SP_THUMBPRINT
$tenant        = $env:SP_TENANT
$endpointTr    = "https://api.cognitive.microsofttranslator.com"

# Le site est en ANGLAIS : on traduit depuis "en". Ne pas mettre l'anglais en cible.
$langueSource = "en"

# dossier = nom du dossier SharePoint / sp = code creation / tr = code Translator
$toutesLangues = @(
    @{ dossier = "fr";     sp = "fr-fr"; tr = "fr" }       # Francais (cible)
    @{ dossier = "es";     sp = "es-es"; tr = "es" }       # Espagnol
    @{ dossier = "de";     sp = "de-de"; tr = "de" }       # Allemand
    @{ dossier = "zh-chs"; sp = "zh-cn"; tr = "zh-Hans" }  # Chinois simplifie
    @{ dossier = "tr";     sp = "tr-tr"; tr = "tr" }       # Turc
    @{ dossier = "ar";     sp = "ar-sa"; tr = "ar" }       # Arabe
    @{ dossier = "ja";     sp = "ja-jp"; tr = "ja" }       # Japonais
)

if ($languesDemandees) {
    $dossiersARegenerer = $languesDemandees -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $languesATraiter = $toutesLangues | Where-Object { $dossiersARegenerer -contains $_.dossier }

    if ($languesATraiter.Count -eq 0) {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body = "Aucune langue valide selectionnee. Langues demandees : $languesDemandees. Langues gerees : $(( $toutesLangues.dossier ) -join ', ')."
        })
        return
    }
} else {
    $languesATraiter = $toutesLangues
    Write-Host "Aucune langue precisee : traitement de TOUTES les langues gerees."
}

# Traduit un fragment HTML via Azure Translator.
# Mode textType=html : les balises sont preservees, seul le texte est traduit
# => la mise en forme de la page ne bouge pas. from = langue source (ici "en").
function Traduire-Html {
    param([string]$html, [string]$langueCible)
    if (-not $html) { return $html }
    $url = "$endpointTr/translate?api-version=3.0&from=$langueSource&to=$langueCible&textType=html"
    $headers = @{
        "Ocp-Apim-Subscription-Key"    = $cleTranslator
        "Ocp-Apim-Subscription-Region" = $region
        "Content-Type"                 = "application/json"
    }
    $corps = ConvertTo-Json @(@{ Text = $html })
    $rep = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $corps
    return $rep[0].translations[0].text
}

$resultats = @()

try {
    # Connexion app-only (certificat)
    Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenant

    $page = Get-PnPPage -Identity $pageSource
    $pageId = $page.PageId

    # Supprime les copies a regenerer ; les langues non demandees ne sont pas touchees
    foreach ($lang in $languesATraiter) {
        $cheminCopie = "SitePages/$($lang.dossier)/$pageSource"
        try {
            Remove-PnPFile -ServerRelativeUrl "/$cheminCopie" -Force -ErrorAction Stop
            Write-Host "Ancienne copie supprimee : $($lang.dossier)"
        } catch {
            Write-Host "Pas de copie existante pour $($lang.dossier) (elle sera creee)."
        }
    }

    # Recree les copies via l'API translations de SharePoint (design frais depuis la source)
    $codesSP = ($languesATraiter | ForEach-Object { '"' + $_.sp + '"' }) -join ","
    $body = @"
{
  "request": {
    "__metadata": { "type": "SP.TranslationStatusCreationRequest" },
    "LanguageCodes": { "results": [ $codesSP ] }
  }
}
"@
    try {
        Invoke-PnPSPRestMethod -Method Post -Url "/_api/sitepages/pages($pageId)/translations/create" -Content $body -ContentType "application/json;odata=verbose" | Out-Null
        Write-Host "Copies (re)creees pour : $(( $languesATraiter.dossier ) -join ', ')"
    } catch {
        Write-Host "Souci a la creation des copies : $_"
    }

    foreach ($lang in $languesATraiter) {
        try {
            $copie = Get-PnPPage -Identity "$($lang.dossier)/$pageSource"

            # Corps de texte
            $blocTexte = $copie.Controls | Where-Object { $_.Type.Name -eq "PageText" }
            if ($blocTexte -and $blocTexte.Text) {
                $blocTexte.Text = Traduire-Html -html $blocTexte.Text -langueCible $lang.tr
            }

            # Titre de la page
            $copie.PageTitle = Traduire-Html -html $page.PageTitle -langueCible $lang.tr

            # Textes dans les webparts (image overlay, hero, titre...)
            foreach ($wp in ($copie.Controls | Where-Object { $_.Type.Name -eq "PageWebPart" })) {
                if (-not $wp.PropertiesJson) { continue }
                $props = $wp.PropertiesJson | ConvertFrom-Json
                $modifie = $false
                foreach ($champ in @("overlayTextHTML","overlayText","htmlTitle","title")) {
                    if ($props.PSObject.Properties.Name -contains $champ -and $props.$champ) {
                        $props.$champ = Traduire-Html -html $props.$champ -langueCible $lang.tr
                        $modifie = $true
                    }
                }
                if ($modifie) {
                    $wp.PropertiesJson = ($props | ConvertTo-Json -Depth 20 -Compress)
                }
            }

            # Sauvegarde en brouillon
            $copie.Save()
            $resultats += "[$($lang.dossier)] REGENERE"
        }
        catch {
            $resultats += "[$($lang.dossier)] ERREUR : $_"
        }
    }

    $message = "Page '$pageSource' - langues traitees : $(( $languesATraiter.dossier ) -join ', ').`n" + ($resultats -join "`n")
    $statut = [HttpStatusCode]::OK
}
catch {
    $message = "ERREUR GLOBALE : $_"
    $statut = [HttpStatusCode]::InternalServerError
}

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = $statut
    Body = $message
})
