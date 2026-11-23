# Azure Function App — HttpTriggerSpTranslator

## Summary

Function retenue pour le projet sp_translator : elle cree les copies multilingues
des pages SharePoint selectionnees et les traduit automatiquement (Azure AI Translator).

Appelee par l'extension SPFx (CustomAction) avec un body JSON :

```json
{
  "siteUrl": "https://3whd6t.sharepoint.com",
  "pages": ["test-mosaic.aspx"],
  "languages": [{ "code": "fr", "azCode": "fr", "id": 1036 }]
}
```

Comportement :
- Cree uniquement les copies manquantes (les traductions existantes ne sont pas ecrasees).
- Traduit titre, topic header, texte et ~20 types de webparts (mode HTML : mise en forme preservee).
- Les copies restent en **brouillon** (mettre `PUBLISH_AFTER_TRANSLATION=true` pour publier).
- Si l'execution depasse 500 s, renvoie **202** avec les pages restantes ; le dialogue SPFx rappelle la function (reprise automatique).

## Variables d'environnement (App Settings / local.settings.json)

| Variable | Role |
|---|---|
| `TRANSLATOR_KEY` / `TRANSLATOR_REGION` | Ressource Azure AI Translator |
| `SP_CLIENT_ID` / `SP_TENANT` | App Entra ID (app-only) |
| `SP_THUMBPRINT` | Empreinte du certificat (sur Azure : definir aussi `WEBSITE_LOAD_CERTIFICATES` avec la meme empreinte) |
| `SP_CERT_PATH` / `SP_CERT_PASSWORD` | Alternative locale : chemin du .pfx + mot de passe |
| `PUBLISH_AFTER_TRANSLATION` | `false` (defaut) = brouillons ; `true` = publie |

## Versions

- PowerShell : **7.4**
- PnP.PowerShell : **3.3.0** (via requirements.psd1)
