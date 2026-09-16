# sp-translator (extension SPFx)

## Résumé

Extension SharePoint Framework du projet **sp_translator** : traduction automatique
des pages SharePoint via l'Azure Function `HttpTriggerSpTranslator` (dossier
`../AzureFunction2`) et Azure AI Translator.

Deux composants dans la même solution :

| Composant | Où il apparaît | GUID |
|---|---|---|
| **TranslatorCommandSet** | Bouton « Traduction » dans la bibliothèque *Pages du site* (sélection d'une ou plusieurs pages) | `ec203189-1da1-424e-9ebb-b4bc9ac12131` |
| **PageTranslatorApplicationCustomizer** | Bouton flottant « 🌐 Traduire la page » en bas à droite quand on **lit une page source** | `5a7c31f0-9d84-4c2e-b1a6-3f8e52d90c47` |

Les deux ouvrent le même dialogue : cases à cocher des langues activées sur le site
(Paramètres du site → Paramètres de langue), puis appel de la function qui crée les
copies multilingues manquantes et les traduit (**brouillons** par défaut). Le bouton
n'est jamais proposé sur une copie traduite (drapeau `_SPIsTranslation`).

## Versions

- SharePoint Framework : **1.21.1** — Node : **22.x** (>= 22.14)
- React 17 / Fluent UI 8 / PnP JS 3

## Configuration

- **URL de la function** : `src/components/constants/Azure.json` → `funcAppUrl`.
  - Test local : `http://localhost:7071/api/HttpTriggerSpTranslator`
  - Prod : `https://<ta-function-app>.azurewebsites.net/api/HttpTriggerSpTranslator?code=<clé>`
  - ⚠️ L'URL est embarquée dans le bundle : la changer **avant** de packager.
- **CORS** : la Function App doit autoriser l'origine `https://3whd6t.sharepoint.com`
  (en local, `local.settings.json` a déjà `"CORS": "*"`).

## Test en local (sans déployer)

1. Lancer la function : dans `../AzureFunction2/SP-Translator-FunctionApp` → `func start`
   (nécessite Azure Functions Core Tools ; le certificat .pfx doit être accessible —
   renseigner `SP_CERT_PATH`/`SP_CERT_PASSWORD` dans `local.settings.json`, ou importer
   le .pfx dans le magasin de certificats utilisateur).
2. Lancer le SPFx : dans ce dossier → `npx gulp serve`
   (config `default` = bouton sur la page `test-mosaic.aspx` ; config `commandSet` =
   bibliothèque des pages : `npx gulp serve --config commandSet`).
3. Accepter « Charger les scripts de débogage » sur la page qui s'ouvre.

## Packaging & déploiement

```powershell
npm run package    # = gulp clean + bundle --ship + package-solution --ship
```

Puis déposer `sharepoint/solution/sp-translator.sppkg` dans le **Catalogue
d'applications** du tenant et cliquer **Déployer** (la solution est en
`skipFeatureDeployment: true` → disponible tenant-wide sans installation par site,
via `ClientSideInstance.xml`).

## Historique

| Version | Date | Commentaires |
| ------- | ---- | ------------ |
| 1.0 | 23 mai 2023 | Version initiale (SPFx 1.16, command set seul) |
| 1.1 | août 2026 | SPFx 1.21.1/Node 22, bouton sur page (application customizer), branchement AzureFunction2, brouillons par défaut |
