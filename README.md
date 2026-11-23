# sp_translator — traduction automatique des pages SharePoint

Solution qui crée les **copies multilingues** des pages d'un site SharePoint Online
et les **traduit automatiquement** avec Azure AI Translator, en conservant la mise
en page, les images, les liens et les paramètres des composants.

L'utilisateur clique sur un bouton dans SharePoint, choisit les langues, et les
copies traduites apparaissent dans les dossiers de langue du site, **en brouillon**,
prêtes à être relues puis publiées.

---

## Comment ça marche

```
  SharePoint                          Azure Function App
  ──────────                          ──────────────────
  Extension SPFx  ── pages+langues ─▶ HttpTriggerSpTranslator (orchestrateur)
  (bouton)                              │ crée les copies (Pending)
     ▲                                  │ 1 message par langue
     │                                  ▼
     │                            file « translation-jobs »
     │                                  │
     │ relit le statut                  ▼
     │ toutes les 5 s             TranslateWorker (×N en parallèle) ◀──▶ Azure AI Translator
     │                                  │
  Copies traduites  ◀──── écrit ────────┘
  /SitePages/en/… (Completed / Failed)
```

1. **L'extension SPFx** envoie à la Function la liste des pages et des langues choisies.
2. **L'orchestrateur** crée en une fois toutes les copies manquantes (statut `Pending`),
   puis dépose un message par langue dans la file d'attente. Il répond en quelques secondes.
3. **Les workers** consomment la file en parallèle : chacun traduit une copie (titre,
   textes, ~20 types de composants) et passe son statut à `Completed` ou `Failed`.
4. **Le dialogue SPFx** relit les statuts toutes les 5 s (15 min maximum) et affiche le résultat,
   y compris le détail des échecs en cas de réussite partielle (bouton *Relancer*).

---

## Contenu du dépôt

| Dossier / fichier | Rôle | Déployé où ? |
|---|---|---|
| `AzureFunction/SP-Translator-FunctionApp/` | Le service de traduction (PowerShell 7.4 + PnP.PowerShell) | **Azure** (Function App) |
| `CustomAction/` | L'extension SharePoint Framework (boutons + dialogue) | **SharePoint** (paquet `.sppkg`) |
| `tests/` | Tests automatiques à lancer avant chaque déploiement | nulle part (poste du développeur) |
| `inspect-page.ps1`, `inspect_homepage.ps1`, `check-page-locks.ps1`, `repair-canvas.ps1` | Outils de diagnostic et de réparation | nulle part (poste du développeur) |
| `PRE-DEPLOY-CHECKLIST.md` | Checklist à dérouler avant chaque mise en production | — |
| `REGLES-IMPLICITES.md` | Comportements non évidents à connaître pour exploiter la solution | — |

### Azure Function

| Fonction | Déclencheur | Rôle |
|---|---|---|
| `HttpTriggerSpTranslator` | HTTP (clé de fonction) | Orchestrateur : valide, crée les copies, remplit la file |
| `TranslateWorker` | File `translation-jobs` | Traduit une copie de langue par message (5 tentatives, puis file *poison*) |
| `shared/translate-common.ps1` | — | Code commun : traduction des composants, statuts, réparation canvas |

### Extension SPFx (`CustomAction/`)

| Composant | Où il apparaît |
|---|---|
| **TranslatorCommandSet** | Bouton « Traduction » dans la bibliothèque *Pages du site* (une ou plusieurs pages sélectionnées) |
| **PageTranslatorApplicationCustomizer** | Bouton flottant « 🌐 Traduire la page » en bas à droite d'une page source |

Le bouton n'est jamais proposé sur une copie traduite, ni en création / modification de page.

---

## Prérequis

**Outils (poste du développeur)**
- PowerShell 7.4 et le module **PnP.PowerShell 3.3.0**
- [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local) v4
- Node.js **22.x** (≥ 22.14) pour SPFx 1.21.1

**Ressources Azure / Microsoft 365**
- Une **Function App** PowerShell 7.4 avec son compte de stockage (`AzureWebJobsStorage`, utilisé par la file)
- Une ressource **Azure AI Translator**
- Une **application Entra ID** avec authentification par certificat (app-only), droits SharePoint `Sites.FullControl.All`
- Sur chaque site cible :
  - les langues cibles activées (*Paramètres du site → Paramètres de langue*) ;
  - deux colonnes **Choix** dans la bibliothèque des pages :
    `MosaicLanguage` et `MosaicTranslationStatus` (`Pending` / `Completed` / `Failed`).

---

## Configuration

### Variables de la Function (App Settings, ou `local.settings.json` en local)

| Variable | Rôle |
|---|---|
| `TRANSLATOR_KEY` / `TRANSLATOR_REGION` | Clé et région de la ressource Azure AI Translator |
| `SP_CLIENT_ID` / `SP_TENANT` | Application Entra ID (ex. `contoso.onmicrosoft.com`) |
| `SP_THUMBPRINT` | Empreinte du certificat. Sur Azure, définir aussi `WEBSITE_LOAD_CERTIFICATES` avec la même valeur |
| `SP_CERT_PATH` / `SP_CERT_PASSWORD` | Alternative en local : chemin du `.pfx` et son mot de passe |
| `PUBLISH_AFTER_TRANSLATION` | `false` (défaut) = copies en brouillon ; `true` = publiées automatiquement |
| `PSWorkerInProcConcurrencyUpperBound` | **`1` obligatoire** : un seul message traité à la fois par instance (PnP n'est pas sûr en parallèle dans un même processus) |

### URL de la Function dans l'extension

Copier `CustomAction/src/components/constants/Azure.json.example` en `Azure.json`
et renseigner `funcAppUrl` :

- en local : `http://localhost:7071/api/HttpTriggerSpTranslator`
- en production : `https://<function-app>.azurewebsites.net/api/HttpTriggerSpTranslator?code=<clé>`

> ⚠️ L'URL est intégrée au bundle : la renseigner **avant** de packager.
> La Function App doit autoriser l'origine du tenant SharePoint en **CORS**.

---

## Lancer en local

```powershell
# 1. La Function (avec Azurite ou un vrai compte de stockage pour la file)
cd AzureFunction/SP-Translator-FunctionApp
func start

# 2. L'extension SPFx, dans un autre terminal
cd CustomAction
npm install
npx gulp serve                       # bouton sur une page
npx gulp serve --config commandSet   # bouton dans la bibliothèque des pages
```

Accepter « Charger les scripts de débogage » sur la page SharePoint qui s'ouvre.

---

## Déployer

Dérouler d'abord **[PRE-DEPLOY-CHECKLIST.md](PRE-DEPLOY-CHECKLIST.md)** (tests bloquants marqués 🔴).

```powershell
# 1. Tests
.\tests\run-tests.ps1                # doit afficher : TOUS LES TESTS PASSENT

# 2. Azure Function
cd AzureFunction/SP-Translator-FunctionApp
func azure functionapp publish <nom-function-app>
# puis, si une fonction à déclencheur de file a été ajoutée ou modifiée :
#   az functionapp restart  +  synchronisation des déclencheurs (syncfunctiontriggers)

# 3. Extension SPFx
cd CustomAction
# incrémenter "version" dans config/package-solution.json
npm run package
```

Déposer ensuite `CustomAction/sharepoint/solution/sp-translator.sppkg` dans le
**catalogue d'applications** du tenant et cliquer **Déployer**. La solution est
disponible sur tous les sites, sans installation site par site.

---

## Outils de diagnostic

Scripts à lancer à la main, uniquement en cas de problème. Ils lisent
`SP_CLIENT_ID` et `SP_THUMBPRINT` dans l'environnement :

```powershell
$env:SP_CLIENT_ID = "<client-id>"; $env:SP_THUMBPRINT = "<empreinte>"
```

| Script | Quand l'utiliser |
|---|---|
| `inspect-page.ps1 -Page "ma-page.aspx" [-Copies]` | Voir tous les composants et textes d'une page (ex. comprendre pourquoi un texte n'est pas traduit). Tenant de test écrit dans le script |
| `inspect_homepage.ps1 -Page ... -SiteUrl ... -Tenant ...` | Même chose, sur n'importe quel site (connexion paramétrable) |
| `check-page-locks.ps1 -SiteUrl ... -Tenant ...` | Une page reste « en cours de modification » : voir qui tient le verrou |
| `repair-canvas.ps1 -SiteUrl ... -Tenant ...` | « Impossible de modifier cette page » : diagnostic seul ; ajouter `-Apply` pour réparer (sauvegarde dans `canvas-backup/`) |

---

## Bon à savoir

- **Une copie existante n'est jamais écrasée** (protection des corrections des relecteurs).
  Pour retraduire, supprimer la copie puis relancer.
- Une copie `Pending` / `Failed` est **reprise automatiquement** au clic suivant.
- Les copies naissent **en brouillon** : invisibles des visiteurs tant qu'elles ne sont pas publiées.
- Le message générique « The page hasn't been translated. » cache la vraie erreur :
  la regarder dans la console du navigateur (F12).
- Le premier appel après un déploiement est lent (1 à 3 min, chargement de PnP).

Toutes les règles détaillées : **[REGLES-IMPLICITES.md](REGLES-IMPLICITES.md)**.

---

## Versions

| Élément | Version |
|---|---|
| Extension SPFx (`package-solution.json`) | 1.6.0.0 |
| SharePoint Framework | 1.21.1 (React 17, Fluent UI 8, PnP JS 3) |
| PowerShell / PnP.PowerShell | 7.4 / 3.3.0 |
