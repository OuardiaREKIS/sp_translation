# Checklist avant deploiement — sp_translator (MOSAIC)

A derouler avant CHAQUE `func azure functionapp publish` et chaque mise a jour
du `.sppkg`. Les 10 tests BLOQUANTS sont marques 🔴 : un seul echec = pas de
deploiement.

## A. Controles du code (10 secondes + 4 cases)

- [ ] 🔴 `.\tests\run-tests.ps1` -> TOUS LES TESTS PASSENT
      (couvre : syntaxe PowerShell, anti-secrets, codes culture LCID,
      detection des titres de bapteme, regles de reparation canvas)
- [ ] `SP_CLIENT_ID` / `SP_TENANT` / `SP_THUMBPRINT` correspondent au tenant cible
- [ ] Version PnP.PowerShell inchangee dans `requirements.psd1` (jamais de
      montee de version PnP en meme temps qu'une evolution de code)
- [ ] 🔴 La creation des copies passe par `translations/create`
      (`Set-PnPPage -Translate` n'a pas ete reintroduit)
- [ ] Aucun `startCoAuth` appele par la fonction ni par `repair-canvas.ps1`
      (validation canvas locale uniquement)

## B. La page canari (une fois creee, elle sert a chaque fois)

Page de test dediee contenant le maximum de composants reels :
titre, plusieurs blocs texte (gras/liens/listes), Hero, Image avec texte
superpose, bouton/lien, webpart Actualites filtre MosaicLanguage,
**fond de section** (obligatoire — bug SectionBackgroundControl),
plusieurs sections/colonnes, tout webpart reellement utilise par l'intranet.

## C-G. Traduction simple (page canari, une langue, ex. -> en)

- [ ] 🔴 La fonction repond succes ; la copie existe au bon emplacement
- [ ] 🔴 Titre traduit, calcule depuis la SOURCE, AUCUN prefixe
      ("Translate to..." / "Traduisez en..." = echec bloquant)
- [ ] 🔴 Contenu traduit : textes, Hero, titres de webparts, titleHTML ;
      accents/Unicode corrects ; pas de residu de langue source
- [ ] 🔴 Elements techniques INTACTS : URLs, images, GUIDs, layout,
      sections, filtres, parametres de webparts
- [ ] La source est intacte, toujours modifiable
- [ ] `MosaicLanguage` = langue cible ; `MosaicTranslationStatus` = Completed
- [ ] 🔴 Fond de section : visible, image/couleur correcte, et
      `.\repair-canvas.ps1` (diagnostic) ne signale RIEN sur la copie

## H-I. Editeur reel et verrous (la validation auto ne remplace pas ca)

- [ ] 🔴 Ouvrir la copie -> **Modifier** s'ouvre (pas de "Impossible de
      modifier cette page pour le moment") -> petite modif -> Enregistrer ->
      rouvrir -> re-Modifier OK
- [ ] 🔴 **Supprimer** la copie : suppression IMMEDIATE, sans message
      "fichier ouvert", sans attendre (test anti-verrou startCoAuth)
- [ ] Recreer la copie ensuite pour poursuivre

## J-K. Multilingue et Actualites

- [ ] `MosaicLanguage` correct sur chaque copie (en n'herite jamais de fr)
- [ ] Filtre du webpart Actualites : champ inchange, seule la VALEUR adaptee ;
      accueil fr -> actus fr, accueil en -> actus en
- [ ] Multi-langues avant release majeure : en (latin), de/fr (accents),
      tr (ç/ğ/ı), zh (CJK), ar (RTL) — titre, corps, statut, editable,
      supprimable pour chacune

## L-N. Reprise et protection de la relecture

- [ ] 🔴 Copie interrompue (Pending/Failed ou titre de bapteme) -> relance ->
      reprise automatique, titre final SANS prefixe, statut Completed
- [ ] 🔴 Copie terminee (voire corrigee par un relecteur) -> relance ->
      NON retraduite, titre non ecrase, corrections humaines preservees,
      aucun doublon

## P-R. Erreurs et publication

- [ ] Erreur Translator ou validation canvas KO -> Failed, PAS de publication,
      erreur logguee (page + langue + cause identifiables), source intacte,
      copie reprenable
- [ ] Avec `PUBLISH_AFTER_TRANSLATION=true` : une copie valide est publiee,
      une Failed/incomplete JAMAIS

## S. Droits utilisateurs — REGLE A DECIDER (pas encore implementee)

Aujourd'hui le bouton est visible par TOUS, lecteurs compris (l'appel
fonctionne car la fonction est app-only). Decision en attente : filtrer
l'affichage aux contributeurs (`SPPermission.editListItems`) et, en
production client, remplacer la cle `?code=` par Easy Auth Entra ID.
Tant que non traite : a mentionner dans la recette client.

## T. Dernier passage avant publish

Sur la page canari : traduire -> titre OK -> contenu OK -> MosaicLanguage OK ->
filtre Actus OK -> Completed -> Modifier -> enregistrer -> re-Modifier ->
supprimer (immediat). Tout vert = GO :

    func azure functionapp publish func-mosaic-traduction-v2

Pour le `.sppkg` : version incrementee dans package-solution.json, upload
catalogue, verifier le HASH du bundle charge dans le manifeste de page.

References : REGLES-IMPLICITES.md, tests/run-tests.ps1, repair-canvas.ps1,
check-page-locks.ps1, inspect_homepage.ps1.
