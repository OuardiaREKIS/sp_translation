# Regles implicites de la solution sp_translator

Les comportements non evidents a connaitre pour exploiter la solution.

## Traduction

1. **Retraduire = supprimer d'abord.** Une copie au titre normal n'est jamais
   ecrasee (protection des corrections des relecteurs). Pour retraduire :
   supprimer la copie dans son dossier de langue, puis relancer.
2. **Reprise automatique.** Copie en statut `Pending`/`Failed`, ou au titre de
   bapteme "Translate to..." = traduction inachevee -> reprise automatique au
   prochain passage, sans suppression.
3. **La copie est un instantane.** Les modifications ulterieures de la source
   ne se propagent pas aux copies (supprimer + retraduire).
4. **Les copies naissent en brouillon** (`PUBLISH_AFTER_TRANSLATION=false`) :
   invisibles des visiteurs tant que non publiees. Circuit de relecture voulu.
5. **La langue cible doit etre activee** dans les parametres de langue du site,
   sinon echec explicite.
6. **Titre de bapteme** : SharePoint nomme les copies "Translate to X : titre" ;
   la fonction l'ecrase par le titre source traduit. Prefixe visible =
   traduction inachevee (voir regle 2).
7. **Conditions d'affichage du bouton** : bibliotheque des pages uniquement,
   page source (jamais une copie), pas en creation/edition (Mode=Edit,
   _layouts). Visible par tous les utilisateurs (pas de filtre de droits).
8. **Jamais touche** : liens, images, IDs, filtres, parametres techniques.
   Texte a proteger de la traduction : attribut `translate="no"`.
   Commentaires : seul l'etat active/desactive est replique.

## Actualites multilingues

9.  **MosaicLanguage : manuel sur les sources, automatique sur les copies.**
    Toute page/actualite creee a la main doit recevoir sa langue (ex. `fr`)
    dans les details de la page ; les copies traduites sont stampees par le
    code. Source sans valeur = invisible des webparts filtres.
10. **Le filtre du webpart Actualites se configure une fois** sur l'accueil
    source ; la valeur est adaptee automatiquement sur les accueils traduits.
11. **Publiee, sinon rien** : seules les pages publiees apparaissent dans le
    webpart Actualites (+ delai d'indexation de quelques minutes).
12. **Bascule de langue native** : un utilisateur voit les copies publiees de
    sa langue de profil M365 ; sinon la source. La propagation d'un changement
    de langue de profil peut prendre des heures (piege de demo).

## Exploitation / deploiement

13. **Colonnes a creer sur chaque site** : `MosaicLanguage` (Choix) et
    `MosaicTranslationStatus` (Choix : Pending/Completed/Failed) sur la
    bibliotheque des pages. Absentes = mode degrade silencieux, sans tracage.
14. **Premier appel lent apres deploiement** (1-3 min, rechargement PnP).
15. **Gros lots** : au-dela d'environ 8 minutes, reponse 202 et reprise
    automatique par le dialogue avec les pages restantes.
16. **"The page hasn't been translated." est un message generique** du dialogue
    SPFx ; la vraie erreur est en console navigateur (ligne
    `HttpTriggerSpTranslator : <code> {...}`).
17. **Le .sppkg est lie a l'environnement** : URL de fonction gravee (rebuild
    par client), CORS par domaine, version de paquet a incrementer et hash du
    bundle a verifier dans le manifeste pour savoir ce qui tourne.
18. **"Page en cours de modification"** = session d'edition ouverte quelque
    part (souvent son propre onglet). Diagnostic : `.\check-page-locks.ps1`.

Outils projet : `repair-canvas.ps1` (reparation canvas fond de section),
`check-page-locks.ps1` (verrous), `inspect_homepage.ps1` (composants de page).
