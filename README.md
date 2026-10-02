# Suivi Coaching

Application de gestion de l'activité de coaching (clients, devis, paiements, bibliothèques d'exercices/aliments/compléments).

## Lancer l'application

Double-clique simplement sur le fichier **`Lancer.bat`**.

La première fois, Windows peut afficher un avertissement de sécurité ("Windows a protégé votre ordinateur" ou similaire) : clique sur **"Informations complémentaires"** puis **"Exécuter quand même"**. C'est normal pour une application non signée numériquement, ce n'est pas un virus.

L'application s'ouvre dans une fenêtre. Aucune installation n'est nécessaire, aucune connexion internet n'est requise pour l'utiliser au quotidien.

Au tout premier lancement (quand aucune base de données n'existe encore), les bibliothèques d'exercices/aliments/compléments sont automatiquement pré-remplies avec les données déjà fournies — un message le confirme.

### Si un double-clic sur "Lancer" ne fait rien du tout

Cela arrive après un **téléchargement depuis internet** (ex. bouton "Download ZIP" sur GitHub) : Windows marque tous les fichiers extraits comme "provenant d'internet" et peut en bloquer silencieusement certains, sans afficher de message. Depuis la version 1.3.1, l'application essaie de se débloquer automatiquement à chaque lancement — si malgré tout rien ne se passe :

1. Avant d'extraire le zip : clique droit sur le fichier `.zip` téléchargé → **Propriétés** → coche **"Débloquer"** en bas → OK. Extrais ensuite normalement.
2. Si le dossier est déjà extrait : ouvre PowerShell dans le dossier `SuiviCoachingApp` et lance `Get-ChildItem -Recurse | Unblock-File`.
3. Vérifie aussi qu'un antivirus ne met pas le dossier en quarantaine.

## Prise en main

L'application se navigue depuis le menu à gauche :

- **Clients** : créer et gérer les fiches de tes clients (identité, contact, objectifs, notes). La case à cocher en haut permet d'afficher aussi les clients archivés. "Archiver ce client" (recommandé pour un client qui arrête) le retire de la liste active sans rien effacer ; "Supprimer définitivement..." efface le client et toutes ses données dans toute l'application — irréversible, pratique pour nettoyer une fiche de test.
- **Administratif** :
  - *Devis* : crée un devis pour un client, puis marque-le "Accepté" ou "Refusé". Un devis accepté peut être transformé en commande — une fois transformé, il disparaît automatiquement de cette liste (il reste consultable dans Commandes). Filtre par statut disponible.
  - *Commandes* : crée une commande (mensuelle, hebdomadaire ou paiement unique) avec un montant total. L'échéancier de paiement est généré automatiquement. Le statut passe tout seul à "Terminée" dès que toutes les échéances sont payées ; il peut aussi être changé à la main (Terminée / Annulée / Réactiver). Filtre par statut disponible.
  - *Paiements / Échéances* : vue de toutes les échéances, avec filtre (en attente / en retard / payées) et bouton pour marquer une échéance comme payée.
- **Bibliothèques** : tes listes réutilisables d'exercices, d'aliments (avec leurs valeurs nutritionnelles), de compléments et de **modèles de séance** (ex. "Haut du corps", "Full body"). Recherche, ajout, modification, suppression, et export en Excel. Un exercice peut avoir une image du mouvement (bouton "Choisir une image...", aperçu affiché) en plus du lien vidéo — les deux ressortent dans l'export PDF d'un programme.
- **Programmes** : choisis un client, crée un programme, ajoute des séances (avec les flèches ▲▼ pour les réordonner) — soit vides, soit directement à partir d'un modèle de séance via "Créer depuis un modèle...", puis ajoute/ajuste des exercices à chaque séance en piochant dans ta bibliothèque (séries, répétitions, récup, tempo, RIR, notes). Le **RIR** (Reps In Reserve, les répétitions gardées en réserve avant l'échec, ex. « 2 » ou « 1-2 ») se saisit à côté du tempo, aussi dans les modèles de séance (il est recopié quand on crée une séance depuis un modèle). Les cases **Tempo** et **RIR** à droite des boutons d'export choisissent si ces colonnes apparaissent dans le PDF, l'Excel du programme et la feuille de séance (choix mémorisé). Les champs séries, charge et récup acceptent des fourchettes libres (ex. "3-4" séries, "35-40 kg", "20-30" s) en plus des valeurs précises, pour donner une indication plutôt qu'un chiffre exact. Le champ Variante précise comment faire l'exercice (ex. "unilatéral", "côté par côté"). À lajout dun exercice (ici comme dans les modèles de séance), la fenêtre "Détail par série" souvre automatiquement, pré-remplie avec le nombre de séries indiqué (le maximum pour une fourchette : "3-4" → 4 lignes) et les valeurs saisies ; clique sur Annuler pour garder des valeurs identiques à chaque série. Ensuite, la case "Détail par série" rouvre cette fenêtre : modifie seulement celles qui doivent différer (pyramide) — le tableau affiche alors le détail complet ("15 / 20 / 25") au lieu de la valeur globale. Clique sur une ligne du tableau pour la modifier directement (plus besoin de la supprimer et la recréer) ; ça marche aussi pour les exercices copiés depuis un modèle de séance (variante et détail par série compris). Chaque séance peut se voir attribuer un jour de la semaine (menu déroulant sous la liste des séances) : un tableau récapitulatif "Semaine type" (Lundi → Dimanche, "Repos" les jours sans séance) apparaît alors en haut de l'export PDF du programme. Exporte le programme en PDF (avec miniature de l'exercice et lien "▶ Video" cliquable — pratique à consulter sur le mobile du client pendant sa séance) ou en Excel pour l'envoyer au client (même présentation que l'onglet TRAINING d'origine, sans les blocs de suivi). Le PDF reprend la présentation de l'onglet TRAINING du fichier d'origine (violet, une ligne par série, nom de la séance en bande verticale). Le bouton "Exporter la feuille de séance..." génère un Excel identique à cet onglet TRAINING, avec un onglet par séance : le programme à gauche (une ligne par série, figé à l'écran) et 12 blocs "SEMAINE 1…12" à droite où le client note, chaque semaine, la DATE puis ses répétitions et charges série par série (et une note par exercice). Chaque case REPS se colore toute seule selon les répétitions prévues pour la série : rouge en dessous de la fourchette (ex. moins de 9 pour « 9-12 »), vert dedans, orange au-dessus (une valeur unique comme « 10 » compte comme fourchette 10-10, « 12+ » n'a pas de maximum ; « max » ou « AMRAP » restent sans couleur). À l'impression, chaque séance tient sur la hauteur d'une page, avec le programme répété sur chaque page (semaines 1 à 6 puis 7 à 12). Le même fichier contient un onglet **TRACKING** pour le suivi quotidien sur 52 semaines (une ligne par jour, dates déjà remplies à partir de la date de début du programme) : le client n'a qu'un seul fichier à remplir. À réimporter ensuite via Suivi > Séances réalisées ou Suivi > Tracking quotidien (les deux boutons font la même chose) : chaque bloc daté devient une séance réalisée et chaque jour rempli du TRACKING va dans le tracking quotidien (les anciennes feuilles et anciens modèles de suivi restent importables).
- **Nutrition** : choisis un client, crée un plan, ajoute un ou plusieurs types de jour (ex. "Jour haut" / "Jour bas"), puis des repas dans chaque type de jour, puis des aliments dans chaque repas avec leur quantité — les kcal/macros et les totaux se calculent automatiquement. Exporte le plan en PDF (présentation de l'onglet NUTRITION d'origine : un "PLAN JOURNALIER" par type de jour avec totaux par repas et encadré RECAP) ou en Excel (même présentation).
- **Tableau de bord** : l'écran d'accueil — nombre de clients actifs, paiements en attente/retard, prochaines échéances, et clients sans bilan depuis 14 jours ou plus.
- **Suivi** : choisis un client, puis :
  - *Questionnaires* : boutons pour envoyer les questionnaires pré-coaching/bilan (ouvre le Google Form dans le navigateur), bouton pour importer les réponses exportées depuis Google Sheets, au format Excel (.xlsx) ou CSV (un assistant te demande une fois quelles colonnes correspondent à la date/au nom/à l'email — mémorisé pour la prochaine fois). Les réponses sont rattachées automatiquement au bon client ; celles qui ne le sont pas peuvent être assignées manuellement via la case à cocher "Afficher les réponses non rattachées".
  - *Tracking quotidien* : télécharge un modèle Excel vierge (onglet TRACKING sur 52 semaines, dates déjà remplies à partir du lundi de la semaine en cours) à envoyer à ton client, puis importe le fichier qu'il t'a renvoyé rempli — sans écraser l'historique existant. Le plus simple : envoie plutôt la feuille de séance (écran Programmes), qui contient déjà cet onglet TRACKING. Les jours laissés vides sont ignorés, et une case laissée vide reste vide (elle n'est pas enregistrée comme 0). Un graphique d'évolution du poids s'affiche automatiquement.
    - **Présentation de l'onglet TRACKING** (même présentation que l'onglet TRACKING d'origine) : un bloc par semaine avec ses en-têtes, DATE et JOUR déjà remplis (semaines du lundi au dimanche), puis une ligne MOYENNE de la semaine. Les colonnes sont regroupées par thème (SOMMEIL, NUTRITION, TRAINING, SANTÉ), chaque thème étant séparé par une bande de couleur et le fond alternant d'un thème à l'autre. Les notes de 1 à 5 ont une liste déroulante et se colorent du rouge au vert. À droite, une seule case **BILAN** par semaine, cliquable vers le formulaire de bilan.
    - **Choix des colonnes** : avant chaque export (feuille de séance ou modèle vierge), une fenêtre te laisse cocher les colonnes à inclure (poids, qualité/durée du sommeil, coucher, lever, énergie, adhésion, non tracké, digestion, séance, nb pas, cardio, motivation, RC repos, pressions systolique/diastolique, notes du jour), activer ou non la case BILAN et indiquer le lien du formulaire de bilan. Ton choix est mémorisé pour les exports suivants. L'import retrouve chaque colonne par son en-tête, quelles que soient les colonnes choisies.
  - *Roadmap hebdo* : suivi de phase semaine par semaine (comme l'onglet ROADMAP de ton ancien fichier Excel). Se remplit à la main semaine par semaine, ou peut être importée d'un coup : télécharge le modèle Excel vierge, remplis-le (une ligne par semaine), puis importe-le — réimporter met à jour les semaines déjà présentes au lieu de les dupliquer.
  - *Journal alimentaire* : importe directement le CSV "Food Diary Report - Detailed Report" que ton client t'envoie depuis FatSecret (Journal alimentaire > Exporter, sur la période voulue) — calories et macros (lipides, glucides, protéines, fibres, sucres, sodium, cholestérol, potassium) par jour. Réimporter un fichier met à jour les jours déjà présents au lieu de les dupliquer.
  - *Séances réalisées* : réimporte la "feuille de séance" remplie par le client (générée depuis l'écran Programmes ; son onglet TRACKING est importé en même temps dans le tracking quotidien) et consulte l'historique de ce qu'il a réellement fait (date, séance, puis détail exercice par exercice en cliquant sur une ligne).
- **Outils** :
  - *Importer les bibliothèques depuis un fichier Excel* : pour récupérer en un clic les exercices, aliments et compléments déjà présents dans ton fichier "SUIVI 2.0.xlsx" (ou tout fichier construit sur le même modèle). Les éléments déjà présents dans l'application ne sont pas dupliqués.
  - *Sauvegarder maintenant* : crée une copie de sécurité de toutes tes données.

## Tes données

Toutes tes données sont stockées uniquement sur ton ordinateur, dans le fichier :

```
Data\suivi_coaching.db
```

Rien n'est envoyé sur internet. Pense à cliquer régulièrement sur **"Sauvegarder maintenant"** (dans Outils) : cela crée une copie datée dans le dossier `Data\Backups`, à conserver idéalement aussi sur une clé USB ou un cloud (Google Drive, OneDrive...) pour être protégé en cas de problème avec l'ordinateur.

## Export PDF

L'export PDF des programmes et plans nutrition utilise Microsoft Edge ou Google Chrome (déjà installés sur la quasi-totalité des ordinateurs Windows récents) pour générer le fichier — l'application ne les installe pas, elle les détecte automatiquement. Si aucun des deux n'est présent sur ton ordinateur, un message te l'indique et tu peux utiliser l'export Excel à la place.

## Mettre à jour l'application

Le numéro de version actuellement installé est affiché en haut de la fenêtre et dans l'écran **Outils**. Le détail de ce qui a changé à chaque version est dans le fichier `CHANGELOG.md`.

Si tu reçois une nouvelle version (nouveau dossier ou fichiers à copier) :

1. **Ne touche jamais au dossier `Data`** : c'est là que sont toutes tes données (clients, paiements, suivis...). Il ne fait pas partie de la mise à jour et ne doit jamais être remplacé, supprimé ou écrasé.
2. Remplace uniquement les autres fichiers/dossiers (`app.ps1`, `Modules\SuiviCoaching`, `UI`, `Lancer.bat`, `README.md`, `CHANGELOG.md`) par les nouveaux.
3. Relance l'application avec `Lancer.bat` comme d'habitude : tes clients, devis, programmes, etc. sont toujours là, seule l'application elle-même a été mise à jour.

En cas de doute avant une mise à jour, clique sur **"Sauvegarder maintenant"** (dans Outils) : ça ne coûte rien et ça protège tes données.

## En cas de souci

- Si l'application ne s'ouvre pas du tout : vérifie que tu n'as pas déplacé ou renommé un des dossiers/fichiers à l'intérieur du dossier `SuiviCoachingApp` (l'application a besoin que tout reste au même endroit les uns par rapport aux autres).
- Si un message d'erreur apparaît pendant l'utilisation, il est affiché en français dans une fenêtre : note ce qu'il dit, ça aide à corriger le problème.

## Suite du projet

Cette version (V1 + V2 + V3) couvre l'intégralité du cahier des charges initial : fiches clients, devis/commandes/paiements, bibliothèques, programmes sportifs et plans nutrition avec export PDF/Excel, import des questionnaires et du suivi client, roadmap hebdo, et tableau de bord.

Pistes d'amélioration possibles pour la suite (non demandées initialement, à discuter si besoin) : export PDF de la roadmap, graphiques supplémentaires (adhésion, sommeil) dans l'historique client.
