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

- **Clients** : créer et gérer les fiches de tes clients (identité, contact, objectifs, notes). La case à cocher en haut permet d'afficher aussi les clients archivés.
- **Administratif** :
  - *Devis* : crée un devis pour un client, puis marque-le "Accepté" ou "Refusé". Un devis accepté peut être transformé en commande — une fois transformé, il disparaît automatiquement de cette liste (il reste consultable dans Commandes). Filtre par statut disponible.
  - *Commandes* : crée une commande (mensuelle, hebdomadaire ou paiement unique) avec un montant total. L'échéancier de paiement est généré automatiquement. Le statut passe tout seul à "Terminée" dès que toutes les échéances sont payées ; il peut aussi être changé à la main (Terminée / Annulée / Réactiver). Filtre par statut disponible.
  - *Paiements / Échéances* : vue de toutes les échéances, avec filtre (en attente / en retard / payées) et bouton pour marquer une échéance comme payée.
- **Bibliothèques** : tes listes réutilisables d'exercices, d'aliments (avec leurs valeurs nutritionnelles), de compléments et de **modèles de séance** (ex. "Haut du corps", "Full body"). Recherche, ajout, modification, suppression, et export en Excel.
- **Programmes** : choisis un client, crée un programme, ajoute des séances (avec les flèches ▲▼ pour les réordonner) — soit vides, soit directement à partir d'un modèle de séance via "Créer depuis un modèle...", puis ajoute/ajuste des exercices à chaque séance en piochant dans ta bibliothèque (séries, répétitions, récup, tempo, notes). Clique sur une ligne du tableau pour la modifier directement (plus besoin de la supprimer et la recréer) ; ça marche aussi pour les exercices copiés depuis un modèle de séance. Exporte le programme en PDF ou Excel pour l'envoyer au client.
- **Nutrition** : choisis un client, crée un plan, ajoute un ou plusieurs types de jour (ex. "Jour haut" / "Jour bas"), puis des repas dans chaque type de jour, puis des aliments dans chaque repas avec leur quantité — les kcal/macros et les totaux se calculent automatiquement. Exporte le plan en PDF ou Excel.
- **Tableau de bord** : l'écran d'accueil — nombre de clients actifs, paiements en attente/retard, prochaines échéances, et clients sans bilan depuis 14 jours ou plus.
- **Suivi** : choisis un client, puis :
  - *Questionnaires* : boutons pour envoyer les questionnaires pré-coaching/bilan (ouvre le Google Form dans le navigateur), bouton pour importer les réponses exportées depuis Google Sheets (un assistant te demande une fois quelles colonnes correspondent à la date/au nom/à l'email — mémorisé pour la prochaine fois). Les réponses sont rattachées automatiquement au bon client ; celles qui ne le sont pas peuvent être assignées manuellement via la case à cocher "Afficher les réponses non rattachées".
  - *Tracking quotidien* : télécharge un modèle Excel vierge à envoyer à ton client, puis importe le fichier qu'il t'a renvoyé rempli — sans écraser l'historique existant. Un graphique d'évolution du poids s'affiche automatiquement.
  - *Roadmap hebdo* : suivi de phase semaine par semaine (comme l'onglet ROADMAP de ton ancien fichier Excel). Se remplit à la main semaine par semaine, ou peut être importée d'un coup : télécharge le modèle Excel vierge, remplis-le (une ligne par semaine), puis importe-le — réimporter met à jour les semaines déjà présentes au lieu de les dupliquer.
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

Pistes d'amélioration possibles pour la suite (non demandées initialement, à discuter si besoin) : import du détail des séances réellement effectuées (charges/répétitions exercice par exercice), export PDF de la roadmap, graphiques supplémentaires (adhésion, sommeil) dans l'historique client.
