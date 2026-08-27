# Changelog

Toutes les évolutions notables de l'application sont documentées ici.

Le format suit les principes de [Keep a Changelog](https://keepachangelog.com/fr/), et le
numéro de version suit le [Semantic Versioning](https://semver.org/lang/fr/) (MAJEUR.MINEUR.CORRECTIF) :
un numéro **MINEUR** augmente quand une fonctionnalité est ajoutée, un **CORRECTIF** quand un bug est corrigé.

## [1.13.0] - 2026-08-27

### Ajouté
- Nouveau champ **Variante** sur les exercices d'un programme et d'un modèle de séance (ex. "unilatéral", "côté par côté") — repris dans l'export PDF, l'export Excel et la feuille de séance.
- Nouveau bouton **"Détail par série..."** : permet, exercice par exercice, de renseigner des répétitions/charge/récup différentes pour chaque série (ex. pyramide 15/20/25 répétitions), en plus de la ligne globale qui reste le mode par défaut. Le détail par série se propage automatiquement quand un modèle de séance est utilisé pour créer une séance, et s'affiche de façon compacte dans les exports ("15 / 20 / 25") quand il est renseigné.

### Corrigé
- La suppression d'un programme ou d'une séance ne nettoyait plus les séances réalisées associées depuis le renommage de cette table (v1.10.0) — corrigé (laissait des lignes orphelines en base, sans impact visible pour l'instant car la fonctionnalité venait d'être introduite).

## [1.12.0] - 2026-08-27

### Ajouté
- Les champs séries, charge et récup d'un exercice (dans un programme et dans un modèle de séance) acceptent désormais des fourchettes libres, ex. "3-4" séries, "35-40 kg" charge, "20-30" secondes de récup, en plus des valeurs précises — pour donner une indication au client plutôt qu'un chiffre exact. Nouveau champ "Charge" (poids visé), absent jusqu'ici, ajouté sur les exercices du programme, des modèles de séance, de la feuille de séance (prévue/réalisée) et des séances réalisées ; répercuté dans l'export PDF, l'export Excel et l'export de la feuille de séance.

## [1.11.0] - 2026-08-27

### Ajouté
- Écran Clients : bouton "Supprimer définitivement..." pour effacer complètement un client et toutes ses données (devis, commandes, échéances, programmes, séances réalisées, plans nutrition, roadmap, suivi quotidien, journal alimentaire, questionnaires...), utile pour nettoyer les fiches de test. Une sauvegarde de la base est créée automatiquement juste avant, au cas où. Action irréversible, distincte de "Archiver" qui reste la solution recommandée pour un client qui arrête le coaching (ses données restent consultables).

## [1.10.0] - 2026-08-27

### Ajouté
- Écran Programmes : bouton "Exporter la feuille de séance..." — génère un Excel avec, pour chaque exercice de chaque séance du programme, les valeurs prévues en référence et des colonnes vides à remplir par le client (date, séries/répétitions/récup/tempo réellement faits, note).
- Écran Suivi > nouvel onglet **Séances réalisées** : bouton pour réimporter la feuille remplie par le client, historique des séances réalisées par client (date, séance, programme), et détail exercice par exercice en cliquant sur une ligne. Réimporter un fichier déjà traité met à jour les données existantes au lieu de les dupliquer.

## [1.9.1] - 2026-08-27

### Corrigé
- Sur certains écrans (Clients, Administratif, Bibliothèques...), le contenu qui dépassait la hauteur de la fenêtre était inaccessible : aucune barre de défilement n'existait nulle part dans l'application. C'était particulièrement visible sur un ordinateur portable en 1920×1080 avec la mise à l'échelle Windows activée (125 % ou 150 %), qui réduit la hauteur réellement disponible. L'application s'ouvre maintenant maximisée par défaut, et une barre de défilement verticale apparaît automatiquement dès qu'un écran dépasse la hauteur visible.

## [1.9.0] - 2026-08-27

### Ajouté
- Bibliothèques > Exercices : possibilité d'attacher une image du mouvement à chaque exercice (bouton "Choisir une image...", aperçu affiché dans la fiche).
- Export PDF d'un programme sportif : chaque exercice affiche désormais sa miniature (si une image a été attachée) et un lien cliquable "▶ Video" vers la vidéo de démonstration (si renseignée) — pensé pour être consulté sur mobile pendant la séance.
- Export Excel d'un programme sportif : ajout de la colonne "Lien video".

## [1.8.0] - 2026-08-27

### Ajouté
- Écran Suivi > nouvel onglet **Journal alimentaire** : importe directement l'export CSV "Food Diary Report - Detailed Report" de FatSecret envoyé par le client (calories, lipides, dont saturés, glucides, fibres, sucres, protéines, sodium, cholestérol, potassium par jour). Seuls les totaux quotidiens sont importés ; réimporter un fichier met à jour les jours déjà présents au lieu de les dupliquer.

## [1.7.0] - 2026-08-27

### Ajouté
- Un devis transformé en commande disparaît automatiquement de la liste des Devis (il reste consultable dans Commandes).
- Une commande passe automatiquement au statut "Terminée" dès que toutes ses échéances sont marquées payées — corrige les commandes déjà entièrement payées qui restaient bloquées sur "Active".
- Gestion manuelle du statut d'une commande (Terminée / Annulée / Réactiver), comme pour les devis.
- Filtres par statut sur les onglets Devis (Tous / En attente / Accepté / Refusé) et Commandes (Toutes / Active / Terminée / Annulée), en plus du filtre déjà existant sur Paiements/Échéances.

## [1.6.0] - 2026-08-27

### Ajouté
- Écran Programmes et onglet Modèles de séance : les lignes d'exercice (séries, répétitions, récup, tempo, notes) peuvent maintenant être modifiées directement en sélectionnant la ligne dans le tableau, sans avoir à la supprimer et la recréer. Le bouton "Ajouter" devient "Enregistrer" (crée une nouvelle ligne si rien n'est sélectionné, modifie la ligne sélectionnée sinon) et un bouton "Nouveau" permet de revenir à l'ajout. Ça s'applique aussi bien aux exercices ajoutés à la main qu'à ceux copiés depuis un modèle de séance.

## [1.5.0] - 2026-08-27

### Ajouté
- Import de la roadmap hebdo (Suivi > Roadmap hebdo) : télécharge un modèle Excel vierge, remplis-le (une ligne par semaine), puis importe-le — comme pour le suivi quotidien. Réimporter un fichier met à jour les semaines déjà existantes (même numéro) au lieu de les dupliquer.

## [1.4.0] - 2026-08-27

### Ajouté
- Modèles de séance réutilisables (Bibliothèques > Modèles de séance) : créer une fois une séance type (ex. "Haut du corps", "Full body") avec sa liste d'exercices, puis l'insérer dans le programme de n'importe quel client via le nouveau bouton "Créer depuis un modèle..." sur l'écran Programmes, sans avoir à la reconstruire à chaque fois.

## [1.3.2] - 2026-08-26

### Ajouté
- Bibliothèques initiales (exercices, aliments, compléments) embarquées dans le dépôt (`DonneesInitiales/`) et chargées automatiquement au tout premier lancement de l'application. Auparavant, comme `Data/` est exclu de git pour protéger les données réelles, un téléchargement frais démarrait avec des bibliothèques vides.

## [1.3.1] - 2026-08-26

### Corrigé
- L'application pouvait échouer silencieusement au démarrage ("rien ne se passe" au double-clic) après un téléchargement depuis internet : Windows bloque les fichiers extraits d'un zip téléchargé, et l'erreur restait invisible (fenêtre masquée). L'application se débloque désormais automatiquement à chaque lancement, et toute erreur de démarrage affiche maintenant un message explicite au lieu de se fermer sans rien afficher.

## [1.3.0] - 2026-08-26

### Ajouté
- Import du questionnaire pré-coaching : complète désormais automatiquement le Téléphone, l'Email et les Objectifs de la fiche client à partir des réponses (colonnes optionnelles à mapper), **uniquement si le champ est encore vide** — aucune donnée déjà saisie n'est jamais écrasée.
- Les numéros de téléphone importés sont lus en texte brut pour éviter la perte d'un éventuel 0 initial.

## [1.2.0] - 2026-08-26

### Ajouté
- Tableau de bord (écran d'accueil) : clients actifs, paiements en attente/retard, prochaines échéances, clients sans bilan récent.
- Écran Suivi : import des réponses aux questionnaires Google Forms (pré-coaching / bilan) avec assistant de correspondance des colonnes et rattachement automatique au client par email/nom.
- Import du suivi quotidien du client depuis un modèle Excel téléchargeable, sans écraser l'historique existant.
- Graphique d'évolution du poids par client.
- Roadmap hebdomadaire par client (suivi de phase semaine par semaine).
- Numéro de version affiché dans la fenêtre et l'écran Outils.

## [1.1.0] - 2026-08-26

### Ajouté
- Écran Programmes : création de séances et association d'exercices depuis la bibliothèque (séries, répétitions, récupération, tempo, notes), réordonnancement des séances.
- Écran Nutrition : plans nutritionnels par client, types de jour, repas, calcul automatique des kcal/macros par repas et par jour.
- Export PDF (via Edge/Chrome en tâche de fond) et Excel des programmes et plans nutrition.

## [1.0.0] - 2026-08-26

### Ajouté
- Version initiale : fiches clients (créer/éditer/archiver).
- Administratif : devis, commandes (mensuel/hebdomadaire/one-shot), génération automatique de l'échéancier, suivi des paiements.
- Bibliothèques réutilisables : exercices, aliments (macros), compléments.
- Import initial des bibliothèques depuis le fichier Excel historique du coach, export Excel des bibliothèques.
- Sauvegarde manuelle de la base de données.
