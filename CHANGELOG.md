# Changelog

Toutes les évolutions notables de l'application sont documentées ici.

Le format suit les principes de [Keep a Changelog](https://keepachangelog.com/fr/), et le
numéro de version suit le [Semantic Versioning](https://semver.org/lang/fr/) (MAJEUR.MINEUR.CORRECTIF) :
un numéro **MINEUR** augmente quand une fonctionnalité est ajoutée, un **CORRECTIF** quand un bug est corrigé.

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
