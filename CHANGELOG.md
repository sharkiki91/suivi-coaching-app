# Changelog

Toutes les évolutions notables de l'application sont documentées ici.

Le format suit les principes de [Keep a Changelog](https://keepachangelog.com/fr/), et le
numéro de version suit le [Semantic Versioning](https://semver.org/lang/fr/) (MAJEUR.MINEUR.CORRECTIF) :
un numéro **MINEUR** augmente quand une fonctionnalité est ajoutée, un **CORRECTIF** quand un bug est corrigé.

## [1.25.1] - 2026-10-10

### Corrigé
- **Administratif – nouvelle commande / nouveau devis** : la liste des clients restait vide si on allait dans Administratif sans être passé par l'écran Clients (au démarrage, l'appli ouvre le tableau de bord). Elle est maintenant chargée à chaque ouverture de l'écran Administratif.
- **TRACKING – date de la semaine 1** : elle était imposée (lundi de la semaine de début du programme, ou de la semaine en cours) et chaque date était figée dans l'Excel. La fenêtre d'options avant l'export (feuille de séance et modèle de suivi quotidien) a maintenant un champ **1er jour de la semaine 1**, pré-rempli comme avant mais modifiable librement (n'importe quel jour de la semaine). Dans le fichier, seule cette 1re date est une valeur : les autres dates (+1 jour) et les JOUR sont des formules, donc changer cette seule case dans Excel / Google Sheets décale tout le tracking.

## [1.25.0] - 2026-10-08

### Ajouté
- **Nutrition – total de la journée** : bande violette en haut de l'écran Nutrition avec le total kcal / protéines / glucides / lipides / fibres de tous les repas du type de jour sélectionné (jusqu'ici seul le total par repas était visible dans l'appli ; le PDF l'avait déjà).
- **Nutrition – modifier et réordonner** : cliquer une ligne d'un repas la charge dans le formulaire, "Modifier la ligne" enregistre le nouvel aliment ou la nouvelle quantité ; flèches ▲▼ pour déplacer une ligne dans le repas et pour changer l'ordre des repas (ex. dîner créé avant le déjeuner). L'ordre est repris dans le PDF et l'Excel.
- **Recettes** : nouvel onglet Bibliothèques > Recettes (intitulé, notes de préparation, ingrédients avec quantités, totaux calculés). "Ajouter la recette" dans un repas copie ses ingrédients (× nombre de portions) ; dans l'appli, le PDF et l'Excel, l'intitulé de la recette s'affiche avec ses ingrédients en retrait en dessous. Une recette se déplace d'un bloc et peut être retirée d'un coup.
- **Tableau d'équivalences** : nouvel onglet Bibliothèques > Équivalences, pré-rempli avec le bloc EQUIVALENCES de l'onglet NUTRITION d'origine. Les quantités équivalentes sont calculées depuis la bibliothèque d'aliments (même apport en calories pour les aliments peu caloriques, sinon en protéines / glucides / lipides selon l'aliment de référence) ; bouton "Suggérer depuis la bibliothèque". Le tableau est ajouté à la fin du PDF et de l'Excel nutrition (case "Avec le tableau d'équivalences", cochée par défaut, mémorisée).
- **Supersets** : bouton "Superset avec le suivant" (Programmes et modèles de séance) pour enchaîner deux exercices (ou plus) sans récup. Numérotation 3A / 3B dans l'appli, numéro en bleu (#3C78D8, couleur de la légende SUPERSET d'origine) dans le PDF, l'Excel et la feuille de séance, avec la légende sous la séance. Recopié lors d'une création de séance depuis un modèle.
- **Roadmap au format d'origine** : "Télécharger la roadmap (Excel)" génère l'onglet ROADMAP comme dans le fichier du coach (en-têtes violets sur deux lignes, DEPENSE → CARDIO / PAS, SEM en lavande, 52 semaines, dates de 7 en 7), pré-rempli avec les semaines du client. Nouveau champ **Précision dépense** (colonne PRECISION DEPENSE). L'import relit ce format, l'ancien modèle, et les onglets ROADMAP des fichiers clients existants (colonnes reconnues par leur titre).

### Corrigé
- Roadmap : une case numérique laissée vide (poids moyen, dépense, cardio, pas) était enregistrée comme 0 ; elle reste maintenant vide. Une ligne vide d'un fichier importé n'efface plus une semaine déjà saisie.
- Deux messages de l'écran Programmes / modèles s'affichaient avec des accents illisibles (« SÃ©lectionne... »).

## [1.24.0] - 2026-10-04

### Ajouté
- **Nombre de semaines réglable** : la fenêtre d'options qui s'ouvre avant l'export (feuille de séance, modèle de tracking) permet de choisir le nombre de blocs SEMAINE du tableau des séances (1 à 52, 12 par défaut) et le nombre de semaines de l'onglet TRACKING (1 à 104, 52 par défaut). Le choix est mémorisé. L'impression, les couleurs des REPS et le réimport s'adaptent.
- **Réordonner les exercices d'une séance** : flèches ▲▼ sous le formulaire (écran Programmes et modèles de séance) pour monter ou descendre l'exercice sélectionné. Un exercice oublié peut ainsi être ajouté en dernier puis remis à sa place sans tout recréer. L'ordre est repris dans le PDF, l'Excel et la feuille de séance.
- **Récap des séries par groupe musculaire** (selon le muscle ciblé de chaque exercice de la bibliothèque), par séance et sur la semaine (toutes les séances du programme, chacune une fois). Une fourchette « 3-4 » compte pour 4, et le détail par série compte série par série. Le récap apparaît :
  - dans l'écran Programmes, sous le formulaire, mis à jour à chaque modification ;
  - dans le PDF du programme, sous la « Semaine type » ;
  - dans un onglet **RECAP SERIES** de l'Excel du programme et de la feuille de séance.

### Corrigé
- **Modifier un exercice déjà créé n'avait pas d'effet visible** quand il avait un détail par série : les reps, la charge ou la récup modifiées sur la ligne étaient enregistrées, mais le détail par série (ce qui s'affiche et s'exporte) gardait les anciennes valeurs. Les champs modifiés sont maintenant appliqués à toutes les séries du détail (les champs non modifiés gardent leur détail, ex. une pyramide de charge), et un nouveau nombre de séries ajoute ou retire des séries. La ligne reste sélectionnée après « Enregistrer », pour voir tout de suite ce qui a été enregistré.
- Changer l'exercice d'une ligne existante dans la liste déroulante n'était pas enregistré : c'est corrigé (programmes et modèles).

## [1.23.0] - 2026-10-02

### Ajouté
- **RIR (Reps In Reserve)** : nouveau champ sur chaque exercice d'une séance et d'un modèle de séance, juste à côté du tempo (répétitions gardées en réserve avant l'échec, texte libre, ex. « 2 » ou « 1-2 »). Il est affiché dans les tableaux d'exercices et recopié lors d'une création de séance depuis un modèle. Les colonnes `rir` sont ajoutées automatiquement à la base.
- **Colonnes Tempo et RIR au choix dans les exports du programme** : deux cases « Tempo » et « RIR » à droite des boutons d'export (écran Programmes) choisissent si ces colonnes apparaissent dans le PDF du programme, l'Excel du programme et la feuille de séance. Le choix est mémorisé et les deux sont cochées par défaut. La zone répétée à l'impression de la feuille de séance et le réimport s'adaptent aux colonnes présentes.

## [1.22.1] - 2026-10-02

### Corrigé
- **Plus de retours à la ligne dans les fichiers Excel** (feuille de séance, Excel du programme, onglet TRACKING, modèle de tracking, Excel du plan nutrition). Chaque colonne est élargie automatiquement selon son contenu réel (en-têtes, noms d'exercices, d'aliments, de repas, dates…), mesuré dans la police de la cellule, pour que tout tienne sur une ligne. Exemples : MUSCLE CIBLE, RÉCUP, TEMPO, et les en-têtes du TRACKING comme MOTIVATION (1-5). Les retours à la ligne voulus (notes d'un exercice sous son nom, « BILAN / SEMAINE n ») sont conservés. Les titres et consignes sur plusieurs colonnes ne sont pas concernés.

## [1.22.0] - 2026-10-02

### Ajouté
- **Choix des colonnes du tracking avant l'export** : une fenêtre s'ouvre avant d'exporter la feuille de séance ou le modèle vierge de tracking. Elle permet de cocher les colonnes à inclure, regroupées par thème, d'activer ou non la case BILAN de la semaine et de saisir le lien du formulaire de bilan. Le choix est mémorisé. Par défaut, les pressions systolique/diastolique, la durée du sommeil et les notes du jour sont décochées.
- **Feuille de séance : REPS colorées selon la fourchette prévue.** Chaque case REPS des blocs SEMAINE se colore d'elle-même quand le client saisit ses répétitions : **rouge** en dessous du minimum prévu pour la série, **vert** dans la fourchette, **orange** au-dessus du maximum (ex. pour « 9-12 » : 8 = rouge, 10 = vert, 13 = orange). Fourchettes reconnues : « 9-12 », « 9 à 12 », « 9/12 », une valeur unique (« 10 » = 10-10) et « 12+ » (pas de maximum). « max », « AMRAP » et les cases vides restent sans couleur. Fonctionne série par série, y compris avec le détail par série (pyramide).
- Nouvelles colonnes de tracking reprises de l'onglet TRACKING d'origine : **RC REPOS** (fréquence cardiaque au repos), **SÉANCE** et **NON TRACKÉ**. Elles sont enregistrées en base (colonnes ajoutées automatiquement) et visibles dans le tableau du tracking quotidien (RC repos, séance).

### Modifié
- **Onglet TRACKING** au style de l'onglet TRACKING d'origine : un bloc par semaine (en-têtes répétés, DATE + JOUR déjà remplis du lundi au dimanche, ligne MOYENNE de la semaine), colonnes regroupées par thème (SOMMEIL, NUTRITION, TRAINING, SANTÉ) séparées par une bande de couleur, fond alterné d'un thème à l'autre pour alléger la lecture, notes de 1 à 5 en liste déroulante colorée du rouge au vert, et **une seule case BILAN par semaine** (fusionnée) contenant le lien vers le formulaire de bilan.
- L'import retrouve chaque colonne par son en-tête, quelles que soient les colonnes exportées. Les fichiers des versions précédentes restent importables.

### Corrigé
- Tracking quotidien : une case laissée vide (poids, sommeil, notes…) était enregistrée comme **0**. Elle reste maintenant vide. Les lignes déjà importées ne sont pas modifiées.
- Une heure de coucher ou de lever qu'Excel avait convertie en format heure est relue correctement (ex. "06:45").
- Feuille de séance Excel : la colonne CHARGE (blocs SEMAINE et programme) est élargie et son en-tête ne passe plus à la ligne, comme pour la case DATE ("CHARGE" et une saisie comme "37.5 kg" tiennent sur une ligne).

## [1.21.0] - 2026-10-02

### Modifié
- **Feuille de séance Excel** : **12 blocs "SEMAINE 1…12"** au lieu de 6 blocs "SÉANCE" (une séance notée par semaine pendant 12 semaines). À l'impression, chaque séance tient sur la hauteur d'une page paysage, avec le programme répété à gauche de chaque page (semaines 1 à 6, puis 7 à 12).
- **Tracking quotidien sur 52 semaines, dans le même fichier** : la feuille de séance contient un onglet **TRACKING** (une ligne par jour sur 52 semaines, dates déjà remplies à partir de la date de début du programme, ou du lundi de la semaine en cours, bande S1…S52 par semaine) : le client ne remplit plus qu'un seul fichier. Le modèle vierge du tracking quotidien utilise ce même onglet (52 semaines à partir du lundi de la semaine en cours).
- **Import unique** : les boutons "Importer" de Séances réalisées et de Tracking quotidien acceptent le même fichier et rangent chaque information à sa place (blocs SEMAINE datés → séances réalisées, jours remplis du TRACKING → tracking quotidien), avec un bilan des deux. Les jours laissés vides sont ignorés ; un jour rempli sans date est signalé. Les anciennes feuilles de séance et anciens modèles de suivi restent importables.

## [1.20.1] - 2026-10-01

### Corrigé
- Feuille de séance Excel : le libellé "DATE" de chaque bloc SÉANCE passait sur deux lignes (colonne trop étroite). La colonne est élargie et le libellé reste sur une seule ligne.

## [1.20.0] - 2026-10-01

### Modifié
- **Feuille de séance Excel** (et Excel du programme) : **un onglet par séance**, nommé comme la séance, au lieu d'un seul onglet avec toutes les séances à la suite. Chaque onglet est réglé pour s'imprimer en entier sur **une seule page** paysage : un tableau de séance n'est plus jamais coupé en deux à l'impression. Le titre et la consigne sont répétés sur chaque onglet. L'import des séances réalisées lit tous les onglets (les feuilles déjà envoyées, en un seul onglet ou au format tableau, restent importables).
- **PDF du programme** : une séance n'est plus coupée entre deux pages — si elle ne tient pas en bas de la page, elle commence sur la suivante. Lignes de séries légèrement resserrées pour qu'une séance de 8 exercices tienne sur la même page que l'en-tête.

## [1.19.0] - 2026-10-01

### Modifié
- **Export Excel du plan nutrition** : même présentation que l'onglet NUTRITION d'origine (et que l'export PDF) — titre et notes du plan, un tableau "PLAN JOURNALIER" par type de jour avec la ligne TOTAL du jour (violet foncé), colonnes Repas / Aliments / Quantité / KCAL / PRO / GLU / LIP / FIB, repas en bande lavande (cellules fusionnées), totaux PRO/GLU/LIP/FIB et kcal de chaque repas à droite, et un encadré RECAP du jour à côté. Valeurs numériques (arrondies à l'unité), impression en paysage sur la largeur d'une page.

## [1.18.0] - 2026-10-01

### Modifié
- **Export Excel du programme** : même présentation que la feuille de séance (onglet TRAINING d'origine : bande violette au nom de la séance, une ligne par série, cellules fusionnées par exercice, couleurs et polices du fichier du coach), sans les 6 blocs de suivi. La date de début et les notes du programme figurent sous le titre.
- **Export PDF du plan nutrition** : nouvelle mise en page calquée sur l'onglet NUTRITION d'origine — un tableau "PLAN JOURNALIER" par type de jour avec la ligne TOTAL du jour (violet foncé), colonnes Repas / Aliments / Quantité / KCAL / PRO / GLU / LIP / FIB, repas en bande lavande, totaux PRO/GLU/LIP/FIB et kcal de chaque repas sur la droite, et un encadré RECAP du jour. Valeurs arrondies à l'unité comme dans l'original, A4 paysage.

## [1.17.0] - 2026-10-01

### Modifié
- **Export PDF du programme** : nouvelle mise en page calquée sur l'onglet TRAINING du fichier d'origine du coach — en-tête violet, tableau "Semaine type" aux couleurs du programme, une bande lavande verticale avec le nom (et le jour) de chaque séance, colonnes # / Exercice / Variante / Set / Reps / Charge / Récup / Tempo / Muscle cible / Lien, **une ligne par série**, noms en violet, séparateur gris entre les exercices, format A4 paysage, un exercice n'est jamais coupé entre deux pages.
- **Feuille de séance Excel (suivi des performances)** : reproduit l'onglet TRAINING d'origine. Pour chaque séance : le programme à gauche (une ligne par série, cellules fusionnées par exercice, mêmes couleurs/polices que l'original, figé à l'écran quand on fait défiler), et à droite 6 blocs "SÉANCE 1…6" (une case DATE, puis # / REPS / CHARGE par série et NOTES par exercice) pour noter 6 séances successives. Impression en paysage sur la largeur d'une page.
- **Import des séances réalisées** : lit ce nouveau format — chaque bloc SÉANCE daté devient une séance réalisée (répétitions/charges assemblées série par série, ex. "12 / 11 / 10 / 9"). Un bloc rempli sans date est signalé. Les anciennes feuilles (format tableau) déjà envoyées aux clients restent importables.

## [1.16.0] - 2026-10-01

### Ajouté
- Programmes et Bibliothèques > Modèles de séance : le détail par série est désormais proposé par défaut. Dès qu'un exercice est ajouté à une séance (ou à un modèle), la fenêtre "Détail par série" s'ouvre automatiquement, pré-remplie avec une ligne par série et les valeurs saisies : il suffit d'ajuster les séries qui diffèrent puis de valider. Cliquer sur Annuler garde l'exercice avec des valeurs identiques à chaque série. L'exercice ajouté reste sélectionné, case "Détail par série" à jour.

### Modifié
- Pour une fourchette de séries ("3-4"), la fenêtre de détail par série propose désormais le maximum (4 lignes, comme la feuille de séance Excel) au lieu du minimum.

### Corrigé
- Un nombre mal saisi (ex. "abc" dans un montant) affichait "Une erreur est survenue" comme s'il s'agissait d'un plantage. Il affiche maintenant une fenêtre "Saisie invalide" qui indique le champ concerné et le format attendu (ex. « abc » n'est pas un nombre valide dans le champ « Montant » — ex : 12,5).

## [1.15.3] - 2026-10-01

### Corrigé
- Suivi > Questionnaires : à l'import d'un fichier Excel, un numéro de téléphone perdait son 0 initial (0612345678 → 612345678) quand sa colonne n'était pas indiquée dans l'assistant d'import (notamment pour les questionnaires bilan). Toutes les colonnes sont désormais lues telles qu'affichées dans Excel ; seule la colonne de date reste interprétée comme une date.

## [1.15.2] - 2026-10-01

### Corrigé
- Administratif : les échéances d'une commande **annulée** ne sont plus comptées comme paiements en attente/en retard dans le tableau de bord, ne passent plus "en retard", et n'apparaissent plus dans les filtres "En attente"/"En retard" de l'onglet Échéances (nouvelle colonne "Commande" pour voir le statut de la commande).
- Administratif : l'échéancier mensuel d'une commande commençant un 29, 30 ou 31 ne "dérive" plus (ex. début le 31/01 : 28/02 puis 31/03, 30/04… au lieu de rester bloqué au 28 tous les mois suivants).
- Suivi > Questionnaires : l'import d'un fichier **CSV** (proposé par la fenêtre de choix de fichier) plantait ; il fonctionne désormais (séparateur virgule ou point-virgule détecté automatiquement).
- Suivi > Questionnaires : les dates des réponses étaient stockées sous une forme illisible ("/Date(…)/") et affichées au format américain ; elles sont maintenant au format JJ/MM/AAAA HH:MM (y compris pour les réponses déjà importées).
- Suivi > Tracking quotidien / Roadmap : un fichier dont le client a supprimé une colonne faisait échouer toutes les lignes ; la colonne manquante est maintenant simplement laissée vide. La ligne d'exemple du modèle n'est plus importée comme une vraie donnée si le client l'a laissée.
- Suivi > Séances réalisées (et tracking/roadmap) : choisir un mauvais fichier affiche maintenant un message clair (colonnes manquantes) au lieu d'une erreur technique.
- Programmes : après "Monter", "Descendre" ou "Enregistrer" le jour d'une séance, la séance reste sélectionnée (avant, la sélection revenait toujours sur la première séance) ; une séance nouvellement ajoutée est directement sélectionnée.
- Les nombres saisis avec des espaces ("10 000") sont acceptés dans les formulaires et dans les imports Excel/FatSecret (ils étaient refusés, ou ignorés silencieusement à l'import).

## [1.15.1] - 2026-10-01

### Corrigé
- Administratif > Commandes : créer une commande directement (sans passer par « Transformer un devis en commande ») provoquait une erreur « variable DevisIdPourCommande non définie ». La variable est désormais initialisée au démarrage.

## [1.15.0] - 2026-09-29

### Ajouté
- Jour de la semaine sur les séances d'un programme (menu déroulant sous la liste des séances, écran Programmes) : quand au moins une séance en a un, un tableau récapitulatif "Semaine type" (Lundi → Dimanche, "Repos" les jours sans séance) apparaît en haut de l'export PDF du programme.

### Modifié
- L'export "Feuille de séance" (Excel à faire remplir par le client) génère désormais une ligne par série de chaque exercice, comme le tableau papier d'origine, avec des colonnes vides "Répétitions" et "Charge" à compléter série par série au lieu d'une seule ligne par exercice. La date de réalisation, la récup/tempo réalisés et la note ne sont à remplir qu'une fois par exercice (sur n'importe laquelle de ses lignes de série). Le réimport via Suivi > Séances réalisées assemble automatiquement les valeurs série par série (ex. "15 / 20 / 25").

## [1.14.1] - 2026-08-27

### Corrigé
- Case "Détail par série" (Programmes et Modèles de séance) : cocher la case ne faisait rien et provoquait l'erreur "la propriété Count est introuvable" lorsqu'aucune série détaillée n'existait encore pour l'exercice. Corrigé.

## [1.14.0] - 2026-08-27

### Modifié
- Refonte du détail par série (Programmes et Modèles de séance) pour le rendre plus intuitif : le bouton "Détail par série..." devient une case à cocher "Détail par série". La cocher ouvre directement une fenêtre pré-remplie avec autant de lignes que le nombre de séries déjà indiqué sur l'exercice, chacune reprenant ses valeurs actuelles (répétitions, charge, récup) — il suffit de modifier les séries qui doivent différer des autres, plus besoin de sélectionner puis valider chaque ligne une par une. Le tableau principal des exercices affiche désormais directement le détail par série ("15 / 20 / 25") dès qu'il est renseigné, sans avoir besoin d'exporter pour le voir.

## [1.13.1] - 2026-08-27

### Corrigé
- Écran Programmes : la ligne pour ajouter/modifier un exercice n'affichait que des cases vides, sans indication de ce que chacune représentait. Une étiquette (Exercice, Variante, Séries, Répétitions, Charge, Récup, Tempo, Notes) apparaît désormais au-dessus de chaque champ — pareil dans l'onglet Modèles de séance et dans la fenêtre "Détail par série...".

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
