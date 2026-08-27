Set-StrictMode -Version Latest

function Initialize-Database {
    param(
        [Parameter(Mandatory)] [string] $DbPath
    )

    $dbFolder = Split-Path -Path $DbPath -Parent
    if (-not (Test-Path $dbFolder)) {
        New-Item -ItemType Directory -Path $dbFolder -Force | Out-Null
    }

    $schema = @"
CREATE TABLE IF NOT EXISTS clients (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom TEXT NOT NULL,
    prenom TEXT NOT NULL,
    email TEXT,
    telephone TEXT,
    date_debut_coaching TEXT,
    objectifs TEXT,
    notes TEXT,
    statut TEXT NOT NULL DEFAULT 'actif' CHECK (statut IN ('actif','archive')),
    date_creation TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS devis (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    prestations TEXT NOT NULL,
    montant REAL NOT NULL,
    duree TEXT,
    date_creation TEXT NOT NULL DEFAULT (datetime('now')),
    statut TEXT NOT NULL DEFAULT 'en_attente' CHECK (statut IN ('en_attente','accepte','refuse'))
);

CREATE TABLE IF NOT EXISTS commandes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    devis_id INTEGER REFERENCES devis(id),
    type_facturation TEXT NOT NULL CHECK (type_facturation IN ('mensuel','hebdomadaire','one_shot')),
    date_debut TEXT NOT NULL,
    date_fin TEXT,
    montant REAL NOT NULL,
    statut TEXT NOT NULL DEFAULT 'active' CHECK (statut IN ('active','terminee','annulee')),
    date_creation TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS echeances (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    commande_id INTEGER NOT NULL REFERENCES commandes(id) ON DELETE CASCADE,
    date_echeance TEXT NOT NULL,
    montant REAL NOT NULL,
    statut TEXT NOT NULL DEFAULT 'en_attente' CHECK (statut IN ('payee','en_attente','en_retard')),
    date_paiement TEXT
);

CREATE INDEX IF NOT EXISTS idx_echeances_statut_date ON echeances(statut, date_echeance);
CREATE INDEX IF NOT EXISTS idx_devis_client ON devis(client_id);
CREATE INDEX IF NOT EXISTS idx_commandes_client ON commandes(client_id);

CREATE TABLE IF NOT EXISTS questionnaires_reponses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER REFERENCES clients(id) ON DELETE CASCADE,
    type TEXT NOT NULL CHECK (type IN ('pre_coaching','bilan')),
    date_reponse TEXT NOT NULL,
    donnees_json TEXT NOT NULL,
    fichier_source TEXT,
    date_import TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS import_mappings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    type TEXT NOT NULL,
    mapping_json TEXT NOT NULL,
    date_maj TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS exercices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom TEXT NOT NULL,
    muscle_cible TEXT,
    variante TEXT,
    lien_video TEXT,
    image_path TEXT
);

CREATE TABLE IF NOT EXISTS aliments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom TEXT NOT NULL,
    quantite_reference REAL NOT NULL DEFAULT 100,
    unite TEXT NOT NULL DEFAULT 'g',
    kcal REAL,
    proteines REAL,
    glucides REAL,
    lipides REAL,
    fibres REAL
);

CREATE TABLE IF NOT EXISTS aliments_micronutriments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aliment_id INTEGER NOT NULL REFERENCES aliments(id) ON DELETE CASCADE,
    nom TEXT NOT NULL,
    valeur REAL,
    unite TEXT
);

CREATE TABLE IF NOT EXISTS complements (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom TEXT NOT NULL,
    dose TEXT,
    timing TEXT,
    note TEXT,
    lien TEXT
);

CREATE TABLE IF NOT EXISTS complements_recommandations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    complement_id INTEGER NOT NULL REFERENCES complements(id),
    dose TEXT,
    timing TEXT,
    note TEXT,
    date_debut TEXT
);

CREATE TABLE IF NOT EXISTS programmes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    nom TEXT NOT NULL,
    date_debut TEXT,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS seances (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    programme_id INTEGER NOT NULL REFERENCES programmes(id) ON DELETE CASCADE,
    nom TEXT NOT NULL,
    ordre INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS seance_exercices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    seance_id INTEGER NOT NULL REFERENCES seances(id) ON DELETE CASCADE,
    exercice_id INTEGER NOT NULL REFERENCES exercices(id),
    ordre INTEGER NOT NULL DEFAULT 0,
    series INTEGER,
    repetitions TEXT,
    recuperation_s INTEGER,
    tempo TEXT,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS seances_realisees (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    seance_id INTEGER NOT NULL REFERENCES seances(id),
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    date_realisation TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS series_realisees (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    seance_realisee_id INTEGER NOT NULL REFERENCES seances_realisees(id) ON DELETE CASCADE,
    seance_exercice_id INTEGER NOT NULL REFERENCES seance_exercices(id),
    numero_serie INTEGER NOT NULL,
    reps_realisees INTEGER,
    charge REAL,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS seance_modeles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom TEXT NOT NULL,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS seance_modele_exercices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    seance_modele_id INTEGER NOT NULL REFERENCES seance_modeles(id) ON DELETE CASCADE,
    exercice_id INTEGER NOT NULL REFERENCES exercices(id),
    ordre INTEGER NOT NULL DEFAULT 0,
    series INTEGER,
    repetitions TEXT,
    recuperation_s INTEGER,
    tempo TEXT,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS plans_nutrition (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    nom TEXT NOT NULL,
    date_debut TEXT,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS types_jour (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plan_nutrition_id INTEGER NOT NULL REFERENCES plans_nutrition(id) ON DELETE CASCADE,
    nom TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS repas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    type_jour_id INTEGER NOT NULL REFERENCES types_jour(id) ON DELETE CASCADE,
    nom TEXT NOT NULL,
    ordre INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS repas_aliments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    repas_id INTEGER NOT NULL REFERENCES repas(id) ON DELETE CASCADE,
    aliment_id INTEGER NOT NULL REFERENCES aliments(id),
    quantite REAL NOT NULL,
    ordre INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS roadmap_semaines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    semaine_numero INTEGER NOT NULL,
    date_debut TEXT,
    phase TEXT,
    nutrition TEXT,
    poids_moyen REAL,
    depense_calorique REAL,
    cardio_minutes REAL,
    pas INTEGER,
    precision_training TEXT,
    evenements TEXT,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS suivi_quotidien (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    date TEXT NOT NULL,
    poids REAL,
    sommeil_heures REAL,
    qualite_sommeil INTEGER,
    heure_coucher TEXT,
    heure_lever TEXT,
    energie INTEGER,
    adhesion_nutrition INTEGER,
    jour_non_tracke INTEGER NOT NULL DEFAULT 0,
    digestion INTEGER,
    seance_realisee_id INTEGER REFERENCES seances_realisees(id),
    nb_pas INTEGER,
    cardio_minutes REAL,
    motivation INTEGER,
    tension_systolique INTEGER,
    tension_diastolique INTEGER,
    bilan TEXT,
    UNIQUE(client_id, date)
);

CREATE INDEX IF NOT EXISTS idx_suivi_client_date ON suivi_quotidien(client_id, date);

CREATE TABLE IF NOT EXISTS journal_alimentaire (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    date TEXT NOT NULL,
    kcal REAL,
    lipides REAL,
    lipides_satures REAL,
    glucides REAL,
    fibres REAL,
    sucres REAL,
    proteines REAL,
    sodium_mg REAL,
    cholesterol_mg REAL,
    potassium_mg REAL,
    UNIQUE(client_id, date)
);

CREATE TABLE IF NOT EXISTS parametres (
    cle TEXT PRIMARY KEY,
    valeur TEXT
);
"@

    Invoke-SqliteQuery -DataSource $DbPath -Query $schema
}

function Backup-Database {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [string] $BackupFolder
    )

    if (-not (Test-Path $BackupFolder)) {
        New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $destination = Join-Path $BackupFolder "suivi_coaching_$timestamp.db"
    Copy-Item -Path $DbPath -Destination $destination -Force
    return $destination
}

Export-ModuleMember -Function Initialize-Database, Backup-Database
