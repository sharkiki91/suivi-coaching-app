Set-StrictMode -Version Latest

# --- Tracking quotidien ---

function Get-SuiviQuotidien {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM suivi_quotidien WHERE client_id = @ClientId ORDER BY date" -SqlParameters @{ ClientId = $ClientId }
}

function Set-SuiviQuotidienJour {
    <# Insere ou met a jour (upsert) la ligne de suivi d'un client pour une date donnee, sans toucher aux autres dates. #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $Date,
        [double] $Poids,
        [double] $SommeilHeures,
        [int] $QualiteSommeil,
        [string] $HeureCoucher,
        [string] $HeureLever,
        [int] $Energie,
        [int] $AdhesionNutrition,
        [int] $Digestion,
        [int] $NbPas,
        [double] $CardioMinutes,
        [int] $Motivation,
        [int] $TensionSystolique,
        [int] $TensionDiastolique,
        [string] $Bilan
    )
    <#
        SQLite embarque dans Windows PowerShell 5.1 (via PSSQLite) est une version ancienne (3.8.x)
        qui ne supporte pas la syntaxe UPSERT "ON CONFLICT ... DO UPDATE" (apparue en 3.24).
        On utilise donc INSERT OR REPLACE, compatible avec toutes les versions : la ligne en conflit
        (meme client_id + date, grace a la contrainte UNIQUE) est supprimee puis reinseree.
    #>
    $query = @"
INSERT OR REPLACE INTO suivi_quotidien (id, client_id, date, poids, sommeil_heures, qualite_sommeil, heure_coucher, heure_lever,
    energie, adhesion_nutrition, digestion, nb_pas, cardio_minutes, motivation, tension_systolique, tension_diastolique, bilan)
VALUES (
    (SELECT id FROM suivi_quotidien WHERE client_id = @ClientId AND date = @Date),
    @ClientId, @Date, @Poids, @SommeilHeures, @QualiteSommeil, @HeureCoucher, @HeureLever,
    @Energie, @AdhesionNutrition, @Digestion, @NbPas, @CardioMinutes, @Motivation, @TensionSystolique, @TensionDiastolique, @Bilan)
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        ClientId = $ClientId; Date = $Date; Poids = $Poids; SommeilHeures = $SommeilHeures; QualiteSommeil = $QualiteSommeil
        HeureCoucher = $HeureCoucher; HeureLever = $HeureLever; Energie = $Energie; AdhesionNutrition = $AdhesionNutrition
        Digestion = $Digestion; NbPas = $NbPas; CardioMinutes = $CardioMinutes; Motivation = $Motivation
        TensionSystolique = $TensionSystolique; TensionDiastolique = $TensionDiastolique; Bilan = $Bilan
    }
}

function Remove-SuiviQuotidienJour {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM suivi_quotidien WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

# --- Roadmap hebdomadaire ---

function Get-RoadmapSemaines {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM roadmap_semaines WHERE client_id = @ClientId ORDER BY semaine_numero" -SqlParameters @{ ClientId = $ClientId }
}

function New-RoadmapSemaine {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [int] $SemaineNumero,
        [string] $DateDebut,
        [string] $Phase,
        [string] $Nutrition,
        [double] $PoidsMoyen,
        [double] $DepenseCalorique,
        [double] $CardioMinutes,
        [int] $Pas,
        [string] $PrecisionTraining,
        [string] $Evenements,
        [string] $Notes
    )
    $query = @"
INSERT INTO roadmap_semaines (client_id, semaine_numero, date_debut, phase, nutrition, poids_moyen, depense_calorique, cardio_minutes, pas, precision_training, evenements, notes)
VALUES (@ClientId, @SemaineNumero, @DateDebut, @Phase, @Nutrition, @PoidsMoyen, @DepenseCalorique, @CardioMinutes, @Pas, @PrecisionTraining, @Evenements, @Notes);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        ClientId = $ClientId; SemaineNumero = $SemaineNumero; DateDebut = $DateDebut; Phase = $Phase; Nutrition = $Nutrition
        PoidsMoyen = $PoidsMoyen; DepenseCalorique = $DepenseCalorique; CardioMinutes = $CardioMinutes; Pas = $Pas
        PrecisionTraining = $PrecisionTraining; Evenements = $Evenements; Notes = $Notes
    }).id
}

function Update-RoadmapSemaine {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [int] $SemaineNumero,
        [string] $DateDebut,
        [string] $Phase,
        [string] $Nutrition,
        [double] $PoidsMoyen,
        [double] $DepenseCalorique,
        [double] $CardioMinutes,
        [int] $Pas,
        [string] $PrecisionTraining,
        [string] $Evenements,
        [string] $Notes
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query @"
UPDATE roadmap_semaines SET semaine_numero = @SemaineNumero, date_debut = @DateDebut, phase = @Phase, nutrition = @Nutrition,
    poids_moyen = @PoidsMoyen, depense_calorique = @DepenseCalorique, cardio_minutes = @CardioMinutes, pas = @Pas,
    precision_training = @PrecisionTraining, evenements = @Evenements, notes = @Notes
WHERE id = @Id
"@ -SqlParameters @{
        Id = $Id; SemaineNumero = $SemaineNumero; DateDebut = $DateDebut; Phase = $Phase; Nutrition = $Nutrition
        PoidsMoyen = $PoidsMoyen; DepenseCalorique = $DepenseCalorique; CardioMinutes = $CardioMinutes; Pas = $Pas
        PrecisionTraining = $PrecisionTraining; Evenements = $Evenements; Notes = $Notes
    }
}

function Remove-RoadmapSemaine {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM roadmap_semaines WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

# --- Journal alimentaire (import FatSecret) ---

function Get-JournalAlimentaire {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM journal_alimentaire WHERE client_id = @ClientId ORDER BY date" -SqlParameters @{ ClientId = $ClientId }
}

function Set-JournalAlimentaireJour {
    <# Insere ou met a jour (upsert) la ligne de journal alimentaire d'un client pour une date donnee. #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $Date,
        [double] $Kcal,
        [double] $Lipides,
        [double] $LipidesSaturees,
        [double] $Glucides,
        [double] $Fibres,
        [double] $Sucres,
        [double] $Proteines,
        [double] $SodiumMg,
        [double] $CholesterolMg,
        [double] $PotassiumMg
    )
    $query = @"
INSERT OR REPLACE INTO journal_alimentaire (id, client_id, date, kcal, lipides, lipides_satures, glucides, fibres, sucres, proteines, sodium_mg, cholesterol_mg, potassium_mg)
VALUES (
    (SELECT id FROM journal_alimentaire WHERE client_id = @ClientId AND date = @Date),
    @ClientId, @Date, @Kcal, @Lipides, @LipidesSaturees, @Glucides, @Fibres, @Sucres, @Proteines, @SodiumMg, @CholesterolMg, @PotassiumMg)
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        ClientId = $ClientId; Date = $Date; Kcal = $Kcal; Lipides = $Lipides; LipidesSaturees = $LipidesSaturees
        Glucides = $Glucides; Fibres = $Fibres; Sucres = $Sucres; Proteines = $Proteines
        SodiumMg = $SodiumMg; CholesterolMg = $CholesterolMg; PotassiumMg = $PotassiumMg
    }
}

function Remove-JournalAlimentaireJour {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM journal_alimentaire WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

Export-ModuleMember -Function Get-SuiviQuotidien, Set-SuiviQuotidienJour, Remove-SuiviQuotidienJour, `
    Get-RoadmapSemaines, New-RoadmapSemaine, Update-RoadmapSemaine, Remove-RoadmapSemaine, `
    Get-JournalAlimentaire, Set-JournalAlimentaireJour, Remove-JournalAlimentaireJour
