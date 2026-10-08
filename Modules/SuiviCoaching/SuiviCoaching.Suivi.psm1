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
    <#
        Insere ou met a jour (upsert) la ligne de suivi d'un client pour une date donnee, sans toucher aux autres dates.
        Une valeur non renseignee ($null ou texte vide) est enregistree vide (NULL), jamais 0 : un poids non pese
        ne doit pas apparaitre comme "0 kg".
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $Date,
        [Nullable[double]] $Poids,
        [Nullable[double]] $SommeilHeures,
        [Nullable[int]] $QualiteSommeil,
        $HeureCoucher,
        $HeureLever,
        [Nullable[int]] $Energie,
        [Nullable[int]] $AdhesionNutrition,
        [bool] $JourNonTracke,
        [Nullable[int]] $Digestion,
        $Seance,
        [Nullable[int]] $NbPas,
        [Nullable[double]] $CardioMinutes,
        [Nullable[int]] $Motivation,
        [Nullable[int]] $FcRepos,
        [Nullable[int]] $TensionSystolique,
        [Nullable[int]] $TensionDiastolique,
        $Bilan
    )
    <#
        SQLite embarque dans Windows PowerShell 5.1 (via PSSQLite) est une version ancienne (3.8.x)
        qui ne supporte pas la syntaxe UPSERT "ON CONFLICT ... DO UPDATE" (apparue en 3.24).
        On utilise donc INSERT OR REPLACE, compatible avec toutes les versions : la ligne en conflit
        (meme client_id + date, grace a la contrainte UNIQUE) est supprimee puis reinseree.
    #>
    $texte = { param($v) if ($null -eq $v -or [string]::IsNullOrWhiteSpace([string]$v)) { [DBNull]::Value } else { ([string]$v).Trim() } }
    $nombre = { param($v) if ($null -eq $v) { [DBNull]::Value } else { $v } }
    $query = @"
INSERT OR REPLACE INTO suivi_quotidien (id, client_id, date, poids, sommeil_heures, qualite_sommeil, heure_coucher, heure_lever,
    energie, adhesion_nutrition, jour_non_tracke, digestion, seance, nb_pas, cardio_minutes, motivation, fc_repos,
    tension_systolique, tension_diastolique, bilan)
VALUES (
    (SELECT id FROM suivi_quotidien WHERE client_id = @ClientId AND date = @Date),
    @ClientId, @Date, @Poids, @SommeilHeures, @QualiteSommeil, @HeureCoucher, @HeureLever,
    @Energie, @AdhesionNutrition, @JourNonTracke, @Digestion, @Seance, @NbPas, @CardioMinutes, @Motivation, @FcRepos,
    @TensionSystolique, @TensionDiastolique, @Bilan)
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        ClientId = $ClientId; Date = $Date; Poids = (& $nombre $Poids); SommeilHeures = (& $nombre $SommeilHeures)
        QualiteSommeil = (& $nombre $QualiteSommeil); HeureCoucher = (& $texte $HeureCoucher); HeureLever = (& $texte $HeureLever)
        Energie = (& $nombre $Energie); AdhesionNutrition = (& $nombre $AdhesionNutrition); JourNonTracke = [int]$JourNonTracke
        Digestion = (& $nombre $Digestion); Seance = (& $texte $Seance); NbPas = (& $nombre $NbPas)
        CardioMinutes = (& $nombre $CardioMinutes); Motivation = (& $nombre $Motivation); FcRepos = (& $nombre $FcRepos)
        TensionSystolique = (& $nombre $TensionSystolique); TensionDiastolique = (& $nombre $TensionDiastolique); Bilan = (& $texte $Bilan)
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
        [Nullable[double]] $PoidsMoyen,
        [Nullable[double]] $DepenseCalorique,
        [Nullable[double]] $CardioMinutes,
        [Nullable[int]] $Pas,
        [string] $PrecisionDepense,
        [string] $PrecisionTraining,
        [string] $Evenements,
        [string] $Notes
    )
    $query = @"
INSERT INTO roadmap_semaines (client_id, semaine_numero, date_debut, phase, nutrition, poids_moyen, depense_calorique, cardio_minutes, pas, precision_depense, precision_training, evenements, notes)
VALUES (@ClientId, @SemaineNumero, @DateDebut, @Phase, @Nutrition, @PoidsMoyen, @DepenseCalorique, @CardioMinutes, @Pas, @PrecisionDepense, @PrecisionTraining, @Evenements, @Notes);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        ClientId = $ClientId; SemaineNumero = $SemaineNumero; DateDebut = $DateDebut; Phase = $Phase; Nutrition = $Nutrition
        PoidsMoyen = $PoidsMoyen; DepenseCalorique = $DepenseCalorique; CardioMinutes = $CardioMinutes; Pas = $Pas
        PrecisionDepense = $PrecisionDepense; PrecisionTraining = $PrecisionTraining; Evenements = $Evenements; Notes = $Notes
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
        [Nullable[double]] $PoidsMoyen,
        [Nullable[double]] $DepenseCalorique,
        [Nullable[double]] $CardioMinutes,
        [Nullable[int]] $Pas,
        [string] $PrecisionDepense,
        [string] $PrecisionTraining,
        [string] $Evenements,
        [string] $Notes
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query @"
UPDATE roadmap_semaines SET semaine_numero = @SemaineNumero, date_debut = @DateDebut, phase = @Phase, nutrition = @Nutrition,
    poids_moyen = @PoidsMoyen, depense_calorique = @DepenseCalorique, cardio_minutes = @CardioMinutes, pas = @Pas, precision_depense = @PrecisionDepense,
    precision_training = @PrecisionTraining, evenements = @Evenements, notes = @Notes
WHERE id = @Id
"@ -SqlParameters @{
        Id = $Id; SemaineNumero = $SemaineNumero; DateDebut = $DateDebut; Phase = $Phase; Nutrition = $Nutrition
        PoidsMoyen = $PoidsMoyen; DepenseCalorique = $DepenseCalorique; CardioMinutes = $CardioMinutes; Pas = $Pas
        PrecisionDepense = $PrecisionDepense; PrecisionTraining = $PrecisionTraining; Evenements = $Evenements; Notes = $Notes
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
