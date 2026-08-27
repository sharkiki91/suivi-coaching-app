Set-StrictMode -Version Latest

# --- Modeles de seance ---

function Get-SeanceModeles {
    param(
        [Parameter(Mandatory)] [string] $DbPath
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM seance_modeles ORDER BY nom"
}

function New-SeanceModele {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [string] $Nom,
        [string] $Notes
    )
    $query = @"
INSERT INTO seance_modeles (nom, notes) VALUES (@Nom, @Notes);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ Nom = $Nom; Notes = $Notes }).id
}

function Update-SeanceModele {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [string] $Nom,
        [string] $Notes
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE seance_modeles SET nom = @Nom, notes = @Notes WHERE id = @Id" `
        -SqlParameters @{ Id = $Id; Nom = $Nom; Notes = $Notes }
}

function Remove-SeanceModele {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    $query = @"
DELETE FROM seance_modele_exercice_series WHERE seance_modele_exercice_id IN (SELECT id FROM seance_modele_exercices WHERE seance_modele_id = @Id);
DELETE FROM seance_modele_exercices WHERE seance_modele_id = @Id;
DELETE FROM seance_modeles WHERE id = @Id;
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ Id = $Id }
}

# --- Exercices d'un modele ---

function Get-SeanceModeleExercices {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceModeleId
    )
    $query = @"
SELECT sme.*, e.nom AS exercice_nom, e.muscle_cible, e.lien_video
FROM seance_modele_exercices sme
JOIN exercices e ON e.id = sme.exercice_id
WHERE sme.seance_modele_id = @SeanceModeleId
ORDER BY sme.ordre, sme.id
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ SeanceModeleId = $SeanceModeleId }
}

function New-SeanceModeleExercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceModeleId,
        [Parameter(Mandatory)] [int] $ExerciceId,
        [string] $Series,
        [string] $Repetitions,
        [string] $Charge,
        [string] $RecuperationS,
        [string] $Tempo,
        [string] $Variante,
        [string] $Notes
    )
    $ordreMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(ordre), -1) AS m FROM seance_modele_exercices WHERE seance_modele_id = @SeanceModeleId" -SqlParameters @{ SeanceModeleId = $SeanceModeleId }).m
    $query = @"
INSERT INTO seance_modele_exercices (seance_modele_id, exercice_id, ordre, series, repetitions, charge, recuperation_s, tempo, variante, notes)
VALUES (@SeanceModeleId, @ExerciceId, @Ordre, @Series, @Repetitions, @Charge, @RecuperationS, @Tempo, @Variante, @Notes);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        SeanceModeleId = $SeanceModeleId; ExerciceId = $ExerciceId; Ordre = ($ordreMax + 1)
        Series = $Series; Repetitions = $Repetitions; Charge = $Charge; RecuperationS = $RecuperationS; Tempo = $Tempo; Variante = $Variante; Notes = $Notes
    }).id
}

function Update-SeanceModeleExercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [string] $Series,
        [string] $Repetitions,
        [string] $Charge,
        [string] $RecuperationS,
        [string] $Tempo,
        [string] $Variante,
        [string] $Notes
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query @"
UPDATE seance_modele_exercices SET series = @Series, repetitions = @Repetitions, charge = @Charge, recuperation_s = @RecuperationS, tempo = @Tempo, variante = @Variante, notes = @Notes
WHERE id = @Id
"@ -SqlParameters @{ Id = $Id; Series = $Series; Repetitions = $Repetitions; Charge = $Charge; RecuperationS = $RecuperationS; Tempo = $Tempo; Variante = $Variante; Notes = $Notes }
}

function Remove-SeanceModeleExercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    $query = @"
DELETE FROM seance_modele_exercice_series WHERE seance_modele_exercice_id = @Id;
DELETE FROM seance_modele_exercices WHERE id = @Id;
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ Id = $Id }
}

# --- Detail par serie d'un exercice de modele (optionnel, pour les schemas type pyramide) ---

function Get-SeanceModeleExerciceSeries {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceModeleExerciceId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM seance_modele_exercice_series WHERE seance_modele_exercice_id = @Id ORDER BY numero_serie" -SqlParameters @{ Id = $SeanceModeleExerciceId }
}

function New-SeanceModeleExerciceSerie {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceModeleExerciceId,
        [string] $Repetitions,
        [string] $Charge,
        [string] $RecuperationS
    )
    $numeroMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(numero_serie), 0) AS m FROM seance_modele_exercice_series WHERE seance_modele_exercice_id = @Id" -SqlParameters @{ Id = $SeanceModeleExerciceId }).m
    $query = @"
INSERT INTO seance_modele_exercice_series (seance_modele_exercice_id, numero_serie, repetitions, charge, recuperation_s)
VALUES (@SeanceModeleExerciceId, @NumeroSerie, @Repetitions, @Charge, @RecuperationS);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        SeanceModeleExerciceId = $SeanceModeleExerciceId; NumeroSerie = ($numeroMax + 1); Repetitions = $Repetitions; Charge = $Charge; RecuperationS = $RecuperationS
    }).id
}

function Remove-SeanceModeleExerciceSeriesTout {
    <# Supprime toutes les series d'un exercice de modele (utilise par la fenetre "Detail par serie" qui remplace tout a chaque validation). #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceModeleExerciceId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM seance_modele_exercice_series WHERE seance_modele_exercice_id = @Id" -SqlParameters @{ Id = $SeanceModeleExerciceId }
}

# --- Utilisation d'un modele dans un programme ---

function New-SeanceDepuisModele {
    <# Cree une nouvelle seance dans le programme donne, en copiant les exercices du modele. Retourne l'id de la seance creee. #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ProgrammeId,
        [Parameter(Mandatory)] [int] $SeanceModeleId,
        [Parameter(Mandatory)] [string] $Nom
    )
    $seanceId = New-Seance -DbPath $DbPath -ProgrammeId $ProgrammeId -Nom $Nom
    $exercices = @(Get-SeanceModeleExercices -DbPath $DbPath -SeanceModeleId $SeanceModeleId)
    foreach ($ex in $exercices) {
        $nouvelExerciceId = New-SeanceExercice -DbPath $DbPath -SeanceId $seanceId -ExerciceId ([int]$ex.exercice_id) `
            -Series $ex.series -Repetitions $ex.repetitions -Charge $ex.charge -RecuperationS $ex.recuperation_s -Tempo $ex.tempo -Variante $ex.variante -Notes $ex.notes
        $seriesModele = @(Get-SeanceModeleExerciceSeries -DbPath $DbPath -SeanceModeleExerciceId ([int]$ex.id))
        foreach ($s in $seriesModele) {
            New-SeanceExerciceSerie -DbPath $DbPath -SeanceExerciceId $nouvelExerciceId -Repetitions $s.repetitions -Charge $s.charge -RecuperationS $s.recuperation_s | Out-Null
        }
    }
    return $seanceId
}

Export-ModuleMember -Function Get-SeanceModeles, New-SeanceModele, Update-SeanceModele, Remove-SeanceModele, `
    Get-SeanceModeleExercices, New-SeanceModeleExercice, Update-SeanceModeleExercice, Remove-SeanceModeleExercice, `
    Get-SeanceModeleExerciceSeries, New-SeanceModeleExerciceSerie, Remove-SeanceModeleExerciceSeriesTout, `
    New-SeanceDepuisModele
