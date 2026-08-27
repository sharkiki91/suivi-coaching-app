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
        [int] $Series,
        [string] $Repetitions,
        [int] $RecuperationS,
        [string] $Tempo,
        [string] $Notes
    )
    $ordreMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(ordre), -1) AS m FROM seance_modele_exercices WHERE seance_modele_id = @SeanceModeleId" -SqlParameters @{ SeanceModeleId = $SeanceModeleId }).m
    $query = @"
INSERT INTO seance_modele_exercices (seance_modele_id, exercice_id, ordre, series, repetitions, recuperation_s, tempo, notes)
VALUES (@SeanceModeleId, @ExerciceId, @Ordre, @Series, @Repetitions, @RecuperationS, @Tempo, @Notes);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        SeanceModeleId = $SeanceModeleId; ExerciceId = $ExerciceId; Ordre = ($ordreMax + 1)
        Series = $Series; Repetitions = $Repetitions; RecuperationS = $RecuperationS; Tempo = $Tempo; Notes = $Notes
    }).id
}

function Remove-SeanceModeleExercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM seance_modele_exercices WHERE id = @Id" -SqlParameters @{ Id = $Id }
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
        New-SeanceExercice -DbPath $DbPath -SeanceId $seanceId -ExerciceId ([int]$ex.exercice_id) `
            -Series $ex.series -Repetitions $ex.repetitions -RecuperationS $ex.recuperation_s -Tempo $ex.tempo -Notes $ex.notes | Out-Null
    }
    return $seanceId
}

Export-ModuleMember -Function Get-SeanceModeles, New-SeanceModele, Update-SeanceModele, Remove-SeanceModele, `
    Get-SeanceModeleExercices, New-SeanceModeleExercice, Remove-SeanceModeleExercice, `
    New-SeanceDepuisModele
