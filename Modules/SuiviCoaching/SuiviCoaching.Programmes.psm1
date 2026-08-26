Set-StrictMode -Version Latest

# --- Programmes ---

function Get-Programmes {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM programmes WHERE client_id = @ClientId ORDER BY date_debut DESC" -SqlParameters @{ ClientId = $ClientId }
}

function New-Programme {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $Nom,
        [string] $DateDebut,
        [string] $Notes
    )
    $query = @"
INSERT INTO programmes (client_id, nom, date_debut, notes) VALUES (@ClientId, @Nom, @DateDebut, @Notes);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ ClientId = $ClientId; Nom = $Nom; DateDebut = $DateDebut; Notes = $Notes }).id
}

function Remove-Programme {
    <# Supprime le programme et tout son contenu (seances, exercices de seance, seances realisees, series). #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    $query = @"
DELETE FROM series_realisees WHERE seance_realisee_id IN (SELECT id FROM seances_realisees WHERE seance_id IN (SELECT id FROM seances WHERE programme_id = @Id));
DELETE FROM seances_realisees WHERE seance_id IN (SELECT id FROM seances WHERE programme_id = @Id);
DELETE FROM seance_exercices WHERE seance_id IN (SELECT id FROM seances WHERE programme_id = @Id);
DELETE FROM seances WHERE programme_id = @Id;
DELETE FROM programmes WHERE id = @Id;
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ Id = $Id }
}

# --- Seances ---

function Get-Seances {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ProgrammeId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM seances WHERE programme_id = @ProgrammeId ORDER BY ordre, id" -SqlParameters @{ ProgrammeId = $ProgrammeId }
}

function New-Seance {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ProgrammeId,
        [Parameter(Mandatory)] [string] $Nom
    )
    $ordreMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(ordre), -1) AS m FROM seances WHERE programme_id = @ProgrammeId" -SqlParameters @{ ProgrammeId = $ProgrammeId }).m
    $query = @"
INSERT INTO seances (programme_id, nom, ordre) VALUES (@ProgrammeId, @Nom, @Ordre);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ ProgrammeId = $ProgrammeId; Nom = $Nom; Ordre = ($ordreMax + 1) }).id
}

function Remove-Seance {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    $query = @"
DELETE FROM series_realisees WHERE seance_realisee_id IN (SELECT id FROM seances_realisees WHERE seance_id = @Id);
DELETE FROM seances_realisees WHERE seance_id = @Id;
DELETE FROM seance_exercices WHERE seance_id = @Id;
DELETE FROM seances WHERE id = @Id;
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ Id = $Id }
}

function Move-Seance {
    <# Echange l'ordre de la seance avec celle juste avant (Direction -1) ou juste apres (Direction 1). #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [int] $ProgrammeId,
        [Parameter(Mandatory)] [ValidateSet(-1, 1)] [int] $Direction
    )
    $seances = @(Get-Seances -DbPath $DbPath -ProgrammeId $ProgrammeId)
    $index = 0
    for ($i = 0; $i -lt $seances.Count; $i++) { if ([int]$seances[$i].id -eq $Id) { $index = $i } }
    $swapIndex = $index + $Direction
    if ($swapIndex -lt 0 -or $swapIndex -ge $seances.Count) { return }
    $a = $seances[$index]; $b = $seances[$swapIndex]
    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE seances SET ordre = @Ordre WHERE id = @Id" -SqlParameters @{ Ordre = $b.ordre; Id = $a.id }
    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE seances SET ordre = @Ordre WHERE id = @Id" -SqlParameters @{ Ordre = $a.ordre; Id = $b.id }
}

# --- Exercices d'une seance ---

function Get-SeanceExercices {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceId
    )
    $query = @"
SELECT se.*, e.nom AS exercice_nom, e.muscle_cible, e.lien_video
FROM seance_exercices se
JOIN exercices e ON e.id = se.exercice_id
WHERE se.seance_id = @SeanceId
ORDER BY se.ordre, se.id
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ SeanceId = $SeanceId }
}

function New-SeanceExercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceId,
        [Parameter(Mandatory)] [int] $ExerciceId,
        [int] $Series,
        [string] $Repetitions,
        [int] $RecuperationS,
        [string] $Tempo,
        [string] $Notes
    )
    $ordreMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(ordre), -1) AS m FROM seance_exercices WHERE seance_id = @SeanceId" -SqlParameters @{ SeanceId = $SeanceId }).m
    $query = @"
INSERT INTO seance_exercices (seance_id, exercice_id, ordre, series, repetitions, recuperation_s, tempo, notes)
VALUES (@SeanceId, @ExerciceId, @Ordre, @Series, @Repetitions, @RecuperationS, @Tempo, @Notes);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        SeanceId = $SeanceId; ExerciceId = $ExerciceId; Ordre = ($ordreMax + 1)
        Series = $Series; Repetitions = $Repetitions; RecuperationS = $RecuperationS; Tempo = $Tempo; Notes = $Notes
    }).id
}

function Update-SeanceExercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [int] $Series,
        [string] $Repetitions,
        [int] $RecuperationS,
        [string] $Tempo,
        [string] $Notes
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query @"
UPDATE seance_exercices SET series = @Series, repetitions = @Repetitions, recuperation_s = @RecuperationS, tempo = @Tempo, notes = @Notes
WHERE id = @Id
"@ -SqlParameters @{ Id = $Id; Series = $Series; Repetitions = $Repetitions; RecuperationS = $RecuperationS; Tempo = $Tempo; Notes = $Notes }
}

function Remove-SeanceExercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM seance_exercices WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

Export-ModuleMember -Function Get-Programmes, New-Programme, Remove-Programme, `
    Get-Seances, New-Seance, Remove-Seance, Move-Seance, `
    Get-SeanceExercices, New-SeanceExercice, Update-SeanceExercice, Remove-SeanceExercice
