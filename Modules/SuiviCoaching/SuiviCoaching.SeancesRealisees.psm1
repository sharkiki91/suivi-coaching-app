Set-StrictMode -Version Latest

# --- Seances realisees ---

function Get-SeancesRealisees {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId
    )
    $query = @"
SELECT sr.*, s.nom AS seance_nom, p.nom AS programme_nom,
    (SELECT COUNT(*) FROM exercices_realises er WHERE er.seance_realisee_id = sr.id) AS nb_exercices
FROM seances_realisees sr
JOIN seances s ON s.id = sr.seance_id
JOIN programmes p ON p.id = s.programme_id
WHERE sr.client_id = @ClientId
ORDER BY sr.date_realisation DESC
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ ClientId = $ClientId }
}

function Get-OuCreerSeanceRealisee {
    <# Retrouve la seance realisee (seance_id + client_id + date) existante, ou la cree si absente. Retourne son id. #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceId,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $DateRealisation
    )
    $existante = Invoke-SqliteQuery -DataSource $DbPath -Query @"
SELECT id FROM seances_realisees WHERE seance_id = @SeanceId AND client_id = @ClientId AND date_realisation = @Date
"@ -SqlParameters @{ SeanceId = $SeanceId; ClientId = $ClientId; Date = $DateRealisation }
    if ($existante) { return [int]$existante.id }

    $query = @"
INSERT INTO seances_realisees (seance_id, client_id, date_realisation) VALUES (@SeanceId, @ClientId, @Date);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ SeanceId = $SeanceId; ClientId = $ClientId; Date = $DateRealisation }).id
}

function Remove-SeanceRealisee {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    $query = @"
DELETE FROM exercices_realises WHERE seance_realisee_id = @Id;
DELETE FROM seances_realisees WHERE id = @Id;
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ Id = $Id }
}

# --- Exercices realises ---

function Get-ExercicesRealises {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceRealiseeId
    )
    $query = @"
SELECT er.*, e.nom AS exercice_nom, e.muscle_cible
FROM exercices_realises er
JOIN seance_exercices se ON se.id = er.seance_exercice_id
JOIN exercices e ON e.id = se.exercice_id
WHERE er.seance_realisee_id = @SeanceRealiseeId
ORDER BY se.ordre, se.id
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ SeanceRealiseeId = $SeanceRealiseeId }
}

function Set-ExerciceRealise {
    <# Insere ou met a jour (upsert) la ligne "realise" d'un exercice pour une seance realisee donnee. #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceRealiseeId,
        [Parameter(Mandatory)] [int] $SeanceExerciceId,
        [string] $Series,
        [string] $Repetitions,
        [string] $Charge,
        [string] $RecuperationS,
        [string] $Tempo,
        [string] $Notes
    )
    <#
        SQLite embarque dans Windows PowerShell 5.1 (via PSSQLite) est une version ancienne (3.8.x)
        qui ne supporte pas la syntaxe UPSERT "ON CONFLICT ... DO UPDATE" (apparue en 3.24).
        On utilise donc INSERT OR REPLACE, compatible avec toutes les versions.
    #>
    $query = @"
INSERT OR REPLACE INTO exercices_realises (id, seance_realisee_id, seance_exercice_id, series, repetitions, charge, recuperation_s, tempo, notes)
VALUES (
    (SELECT id FROM exercices_realises WHERE seance_realisee_id = @SeanceRealiseeId AND seance_exercice_id = @SeanceExerciceId),
    @SeanceRealiseeId, @SeanceExerciceId, @Series, @Repetitions, @Charge, @RecuperationS, @Tempo, @Notes)
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        SeanceRealiseeId = $SeanceRealiseeId; SeanceExerciceId = $SeanceExerciceId
        Series = $Series; Repetitions = $Repetitions; Charge = $Charge; RecuperationS = $RecuperationS; Tempo = $Tempo; Notes = $Notes
    }
}

Export-ModuleMember -Function Get-SeancesRealisees, Get-OuCreerSeanceRealisee, Remove-SeanceRealisee, `
    Get-ExercicesRealises, Set-ExerciceRealise
