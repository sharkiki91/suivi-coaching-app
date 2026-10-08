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
    <# Supprime le programme et tout son contenu (seances, exercices de seance, series par exercice, seances realisees). #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    $query = @"
DELETE FROM exercices_realises WHERE seance_realisee_id IN (SELECT id FROM seances_realisees WHERE seance_id IN (SELECT id FROM seances WHERE programme_id = @Id));
DELETE FROM seances_realisees WHERE seance_id IN (SELECT id FROM seances WHERE programme_id = @Id);
DELETE FROM seance_exercice_series WHERE seance_exercice_id IN (SELECT id FROM seance_exercices WHERE seance_id IN (SELECT id FROM seances WHERE programme_id = @Id));
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
    $query = @"
SELECT *,
    nom || CASE WHEN jour_semaine IS NOT NULL AND jour_semaine <> '' THEN ' (' || jour_semaine || ')' ELSE '' END AS affichage
FROM seances WHERE programme_id = @ProgrammeId ORDER BY ordre, id
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ ProgrammeId = $ProgrammeId }
}

function Update-SeanceJour {
    <# Definit (ou retire, si Jour est vide/absent) le jour de la semaine associe a une seance. #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [string] $Jour
    )
    $jourNormalise = if ([string]::IsNullOrWhiteSpace($Jour)) { $null } else { $Jour.Trim() }
    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE seances SET jour_semaine = @Jour WHERE id = @Id" `
        -SqlParameters @{ Id = $Id; Jour = $jourNormalise }
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
DELETE FROM exercices_realises WHERE seance_realisee_id IN (SELECT id FROM seances_realisees WHERE seance_id = @Id);
DELETE FROM seances_realisees WHERE seance_id = @Id;
DELETE FROM seance_exercice_series WHERE seance_exercice_id IN (SELECT id FROM seance_exercices WHERE seance_id = @Id);
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
SELECT se.*, e.nom AS exercice_nom, e.muscle_cible, e.lien_video, e.image_path
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
        [string] $Series,
        [string] $Repetitions,
        [string] $Charge,
        [string] $RecuperationS,
        [string] $Tempo,
        [string] $Rir,
        [string] $Variante,
        [string] $Notes,
        [int] $Superset = 0
    )
    $ordreMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(ordre), -1) AS m FROM seance_exercices WHERE seance_id = @SeanceId" -SqlParameters @{ SeanceId = $SeanceId }).m
    $query = @"
INSERT INTO seance_exercices (seance_id, exercice_id, ordre, series, repetitions, charge, recuperation_s, tempo, rir, variante, notes, superset)
VALUES (@SeanceId, @ExerciceId, @Ordre, @Series, @Repetitions, @Charge, @RecuperationS, @Tempo, @Rir, @Variante, @Notes, @Superset);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        SeanceId = $SeanceId; ExerciceId = $ExerciceId; Ordre = ($ordreMax + 1)
        Series = $Series; Repetitions = $Repetitions; Charge = $Charge; RecuperationS = $RecuperationS; Tempo = $Tempo; Rir = $Rir; Variante = $Variante; Notes = $Notes; Superset = $Superset
    }).id
}

function Update-SeanceExercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [int] $ExerciceId = 0,   # > 0 : remplace l'exercice de la ligne (choix dans la liste)
        [string] $Series,
        [string] $Repetitions,
        [string] $Charge,
        [string] $RecuperationS,
        [string] $Tempo,
        [string] $Rir,
        [string] $Variante,
        [string] $Notes
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query @"
UPDATE seance_exercices SET exercice_id = CASE WHEN @ExerciceId > 0 THEN @ExerciceId ELSE exercice_id END, series = @Series, repetitions = @Repetitions, charge = @Charge, recuperation_s = @RecuperationS, tempo = @Tempo, rir = @Rir, variante = @Variante, notes = @Notes
WHERE id = @Id
"@ -SqlParameters @{ Id = $Id; ExerciceId = $ExerciceId; Series = $Series; Repetitions = $Repetitions; Charge = $Charge; RecuperationS = $RecuperationS; Tempo = $Tempo; Rir = $Rir; Variante = $Variante; Notes = $Notes }
}

function Remove-SeanceExercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    $query = @"
DELETE FROM seance_exercice_series WHERE seance_exercice_id = @Id;
DELETE FROM seance_exercices WHERE id = @Id;
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ Id = $Id }
}

# --- Detail par serie d'un exercice de seance (optionnel, pour les schemas type pyramide) ---

function Get-SeanceExerciceSeries {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceExerciceId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM seance_exercice_series WHERE seance_exercice_id = @Id ORDER BY numero_serie" -SqlParameters @{ Id = $SeanceExerciceId }
}

function New-SeanceExerciceSerie {
    <# Ajoute une serie a la fin (numero_serie = MAX + 1). #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceExerciceId,
        [string] $Repetitions,
        [string] $Charge,
        [string] $RecuperationS
    )
    $numeroMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(numero_serie), 0) AS m FROM seance_exercice_series WHERE seance_exercice_id = @Id" -SqlParameters @{ Id = $SeanceExerciceId }).m
    $query = @"
INSERT INTO seance_exercice_series (seance_exercice_id, numero_serie, repetitions, charge, recuperation_s)
VALUES (@SeanceExerciceId, @NumeroSerie, @Repetitions, @Charge, @RecuperationS);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        SeanceExerciceId = $SeanceExerciceId; NumeroSerie = ($numeroMax + 1); Repetitions = $Repetitions; Charge = $Charge; RecuperationS = $RecuperationS
    }).id
}

function Remove-SeanceExerciceSeriesTout {
    <# Supprime toutes les series d'un exercice (utilise par la fenetre "Detail par serie" qui remplace tout a chaque validation). #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $SeanceExerciceId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM seance_exercice_series WHERE seance_exercice_id = @Id" -SqlParameters @{ Id = $SeanceExerciceId }
}

# --- Partage entre exercices de seance (Programme) et exercices de modele (Modele) ---

function Get-TablesExercice {
    <# Noms de tables/colonnes selon le contexte (liste fermee : jamais de nom de table venant de l'exterieur dans le SQL). #>
    param([Parameter(Mandatory)] [ValidateSet('Programme', 'Modele')] [string] $Contexte)
    if ($Contexte -eq 'Programme') {
        return @{ Lignes = 'seance_exercices'; Parent = 'seance_id'; Series = 'seance_exercice_series'; CleSerie = 'seance_exercice_id' }
    }
    return @{ Lignes = 'seance_modele_exercices'; Parent = 'seance_modele_id'; Series = 'seance_modele_exercice_series'; CleSerie = 'seance_modele_exercice_id' }
}

function Move-LigneExercice {
    <#
        Monte (Direction -1) ou descend (Direction 1) un exercice dans sa seance (ou son modele). Les
        positions de toute la seance sont renumerotees 0, 1, 2... au passage (les anciens exercices
        pouvaient avoir le meme numero d'ordre).
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [ValidateSet('Programme', 'Modele')] [string] $Contexte,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [ValidateSet(-1, 1)] [int] $Direction
    )
    $t = Get-TablesExercice -Contexte $Contexte
    $parent = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT $($t.Parent) AS p FROM $($t.Lignes) WHERE id = @Id" -SqlParameters @{ Id = $Id }).p
    if ($null -eq $parent) { return }
    $ids = @(Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT id FROM $($t.Lignes) WHERE $($t.Parent) = @P ORDER BY ordre, id" -SqlParameters @{ P = $parent } | ForEach-Object { [int]$_.id })
    $index = [array]::IndexOf($ids, $Id)
    $cible = $index + $Direction
    if ($index -lt 0 -or $cible -lt 0 -or $cible -ge $ids.Count) { return }
    $ids[$index] = $ids[$cible]; $ids[$cible] = $Id
    for ($i = 0; $i -lt $ids.Count; $i++) {
        Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE $($t.Lignes) SET ordre = @Ordre WHERE id = @Id" -SqlParameters @{ Ordre = $i; Id = $ids[$i] }
    }
}

function Update-SeriesDetailDepuisLigne {
    <#
        Apres une modification de la ligne d'un exercice qui a un detail par serie : les champs modifies
        sur la ligne (-Champs, ex. @{ repetitions = '12' }) sont appliques a toutes les series du detail,
        les autres champs gardent leur detail (une pyramide de charge reste intacte si seules les reps
        changent). -NbSeries > 0 ajuste le nombre de series (ajout en recopiant la derniere, ou retrait
        des dernieres). Sans detail par serie, rien a faire : la ligne fait deja foi.
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [ValidateSet('Programme', 'Modele')] [string] $Contexte,
        [Parameter(Mandatory)] [int] $LigneId,
        [hashtable] $Champs = @{},
        [int] $NbSeries = 0
    )
    $t = Get-TablesExercice -Contexte $Contexte
    $series = @(Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM $($t.Series) WHERE $($t.CleSerie) = @Id ORDER BY numero_serie" -SqlParameters @{ Id = $LigneId })
    if ($series.Count -eq 0) { return }
    foreach ($champ in $Champs.Keys) {
        if ($champ -notin @('repetitions', 'charge', 'recuperation_s')) { continue }
        $valeur = if ([string]::IsNullOrWhiteSpace([string]$Champs[$champ])) { [DBNull]::Value } else { [string]$Champs[$champ] }
        Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE $($t.Series) SET $champ = @V WHERE $($t.CleSerie) = @Id" -SqlParameters @{ V = $valeur; Id = $LigneId }
    }
    if ($NbSeries -gt 0 -and $NbSeries -ne $series.Count) {
        if ($NbSeries -lt $series.Count) {
            Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM $($t.Series) WHERE $($t.CleSerie) = @Id AND numero_serie > @N" -SqlParameters @{ Id = $LigneId; N = $NbSeries }
        } else {
            $derniere = Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM $($t.Series) WHERE $($t.CleSerie) = @Id ORDER BY numero_serie DESC LIMIT 1" -SqlParameters @{ Id = $LigneId }
            for ($n = $series.Count + 1; $n -le $NbSeries; $n++) {
                Invoke-SqliteQuery -DataSource $DbPath -Query "INSERT INTO $($t.Series) ($($t.CleSerie), numero_serie, repetitions, charge, recuperation_s) VALUES (@Id, @N, @R, @C, @Rec)" `
                    -SqlParameters @{ Id = $LigneId; N = $n; R = $derniere.repetitions; C = $derniere.charge; Rec = $derniere.recuperation_s }
            }
        }
    }
}

function Move-SeanceExercice {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id, [Parameter(Mandatory)] [ValidateSet(-1, 1)] [int] $Direction)
    Move-LigneExercice -DbPath $DbPath -Contexte 'Programme' -Id $Id -Direction $Direction
}

function Switch-SupersetExercice {
    <# Lie (ou delie) l'exercice a l'exercice suivant de la seance / du modele : superset = enchaines sans recup entre les deux. #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [ValidateSet('Programme', 'Modele')] [string] $Contexte,
        [Parameter(Mandatory)] [int] $Id
    )
    $t = Get-TablesExercice -Contexte $Contexte
    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE $($t.Lignes) SET superset = CASE WHEN COALESCE(superset, 0) = 1 THEN 0 ELSE 1 END WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

function Get-NumerosExercices {
    <#
        Numerote les exercices d'une seance (dans l'ordre) comme sur la feuille du coach : 1, 2, 3...
        Un superset partage un seul numero avec une lettre par exercice (3A, 3B). Retourne, pour chaque
        exercice, un objet { Numero ; EstSuperset }.
    #>
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [array] $Exercices)
    $lettres = 'ABCDEFGHIJ'
    $numero = 0; $rang = 0; $lieAuPrecedent = $false
    for ($i = 0; $i -lt $Exercices.Count; $i++) {
        $lieAuSuivant = ($i -lt $Exercices.Count - 1) -and ([string]$Exercices[$i].superset -eq '1')
        if ($lieAuPrecedent) { $rang++ } else { $numero++; $rang = 0 }
        $estSuperset = $lieAuPrecedent -or $lieAuSuivant
        $texte = if ($estSuperset) { "$numero$($lettres[[math]::Min($rang, 9)])" } else { [string]$numero }
        [pscustomobject]@{ Numero = $texte; EstSuperset = $estSuperset }
        $lieAuPrecedent = $lieAuSuivant
    }
}

Export-ModuleMember -Function Switch-SupersetExercice, Get-NumerosExercices, Get-Programmes, New-Programme, Remove-Programme, `
    Get-Seances, New-Seance, Remove-Seance, Move-Seance, Update-SeanceJour, `
    Get-SeanceExercices, New-SeanceExercice, Update-SeanceExercice, Remove-SeanceExercice, `
    Get-SeanceExerciceSeries, New-SeanceExerciceSerie, Remove-SeanceExerciceSeriesTout, `
    Move-LigneExercice, Update-SeriesDetailDepuisLigne, Move-SeanceExercice
