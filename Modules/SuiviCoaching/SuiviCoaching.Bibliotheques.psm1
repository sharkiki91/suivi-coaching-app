Set-StrictMode -Version Latest

# --- Exercices ---

function Get-Exercices {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [string] $Recherche
    )

    $query = "SELECT * FROM exercices"
    $params = @{}
    if ($Recherche) {
        $query += " WHERE nom LIKE @Recherche OR muscle_cible LIKE @Recherche"
        $params.Recherche = "%$Recherche%"
    }
    $query += " ORDER BY muscle_cible, nom"
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters $params
}

function New-Exercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [string] $Nom,
        [string] $MuscleCible,
        [string] $Variante,
        [string] $LienVideo,
        [string] $ImagePath
    )

    $query = @"
INSERT INTO exercices (nom, muscle_cible, variante, lien_video, image_path)
VALUES (@Nom, @MuscleCible, @Variante, @LienVideo, @ImagePath);
SELECT last_insert_rowid() AS id;
"@
    $result = Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        Nom = $Nom; MuscleCible = $MuscleCible; Variante = $Variante; LienVideo = $LienVideo; ImagePath = $ImagePath
    }
    return $result.id
}

function Update-Exercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [string] $Nom,
        [string] $MuscleCible,
        [string] $Variante,
        [string] $LienVideo,
        [string] $ImagePath
    )

    Invoke-SqliteQuery -DataSource $DbPath -Query @"
UPDATE exercices SET nom = @Nom, muscle_cible = @MuscleCible, variante = @Variante,
    lien_video = @LienVideo, image_path = @ImagePath
WHERE id = @Id
"@ -SqlParameters @{ Id = $Id; Nom = $Nom; MuscleCible = $MuscleCible; Variante = $Variante; LienVideo = $LienVideo; ImagePath = $ImagePath }
}

function Remove-Exercice {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM exercices WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

# --- Aliments ---

function Get-Aliments {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [string] $Recherche
    )

    $query = "SELECT * FROM aliments"
    $params = @{}
    if ($Recherche) {
        $query += " WHERE nom LIKE @Recherche"
        $params.Recherche = "%$Recherche%"
    }
    $query += " ORDER BY nom"
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters $params
}

function New-Aliment {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [string] $Nom,
        [double] $QuantiteReference = 100,
        [string] $Unite = 'g',
        [double] $Kcal,
        [double] $Proteines,
        [double] $Glucides,
        [double] $Lipides,
        [double] $Fibres
    )

    $query = @"
INSERT INTO aliments (nom, quantite_reference, unite, kcal, proteines, glucides, lipides, fibres)
VALUES (@Nom, @QuantiteReference, @Unite, @Kcal, @Proteines, @Glucides, @Lipides, @Fibres);
SELECT last_insert_rowid() AS id;
"@
    $result = Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        Nom = $Nom; QuantiteReference = $QuantiteReference; Unite = $Unite
        Kcal = $Kcal; Proteines = $Proteines; Glucides = $Glucides; Lipides = $Lipides; Fibres = $Fibres
    }
    return $result.id
}

function Update-Aliment {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [string] $Nom,
        [double] $QuantiteReference = 100,
        [string] $Unite = 'g',
        [double] $Kcal,
        [double] $Proteines,
        [double] $Glucides,
        [double] $Lipides,
        [double] $Fibres
    )

    Invoke-SqliteQuery -DataSource $DbPath -Query @"
UPDATE aliments SET nom = @Nom, quantite_reference = @QuantiteReference, unite = @Unite,
    kcal = @Kcal, proteines = @Proteines, glucides = @Glucides, lipides = @Lipides, fibres = @Fibres
WHERE id = @Id
"@ -SqlParameters @{
        Id = $Id; Nom = $Nom; QuantiteReference = $QuantiteReference; Unite = $Unite
        Kcal = $Kcal; Proteines = $Proteines; Glucides = $Glucides; Lipides = $Lipides; Fibres = $Fibres
    }
}

function Remove-Aliment {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM aliments WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

# --- Complements ---

function Get-Complements {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [string] $Recherche
    )

    $query = "SELECT * FROM complements"
    $params = @{}
    if ($Recherche) {
        $query += " WHERE nom LIKE @Recherche"
        $params.Recherche = "%$Recherche%"
    }
    $query += " ORDER BY nom"
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters $params
}

function New-Complement {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [string] $Nom,
        [string] $Dose,
        [string] $Timing,
        [string] $Note,
        [string] $Lien
    )

    $query = @"
INSERT INTO complements (nom, dose, timing, note, lien)
VALUES (@Nom, @Dose, @Timing, @Note, @Lien);
SELECT last_insert_rowid() AS id;
"@
    $result = Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        Nom = $Nom; Dose = $Dose; Timing = $Timing; Note = $Note; Lien = $Lien
    }
    return $result.id
}

function Update-Complement {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [string] $Nom,
        [string] $Dose,
        [string] $Timing,
        [string] $Note,
        [string] $Lien
    )

    Invoke-SqliteQuery -DataSource $DbPath -Query @"
UPDATE complements SET nom = @Nom, dose = @Dose, timing = @Timing, note = @Note, lien = @Lien
WHERE id = @Id
"@ -SqlParameters @{ Id = $Id; Nom = $Nom; Dose = $Dose; Timing = $Timing; Note = $Note; Lien = $Lien }
}

function Remove-Complement {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM complements WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

Export-ModuleMember -Function Get-Exercices, New-Exercice, Update-Exercice, Remove-Exercice, `
    Get-Aliments, New-Aliment, Update-Aliment, Remove-Aliment, `
    Get-Complements, New-Complement, Update-Complement, Remove-Complement
