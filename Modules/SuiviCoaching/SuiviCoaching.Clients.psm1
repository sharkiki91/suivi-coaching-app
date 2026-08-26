Set-StrictMode -Version Latest

function Get-Clients {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [switch] $InclureArchives
    )

    $query = "SELECT * FROM clients"
    if (-not $InclureArchives) {
        $query += " WHERE statut = 'actif'"
    }
    $query += " ORDER BY nom, prenom"

    Invoke-SqliteQuery -DataSource $DbPath -Query $query
}

function Get-Client {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )

    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM clients WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

function New-Client {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [string] $Nom,
        [Parameter(Mandatory)] [string] $Prenom,
        [string] $Email,
        [string] $Telephone,
        [string] $DateDebutCoaching,
        [string] $Objectifs,
        [string] $Notes
    )

    $query = @"
INSERT INTO clients (nom, prenom, email, telephone, date_debut_coaching, objectifs, notes)
VALUES (@Nom, @Prenom, @Email, @Telephone, @DateDebutCoaching, @Objectifs, @Notes);
SELECT last_insert_rowid() AS id;
"@

    $result = Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        Nom = $Nom; Prenom = $Prenom; Email = $Email; Telephone = $Telephone
        DateDebutCoaching = $DateDebutCoaching; Objectifs = $Objectifs; Notes = $Notes
    }
    return $result.id
}

function Update-Client {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [string] $Nom,
        [Parameter(Mandatory)] [string] $Prenom,
        [string] $Email,
        [string] $Telephone,
        [string] $DateDebutCoaching,
        [string] $Objectifs,
        [string] $Notes
    )

    $query = @"
UPDATE clients
SET nom = @Nom, prenom = @Prenom, email = @Email, telephone = @Telephone,
    date_debut_coaching = @DateDebutCoaching, objectifs = @Objectifs, notes = @Notes
WHERE id = @Id
"@

    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        Id = $Id; Nom = $Nom; Prenom = $Prenom; Email = $Email; Telephone = $Telephone
        DateDebutCoaching = $DateDebutCoaching; Objectifs = $Objectifs; Notes = $Notes
    }
}

function Set-ClientStatut {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [ValidateSet('actif', 'archive')] [string] $Statut
    )

    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE clients SET statut = @Statut WHERE id = @Id" -SqlParameters @{ Statut = $Statut; Id = $Id }
}

Export-ModuleMember -Function Get-Clients, Get-Client, New-Client, Update-Client, Set-ClientStatut
