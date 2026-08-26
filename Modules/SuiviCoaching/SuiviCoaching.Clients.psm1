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

function Set-ClientChampsSiVide {
    <#
        Complete Telephone / Email / Objectifs sur la fiche d'un client, mais uniquement
        pour les champs actuellement vides (une valeur deja saisie n'est jamais ecrasee).
        Utilise par l'import du questionnaire pre-coaching. Renvoie $true si au moins un
        champ a ete complete.
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [string] $Telephone,
        [string] $Email,
        [string] $Objectifs
    )

    $client = Get-Client -DbPath $DbPath -Id $Id
    if (-not $client) { return $false }

    $modifie = $false
    $nouveauTelephone = $client.telephone
    if ([string]::IsNullOrWhiteSpace($client.telephone) -and -not [string]::IsNullOrWhiteSpace($Telephone)) {
        $nouveauTelephone = $Telephone.Trim(); $modifie = $true
    }
    $nouvelEmail = $client.email
    if ([string]::IsNullOrWhiteSpace($client.email) -and -not [string]::IsNullOrWhiteSpace($Email)) {
        $nouvelEmail = $Email.Trim(); $modifie = $true
    }
    $nouveauxObjectifs = $client.objectifs
    if ([string]::IsNullOrWhiteSpace($client.objectifs) -and -not [string]::IsNullOrWhiteSpace($Objectifs)) {
        $nouveauxObjectifs = $Objectifs.Trim(); $modifie = $true
    }

    if (-not $modifie) { return $false }

    Invoke-SqliteQuery -DataSource $DbPath -Query @"
UPDATE clients SET telephone = @Telephone, email = @Email, objectifs = @Objectifs WHERE id = @Id
"@ -SqlParameters @{ Telephone = $nouveauTelephone; Email = $nouvelEmail; Objectifs = $nouveauxObjectifs; Id = $Id }
    return $true
}

function Set-ClientStatut {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [ValidateSet('actif', 'archive')] [string] $Statut
    )

    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE clients SET statut = @Statut WHERE id = @Id" -SqlParameters @{ Statut = $Statut; Id = $Id }
}

Export-ModuleMember -Function Get-Clients, Get-Client, New-Client, Update-Client, Set-ClientChampsSiVide, Set-ClientStatut
