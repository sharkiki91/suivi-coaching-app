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

function Remove-Client {
    <#
        Supprime definitivement un client et TOUTES ses donnees dans l'application (devis, commandes et
        echeances, questionnaires, complements recommandes, programmes/seances/exercices, seances
        realisees, plans nutrition/types de jour/repas/aliments, roadmap, suivi quotidien, journal
        alimentaire). Irreversible : appeler Backup-Database avant, cote appelant, si on veut pouvoir
        revenir en arriere.
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    $query = @"
DELETE FROM echeances WHERE commande_id IN (SELECT id FROM commandes WHERE client_id = @Id);
DELETE FROM commandes WHERE client_id = @Id;
DELETE FROM devis WHERE client_id = @Id;
DELETE FROM questionnaires_reponses WHERE client_id = @Id;
DELETE FROM complements_recommandations WHERE client_id = @Id;
DELETE FROM exercices_realises WHERE seance_realisee_id IN (SELECT id FROM seances_realisees WHERE client_id = @Id);
DELETE FROM seances_realisees WHERE client_id = @Id;
DELETE FROM seance_exercices WHERE seance_id IN (SELECT id FROM seances WHERE programme_id IN (SELECT id FROM programmes WHERE client_id = @Id));
DELETE FROM seances WHERE programme_id IN (SELECT id FROM programmes WHERE client_id = @Id);
DELETE FROM programmes WHERE client_id = @Id;
DELETE FROM repas_aliments WHERE repas_id IN (SELECT id FROM repas WHERE type_jour_id IN (SELECT id FROM types_jour WHERE plan_nutrition_id IN (SELECT id FROM plans_nutrition WHERE client_id = @Id)));
DELETE FROM repas WHERE type_jour_id IN (SELECT id FROM types_jour WHERE plan_nutrition_id IN (SELECT id FROM plans_nutrition WHERE client_id = @Id));
DELETE FROM types_jour WHERE plan_nutrition_id IN (SELECT id FROM plans_nutrition WHERE client_id = @Id);
DELETE FROM plans_nutrition WHERE client_id = @Id;
DELETE FROM roadmap_semaines WHERE client_id = @Id;
DELETE FROM suivi_quotidien WHERE client_id = @Id;
DELETE FROM journal_alimentaire WHERE client_id = @Id;
DELETE FROM clients WHERE id = @Id;
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ Id = $Id }
}

function Set-ClientStatut {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [ValidateSet('actif', 'archive')] [string] $Statut
    )

    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE clients SET statut = @Statut WHERE id = @Id" -SqlParameters @{ Statut = $Statut; Id = $Id }
}

Export-ModuleMember -Function Get-Clients, Get-Client, New-Client, Update-Client, Remove-Client, Set-ClientChampsSiVide, Set-ClientStatut
