Set-StrictMode -Version Latest

# --- Devis ---

function Get-Devis {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [int] $ClientId
    )

    $query = @"
SELECT d.*, c.nom || ' ' || c.prenom AS client_nom
FROM devis d
JOIN clients c ON c.id = d.client_id
"@
    if ($ClientId) {
        $query += " WHERE d.client_id = @ClientId"
        $query += " ORDER BY d.date_creation DESC"
        return Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ ClientId = $ClientId }
    }
    $query += " ORDER BY d.date_creation DESC"
    Invoke-SqliteQuery -DataSource $DbPath -Query $query
}

function New-Devis {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $Prestations,
        [Parameter(Mandatory)] [double] $Montant,
        [string] $Duree
    )

    $query = @"
INSERT INTO devis (client_id, prestations, montant, duree)
VALUES (@ClientId, @Prestations, @Montant, @Duree);
SELECT last_insert_rowid() AS id;
"@
    $result = Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        ClientId = $ClientId; Prestations = $Prestations; Montant = $Montant; Duree = $Duree
    }
    return $result.id
}

function Set-DevisStatut {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [ValidateSet('en_attente', 'accepte', 'refuse')] [string] $Statut
    )

    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE devis SET statut = @Statut WHERE id = @Id" -SqlParameters @{ Statut = $Statut; Id = $Id }
}

# --- Commandes ---

function Get-Commandes {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [int] $ClientId
    )

    $query = @"
SELECT cmd.*, c.nom || ' ' || c.prenom AS client_nom
FROM commandes cmd
JOIN clients c ON c.id = cmd.client_id
"@
    if ($ClientId) {
        $query += " WHERE cmd.client_id = @ClientId ORDER BY cmd.date_debut DESC"
        return Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ ClientId = $ClientId }
    }
    $query += " ORDER BY cmd.date_debut DESC"
    Invoke-SqliteQuery -DataSource $DbPath -Query $query
}

function New-Commande {
    <#
        Cree une commande et genere automatiquement son echeancier.
        - one_shot      : une seule echeance a DateDebut
        - hebdomadaire  : une echeance tous les 7 jours entre DateDebut et DateFin
        - mensuel       : une echeance par mois (meme jour que DateDebut) entre DateDebut et DateFin
        Le montant total est reparti a parts egales sur les echeances (arrondi au centime,
        la derniere echeance absorbe l'ecart d'arrondi).
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [int] $DevisId,
        [Parameter(Mandatory)] [ValidateSet('mensuel', 'hebdomadaire', 'one_shot')] [string] $TypeFacturation,
        [Parameter(Mandatory)] [datetime] $DateDebut,
        [datetime] $DateFin,
        [Parameter(Mandatory)] [double] $Montant
    )

    if ($TypeFacturation -ne 'one_shot' -and -not $DateFin) {
        throw "Une date de fin est requise pour une facturation mensuelle ou hebdomadaire."
    }

    $devisIdValue = if ($DevisId) { $DevisId } else { [DBNull]::Value }
    $dateFinValue = if ($DateFin) { $DateFin.ToString('yyyy-MM-dd') } else { [DBNull]::Value }

    $insertQuery = @"
INSERT INTO commandes (client_id, devis_id, type_facturation, date_debut, date_fin, montant)
VALUES (@ClientId, @DevisId, @TypeFacturation, @DateDebut, @DateFin, @Montant);
SELECT last_insert_rowid() AS id;
"@
    $result = Invoke-SqliteQuery -DataSource $DbPath -Query $insertQuery -SqlParameters @{
        ClientId = $ClientId; DevisId = $devisIdValue; TypeFacturation = $TypeFacturation
        DateDebut = $DateDebut.ToString('yyyy-MM-dd'); DateFin = $dateFinValue; Montant = $Montant
    }
    $commandeId = $result.id

    $dates = New-Object System.Collections.Generic.List[datetime]
    switch ($TypeFacturation) {
        'one_shot' {
            $dates.Add($DateDebut)
        }
        'hebdomadaire' {
            $current = $DateDebut
            while ($current -le $DateFin) {
                $dates.Add($current)
                $current = $current.AddDays(7)
            }
        }
        'mensuel' {
            $current = $DateDebut
            while ($current -le $DateFin) {
                $dates.Add($current)
                $current = $current.AddMonths(1)
            }
        }
    }

    if ($dates.Count -eq 0) { $dates.Add($DateDebut) }

    $montantParEcheance = [math]::Round($Montant / $dates.Count, 2)
    for ($i = 0; $i -lt $dates.Count; $i++) {
        $montantEcheance = if ($i -eq $dates.Count - 1) {
            [math]::Round($Montant - ($montantParEcheance * ($dates.Count - 1)), 2)
        } else {
            $montantParEcheance
        }
        Invoke-SqliteQuery -DataSource $DbPath -Query @"
INSERT INTO echeances (commande_id, date_echeance, montant)
VALUES (@CommandeId, @DateEcheance, @Montant)
"@ -SqlParameters @{ CommandeId = $commandeId; DateEcheance = $dates[$i].ToString('yyyy-MM-dd'); Montant = $montantEcheance }
    }

    return $commandeId
}

# --- Echeances ---

function Update-EcheancesRetard {
    param(
        [Parameter(Mandatory)] [string] $DbPath
    )

    Invoke-SqliteQuery -DataSource $DbPath -Query @"
UPDATE echeances
SET statut = 'en_retard'
WHERE statut = 'en_attente' AND date_echeance < date('now')
"@
}

function Get-Echeances {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [int] $ClientId,
        [int] $CommandeId
    )

    $query = @"
SELECT e.*, cmd.type_facturation, c.id AS client_id, c.nom || ' ' || c.prenom AS client_nom
FROM echeances e
JOIN commandes cmd ON cmd.id = e.commande_id
JOIN clients c ON c.id = cmd.client_id
"@
    $conditions = New-Object System.Collections.Generic.List[string]
    $params = @{}
    if ($ClientId) { $conditions.Add("c.id = @ClientId"); $params.ClientId = $ClientId }
    if ($CommandeId) { $conditions.Add("cmd.id = @CommandeId"); $params.CommandeId = $CommandeId }
    if ($conditions.Count -gt 0) {
        $query += " WHERE " + ($conditions -join " AND ")
    }
    $query += " ORDER BY e.date_echeance"

    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters $params
}

function Set-EcheanceStatut {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [ValidateSet('payee', 'en_attente', 'en_retard')] [string] $Statut
    )

    $datePaiement = if ($Statut -eq 'payee') { (Get-Date).ToString('yyyy-MM-dd') } else { [DBNull]::Value }
    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE echeances SET statut = @Statut, date_paiement = @DatePaiement WHERE id = @Id" `
        -SqlParameters @{ Statut = $Statut; DatePaiement = $datePaiement; Id = $Id }
}

Export-ModuleMember -Function Get-Devis, New-Devis, Set-DevisStatut, Get-Commandes, New-Commande, Update-EcheancesRetard, Get-Echeances, Set-EcheanceStatut
