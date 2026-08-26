Set-StrictMode -Version Latest

function Get-StatsDashboard {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [int] $JoursSansBilanSeuil = 14
    )

    $clientsActifs = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COUNT(*) AS n FROM clients WHERE statut = 'actif'").n

    $enAttente = Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COUNT(*) AS n, COALESCE(SUM(montant), 0) AS total FROM echeances WHERE statut = 'en_attente'"
    $enRetard = Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COUNT(*) AS n, COALESCE(SUM(montant), 0) AS total FROM echeances WHERE statut = 'en_retard'"

    $prochaines = Invoke-SqliteQuery -DataSource $DbPath -Query @"
SELECT e.date_echeance, e.montant, c.nom || ' ' || c.prenom AS client_nom
FROM echeances e
JOIN commandes cmd ON cmd.id = e.commande_id
JOIN clients c ON c.id = cmd.client_id
WHERE e.statut IN ('en_attente', 'en_retard')
ORDER BY e.date_echeance
LIMIT 10
"@

    $clientsSansBilan = Invoke-SqliteQuery -DataSource $DbPath -Query @"
SELECT c.id, c.nom || ' ' || c.prenom AS client_nom,
    (SELECT MAX(date_reponse) FROM questionnaires_reponses WHERE client_id = c.id AND type = 'bilan') AS dernier_bilan,
    c.date_debut_coaching
FROM clients c
WHERE c.statut = 'actif'
"@ | Where-Object {
        $reference = if ($_.dernier_bilan) { $_.dernier_bilan } else { $_.date_debut_coaching }
        if (-not $reference) { return $false }
        ((Get-Date) - [datetime]$reference).Days -ge $JoursSansBilanSeuil
    } | ForEach-Object {
        $reference = if ($_.dernier_bilan) { $_.dernier_bilan } else { $_.date_debut_coaching }
        [pscustomobject]@{
            client_nom = $_.client_nom
            dernier_bilan = if ($_.dernier_bilan) { $_.dernier_bilan } else { 'Jamais' }
            jours_ecoules = ((Get-Date) - [datetime]$reference).Days
        }
    }

    [pscustomobject]@{
        ClientsActifs = [int]$clientsActifs
        PaiementsEnAttenteNombre = [int]$enAttente.n
        PaiementsEnAttenteMontant = [double]$enAttente.total
        PaiementsEnRetardNombre = [int]$enRetard.n
        PaiementsEnRetardMontant = [double]$enRetard.total
        ProchainesEcheances = @($prochaines)
        ClientsSansBilanRecent = @($clientsSansBilan)
    }
}

Export-ModuleMember -Function Get-StatsDashboard
