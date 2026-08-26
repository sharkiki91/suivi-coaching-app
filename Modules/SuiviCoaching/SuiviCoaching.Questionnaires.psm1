Set-StrictMode -Version Latest

function Get-QuestionnairesReponses {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [int] $ClientId,
        [switch] $NonRattachees
    )
    $query = "SELECT qr.*, c.nom || ' ' || c.prenom AS client_nom FROM questionnaires_reponses qr LEFT JOIN clients c ON c.id = qr.client_id"
    $params = @{}
    if ($NonRattachees) {
        $query += " WHERE qr.client_id IS NULL"
    } elseif ($ClientId) {
        $query += " WHERE qr.client_id = @ClientId"
        $params.ClientId = $ClientId
    }
    $query += " ORDER BY qr.date_reponse DESC"
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters $params
}

function New-QuestionnaireReponse {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [int] $ClientId,
        [Parameter(Mandatory)] [ValidateSet('pre_coaching', 'bilan')] [string] $Type,
        [Parameter(Mandatory)] [string] $DateReponse,
        [Parameter(Mandatory)] [string] $DonneesJson,
        [string] $FichierSource
    )
    $clientIdValue = if ($ClientId) { $ClientId } else { [DBNull]::Value }
    $query = @"
INSERT INTO questionnaires_reponses (client_id, type, date_reponse, donnees_json, fichier_source)
VALUES (@ClientId, @Type, @DateReponse, @DonneesJson, @FichierSource);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
        ClientId = $clientIdValue; Type = $Type; DateReponse = $DateReponse; DonneesJson = $DonneesJson; FichierSource = $FichierSource
    }).id
}

function Set-QuestionnaireReponseClient {
    <# Rattache (ou modifie le rattachement) d'une reponse a un client, manuellement. #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [int] $ClientId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE questionnaires_reponses SET client_id = @ClientId WHERE id = @Id" -SqlParameters @{ ClientId = $ClientId; Id = $Id }
}

function Remove-QuestionnaireReponse {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM questionnaires_reponses WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

function Get-ImportMapping {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [string] $Type
    )
    $row = Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM import_mappings WHERE type = @Type ORDER BY date_maj DESC LIMIT 1" -SqlParameters @{ Type = $Type }
    if ($row) { return $row.mapping_json | ConvertFrom-Json }
    return $null
}

function Save-ImportMapping {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [string] $Type,
        [Parameter(Mandatory)] [string] $MappingJson
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM import_mappings WHERE type = @Type" -SqlParameters @{ Type = $Type }
    Invoke-SqliteQuery -DataSource $DbPath -Query "INSERT INTO import_mappings (type, mapping_json) VALUES (@Type, @MappingJson)" -SqlParameters @{ Type = $Type; MappingJson = $MappingJson }
}

Export-ModuleMember -Function Get-QuestionnairesReponses, New-QuestionnaireReponse, Set-QuestionnaireReponseClient, Remove-QuestionnaireReponse, `
    Get-ImportMapping, Save-ImportMapping
