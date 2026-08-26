Set-StrictMode -Version Latest

# --- Plans nutrition ---

function Get-PlansNutrition {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM plans_nutrition WHERE client_id = @ClientId ORDER BY date_debut DESC" -SqlParameters @{ ClientId = $ClientId }
}

function New-PlanNutrition {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $Nom,
        [string] $DateDebut,
        [string] $Notes
    )
    $query = @"
INSERT INTO plans_nutrition (client_id, nom, date_debut, notes) VALUES (@ClientId, @Nom, @DateDebut, @Notes);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ ClientId = $ClientId; Nom = $Nom; DateDebut = $DateDebut; Notes = $Notes }).id
}

function Remove-PlanNutrition {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    $query = @"
DELETE FROM repas_aliments WHERE repas_id IN (
    SELECT r.id FROM repas r JOIN types_jour tj ON tj.id = r.type_jour_id WHERE tj.plan_nutrition_id = @Id
);
DELETE FROM repas WHERE type_jour_id IN (SELECT id FROM types_jour WHERE plan_nutrition_id = @Id);
DELETE FROM types_jour WHERE plan_nutrition_id = @Id;
DELETE FROM plans_nutrition WHERE id = @Id;
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ Id = $Id }
}

# --- Types de jour ---

function Get-TypesJour {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $PlanNutritionId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM types_jour WHERE plan_nutrition_id = @PlanNutritionId ORDER BY id" -SqlParameters @{ PlanNutritionId = $PlanNutritionId }
}

function New-TypeJour {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $PlanNutritionId,
        [Parameter(Mandatory)] [string] $Nom
    )
    $query = @"
INSERT INTO types_jour (plan_nutrition_id, nom) VALUES (@PlanNutritionId, @Nom);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ PlanNutritionId = $PlanNutritionId; Nom = $Nom }).id
}

function Remove-TypeJour {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    $query = @"
DELETE FROM repas_aliments WHERE repas_id IN (SELECT id FROM repas WHERE type_jour_id = @Id);
DELETE FROM repas WHERE type_jour_id = @Id;
DELETE FROM types_jour WHERE id = @Id;
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ Id = $Id }
}

# --- Repas ---

function Get-Repas {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $TypeJourId
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM repas WHERE type_jour_id = @TypeJourId ORDER BY ordre, id" -SqlParameters @{ TypeJourId = $TypeJourId }
}

function New-Repas {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $TypeJourId,
        [Parameter(Mandatory)] [string] $Nom
    )
    $ordreMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(ordre), -1) AS m FROM repas WHERE type_jour_id = @TypeJourId" -SqlParameters @{ TypeJourId = $TypeJourId }).m
    $query = @"
INSERT INTO repas (type_jour_id, nom, ordre) VALUES (@TypeJourId, @Nom, @Ordre);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ TypeJourId = $TypeJourId; Nom = $Nom; Ordre = ($ordreMax + 1) }).id
}

function Remove-Repas {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM repas_aliments WHERE repas_id = @Id; DELETE FROM repas WHERE id = @Id;" -SqlParameters @{ Id = $Id }
}

# --- Aliments d'un repas ---

function Get-RepasAliments {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $RepasId
    )
    $query = @"
SELECT ra.*, a.nom AS aliment_nom, a.quantite_reference, a.unite,
    ROUND(a.kcal * ra.quantite / a.quantite_reference, 1) AS kcal_calc,
    ROUND(a.proteines * ra.quantite / a.quantite_reference, 1) AS proteines_calc,
    ROUND(a.glucides * ra.quantite / a.quantite_reference, 1) AS glucides_calc,
    ROUND(a.lipides * ra.quantite / a.quantite_reference, 1) AS lipides_calc,
    ROUND(a.fibres * ra.quantite / a.quantite_reference, 1) AS fibres_calc
FROM repas_aliments ra
JOIN aliments a ON a.id = ra.aliment_id
WHERE ra.repas_id = @RepasId
ORDER BY ra.ordre, ra.id
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ RepasId = $RepasId }
}

function New-RepasAliment {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $RepasId,
        [Parameter(Mandatory)] [int] $AlimentId,
        [Parameter(Mandatory)] [double] $Quantite
    )
    $ordreMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(ordre), -1) AS m FROM repas_aliments WHERE repas_id = @RepasId" -SqlParameters @{ RepasId = $RepasId }).m
    $query = @"
INSERT INTO repas_aliments (repas_id, aliment_id, quantite, ordre) VALUES (@RepasId, @AlimentId, @Quantite, @Ordre);
SELECT last_insert_rowid() AS id;
"@
    (Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ RepasId = $RepasId; AlimentId = $AlimentId; Quantite = $Quantite; Ordre = ($ordreMax + 1) }).id
}

function Remove-RepasAliment {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM repas_aliments WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

function Get-TotauxRepas {
    <# Calcule les totaux kcal/macros d'un repas a partir des lignes retournees par Get-RepasAliments. #>
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [array] $Lignes)
    $totaux = [pscustomobject]@{
        Kcal = 0.0; Proteines = 0.0; Glucides = 0.0; Lipides = 0.0; Fibres = 0.0
    }
    foreach ($l in $Lignes) {
        $totaux.Kcal += [double]$l.kcal_calc
        $totaux.Proteines += [double]$l.proteines_calc
        $totaux.Glucides += [double]$l.glucides_calc
        $totaux.Lipides += [double]$l.lipides_calc
        $totaux.Fibres += [double]$l.fibres_calc
    }
    $totaux
}

Export-ModuleMember -Function Get-PlansNutrition, New-PlanNutrition, Remove-PlanNutrition, `
    Get-TypesJour, New-TypeJour, Remove-TypeJour, `
    Get-Repas, New-Repas, Remove-Repas, `
    Get-RepasAliments, New-RepasAliment, Remove-RepasAliment, Get-TotauxRepas
