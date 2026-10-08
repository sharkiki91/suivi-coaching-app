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

function Get-TotauxTypeJour {
    <# Totaux kcal/macros d'une journee entiere (tous les repas du type de jour). #>
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $TypeJourId)
    $lignes = @()
    foreach ($r in @(Get-Repas -DbPath $DbPath -TypeJourId $TypeJourId)) { $lignes += @(Get-RepasAliments -DbPath $DbPath -RepasId ([int]$r.id)) }
    Get-TotauxRepas -Lignes $lignes
}

function Update-RepasAliment {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $Id,
        [Parameter(Mandatory)] [int] $AlimentId,
        [Parameter(Mandatory)] [double] $Quantite
    )
    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE repas_aliments SET aliment_id = @AlimentId, quantite = @Quantite WHERE id = @Id" -SqlParameters @{ Id = $Id; AlimentId = $AlimentId; Quantite = $Quantite }
}

function Set-OrdreLignes {
    <# Renumerote 0, 1, 2... les lignes de $Table dans l'ordre des ids donnes. #>
    param([string] $DbPath, [string] $Table, [int[]] $Ids)
    for ($i = 0; $i -lt $Ids.Count; $i++) {
        Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE $Table SET ordre = @Ordre WHERE id = @Id" -SqlParameters @{ Ordre = $i; Id = $Ids[$i] }
    }
}

function Move-LigneOrdonnee {
    <# Monte (-1) ou descend (1) une ligne parmi ses soeurs (meme $ColonneParent), en renumerotant l'ordre. #>
    param([string] $DbPath, [string] $Table, [string] $ColonneParent, [int] $Id, [int] $Direction)
    $parent = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT $ColonneParent AS p FROM $Table WHERE id = @Id" -SqlParameters @{ Id = $Id }).p
    if ($null -eq $parent) { return }
    $ids = @(Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT id FROM $Table WHERE $ColonneParent = @P ORDER BY ordre, id" -SqlParameters @{ P = $parent } | ForEach-Object { [int]$_.id })
    $index = [array]::IndexOf($ids, $Id)
    $cible = $index + $Direction
    if ($index -lt 0 -or $cible -lt 0 -or $cible -ge $ids.Count) { return }
    $ids[$index] = $ids[$cible]; $ids[$cible] = $Id
    Set-OrdreLignes -DbPath $DbPath -Table $Table -Ids $ids
}

function Move-Repas {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id, [Parameter(Mandatory)] [ValidateSet(-1, 1)] [int] $Direction)
    Move-LigneOrdonnee -DbPath $DbPath -Table 'repas' -ColonneParent 'type_jour_id' -Id $Id -Direction $Direction
}

function Move-RepasAliment {
    <#
        Monte/descend une ligne d'un repas. Les ingredients d'une recette ajoutee au repas forment un bloc :
        deplacer l'un d'eux deplace toute la recette, et une ligne seule saute une recette entiere.
    #>
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id, [Parameter(Mandatory)] [ValidateSet(-1, 1)] [int] $Direction)
    $repasId = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT repas_id AS p FROM repas_aliments WHERE id = @Id" -SqlParameters @{ Id = $Id }).p
    if ($null -eq $repasId) { return }
    $lignes = @(Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT id, recette_groupe FROM repas_aliments WHERE repas_id = @P ORDER BY ordre, id" -SqlParameters @{ P = $repasId })
    $blocs = New-Object System.Collections.ArrayList
    $groupePrecedent = $null
    foreach ($l in $lignes) {
        $groupe = if ($l.recette_groupe -is [DBNull] -or $null -eq $l.recette_groupe) { $null } else { [int]$l.recette_groupe }
        if ($null -ne $groupe -and $groupe -eq $groupePrecedent) { $blocs[$blocs.Count - 1].Add([int]$l.id) | Out-Null }
        else { $nouveau = New-Object System.Collections.ArrayList; $nouveau.Add([int]$l.id) | Out-Null; $blocs.Add($nouveau) | Out-Null }
        $groupePrecedent = $groupe
    }
    $index = -1
    for ($i = 0; $i -lt $blocs.Count; $i++) { if ($blocs[$i].Contains($Id)) { $index = $i } }
    $cible = $index + $Direction
    if ($index -lt 0 -or $cible -lt 0 -or $cible -ge $blocs.Count) { return }
    $tmp = $blocs[$index]; $blocs[$index] = $blocs[$cible]; $blocs[$cible] = $tmp
    [int[]]$ids = @(foreach ($b in $blocs) { foreach ($x in $b) { $x } })
    Set-OrdreLignes -DbPath $DbPath -Table 'repas_aliments' -Ids $ids
}

# --- Recettes (bibliotheque) ---

function Get-Recettes {
    param([Parameter(Mandatory)] [string] $DbPath)
    Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM recettes ORDER BY nom"
}

function New-Recette {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [string] $Nom, [string] $Notes)
    (Invoke-SqliteQuery -DataSource $DbPath -Query "INSERT INTO recettes (nom, notes) VALUES (@Nom, @Notes); SELECT last_insert_rowid() AS id;" -SqlParameters @{ Nom = $Nom; Notes = $Notes }).id
}

function Update-Recette {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id, [Parameter(Mandatory)] [string] $Nom, [string] $Notes)
    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE recettes SET nom = @Nom, notes = @Notes WHERE id = @Id" -SqlParameters @{ Id = $Id; Nom = $Nom; Notes = $Notes }
}

function Remove-Recette {
    <# Supprime la recette de la bibliotheque ; les repas ou elle a deja ete ajoutee gardent leur copie. #>
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id)
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM recette_ingredients WHERE recette_id = @Id; DELETE FROM recettes WHERE id = @Id;" -SqlParameters @{ Id = $Id }
}

function Get-RecetteIngredients {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $RecetteId)
    $query = @"
SELECT ri.*, a.nom AS aliment_nom, a.quantite_reference, a.unite,
    ROUND(a.kcal * ri.quantite / a.quantite_reference, 1) AS kcal_calc,
    ROUND(a.proteines * ri.quantite / a.quantite_reference, 1) AS proteines_calc,
    ROUND(a.glucides * ri.quantite / a.quantite_reference, 1) AS glucides_calc,
    ROUND(a.lipides * ri.quantite / a.quantite_reference, 1) AS lipides_calc,
    ROUND(a.fibres * ri.quantite / a.quantite_reference, 1) AS fibres_calc
FROM recette_ingredients ri
JOIN aliments a ON a.id = ri.aliment_id
WHERE ri.recette_id = @RecetteId
ORDER BY ri.ordre, ri.id
"@
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{ RecetteId = $RecetteId }
}

function New-RecetteIngredient {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $RecetteId, [Parameter(Mandatory)] [int] $AlimentId, [Parameter(Mandatory)] [double] $Quantite)
    $ordreMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(ordre), -1) AS m FROM recette_ingredients WHERE recette_id = @R" -SqlParameters @{ R = $RecetteId }).m
    (Invoke-SqliteQuery -DataSource $DbPath -Query "INSERT INTO recette_ingredients (recette_id, aliment_id, quantite, ordre) VALUES (@R, @A, @Q, @O); SELECT last_insert_rowid() AS id;" `
        -SqlParameters @{ R = $RecetteId; A = $AlimentId; Q = $Quantite; O = ($ordreMax + 1) }).id
}

function Update-RecetteIngredient {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id, [Parameter(Mandatory)] [int] $AlimentId, [Parameter(Mandatory)] [double] $Quantite)
    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE recette_ingredients SET aliment_id = @A, quantite = @Q WHERE id = @Id" -SqlParameters @{ Id = $Id; A = $AlimentId; Q = $Quantite }
}

function Remove-RecetteIngredient {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id)
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM recette_ingredients WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

function Move-RecetteIngredient {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id, [Parameter(Mandatory)] [ValidateSet(-1, 1)] [int] $Direction)
    Move-LigneOrdonnee -DbPath $DbPath -Table 'recette_ingredients' -ColonneParent 'recette_id' -Id $Id -Direction $Direction
}

function Add-RecetteAuRepas {
    <#
        Copie les ingredients d'une recette a la fin d'un repas (quantites x $Portions). Copie ponctuelle,
        comme les modeles de seance : modifier la recette ensuite ne change pas les repas deja faits.
        Retourne le numero de groupe commun aux lignes ajoutees.
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $RepasId,
        [Parameter(Mandatory)] [int] $RecetteId,
        [double] $Portions = 1
    )
    if ($Portions -le 0) { throw "Le nombre de portions doit etre superieur a 0." }
    $recette = Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT nom FROM recettes WHERE id = @Id" -SqlParameters @{ Id = $RecetteId }
    if (-not $recette) { throw "Recette introuvable." }
    $ingredients = @(Get-RecetteIngredients -DbPath $DbPath -RecetteId $RecetteId)
    if ($ingredients.Count -eq 0) { throw "Cette recette n'a encore aucun ingredient." }
    $nom = [string]$recette.nom
    if ($Portions -ne 1) { $nom += " (x$($Portions.ToString([System.Globalization.CultureInfo]::InvariantCulture).Replace('.', ',')))" }
    $groupe = [int](Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(recette_groupe), 0) + 1 AS g FROM repas_aliments").g
    $ordre = [int](Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(ordre), -1) AS m FROM repas_aliments WHERE repas_id = @R" -SqlParameters @{ R = $RepasId }).m
    foreach ($i in $ingredients) {
        $ordre++
        Invoke-SqliteQuery -DataSource $DbPath -Query @"
INSERT INTO repas_aliments (repas_id, aliment_id, quantite, ordre, recette_nom, recette_groupe) VALUES (@R, @A, @Q, @O, @N, @G)
"@ -SqlParameters @{ R = $RepasId; A = [int]$i.aliment_id; Q = [math]::Round([double]$i.quantite * $Portions, 1); O = $ordre; N = $nom; G = $groupe }
    }
    return $groupe
}

function Remove-RecetteDuRepas {
    <# Retire du repas tous les ingredients d'une recette ajoutee (meme groupe). #>
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $RepasId, [Parameter(Mandatory)] [int] $Groupe)
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM repas_aliments WHERE repas_id = @R AND recette_groupe = @G" -SqlParameters @{ R = $RepasId; G = $Groupe }
}

# --- Tableau d'equivalences (bibliotheque) ---

function Get-BaseEquivalence {
    <#
        Nutriment sur lequel deux aliments sont juges equivalents : les calories pour un aliment peu
        calorique (legumes), sinon le macro qui apporte le plus d'energie (riz -> glucides, poulet -> proteines).
    #>
    param([Parameter(Mandatory)] $Aliment)
    $ref = [double]$Aliment.quantite_reference; if ($ref -le 0) { $ref = 100 }
    $kcal100 = [double](Get-NombreOuZero $Aliment.kcal) * 100 / $ref
    if ($kcal100 -lt 60) { return 'kcal' }
    $energie = @{
        proteines = 4 * [double](Get-NombreOuZero $Aliment.proteines)
        glucides = 4 * [double](Get-NombreOuZero $Aliment.glucides)
        lipides = 9 * [double](Get-NombreOuZero $Aliment.lipides)
    }
    ($energie.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
}

function Get-NombreOuZero {
    param($Valeur)
    if ($null -eq $Valeur -or $Valeur -is [DBNull] -or [string]::IsNullOrWhiteSpace([string]$Valeur)) { return 0.0 }
    [double]$Valeur
}

function Get-QuantiteEquivalente {
    <# Quantite de $Aliment apportant autant de $Base que $QuantiteReference de $Reference (arrondie a 5 g/ml pres). Null si impossible. #>
    param([Parameter(Mandatory)] $Reference, [Parameter(Mandatory)] [double] $QuantiteReference, [Parameter(Mandatory)] $Aliment, [Parameter(Mandatory)] [string] $Base)
    $parUniteRef = (Get-NombreOuZero $Reference.$Base) / [double]$Reference.quantite_reference
    $parUnite = (Get-NombreOuZero $Aliment.$Base) / [double]$Aliment.quantite_reference
    if ($parUnite -le 0 -or $parUniteRef -le 0) { return $null }
    $q = $QuantiteReference * $parUniteRef / $parUnite
    if ([string]$Aliment.unite -in @('g', 'ml')) { return [math]::Max(5, [math]::Round($q / 5) * 5) }
    [math]::Round($q, 1)
}

function Format-QuantiteAliment {
    param([double] $Quantite, [string] $Unite)
    $q = if ($Quantite -eq [math]::Floor($Quantite)) { [string][int]$Quantite } else { $Quantite.ToString('0.#', [System.Globalization.CultureInfo]::GetCultureInfo('fr-FR')) }
    if ($Unite -in @('g', 'ml')) { "$q $Unite" } else { "$q $Unite".Trim() }
}

function Get-EquivalencesGroupes {
    <#
        Groupes du tableau d'equivalences avec, pour chaque aliment equivalent, la quantite calculee
        depuis la bibliotheque d'aliments (meme apport de la base du groupe que l'aliment de reference).
    #>
    param([Parameter(Mandatory)] [string] $DbPath)
    $aliments = @{}
    foreach ($a in @(Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM aliments")) { $aliments[[int]$a.id] = $a }
    $libelles = @{ kcal = 'calories'; proteines = 'protéines'; glucides = 'glucides'; lipides = 'lipides' }
    $groupes = @(Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM equivalences_groupes ORDER BY ordre, id")
    foreach ($g in $groupes) {
        if (-not $aliments.ContainsKey([int]$g.aliment_id)) { continue }
        $ref = $aliments[[int]$g.aliment_id]
        $base = Get-BaseEquivalence -Aliment $ref
        $equivalents = foreach ($e in @(Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM equivalences_aliments WHERE groupe_id = @G ORDER BY ordre, id" -SqlParameters @{ G = [int]$g.id })) {
            if (-not $aliments.ContainsKey([int]$e.aliment_id)) { continue }
            $a = $aliments[[int]$e.aliment_id]
            $q = Get-QuantiteEquivalente -Reference $ref -QuantiteReference ([double]$g.quantite) -Aliment $a -Base $base
            [pscustomobject]@{
                id = [int]$e.id; aliment_id = [int]$e.aliment_id; aliment_nom = [string]$a.nom; quantite = $q; unite = [string]$a.unite
                quantite_affichage = $(if ($null -eq $q) { '-' } else { Format-QuantiteAliment $q ([string]$a.unite) })
                affichage = $(if ($null -eq $q) { [string]$a.nom } else { "$($a.nom) ($(Format-QuantiteAliment $q ([string]$a.unite)))" })
            }
        }
        [pscustomobject]@{
            id = [int]$g.id; aliment_id = [int]$g.aliment_id; aliment_nom = [string]$ref.nom; quantite = [double]$g.quantite; unite = [string]$ref.unite
            base = $base; base_libelle = $libelles[$base]
            affichage = "$($ref.nom) ($(Format-QuantiteAliment ([double]$g.quantite) ([string]$ref.unite)))"
            Equivalents = @($equivalents)
        }
    }
}

function New-EquivalenceGroupe {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $AlimentId, [Parameter(Mandatory)] [double] $Quantite)
    $ordreMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(ordre), -1) AS m FROM equivalences_groupes").m
    (Invoke-SqliteQuery -DataSource $DbPath -Query "INSERT INTO equivalences_groupes (aliment_id, quantite, ordre) VALUES (@A, @Q, @O); SELECT last_insert_rowid() AS id;" `
        -SqlParameters @{ A = $AlimentId; Q = $Quantite; O = ($ordreMax + 1) }).id
}

function Update-EquivalenceGroupe {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id, [Parameter(Mandatory)] [int] $AlimentId, [Parameter(Mandatory)] [double] $Quantite)
    Invoke-SqliteQuery -DataSource $DbPath -Query "UPDATE equivalences_groupes SET aliment_id = @A, quantite = @Q WHERE id = @Id" -SqlParameters @{ Id = $Id; A = $AlimentId; Q = $Quantite }
}

function Remove-EquivalenceGroupe {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id)
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM equivalences_aliments WHERE groupe_id = @Id; DELETE FROM equivalences_groupes WHERE id = @Id;" -SqlParameters @{ Id = $Id }
}

function Move-EquivalenceGroupe {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id, [Parameter(Mandatory)] [ValidateSet(-1, 1)] [int] $Direction)
    $ids = @(Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT id FROM equivalences_groupes ORDER BY ordre, id" | ForEach-Object { [int]$_.id })
    $index = [array]::IndexOf($ids, $Id); $cible = $index + $Direction
    if ($index -lt 0 -or $cible -lt 0 -or $cible -ge $ids.Count) { return }
    $ids[$index] = $ids[$cible]; $ids[$cible] = $Id
    Set-OrdreLignes -DbPath $DbPath -Table 'equivalences_groupes' -Ids $ids
}

function Add-EquivalenceAliment {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $GroupeId, [Parameter(Mandatory)] [int] $AlimentId)
    $existe = Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT id FROM equivalences_aliments WHERE groupe_id = @G AND aliment_id = @A" -SqlParameters @{ G = $GroupeId; A = $AlimentId }
    if ($existe) { return }
    $ordreMax = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COALESCE(MAX(ordre), -1) AS m FROM equivalences_aliments WHERE groupe_id = @G" -SqlParameters @{ G = $GroupeId }).m
    Invoke-SqliteQuery -DataSource $DbPath -Query "INSERT INTO equivalences_aliments (groupe_id, aliment_id, ordre) VALUES (@G, @A, @O)" -SqlParameters @{ G = $GroupeId; A = $AlimentId; O = ($ordreMax + 1) }
}

function Remove-EquivalenceAliment {
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $Id)
    Invoke-SqliteQuery -DataSource $DbPath -Query "DELETE FROM equivalences_aliments WHERE id = @Id" -SqlParameters @{ Id = $Id }
}

function Get-SuggestionsEquivalence {
    <#
        Propose des aliments de la bibliotheque equivalents a la reference d'un groupe : meme base
        (calories / proteines / glucides / lipides) et repartition des macros la plus proche.
    #>
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $GroupeId, [int] $Nombre = 5)
    $g = Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM equivalences_groupes WHERE id = @Id" -SqlParameters @{ Id = $GroupeId }
    if (-not $g) { return @() }
    $aliments = @(Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT * FROM aliments")
    $ref = $aliments | Where-Object { [int]$_.id -eq [int]$g.aliment_id } | Select-Object -First 1
    if (-not $ref) { return @() }
    $dejaLa = @(Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT aliment_id FROM equivalences_aliments WHERE groupe_id = @G" -SqlParameters @{ G = $GroupeId } | ForEach-Object { [int]$_.aliment_id })
    $base = Get-BaseEquivalence -Aliment $ref
    $profil = {
        param($a)
        $p = 4 * (Get-NombreOuZero $a.proteines); $gl = 4 * (Get-NombreOuZero $a.glucides); $l = 9 * (Get-NombreOuZero $a.lipides)
        $tot = $p + $gl + $l; if ($tot -le 0) { $tot = 1 }
        @(($p / $tot), ($gl / $tot), ($l / $tot))
    }
    $pr = & $profil $ref
    $candidats = foreach ($a in $aliments) {
        if ([int]$a.id -eq [int]$ref.id -or $dejaLa -contains [int]$a.id) { continue }
        if ((Get-BaseEquivalence -Aliment $a) -ne $base) { continue }
        if ($null -eq (Get-QuantiteEquivalente -Reference $ref -QuantiteReference ([double]$g.quantite) -Aliment $a -Base $base)) { continue }
        $pa = & $profil $a
        $distance = [math]::Sqrt([math]::Pow($pa[0] - $pr[0], 2) + [math]::Pow($pa[1] - $pr[1], 2) + [math]::Pow($pa[2] - $pr[2], 2))
        [pscustomobject]@{ Aliment = $a; Distance = $distance }
    }
    @($candidats | Sort-Object Distance | Select-Object -First $Nombre | ForEach-Object { $_.Aliment })
}

function Initialize-EquivalencesParDefaut {
    <#
        Une seule fois (reglage "equivalences_initialisees") et seulement si le tableau est vide : reprend
        le bloc EQUIVALENCES de l'onglet NUTRITION d'origine du coach, en retrouvant les aliments dans la
        bibliotheque par leur nom (un aliment introuvable est simplement ignore).
    #>
    param([Parameter(Mandatory)] [string] $DbPath)
    if ((Get-Parametre -DbPath $DbPath -Cle 'equivalences_initialisees') -eq '1') { return }
    $nb = [int](Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COUNT(*) AS n FROM equivalences_groupes").n
    if ($nb -eq 0) {
        $tableau = @(
            @{ Ref = 'Brocolis'; Qte = 100; Eq = @('Haricots verts', 'Epinards%', 'Chou-fleur', 'Courgettes') }
            @{ Ref = 'Patate douce bouillie'; Qte = 100; Eq = @('Pomme de terre bouillie') }
            @{ Ref = 'Riz basmati%'; Qte = 100; Eq = @('P_tes%', 'Quinoa%', 'Pomme de terre bouillie', 'Semoule%') }
            @{ Ref = 'Blanc de poulet%'; Qte = 100; Eq = @('Colin%', '%cabillaud%', '%merlan%') }
            @{ Ref = 'Steak hach_ 5%'; Qte = 100; Eq = @('Thon rouge') }
            @{ Ref = 'Skyr'; Qte = 200; Eq = @('Prot_ine isolate') }
            @{ Ref = 'Pomme "%'; Qte = 150; Eq = @('Banane', 'Kiwi vert') }
        )
        $trouver = { param([string]$Motif) (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT id FROM aliments WHERE nom LIKE @M ORDER BY id LIMIT 1" -SqlParameters @{ M = $Motif }) }
        foreach ($g in $tableau) {
            $ref = & $trouver $g.Ref
            if (-not $ref) { continue }
            $groupeId = New-EquivalenceGroupe -DbPath $DbPath -AlimentId ([int]$ref.id) -Quantite $g.Qte
            foreach ($motif in $g.Eq) {
                $a = & $trouver $motif
                if ($a) { Add-EquivalenceAliment -DbPath $DbPath -GroupeId ([int]$groupeId) -AlimentId ([int]$a.id) }
            }
        }
    }
    Set-Parametre -DbPath $DbPath -Cle 'equivalences_initialisees' -Valeur '1'
}

Export-ModuleMember -Function Initialize-EquivalencesParDefaut, Get-PlansNutrition, New-PlanNutrition, Remove-PlanNutrition, `
    Get-TypesJour, New-TypeJour, Remove-TypeJour, `
    Get-Repas, New-Repas, Remove-Repas, Move-Repas, `
    Get-RepasAliments, New-RepasAliment, Update-RepasAliment, Remove-RepasAliment, Move-RepasAliment, Get-TotauxRepas, Get-TotauxTypeJour, `
    Get-Recettes, New-Recette, Update-Recette, Remove-Recette, `
    Get-RecetteIngredients, New-RecetteIngredient, Update-RecetteIngredient, Remove-RecetteIngredient, Move-RecetteIngredient, `
    Add-RecetteAuRepas, Remove-RecetteDuRepas, `
    Get-EquivalencesGroupes, New-EquivalenceGroupe, Update-EquivalenceGroupe, Remove-EquivalenceGroupe, Move-EquivalenceGroupe, `
    Add-EquivalenceAliment, Remove-EquivalenceAliment, Get-SuggestionsEquivalence, Format-QuantiteAliment
