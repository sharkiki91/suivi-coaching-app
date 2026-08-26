Set-StrictMode -Version Latest

function ConvertTo-DoubleTolerant {
    <#
        Convertit une valeur de cellule Excel en double, en tolerant les fautes de
        saisie courantes (virgule ou point-virgule utilises comme separateur decimal).
        Retourne $null si la valeur est vide ou non convertible.
    #>
    param($Valeur)

    if ($null -eq $Valeur) { return $null }
    $texte = [string]$Valeur
    if ([string]::IsNullOrWhiteSpace($texte)) { return $null }

    $texte = $texte.Replace(';', '.').Replace(',', '.')
    $parsed = 0.0
    if ([double]::TryParse($texte, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Import-BibliothequesDepuisExcel {
    <#
        Importe les bibliotheques (exercices, aliments, complements) depuis le fichier
        Excel historique du coach ("SUIVI 2.0.xlsx" ou un fichier structure de la meme
        maniere : onglets DATABASETRAINING, DATABASECOMPLEMENTS, NUTRITION).
        Les entrees dont le nom existe deja dans la bibliotheque sont ignorees (pas de doublon).
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [string] $ExcelPath
    )

    $resultat = [ordered]@{
        ExercicesImportes = 0
        AlimentsImportes = 0
        ComplementsImportes = 0
        Erreurs = New-Object System.Collections.Generic.List[string]
    }

    try {
        $exercicesExistants = @(Get-Exercices -DbPath $DbPath | ForEach-Object { $_.nom.ToUpperInvariant() })
        $lignes = Import-Excel -Path $ExcelPath -WorksheetName 'DATABASETRAINING' -StartRow 3 -EndRow 125 -StartColumn 2 -EndColumn 4
        foreach ($ligne in $lignes) {
            $nom = [string]$ligne.EXERCICE
            if ([string]::IsNullOrWhiteSpace($nom)) { continue }
            if ($exercicesExistants -contains $nom.ToUpperInvariant()) { continue }
            New-Exercice -DbPath $DbPath -Nom $nom -MuscleCible $ligne.'MUSCLE CIBLE' -LienVideo $ligne.LIEN | Out-Null
            $resultat.ExercicesImportes++
        }
    } catch {
        $resultat.Erreurs.Add("Exercices (onglet DATABASETRAINING) : $($_.Exception.Message)")
    }

    try {
        $alimentsExistants = @(Get-Aliments -DbPath $DbPath | ForEach-Object { $_.nom.ToUpperInvariant() })
        $lignes = Import-Excel -Path $ExcelPath -WorksheetName 'NUTRITION' -StartRow 2 -EndRow 96 -StartColumn 21 -EndColumn 27
        foreach ($ligne in $lignes) {
            $nom = [string]$ligne.Aliment
            if ([string]::IsNullOrWhiteSpace($nom)) { continue }
            if ($alimentsExistants -contains $nom.ToUpperInvariant()) { continue }
            try {
                $quantite = ConvertTo-DoubleTolerant $ligne.'Quantité'
                if (-not $quantite) { $quantite = 100 }
                New-Aliment -DbPath $DbPath -Nom $nom -QuantiteReference $quantite `
                    -Kcal (ConvertTo-DoubleTolerant $ligne.Kcal) -Proteines (ConvertTo-DoubleTolerant $ligne.'Protéines') `
                    -Glucides (ConvertTo-DoubleTolerant $ligne.Glucides) -Lipides (ConvertTo-DoubleTolerant $ligne.Lipides) `
                    -Fibres (ConvertTo-DoubleTolerant $ligne.Fibres) | Out-Null
                $resultat.AlimentsImportes++
            } catch {
                $resultat.Erreurs.Add("Aliment '$nom' ignore : $($_.Exception.Message)")
            }
        }
    } catch {
        $resultat.Erreurs.Add("Aliments (onglet NUTRITION) : $($_.Exception.Message)")
    }

    try {
        $complementsExistants = @(Get-Complements -DbPath $DbPath | ForEach-Object { $_.nom.ToUpperInvariant() })
        $lignes = Import-Excel -Path $ExcelPath -WorksheetName 'DATABASECOMPLEMENTS' -StartRow 3 -EndRow 22 -StartColumn 2 -EndColumn 4
        foreach ($ligne in $lignes) {
            $nom = [string]$ligne.COMPLEMENT
            if ([string]::IsNullOrWhiteSpace($nom)) { continue }
            if ($complementsExistants -contains $nom.ToUpperInvariant()) { continue }
            New-Complement -DbPath $DbPath -Nom $nom -Dose $ligne.DOSE -Lien $ligne.LIEN | Out-Null
            $resultat.ComplementsImportes++
        }
    } catch {
        $resultat.Erreurs.Add("Complements (onglet DATABASECOMPLEMENTS) : $($_.Exception.Message)")
    }

    return [pscustomobject]$resultat
}

function Export-DonneesVersExcel {
    param(
        [Parameter(Mandatory)] $Donnees,
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $WorksheetName
    )

    if (Test-Path $Path) { Remove-Item $Path -Force }
    $Donnees | Export-Excel -Path $Path -WorksheetName $WorksheetName -AutoSize -TableStyle Medium2 -FreezeTopRow
}

function ConvertTo-DateIso {
    <# Convertit une date Excel (objet DateTime ou texte JJ/MM/AAAA, AAAA-MM-JJ...) en texte AAAA-MM-JJ. Retourne $null si non reconnue. #>
    param($Valeur)

    if ($null -eq $Valeur) { return $null }
    if ($Valeur -is [datetime]) { return $Valeur.ToString('yyyy-MM-dd') }
    $texte = [string]$Valeur
    if ([string]::IsNullOrWhiteSpace($texte)) { return $null }

    $formats = @('dd/MM/yyyy', 'yyyy-MM-dd', 'd/M/yyyy', 'dd-MM-yyyy', 'dd/MM/yyyy HH:mm:ss')
    $dt = [datetime]::MinValue
    foreach ($fmt in $formats) {
        if ([datetime]::TryParseExact($texte, $fmt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
            return $dt.ToString('yyyy-MM-dd')
        }
    }
    if ([datetime]::TryParse($texte, [System.Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return $dt.ToString('yyyy-MM-dd')
    }
    return $null
}

function Get-TexteNormalise {
    <# Minuscules, sans accents, espaces normalises : pour comparer des noms de facon tolerante. #>
    param([string] $Texte)
    if ([string]::IsNullOrWhiteSpace($Texte)) { return '' }
    $t = $Texte.Trim().ToLowerInvariant().Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $t.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return ($sb.ToString() -replace '\s+', ' ').Trim()
}

function Find-ClientParNomOuEmail {
    param($Clients, [string] $Nom, [string] $Email)
    if ($Email) {
        $emailNorm = $Email.Trim().ToLowerInvariant()
        $match = $Clients | Where-Object { $_.email -and $_.email.Trim().ToLowerInvariant() -eq $emailNorm } | Select-Object -First 1
        if ($match) { return [int]$match.id }
    }
    if ($Nom) {
        $nomNorm = Get-TexteNormalise $Nom
        if ($nomNorm) {
            foreach ($c in $Clients) {
                $nomClientNorm = Get-TexteNormalise $c.nom
                $prenomClientNorm = Get-TexteNormalise $c.prenom
                if ($nomClientNorm -and $prenomClientNorm -and $nomNorm.Contains($nomClientNorm) -and $nomNorm.Contains($prenomClientNorm)) {
                    return [int]$c.id
                }
            }
        }
    }
    return $null
}

function Get-EnTetesExcel {
    <# Renvoie la liste des en-tetes (ligne 1) d'un fichier Excel, meme s'il n'y a aucune ligne de donnees. #>
    param([Parameter(Mandatory)] [string] $ExcelPath)
    $premiereLigne = Import-Excel -Path $ExcelPath -NoHeader | Select-Object -First 1
    if (-not $premiereLigne) { return @() }
    return @($premiereLigne.PSObject.Properties.Value | Where-Object { $_ })
}

function Import-QuestionnaireDepuisExcel {
    <#
        Importe les reponses d'un questionnaire Google Forms exporte en Excel/CSV.
        Chaque colonne du fichier est conservee telle quelle dans donnees_json.
        Le rattachement au client se fait par email (colonne ColonneEmail, si fournie)
        puis par nom (colonne ColonneNom, si fournie) ; sans correspondance, la reponse
        reste non rattachee et pourra etre assignee manuellement dans l'application.
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [ValidateSet('pre_coaching', 'bilan')] [string] $Type,
        [Parameter(Mandatory)] [string] $ExcelPath,
        [string] $ColonneDate,
        [string] $ColonneNom,
        [string] $ColonneEmail
    )

    $resultat = [ordered]@{ Importees = 0; Rattachees = 0; NonRattachees = 0; Erreurs = New-Object System.Collections.Generic.List[string] }
    $clients = @(Get-Clients -DbPath $DbPath -InclureArchives)
    $lignes = @(Import-Excel -Path $ExcelPath)
    $nomFichier = Split-Path $ExcelPath -Leaf

    foreach ($ligne in $lignes) {
        try {
            $dateIso = $null
            if ($ColonneDate) { $dateIso = ConvertTo-DateIso $ligne.$ColonneDate }
            if (-not $dateIso) { $dateIso = (Get-Date).ToString('yyyy-MM-dd') }

            $nomValeur = if ($ColonneNom) { [string]$ligne.$ColonneNom } else { $null }
            $emailValeur = if ($ColonneEmail) { [string]$ligne.$ColonneEmail } else { $null }
            $clientId = Find-ClientParNomOuEmail -Clients $clients -Nom $nomValeur -Email $emailValeur

            $dict = [ordered]@{}
            foreach ($prop in $ligne.PSObject.Properties) { $dict[$prop.Name] = $prop.Value }
            $json = $dict | ConvertTo-Json -Compress

            New-QuestionnaireReponse -DbPath $DbPath -ClientId $clientId -Type $Type -DateReponse $dateIso -DonneesJson $json -FichierSource $nomFichier | Out-Null
            $resultat.Importees++
            if ($clientId) { $resultat.Rattachees++ } else { $resultat.NonRattachees++ }
        } catch {
            $resultat.Erreurs.Add($_.Exception.Message)
        }
    }
    return [pscustomobject]$resultat
}

function Export-ModeleTrackingExcel {
    <# Genere un fichier Excel vierge (avec une ligne d'exemple) au format impose pour le suivi quotidien du client. #>
    param([Parameter(Mandatory)] [string] $Path)

    $exemple = [pscustomobject]@{
        'Date' = '01/09/2026'
        'Poids (kg)' = 82.5
        'Sommeil (h)' = 7.5
        'Qualite sommeil (1-5)' = 4
        'Heure coucher' = '23:00'
        'Heure lever' = '07:00'
        'Energie (1-5)' = 3
        'Adhesion nutrition (1-5)' = 4
        'Digestion (1-5)' = 4
        'Nb pas' = 8500
        'Cardio (min)' = 20
        'Motivation (1-5)' = 4
        'Tension systolique' = $null
        'Tension diastolique' = $null
        'Bilan' = 'Exemple de ligne a remplacer - une ligne par jour'
    }
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $exemple | Export-Excel -Path $Path -WorksheetName 'Suivi' -AutoSize -TableStyle Medium2 -FreezeTopRow
}

function Import-TrackingDepuisExcel {
    <# Importe le suivi quotidien d'un client depuis un fichier au format du modele (Export-ModeleTrackingExcel). #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $ExcelPath
    )

    $resultat = [ordered]@{ Importes = 0; IgnoresSansDate = 0; Erreurs = New-Object System.Collections.Generic.List[string] }
    $lignes = @(Import-Excel -Path $ExcelPath)

    foreach ($ligne in $lignes) {
        $dateIso = ConvertTo-DateIso $ligne.'Date'
        if (-not $dateIso) { $resultat.IgnoresSansDate++; continue }
        try {
            Set-SuiviQuotidienJour -DbPath $DbPath -ClientId $ClientId -Date $dateIso `
                -Poids (ConvertTo-DoubleTolerant $ligne.'Poids (kg)') `
                -SommeilHeures (ConvertTo-DoubleTolerant $ligne.'Sommeil (h)') `
                -QualiteSommeil (ConvertTo-DoubleTolerant $ligne.'Qualite sommeil (1-5)') `
                -HeureCoucher ([string]$ligne.'Heure coucher') -HeureLever ([string]$ligne.'Heure lever') `
                -Energie (ConvertTo-DoubleTolerant $ligne.'Energie (1-5)') `
                -AdhesionNutrition (ConvertTo-DoubleTolerant $ligne.'Adhesion nutrition (1-5)') `
                -Digestion (ConvertTo-DoubleTolerant $ligne.'Digestion (1-5)') `
                -NbPas (ConvertTo-DoubleTolerant $ligne.'Nb pas') -CardioMinutes (ConvertTo-DoubleTolerant $ligne.'Cardio (min)') `
                -Motivation (ConvertTo-DoubleTolerant $ligne.'Motivation (1-5)') `
                -TensionSystolique (ConvertTo-DoubleTolerant $ligne.'Tension systolique') -TensionDiastolique (ConvertTo-DoubleTolerant $ligne.'Tension diastolique') `
                -Bilan ([string]$ligne.'Bilan')
            $resultat.Importes++
        } catch {
            $resultat.Erreurs.Add("Ligne du $dateIso : $($_.Exception.Message)")
        }
    }
    return [pscustomobject]$resultat
}

Export-ModuleMember -Function Import-BibliothequesDepuisExcel, Export-DonneesVersExcel, Get-EnTetesExcel, `
    Import-QuestionnaireDepuisExcel, Export-ModeleTrackingExcel, Import-TrackingDepuisExcel
