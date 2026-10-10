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

    $texte = ($texte -replace '\s', '').Replace(';', '.').Replace(',', '.')
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
    # Cellule date lue directement (EPPlus) : Excel stocke les dates en nombre de jours depuis 1900
    if ($Valeur -is [double] -or $Valeur -is [int]) {
        $n = [double]$Valeur
        if ($n -ge 20000 -and $n -le 80000) { return [datetime]::FromOADate($n).ToString('yyyy-MM-dd') }
        return $null
    }
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

function Get-ValeurColonne {
    <#
        Lit la colonne $Nom d'une ligne importee, ou $null si le fichier n'a pas cette colonne
        (sous Set-StrictMode, $ligne.'Colonne absente' leve une erreur au lieu de renvoyer $null).
    #>
    param($Ligne, [string] $Nom)
    $prop = $Ligne.PSObject.Properties[$Nom]
    if ($prop) { return $prop.Value }
    return $null
}

function Test-ColonnesRequises {
    <# Leve une erreur lisible si le fichier ne contient pas les colonnes attendues (mauvais fichier choisi). #>
    param($Lignes, [string[]] $Colonnes, [string] $DescriptionFichier)
    if (@($Lignes).Count -eq 0) { return }
    $presentes = @($Lignes)[0].PSObject.Properties.Name
    $manquantes = @($Colonnes | Where-Object { $presentes -notcontains $_ })
    if ($manquantes.Count -gt 0) {
        throw "Ce fichier ne ressemble pas a $DescriptionFichier (colonne(s) manquante(s) : $($manquantes -join ', '))."
    }
}

function Get-SeparateurCsv {
    <# Devine le separateur d'un CSV (virgule pour Google Forms/Sheets, point-virgule pour Excel en francais). #>
    param([Parameter(Mandatory)] [string] $CsvPath)
    $premiere = Get-Content -Path $CsvPath -TotalCount 1 -Encoding UTF8
    if ($null -eq $premiere) { return ',' }
    if (([regex]::Matches($premiere, ';')).Count -gt ([regex]::Matches($premiere, ',')).Count) { return ';' }
    return ','
}

function Test-EstCsv { param([string] $Path) return ([System.IO.Path]::GetExtension($Path) -eq '.csv') }

function Get-EnTetesExcel {
    <# Renvoie la liste des en-tetes (ligne 1) d'un fichier Excel ou CSV, meme s'il n'y a aucune ligne de donnees. #>
    param([Parameter(Mandatory)] [string] $ExcelPath)
    if (Test-EstCsv $ExcelPath) {
        Add-Type -AssemblyName Microsoft.VisualBasic
        $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($ExcelPath, [System.Text.Encoding]::UTF8)
        try {
            $parser.SetDelimiters((Get-SeparateurCsv $ExcelPath))
            $parser.HasFieldsEnclosedInQuotes = $true
            if ($parser.EndOfData) { return @() }
            return @($parser.ReadFields() | Where-Object { $_ })
        } finally { $parser.Close() }
    }
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

        Pour le questionnaire pre-coaching, si ColonneTelephone/ColonneObjectif* sont
        fournies, les champs Telephone/Email/Objectifs de la fiche client sont completes
        automatiquement -- mais uniquement s'ils sont actuellement vides (jamais d'ecrasement).
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [ValidateSet('pre_coaching', 'bilan')] [string] $Type,
        [Parameter(Mandatory)] [string] $ExcelPath,
        [string] $ColonneDate,
        [string] $ColonneNom,
        [string] $ColonneEmail,
        [string] $ColonneTelephone,
        [string] $ColonneObjectifLongTerme,
        [string] $ColonneObjectifCourtTerme,
        [string] $ColonneMoyens
    )

    $resultat = [ordered]@{ Importees = 0; Rattachees = 0; NonRattachees = 0; FichesCompletees = 0; Erreurs = New-Object System.Collections.Generic.List[string] }
    $clients = @(Get-Clients -DbPath $DbPath -InclureArchives)
    # -AsText lit les cellules telles qu'affichees dans Excel : evite qu'un numero de telephone (souvent commencant
    # par 0) soit lu comme un nombre et perde son 0 initial. Applique a toutes les colonnes (le telephone peut se
    # trouver dans une colonne non indiquee dans l'assistant, ex. questionnaire bilan), sauf la date qui doit rester une date.
    $colonnesTexte = @(Get-EnTetesExcel -ExcelPath $ExcelPath | Where-Object { $_ -ne $ColonneDate })
    $lignes = if (Test-EstCsv $ExcelPath) {
        # Export CSV de Google Forms/Sheets : tout est deja du texte (le 0 initial des telephones est conserve).
        @(Import-Csv -Path $ExcelPath -Delimiter (Get-SeparateurCsv $ExcelPath) -Encoding UTF8)
    } elseif ($colonnesTexte.Count -gt 0) { @(Import-Excel -Path $ExcelPath -AsText $colonnesTexte) } else { @(Import-Excel -Path $ExcelPath) }
    $nomFichier = Split-Path $ExcelPath -Leaf

    foreach ($ligne in $lignes) {
        try {
            $dateIso = $null
            if ($ColonneDate) { $dateIso = ConvertTo-DateIso (Get-ValeurColonne $ligne $ColonneDate) }
            if (-not $dateIso) { $dateIso = (Get-Date).ToString('yyyy-MM-dd') }

            $nomValeur = if ($ColonneNom) { [string](Get-ValeurColonne $ligne $ColonneNom) } else { $null }
            $emailValeur = if ($ColonneEmail) { [string](Get-ValeurColonne $ligne $ColonneEmail) } else { $null }
            $clientId = Find-ClientParNomOuEmail -Clients $clients -Nom $nomValeur -Email $emailValeur

            $dict = [ordered]@{}
            foreach ($prop in $ligne.PSObject.Properties) {
                # ConvertTo-Json (PS 5.1) ecrirait une date Excel sous la forme "\/Date(1790770917240)\/" : on la stocke lisible.
                $dict[$prop.Name] = if ($prop.Value -is [datetime]) { $prop.Value.ToString('dd/MM/yyyy HH:mm') } else { $prop.Value }
            }
            $json = $dict | ConvertTo-Json -Compress

            New-QuestionnaireReponse -DbPath $DbPath -ClientId $clientId -Type $Type -DateReponse $dateIso -DonneesJson $json -FichierSource $nomFichier | Out-Null
            $resultat.Importees++
            if ($clientId) { $resultat.Rattachees++ } else { $resultat.NonRattachees++ }

            if ($clientId -and $Type -eq 'pre_coaching') {
                $telephoneValeur = if ($ColonneTelephone) { [string](Get-ValeurColonne $ligne $ColonneTelephone) } else { $null }
                $morceauxObjectifs = @()
                foreach ($col in @($ColonneObjectifLongTerme, $ColonneObjectifCourtTerme, $ColonneMoyens)) {
                    if ($col -and -not [string]::IsNullOrWhiteSpace([string](Get-ValeurColonne $ligne $col))) { $morceauxObjectifs += [string](Get-ValeurColonne $ligne $col) }
                }
                $objectifsValeur = if ($morceauxObjectifs.Count -gt 0) { $morceauxObjectifs -join "`r`n`r`n" } else { $null }

                if ($telephoneValeur -or $emailValeur -or $objectifsValeur) {
                    $complete = Set-ClientChampsSiVide -DbPath $DbPath -Id $clientId -Telephone $telephoneValeur -Email $emailValeur -Objectifs $objectifsValeur
                    if ($complete) { $resultat.FichesCompletees++ }
                }
            }
        } catch {
            $resultat.Erreurs.Add($_.Exception.Message)
        }
    }
    return [pscustomobject]$resultat
}

function Export-ModeleTrackingExcel {
    <#
        Genere le modele de suivi quotidien a envoyer au client : l'onglet TRACKING (52 semaines, dates
        deja remplies a partir de -DateDebut, lundi de cette semaine par defaut), avec
        les colonnes et la case BILAN choisies par le coach (-DbPath). C'est le meme onglet que celui
        inclus dans la feuille de seance (Programmes > Exporter la feuille de seance).
    #>
    param(
        [Parameter(Mandatory)] [string] $Path,
        [string] $DbPath,
        [datetime] $DateDebut = (Get-LundiCetteSemaine)
    )
    $reglages = if ($DbPath) { Get-ReglagesTracking -DbPath $DbPath } else { $null }
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $pkg = Open-ExcelPackage -Path $Path -Create
    try {
        if ($reglages) {
            Add-OngletTrackingExcel -Pkg $pkg -DateDebut $DateDebut.Date -Colonnes $reglages.Colonnes -AvecBilan $reglages.AvecBilan -LienBilan $reglages.LienBilan -NbSemaines $reglages.NbSemaines
        } else {
            Add-OngletTrackingExcel -Pkg $pkg -DateDebut $DateDebut.Date
        }
    } finally {
        Close-ExcelPackage $pkg
    }
}

function Get-CleTrackingDepuisEntete {
    <# Cle du catalogue TRACKING correspondant a un en-tete de colonne (libelle actuel ou ancien nom), 'date', ou $null. #>
    param([string] $Entete)
    $n = Get-TexteNormalise $Entete
    if (-not $n) { return $null }
    if ($n -eq 'date') { return 'date' }
    foreach ($def in Get-CatalogueTracking) {
        if ((Get-TexteNormalise $def.Libelle) -eq $n) { return $def.Cle }
        foreach ($ancien in $def.Anciens) { if ((Get-TexteNormalise $ancien) -eq $n) { return $def.Cle } }
    }
    return $null
}

function ConvertTo-HeureImport {
    <# "23:00" reste "23:00" ; une heure lue comme fraction de jour Excel (0.958...) redevient "23:00". #>
    param($Valeur)
    if ($Valeur -is [double] -and $Valeur -ge 0 -and $Valeur -lt 1) { return [TimeSpan]::FromDays($Valeur).ToString('hh\:mm') }
    if ($Valeur -is [datetime]) { return $Valeur.ToString('HH:mm') }
    return Get-TexteImportOuNull $Valeur
}

function Import-LigneTracking {
    <#
        Enregistre un jour de suivi a partir de $Valeurs (hashtable cle du catalogue TRACKING -> valeur
        lue, plus 'date'). Seules les colonnes presentes dans le fichier sont renseignees, les autres
        restent vides. Un jour sans aucune valeur saisie (date seule, deja pre-remplie dans le modele)
        est ignore sans bruit ; un jour rempli sans date est compte dans IgnoresSansDate.
    #>
    param([string] $DbPath, [int] $ClientId, [hashtable] $Valeurs, $Resultat)

    if ([string]$Valeurs['bilan'] -like 'Exemple de ligne a remplacer*') { return }
    $parametres = @{}
    foreach ($def in Get-CatalogueTracking) {
        if (-not $Valeurs.ContainsKey($def.Cle)) { continue }
        $brut = $Valeurs[$def.Cle]
        $valeur = switch ($def.Type) {
            'heure'  { ConvertTo-HeureImport $brut }
            'texte'  { Get-TexteImportOuNull $brut }
            'ouinon' { if ([string]$brut -match '^\s*(x|oui|o|1|vrai|true)\s*$') { $true } else { $null } }
            default  { ConvertTo-DoubleTolerant $brut }
        }
        if ($null -ne $valeur) { $parametres[$def.Param] = $valeur }
    }
    if ($parametres.Count -eq 0) { return }
    $dateIso = ConvertTo-DateIso $Valeurs['date']
    if (-not $dateIso) { $Resultat.IgnoresSansDate++; return }
    try {
        Set-SuiviQuotidienJour -DbPath $DbPath -ClientId $ClientId -Date $dateIso @parametres
        $Resultat.Importes++
    } catch {
        $Resultat.Erreurs.Add("Suivi du $dateIso : $($_.Exception.Message)")
    }
}

function Import-TrackingOngletCoach {
    <#
        Lit un onglet TRACKING genere par Add-OngletTrackingExcel : chaque ligne d'en-tetes est reperee
        par "T|" en colonne A (masquee), chaque jour par "J|". Chaque valeur est retrouvee par son
        en-tete, donc quelles que soient les colonnes choisies a l'export (bandes de theme, JOUR, BILAN
        et MOYENNE sont ignores).
    #>
    param([string] $DbPath, [int] $ClientId, $Ws, $Resultat)

    $finLigne = $Ws.Dimension.End.Row; $finCol = $Ws.Dimension.End.Column
    $colonnes = @{}   # numero de colonne -> cle du catalogue
    for ($r = 1; $r -le $finLigne; $r++) {
        $repere = [string]$Ws.Cells[$r, 1].Value
        if ($repere -like 'T|*') {
            $colonnes = @{}
            for ($c = 2; $c -le $finCol; $c++) {
                $cle = Get-CleTrackingDepuisEntete ([string]$Ws.Cells[$r, $c].Text)
                if ($cle) { $colonnes[$c] = $cle }
            }
            continue
        }
        if ($repere -notlike 'J|*' -or $colonnes.Count -eq 0) { continue }
        $valeurs = @{}
        foreach ($c in $colonnes.Keys) { $valeurs[$colonnes[$c]] = $Ws.Cells[$r, $c].Value }
        Import-LigneTracking -DbPath $DbPath -ClientId $ClientId -Valeurs $valeurs -Resultat $Resultat
    }
}

function Import-TrackingDepuisExcel {
    <#
        Importe le suivi quotidien d'un client : onglet TRACKING (modele actuel ou feuille de seance),
        ou ancien modele en tableau (une ligne par jour, colonnes 'Date', 'Poids (kg)'...).
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $ExcelPath
    )

    $resultat = [ordered]@{ Importes = 0; IgnoresSansDate = 0; Erreurs = New-Object System.Collections.Generic.List[string] }
    $ongletTrouve = $false
    $pkg = Open-ExcelPackage -Path $ExcelPath
    try {
        foreach ($ws in $pkg.Workbook.Worksheets) {
            if (-not $ws.Dimension) { continue }
            if (Test-OngletAvecRepere -Ws $ws -Repere 'T|*') {
                Import-TrackingOngletCoach -DbPath $DbPath -ClientId $ClientId -Ws $ws -Resultat $resultat
                $ongletTrouve = $true
            }
        }
    } finally {
        Close-ExcelPackage $pkg -NoSave
    }
    if ($ongletTrouve) { return [pscustomobject]$resultat }

    $lignes = @(Import-Excel -Path $ExcelPath)
    Test-ColonnesRequises -Lignes $lignes -Colonnes @('Date') -DescriptionFichier 'un suivi quotidien (modele a telecharger depuis l''application)'
    foreach ($ligne in $lignes) {
        $valeurs = @{}
        foreach ($prop in $ligne.PSObject.Properties) {
            $cle = Get-CleTrackingDepuisEntete $prop.Name
            if ($cle) { $valeurs[$cle] = $prop.Value }
        }
        Import-LigneTracking -DbPath $DbPath -ClientId $ClientId -Valeurs $valeurs -Resultat $resultat
    }
    return [pscustomobject]$resultat
}

function Test-OngletAvecRepere {
    <# Vrai si la colonne A (masquee) de l'onglet contient le repere technique donne dans ses 30 premieres lignes. #>
    param($Ws, [string] $Repere)
    for ($r = 1; $r -le [math]::Min($Ws.Dimension.End.Row, 30); $r++) { if ([string]$Ws.Cells[$r, 1].Value -like $Repere) { return $true } }
    return $false
}

function Get-LignesRoadmapFormatCoach {
    <#
        Lit un onglet ROADMAP au format du coach (modele de l'application, ou son fichier d'origine) :
        en-tete "SEM" en B3, une semaine par ligne a partir de la ligne 5. Les colonnes sont reconnues par
        leur titre (lignes 3-4) car certains fichiers clients n'ont pas toutes les colonnes (ex. sans POIDS
        MOYEN / CARDIO / PAS). Retourne $null si le fichier n'a pas ce format. Chaque ligne = hashtable
        cle roadmap_semaines -> valeur brute de la cellule (seulement les colonnes presentes).
    #>
    param([Parameter(Mandatory)] [string] $ExcelPath)
    $pkg = Open-ExcelPackage -Path $ExcelPath
    try {
        $ws = $null
        foreach ($w in $pkg.Workbook.Worksheets) {
            if ($w.Dimension -and ([string]$w.Cells[3, 2].Value).Trim() -eq 'SEM') { $ws = $w; break }
        }
        if (-not $ws) { return $null }
        $positions = @{}   # cle -> numero de colonne
        foreach ($c in @(Get-ColonnesRoadmap)) {
            for ($col = 2; $col -le $ws.Dimension.End.Column; $col++) {
                $titres = @(([string]$ws.Cells[3, $col].Text).Trim().ToUpperInvariant(), ([string]$ws.Cells[4, $col].Text).Trim().ToUpperInvariant())
                if ($titres -contains $c.Titre) { $positions[$c.Cle] = $col; break }
            }
        }
        $lignes = New-Object System.Collections.ArrayList
        for ($r = 5; $r -le $ws.Dimension.End.Row; $r++) {
            $valeurs = @{}
            foreach ($cle in $positions.Keys) { $valeurs[$cle] = $ws.Cells[$r, $positions[$cle]].Value }
            $lignes.Add($valeurs) | Out-Null
        }
        return , $lignes
    } finally {
        Close-ExcelPackage $pkg -NoSave
    }
}

function Import-RoadmapDepuisExcel {
    <#
        Importe la roadmap hebdo d'un client : modele au format ROADMAP du coach (Export-ModeleRoadmapExcel),
        ou ancien modele "une colonne par champ" (en-tete Semaine). Une semaine deja presente (meme numero)
        est mise a jour plutot que dupliquee ; une semaine nouvelle sans aucune information (juste le numero
        et la date pre-remplis par le modele) est ignoree.
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $ExcelPath
    )

    $resultat = [ordered]@{ Importees = 0; IgnoreesSansNumero = 0; Erreurs = New-Object System.Collections.Generic.List[string] }
    $existantes = @{}
    foreach ($s in @(Get-RoadmapSemaines -DbPath $DbPath -ClientId $ClientId)) { $existantes[[int]$s.semaine_numero] = $s }

    $formatCoach = Get-LignesRoadmapFormatCoach -ExcelPath $ExcelPath
    if ($null -ne $formatCoach) {
        $lignes = @($formatCoach)
    } else {
        $brutes = @(Import-Excel -Path $ExcelPath)
        Test-ColonnesRequises -Lignes $brutes -Colonnes @('Semaine') -DescriptionFichier 'une roadmap (modele a telecharger depuis l''application)'
        $lignes = foreach ($ligne in $brutes) {
            $v = { param([string]$Nom) Get-ValeurColonne $ligne $Nom }
            if ([string](& $v 'Notes') -like 'Exemple de ligne a remplacer*') { continue }
            @{
                semaine_numero = (& $v 'Semaine'); date_debut = (& $v 'Date debut'); phase = (& $v 'Phase'); nutrition = (& $v 'Nutrition')
                poids_moyen = (& $v 'Poids moyen (kg)'); depense_calorique = (& $v 'Depense calorique'); cardio_minutes = (& $v 'Cardio (min)')
                pas = (& $v 'Pas'); precision_training = (& $v 'Precision training'); evenements = (& $v 'Evenements'); notes = (& $v 'Notes')
            }
        }
    }

    foreach ($l in @($lignes)) {
        $numeroDouble = ConvertTo-DoubleTolerant $l['semaine_numero']
        if ($null -eq $numeroDouble) {
            $remplie = @($l.Keys | Where-Object { $_ -ne 'semaine_numero' -and -not [string]::IsNullOrWhiteSpace([string]$l[$_]) }).Count -gt 0
            if ($remplie) { $resultat.IgnoreesSansNumero++ }
            continue
        }
        $numero = [int]$numeroDouble
        # Ligne sans aucune information (juste le numero et la date du modele) : ignoree, elle n'efface rien
        $remplie = @($l.Keys | Where-Object { $_ -notin @('semaine_numero', 'date_debut') -and -not [string]::IsNullOrWhiteSpace([string]$l[$_]) }).Count -gt 0
        if (-not $remplie) { continue }
        $existante = if ($existantes.ContainsKey($numero)) { $existantes[$numero] } else { $null }
        # Colonne absente du fichier (ex. "depense calorique", ou fichier client sans POIDS MOYEN) : on garde la valeur de l'appli
        $val = { param([string]$Cle) if ($l.ContainsKey($Cle)) { $l[$Cle] } elseif ($existante) { $existante.$Cle } else { $null } }
        try {
            $champs = @{
                DbPath = $DbPath
                SemaineNumero = $numero
                DateDebut = (ConvertTo-DateIso (& $val 'date_debut'))
                Phase = (Get-TexteImportOuNull (& $val 'phase'))
                Nutrition = (Get-TexteImportOuNull (& $val 'nutrition'))
                PoidsMoyen = (ConvertTo-DoubleTolerant (& $val 'poids_moyen'))
                DepenseCalorique = (ConvertTo-DoubleTolerant (& $val 'depense_calorique'))
                CardioMinutes = (ConvertTo-DoubleTolerant (& $val 'cardio_minutes'))
                Pas = (ConvertTo-DoubleTolerant (& $val 'pas'))
                PrecisionDepense = (Get-TexteImportOuNull (& $val 'precision_depense'))
                PrecisionTraining = (Get-TexteImportOuNull (& $val 'precision_training'))
                Evenements = (Get-TexteImportOuNull (& $val 'evenements'))
                Notes = (Get-TexteImportOuNull (& $val 'notes'))
            }
            if ($existante) {
                Update-RoadmapSemaine -Id ([int]$existante.id) @champs
            } else {
                New-RoadmapSemaine -ClientId $ClientId @champs | Out-Null
            }
            $resultat.Importees++
        } catch {
            $resultat.Erreurs.Add("Semaine $numero : $($_.Exception.Message)")
        }
    }
    return [pscustomobject]$resultat
}

function Get-TexteImportOuNull {
    param($Valeur)
    if ($null -eq $Valeur) { return $null }
    $texte = ([string]$Valeur).Trim()
    if ([string]::IsNullOrWhiteSpace($texte)) { return $null }
    return $texte
}

function Convert-DoubleFatSecret {
    <# Convertit un champ numerique FatSecret (virgule decimale, ou vide) en double. Retourne $null si vide/non convertible. #>
    param($Valeur)
    if ($null -eq $Valeur) { return $null }
    $texte = (([string]$Valeur) -replace '\s', '').Replace(',', '.')
    if ([string]::IsNullOrWhiteSpace($texte)) { return $null }
    $parsed = 0.0
    if ([double]::TryParse($texte, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Import-JournalAlimentaireDepuisFatSecret {
    <#
        Importe un export "Food Diary Report - Detailed Report" de FatSecret (CSV) pour un client.
        Seuls les totaux quotidiens sont importes (pas le detail par repas/aliment). Une date deja
        presente pour ce client est mise a jour plutot que dupliquee.
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $CsvPath
    )

    Add-Type -AssemblyName Microsoft.VisualBasic

    $moisFr = @{
        'janvier' = 1; 'février' = 2; 'mars' = 3; 'avril' = 4; 'mai' = 5; 'juin' = 6
        'juillet' = 7; 'août' = 8; 'septembre' = 9; 'octobre' = 10; 'novembre' = 11; 'décembre' = 12
    }
    $regexJour = '^(?:lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche),\s*(\S+)\s+(\d{1,2}),\s*(\d{4})$'

    $resultat = [ordered]@{ Importes = 0; Erreurs = New-Object System.Collections.Generic.List[string] }

    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($CsvPath, [System.Text.Encoding]::UTF8)
    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters(',')
    $parser.HasFieldsEnclosedInQuotes = $true
    try {
        while (-not $parser.EndOfData) {
            $champs = $parser.ReadFields()
            if (-not $champs -or $champs.Count -lt 11) { continue }
            $premier = $champs[0].Trim()
            if ($premier -notmatch $regexJour) { continue }

            $mois = $moisFr[$Matches[1].ToLower()]
            if (-not $mois) { $resultat.Erreurs.Add("Mois non reconnu ('$($Matches[1])') sur la ligne : $premier"); continue }
            try {
                $date = (Get-Date -Year ([int]$Matches[3]) -Month $mois -Day ([int]$Matches[2])).ToString('yyyy-MM-dd')
            } catch {
                $resultat.Erreurs.Add("Date invalide sur la ligne : $premier"); continue
            }

            try {
                Set-JournalAlimentaireJour -DbPath $DbPath -ClientId $ClientId -Date $date `
                    -Kcal (Convert-DoubleFatSecret $champs[1]) -Lipides (Convert-DoubleFatSecret $champs[2]) `
                    -LipidesSaturees (Convert-DoubleFatSecret $champs[3]) -Glucides (Convert-DoubleFatSecret $champs[4]) `
                    -Fibres (Convert-DoubleFatSecret $champs[5]) -Sucres (Convert-DoubleFatSecret $champs[6]) `
                    -Proteines (Convert-DoubleFatSecret $champs[7]) -SodiumMg (Convert-DoubleFatSecret $champs[8]) `
                    -CholesterolMg (Convert-DoubleFatSecret $champs[9]) -PotassiumMg (Convert-DoubleFatSecret $champs[10])
                $resultat.Importes++
            } catch {
                $resultat.Erreurs.Add("Jour $date : $($_.Exception.Message)")
            }
        }
    } finally {
        $parser.Close()
    }

    return [pscustomobject]$resultat
}

function Import-FeuilleSeanceFormatCoach {
    <#
        Lit la feuille de suivi generee par Export-FeuilleSeanceExcel : pour chaque seance, une ligne
        "D|seanceId" (colonne A masquee) portant les cases DATE des blocs SEANCE 1..6, puis une ligne
        "S|seanceId|seanceExerciceId|serie" par serie, avec REPS / CHARGE (et NOTES sur la 1re serie
        de l'exercice, cellule fusionnee) sous chaque bloc. Chaque bloc renseigne = une seance realisee
        a sa date ; un bloc rempli sans date est compte dans IgnoreesSansDate.
    #>
    param([string] $DbPath, [int] $ClientId, $Ws, $Resultat)

    $finLigne = $Ws.Dimension.End.Row; $finCol = $Ws.Dimension.End.Column
    $blocs = @(); $dates = @{}
    $saisies = [ordered]@{}   # "colonneBloc|seanceExerciceId" -> valeurs saisies pour cet exercice dans ce bloc
    for ($r = 1; $r -le $finLigne; $r++) {
        $repere = [string]$Ws.Cells[$r, 1].Value
        if ($repere -like 'D|*') {
            # Les blocs sont retrouves par leur case "DATE" (resiste a un decalage de colonnes)
            $blocs = @(for ($c = 2; $c -le $finCol; $c++) { if ([string]$Ws.Cells[$r, $c].Text -eq 'DATE') { $c } })
            $dates = @{}
            foreach ($c0 in $blocs) { $dates[$c0] = ConvertTo-DateIso $Ws.Cells[$r, ($c0 + 1)].Value }
            continue
        }
        if ($repere -notlike 'S|*') { continue }
        $morceaux = $repere.Split('|')
        if ($morceaux.Count -lt 4) { continue }
        $seanceId = [int]$morceaux[1]; $seId = [int]$morceaux[2]; $numSerie = [int]$morceaux[3]
        foreach ($c0 in $blocs) {
            $cle = "$c0|$seId"
            if (-not $saisies.Contains($cle)) {
                $saisies[$cle] = [pscustomobject]@{
                    SeanceId = $seanceId; SeanceExerciceId = $seId; Date = $dates[$c0]; Bloc = [array]::IndexOf($blocs, $c0) + 1
                    Exercice = [string]$Ws.Cells[$r, 4].Text
                    Reps = New-Object System.Collections.Generic.List[string]; Charges = New-Object System.Collections.Generic.List[string]; Notes = $null
                }
            }
            $saisie = $saisies[$cle]
            $saisie.Reps.Add((Get-TexteImportOuNull $Ws.Cells[$r, ($c0 + 1)].Value))
            $saisie.Charges.Add((Get-TexteImportOuNull $Ws.Cells[$r, ($c0 + 2)].Value))
            $note = Get-TexteImportOuNull $Ws.Cells[$r, ($c0 + 3)].Value
            if ($note -and -not $saisie.Notes) { $saisie.Notes = $note }
        }
    }

    foreach ($saisie in $saisies.Values) {
        $repsRenseignees = @($saisie.Reps | Where-Object { $_ })
        $chargesRenseignees = @($saisie.Charges | Where-Object { $_ })
        if ($repsRenseignees.Count -eq 0 -and $chargesRenseignees.Count -eq 0 -and -not $saisie.Notes) { continue }
        if (-not $saisie.Date) { $Resultat.IgnoreesSansDate++; continue }
        $reps = if ($repsRenseignees.Count -gt 0) { ($saisie.Reps | ForEach-Object { if ($_) { $_ } else { '-' } }) -join ' / ' } else { $null }
        $charge = if ($chargesRenseignees.Count -gt 0) { ($saisie.Charges | ForEach-Object { if ($_) { $_ } else { '-' } }) -join ' / ' } else { $null }
        try {
            $seanceRealiseeId = Get-OuCreerSeanceRealisee -DbPath $DbPath -SeanceId $saisie.SeanceId -ClientId $ClientId -DateRealisation $saisie.Date
            Set-ExerciceRealise -DbPath $DbPath -SeanceRealiseeId $seanceRealiseeId -SeanceExerciceId $saisie.SeanceExerciceId `
                -Series ([string]$saisie.Reps.Count) -Repetitions $reps -Charge $charge -Notes $saisie.Notes
            $Resultat.Importees++
        } catch {
            $Resultat.Erreurs.Add("SEANCE $($saisie.Bloc) - '$($saisie.Exercice)' du $($saisie.Date) : $($_.Exception.Message)")
        }
    }
}

function Import-SeanceRealiseeDepuisExcel {
    <#
        Importe une feuille de seance remplie par le client (generee par Export-FeuilleSeanceExcel) :
        une ligne par serie realisee, regroupees ici par exercice (SeanceExerciceId). La date de
        realisation et les valeurs recup/tempo/notes n'ont besoin d'etre renseignees que sur une seule
        ligne de serie de l'exercice ; les repetitions et charges realisees sont assemblees serie par
        serie (ex. "15 / 20 / 25", "-" pour une serie non renseignee). Un exercice sans aucune date ou
        sans aucune valeur realisee dans son groupe de lignes est ignore. Reimporter un fichier deja
        traite met a jour les lignes existantes (meme seance + meme date) plutot que de les dupliquer.
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $ExcelPath
    )

    $resultat = [ordered]@{ Importees = 0; IgnoreesSansDate = 0; Erreurs = New-Object System.Collections.Generic.List[string] }

    # Format actuel (v1.17+) : feuille "TRAINING" calquee sur le fichier du coach, reperee par la colonne A masquee.
    # Un onglet par seance (v1.20+) ; les feuilles v1.17-1.19 avaient toutes les seances dans un seul onglet.
    $formatCoachTrouve = $false
    $pkg = Open-ExcelPackage -Path $ExcelPath
    try {
        foreach ($ws in $pkg.Workbook.Worksheets) {
            if (-not $ws.Dimension) { continue }
            if (Test-OngletAvecRepere -Ws $ws -Repere 'D|*') {
                Import-FeuilleSeanceFormatCoach -DbPath $DbPath -ClientId $ClientId -Ws $ws -Resultat $resultat
                $formatCoachTrouve = $true
            }
        }
    } finally {
        Close-ExcelPackage $pkg -NoSave
    }
    if ($formatCoachTrouve) { return [pscustomobject]$resultat }

    # Ancien format (une ligne par serie, en tableau) : feuilles deja envoyees aux clients avant la v1.17.
    $lignes = @(Import-Excel -Path $ExcelPath)
    Test-ColonnesRequises -Lignes $lignes -Colonnes @('SeanceId', 'SeanceExerciceId', 'Serie', 'Date de realisation', 'Repetitions realisees', 'Charge realisee') `
        -DescriptionFichier 'une feuille de seance (generee depuis Programmes > Feuille de seance)'
    $groupes = @($lignes | Group-Object -Property SeanceExerciceId)

    foreach ($groupe in $groupes) {
        $lignesExercice = @($groupe.Group | Sort-Object { [int]$_.Serie })
        $premiere = $lignesExercice[0]

        $dateIso = $null
        foreach ($l in $lignesExercice) {
            $dateIso = ConvertTo-DateIso $l.'Date de realisation'
            if ($dateIso) { break }
        }
        if (-not $dateIso) { $resultat.IgnoreesSansDate++; continue }

        $repsValeurs = @($lignesExercice | ForEach-Object { Get-TexteImportOuNull $_.'Repetitions realisees' })
        $chargeValeurs = @($lignesExercice | ForEach-Object { Get-TexteImportOuNull $_.'Charge realisee' })
        $recup = @($lignesExercice | ForEach-Object { Get-TexteImportOuNull (Get-ValeurColonne $_ 'Recup realisee (s)') } | Where-Object { $_ }) | Select-Object -First 1
        $tempo = @($lignesExercice | ForEach-Object { Get-TexteImportOuNull (Get-ValeurColonne $_ 'Tempo realise') } | Where-Object { $_ }) | Select-Object -First 1
        $notes = @($lignesExercice | ForEach-Object { Get-TexteImportOuNull (Get-ValeurColonne $_ 'Notes') } | Where-Object { $_ }) | Select-Object -First 1

        $repsRenseignees = @($repsValeurs | Where-Object { $_ })
        $chargeRenseignees = @($chargeValeurs | Where-Object { $_ })
        $tousVides = -not ($repsRenseignees.Count -gt 0 -or $chargeRenseignees.Count -gt 0 -or $recup -or $tempo -or $notes)
        if ($tousVides) { continue }

        $reps = if ($repsRenseignees.Count -gt 0) { ($repsValeurs | ForEach-Object { if ($_) { $_ } else { '-' } }) -join ' / ' } else { $null }
        $charge = if ($chargeRenseignees.Count -gt 0) { ($chargeValeurs | ForEach-Object { if ($_) { $_ } else { '-' } }) -join ' / ' } else { $null }

        try {
            $seanceId = [int]$premiere.SeanceId
            $seanceExerciceId = [int]$premiere.SeanceExerciceId
            $seanceRealiseeId = Get-OuCreerSeanceRealisee -DbPath $DbPath -SeanceId $seanceId -ClientId $ClientId -DateRealisation $dateIso
            Set-ExerciceRealise -DbPath $DbPath -SeanceRealiseeId $seanceRealiseeId -SeanceExerciceId $seanceExerciceId `
                -Series ([string]$lignesExercice.Count) -Repetitions $reps -Charge $charge -RecuperationS $recup -Tempo $tempo -Notes $notes
            $resultat.Importees++
        } catch {
            $resultat.Erreurs.Add("Ligne '$(Get-ValeurColonne $premiere 'Exercice')' du $dateIso : $($_.Exception.Message)")
        }
    }
    return [pscustomobject]$resultat
}

function Import-FichierSuiviClient {
    <#
        Import unique du fichier renvoye par le client (bouton d'import des seances realisees comme
        du tracking) : chaque onglet est reconnu a son repere en colonne A masquee et range au bon
        endroit — onglets de seance ("D|") -> seances realisees, onglet TRACKING ("T|") -> suivi
        quotidien. Les anciens fichiers en tableau (feuille de seance ou modele de suivi) restent acceptes.
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ClientId,
        [Parameter(Mandatory)] [string] $ExcelPath
    )

    $erreurs = New-Object System.Collections.Generic.List[string]
    $resSeances = [ordered]@{ Importees = 0; IgnoreesSansDate = 0; Erreurs = $erreurs }
    $resJours = [ordered]@{ Importes = 0; IgnoresSansDate = 0; Erreurs = $erreurs }
    $seancesTrouvees = $false; $trackingTrouve = $false
    $pkg = Open-ExcelPackage -Path $ExcelPath
    try {
        foreach ($ws in $pkg.Workbook.Worksheets) {
            if (-not $ws.Dimension) { continue }
            if (Test-OngletAvecRepere -Ws $ws -Repere 'D|*') {
                Import-FeuilleSeanceFormatCoach -DbPath $DbPath -ClientId $ClientId -Ws $ws -Resultat $resSeances
                $seancesTrouvees = $true
            } elseif (Test-OngletAvecRepere -Ws $ws -Repere 'T|*') {
                Import-TrackingOngletCoach -DbPath $DbPath -ClientId $ClientId -Ws $ws -Resultat $resJours
                $trackingTrouve = $true
            }
        }
    } finally {
        Close-ExcelPackage $pkg -NoSave
    }

    if (-not $seancesTrouvees -and -not $trackingTrouve) {
        # Ancien format en tableau : feuille de seance (colonne SeanceExerciceId) ou modele de suivi quotidien
        $premiere = @(Import-Excel -Path $ExcelPath) | Select-Object -First 1
        if ($premiere -and $premiere.PSObject.Properties['SeanceExerciceId']) {
            $r = Import-SeanceRealiseeDepuisExcel -DbPath $DbPath -ClientId $ClientId -ExcelPath $ExcelPath
            $resSeances.Importees = $r.Importees; $resSeances.IgnoreesSansDate = $r.IgnoreesSansDate; $erreurs.AddRange($r.Erreurs)
            $seancesTrouvees = $true
        } else {
            $r = Import-TrackingDepuisExcel -DbPath $DbPath -ClientId $ClientId -ExcelPath $ExcelPath
            $resJours.Importes = $r.Importes; $resJours.IgnoresSansDate = $r.IgnoresSansDate; $erreurs.AddRange($r.Erreurs)
            $trackingTrouve = $true
        }
    }

    return [pscustomobject]@{
        SeancesTrouvees = $seancesTrouvees; SeancesImportees = $resSeances.Importees; SeancesSansDate = $resSeances.IgnoreesSansDate
        TrackingTrouve = $trackingTrouve; JoursImportes = $resJours.Importes; JoursSansDate = $resJours.IgnoresSansDate
        Erreurs = $erreurs
    }
}

Export-ModuleMember -Function Import-BibliothequesDepuisExcel, Export-DonneesVersExcel, Get-EnTetesExcel, `
    Import-QuestionnaireDepuisExcel, Export-ModeleTrackingExcel, Import-TrackingDepuisExcel, `
    Import-RoadmapDepuisExcel, Import-JournalAlimentaireDepuisFatSecret, `
    Import-SeanceRealiseeDepuisExcel, Import-FichierSuiviClient
