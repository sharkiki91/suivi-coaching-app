Set-StrictMode -Version Latest

function Find-NavigateurPdf {
    <# Cherche un navigateur Chromium installe (Edge en priorite, puis Chrome) capable d'exporter en PDF en mode headless. #>
    $chemins = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )
    foreach ($chemin in $chemins) {
        if ($chemin -and (Test-Path $chemin)) { return $chemin }
    }
    return $null
}

function ConvertTo-PdfDepuisHtml {
    param(
        [Parameter(Mandatory)] [string] $HtmlPath,
        [Parameter(Mandatory)] [string] $PdfPath
    )
    $navigateur = Find-NavigateurPdf
    if (-not $navigateur) {
        throw "Aucun navigateur compatible (Edge ou Chrome) n'a ete trouve sur cet ordinateur. Installe Microsoft Edge ou Google Chrome pour pouvoir exporter en PDF, ou utilise l'export Excel a la place."
    }
    $args = @(
        "--headless=new"
        "--disable-gpu"
        "--no-margins"
        "--print-to-pdf=`"$PdfPath`""
        "--no-pdf-header-footer"
        "`"$HtmlPath`""
    )
    $process = Start-Process -FilePath $navigateur -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
    if (-not (Test-Path $PdfPath)) {
        throw "La conversion en PDF a echoue (code de sortie $($process.ExitCode))."
    }
}

function HtmlEncode { param([string]$Texte) [System.Net.WebUtility]::HtmlEncode($Texte) }

function Get-ImageDataUri {
    <# Lit un fichier image et le retourne encode en data URI base64 (pour l'incruster directement dans le HTML/PDF). Retourne $null si le fichier est introuvable. #>
    param([string] $Path)
    if (-not $Path -or -not (Test-Path $Path)) { return $null }
    $mimeTypes = @{ '.jpg' = 'image/jpeg'; '.jpeg' = 'image/jpeg'; '.png' = 'image/png'; '.gif' = 'image/gif'; '.webp' = 'image/webp' }
    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if (-not $mimeTypes.ContainsKey($extension)) { return $null }
    try {
        $octets = [System.IO.File]::ReadAllBytes($Path)
        return "data:$($mimeTypes[$extension]);base64,$([Convert]::ToBase64String($octets))"
    } catch {
        return $null
    }
}

function Get-ValeurAvecDetailSeries {
    <#
        Retourne la valeur a afficher pour un champ (repetitions/charge/recuperation_s) d'un exercice :
        si un detail par serie a ete saisi, une valeur par serie separee par " / " (ex. "15 / 20 / 25") ;
        sinon la valeur globale saisie sur la ligne de l'exercice.
    #>
    param(
        [array] $SeriesDetail,
        [string] $ValeurGlobale,
        [string] $NomChamp
    )
    if ($SeriesDetail -and $SeriesDetail.Count -gt 0) {
        $valeurs = @($SeriesDetail | ForEach-Object { $_.$NomChamp })
        $nonVides = @($valeurs | Where-Object { $_ })
        if ($nonVides.Count -gt 0) {
            return (($valeurs | ForEach-Object { if ($_) { [string]$_ } else { '-' } }) -join ' / ')
        }
    }
    return $ValeurGlobale
}

function Get-NombreSeriesExport {
    <#
        Nombre de lignes (une par serie) a generer pour un exercice dans la feuille de seance a
        remplir : le nombre de series avec detail si saisi, sinon la valeur haute du champ "series"
        global (qui peut etre une fourchette, ex. "3-4"), sinon 1 par defaut.
    #>
    param([array] $SeriesDetail, [string] $SeriesGlobal)
    if ($SeriesDetail -and $SeriesDetail.Count -gt 0) { return $SeriesDetail.Count }
    if ($SeriesGlobal) {
        $nombres = @([regex]::Matches($SeriesGlobal, '\d+') | ForEach-Object { [int]$_.Value })
        if ($nombres.Count -gt 0) { return ($nombres | Measure-Object -Maximum).Maximum }
    }
    return 1
}

function Get-ValeurSeriePrevue {
    <# Valeur prevue (repetitions/charge) pour une serie precise : le detail par serie s'il existe pour ce numero, sinon la valeur globale. #>
    param([array] $SeriesDetail, [int] $NumeroSerie, [string] $ValeurGlobale, [string] $NomChamp)
    $ligneSerie = $SeriesDetail | Where-Object { [int]$_.numero_serie -eq $NumeroSerie } | Select-Object -First 1
    if ($ligneSerie -and $ligneSerie.$NomChamp) { return [string]$ligneSerie.$NomChamp }
    return $ValeurGlobale
}

# ================= PROGRAMMES =================

function Export-ProgrammePdf {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ProgrammeId,
        [Parameter(Mandatory)] [string] $Path
    )

    $prog = (Invoke-SqliteQuery -DataSource $DbPath -Query @"
SELECT p.*, c.nom AS client_nom, c.prenom AS client_prenom
FROM programmes p JOIN clients c ON c.id = p.client_id
WHERE p.id = @Id
"@ -SqlParameters @{ Id = $ProgrammeId })

    $seances = @(Get-Seances -DbPath $DbPath -ProgrammeId $ProgrammeId)
    $dossierData = Split-Path -Path $DbPath -Parent
    $v = $Script:CouleurViolet; $lav = $Script:CouleurLavande; $sep = $Script:CouleurSeparateur

    # Mise en page calquee sur l'onglet TRAINING du fichier d'origine du coach (violet/lavande, Arial gras,
    # bande verticale au nom de la seance, une ligne par serie, trait gris entre les exercices).
    $style = @"
<style>
    @page { size: A4 landscape; margin: 10mm; }
    * { box-sizing: border-box; }
    body { font-family: Arial, Helvetica, sans-serif; color: #000; margin: 0; font-size: 11px; }
    .entete { background: $v; color: #fff; padding: 10px 14px; font-weight: bold; font-size: 18px; letter-spacing: .5px; }
    .entete .client { font-size: 13px; font-weight: normal; margin-top: 3px; }
    .notes-prog { color: $v; font-style: italic; margin: 8px 2px 12px; }
    table { border-collapse: collapse; width: 100%; }
    .semaine { margin-bottom: 16px; table-layout: fixed; }
    .semaine th { background: $v; color: #fff; font-size: 11px; padding: 5px; border: 1px solid #fff; }
    .semaine td { text-align: center; font-weight: bold; color: $v; padding: 8px 4px; border: 1px solid $sep; font-size: 11px; }
    .semaine td.repos { color: #999; font-weight: normal; }
    .bloc { display: flex; margin-bottom: 16px; border: 2px solid $v; break-inside: avoid; page-break-inside: avoid; }   /* une seance n'est jamais coupee entre deux pages */
    .bloc .bande { background: $lav; color: #fff; font-size: 17px; font-weight: bold; width: 46px; min-width: 46px; border-right: 2px solid $v; display: flex; align-items: center; justify-content: center; }
    .bloc .bande div { writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: center; }
    .bloc .bande .jour { font-size: 11px; font-weight: normal; }
    .seance { flex: 1; }
    .seance th { background: $v; color: #fff; font-size: 10px; font-weight: bold; padding: 6px 3px; border-right: 2px solid $sep; text-transform: uppercase; }
    .seance td { text-align: center; vertical-align: middle; font-weight: bold; font-size: 10.5px; padding: 1px 4px; line-height: 1.25; }
    .seance td.num { background: $v; color: #fff; width: 26px; }
    .seance td.violet { color: $v; text-transform: uppercase; }
    .seance td.nom { width: 24%; }
    .seance td.nom .note { display: block; text-transform: none; font-weight: normal; font-style: italic; color: #555; font-size: 10px; margin-top: 2px; }
    .seance td.set { width: 34px; }
    .seance { table-layout: fixed; }
    .seance tbody.exercice { break-inside: avoid; }
    .seance tbody.exercice + tbody.exercice { border-top: 2px solid $sep; }
    .seance tbody.exercice tr:not(:last-child) td.serie { border-bottom: 1px solid #ECE8F5; }
    .seance td.serie { border-left: 1px solid #ECE8F5; }
    .miniature-exercice { width: 46px; height: 46px; object-fit: cover; border-radius: 3px; display: block; margin: 0 auto 3px; }
    a.lien-video { color: $v; text-decoration: none; white-space: nowrap; }
    .vide { color: #999; font-style: italic; padding: 10px; }
</style>
"@
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<html><head><meta charset='utf-8'>$style</head><body>")
    [void]$sb.Append("<div class='entete'>PROGRAMME $(HtmlEncode ([string]$prog.nom).ToUpperInvariant())")
    $sousTitre = "$(HtmlEncode $prog.client_prenom) $(HtmlEncode $prog.client_nom)"
    if ($prog.date_debut) { $sousTitre += " &mdash; d&eacute;but le $(HtmlEncode ([datetime]$prog.date_debut).ToString('dd/MM/yyyy'))" }
    [void]$sb.Append("<div class='client'>$sousTitre</div></div>")
    if ($prog.notes) { [void]$sb.Append("<p class='notes-prog'>$(HtmlEncode $prog.notes)</p>") } else { [void]$sb.Append("<div style='height:10px'></div>") }

    $joursSemaine = @('Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche')
    $seancesAvecJour = @($seances | Where-Object { $_.jour_semaine })
    if ($seancesAvecJour.Count -gt 0) {
        [void]$sb.Append("<table class='semaine'><tr>")
        foreach ($jour in $joursSemaine) { [void]$sb.Append("<th>$($jour.ToUpperInvariant())</th>") }
        [void]$sb.Append("</tr><tr>")
        foreach ($jour in $joursSemaine) {
            $seancesDuJour = @($seancesAvecJour | Where-Object { $_.jour_semaine -eq $jour })
            if ($seancesDuJour.Count -gt 0) { [void]$sb.Append("<td>$(($seancesDuJour | ForEach-Object { HtmlEncode ([string]$_.nom).ToUpperInvariant() }) -join '<br>')</td>") }
            else { [void]$sb.Append("<td class='repos'>Repos</td>") }
        }
        [void]$sb.Append("</tr></table>")
    }

    foreach ($s in $seances) {
        $exercices = @(Get-SeanceExercices -DbPath $DbPath -SeanceId ([int]$s.id))
        # Colonne image seulement si au moins un exercice de la seance en a une
        $avecImage = @($exercices | Where-Object { $_.image_path }).Count -gt 0
        $jourHtml = if ($s.jour_semaine) { "<br><span class='jour'>$(HtmlEncode ([string]$s.jour_semaine).ToUpperInvariant())</span>" } else { '' }
        [void]$sb.Append("<div class='bloc'><div class='bande'><div>$(HtmlEncode ([string]$s.nom).ToUpperInvariant())$jourHtml</div></div>")
        [void]$sb.Append("<table class='seance'><colgroup><col style='width:28px'><col style='width:24%'><col style='width:9%'><col style='width:36px'><col style='width:8%'><col style='width:9%'><col style='width:8%'><col style='width:7%'><col style='width:13%'><col style='width:9%'></colgroup><thead><tr>")
        [void]$sb.Append("<th>#</th><th>Exercice</th><th>Variante</th><th>Set</th><th>Reps</th><th>Charge</th><th>R&eacute;cup (s)</th><th>Tempo</th><th>Muscle cible</th><th>Lien</th></tr></thead>")
        if ($exercices.Count -eq 0) {
            [void]$sb.Append("<tbody><tr><td colspan='10' class='vide'>Aucun exercice dans cette s&eacute;ance.</td></tr></tbody></table></div>")
            continue
        }
        $numero = 0
        foreach ($e in $exercices) {
            $numero++
            $detail = @(Get-SeanceExerciceSeries -DbPath $DbPath -SeanceExerciceId ([int]$e.id))
            $series = @(Get-LignesSeriesExercice -Exercice $e -SeriesDetail $detail)
            $n = $series.Count
            $recups = @($series | ForEach-Object { [string]$_.Recup } | Select-Object -Unique)
            $recupUnique = $recups.Count -le 1

            $nomHtml = HtmlEncode ([string]$e.exercice_nom)
            if ($avecImage -and $e.image_path) {
                $dataUri = Get-ImageDataUri -Path (Join-Path $dossierData $e.image_path)
                if ($dataUri) { $nomHtml = "<img class='miniature-exercice' src='$dataUri' alt='' />$nomHtml" }
            }
            if ($e.notes) { $nomHtml += "<span class='note'>$(HtmlEncode $e.notes)</span>" }
            $lienHtml = if ($e.lien_video) { "<a class='lien-video' href='$(HtmlEncode $e.lien_video)'>&#9654; VID&Eacute;O</a>" } else { '' }

            [void]$sb.Append("<tbody class='exercice'>")
            foreach ($sr in $series) {
                [void]$sb.Append('<tr>')
                if ($sr.Numero -eq 1) {
                    [void]$sb.Append("<td class='num' rowspan='$n'>$numero</td>")
                    [void]$sb.Append("<td class='violet nom' rowspan='$n'>$nomHtml</td>")
                    [void]$sb.Append("<td class='violet' rowspan='$n'>$(HtmlEncode $e.variante)</td>")
                }
                [void]$sb.Append("<td class='set serie'>$($sr.Numero)</td><td class='serie'>$(HtmlEncode $sr.Repetitions)</td><td class='serie'>$(HtmlEncode $sr.Charge)</td>")
                if (-not $recupUnique) { [void]$sb.Append("<td class='serie'>$(HtmlEncode $sr.Recup)</td>") }
                elseif ($sr.Numero -eq 1) { [void]$sb.Append("<td rowspan='$n'>$(HtmlEncode $recups[0])</td>") }
                if ($sr.Numero -eq 1) {
                    [void]$sb.Append("<td rowspan='$n'>$(HtmlEncode $e.tempo)</td>")
                    [void]$sb.Append("<td class='violet' rowspan='$n'>$(HtmlEncode $e.muscle_cible)</td>")
                    [void]$sb.Append("<td rowspan='$n'>$lienHtml</td>")
                }
                [void]$sb.Append('</tr>')
            }
            [void]$sb.Append('</tbody>')
        }
        [void]$sb.Append('</table></div>')
    }
    [void]$sb.Append("</body></html>")

    $tempHtml = [System.IO.Path]::GetTempFileName() + ".html"
    [System.IO.File]::WriteAllText($tempHtml, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
    try {
        ConvertTo-PdfDepuisHtml -HtmlPath $tempHtml -PdfPath $Path
    } finally {
        Remove-Item $tempHtml -Force -ErrorAction SilentlyContinue
    }
}

function Export-ProgrammeExcel {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ProgrammeId,
        [Parameter(Mandatory)] [string] $Path
    )

    # Meme presentation que la feuille de seance (onglet TRAINING d'origine), sans les blocs de suivi
    $prog = Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT notes, date_debut FROM programmes WHERE id = @Id" -SqlParameters @{ Id = $ProgrammeId }
    $morceaux = @()
    if ($prog.date_debut) { $morceaux += "Debut le $(([datetime]$prog.date_debut).ToString('dd/MM/yyyy'))" }
    if ($prog.notes) { $morceaux += [string]$prog.notes }
    Export-FeuilleSeanceExcel -DbPath $DbPath -ProgrammeId $ProgrammeId -Path $Path -NbBlocs 0 -TexteConsigne ($morceaux -join ' - ')
}

# --- Charte graphique du fichier d'origine du coach (onglet TRAINING de "SUIVI 2.0.xlsx") ---
$Script:CouleurViolet = '#674EA7'     # en-tetes, colonne #, noms d'exercices
$Script:CouleurLavande = '#8E7CC3'    # bande du nom de seance, titres "SEANCE n"
$Script:CouleurSeparateur = '#B7B7B7' # trait entre deux exercices
$Script:CouleurLigneSerie = '#D9D2E9' # trait leger entre deux series (zones a remplir)
$Script:NbBlocsSuivi = 6              # nombre de seances a noter cote a cote (comme le fichier d'origine)

function Set-StyleExcel {
    <# Applique un style "charte coach" a une plage EPPlus. #>
    param(
        [Parameter(Mandatory)] $Plage,
        [string] $Fond, [string] $Couleur, [int] $Taille = 9,
        [switch] $Gras, [switch] $Italique, [switch] $Gauche, [int] $Rotation = 0
    )
    $st = $Plage.Style
    $st.Font.Name = 'Arial'; $st.Font.Size = $Taille; $st.Font.Bold = [bool]$Gras; $st.Font.Italic = [bool]$Italique
    if ($Couleur) { $st.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml($Couleur)) }
    if ($Fond) {
        $st.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $st.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($Fond))
    }
    $st.HorizontalAlignment = if ($Gauche) { [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left } else { [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center }
    $st.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $st.WrapText = $true
    if ($Rotation) { $st.TextRotation = $Rotation }
}

function Set-BordureExcel {
    param([Parameter(Mandatory)] $Plage, [string[]] $Cotes = @('Bottom'), [string] $Couleur = $Script:CouleurSeparateur, [string] $Epaisseur = 'Medium')
    foreach ($cote in $Cotes) {
        $b = $Plage.Style.Border.$cote
        $b.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::$Epaisseur
        $b.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml($Couleur))
    }
}

function Set-FusionExcel {
    <# Fusionne une plage (si elle fait plus d'une cellule) et y ecrit une valeur. #>
    param([Parameter(Mandatory)] $Ws, [int] $L1, [int] $C1, [int] $L2, [int] $C2, $Valeur)
    $plage = $Ws.Cells[$L1, $C1, $L2, $C2]
    if ($L2 -gt $L1 -or $C2 -gt $C1) { $plage.Merge = $true }
    if ($null -ne $Valeur -and "$Valeur" -ne '') { $Ws.Cells[$L1, $C1].Value = $Valeur }
    return ,$plage   # virgule : sinon PowerShell deroule la plage cellule par cellule
}

function Get-LignesSeriesExercice {
    <#
        Une entree par serie a afficher pour un exercice : numero, repetitions/charge/recup prevues
        (detail par serie si saisi, sinon la valeur globale de l'exercice).
    #>
    param([Parameter(Mandatory)] $Exercice, [array] $SeriesDetail)
    $nb = Get-NombreSeriesExport -SeriesDetail $SeriesDetail -SeriesGlobal $Exercice.series
    for ($n = 1; $n -le $nb; $n++) {
        [pscustomobject]@{
            Numero = $n
            Repetitions = Get-ValeurSeriePrevue -SeriesDetail $SeriesDetail -NumeroSerie $n -ValeurGlobale $Exercice.repetitions -NomChamp 'repetitions'
            Charge = Get-ValeurSeriePrevue -SeriesDetail $SeriesDetail -NumeroSerie $n -ValeurGlobale $Exercice.charge -NomChamp 'charge'
            Recup = Get-ValeurSeriePrevue -SeriesDetail $SeriesDetail -NumeroSerie $n -ValeurGlobale $Exercice.recuperation_s -NomChamp 'recuperation_s'
        }
    }
}

function Export-FeuilleSeanceExcel {
    <#
        Genere la feuille de suivi des performances a remplir par le client, sur le modele de
        l'onglet TRAINING du fichier d'origine du coach : pour chaque seance, le programme a gauche
        (une ligne par serie : #, exercice, variante, set, reps, charge, recup, tempo, muscle, lien)
        et, a droite, 6 blocs "SEANCE 1..6" cote a cote (DATE, puis REPS / CHARGE par serie et NOTES
        par exercice) pour noter 6 seances successives. La partie programme reste figee a l'ecran.
        Un onglet par seance (nomme comme la seance), imprime en entier sur une seule page paysage.

        La colonne A (masquee) contient des reperes techniques ("D|seanceId" sur la ligne des dates,
        "S|seanceId|seanceExerciceId|serie" sur chaque ligne de serie) relus par
        Import-SeanceRealiseeDepuisExcel : ne pas la supprimer.

        -NbBlocs 0 produit le programme seul (meme presentation, sans les blocs de suivi) : utilise
        par Export-ProgrammeExcel, avec -TexteConsigne pour la ligne d'explication sous le titre.
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ProgrammeId,
        [Parameter(Mandatory)] [string] $Path,
        [int] $NbBlocs = $Script:NbBlocsSuivi,
        [string] $TexteConsigne
    )

    $prog = Invoke-SqliteQuery -DataSource $DbPath -Query @"
SELECT p.nom, c.nom AS client_nom, c.prenom AS client_prenom
FROM programmes p JOIN clients c ON c.id = p.client_id WHERE p.id = @Id
"@ -SqlParameters @{ Id = $ProgrammeId }
    $seances = @(Get-Seances -DbPath $DbPath -ProgrammeId $ProgrammeId)

    # Colonnes : A repere masque | B bande seance | C..L programme | M espace | puis 6 blocs de 4 colonnes + 1 espace
    $colonnesProgramme = [ordered]@{ '#' = 4; 'EXERCICE' = 24; 'VARIANTE' = 11; 'SET' = 4.5; 'REPS' = 7; 'CHARGE' = 8; 'RECUP (s)' = 8; 'TEMPO' = 7; 'MUSCLE CIBLE' = 12; 'LIEN' = 7 }
    $cB = 2; $cDebut = 3; $cFin = $cDebut + $colonnesProgramme.Count - 1   # C..L
    $cPremierBloc = $cFin + 2                                             # N
    $largeurBloc = 5                                                      # #, REPS, CHARGE, NOTES + espace

    if (Test-Path $Path) { Remove-Item $Path -Force }
    $pkg = Open-ExcelPackage -Path $Path -Create
    try {
        $titre = "PROGRAMME $($prog.nom) - $($prog.client_prenom) $($prog.client_nom)".ToUpperInvariant()
        $consigne = if ($NbBlocs -gt 0) { "A chaque seance : note la DATE en haut d'un bloc SEANCE, puis tes repetitions et la charge serie par serie (et une note si besoin). Seance suivante = bloc suivant." } else { $TexteConsigne }
        $nomsOnglets = @{}

        # Un onglet par seance : chaque tableau s'imprime en entier sur une seule page, sans etre coupe en deux.
        foreach ($s in $seances) {
            $exercices = @(Get-SeanceExercices -DbPath $DbPath -SeanceId ([int]$s.id))
            if ($exercices.Count -eq 0) { continue }

            # Nom d'onglet Excel : 31 caracteres max, sans [ ] : * ? / \, et unique
            $nomOnglet = (([string]$s.nom) -replace '[\[\]:*?/\\]', '-').Trim()
            if (-not $nomOnglet) { $nomOnglet = 'Seance' }
            if ($nomOnglet.Length -gt 28) { $nomOnglet = $nomOnglet.Substring(0, 28).Trim() }
            $base = $nomOnglet; $k = 2
            while ($nomsOnglets.ContainsKey($nomOnglet.ToUpperInvariant())) { $nomOnglet = "$base ($k)"; $k++ }
            $nomsOnglets[$nomOnglet.ToUpperInvariant()] = $true

            $ws = Add-Worksheet -ExcelPackage $pkg -WorksheetName $nomOnglet
            $ws.View.ShowGridLines = $false
            $ws.Column(1).Hidden = $true
            $ws.Column($cB).Width = 6
            $i = 0; foreach ($k in $colonnesProgramme.Keys) { $ws.Column($cDebut + $i).Width = $colonnesProgramme[$k]; $i++ }
            $ws.Column($cFin + 1).Width = 2.5
            for ($b = 0; $b -lt $NbBlocs; $b++) {
                $c0 = $cPremierBloc + $b * $largeurBloc
                $ws.Column($c0).Width = 6.5;   # assez large pour "DATE" sur une ligne
                $ws.Column($c0 + 1).Width = 7; $ws.Column($c0 + 2).Width = 8; $ws.Column($c0 + 3).Width = 16; $ws.Column($c0 + 4).Width = 2.5
            }

            # Titre + consigne (repetes sur chaque onglet)
            Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 1 -C1 $cB -L2 1 -C2 $cFin -Valeur $titre) -Fond $Script:CouleurViolet -Couleur '#FFFFFF' -Taille 12 -Gras
            $ws.Row(1).Height = 24
            Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 2 -C1 $cB -L2 2 -C2 $cFin -Valeur $consigne) -Couleur $Script:CouleurViolet -Italique -Gauche
            $ws.Row(2).Height = 26

            $ligne = 4
            $lTitre = $ligne; $lDate = $ligne + 1; $lEntete = $ligne + 2
            $ws.Cells[$lDate, 1].Value = "D|$($s.id)"
            $ws.Row($lTitre).Height = 18; $ws.Row($lDate).Height = 20; $ws.Row($lEntete).Height = 18

            # En-tetes du programme (sur 2 lignes, comme l'original)
            $i = 0
            foreach ($k in $colonnesProgramme.Keys) {
                $p = Set-FusionExcel -Ws $ws -L1 $lDate -C1 ($cDebut + $i) -L2 $lEntete -C2 ($cDebut + $i) -Valeur $k
                Set-StyleExcel -Plage $p -Fond $Script:CouleurViolet -Couleur '#FFFFFF' -Gras
                Set-BordureExcel -Plage $p -Cotes @('Right') -Couleur $Script:CouleurSeparateur
                $i++
            }
            # Blocs SEANCE n : titre, ligne DATE a remplir, en-tetes
            for ($b = 0; $b -lt $NbBlocs; $b++) {
                $c0 = $cPremierBloc + $b * $largeurBloc
                Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $lTitre -C1 $c0 -L2 $lTitre -C2 ($c0 + 3) -Valeur "SEANCE $($b + 1)") -Fond $Script:CouleurLavande -Couleur '#FFFFFF' -Gras -Taille 10
                Set-StyleExcel -Plage $ws.Cells[$lDate, $c0] -Fond $Script:CouleurViolet -Couleur '#FFFFFF' -Gras -Taille 8
                $ws.Cells[$lDate, $c0].Style.WrapText = $false
                $ws.Cells[$lDate, $c0].Value = 'DATE'
                $pDate = Set-FusionExcel -Ws $ws -L1 $lDate -C1 ($c0 + 1) -L2 $lDate -C2 ($c0 + 3) -Valeur $null
                Set-StyleExcel -Plage $pDate -Fond '#F3F0FA' -Couleur '#000000' -Gras -Taille 10
                $pDate.Style.Numberformat.Format = 'dd/mm/yyyy'
                Set-BordureExcel -Plage $pDate -Cotes @('Top', 'Bottom', 'Left', 'Right') -Couleur $Script:CouleurViolet -Epaisseur 'Thin'
                $j = 0
                foreach ($k in @('#', 'REPS', 'CHARGE', 'NOTES')) {
                    $ws.Cells[$lEntete, ($c0 + $j)].Value = $k
                    Set-StyleExcel -Plage $ws.Cells[$lEntete, ($c0 + $j)] -Fond $Script:CouleurViolet -Couleur '#FFFFFF' -Gras
                    $j++
                }
                Set-BordureExcel -Plage $ws.Cells[$lEntete, $c0, $lEntete, ($c0 + 3)] -Cotes @('Bottom') -Couleur $Script:CouleurSeparateur
            }

            # Une ligne par serie
            $ligne = $lEntete + 1
            $numeroExercice = 0
            foreach ($e in $exercices) {
                $numeroExercice++
                $detail = @(Get-SeanceExerciceSeries -DbPath $DbPath -SeanceExerciceId ([int]$e.id))
                $series = @(Get-LignesSeriesExercice -Exercice $e -SeriesDetail $detail)
                $l1 = $ligne; $l2 = $ligne + $series.Count - 1
                foreach ($sr in $series) {
                    $l = $l1 + $sr.Numero - 1
                    $ws.Row($l).Height = 16
                    $ws.Cells[$l, 1].Value = "S|$($s.id)|$($e.id)|$($sr.Numero)"
                    $ws.Cells[$l, ($cDebut + 3)].Value = $sr.Numero
                    $ws.Cells[$l, ($cDebut + 4)].Value = [string]$sr.Repetitions
                    $ws.Cells[$l, ($cDebut + 5)].Value = [string]$sr.Charge
                    Set-StyleExcel -Plage $ws.Cells[$l, ($cDebut + 3), $l, ($cDebut + 5)] -Couleur '#000000' -Gras
                }
                $nom = ([string]$e.exercice_nom).ToUpperInvariant()
                if ($e.notes) { $nom += "`n($($e.notes))" }
                Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $l1 -C1 $cDebut -L2 $l2 -C2 $cDebut -Valeur $numeroExercice) -Fond $Script:CouleurViolet -Couleur '#FFFFFF' -Gras
                Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $l1 -C1 ($cDebut + 1) -L2 $l2 -C2 ($cDebut + 1) -Valeur $nom) -Couleur $Script:CouleurViolet -Gras
                Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $l1 -C1 ($cDebut + 2) -L2 $l2 -C2 ($cDebut + 2) -Valeur ([string]$e.variante).ToUpperInvariant()) -Couleur $Script:CouleurViolet -Gras
                # Recup : une seule cellule si identique pour toutes les series (comme l'original), sinon serie par serie
                $recups = @($series | ForEach-Object { [string]$_.Recup } | Select-Object -Unique)
                if ($recups.Count -le 1) {
                    Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $l1 -C1 ($cDebut + 6) -L2 $l2 -C2 ($cDebut + 6) -Valeur ([string]$recups[0])) -Couleur '#000000' -Gras
                } else {
                    foreach ($sr in $series) { $ws.Cells[($l1 + $sr.Numero - 1), ($cDebut + 6)].Value = [string]$sr.Recup }
                    Set-StyleExcel -Plage $ws.Cells[$l1, ($cDebut + 6), $l2, ($cDebut + 6)] -Couleur '#000000' -Gras
                }
                Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $l1 -C1 ($cDebut + 7) -L2 $l2 -C2 ($cDebut + 7) -Valeur ([string]$e.tempo)) -Couleur '#000000' -Gras
                Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $l1 -C1 ($cDebut + 8) -L2 $l2 -C2 ($cDebut + 8) -Valeur ([string]$e.muscle_cible).ToUpperInvariant()) -Couleur $Script:CouleurViolet -Gras
                $pLien = Set-FusionExcel -Ws $ws -L1 $l1 -C1 ($cDebut + 9) -L2 $l2 -C2 ($cDebut + 9) -Valeur $null
                Set-StyleExcel -Plage $pLien -Couleur $Script:CouleurViolet -Gras
                if ($e.lien_video) {
                    try { $ws.Cells[$l1, ($cDebut + 9)].Hyperlink = New-Object System.Uri([string]$e.lien_video) } catch { }
                    $ws.Cells[$l1, ($cDebut + 9)].Value = 'VIDEO'
                    $ws.Cells[$l1, ($cDebut + 9)].Style.Font.UnderLine = $true
                }
                Set-BordureExcel -Plage $ws.Cells[$l1, $cDebut, $l2, $cFin] -Cotes @('Right') -Couleur $Script:CouleurSeparateur -Epaisseur 'Thin'
                Set-BordureExcel -Plage $ws.Cells[$l2, $cDebut, $l2, $cFin] -Cotes @('Bottom')

                for ($b = 0; $b -lt $NbBlocs; $b++) {
                    $c0 = $cPremierBloc + $b * $largeurBloc
                    foreach ($sr in $series) { $ws.Cells[($l1 + $sr.Numero - 1), $c0].Value = $sr.Numero }
                    Set-StyleExcel -Plage $ws.Cells[$l1, $c0, $l2, $c0] -Fond $Script:CouleurViolet -Couleur '#FFFFFF' -Gras
                    Set-StyleExcel -Plage $ws.Cells[$l1, ($c0 + 1), $l2, ($c0 + 2)] -Couleur '#000000' -Gras -Taille 10
                    if ($series.Count -gt 1) { Set-BordureExcel -Plage $ws.Cells[$l1, ($c0 + 1), ($l2 - 1), ($c0 + 2)] -Cotes @('Bottom') -Couleur $Script:CouleurLigneSerie -Epaisseur 'Thin' }
                    Set-BordureExcel -Plage $ws.Cells[$l1, ($c0 + 1), $l2, ($c0 + 2)] -Cotes @('Right') -Couleur $Script:CouleurLigneSerie -Epaisseur 'Thin'
                    Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $l1 -C1 ($c0 + 3) -L2 $l2 -C2 ($c0 + 3) -Valeur $null) -Couleur '#000000' -Taille 9 -Gauche
                    Set-BordureExcel -Plage $ws.Cells[$l1, ($c0 + 3), $l2, ($c0 + 3)] -Cotes @('Right') -Couleur $Script:CouleurViolet
                    Set-BordureExcel -Plage $ws.Cells[$l2, $c0, $l2, ($c0 + 3)] -Cotes @('Bottom') -Couleur $Script:CouleurViolet
                }
                $ligne = $l2 + 1
            }

            # Bande verticale du nom de seance (+ jour), sur toute la hauteur du bloc
            $nomSeance = ([string]$s.nom).ToUpperInvariant()
            if ($s.jour_semaine) { $nomSeance += " ($(([string]$s.jour_semaine).ToUpperInvariant()))" }
            Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $lTitre -C1 $cB -L2 ($ligne - 1) -C2 $cB -Valeur $nomSeance) -Fond $Script:CouleurLavande -Couleur '#FFFFFF' -Gras -Taille 14 -Rotation 90
            Set-BordureExcel -Plage $ws.Cells[$lTitre, $cB, ($ligne - 1), $cB] -Cotes @('Right') -Couleur $Script:CouleurViolet

            # La partie programme reste visible quand on fait defiler les blocs SEANCE vers la droite
            if ($NbBlocs -gt 0) { $ws.View.FreezePanes(1, ($cFin + 2)) }
            # Impression : tout l'onglet (donc toute la seance) sur une seule page paysage
            $ws.PrinterSettings.Orientation = [OfficeOpenXml.eOrientation]::Landscape
            $ws.PrinterSettings.FitToPage = $true; $ws.PrinterSettings.FitToWidth = 1; $ws.PrinterSettings.FitToHeight = 1
            $ws.PrinterSettings.HorizontalCentered = $true
            $ws.PrinterSettings.TopMargin = 0.4; $ws.PrinterSettings.BottomMargin = 0.4; $ws.PrinterSettings.LeftMargin = 0.3; $ws.PrinterSettings.RightMargin = 0.3
        }

        # Un classeur Excel doit contenir au moins un onglet
        if ($nomsOnglets.Count -eq 0) {
            $ws = Add-Worksheet -ExcelPackage $pkg -WorksheetName 'TRAINING'
            $ws.Cells[1, 2].Value = 'Aucune seance avec des exercices dans ce programme.'
        }
    } finally {
        Close-ExcelPackage $pkg
    }
}

# ================= NUTRITION =================

# --- Charte de l'onglet NUTRITION d'origine ---
$Script:CouleurVioletFonce = '#351C75'  # ligne "PLAN JOURNALIER / TOTAL"
$Script:CouleurLilas = '#B4A7D6'        # en-tetes KCAL..FIB et etiquettes PRO/GLU/LIP/FIB
$Script:CouleurGrisClair = '#D9D9D9'    # traits entre les aliments

function Format-Nutri { param($Valeur) if ($null -eq $Valeur -or "$Valeur" -eq '') { return '' }; return [string][math]::Round([double]$Valeur) }

function Get-DonneesJourNutrition {
    <#
        Prepare un type de jour pour l'export : repas (avec au moins 4 lignes, comme le fichier
        d'origine, pour loger les totaux PRO/GLU/LIP/FIB du repas), totaux par repas et totaux du jour.
    #>
    param([Parameter(Mandatory)] [string] $DbPath, [Parameter(Mandatory)] [int] $TypeJourId)
    $jour = [pscustomobject]@{ Kcal = 0.0; Proteines = 0.0; Glucides = 0.0; Lipides = 0.0; Fibres = 0.0 }
    $repasListe = foreach ($r in @(Get-Repas -DbPath $DbPath -TypeJourId $TypeJourId)) {
        $lignes = @(Get-RepasAliments -DbPath $DbPath -RepasId ([int]$r.id))
        $tot = Get-TotauxRepas -Lignes $lignes
        foreach ($k in 'Kcal', 'Proteines', 'Glucides', 'Lipides', 'Fibres') { $jour.$k += $tot.$k }
        [pscustomobject]@{ Nom = [string]$r.nom; Lignes = $lignes; Totaux = $tot; NbLignes = [math]::Max(4, $lignes.Count) }
    }
    [pscustomobject]@{ Repas = @($repasListe); Totaux = $jour }
}

function Get-RepartitionTotaux {
    <# Repartit $NbLignes lignes en 4 groupes (PRO, GLU, LIP, FIB) : retourne la taille de chaque groupe. #>
    param([int] $NbLignes)
    $base = [math]::Floor($NbLignes / 4); $reste = $NbLignes % 4
    [int[]]@(0..3 | ForEach-Object { [int]($base + $(if ($_ -lt $reste) { 1 } else { 0 })) })   # [int] : IndexOf compare ensuite a un [int]
}

function Export-PlanNutritionPdf {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $PlanNutritionId,
        [Parameter(Mandatory)] [string] $Path
    )

    $plan = (Invoke-SqliteQuery -DataSource $DbPath -Query @"
SELECT pn.*, c.nom AS client_nom, c.prenom AS client_prenom
FROM plans_nutrition pn JOIN clients c ON c.id = pn.client_id
WHERE pn.id = @Id
"@ -SqlParameters @{ Id = $PlanNutritionId })

    $typesJour = @(Get-TypesJour -DbPath $DbPath -PlanNutritionId $PlanNutritionId)
    $v = $Script:CouleurViolet; $lav = $Script:CouleurLavande; $vf = $Script:CouleurVioletFonce; $lil = $Script:CouleurLilas; $gc = $Script:CouleurGrisClair

    # Mise en page calquee sur l'onglet NUTRITION du fichier d'origine du coach.
    $style = @"
<style>
    @page { size: A4 landscape; margin: 10mm; }
    * { box-sizing: border-box; }
    body { font-family: Arial, Helvetica, sans-serif; color: #000; margin: 0; font-size: 11px; }
    .entete { background: $v; color: #fff; padding: 10px 14px; font-weight: bold; font-size: 18px; letter-spacing: .5px; }
    .entete .client { font-size: 13px; font-weight: normal; margin-top: 3px; }
    .notes-plan { color: $v; font-style: italic; margin: 8px 2px 12px; }
    .jour { display: flex; gap: 18px; align-items: flex-start; margin-bottom: 18px; break-inside: avoid; }
    table { border-collapse: collapse; }
    .plan { flex: 1; table-layout: fixed; border: 2px solid $v; }
    .plan th, .plan td { text-align: center; vertical-align: middle; padding: 3px 4px; }
    .plan .l1 th { background: $vf; color: #fff; font-size: 11px; padding: 6px 4px; }
    .plan .l2 th { background: $v; color: #fff; font-size: 10px; }
    .plan .l2 th.nut { background: $lil; }
    .plan td { font-size: 11px; border-bottom: 1px solid $gc; }
    .plan td.aliment { text-align: left; }
    .plan td.repas { background: $lav; color: #fff; font-weight: bold; font-size: 11px; text-transform: uppercase; border-bottom: none; }
    .plan td.etiq { background: $lil; color: #fff; font-weight: bold; font-size: 10px; }
    .plan td.tot { font-weight: bold; }
    .plan tbody.repas-bloc { break-inside: avoid; }
    .plan tbody.repas-bloc + tbody.repas-bloc { border-top: 2px solid $v; }
    .recap { width: 170px; border: 2px solid $v; }
    .recap th { background: $v; color: #fff; padding: 10px 4px; font-size: 11px; }
    .recap td { padding: 9px 6px; text-align: center; font-weight: bold; border-bottom: 1px solid $gc; }
    .recap td.etiq { background: $lil; color: #fff; font-size: 10px; width: 50%; }
    .vide { color: #999; font-style: italic; }
</style>
"@
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<html><head><meta charset='utf-8'>$style</head><body>")
    [void]$sb.Append("<div class='entete'>PLAN NUTRITION $(HtmlEncode ([string]$plan.nom).ToUpperInvariant())")
    $sousTitre = "$(HtmlEncode $plan.client_prenom) $(HtmlEncode $plan.client_nom)"
    if ($plan.date_debut) { $sousTitre += " &mdash; d&eacute;but le $(HtmlEncode ([datetime]$plan.date_debut).ToString('dd/MM/yyyy'))" }
    [void]$sb.Append("<div class='client'>$sousTitre</div></div>")
    if ($plan.notes) { [void]$sb.Append("<p class='notes-plan'>$(HtmlEncode $plan.notes)</p>") } else { [void]$sb.Append("<div style='height:10px'></div>") }

    $etiquettes = @(@('PRO', 'Proteines'), @('GLU', 'Glucides'), @('LIP', 'Lipides'), @('FIB', 'Fibres'))
    foreach ($tj in $typesJour) {
        $d = Get-DonneesJourNutrition -DbPath $DbPath -TypeJourId ([int]$tj.id)
        $t = $d.Totaux
        [void]$sb.Append("<div class='jour'><table class='plan'><colgroup><col style='width:11%'><col style='width:25%'><col style='width:10%'><col style='width:8%'><col style='width:7%'><col style='width:7%'><col style='width:7%'><col style='width:7%'><col style='width:7%'><col style='width:6%'><col style='width:7%'></colgroup>")
        [void]$sb.Append("<thead><tr class='l1'><th colspan='2'>PLAN JOURNALIER &mdash; $(HtmlEncode ([string]$tj.nom).ToUpperInvariant())</th><th>TOTAL</th><th>$(Format-Nutri $t.Kcal)</th><th>$(Format-Nutri $t.Proteines)</th><th>$(Format-Nutri $t.Glucides)</th><th>$(Format-Nutri $t.Lipides)</th><th>$(Format-Nutri $t.Fibres)</th><th colspan='3'></th></tr>")
        [void]$sb.Append("<tr class='l2'><th>REPAS</th><th>ALIMENTS</th><th>QUANTIT&Eacute;</th><th class='nut'>KCAL</th><th class='nut'>PRO</th><th class='nut'>GLU</th><th class='nut'>LIP</th><th class='nut'>FIB</th><th colspan='3'>TOTAUX</th></tr></thead>")
        if ($d.Repas.Count -eq 0) { [void]$sb.Append("<tbody><tr><td colspan='11' class='vide'>Aucun repas pour ce type de jour.</td></tr></tbody>") }
        foreach ($r in $d.Repas) {
            $n = $r.NbLignes
            $groupes = Get-RepartitionTotaux -NbLignes $n
            [int[]]$debutsGroupes = @(0, $groupes[0], ($groupes[0] + $groupes[1]), ($groupes[0] + $groupes[1] + $groupes[2]))
            [void]$sb.Append("<tbody class='repas-bloc'>")
            for ($i = 0; $i -lt $n; $i++) {
                [void]$sb.Append('<tr>')
                if ($i -eq 0) { [void]$sb.Append("<td class='repas' rowspan='$n'>$(HtmlEncode $r.Nom)</td>") }
                if ($i -lt $r.Lignes.Count) {
                    $l = $r.Lignes[$i]
                    [void]$sb.Append("<td class='aliment'>$(HtmlEncode $l.aliment_nom)</td><td>$($l.quantite) $(HtmlEncode $l.unite)</td><td>$(Format-Nutri $l.kcal_calc)</td><td>$(Format-Nutri $l.proteines_calc)</td><td>$(Format-Nutri $l.glucides_calc)</td><td>$(Format-Nutri $l.lipides_calc)</td><td>$(Format-Nutri $l.fibres_calc)</td>")
                } else {
                    [void]$sb.Append("<td class='aliment'>&nbsp;</td><td></td><td></td><td></td><td></td><td></td><td></td>")
                }
                $g = [array]::IndexOf($debutsGroupes, $i)
                if ($g -ge 0 -and $groupes[$g] -gt 0) {
                    [void]$sb.Append("<td class='etiq' rowspan='$($groupes[$g])'>$($etiquettes[$g][0])</td><td class='tot' rowspan='$($groupes[$g])'>$(Format-Nutri $r.Totaux.($etiquettes[$g][1]))</td>")
                }
                if ($i -eq 0) { [void]$sb.Append("<td class='tot' rowspan='$n'>$(Format-Nutri $r.Totaux.Kcal)<br><span style='font-weight:normal;font-size:9px'>kcal</span></td>") }
                [void]$sb.Append('</tr>')
            }
            [void]$sb.Append('</tbody>')
        }
        [void]$sb.Append("</table><table class='recap'><tr><th colspan='2'>RECAP</th></tr>")
        foreach ($x in @(,@('KCAL', 'Kcal')) + $etiquettes) { [void]$sb.Append("<tr><td class='etiq'>$($x[0])</td><td>$(Format-Nutri $t.($x[1]))</td></tr>") }
        [void]$sb.Append("</table></div>")
    }
    [void]$sb.Append("</body></html>")

    $tempHtml = [System.IO.Path]::GetTempFileName() + ".html"
    [System.IO.File]::WriteAllText($tempHtml, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
    try {
        ConvertTo-PdfDepuisHtml -HtmlPath $tempHtml -PdfPath $Path
    } finally {
        Remove-Item $tempHtml -Force -ErrorAction SilentlyContinue
    }
}

function Export-PlanNutritionExcel {
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $PlanNutritionId,
        [Parameter(Mandatory)] [string] $Path
    )

    <#
        Meme presentation que l'onglet NUTRITION du fichier d'origine (et que l'export PDF) : un tableau
        "PLAN JOURNALIER" par type de jour, repas en bande lavande, totaux PRO/GLU/LIP/FIB et kcal par
        repas a droite, et un encadre RECAP du jour a cote.
    #>
    $plan = Invoke-SqliteQuery -DataSource $DbPath -Query @"
SELECT pn.nom, pn.date_debut, pn.notes, c.nom AS client_nom, c.prenom AS client_prenom
FROM plans_nutrition pn JOIN clients c ON c.id = pn.client_id WHERE pn.id = @Id
"@ -SqlParameters @{ Id = $PlanNutritionId }
    $typesJour = @(Get-TypesJour -DbPath $DbPath -PlanNutritionId $PlanNutritionId)
    $etiquettes = @(@('PRO', 'Proteines'), @('GLU', 'Glucides'), @('LIP', 'Lipides'), @('FIB', 'Fibres'))
    $blanc = '#FFFFFF'

    # Colonnes : B repas | C aliment | D quantite | E..I kcal/pro/glu/lip/fib | J..K etiquette+valeur | L kcal repas | M espace | N..O RECAP
    $cRepas = 2; $cAliment = 3; $cQte = 4; $cKcal = 5; $cEtiq = 10; $cVal = 11; $cKcalRepas = 12; $cRecap = 14
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $pkg = Open-ExcelPackage -Path $Path -Create
    try {
        $ws = Add-Worksheet -ExcelPackage $pkg -WorksheetName 'NUTRITION'
        $ws.View.ShowGridLines = $false
        $largeurs = @{ 1 = 2; 2 = 13; 3 = 28; 4 = 11; 5 = 9; 6 = 8; 7 = 8; 8 = 8; 9 = 8; 10 = 7; 11 = 8; 12 = 10; 13 = 3; 14 = 10; 15 = 11 }
        foreach ($c in $largeurs.Keys) { $ws.Column($c).Width = $largeurs[$c] }

        $titre = "PLAN NUTRITION $($plan.nom) - $($plan.client_prenom) $($plan.client_nom)".ToUpperInvariant()
        Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 1 -C1 $cRepas -L2 1 -C2 $cKcalRepas -Valeur $titre) -Fond $Script:CouleurViolet -Couleur $blanc -Taille 12 -Gras
        $ws.Row(1).Height = 24
        $morceaux = @()
        if ($plan.date_debut) { $morceaux += "Debut le $(([datetime]$plan.date_debut).ToString('dd/MM/yyyy'))" }
        if ($plan.notes) { $morceaux += [string]$plan.notes }
        Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 2 -C1 $cRepas -L2 2 -C2 $cKcalRepas -Valeur ($morceaux -join ' - ')) -Couleur $Script:CouleurViolet -Italique -Gauche
        $ws.Row(2).Height = 26

        $ligne = 4
        foreach ($tj in $typesJour) {
            $d = Get-DonneesJourNutrition -DbPath $DbPath -TypeJourId ([int]$tj.id)
            $t = $d.Totaux
            $l1 = $ligne; $l2 = $ligne + 1

            # Ligne 1 : PLAN JOURNALIER + TOTAL du jour (violet fonce)
            Set-FusionExcel -Ws $ws -L1 $l1 -C1 $cRepas -L2 $l1 -C2 $cAliment -Valeur "PLAN JOURNALIER - $(([string]$tj.nom).ToUpperInvariant())" | Out-Null
            $ws.Cells[$l1, $cQte].Value = 'TOTAL'
            $i = 0; foreach ($k in 'Kcal', 'Proteines', 'Glucides', 'Lipides', 'Fibres') { $ws.Cells[$l1, ($cKcal + $i)].Value = [double](Format-Nutri $t.$k); $i++ }
            Set-FusionExcel -Ws $ws -L1 $l1 -C1 $cEtiq -L2 $l1 -C2 $cKcalRepas -Valeur $null | Out-Null
            Set-StyleExcel -Plage $ws.Cells[$l1, $cRepas, $l1, $cKcalRepas] -Fond $Script:CouleurVioletFonce -Couleur $blanc -Gras
            # Ligne 2 : en-tetes
            $entetes = @{ $cRepas = 'REPAS'; $cAliment = 'ALIMENTS'; $cQte = 'QUANTITE'; 5 = 'KCAL'; 6 = 'PRO'; 7 = 'GLU'; 8 = 'LIP'; 9 = 'FIB' }
            foreach ($c in $entetes.Keys) { $ws.Cells[$l2, $c].Value = $entetes[$c] }
            Set-FusionExcel -Ws $ws -L1 $l2 -C1 $cEtiq -L2 $l2 -C2 $cKcalRepas -Valeur 'TOTAUX' | Out-Null
            Set-StyleExcel -Plage $ws.Cells[$l2, $cRepas, $l2, $cKcalRepas] -Fond $Script:CouleurViolet -Couleur $blanc -Gras
            Set-StyleExcel -Plage $ws.Cells[$l2, $cKcal, $l2, ($cKcal + 4)] -Fond $Script:CouleurLilas -Couleur $blanc -Gras
            $ws.Row($l1).Height = 18; $ws.Row($l2).Height = 18

            $ligne = $l2 + 1
            foreach ($r in $d.Repas) {
                $n = $r.NbLignes; $r1 = $ligne; $r2 = $ligne + $n - 1
                Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $r1 -C1 $cRepas -L2 $r2 -C2 $cRepas -Valeur ([string]$r.Nom).ToUpperInvariant()) -Fond $Script:CouleurLavande -Couleur $blanc -Gras
                for ($i = 0; $i -lt $n; $i++) {
                    $l = $r1 + $i
                    $ws.Row($l).Height = 16
                    if ($i -lt $r.Lignes.Count) {
                        $a = $r.Lignes[$i]
                        $ws.Cells[$l, $cAliment].Value = [string]$a.aliment_nom
                        $ws.Cells[$l, $cQte].Value = "$($a.quantite) $($a.unite)".Trim()
                        $j = 0; foreach ($k in 'kcal_calc', 'proteines_calc', 'glucides_calc', 'lipides_calc', 'fibres_calc') { $ws.Cells[$l, ($cKcal + $j)].Value = [double](Format-Nutri $a.$k); $j++ }
                    }
                }
                Set-StyleExcel -Plage $ws.Cells[$r1, $cAliment, $r2, $cAliment] -Couleur '#000000' -Taille 10 -Gauche
                Set-StyleExcel -Plage $ws.Cells[$r1, $cQte, $r2, ($cKcal + 4)] -Couleur '#000000'
                Set-BordureExcel -Plage $ws.Cells[$r1, $cAliment, $r2, ($cKcal + 4)] -Cotes @('Bottom') -Couleur $Script:CouleurGrisClair -Epaisseur 'Thin'
                # Totaux du repas : PRO / GLU / LIP / FIB repartis sur la hauteur du repas, kcal sur toute la hauteur
                $groupes = Get-RepartitionTotaux -NbLignes $n
                $debut = $r1
                for ($g = 0; $g -lt 4; $g++) {
                    $fin = $debut + $groupes[$g] - 1
                    Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $debut -C1 $cEtiq -L2 $fin -C2 $cEtiq -Valeur $etiquettes[$g][0]) -Fond $Script:CouleurLilas -Couleur $blanc -Gras
                    Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $debut -C1 $cVal -L2 $fin -C2 $cVal -Valeur ([double](Format-Nutri $r.Totaux.($etiquettes[$g][1])))) -Couleur '#000000' -Gras
                    Set-BordureExcel -Plage $ws.Cells[$fin, $cEtiq, $fin, $cVal] -Cotes @('Bottom') -Couleur $Script:CouleurGrisClair -Epaisseur 'Thin'
                    $debut = $fin + 1
                }
                Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $r1 -C1 $cKcalRepas -L2 $r2 -C2 $cKcalRepas -Valeur ([double](Format-Nutri $r.Totaux.Kcal))) -Couleur '#000000' -Gras -Taille 10
                Set-BordureExcel -Plage $ws.Cells[$r2, $cRepas, $r2, $cKcalRepas] -Cotes @('Bottom') -Couleur $Script:CouleurViolet
                $ligne = $r2 + 1
            }
            if ($d.Repas.Count -eq 0) {
                Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $ligne -C1 $cRepas -L2 $ligne -C2 $cKcalRepas -Valeur 'Aucun repas pour ce type de jour.') -Couleur '#999999' -Italique
                $ligne++
            }
            # Cadre du tableau
            Set-BordureExcel -Plage $ws.Cells[$l1, $cRepas, ($ligne - 1), $cRepas] -Cotes @('Left') -Couleur $Script:CouleurViolet
            Set-BordureExcel -Plage $ws.Cells[$l1, $cKcalRepas, ($ligne - 1), $cKcalRepas] -Cotes @('Right') -Couleur $Script:CouleurViolet

            # Encadre RECAP du jour
            Set-StyleExcel -Plage (Set-FusionExcel -Ws $ws -L1 $l1 -C1 $cRecap -L2 $l2 -C2 ($cRecap + 1) -Valeur 'RECAP') -Fond $Script:CouleurViolet -Couleur $blanc -Gras
            $lr = $l2 + 1
            foreach ($x in @(,@('KCAL', 'Kcal')) + $etiquettes) {
                $ws.Cells[$lr, $cRecap].Value = $x[0]
                $ws.Cells[$lr, ($cRecap + 1)].Value = [double](Format-Nutri $t.($x[1]))
                Set-StyleExcel -Plage $ws.Cells[$lr, $cRecap] -Fond $Script:CouleurLilas -Couleur $blanc -Gras
                Set-StyleExcel -Plage $ws.Cells[$lr, ($cRecap + 1)] -Couleur '#000000' -Gras -Taille 12
                Set-BordureExcel -Plage $ws.Cells[$lr, $cRecap, $lr, ($cRecap + 1)] -Cotes @('Bottom') -Couleur $Script:CouleurGrisClair -Epaisseur 'Thin'
                $lr++
            }
            Set-BordureExcel -Plage $ws.Cells[$l1, $cRecap, ($lr - 1), ($cRecap + 1)] -Cotes @('Left', 'Right') -Couleur $Script:CouleurViolet
            Set-BordureExcel -Plage $ws.Cells[($lr - 1), $cRecap, ($lr - 1), ($cRecap + 1)] -Cotes @('Bottom') -Couleur $Script:CouleurViolet

            $ligne = [math]::Max($ligne, $lr) + 2
        }
        $ws.PrinterSettings.Orientation = [OfficeOpenXml.eOrientation]::Landscape
        $ws.PrinterSettings.FitToPage = $true; $ws.PrinterSettings.FitToWidth = 1; $ws.PrinterSettings.FitToHeight = 0
    } finally {
        Close-ExcelPackage $pkg
    }
}

Export-ModuleMember -Function Find-NavigateurPdf, ConvertTo-PdfDepuisHtml, Export-ProgrammePdf, Export-ProgrammeExcel, `
    Export-FeuilleSeanceExcel, Export-PlanNutritionPdf, Export-PlanNutritionExcel, Get-ValeurAvecDetailSeries
