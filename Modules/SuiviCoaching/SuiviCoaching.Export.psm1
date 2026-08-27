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

function Get-StyleHtmlBase {
    return @"
<style>
    body { font-family: 'Segoe UI', Arial, sans-serif; color: #222; margin: 30px; }
    h1 { color: #2C3E50; margin-bottom: 0; }
    .sous-titre { color: #666; margin-top: 4px; margin-bottom: 24px; }
    h2 { color: #2C3E50; border-bottom: 2px solid #2C3E50; padding-bottom: 4px; margin-top: 28px; }
    h3 { color: #27AE60; margin-bottom: 6px; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 16px; }
    th, td { border: 1px solid #ddd; padding: 6px 10px; text-align: left; font-size: 13px; }
    th { background-color: #2C3E50; color: white; }
    tr:nth-child(even) { background-color: #F5F7FA; }
    .totaux { font-weight: bold; background-color: #E8F6EF; }
    .notes { font-style: italic; color: #555; margin-top: 4px; }
    .miniature-exercice { width: 60px; height: 60px; object-fit: cover; border-radius: 4px; display: block; }
    .lien-video { color: #27AE60; font-weight: bold; text-decoration: none; white-space: nowrap; }
</style>
"@
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

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<html><head><meta charset='utf-8'>$(Get-StyleHtmlBase)</head><body>")
    [void]$sb.Append("<h1>Programme sportif</h1>")
    [void]$sb.Append("<div class='sous-titre'>$(HtmlEncode $prog.nom) &mdash; $(HtmlEncode $prog.client_prenom) $(HtmlEncode $prog.client_nom)")
    if ($prog.date_debut) { [void]$sb.Append(" &mdash; debut le $(HtmlEncode $prog.date_debut)") }
    [void]$sb.Append("</div>")
    if ($prog.notes) { [void]$sb.Append("<p class='notes'>$(HtmlEncode $prog.notes)</p>") }

    foreach ($s in $seances) {
        [void]$sb.Append("<h2>$(HtmlEncode $s.nom)</h2>")
        $exercices = @(Get-SeanceExercices -DbPath $DbPath -SeanceId ([int]$s.id))
        if ($exercices.Count -eq 0) {
            [void]$sb.Append("<p class='notes'>Aucun exercice dans cette seance.</p>")
            continue
        }
        [void]$sb.Append("<table><tr><th>Image</th><th>Exercice</th><th>Muscle cible</th><th>Series</th><th>Repetitions</th><th>Charge</th><th>Recup</th><th>Tempo</th><th>Video</th><th>Notes</th></tr>")
        foreach ($e in $exercices) {
            $recup = if ($e.recuperation_s) { "$($e.recuperation_s) s" } else { "" }
            $seriesTexte = if ($e.series) { [string]$e.series } else { '' }
            $imageTd = ''
            if ($e.image_path) {
                $dataUri = Get-ImageDataUri -Path (Join-Path $dossierData $e.image_path)
                if ($dataUri) { $imageTd = "<img class='miniature-exercice' src='$dataUri' alt='' />" }
            }
            $videoTd = ''
            if ($e.lien_video) {
                $lienEncode = HtmlEncode $e.lien_video
                $videoTd = "<a class='lien-video' href='$lienEncode'>&#9654; Video</a>"
            }
            [void]$sb.Append("<tr><td>$imageTd</td><td>$(HtmlEncode $e.exercice_nom)</td><td>$(HtmlEncode $e.muscle_cible)</td><td>$(HtmlEncode $seriesTexte)</td><td>$(HtmlEncode $e.repetitions)</td><td>$(HtmlEncode $e.charge)</td><td>$(HtmlEncode $recup)</td><td>$(HtmlEncode $e.tempo)</td><td>$videoTd</td><td>$(HtmlEncode $e.notes)</td></tr>")
        }
        [void]$sb.Append("</table>")
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

    $seances = @(Get-Seances -DbPath $DbPath -ProgrammeId $ProgrammeId)
    $lignes = New-Object System.Collections.Generic.List[object]
    foreach ($s in $seances) {
        $exercices = @(Get-SeanceExercices -DbPath $DbPath -SeanceId ([int]$s.id))
        foreach ($e in $exercices) {
            $lignes.Add([pscustomobject]@{
                Seance = $s.nom
                Exercice = $e.exercice_nom
                'Muscle cible' = $e.muscle_cible
                Series = $e.series
                Repetitions = $e.repetitions
                Charge = $e.charge
                'Recup (s)' = $e.recuperation_s
                Tempo = $e.tempo
                'Lien video' = $e.lien_video
                Notes = $e.notes
            })
        }
    }
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $lignes | Export-Excel -Path $Path -WorksheetName 'Programme' -AutoSize -TableStyle Medium2 -FreezeTopRow
}

function Export-FeuilleSeanceExcel {
    <#
        Genere un fichier Excel a remplir par le client : une ligne par exercice de chaque seance du
        programme, avec les valeurs prevues en reference et des colonnes vides a completer (date de
        realisation, series/repetitions/recup/tempo reellement faits, note). A reimporter ensuite via
        Import-SeanceRealiseeDepuisExcel.
    #>
    param(
        [Parameter(Mandatory)] [string] $DbPath,
        [Parameter(Mandatory)] [int] $ProgrammeId,
        [Parameter(Mandatory)] [string] $Path
    )

    $seances = @(Get-Seances -DbPath $DbPath -ProgrammeId $ProgrammeId)
    $lignes = New-Object System.Collections.Generic.List[object]
    foreach ($s in $seances) {
        $exercices = @(Get-SeanceExercices -DbPath $DbPath -SeanceId ([int]$s.id))
        foreach ($e in $exercices) {
            $lignes.Add([pscustomobject]@{
                SeanceId = $s.id
                Seance = $s.nom
                SeanceExerciceId = $e.id
                Exercice = $e.exercice_nom
                'Series prevues' = $e.series
                'Repetitions prevues' = $e.repetitions
                'Charge prevue' = $e.charge
                'Recup prevue (s)' = $e.recuperation_s
                'Tempo prevu' = $e.tempo
                'Date de realisation' = $null
                'Series realisees' = $null
                'Repetitions realisees' = $null
                'Charge realisee' = $null
                'Recup realisee (s)' = $null
                'Tempo realise' = $null
                Notes = $null
            })
        }
    }
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $lignes | Export-Excel -Path $Path -WorksheetName 'Feuille de seance' -AutoSize -TableStyle Medium2 -FreezeTopRow
}

# ================= NUTRITION =================

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

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<html><head><meta charset='utf-8'>$(Get-StyleHtmlBase)</head><body>")
    [void]$sb.Append("<h1>Plan nutritionnel</h1>")
    [void]$sb.Append("<div class='sous-titre'>$(HtmlEncode $plan.nom) &mdash; $(HtmlEncode $plan.client_prenom) $(HtmlEncode $plan.client_nom)")
    if ($plan.date_debut) { [void]$sb.Append(" &mdash; debut le $(HtmlEncode $plan.date_debut)") }
    [void]$sb.Append("</div>")
    if ($plan.notes) { [void]$sb.Append("<p class='notes'>$(HtmlEncode $plan.notes)</p>") }

    foreach ($tj in $typesJour) {
        [void]$sb.Append("<h2>$(HtmlEncode $tj.nom)</h2>")
        $repasListe = @(Get-Repas -DbPath $DbPath -TypeJourId ([int]$tj.id))
        $totalJour = [pscustomobject]@{ Kcal = 0.0; Proteines = 0.0; Glucides = 0.0; Lipides = 0.0; Fibres = 0.0 }

        foreach ($r in $repasListe) {
            [void]$sb.Append("<h3>$(HtmlEncode $r.nom)</h3>")
            $lignes = @(Get-RepasAliments -DbPath $DbPath -RepasId ([int]$r.id))
            if ($lignes.Count -eq 0) {
                [void]$sb.Append("<p class='notes'>Aucun aliment dans ce repas.</p>")
                continue
            }
            [void]$sb.Append("<table><tr><th>Aliment</th><th>Quantite</th><th>Kcal</th><th>Proteines</th><th>Glucides</th><th>Lipides</th><th>Fibres</th></tr>")
            $totalRepas = Get-TotauxRepas -Lignes $lignes
            foreach ($l in $lignes) {
                [void]$sb.Append("<tr><td>$(HtmlEncode $l.aliment_nom)</td><td>$($l.quantite) $(HtmlEncode $l.unite)</td><td>$($l.kcal_calc)</td><td>$($l.proteines_calc)</td><td>$($l.glucides_calc)</td><td>$($l.lipides_calc)</td><td>$($l.fibres_calc)</td></tr>")
            }
            [void]$sb.Append("<tr class='totaux'><td>Total repas</td><td></td><td>$($totalRepas.Kcal)</td><td>$($totalRepas.Proteines)</td><td>$($totalRepas.Glucides)</td><td>$($totalRepas.Lipides)</td><td>$($totalRepas.Fibres)</td></tr>")
            [void]$sb.Append("</table>")
            $totalJour.Kcal += $totalRepas.Kcal; $totalJour.Proteines += $totalRepas.Proteines
            $totalJour.Glucides += $totalRepas.Glucides; $totalJour.Lipides += $totalRepas.Lipides; $totalJour.Fibres += $totalRepas.Fibres
        }

        [void]$sb.Append("<table><tr class='totaux'><td>TOTAL JOURNEE</td><td>$($totalJour.Kcal) kcal</td><td>Proteines : $($totalJour.Proteines) g</td><td>Glucides : $($totalJour.Glucides) g</td><td>Lipides : $($totalJour.Lipides) g</td><td>Fibres : $($totalJour.Fibres) g</td></tr></table>")
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

    $typesJour = @(Get-TypesJour -DbPath $DbPath -PlanNutritionId $PlanNutritionId)
    $lignes = New-Object System.Collections.Generic.List[object]
    foreach ($tj in $typesJour) {
        $repasListe = @(Get-Repas -DbPath $DbPath -TypeJourId ([int]$tj.id))
        foreach ($r in $repasListe) {
            $alimentsLignes = @(Get-RepasAliments -DbPath $DbPath -RepasId ([int]$r.id))
            foreach ($l in $alimentsLignes) {
                $lignes.Add([pscustomobject]@{
                    'Type de jour' = $tj.nom
                    Repas = $r.nom
                    Aliment = $l.aliment_nom
                    Quantite = $l.quantite
                    Unite = $l.unite
                    Kcal = $l.kcal_calc
                    Proteines = $l.proteines_calc
                    Glucides = $l.glucides_calc
                    Lipides = $l.lipides_calc
                    Fibres = $l.fibres_calc
                })
            }
        }
    }
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $lignes | Export-Excel -Path $Path -WorksheetName 'Plan nutrition' -AutoSize -TableStyle Medium2 -FreezeTopRow
}

Export-ModuleMember -Function Find-NavigateurPdf, ConvertTo-PdfDepuisHtml, Export-ProgrammePdf, Export-ProgrammeExcel, `
    Export-FeuilleSeanceExcel, Export-PlanNutritionPdf, Export-PlanNutritionExcel
