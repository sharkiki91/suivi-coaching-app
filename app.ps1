Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$AppVersion = '1.8.0'
$AppRoot = $PSScriptRoot
$DbPath = Join-Path $AppRoot 'Data\suivi_coaching.db'
$BackupFolder = Join-Path $AppRoot 'Data\Backups'

function Show-ErreurDemarrage {
    <# Utilise avant que PresentationFramework soit forcement charge : repose sur Windows Forms, quasi toujours disponible. #>
    param([string]$Message)
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show($Message, 'Suivi Coaching - Erreur au demarrage', 'OK', 'Error') | Out-Null
    } catch {
        # Si meme Windows Forms est indisponible, on tente au moins un affichage console.
        Write-Host $Message
    }
}

try {
    # Quand l'appli est telechargee depuis internet (zip GitHub), Windows marque tous les fichiers extraits
    # comme "provenant d'internet" et peut bloquer silencieusement les modules/DLL. On les debloque a chaque
    # lancement : operation rapide et sans effet si les fichiers sont deja debloques.
    Get-ChildItem -Path $AppRoot -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

    Import-Module (Join-Path $AppRoot 'Modules\PSSQLite\1.1.0\PSSQLite.psd1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\ImportExcel\7.8.10\ImportExcel.psd1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.Database.psm1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.Clients.psm1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.Administratif.psm1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.Bibliotheques.psm1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.Import.psm1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.Programmes.psm1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.ModelesSeance.psm1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.Nutrition.psm1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.Export.psm1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.Suivi.psm1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.Questionnaires.psm1') -Force
    Import-Module (Join-Path $AppRoot 'Modules\SuiviCoaching\SuiviCoaching.Dashboard.psm1') -Force
} catch {
    Show-ErreurDemarrage "L'application n'a pas pu demarrer :`n`n$($_.Exception.Message)`n`nSi tu viens de telecharger l'application, verifie qu'aucun antivirus ne bloque le dossier, ou essaie de re-extraire le fichier zip apres avoir clique droit dessus > Proprietes > Debloquer."
    exit 1
}

function Show-Info { param([string]$Message, [string]$Titre = 'Suivi Coaching')
    [System.Windows.MessageBox]::Show($Message, $Titre, 'OK', 'Information') | Out-Null
}
function Show-Erreur { param([string]$Message, [string]$Titre = 'Erreur')
    [System.Windows.MessageBox]::Show($Message, $Titre, 'OK', 'Error') | Out-Null
}
function Show-Confirmation { param([string]$Message, [string]$Titre = 'Confirmation')
    $reponse = [System.Windows.MessageBox]::Show($Message, $Titre, 'YesNo', 'Question')
    return $reponse -eq 'Yes'
}
function Invoke-Protege {
    <# Execute un bloc et affiche une boite d'erreur lisible en cas d'echec, sans planter l'appli. #>
    param([Parameter(Mandatory)] [scriptblock] $Bloc)
    try {
        & $Bloc
    } catch {
        Show-Erreur "Une erreur est survenue :`n`n$($_.Exception.Message)"
    }
}
function Get-TexteOuNull { param([string]$Texte)
    if ([string]::IsNullOrWhiteSpace($Texte)) { return $null }
    return $Texte.Trim()
}
function Get-DoubleOuNull { param([string]$Texte)
    if ([string]::IsNullOrWhiteSpace($Texte)) { return $null }
    $valeur = 0.0
    if ([double]::TryParse($Texte.Replace(',', '.'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$valeur)) {
        return $valeur
    }
    throw "'$Texte' n'est pas un nombre valide."
}
function Get-IntOuNull { param([string]$Texte)
    if ([string]::IsNullOrWhiteSpace($Texte)) { return $null }
    $valeur = 0
    if ([int]::TryParse($Texte, [ref]$valeur)) { return $valeur }
    throw "'$Texte' n'est pas un nombre entier valide."
}

$estPremierLancement = -not (Test-Path $DbPath)

try {
    Initialize-Database -DbPath $DbPath
} catch {
    Show-Erreur "Impossible d'initialiser la base de donnees :`n$($_.Exception.Message)"
    exit 1
}

if ($estPremierLancement) {
    $fichierBibliothequesInitiales = Join-Path $AppRoot 'DonneesInitiales\Bibliotheques-initiales.xlsx'
    if (Test-Path $fichierBibliothequesInitiales) {
        try {
            $resSeed = Import-BibliothequesDepuisExcel -DbPath $DbPath -ExcelPath $fichierBibliothequesInitiales
            if ($resSeed.ExercicesImportes -gt 0 -or $resSeed.AlimentsImportes -gt 0 -or $resSeed.ComplementsImportes -gt 0) {
                Show-Info "Premier lancement : bibliotheques initiales chargees.`n`nExercices : $($resSeed.ExercicesImportes)`nAliments : $($resSeed.AlimentsImportes)`nComplements : $($resSeed.ComplementsImportes)"
            }
        } catch {
            # Non bloquant : l'appli reste utilisable meme si le chargement initial echoue (les bibliotheques resteront simplement vides).
        }
    }
}

function Import-XamlWindow {
    <# Charge une fenetre XAML en forcant l'UTF-8, quel que soit le BOM du fichier. #>
    param([Parameter(Mandatory)] [string] $Path)
    $texte = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    [xml]$xamlDoc = $texte
    $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
    [System.Windows.Markup.XamlReader]::Load($reader)
}

try {
    $xamlPath = Join-Path $AppRoot 'UI\MainWindow.xaml'
    $Window = Import-XamlWindow -Path $xamlPath
    $Window.Title = "Suivi Coaching v$AppVersion"
} catch {
    Show-Erreur "Impossible de charger l'interface :`n`n$($_.Exception.Message)"
    exit 1
}

function Get-Ctrl { param([string]$Name) $Window.FindName($Name) }

function Show-DialogNouvelElement {
    <#
        Ouvre la boite de dialogue de creation (programme / plan nutrition).
        Retourne un hashtable @{ Nom; DateDebut; Notes } ou $null si annule.
    #>
    param([Parameter(Mandatory)] [string] $Titre)

    $dlg = Import-XamlWindow -Path (Join-Path $AppRoot 'UI\DialogNouvelElement.xaml')
    $dlg.Owner = $Window
    $dlg.FindName('TxtTitreDialog').Text = $Titre
    $txtNom = $dlg.FindName('TxtDialogNom')
    $dateDebut = $dlg.FindName('DateDialogDebut')
    $txtNotes = $dlg.FindName('TxtDialogNotes')

    $Script:ResultatDialog = $null
    $dlg.FindName('BtnDialogAnnuler').Add_Click({ $dlg.DialogResult = $false })
    $dlg.FindName('BtnDialogValider').Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtNom.Text)) {
            Show-Erreur "Le nom est obligatoire."
            return
        }
        $Script:ResultatDialog = @{
            Nom = $txtNom.Text.Trim()
            DateDebut = if ($dateDebut.SelectedDate) { $dateDebut.SelectedDate.ToString('yyyy-MM-dd') } else { $null }
            Notes = if ($txtNotes.Text) { $txtNotes.Text.Trim() } else { $null }
        }
        $dlg.DialogResult = $true
    })

    if ($dlg.ShowDialog()) { return $Script:ResultatDialog }
    return $null
}

function Show-DialogChoixModele {
    <#
        Ouvre la boite de dialogue de creation d'une seance a partir d'un modele.
        Retourne un hashtable @{ SeanceModeleId; Nom } ou $null si annule.
    #>
    param([Parameter(Mandatory)] [array] $Modeles)

    $dlg = Import-XamlWindow -Path (Join-Path $AppRoot 'UI\DialogChoixModele.xaml')
    $dlg.Owner = $Window
    $cmbModele = $dlg.FindName('CmbDialogModele')
    $txtNom = $dlg.FindName('TxtDialogNomSeance')
    $cmbModele.ItemsSource = $Modeles
    if ($Modeles.Count -gt 0) { $cmbModele.SelectedIndex = 0 }
    $cmbModele.Add_SelectionChanged({
        if ($cmbModele.SelectedItem) { $txtNom.Text = [string]$cmbModele.SelectedItem.nom }
    })
    if ($Modeles.Count -gt 0) { $txtNom.Text = [string]$Modeles[0].nom }

    $Script:ResultatDialog = $null
    $dlg.FindName('BtnDialogAnnuler').Add_Click({ $dlg.DialogResult = $false })
    $dlg.FindName('BtnDialogValider').Add_Click({
        if (-not $cmbModele.SelectedItem) { Show-Erreur "Selectionne un modele."; return }
        if ([string]::IsNullOrWhiteSpace($txtNom.Text)) { Show-Erreur "Le nom de la seance est obligatoire."; return }
        $Script:ResultatDialog = @{
            SeanceModeleId = $cmbModele.SelectedItem.id
            Nom = $txtNom.Text.Trim()
        }
        $dlg.DialogResult = $true
    })

    if ($dlg.ShowDialog()) { return $Script:ResultatDialog }
    return $null
}

function Show-DialogImportQuestionnaire {
    <#
        Ouvre la boite de dialogue de mapping des colonnes pour l'import d'un questionnaire.
        Retourne un hashtable @{ Type; ColonneDate; ColonneNom; ColonneEmail; ColonneTelephone;
        ColonneObjectifLongTerme; ColonneObjectifCourtTerme; ColonneMoyens } ou $null si annule.
    #>
    param([Parameter(Mandatory)] [string[]] $Entetes)

    $dlg = Import-XamlWindow -Path (Join-Path $AppRoot 'UI\DialogImportQuestionnaire.xaml')
    $dlg.Owner = $Window
    $cmbType = $dlg.FindName('CmbTypeQuestionnaire')
    $cmbDate = $dlg.FindName('CmbColonneDate')
    $cmbNom = $dlg.FindName('CmbColonneNom')
    $cmbEmail = $dlg.FindName('CmbColonneEmail')
    $panelPreCoaching = $dlg.FindName('PanelChampsPreCoaching')
    $cmbTelephone = $dlg.FindName('CmbColonneTelephone')
    $cmbObjectifLong = $dlg.FindName('CmbColonneObjectifLongTerme')
    $cmbObjectifCourt = $dlg.FindName('CmbColonneObjectifCourtTerme')
    $cmbMoyens = $dlg.FindName('CmbColonneMoyens')
    $combosOptionnelsChampsClient = @($cmbTelephone, $cmbObjectifLong, $cmbObjectifCourt, $cmbMoyens)

    $optionsAvecVide = @('(aucune)') + $Entetes
    $cmbDate.ItemsSource = $Entetes
    $cmbNom.ItemsSource = $optionsAvecVide
    $cmbEmail.ItemsSource = $optionsAvecVide
    $cmbNom.SelectedIndex = 0
    $cmbEmail.SelectedIndex = 0
    foreach ($cmb in $combosOptionnelsChampsClient) { $cmb.ItemsSource = $optionsAvecVide; $cmb.SelectedIndex = 0 }

    $appliquerMapping = {
        $typeItem = $cmbType.SelectedItem
        if (-not $typeItem) { return }
        $panelPreCoaching.Visibility = if ($typeItem.Tag -eq 'pre_coaching') { 'Visible' } else { 'Collapsed' }
        $mapping = Get-ImportMapping -DbPath $DbPath -Type $typeItem.Tag
        if (-not $mapping) { return }
        if ($mapping.ColonneDate -and ($Entetes -contains $mapping.ColonneDate)) { $cmbDate.SelectedItem = $mapping.ColonneDate }
        if ($mapping.ColonneNom -and ($Entetes -contains $mapping.ColonneNom)) { $cmbNom.SelectedItem = $mapping.ColonneNom }
        if ($mapping.ColonneEmail -and ($Entetes -contains $mapping.ColonneEmail)) { $cmbEmail.SelectedItem = $mapping.ColonneEmail }
        if ($mapping.ColonneTelephone -and ($Entetes -contains $mapping.ColonneTelephone)) { $cmbTelephone.SelectedItem = $mapping.ColonneTelephone }
        if ($mapping.ColonneObjectifLongTerme -and ($Entetes -contains $mapping.ColonneObjectifLongTerme)) { $cmbObjectifLong.SelectedItem = $mapping.ColonneObjectifLongTerme }
        if ($mapping.ColonneObjectifCourtTerme -and ($Entetes -contains $mapping.ColonneObjectifCourtTerme)) { $cmbObjectifCourt.SelectedItem = $mapping.ColonneObjectifCourtTerme }
        if ($mapping.ColonneMoyens -and ($Entetes -contains $mapping.ColonneMoyens)) { $cmbMoyens.SelectedItem = $mapping.ColonneMoyens }
    }
    $cmbType.Add_SelectionChanged($appliquerMapping)
    $cmbType.SelectedIndex = 0

    $Script:ResultatDialog = $null
    $dlg.FindName('BtnDialogAnnuler').Add_Click({ $dlg.DialogResult = $false })
    $dlg.FindName('BtnDialogValider').Add_Click({
        if (-not $cmbType.SelectedItem) { Show-Erreur "Sélectionne un type de questionnaire."; return }
        if (-not $cmbDate.SelectedItem) { Show-Erreur "Sélectionne la colonne contenant la date."; return }
        function Get-ValeurComboOuNull($cmb) {
            if ($cmb.SelectedItem -and $cmb.SelectedItem -ne '(aucune)') { return [string]$cmb.SelectedItem }
            return $null
        }
        $Script:ResultatDialog = @{
            Type = [string]$cmbType.SelectedItem.Tag
            ColonneDate = [string]$cmbDate.SelectedItem
            ColonneNom = Get-ValeurComboOuNull $cmbNom
            ColonneEmail = Get-ValeurComboOuNull $cmbEmail
            ColonneTelephone = Get-ValeurComboOuNull $cmbTelephone
            ColonneObjectifLongTerme = Get-ValeurComboOuNull $cmbObjectifLong
            ColonneObjectifCourtTerme = Get-ValeurComboOuNull $cmbObjectifCourt
            ColonneMoyens = Get-ValeurComboOuNull $cmbMoyens
        }
        $dlg.DialogResult = $true
    })

    if ($dlg.ShowDialog()) { return $Script:ResultatDialog }
    return $null
}

# Navigation
$PanelDashboard = Get-Ctrl 'PanelDashboard'
$PanelClients = Get-Ctrl 'PanelClients'
$PanelAdministratif = Get-Ctrl 'PanelAdministratif'
$PanelBibliotheques = Get-Ctrl 'PanelBibliotheques'
$PanelProgrammes = Get-Ctrl 'PanelProgrammes'
$PanelNutrition = Get-Ctrl 'PanelNutrition'
$PanelSuivi = Get-Ctrl 'PanelSuivi'
$PanelOutils = Get-Ctrl 'PanelOutils'
$AllPanels = @($PanelDashboard, $PanelClients, $PanelAdministratif, $PanelBibliotheques, $PanelProgrammes, $PanelNutrition, $PanelSuivi, $PanelOutils)

function Show-Panel {
    param($Panel)
    foreach ($p in $AllPanels) { $p.Visibility = 'Collapsed' }
    $Panel.Visibility = 'Visible'
}

(Get-Ctrl 'BtnNavDashboard').Add_Click({ Show-Panel $PanelDashboard; Update-VueDashboard })
(Get-Ctrl 'BtnNavClients').Add_Click({ Show-Panel $PanelClients; Update-VueClients })
(Get-Ctrl 'BtnNavAdministratif').Add_Click({ Show-Panel $PanelAdministratif; Update-VueAdministratif })
(Get-Ctrl 'BtnNavBibliotheques').Add_Click({ Show-Panel $PanelBibliotheques; Update-VueBibliotheques })
(Get-Ctrl 'BtnNavProgrammes').Add_Click({ Show-Panel $PanelProgrammes; Update-VueProgrammesClients })
(Get-Ctrl 'BtnNavNutrition').Add_Click({ Show-Panel $PanelNutrition; Update-VueNutritionClients })
(Get-Ctrl 'BtnNavSuivi').Add_Click({ Show-Panel $PanelSuivi; Update-VueSuiviClients })
(Get-Ctrl 'BtnNavOutils').Add_Click({ Show-Panel $PanelOutils })

# ============================== CLIENTS ==============================

$GridClients = Get-Ctrl 'GridClients'
$ChkAfficherArchives = Get-Ctrl 'ChkAfficherArchives'
$TxtClientNom = Get-Ctrl 'TxtClientNom'
$TxtClientPrenom = Get-Ctrl 'TxtClientPrenom'
$TxtClientEmail = Get-Ctrl 'TxtClientEmail'
$TxtClientTelephone = Get-Ctrl 'TxtClientTelephone'
$DateClientDebut = Get-Ctrl 'DateClientDebut'
$TxtClientObjectifs = Get-Ctrl 'TxtClientObjectifs'
$TxtClientNotes = Get-Ctrl 'TxtClientNotes'
$Script:SelectedClientId = $null

function Update-VueClients {
    $GridClients.ItemsSource = @(Get-Clients -DbPath $DbPath -InclureArchives:$ChkAfficherArchives.IsChecked)
    Update-CombosClients
}

function Clear-FormClient {
    $Script:SelectedClientId = $null
    $TxtClientNom.Text = ''
    $TxtClientPrenom.Text = ''
    $TxtClientEmail.Text = ''
    $TxtClientTelephone.Text = ''
    $DateClientDebut.SelectedDate = $null
    $TxtClientObjectifs.Text = ''
    $TxtClientNotes.Text = ''
    $GridClients.SelectedItem = $null
}

$GridClients.Add_SelectionChanged({
    $item = $GridClients.SelectedItem
    if ($null -eq $item) { return }
    $Script:SelectedClientId = [int]$item.id
    $TxtClientNom.Text = [string]$item.nom
    $TxtClientPrenom.Text = [string]$item.prenom
    $TxtClientEmail.Text = [string]$item.email
    $TxtClientTelephone.Text = [string]$item.telephone
    $TxtClientObjectifs.Text = [string]$item.objectifs
    $TxtClientNotes.Text = [string]$item.notes
    $DateClientDebut.SelectedDate = if ($item.date_debut_coaching) { [datetime]$item.date_debut_coaching } else { $null }
})

(Get-Ctrl 'BtnClientNouveau').Add_Click({ Clear-FormClient })

(Get-Ctrl 'BtnClientEnregistrer').Add_Click({
    Invoke-Protege {
        if ([string]::IsNullOrWhiteSpace($TxtClientNom.Text) -or [string]::IsNullOrWhiteSpace($TxtClientPrenom.Text)) {
            Show-Erreur "Le nom et le prenom sont obligatoires."
            return
        }
        $dateDebut = if ($DateClientDebut.SelectedDate) { $DateClientDebut.SelectedDate.ToString('yyyy-MM-dd') } else { $null }
        if ($Script:SelectedClientId) {
            Update-Client -DbPath $DbPath -Id $Script:SelectedClientId -Nom $TxtClientNom.Text.Trim() -Prenom $TxtClientPrenom.Text.Trim() `
                -Email (Get-TexteOuNull $TxtClientEmail.Text) -Telephone (Get-TexteOuNull $TxtClientTelephone.Text) `
                -DateDebutCoaching $dateDebut -Objectifs (Get-TexteOuNull $TxtClientObjectifs.Text) -Notes (Get-TexteOuNull $TxtClientNotes.Text)
        } else {
            $Script:SelectedClientId = New-Client -DbPath $DbPath -Nom $TxtClientNom.Text.Trim() -Prenom $TxtClientPrenom.Text.Trim() `
                -Email (Get-TexteOuNull $TxtClientEmail.Text) -Telephone (Get-TexteOuNull $TxtClientTelephone.Text) `
                -DateDebutCoaching $dateDebut -Objectifs (Get-TexteOuNull $TxtClientObjectifs.Text) -Notes (Get-TexteOuNull $TxtClientNotes.Text)
        }
        Update-VueClients
        Show-Info "Fiche client enregistree."
    }
})

(Get-Ctrl 'BtnClientArchiver').Add_Click({
    Invoke-Protege {
        if (-not $Script:SelectedClientId) { Show-Erreur "Selectionne d'abord un client."; return }
        if (Show-Confirmation "Archiver ce client ? Il n'apparaitra plus dans la liste active.") {
            Set-ClientStatut -DbPath $DbPath -Id $Script:SelectedClientId -Statut 'archive'
            Clear-FormClient
            Update-VueClients
        }
    }
})

$ChkAfficherArchives.Add_Click({ Update-VueClients })

# ============================== ADMINISTRATIF ==============================

$GridDevis = Get-Ctrl 'GridDevis'
$CmbDevisClient = Get-Ctrl 'CmbDevisClient'
$TxtDevisPrestations = Get-Ctrl 'TxtDevisPrestations'
$TxtDevisMontant = Get-Ctrl 'TxtDevisMontant'
$TxtDevisDuree = Get-Ctrl 'TxtDevisDuree'
$CmbDevisFiltre = Get-Ctrl 'CmbDevisFiltre'

$GridCommandes = Get-Ctrl 'GridCommandes'
$CmbCommandeClient = Get-Ctrl 'CmbCommandeClient'
$CmbCommandeType = Get-Ctrl 'CmbCommandeType'
$DateCommandeDebut = Get-Ctrl 'DateCommandeDebut'
$DateCommandeFin = Get-Ctrl 'DateCommandeFin'
$TxtCommandeMontant = Get-Ctrl 'TxtCommandeMontant'
$CmbCommandeFiltre = Get-Ctrl 'CmbCommandeFiltre'

$GridEcheances = Get-Ctrl 'GridEcheances'
$CmbEcheanceFiltre = Get-Ctrl 'CmbEcheanceFiltre'

function Update-CombosClients {
    $clients = @(Get-Clients -DbPath $DbPath | ForEach-Object {
        [pscustomobject]@{ id = $_.id; affichage = "$($_.nom) $($_.prenom)" }
    })
    $CmbDevisClient.ItemsSource = $clients
    $CmbCommandeClient.ItemsSource = $clients
}

function Update-VueAdministratif {
    Update-EcheancesRetard -DbPath $DbPath
    Update-CommandesTermineesAuto -DbPath $DbPath
    Update-VueDevis
    Update-VueCommandes
    Update-VueEcheances
}

function Update-VueDevis {
    $tous = @(Get-Devis -DbPath $DbPath)
    $filtreItem = $CmbDevisFiltre.SelectedItem
    $filtre = if ($filtreItem) { $filtreItem.Content } else { 'Tous' }
    $mapStatuts = @{ 'En attente' = 'en_attente'; 'Accepté' = 'accepte'; 'Refusé' = 'refuse' }
    if ($mapStatuts.ContainsKey($filtre)) {
        $tous = $tous | Where-Object { $_.statut -eq $mapStatuts[$filtre] }
    }
    $GridDevis.ItemsSource = @($tous)
}

function Update-VueCommandes {
    $toutes = @(Get-Commandes -DbPath $DbPath)
    $filtreItem = $CmbCommandeFiltre.SelectedItem
    $filtre = if ($filtreItem) { $filtreItem.Content } else { 'Toutes' }
    $mapStatuts = @{ 'Active' = 'active'; 'Terminée' = 'terminee'; 'Annulée' = 'annulee' }
    if ($mapStatuts.ContainsKey($filtre)) {
        $toutes = $toutes | Where-Object { $_.statut -eq $mapStatuts[$filtre] }
    }
    $GridCommandes.ItemsSource = @($toutes)
}

function Update-VueEcheances {
    $toutes = @(Get-Echeances -DbPath $DbPath)
    $filtreItem = $CmbEcheanceFiltre.SelectedItem
    $filtre = if ($filtreItem) { $filtreItem.Content } else { 'Toutes' }
    $mapStatuts = @{ 'En attente' = 'en_attente'; 'En retard' = 'en_retard'; 'Payées' = 'payee' }
    if ($mapStatuts.ContainsKey($filtre)) {
        $toutes = $toutes | Where-Object { $_.statut -eq $mapStatuts[$filtre] }
    }
    $GridEcheances.ItemsSource = @($toutes)
}

$CmbDevisFiltre.Add_SelectionChanged({ Update-VueDevis })
$CmbCommandeFiltre.Add_SelectionChanged({ Update-VueCommandes })

(Get-Ctrl 'BtnDevisCreer').Add_Click({
    Invoke-Protege {
        if (-not $CmbDevisClient.SelectedItem) { Show-Erreur "Selectionne un client."; return }
        if ([string]::IsNullOrWhiteSpace($TxtDevisPrestations.Text)) { Show-Erreur "Les prestations sont obligatoires."; return }
        $montant = Get-DoubleOuNull $TxtDevisMontant.Text
        if (-not $montant) { Show-Erreur "Indique un montant valide."; return }
        New-Devis -DbPath $DbPath -ClientId $CmbDevisClient.SelectedItem.id -Prestations $TxtDevisPrestations.Text.Trim() `
            -Montant $montant -Duree (Get-TexteOuNull $TxtDevisDuree.Text) | Out-Null
        $TxtDevisPrestations.Text = ''; $TxtDevisMontant.Text = ''; $TxtDevisDuree.Text = ''
        Update-VueAdministratif
        Show-Info "Devis cree."
    }
})

(Get-Ctrl 'BtnDevisAccepter').Add_Click({
    Invoke-Protege {
        $item = $GridDevis.SelectedItem
        if (-not $item) { Show-Erreur "Selectionne un devis."; return }
        Set-DevisStatut -DbPath $DbPath -Id ([int]$item.id) -Statut 'accepte'
        Update-VueAdministratif
    }
})
(Get-Ctrl 'BtnDevisRefuser').Add_Click({
    Invoke-Protege {
        $item = $GridDevis.SelectedItem
        if (-not $item) { Show-Erreur "Selectionne un devis."; return }
        Set-DevisStatut -DbPath $DbPath -Id ([int]$item.id) -Statut 'refuse'
        Update-VueAdministratif
    }
})
(Get-Ctrl 'BtnDevisTransformer').Add_Click({
    Invoke-Protege {
        $item = $GridDevis.SelectedItem
        if (-not $item) { Show-Erreur "Selectionne un devis."; return }
        if ($item.statut -ne 'accepte') { Show-Erreur "Seul un devis accepte peut etre transforme en commande."; return }
        foreach ($c in $CmbCommandeClient.ItemsSource) { if ($c.id -eq $item.client_id) { $CmbCommandeClient.SelectedItem = $c } }
        $TxtCommandeMontant.Text = [string]$item.montant
        $Script:DevisIdPourCommande = [int]$item.id
        (Get-Ctrl 'TabAdministratif').SelectedIndex = 1
        Show-Info "Choisis le type de facturation et les dates, puis clique sur 'Creer la commande'."
    }
})

(Get-Ctrl 'BtnCommandeCreer').Add_Click({
    Invoke-Protege {
        if (-not $CmbCommandeClient.SelectedItem) { Show-Erreur "Selectionne un client."; return }
        if (-not $CmbCommandeType.SelectedItem) { Show-Erreur "Selectionne un type de facturation."; return }
        if (-not $DateCommandeDebut.SelectedDate) { Show-Erreur "Indique une date de debut."; return }
        $type = $CmbCommandeType.SelectedItem.Content
        $montant = Get-DoubleOuNull $TxtCommandeMontant.Text
        if (-not $montant) { Show-Erreur "Indique un montant valide."; return }
        if ($type -ne 'one_shot' -and -not $DateCommandeFin.SelectedDate) {
            Show-Erreur "Une date de fin est necessaire pour une facturation mensuelle ou hebdomadaire."; return
        }

        $params = @{
            DbPath = $DbPath
            ClientId = $CmbCommandeClient.SelectedItem.id
            TypeFacturation = $type
            DateDebut = $DateCommandeDebut.SelectedDate
            Montant = $montant
        }
        if ($DateCommandeFin.SelectedDate) { $params.DateFin = $DateCommandeFin.SelectedDate }
        if ($Script:DevisIdPourCommande) { $params.DevisId = $Script:DevisIdPourCommande; $Script:DevisIdPourCommande = $null }

        New-Commande @params | Out-Null
        $TxtCommandeMontant.Text = ''; $DateCommandeDebut.SelectedDate = $null; $DateCommandeFin.SelectedDate = $null
        Update-VueAdministratif
        Show-Info "Commande creee et echeancier genere."
    }
})

(Get-Ctrl 'BtnCommandeTerminer').Add_Click({
    Invoke-Protege {
        $item = $GridCommandes.SelectedItem
        if (-not $item) { Show-Erreur "Selectionne une commande."; return }
        Set-CommandeStatut -DbPath $DbPath -Id ([int]$item.id) -Statut 'terminee'
        Update-VueAdministratif
    }
})
(Get-Ctrl 'BtnCommandeAnnuler').Add_Click({
    Invoke-Protege {
        $item = $GridCommandes.SelectedItem
        if (-not $item) { Show-Erreur "Selectionne une commande."; return }
        Set-CommandeStatut -DbPath $DbPath -Id ([int]$item.id) -Statut 'annulee'
        Update-VueAdministratif
    }
})
(Get-Ctrl 'BtnCommandeReactiver').Add_Click({
    Invoke-Protege {
        $item = $GridCommandes.SelectedItem
        if (-not $item) { Show-Erreur "Selectionne une commande."; return }
        Set-CommandeStatut -DbPath $DbPath -Id ([int]$item.id) -Statut 'active'
        Update-VueAdministratif
    }
})

$CmbEcheanceFiltre.Add_SelectionChanged({ Update-VueEcheances })

(Get-Ctrl 'BtnEcheanceMarquerPayee').Add_Click({
    Invoke-Protege {
        $item = $GridEcheances.SelectedItem
        if (-not $item) { Show-Erreur "Selectionne une echeance."; return }
        Set-EcheanceStatut -DbPath $DbPath -Id ([int]$item.id) -Statut 'payee'
        Update-VueAdministratif
    }
})

# ============================== BIBLIOTHEQUES ==============================

function Update-VueBibliotheques {
    Update-VueExercices
    Update-VueAliments
    Update-VueComplements
    Update-VueModeles
}

# --- Exercices ---
$GridExercices = Get-Ctrl 'GridExercices'
$TxtExerciceRecherche = Get-Ctrl 'TxtExerciceRecherche'
$TxtExerciceNom = Get-Ctrl 'TxtExerciceNom'
$TxtExerciceMuscle = Get-Ctrl 'TxtExerciceMuscle'
$TxtExerciceVariante = Get-Ctrl 'TxtExerciceVariante'
$TxtExerciceLien = Get-Ctrl 'TxtExerciceLien'
$Script:SelectedExerciceId = $null

function Update-VueExercices {
    $GridExercices.ItemsSource = @(Get-Exercices -DbPath $DbPath -Recherche (Get-TexteOuNull $TxtExerciceRecherche.Text))
}
function Clear-FormExercice {
    $Script:SelectedExerciceId = $null
    $TxtExerciceNom.Text = ''; $TxtExerciceMuscle.Text = ''; $TxtExerciceVariante.Text = ''; $TxtExerciceLien.Text = ''
    $GridExercices.SelectedItem = $null
}
$GridExercices.Add_SelectionChanged({
    $item = $GridExercices.SelectedItem
    if ($null -eq $item) { return }
    $Script:SelectedExerciceId = [int]$item.id
    $TxtExerciceNom.Text = [string]$item.nom
    $TxtExerciceMuscle.Text = [string]$item.muscle_cible
    $TxtExerciceVariante.Text = [string]$item.variante
    $TxtExerciceLien.Text = [string]$item.lien_video
})
(Get-Ctrl 'BtnExerciceRechercher').Add_Click({ Update-VueExercices })
(Get-Ctrl 'BtnExerciceNouveau').Add_Click({ Clear-FormExercice })
(Get-Ctrl 'BtnExerciceEnregistrer').Add_Click({
    Invoke-Protege {
        if ([string]::IsNullOrWhiteSpace($TxtExerciceNom.Text)) { Show-Erreur "Le nom est obligatoire."; return }
        if ($Script:SelectedExerciceId) {
            Update-Exercice -DbPath $DbPath -Id $Script:SelectedExerciceId -Nom $TxtExerciceNom.Text.Trim() `
                -MuscleCible (Get-TexteOuNull $TxtExerciceMuscle.Text) -Variante (Get-TexteOuNull $TxtExerciceVariante.Text) -LienVideo (Get-TexteOuNull $TxtExerciceLien.Text)
        } else {
            New-Exercice -DbPath $DbPath -Nom $TxtExerciceNom.Text.Trim() `
                -MuscleCible (Get-TexteOuNull $TxtExerciceMuscle.Text) -Variante (Get-TexteOuNull $TxtExerciceVariante.Text) -LienVideo (Get-TexteOuNull $TxtExerciceLien.Text) | Out-Null
        }
        Clear-FormExercice
        Update-VueExercices
    }
})
(Get-Ctrl 'BtnExerciceSupprimer').Add_Click({
    Invoke-Protege {
        if (-not $Script:SelectedExerciceId) { Show-Erreur "Selectionne un exercice."; return }
        if (Show-Confirmation "Supprimer cet exercice ?") {
            Remove-Exercice -DbPath $DbPath -Id $Script:SelectedExerciceId
            Clear-FormExercice
            Update-VueExercices
        }
    }
})
(Get-Ctrl 'BtnExerciceExporter').Add_Click({
    Invoke-Protege {
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = 'Fichier Excel (*.xlsx)|*.xlsx'
        $dialog.FileName = 'Exercices.xlsx'
        if ($dialog.ShowDialog()) {
            Export-DonneesVersExcel -Donnees (Get-Exercices -DbPath $DbPath) -Path $dialog.FileName -WorksheetName 'Exercices'
            Show-Info "Export termine."
        }
    }
})

# --- Aliments ---
$GridAliments = Get-Ctrl 'GridAliments'
$TxtAlimentRecherche = Get-Ctrl 'TxtAlimentRecherche'
$TxtAlimentNom = Get-Ctrl 'TxtAlimentNom'
$TxtAlimentQuantite = Get-Ctrl 'TxtAlimentQuantite'
$TxtAlimentUnite = Get-Ctrl 'TxtAlimentUnite'
$TxtAlimentKcal = Get-Ctrl 'TxtAlimentKcal'
$TxtAlimentProteines = Get-Ctrl 'TxtAlimentProteines'
$TxtAlimentGlucides = Get-Ctrl 'TxtAlimentGlucides'
$TxtAlimentLipides = Get-Ctrl 'TxtAlimentLipides'
$TxtAlimentFibres = Get-Ctrl 'TxtAlimentFibres'
$Script:SelectedAlimentId = $null

function Update-VueAliments {
    $GridAliments.ItemsSource = @(Get-Aliments -DbPath $DbPath -Recherche (Get-TexteOuNull $TxtAlimentRecherche.Text))
}
function Clear-FormAliment {
    $Script:SelectedAlimentId = $null
    $TxtAlimentNom.Text = ''; $TxtAlimentQuantite.Text = '100'; $TxtAlimentUnite.Text = 'g'
    $TxtAlimentKcal.Text = ''; $TxtAlimentProteines.Text = ''; $TxtAlimentGlucides.Text = ''
    $TxtAlimentLipides.Text = ''; $TxtAlimentFibres.Text = ''
    $GridAliments.SelectedItem = $null
}
$GridAliments.Add_SelectionChanged({
    $item = $GridAliments.SelectedItem
    if ($null -eq $item) { return }
    $Script:SelectedAlimentId = [int]$item.id
    $TxtAlimentNom.Text = [string]$item.nom
    $TxtAlimentQuantite.Text = [string]$item.quantite_reference
    $TxtAlimentUnite.Text = [string]$item.unite
    $TxtAlimentKcal.Text = [string]$item.kcal
    $TxtAlimentProteines.Text = [string]$item.proteines
    $TxtAlimentGlucides.Text = [string]$item.glucides
    $TxtAlimentLipides.Text = [string]$item.lipides
    $TxtAlimentFibres.Text = [string]$item.fibres
})
(Get-Ctrl 'BtnAlimentRechercher').Add_Click({ Update-VueAliments })
(Get-Ctrl 'BtnAlimentNouveau').Add_Click({ Clear-FormAliment })
(Get-Ctrl 'BtnAlimentEnregistrer').Add_Click({
    Invoke-Protege {
        if ([string]::IsNullOrWhiteSpace($TxtAlimentNom.Text)) { Show-Erreur "Le nom est obligatoire."; return }
        $alimentArgs = @{
            Nom = $TxtAlimentNom.Text.Trim()
            QuantiteReference = (Get-DoubleOuNull $TxtAlimentQuantite.Text)
            Unite = (Get-TexteOuNull $TxtAlimentUnite.Text)
            Kcal = (Get-DoubleOuNull $TxtAlimentKcal.Text)
            Proteines = (Get-DoubleOuNull $TxtAlimentProteines.Text)
            Glucides = (Get-DoubleOuNull $TxtAlimentGlucides.Text)
            Lipides = (Get-DoubleOuNull $TxtAlimentLipides.Text)
            Fibres = (Get-DoubleOuNull $TxtAlimentFibres.Text)
        }
        if ($Script:SelectedAlimentId) {
            Update-Aliment -DbPath $DbPath -Id $Script:SelectedAlimentId @alimentArgs
        } else {
            New-Aliment -DbPath $DbPath @alimentArgs | Out-Null
        }
        Clear-FormAliment
        Update-VueAliments
    }
})
(Get-Ctrl 'BtnAlimentSupprimer').Add_Click({
    Invoke-Protege {
        if (-not $Script:SelectedAlimentId) { Show-Erreur "Selectionne un aliment."; return }
        if (Show-Confirmation "Supprimer cet aliment ?") {
            Remove-Aliment -DbPath $DbPath -Id $Script:SelectedAlimentId
            Clear-FormAliment
            Update-VueAliments
        }
    }
})
(Get-Ctrl 'BtnAlimentExporter').Add_Click({
    Invoke-Protege {
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = 'Fichier Excel (*.xlsx)|*.xlsx'
        $dialog.FileName = 'Aliments.xlsx'
        if ($dialog.ShowDialog()) {
            Export-DonneesVersExcel -Donnees (Get-Aliments -DbPath $DbPath) -Path $dialog.FileName -WorksheetName 'Aliments'
            Show-Info "Export termine."
        }
    }
})

# --- Complements ---
$GridComplements = Get-Ctrl 'GridComplements'
$TxtComplementRecherche = Get-Ctrl 'TxtComplementRecherche'
$TxtComplementNom = Get-Ctrl 'TxtComplementNom'
$TxtComplementDose = Get-Ctrl 'TxtComplementDose'
$TxtComplementTiming = Get-Ctrl 'TxtComplementTiming'
$TxtComplementNote = Get-Ctrl 'TxtComplementNote'
$TxtComplementLien = Get-Ctrl 'TxtComplementLien'
$Script:SelectedComplementId = $null

function Update-VueComplements {
    $GridComplements.ItemsSource = @(Get-Complements -DbPath $DbPath -Recherche (Get-TexteOuNull $TxtComplementRecherche.Text))
}
function Clear-FormComplement {
    $Script:SelectedComplementId = $null
    $TxtComplementNom.Text = ''; $TxtComplementDose.Text = ''; $TxtComplementTiming.Text = ''
    $TxtComplementNote.Text = ''; $TxtComplementLien.Text = ''
    $GridComplements.SelectedItem = $null
}
$GridComplements.Add_SelectionChanged({
    $item = $GridComplements.SelectedItem
    if ($null -eq $item) { return }
    $Script:SelectedComplementId = [int]$item.id
    $TxtComplementNom.Text = [string]$item.nom
    $TxtComplementDose.Text = [string]$item.dose
    $TxtComplementTiming.Text = [string]$item.timing
    $TxtComplementNote.Text = [string]$item.note
    $TxtComplementLien.Text = [string]$item.lien
})
(Get-Ctrl 'BtnComplementRechercher').Add_Click({ Update-VueComplements })
(Get-Ctrl 'BtnComplementNouveau').Add_Click({ Clear-FormComplement })
(Get-Ctrl 'BtnComplementEnregistrer').Add_Click({
    Invoke-Protege {
        if ([string]::IsNullOrWhiteSpace($TxtComplementNom.Text)) { Show-Erreur "Le nom est obligatoire."; return }
        if ($Script:SelectedComplementId) {
            Update-Complement -DbPath $DbPath -Id $Script:SelectedComplementId -Nom $TxtComplementNom.Text.Trim() `
                -Dose (Get-TexteOuNull $TxtComplementDose.Text) -Timing (Get-TexteOuNull $TxtComplementTiming.Text) `
                -Note (Get-TexteOuNull $TxtComplementNote.Text) -Lien (Get-TexteOuNull $TxtComplementLien.Text)
        } else {
            New-Complement -DbPath $DbPath -Nom $TxtComplementNom.Text.Trim() `
                -Dose (Get-TexteOuNull $TxtComplementDose.Text) -Timing (Get-TexteOuNull $TxtComplementTiming.Text) `
                -Note (Get-TexteOuNull $TxtComplementNote.Text) -Lien (Get-TexteOuNull $TxtComplementLien.Text) | Out-Null
        }
        Clear-FormComplement
        Update-VueComplements
    }
})
(Get-Ctrl 'BtnComplementSupprimer').Add_Click({
    Invoke-Protege {
        if (-not $Script:SelectedComplementId) { Show-Erreur "Selectionne un complement."; return }
        if (Show-Confirmation "Supprimer ce complement ?") {
            Remove-Complement -DbPath $DbPath -Id $Script:SelectedComplementId
            Clear-FormComplement
            Update-VueComplements
        }
    }
})
(Get-Ctrl 'BtnComplementExporter').Add_Click({
    Invoke-Protege {
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = 'Fichier Excel (*.xlsx)|*.xlsx'
        $dialog.FileName = 'Complements.xlsx'
        if ($dialog.ShowDialog()) {
            Export-DonneesVersExcel -Donnees (Get-Complements -DbPath $DbPath) -Path $dialog.FileName -WorksheetName 'Complements'
            Show-Info "Export termine."
        }
    }
})

# --- Modeles de seance ---
$ListeModeles = Get-Ctrl 'ListeModeles'
$TxtModeleNom = Get-Ctrl 'TxtModeleNom'
$TxtModeleNotes = Get-Ctrl 'TxtModeleNotes'
$GridModeleExercices = Get-Ctrl 'GridModeleExercices'
$CmbModeleExerciceAAjouter = Get-Ctrl 'CmbModeleExerciceAAjouter'
$TxtModExASeries = Get-Ctrl 'TxtModExASeries'
$TxtModExARepetitions = Get-Ctrl 'TxtModExARepetitions'
$TxtModExARecup = Get-Ctrl 'TxtModExARecup'
$TxtModExATempo = Get-Ctrl 'TxtModExATempo'
$TxtModExANotes = Get-Ctrl 'TxtModExANotes'
$Script:SelectedModeleId = $null
$Script:SelectedModeleExerciceId = $null

function Update-VueModeles {
    $ListeModeles.ItemsSource = @(Get-SeanceModeles -DbPath $DbPath)
    $CmbModeleExerciceAAjouter.ItemsSource = @(Get-Exercices -DbPath $DbPath | ForEach-Object {
        [pscustomobject]@{ id = $_.id; affichage = "$($_.nom) ($($_.muscle_cible))" }
    })
    $GridModeleExercices.ItemsSource = $null
}
function Clear-FormModele {
    $Script:SelectedModeleId = $null
    $TxtModeleNom.Text = ''; $TxtModeleNotes.Text = ''
    $ListeModeles.SelectedItem = $null
    $GridModeleExercices.ItemsSource = $null
}
function Clear-FormModeleExercice {
    $Script:SelectedModeleExerciceId = $null
    $CmbModeleExerciceAAjouter.SelectedItem = $null
    $TxtModExASeries.Text = ''; $TxtModExARepetitions.Text = ''; $TxtModExARecup.Text = ''; $TxtModExATempo.Text = ''; $TxtModExANotes.Text = ''
    $GridModeleExercices.SelectedItem = $null
}
function Update-VueModeleExercices {
    $GridModeleExercices.ItemsSource = $null
    Clear-FormModeleExercice
    if (-not $ListeModeles.SelectedItem) { return }
    $GridModeleExercices.ItemsSource = @(Get-SeanceModeleExercices -DbPath $DbPath -SeanceModeleId $ListeModeles.SelectedItem.id)
}
$ListeModeles.Add_SelectionChanged({
    $item = $ListeModeles.SelectedItem
    if ($null -eq $item) { return }
    $Script:SelectedModeleId = [int]$item.id
    $TxtModeleNom.Text = [string]$item.nom
    $TxtModeleNotes.Text = [string]$item.notes
    Update-VueModeleExercices
})
$GridModeleExercices.Add_SelectionChanged({
    $item = $GridModeleExercices.SelectedItem
    if ($null -eq $item) { return }
    $Script:SelectedModeleExerciceId = [int]$item.id
    $CmbModeleExerciceAAjouter.SelectedItem = $CmbModeleExerciceAAjouter.Items | Where-Object { [int]$_.id -eq [int]$item.exercice_id }
    $TxtModExASeries.Text = [string]$item.series
    $TxtModExARepetitions.Text = [string]$item.repetitions
    $TxtModExARecup.Text = [string]$item.recuperation_s
    $TxtModExATempo.Text = [string]$item.tempo
    $TxtModExANotes.Text = [string]$item.notes
})
(Get-Ctrl 'BtnModeleNouveau').Add_Click({ Clear-FormModele })
(Get-Ctrl 'BtnModeleEnregistrer').Add_Click({
    Invoke-Protege {
        if ([string]::IsNullOrWhiteSpace($TxtModeleNom.Text)) { Show-Erreur "Le nom est obligatoire."; return }
        if ($Script:SelectedModeleId) {
            Update-SeanceModele -DbPath $DbPath -Id $Script:SelectedModeleId -Nom $TxtModeleNom.Text.Trim() -Notes (Get-TexteOuNull $TxtModeleNotes.Text)
        } else {
            New-SeanceModele -DbPath $DbPath -Nom $TxtModeleNom.Text.Trim() -Notes (Get-TexteOuNull $TxtModeleNotes.Text) | Out-Null
        }
        Clear-FormModele
        Update-VueModeles
    }
})
(Get-Ctrl 'BtnModeleSupprimer').Add_Click({
    Invoke-Protege {
        if (-not $Script:SelectedModeleId) { Show-Erreur "Selectionne un modele."; return }
        if (Show-Confirmation "Supprimer ce modele et ses exercices ?") {
            Remove-SeanceModele -DbPath $DbPath -Id $Script:SelectedModeleId
            Clear-FormModele
            Update-VueModeles
        }
    }
})
(Get-Ctrl 'BtnModeleExerciceNouveau').Add_Click({ Clear-FormModeleExercice })

(Get-Ctrl 'BtnModeleExerciceAjouter').Add_Click({
    Invoke-Protege {
        if (-not $ListeModeles.SelectedItem) { Show-Erreur "Selectionne d'abord un modele."; return }
        if ($Script:SelectedModeleExerciceId) {
            Update-SeanceModeleExercice -DbPath $DbPath -Id $Script:SelectedModeleExerciceId `
                -Series (Get-IntOuNull $TxtModExASeries.Text) -Repetitions (Get-TexteOuNull $TxtModExARepetitions.Text) `
                -RecuperationS (Get-IntOuNull $TxtModExARecup.Text) -Tempo (Get-TexteOuNull $TxtModExATempo.Text) -Notes (Get-TexteOuNull $TxtModExANotes.Text)
        } else {
            if (-not $CmbModeleExerciceAAjouter.SelectedItem) { Show-Erreur "Selectionne un exercice dans la liste."; return }
            New-SeanceModeleExercice -DbPath $DbPath -SeanceModeleId $ListeModeles.SelectedItem.id -ExerciceId $CmbModeleExerciceAAjouter.SelectedItem.id `
                -Series (Get-IntOuNull $TxtModExASeries.Text) -Repetitions (Get-TexteOuNull $TxtModExARepetitions.Text) `
                -RecuperationS (Get-IntOuNull $TxtModExARecup.Text) -Tempo (Get-TexteOuNull $TxtModExATempo.Text) -Notes (Get-TexteOuNull $TxtModExANotes.Text) | Out-Null
        }
        Update-VueModeleExercices
    }
})
(Get-Ctrl 'BtnModeleExerciceSupprimer').Add_Click({
    Invoke-Protege {
        $item = $GridModeleExercices.SelectedItem
        if (-not $item) { Show-Erreur "Selectionne un exercice dans le tableau."; return }
        Remove-SeanceModeleExercice -DbPath $DbPath -Id ([int]$item.id)
        Update-VueModeleExercices
    }
})

# ============================== PROGRAMMES ==============================

$CmbProgrammeClient = Get-Ctrl 'CmbProgrammeClient'
$CmbProgrammeSelection = Get-Ctrl 'CmbProgrammeSelection'
$ListeSeances = Get-Ctrl 'ListeSeances'
$TxtNouvelleSeance = Get-Ctrl 'TxtNouvelleSeance'
$GridSeanceExercices = Get-Ctrl 'GridSeanceExercices'
$CmbExerciceAAjouter = Get-Ctrl 'CmbExerciceAAjouter'
$TxtExASeries = Get-Ctrl 'TxtExASeries'
$TxtExARepetitions = Get-Ctrl 'TxtExARepetitions'
$TxtExARecup = Get-Ctrl 'TxtExARecup'
$TxtExATempo = Get-Ctrl 'TxtExATempo'
$TxtExANotes = Get-Ctrl 'TxtExANotes'
$Script:SelectedSeanceExerciceId = $null

function Clear-FormSeanceExercice {
    $Script:SelectedSeanceExerciceId = $null
    $CmbExerciceAAjouter.SelectedItem = $null
    $TxtExASeries.Text = ''; $TxtExARepetitions.Text = ''; $TxtExARecup.Text = ''; $TxtExATempo.Text = ''; $TxtExANotes.Text = ''
    $GridSeanceExercices.SelectedItem = $null
}

function Update-VueProgrammesClients {
    $clients = @(Get-Clients -DbPath $DbPath | ForEach-Object {
        [pscustomobject]@{ id = $_.id; affichage = "$($_.nom) $($_.prenom)" }
    })
    $CmbProgrammeClient.ItemsSource = $clients
    $CmbExerciceAAjouter.ItemsSource = @(Get-Exercices -DbPath $DbPath | ForEach-Object {
        [pscustomobject]@{ id = $_.id; affichage = "$($_.nom) ($($_.muscle_cible))" }
    })
    if ($clients.Count -gt 0 -and -not $CmbProgrammeClient.SelectedItem) { $CmbProgrammeClient.SelectedIndex = 0 }
}

function Update-VueProgrammesPourClient {
    $CmbProgrammeSelection.ItemsSource = $null
    $ListeSeances.ItemsSource = $null
    $GridSeanceExercices.ItemsSource = $null
    if (-not $CmbProgrammeClient.SelectedItem) { return }
    $programmes = @(Get-Programmes -DbPath $DbPath -ClientId $CmbProgrammeClient.SelectedItem.id)
    $CmbProgrammeSelection.ItemsSource = $programmes
    if ($programmes.Count -gt 0) { $CmbProgrammeSelection.SelectedIndex = 0 }
}

function Update-VueSeances {
    $ListeSeances.ItemsSource = $null
    $GridSeanceExercices.ItemsSource = $null
    if (-not $CmbProgrammeSelection.SelectedItem) { return }
    $seances = @(Get-Seances -DbPath $DbPath -ProgrammeId $CmbProgrammeSelection.SelectedItem.id)
    $ListeSeances.ItemsSource = $seances
    if ($seances.Count -gt 0) { $ListeSeances.SelectedIndex = 0 }
}

function Update-VueSeanceExercices {
    $GridSeanceExercices.ItemsSource = $null
    Clear-FormSeanceExercice
    if (-not $ListeSeances.SelectedItem) { return }
    $GridSeanceExercices.ItemsSource = @(Get-SeanceExercices -DbPath $DbPath -SeanceId $ListeSeances.SelectedItem.id)
}

$CmbProgrammeClient.Add_SelectionChanged({ Update-VueProgrammesPourClient })
$CmbProgrammeSelection.Add_SelectionChanged({ Update-VueSeances })
$ListeSeances.Add_SelectionChanged({ Update-VueSeanceExercices })
$GridSeanceExercices.Add_SelectionChanged({
    $item = $GridSeanceExercices.SelectedItem
    if ($null -eq $item) { return }
    $Script:SelectedSeanceExerciceId = [int]$item.id
    $CmbExerciceAAjouter.SelectedItem = $CmbExerciceAAjouter.Items | Where-Object { [int]$_.id -eq [int]$item.exercice_id }
    $TxtExASeries.Text = [string]$item.series
    $TxtExARepetitions.Text = [string]$item.repetitions
    $TxtExARecup.Text = [string]$item.recuperation_s
    $TxtExATempo.Text = [string]$item.tempo
    $TxtExANotes.Text = [string]$item.notes
})

(Get-Ctrl 'BtnProgrammeNouveau').Add_Click({
    Invoke-Protege {
        if (-not $CmbProgrammeClient.SelectedItem) { Show-Erreur "Selectionne d'abord un client."; return }
        $resultat = Show-DialogNouvelElement -Titre "Nouveau programme sportif"
        if (-not $resultat) { return }
        New-Programme -DbPath $DbPath -ClientId $CmbProgrammeClient.SelectedItem.id -Nom $resultat.Nom -DateDebut $resultat.DateDebut -Notes $resultat.Notes | Out-Null
        Update-VueProgrammesPourClient
    }
})

(Get-Ctrl 'BtnProgrammeSupprimer').Add_Click({
    Invoke-Protege {
        if (-not $CmbProgrammeSelection.SelectedItem) { Show-Erreur "Selectionne un programme."; return }
        if (Show-Confirmation "Supprimer ce programme et tout son contenu (seances, exercices) ?") {
            Remove-Programme -DbPath $DbPath -Id $CmbProgrammeSelection.SelectedItem.id
            Update-VueProgrammesPourClient
        }
    }
})

(Get-Ctrl 'BtnProgrammeExporterPdf').Add_Click({
    Invoke-Protege {
        if (-not $CmbProgrammeSelection.SelectedItem) { Show-Erreur "Selectionne un programme."; return }
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = 'Fichier PDF (*.pdf)|*.pdf'
        $dialog.FileName = "Programme.pdf"
        if ($dialog.ShowDialog()) {
            Export-ProgrammePdf -DbPath $DbPath -ProgrammeId $CmbProgrammeSelection.SelectedItem.id -Path $dialog.FileName
            Show-Info "Export PDF termine."
        }
    }
})

(Get-Ctrl 'BtnProgrammeExporterExcel').Add_Click({
    Invoke-Protege {
        if (-not $CmbProgrammeSelection.SelectedItem) { Show-Erreur "Selectionne un programme."; return }
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = 'Fichier Excel (*.xlsx)|*.xlsx'
        $dialog.FileName = "Programme.xlsx"
        if ($dialog.ShowDialog()) {
            Export-ProgrammeExcel -DbPath $DbPath -ProgrammeId $CmbProgrammeSelection.SelectedItem.id -Path $dialog.FileName
            Show-Info "Export Excel termine."
        }
    }
})

(Get-Ctrl 'BtnSeanceAjouter').Add_Click({
    Invoke-Protege {
        if (-not $CmbProgrammeSelection.SelectedItem) { Show-Erreur "Selectionne d'abord un programme."; return }
        if ([string]::IsNullOrWhiteSpace($TxtNouvelleSeance.Text)) { Show-Erreur "Indique un nom de seance."; return }
        New-Seance -DbPath $DbPath -ProgrammeId $CmbProgrammeSelection.SelectedItem.id -Nom $TxtNouvelleSeance.Text.Trim() | Out-Null
        $TxtNouvelleSeance.Text = ''
        Update-VueSeances
    }
})

(Get-Ctrl 'BtnSeanceDepuisModele').Add_Click({
    Invoke-Protege {
        if (-not $CmbProgrammeSelection.SelectedItem) { Show-Erreur "Selectionne d'abord un programme."; return }
        $modeles = @(Get-SeanceModeles -DbPath $DbPath)
        if ($modeles.Count -eq 0) { Show-Erreur "Aucun modele de seance disponible. Cree d'abord un modele dans Bibliotheques > Modeles de seance."; return }
        $resultat = Show-DialogChoixModele -Modeles $modeles
        if (-not $resultat) { return }
        New-SeanceDepuisModele -DbPath $DbPath -ProgrammeId $CmbProgrammeSelection.SelectedItem.id -SeanceModeleId $resultat.SeanceModeleId -Nom $resultat.Nom | Out-Null
        Update-VueSeances
    }
})

(Get-Ctrl 'BtnSeanceMonter').Add_Click({
    Invoke-Protege {
        if (-not $ListeSeances.SelectedItem) { Show-Erreur "Selectionne une seance."; return }
        Move-Seance -DbPath $DbPath -Id $ListeSeances.SelectedItem.id -ProgrammeId $CmbProgrammeSelection.SelectedItem.id -Direction -1
        Update-VueSeances
    }
})
(Get-Ctrl 'BtnSeanceDescendre').Add_Click({
    Invoke-Protege {
        if (-not $ListeSeances.SelectedItem) { Show-Erreur "Selectionne une seance."; return }
        Move-Seance -DbPath $DbPath -Id $ListeSeances.SelectedItem.id -ProgrammeId $CmbProgrammeSelection.SelectedItem.id -Direction 1
        Update-VueSeances
    }
})
(Get-Ctrl 'BtnSeanceSupprimer').Add_Click({
    Invoke-Protege {
        if (-not $ListeSeances.SelectedItem) { Show-Erreur "Selectionne une seance."; return }
        if (Show-Confirmation "Supprimer cette seance et ses exercices ?") {
            Remove-Seance -DbPath $DbPath -Id $ListeSeances.SelectedItem.id
            Update-VueSeances
        }
    }
})

(Get-Ctrl 'BtnExerciceSeanceNouveau').Add_Click({ Clear-FormSeanceExercice })

(Get-Ctrl 'BtnExerciceAjouter').Add_Click({
    Invoke-Protege {
        if (-not $ListeSeances.SelectedItem) { Show-Erreur "Selectionne d'abord une seance."; return }
        if ($Script:SelectedSeanceExerciceId) {
            Update-SeanceExercice -DbPath $DbPath -Id $Script:SelectedSeanceExerciceId `
                -Series (Get-IntOuNull $TxtExASeries.Text) -Repetitions (Get-TexteOuNull $TxtExARepetitions.Text) `
                -RecuperationS (Get-IntOuNull $TxtExARecup.Text) -Tempo (Get-TexteOuNull $TxtExATempo.Text) -Notes (Get-TexteOuNull $TxtExANotes.Text)
        } else {
            if (-not $CmbExerciceAAjouter.SelectedItem) { Show-Erreur "Selectionne un exercice dans la liste."; return }
            New-SeanceExercice -DbPath $DbPath -SeanceId $ListeSeances.SelectedItem.id -ExerciceId $CmbExerciceAAjouter.SelectedItem.id `
                -Series (Get-IntOuNull $TxtExASeries.Text) -Repetitions (Get-TexteOuNull $TxtExARepetitions.Text) `
                -RecuperationS (Get-IntOuNull $TxtExARecup.Text) -Tempo (Get-TexteOuNull $TxtExATempo.Text) -Notes (Get-TexteOuNull $TxtExANotes.Text) | Out-Null
        }
        Update-VueSeanceExercices
    }
})

(Get-Ctrl 'BtnExerciceSupprimerDeSeance').Add_Click({
    Invoke-Protege {
        $item = $GridSeanceExercices.SelectedItem
        if (-not $item) { Show-Erreur "Selectionne un exercice dans le tableau."; return }
        Remove-SeanceExercice -DbPath $DbPath -Id ([int]$item.id)
        Update-VueSeanceExercices
    }
})

# ============================== NUTRITION ==============================

$CmbNutritionClient = Get-Ctrl 'CmbNutritionClient'
$CmbPlanSelection = Get-Ctrl 'CmbPlanSelection'
$ListeTypesJour = Get-Ctrl 'ListeTypesJour'
$TxtNouveauTypeJour = Get-Ctrl 'TxtNouveauTypeJour'
$ListeRepas = Get-Ctrl 'ListeRepas'
$TxtNouveauRepas = Get-Ctrl 'TxtNouveauRepas'
$GridRepasAliments = Get-Ctrl 'GridRepasAliments'
$TxtTotauxRepas = Get-Ctrl 'TxtTotauxRepas'
$CmbAlimentAAjouter = Get-Ctrl 'CmbAlimentAAjouter'
$TxtQuantiteAliment = Get-Ctrl 'TxtQuantiteAliment'

function Update-VueNutritionClients {
    $clients = @(Get-Clients -DbPath $DbPath | ForEach-Object {
        [pscustomobject]@{ id = $_.id; affichage = "$($_.nom) $($_.prenom)" }
    })
    $CmbNutritionClient.ItemsSource = $clients
    $CmbAlimentAAjouter.ItemsSource = @(Get-Aliments -DbPath $DbPath | ForEach-Object {
        [pscustomobject]@{ id = $_.id; affichage = "$($_.nom) ($($_.quantite_reference) $($_.unite) = $($_.kcal) kcal)" }
    })
    if ($clients.Count -gt 0 -and -not $CmbNutritionClient.SelectedItem) { $CmbNutritionClient.SelectedIndex = 0 }
}

function Update-VuePlansPourClient {
    $CmbPlanSelection.ItemsSource = $null
    $ListeTypesJour.ItemsSource = $null
    $ListeRepas.ItemsSource = $null
    $GridRepasAliments.ItemsSource = $null
    if (-not $CmbNutritionClient.SelectedItem) { return }
    $plans = @(Get-PlansNutrition -DbPath $DbPath -ClientId $CmbNutritionClient.SelectedItem.id)
    $CmbPlanSelection.ItemsSource = $plans
    if ($plans.Count -gt 0) { $CmbPlanSelection.SelectedIndex = 0 }
}

function Update-VueTypesJour {
    $ListeTypesJour.ItemsSource = $null
    $ListeRepas.ItemsSource = $null
    $GridRepasAliments.ItemsSource = $null
    if (-not $CmbPlanSelection.SelectedItem) { return }
    $types = @(Get-TypesJour -DbPath $DbPath -PlanNutritionId $CmbPlanSelection.SelectedItem.id)
    $ListeTypesJour.ItemsSource = $types
    if ($types.Count -gt 0) { $ListeTypesJour.SelectedIndex = 0 }
}

function Update-VueRepas {
    $ListeRepas.ItemsSource = $null
    $GridRepasAliments.ItemsSource = $null
    if (-not $ListeTypesJour.SelectedItem) { return }
    $repasListe = @(Get-Repas -DbPath $DbPath -TypeJourId $ListeTypesJour.SelectedItem.id)
    $ListeRepas.ItemsSource = $repasListe
    if ($repasListe.Count -gt 0) { $ListeRepas.SelectedIndex = 0 }
}

function Update-VueRepasAliments {
    $GridRepasAliments.ItemsSource = $null
    $TxtTotauxRepas.Text = ''
    if (-not $ListeRepas.SelectedItem) { return }
    $lignes = @(Get-RepasAliments -DbPath $DbPath -RepasId $ListeRepas.SelectedItem.id)
    $GridRepasAliments.ItemsSource = $lignes
    $totaux = Get-TotauxRepas -Lignes $lignes
    $TxtTotauxRepas.Text = "Total repas : $($totaux.Kcal) kcal — Proteines $($totaux.Proteines) g — Glucides $($totaux.Glucides) g — Lipides $($totaux.Lipides) g — Fibres $($totaux.Fibres) g"
}

$CmbNutritionClient.Add_SelectionChanged({ Update-VuePlansPourClient })
$CmbPlanSelection.Add_SelectionChanged({ Update-VueTypesJour })
$ListeTypesJour.Add_SelectionChanged({ Update-VueRepas })
$ListeRepas.Add_SelectionChanged({ Update-VueRepasAliments })

(Get-Ctrl 'BtnPlanNouveau').Add_Click({
    Invoke-Protege {
        if (-not $CmbNutritionClient.SelectedItem) { Show-Erreur "Selectionne d'abord un client."; return }
        $resultat = Show-DialogNouvelElement -Titre "Nouveau plan nutritionnel"
        if (-not $resultat) { return }
        New-PlanNutrition -DbPath $DbPath -ClientId $CmbNutritionClient.SelectedItem.id -Nom $resultat.Nom -DateDebut $resultat.DateDebut -Notes $resultat.Notes | Out-Null
        Update-VuePlansPourClient
    }
})

(Get-Ctrl 'BtnPlanSupprimer').Add_Click({
    Invoke-Protege {
        if (-not $CmbPlanSelection.SelectedItem) { Show-Erreur "Selectionne un plan."; return }
        if (Show-Confirmation "Supprimer ce plan et tout son contenu ?") {
            Remove-PlanNutrition -DbPath $DbPath -Id $CmbPlanSelection.SelectedItem.id
            Update-VuePlansPourClient
        }
    }
})

(Get-Ctrl 'BtnPlanExporterPdf').Add_Click({
    Invoke-Protege {
        if (-not $CmbPlanSelection.SelectedItem) { Show-Erreur "Selectionne un plan."; return }
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = 'Fichier PDF (*.pdf)|*.pdf'
        $dialog.FileName = "PlanNutrition.pdf"
        if ($dialog.ShowDialog()) {
            Export-PlanNutritionPdf -DbPath $DbPath -PlanNutritionId $CmbPlanSelection.SelectedItem.id -Path $dialog.FileName
            Show-Info "Export PDF termine."
        }
    }
})

(Get-Ctrl 'BtnPlanExporterExcel').Add_Click({
    Invoke-Protege {
        if (-not $CmbPlanSelection.SelectedItem) { Show-Erreur "Selectionne un plan."; return }
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = 'Fichier Excel (*.xlsx)|*.xlsx'
        $dialog.FileName = "PlanNutrition.xlsx"
        if ($dialog.ShowDialog()) {
            Export-PlanNutritionExcel -DbPath $DbPath -PlanNutritionId $CmbPlanSelection.SelectedItem.id -Path $dialog.FileName
            Show-Info "Export Excel termine."
        }
    }
})

(Get-Ctrl 'BtnTypeJourAjouter').Add_Click({
    Invoke-Protege {
        if (-not $CmbPlanSelection.SelectedItem) { Show-Erreur "Selectionne d'abord un plan."; return }
        if ([string]::IsNullOrWhiteSpace($TxtNouveauTypeJour.Text)) { Show-Erreur "Indique un nom (ex: Jour haut)."; return }
        New-TypeJour -DbPath $DbPath -PlanNutritionId $CmbPlanSelection.SelectedItem.id -Nom $TxtNouveauTypeJour.Text.Trim() | Out-Null
        $TxtNouveauTypeJour.Text = ''
        Update-VueTypesJour
    }
})
(Get-Ctrl 'BtnTypeJourSupprimer').Add_Click({
    Invoke-Protege {
        if (-not $ListeTypesJour.SelectedItem) { Show-Erreur "Selectionne un type de jour."; return }
        if (Show-Confirmation "Supprimer ce type de jour et ses repas ?") {
            Remove-TypeJour -DbPath $DbPath -Id $ListeTypesJour.SelectedItem.id
            Update-VueTypesJour
        }
    }
})

(Get-Ctrl 'BtnRepasAjouter').Add_Click({
    Invoke-Protege {
        if (-not $ListeTypesJour.SelectedItem) { Show-Erreur "Selectionne d'abord un type de jour."; return }
        if ([string]::IsNullOrWhiteSpace($TxtNouveauRepas.Text)) { Show-Erreur "Indique un nom de repas."; return }
        New-Repas -DbPath $DbPath -TypeJourId $ListeTypesJour.SelectedItem.id -Nom $TxtNouveauRepas.Text.Trim() | Out-Null
        $TxtNouveauRepas.Text = ''
        Update-VueRepas
    }
})
(Get-Ctrl 'BtnRepasSupprimer').Add_Click({
    Invoke-Protege {
        if (-not $ListeRepas.SelectedItem) { Show-Erreur "Selectionne un repas."; return }
        if (Show-Confirmation "Supprimer ce repas ?") {
            Remove-Repas -DbPath $DbPath -Id $ListeRepas.SelectedItem.id
            Update-VueRepas
        }
    }
})

(Get-Ctrl 'BtnAlimentAjouterRepas').Add_Click({
    Invoke-Protege {
        if (-not $ListeRepas.SelectedItem) { Show-Erreur "Selectionne d'abord un repas."; return }
        if (-not $CmbAlimentAAjouter.SelectedItem) { Show-Erreur "Selectionne un aliment dans la liste."; return }
        $quantite = Get-DoubleOuNull $TxtQuantiteAliment.Text
        if (-not $quantite) { Show-Erreur "Indique une quantite valide."; return }
        New-RepasAliment -DbPath $DbPath -RepasId $ListeRepas.SelectedItem.id -AlimentId $CmbAlimentAAjouter.SelectedItem.id -Quantite $quantite | Out-Null
        $TxtQuantiteAliment.Text = '100'
        Update-VueRepasAliments
    }
})
(Get-Ctrl 'BtnAlimentSupprimerDeRepas').Add_Click({
    Invoke-Protege {
        $item = $GridRepasAliments.SelectedItem
        if (-not $item) { Show-Erreur "Selectionne un aliment dans le tableau."; return }
        Remove-RepasAliment -DbPath $DbPath -Id ([int]$item.id)
        Update-VueRepasAliments
    }
})

# ============================== TABLEAU DE BORD ==============================

$TxtStatClientsActifs = Get-Ctrl 'TxtStatClientsActifs'
$TxtStatEnAttente = Get-Ctrl 'TxtStatEnAttente'
$TxtStatEnRetard = Get-Ctrl 'TxtStatEnRetard'
$GridDashboardEcheances = Get-Ctrl 'GridDashboardEcheances'
$GridDashboardSansBilan = Get-Ctrl 'GridDashboardSansBilan'

function Update-VueDashboard {
    Update-EcheancesRetard -DbPath $DbPath
    $stats = Get-StatsDashboard -DbPath $DbPath
    $TxtStatClientsActifs.Text = [string]$stats.ClientsActifs
    $TxtStatEnAttente.Text = "$($stats.PaiementsEnAttenteNombre) ($($stats.PaiementsEnAttenteMontant) €)"
    $TxtStatEnRetard.Text = "$($stats.PaiementsEnRetardNombre) ($($stats.PaiementsEnRetardMontant) €)"
    $GridDashboardEcheances.ItemsSource = $stats.ProchainesEcheances
    $GridDashboardSansBilan.ItemsSource = $stats.ClientsSansBilanRecent
}

(Get-Ctrl 'BtnDashboardActualiser').Add_Click({ Update-VueDashboard })

# ============================== SUIVI ==============================

$CmbSuiviClient = Get-Ctrl 'CmbSuiviClient'

# --- Questionnaires ---
$ChkQuestionnaireNonRattachees = Get-Ctrl 'ChkQuestionnaireNonRattachees'
$GridQuestionnaires = Get-Ctrl 'GridQuestionnaires'
$TxtDetailReponse = Get-Ctrl 'TxtDetailReponse'

# --- Tracking ---
$GridSuiviQuotidien = Get-Ctrl 'GridSuiviQuotidien'
$CanvasPoids = Get-Ctrl 'CanvasPoids'
$Script:DernieresLignesSuivi = @()

# --- Roadmap ---
$GridRoadmap = Get-Ctrl 'GridRoadmap'
$TxtRoadmapSemaine = Get-Ctrl 'TxtRoadmapSemaine'
$DateRoadmapDebut = Get-Ctrl 'DateRoadmapDebut'
$TxtRoadmapPhase = Get-Ctrl 'TxtRoadmapPhase'
$TxtRoadmapNutrition = Get-Ctrl 'TxtRoadmapNutrition'
$TxtRoadmapPoidsMoyen = Get-Ctrl 'TxtRoadmapPoidsMoyen'
$TxtRoadmapDepense = Get-Ctrl 'TxtRoadmapDepense'
$TxtRoadmapCardio = Get-Ctrl 'TxtRoadmapCardio'
$TxtRoadmapPas = Get-Ctrl 'TxtRoadmapPas'
$TxtRoadmapPrecisionTraining = Get-Ctrl 'TxtRoadmapPrecisionTraining'
$TxtRoadmapEvenements = Get-Ctrl 'TxtRoadmapEvenements'
$TxtRoadmapNotes = Get-Ctrl 'TxtRoadmapNotes'
$Script:SelectedRoadmapId = $null

# --- Journal alimentaire ---
$GridJournalAlimentaire = Get-Ctrl 'GridJournalAlimentaire'

function Update-VueSuiviClients {
    $clients = @(Get-Clients -DbPath $DbPath | ForEach-Object {
        [pscustomobject]@{ id = $_.id; affichage = "$($_.nom) $($_.prenom)" }
    })
    $CmbSuiviClient.ItemsSource = $clients
    if ($clients.Count -gt 0 -and -not $CmbSuiviClient.SelectedItem) { $CmbSuiviClient.SelectedIndex = 0 }
    elseif ($CmbSuiviClient.SelectedItem) { Update-VueSuiviComplet }
}

function Update-VueSuiviComplet {
    Update-VueQuestionnaires
    Update-VueSuiviQuotidien
    Update-VueRoadmap
    Update-VueJournalAlimentaire
}

$CmbSuiviClient.Add_SelectionChanged({ Update-VueSuiviComplet })

# --- Questionnaires : logique ---

function Format-DetailReponse {
    param($DonneesJson)
    if (-not $DonneesJson) { return '' }
    try {
        $obj = $DonneesJson | ConvertFrom-Json
        $lignes = foreach ($prop in $obj.PSObject.Properties) { "$($prop.Name) : $($prop.Value)" }
        return ($lignes -join "`r`n")
    } catch {
        return [string]$DonneesJson
    }
}

function Update-VueQuestionnaires {
    $TxtDetailReponse.Text = "Sélectionne une réponse pour voir le détail."
    if ($ChkQuestionnaireNonRattachees.IsChecked) {
        $GridQuestionnaires.ItemsSource = @(Get-QuestionnairesReponses -DbPath $DbPath -NonRattachees)
    } elseif ($CmbSuiviClient.SelectedItem) {
        $GridQuestionnaires.ItemsSource = @(Get-QuestionnairesReponses -DbPath $DbPath -ClientId $CmbSuiviClient.SelectedItem.id)
    } else {
        $GridQuestionnaires.ItemsSource = $null
    }
}

$ChkQuestionnaireNonRattachees.Add_Click({ Update-VueQuestionnaires })

$GridQuestionnaires.Add_SelectionChanged({
    $item = $GridQuestionnaires.SelectedItem
    $TxtDetailReponse.Text = if ($item) { Format-DetailReponse $item.donnees_json } else { "Sélectionne une réponse pour voir le détail." }
})

(Get-Ctrl 'BtnEnvoyerPreCoaching').Add_Click({
    Invoke-Protege { Start-Process 'https://docs.google.com/forms/d/e/1FAIpQLSdfSKg4FTRaM84Ig-sOaCVPnM83MVWNbEdBRv-E4im6EAiWZw/viewform' }
})
(Get-Ctrl 'BtnEnvoyerBilan').Add_Click({
    Invoke-Protege { Start-Process 'https://docs.google.com/forms/d/e/1FAIpQLSezeUg3guyJ7A3Ly2_7RzJqRninppn0SZSMR_2hhLxucAKaYQ/viewform' }
})

(Get-Ctrl 'BtnImporterQuestionnaire').Add_Click({
    Invoke-Protege {
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = 'Fichier Excel ou CSV (*.xlsx;*.csv)|*.xlsx;*.csv'
        if (-not $dialog.ShowDialog()) { return }

        $entetes = Get-EnTetesExcel -ExcelPath $dialog.FileName
        if ($entetes.Count -eq 0) {
            Show-Erreur "Le fichier ne contient aucune colonne exploitable."
            return
        }
        $resultat = Show-DialogImportQuestionnaire -Entetes $entetes
        if (-not $resultat) { return }

        $res = Import-QuestionnaireDepuisExcel -DbPath $DbPath -Type $resultat.Type -ExcelPath $dialog.FileName `
            -ColonneDate $resultat.ColonneDate -ColonneNom $resultat.ColonneNom -ColonneEmail $resultat.ColonneEmail `
            -ColonneTelephone $resultat.ColonneTelephone -ColonneObjectifLongTerme $resultat.ColonneObjectifLongTerme `
            -ColonneObjectifCourtTerme $resultat.ColonneObjectifCourtTerme -ColonneMoyens $resultat.ColonneMoyens
        Save-ImportMapping -DbPath $DbPath -Type $resultat.Type -MappingJson ($resultat | ConvertTo-Json -Compress)

        $message = "Réponses importées : $($res.Importees)`nRattachées automatiquement : $($res.Rattachees)`nNon rattachées : $($res.NonRattachees)"
        if ($resultat.Type -eq 'pre_coaching') { $message += "`nFiches client complétées : $($res.FichesCompletees)" }
        Show-Info $message
        Update-VueQuestionnaires
    }
})

(Get-Ctrl 'BtnRattacherReponse').Add_Click({
    Invoke-Protege {
        $item = $GridQuestionnaires.SelectedItem
        if (-not $item) { Show-Erreur "Sélectionne une réponse."; return }
        if (-not $CmbSuiviClient.SelectedItem) { Show-Erreur "Sélectionne le client auquel rattacher cette réponse."; return }
        Set-QuestionnaireReponseClient -DbPath $DbPath -Id ([int]$item.id) -ClientId $CmbSuiviClient.SelectedItem.id
        Update-VueQuestionnaires
    }
})

# --- Tracking quotidien : logique ---

function Dessiner-GraphiquePoids {
    $CanvasPoids.Children.Clear()
    $points = @($Script:DernieresLignesSuivi | Where-Object { $_.poids } | Sort-Object date)
    if ($points.Count -lt 2) { return }

    $largeur = $CanvasPoids.ActualWidth
    $hauteur = $CanvasPoids.ActualHeight
    if ($largeur -le 10 -or $hauteur -le 10) { return }

    $valeurs = $points | ForEach-Object { [double]$_.poids }
    $min = ($valeurs | Measure-Object -Minimum).Minimum
    $max = ($valeurs | Measure-Object -Maximum).Maximum
    if ($max -eq $min) { $max = $min + 1 }
    $margeGauche = 40; $margeAutres = 10

    $polyline = New-Object System.Windows.Shapes.Polyline
    $polyline.Stroke = [System.Windows.Media.Brushes]::SteelBlue
    $polyline.StrokeThickness = 2
    $pts = New-Object System.Windows.Media.PointCollection
    for ($i = 0; $i -lt $points.Count; $i++) {
        $x = $margeGauche + ($i / [math]::Max(1, ($points.Count - 1))) * ($largeur - $margeGauche - $margeAutres)
        $y = ($hauteur - $margeAutres) - ((([double]$points[$i].poids - $min) / ($max - $min)) * ($hauteur - 2 * $margeAutres))
        [void]$pts.Add((New-Object System.Windows.Point($x, $y)))
    }
    $polyline.Points = $pts
    [void]$CanvasPoids.Children.Add($polyline)

    foreach ($v in @($max, $min)) {
        $y = ($hauteur - $margeAutres) - ((($v - $min) / ($max - $min)) * ($hauteur - 2 * $margeAutres))
        $txt = New-Object System.Windows.Controls.TextBlock
        $txt.Text = [string][math]::Round($v, 1)
        $txt.FontSize = 11
        $txt.Foreground = [System.Windows.Media.Brushes]::Gray
        [System.Windows.Controls.Canvas]::SetLeft($txt, 0)
        [System.Windows.Controls.Canvas]::SetTop($txt, $y - 7)
        [void]$CanvasPoids.Children.Add($txt)
    }
}

$CanvasPoids.Add_SizeChanged({ Dessiner-GraphiquePoids })

function Update-VueSuiviQuotidien {
    if (-not $CmbSuiviClient.SelectedItem) {
        $GridSuiviQuotidien.ItemsSource = $null
        $Script:DernieresLignesSuivi = @()
        Dessiner-GraphiquePoids
        return
    }
    $lignes = @(Get-SuiviQuotidien -DbPath $DbPath -ClientId $CmbSuiviClient.SelectedItem.id)
    $GridSuiviQuotidien.ItemsSource = $lignes
    $Script:DernieresLignesSuivi = $lignes
    Dessiner-GraphiquePoids
}

(Get-Ctrl 'BtnTelechargerModele').Add_Click({
    Invoke-Protege {
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = 'Fichier Excel (*.xlsx)|*.xlsx'
        $dialog.FileName = 'Modele_Suivi_Quotidien.xlsx'
        if ($dialog.ShowDialog()) {
            Export-ModeleTrackingExcel -Path $dialog.FileName
            Show-Info "Modèle créé. Envoie ce fichier à ton client pour qu'il le remplisse."
        }
    }
})

(Get-Ctrl 'BtnImporterTracking').Add_Click({
    Invoke-Protege {
        if (-not $CmbSuiviClient.SelectedItem) { Show-Erreur "Sélectionne d'abord un client."; return }
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = 'Fichier Excel (*.xlsx)|*.xlsx'
        if ($dialog.ShowDialog()) {
            $res = Import-TrackingDepuisExcel -DbPath $DbPath -ClientId $CmbSuiviClient.SelectedItem.id -ExcelPath $dialog.FileName
            $message = "Jours importés : $($res.Importes)`nLignes sans date ignorées : $($res.IgnoresSansDate)"
            if ($res.Erreurs.Count -gt 0) { $message += "`n`nRemarques :`n" + ($res.Erreurs -join "`n") }
            Show-Info $message "Import terminé"
            Update-VueSuiviQuotidien
        }
    }
})

(Get-Ctrl 'BtnSupprimerSuiviJour').Add_Click({
    Invoke-Protege {
        $item = $GridSuiviQuotidien.SelectedItem
        if (-not $item) { Show-Erreur "Sélectionne une ligne."; return }
        if (Show-Confirmation "Supprimer le suivi du $($item.date) ?") {
            Remove-SuiviQuotidienJour -DbPath $DbPath -Id ([int]$item.id)
            Update-VueSuiviQuotidien
        }
    }
})

# --- Roadmap : logique ---

function Clear-FormRoadmap {
    $Script:SelectedRoadmapId = $null
    $TxtRoadmapSemaine.Text = ''; $DateRoadmapDebut.SelectedDate = $null; $TxtRoadmapPhase.Text = ''
    $TxtRoadmapNutrition.Text = ''; $TxtRoadmapPoidsMoyen.Text = ''; $TxtRoadmapDepense.Text = ''
    $TxtRoadmapCardio.Text = ''; $TxtRoadmapPas.Text = ''; $TxtRoadmapPrecisionTraining.Text = ''
    $TxtRoadmapEvenements.Text = ''; $TxtRoadmapNotes.Text = ''
    $GridRoadmap.SelectedItem = $null
}

function Update-VueRoadmap {
    Clear-FormRoadmap
    if (-not $CmbSuiviClient.SelectedItem) { $GridRoadmap.ItemsSource = $null; return }
    $GridRoadmap.ItemsSource = @(Get-RoadmapSemaines -DbPath $DbPath -ClientId $CmbSuiviClient.SelectedItem.id)
}

$GridRoadmap.Add_SelectionChanged({
    $item = $GridRoadmap.SelectedItem
    if ($null -eq $item) { return }
    $Script:SelectedRoadmapId = [int]$item.id
    $TxtRoadmapSemaine.Text = [string]$item.semaine_numero
    $DateRoadmapDebut.SelectedDate = if ($item.date_debut) { [datetime]$item.date_debut } else { $null }
    $TxtRoadmapPhase.Text = [string]$item.phase
    $TxtRoadmapNutrition.Text = [string]$item.nutrition
    $TxtRoadmapPoidsMoyen.Text = [string]$item.poids_moyen
    $TxtRoadmapDepense.Text = [string]$item.depense_calorique
    $TxtRoadmapCardio.Text = [string]$item.cardio_minutes
    $TxtRoadmapPas.Text = [string]$item.pas
    $TxtRoadmapPrecisionTraining.Text = [string]$item.precision_training
    $TxtRoadmapEvenements.Text = [string]$item.evenements
    $TxtRoadmapNotes.Text = [string]$item.notes
})

(Get-Ctrl 'BtnRoadmapNouveau').Add_Click({ Clear-FormRoadmap })

(Get-Ctrl 'BtnRoadmapEnregistrer').Add_Click({
    Invoke-Protege {
        if (-not $CmbSuiviClient.SelectedItem) { Show-Erreur "Sélectionne d'abord un client."; return }
        $semaine = Get-IntOuNull $TxtRoadmapSemaine.Text
        if (-not $semaine) { Show-Erreur "Indique un numéro de semaine."; return }
        $dateDebut = if ($DateRoadmapDebut.SelectedDate) { $DateRoadmapDebut.SelectedDate.ToString('yyyy-MM-dd') } else { $null }
        $roadmapArgs = @{
            SemaineNumero = $semaine; DateDebut = $dateDebut; Phase = (Get-TexteOuNull $TxtRoadmapPhase.Text)
            Nutrition = (Get-TexteOuNull $TxtRoadmapNutrition.Text); PoidsMoyen = (Get-DoubleOuNull $TxtRoadmapPoidsMoyen.Text)
            DepenseCalorique = (Get-DoubleOuNull $TxtRoadmapDepense.Text); CardioMinutes = (Get-DoubleOuNull $TxtRoadmapCardio.Text)
            Pas = (Get-IntOuNull $TxtRoadmapPas.Text); PrecisionTraining = (Get-TexteOuNull $TxtRoadmapPrecisionTraining.Text)
            Evenements = (Get-TexteOuNull $TxtRoadmapEvenements.Text); Notes = (Get-TexteOuNull $TxtRoadmapNotes.Text)
        }
        if ($Script:SelectedRoadmapId) {
            Update-RoadmapSemaine -DbPath $DbPath -Id $Script:SelectedRoadmapId @roadmapArgs
        } else {
            New-RoadmapSemaine -DbPath $DbPath -ClientId $CmbSuiviClient.SelectedItem.id @roadmapArgs | Out-Null
        }
        Update-VueRoadmap
    }
})

(Get-Ctrl 'BtnRoadmapSupprimer').Add_Click({
    Invoke-Protege {
        if (-not $Script:SelectedRoadmapId) { Show-Erreur "Sélectionne une semaine."; return }
        if (Show-Confirmation "Supprimer cette semaine de roadmap ?") {
            Remove-RoadmapSemaine -DbPath $DbPath -Id $Script:SelectedRoadmapId
            Update-VueRoadmap
        }
    }
})

(Get-Ctrl 'BtnTelechargerModeleRoadmap').Add_Click({
    Invoke-Protege {
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = 'Fichier Excel (*.xlsx)|*.xlsx'
        $dialog.FileName = 'Modele_Roadmap.xlsx'
        if ($dialog.ShowDialog()) {
            Export-ModeleRoadmapExcel -Path $dialog.FileName
            Show-Info "Modèle créé. Remplis-le (une ligne par semaine) puis réimporte-le via ""Importer la roadmap remplie...""."
        }
    }
})

(Get-Ctrl 'BtnImporterRoadmap').Add_Click({
    Invoke-Protege {
        if (-not $CmbSuiviClient.SelectedItem) { Show-Erreur "Sélectionne d'abord un client."; return }
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = 'Fichier Excel (*.xlsx)|*.xlsx'
        if ($dialog.ShowDialog()) {
            $res = Import-RoadmapDepuisExcel -DbPath $DbPath -ClientId $CmbSuiviClient.SelectedItem.id -ExcelPath $dialog.FileName
            $message = "Semaines importées : $($res.Importees)`nLignes sans numéro de semaine ignorées : $($res.IgnoreesSansNumero)"
            if ($res.Erreurs.Count -gt 0) { $message += "`n`nRemarques :`n" + ($res.Erreurs -join "`n") }
            Show-Info $message "Import terminé"
            Update-VueRoadmap
        }
    }
})

# --- Journal alimentaire : logique ---

function Update-VueJournalAlimentaire {
    $GridJournalAlimentaire.ItemsSource = $null
    if (-not $CmbSuiviClient.SelectedItem) { return }
    $GridJournalAlimentaire.ItemsSource = @(Get-JournalAlimentaire -DbPath $DbPath -ClientId $CmbSuiviClient.SelectedItem.id)
}

(Get-Ctrl 'BtnImporterJournalAlimentaire').Add_Click({
    Invoke-Protege {
        if (-not $CmbSuiviClient.SelectedItem) { Show-Erreur "Sélectionne d'abord un client."; return }
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = 'Fichier CSV (*.csv)|*.csv'
        if ($dialog.ShowDialog()) {
            $res = Import-JournalAlimentaireDepuisFatSecret -DbPath $DbPath -ClientId $CmbSuiviClient.SelectedItem.id -CsvPath $dialog.FileName
            $message = "Jours importés : $($res.Importes)"
            if ($res.Erreurs.Count -gt 0) { $message += "`n`nRemarques :`n" + ($res.Erreurs -join "`n") }
            Show-Info $message "Import terminé"
            Update-VueJournalAlimentaire
        }
    }
})

(Get-Ctrl 'BtnSupprimerJourJournal').Add_Click({
    Invoke-Protege {
        $item = $GridJournalAlimentaire.SelectedItem
        if (-not $item) { Show-Erreur "Sélectionne une ligne."; return }
        if (Show-Confirmation "Supprimer le journal du $($item.date) ?") {
            Remove-JournalAlimentaireJour -DbPath $DbPath -Id ([int]$item.id)
            Update-VueJournalAlimentaire
        }
    }
})

# ============================== OUTILS ==============================

(Get-Ctrl 'BtnImporterBibliotheques').Add_Click({
    Invoke-Protege {
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = 'Fichier Excel (*.xlsx)|*.xlsx'
        if ($dialog.ShowDialog()) {
            $res = Import-BibliothequesDepuisExcel -DbPath $DbPath -ExcelPath $dialog.FileName
            $message = "Exercices importes : $($res.ExercicesImportes)`nAliments importes : $($res.AlimentsImportes)`nComplements importes : $($res.ComplementsImportes)"
            if ($res.Erreurs.Count -gt 0) {
                $message += "`n`nRemarques :`n" + ($res.Erreurs -join "`n")
            }
            Show-Info $message "Import termine"
            Update-VueBibliotheques
        }
    }
})

(Get-Ctrl 'BtnSauvegarder').Add_Click({
    Invoke-Protege {
        $chemin = Backup-Database -DbPath $DbPath -BackupFolder $BackupFolder
        Show-Info "Sauvegarde creee :`n$chemin"
    }
})

(Get-Ctrl 'TxtVersionInstallee').Text = "Version installee : v$AppVersion"

# ============================== DEMARRAGE ==============================

Update-VueDashboard
$Window.ShowDialog() | Out-Null
