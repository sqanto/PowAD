
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#Form Start
$objForm = New-Object System.Windows.Forms.Form
$objForm.StartPosition = "CenterScreen"
$objForm.Size = New-Object System.Drawing.Size(1000,620)
$objForm.Text = "Edit User"

$objForm.AllowDrop = $true
$objForm.Add_DragEnter($handler)
$objForm.Controls.Add($listBox)

function search

{
add-pssnapin Quest.ActiveRoles.ADManagement
$global:importUser = Get-QADUser $searcheing.Text
}

function uebertrag
{
$objTextBoxV.Text = $global:importUser.FirstName
$objTextBoxN.Text = $global:importUser.Lastname
}


function searchframe
{
$Searchframe = New-Object System.Windows.Forms.Form
$Searchframe.StartPosition = "CenterScreen"
$Searchframe.Size = New-Object System.Drawing.Size(200,200)
$Searchframe.Text = "Search User"


$searchtxt = New-Object System.Windows.Forms.Label
$searchtxt.Location = New-Object System.Drawing.Size(10,10) 
$searchtxt.Size = New-Object System.Drawing.Size(399,20) 
$searchtxt.Text = "User eingeben (Anmeldename)"
$Searchframe.Controls.Add($searchtxt)

$global:searcheing = New-Object System.Windows.Forms.TextBox 
$searcheing.Location = New-Object System.Drawing.Size(10,30) 
$searcheing.Size = New-Object System.Drawing.Size(150,20) 
$Searchframe.Controls.Add($searcheing)

#button

$searchbutto = New-Object System.Windows.Forms.Button
$searchbutto.Location = New-Object System.Drawing.Size(10,70)
$searchbutto.Size = New-Object System.Drawing.Size(75,23)
$searchbutto.Text = "Search"
$searchbutto.Add_Click({. search})  #start funktion
$Searchframe.Controls.Add($searchbutto) 

$uebertrag = New-Object System.Windows.Forms.Button
$uebertrag.Location = New-Object System.Drawing.Size(10,120)
$uebertrag.Size = New-Object System.Drawing.Size(75,23)
$uebertrag.Text = "Übertrag"
$uebertrag.Add_Click({. uebertrag})  #start funktion
$Searchframe.Controls.Add($uebertrag) 


[void] $Searchframe.ShowDialog()

}

#Text
$objLabelOT = New-Object System.Windows.Forms.Label
$objLabelOT.Location = New-Object System.Drawing.Size(10,10) 
$objLabelOT.Size = New-Object System.Drawing.Size(100,20) 
$objLabelOT.Text = "Allgemeine Infos"
$objForm.Controls.Add($objLabelOT)

$objLabelV = New-Object System.Windows.Forms.Label
$objLabelV.Location = New-Object System.Drawing.Size(10,50) 
$objLabelV.Size = New-Object System.Drawing.Size(100,20) 
$objLabelV.Text = "Vorname:"
$objForm.Controls.Add($objLabelV)

$objLabelN = New-Object System.Windows.Forms.Label
$objLabelN.Location = New-Object System.Drawing.Size(10,80) 
$objLabelN.Size = New-Object System.Drawing.Size(100,20) 
$objLabelN.Text = "Nachnahme:"
$objForm.Controls.Add($objLabelN)

$objLabelA = New-Object System.Windows.Forms.Label
$objLabelA.Location = New-Object System.Drawing.Size(10,110) 
$objLabelA.Size = New-Object System.Drawing.Size(100,20) 
$objLabelA.Text = "Anmeldename:" #
$objForm.Controls.Add($objLabelA)

$objLabelB = New-Object System.Windows.Forms.Label
$objLabelB.Location = New-Object System.Drawing.Size(10,140) 
$objLabelB.Size = New-Object System.Drawing.Size(100,20) 
$objLabelB.Text = "Beschreibung:"
$objForm.Controls.Add($objLabelB)

$objLabelBu = New-Object System.Windows.Forms.Label
$objLabelBu.Location = New-Object System.Drawing.Size(10,170) 
$objLabelBu.Size = New-Object System.Drawing.Size(100,20) 
$objLabelBu.Text = "Büro:"
$objForm.Controls.Add($objLabelBu)

$objLabelRu = New-Object System.Windows.Forms.Label
$objLabelRu.Location = New-Object System.Drawing.Size(10,200) 
$objLabelRu.Size = New-Object System.Drawing.Size(100,20) 
$objLabelRu.Text = "Rufnummer:"
$objForm.Controls.Add($objLabelRu)

$objLabelEm = New-Object System.Windows.Forms.Label
$objLabelEm.Location = New-Object System.Drawing.Size(10,230) 
$objLabelEm.Size = New-Object System.Drawing.Size(100,20) 
$objLabelEm.Text = "E-Mail:"
$objForm.Controls.Add($objLabelEm)

$objLabelAb = New-Object System.Windows.Forms.Label
$objLabelAb.Location = New-Object System.Drawing.Size(300,50) 
$objLabelAb.Size = New-Object System.Drawing.Size(100,20) 
$objLabelAb.Text = "Abteilung:"
$objForm.Controls.Add($objLabelAb)

$objLabelPo = New-Object System.Windows.Forms.Label
$objLabelPo.Location = New-Object System.Drawing.Size(300,80) 
$objLabelPo.Size = New-Object System.Drawing.Size(100,20) 
$objLabelPo.Text = "Position:"
$objForm.Controls.Add($objLabelPo)

$objLabelAn = New-Object System.Windows.Forms.Label
$objLabelAn.Location = New-Object System.Drawing.Size(300,80) 
$objLabelAn.Size = New-Object System.Drawing.Size(100,20) 
$objLabelAn.Text = "Anzeigename:"#Anzeigename
$objForm.Controls.Add($objLabelAn)

$objLabelP = New-Object System.Windows.Forms.Label
$objLabelP.Location = New-Object System.Drawing.Size(300,110) 
$objLabelP.Size = New-Object System.Drawing.Size(100,20) 
$objLabelP.Text = "PräWindows:"
$objForm.Controls.Add($objLabelP)

$objLabelAs = New-Object System.Windows.Forms.Label
$objLabelAs.Location = New-Object System.Drawing.Size(300,140) 
$objLabelAs.Size = New-Object System.Drawing.Size(100,20) 
$objLabelAs.Text = "Anmeldescript:"
$objForm.Controls.Add($objLabelAs)

$objLabelPp = New-Object System.Windows.Forms.Label
$objLabelPp.Location = New-Object System.Drawing.Size(300,170) 
$objLabelPp.Size = New-Object System.Drawing.Size(100,20) 
$objLabelPp.Text = "Profilpfad:"
$objForm.Controls.Add($objLabelPp)

$objLabelK = New-Object System.Windows.Forms.Label
$objLabelK.Location = New-Object System.Drawing.Size(300,200) 
$objLabelK.Size = New-Object System.Drawing.Size(100,20) 
$objLabelK.Text = "Kostenstelle:"
$objForm.Controls.Add($objLabelK)

$objLabelAna = New-Object System.Windows.Forms.Label
$objLabelAna.Location = New-Object System.Drawing.Size(300,230) 
$objLabelAna.Size = New-Object System.Drawing.Size(100,20) 
$objLabelAna.Text = "Anzeigename:"
$objForm.Controls.Add($objLabelAna)

$objLabelPat = New-Object System.Windows.Forms.Label
$objLabelPat.Location = New-Object System.Drawing.Size(600,50) 
$objLabelPat.Size = New-Object System.Drawing.Size(100,20) 
$objLabelPat.Text = "Path:"
$objForm.Controls.Add($objLabelPat)

$objLabelAV = New-Object System.Windows.Forms.Label
$objLabelAV.Location = New-Object System.Drawing.Size(600,80) 
$objLabelAV.Size = New-Object System.Drawing.Size(100,20) 
$objLabelAV.Text = "AVP:"
$objForm.Controls.Add($objLabelAV)

$objLabelPnu = New-Object System.Windows.Forms.Label
$objLabelPnu.Location = New-Object System.Drawing.Size(600,110) 
$objLabelPnu.Size = New-Object System.Drawing.Size(100,20) 
$objLabelPnu.Text = "Personalnummer:"
$objForm.Controls.Add($objLabelPnu)

$objLabelKgr = New-Object System.Windows.Forms.Label
$objLabelKgr.Location = New-Object System.Drawing.Size(600,140) 
$objLabelKgr.Size = New-Object System.Drawing.Size(100,20) 
$objLabelKgr.Text = "Kostenstellgrp:"
$objForm.Controls.Add($objLabelKgr)

$objLabelFn = New-Object System.Windows.Forms.Label
$objLabelFn.Location = New-Object System.Drawing.Size(600,170) 
$objLabelFn.Size = New-Object System.Drawing.Size(100,20) 
$objLabelFn.Text = "Firmennummer:"
$objForm.Controls.Add($objLabelFn)

$objLabelAk = New-Object System.Windows.Forms.Label
$objLabelAk.Location = New-Object System.Drawing.Size(600,200) 
$objLabelAk.Size = New-Object System.Drawing.Size(100,20) 
$objLabelAk.Text = "Ak-Gruppe:"
$objForm.Controls.Add($objLabelAk)

#Texteingabe

$objTextBoxV = New-Object System.Windows.Forms.TextBox 
$objTextBoxV.Location = New-Object System.Drawing.Size(110,50) 
$objTextBoxV.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxV)

$objTextBoxN = New-Object System.Windows.Forms.TextBox 
$objTextBoxN.Location = New-Object System.Drawing.Size(110,80) 
$objTextBoxN.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxN)

$objTextBoxA = New-Object System.Windows.Forms.TextBox 
$objTextBoxA.Location = New-Object System.Drawing.Size(110,110) 
$objTextBoxA.Size = New-Object System.Drawing.Size(150,20) 
$objTextBoxA.Add_Click({. CrAUser})
$objForm.Controls.Add($objTextBoxA)

$objTextBoxB = New-Object System.Windows.Forms.TextBox 
$objTextBoxB.Location = New-Object System.Drawing.Size(110,140) 
$objTextBoxB.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxB)

$objTextBoxBu = New-Object System.Windows.Forms.TextBox 
$objTextBoxBu.Location = New-Object System.Drawing.Size(110,170) 
$objTextBoxBu.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxBu)

$objTextBoxRu = New-Object System.Windows.Forms.TextBox 
$objTextBoxRu.Location = New-Object System.Drawing.Size(110,200) 
$objTextBoxRu.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxRu)

$objTextBoxEm = New-Object System.Windows.Forms.TextBox 
$objTextBoxEm.Location = New-Object System.Drawing.Size(110,230) 
$objTextBoxEm.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxEm)

$objTextBoxAb = New-Object System.Windows.Forms.TextBox 
$objTextBoxAb.Location = New-Object System.Drawing.Size(410,50) 
$objTextBoxAb.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxAb)

$objTextBoxPo = New-Object System.Windows.Forms.TextBox 
$objTextBoxPo.Location = New-Object System.Drawing.Size(410,80) 
$objTextBoxPo.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxPo)

$objTextBoxAn = New-Object System.Windows.Forms.TextBox 
$objTextBoxAn.Location = New-Object System.Drawing.Size(410,110) 
$objTextBoxAn.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxAn)

$objTextBoxAs = New-Object System.Windows.Forms.TextBox 
$objTextBoxAs.Location = New-Object System.Drawing.Size(410,140) 
$objTextBoxAs.Size = New-Object System.Drawing.Size(150,20) 
$objTextBoxAs.Text = "logon.vbs"
$objForm.Controls.Add($objTextBoxAs)

$objTextBoxPp = New-Object System.Windows.Forms.TextBox 
$objTextBoxPp.Location = New-Object System.Drawing.Size(410,170) 
$objTextBoxPp.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxPp)

$objTextBoxK = New-Object System.Windows.Forms.TextBox 
$objTextBoxK.Location = New-Object System.Drawing.Size(410,200) 
$objTextBoxK.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxK)

$objTextBoxAna = New-Object System.Windows.Forms.TextBox 
$objTextBoxAna.Location = New-Object System.Drawing.Size(410,230) 
$objTextBoxAna.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxAna)

$objCombobox = New-Object System.Windows.Forms.Combobox 
$objCombobox.Location = New-Object System.Drawing.Size(710,50) 
$objCombobox.Size = New-Object System.Drawing.Size(260,20) 

[void] $objCombobox.Items.Add("OU=Domain User,DC=berlin,DC=strato,DC=de")
[void] $objCombobox.Items.Add("OU=Test User OU,OU=Domain User,DC=berlin,DC=strato,DC=de")
[void] $objCombobox.Items.Add("Item 3")
[void] $objCombobox.Items.Add("Item 4")
[void] $objCombobox.Items.Add("Item 5")
$objCombobox.SelectedItem = "OU=Domain User,DC=berlin,DC=strato,DC=de"

$objForm.Controls.Add($objCombobox) 

$objTextBoxAV = New-Object System.Windows.Forms.TextBox 
$objTextBoxAV.Location = New-Object System.Drawing.Size(710,80) 
$objTextBoxAV.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxAV)

$objTextBoxPnu = New-Object System.Windows.Forms.TextBox 
$objTextBoxPnu.Location = New-Object System.Drawing.Size(710,110) 
$objTextBoxPnu.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxPnu)

$objTextBoxKgr = New-Object System.Windows.Forms.TextBox 
$objTextBoxKgr.Location = New-Object System.Drawing.Size(710,140) 
$objTextBoxKgr.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxKgr)

$objTextBoxFn = New-Object System.Windows.Forms.TextBox 
$objTextBoxFn.Location = New-Object System.Drawing.Size(710,170) 
$objTextBoxFn.Size = New-Object System.Drawing.Size(150,20) 
$objTextBoxFn.Text = "111"
$objForm.Controls.Add($objTextBoxFn)

$objTextBoxAk = New-Object System.Windows.Forms.TextBox 
$objTextBoxAk.Location = New-Object System.Drawing.Size(710,200) 
$objTextBoxAk.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxAk)


#Button
$global:NUserButton = New-Object System.Windows.Forms.Button
$NUserButton.Location = New-Object System.Drawing.Size(600,230)
$NUserButton.Size = New-Object System.Drawing.Size(75,23)
$NUserButton.Text = "Edit User"
$NUserButton.Enabled = $false
$NUserButton.Visible = $false
$NUserButton.Add_Click({. NewUser})  #start funktion
$objForm.Controls.Add($NUserButton) 

$CheckallButton = New-Object System.Windows.Forms.Button
$CheckallButton.Location = New-Object System.Drawing.Size(700,230)
$CheckallButton.Size = New-Object System.Drawing.Size(75,23)
$CheckallButton.Text = "Check All"
$CheckallButton.Add_Click({. check})  #start funktion
$objForm.Controls.Add($CheckallButton) 

$CheckallButton = New-Object System.Windows.Forms.Button
$CheckallButton.Location = New-Object System.Drawing.Size(700,400)
$CheckallButton.Size = New-Object System.Drawing.Size(75,23)
$CheckallButton.Text = "Search User"
$CheckallButton.Add_Click({. searchframe})  #start funktion
$objForm.Controls.Add($CheckallButton) 

#Feature L
$objLabelAk = New-Object System.Windows.Forms.Label
$objLabelAk.Location = New-Object System.Drawing.Size(10,300) 
$objLabelAk.Size = New-Object System.Drawing.Size(200,20) 
$objLabelAk.Text = "Feature"
$objForm.Controls.Add($objLabelAk)

## Festure on/of
$objTypeCheckboxAg = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxAg.Location = New-Object System.Drawing.Size(10,350) 
$objTypeCheckboxAg.Size = New-Object System.Drawing.Size(200,20)
$objTypeCheckboxAg.Text = "Account gesperrt"
$objForm.Controls.Add($objTypeCheckboxAg)

$objTypeCheckboxKae = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxKae.Location = New-Object System.Drawing.Size(10,380) 
$objTypeCheckboxKae.Size = New-Object System.Drawing.Size(200,20)
$objTypeCheckboxKae.Text = "Kennwort muß geändert werden"
$objForm.Controls.Add($objTypeCheckboxKae)

$objTypeCheckboxKlna = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxKlna.Location = New-Object System.Drawing.Size(10,410) 
$objTypeCheckboxKlna.Size = New-Object System.Drawing.Size(200,20)
$objTypeCheckboxKlna.Text = "Kennwort läuft nie ab"
$objForm.Controls.Add($objTypeCheckboxKlna)

$objTypeCheckboxKna = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxKna.Location = New-Object System.Drawing.Size(10,440) 
$objTypeCheckboxKna.Size = New-Object System.Drawing.Size(270,20)
$objTypeCheckboxKna.Text = "Kennwort kann nicht geändert werden werden"
$objForm.Controls.Add($objTypeCheckboxKna)

$objTypeCheckboxExchange = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxExchange.Location = New-Object System.Drawing.Size(10,470) 
$objTypeCheckboxExchange.Size = New-Object System.Drawing.Size(200,20)
$objTypeCheckboxExchange.Text = "E- Postfach erstellen"
$objForm.Controls.Add($objTypeCheckboxExchange)

$objLabelABL = New-Object System.Windows.Forms.Label
$objLabelABL.Location = New-Object System.Drawing.Size(310,300) 
$objLabelABL.Size = New-Object System.Drawing.Size(100,20) 
$objLabelABL.Text = "Ablauf"
$objForm.Controls.Add($objLabelABL)

$objTypeCheckboxnie = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxnie.Location = New-Object System.Drawing.Size(310,350) 
$objTypeCheckboxnie.Size = New-Object System.Drawing.Size(40,20)
$objTypeCheckboxnie.Text = "Nie"
$objForm.Controls.Add($objTypeCheckboxnie)

$objTypeCheckboxam = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxam.Location = New-Object System.Drawing.Size(370,350) 
$objTypeCheckboxam.Size = New-Object System.Drawing.Size(200,20)
$objTypeCheckboxam.Text = "Am:"
$objForm.Controls.Add($objTypeCheckboxam)






$objLabelFeature = New-Object System.Windows.Forms.Label
$objLabelFeature.Location = New-Object System.Drawing.Size(10,570) 
$objLabelFeature.Size = New-Object System.Drawing.Size(400,40) 
$objLabelFeature.Text = "Fehlde Feature: Position  :  account sperren läüft ab : exchange " #
$objForm.Controls.Add($objLabelFeature)



. calender



[void] $objForm.ShowDialog()