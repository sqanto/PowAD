#Create User


[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#Default

$global:confpath = $env:USERPROFILE + "\AppData\Roaming\"
$global:confpath2 = $confpath + "Psadmin"

# All Function
function Import-Excel
{
  param (
    [string]$FileName,
    [string]$WorksheetName,
    [bool]$DisplayProgress = $true
  )

  if ($FileName -eq "") {
    throw "Please provide path to the Excel file"
    Exit
  }

  if (-not (Test-Path $FileName)) {
    throw "Path '$FileName' does not exist."
    exit
  }

  $FileName = Resolve-Path $FileName
  $excel = New-Object -com "Excel.Application"
  $excel.Visible = $false
  $workbook = $excel.workbooks.open($FileName)

  if (-not $WorksheetName) {
    Write-Warning "Defaulting to the first worksheet in workbook."
    $sheet = $workbook.ActiveSheet
  } else {
    $sheet = $workbook.Sheets.Item($WorksheetName)
  }
  
  if (-not $sheet)
  {
    throw "Unable to open worksheet $WorksheetName"
    exit
  }
  
  $sheetName = $sheet.Name
  $columns = $sheet.UsedRange.Columns.Count
  $lines = $sheet.UsedRange.Rows.Count
  
  Write-Warning "Worksheet $sheetName contains $columns columns and $lines lines of data"
  
  $fields = @()
  
  for ($column = 1; $column -le $columns; $column ++) {
    $fieldName = $sheet.Cells.Item.Invoke(1, $column).Value2
    if ($fieldName -eq $null) {
      $fieldName = "Column" + $column.ToString()
    }
    $fields += $fieldName
  }
  
  $line = 2
  
  
  for ($line = 2; $line -le $lines; $line ++) {
    $values = New-Object object[] $columns
    for ($column = 1; $column -le $columns; $column++) {
      $values[$column - 1] = $sheet.Cells.Item.Invoke($line, $column).Value2
    }  
  
    $row = New-Object psobject
    $fields | foreach-object -begin {$i = 0} -process {
      $row | Add-Member -MemberType noteproperty -Name $fields[$i] -Value $values[$i]; $i++
    }
    $row
    $percents = [math]::round((($line/$lines) * 100), 0)
    if ($DisplayProgress) {
      Write-Progress -Activity:"Importing from Excel file $FileName" -Status:"Imported $line of total $lines lines ($percents%)" -PercentComplete:$percents
    }
  }
  $workbook.Close()
  $excel.Quit()
}

function creatmail
{
$global:email = $Anmeldename + "@strato.de"
$objTextBoxEm.Text = $global:email 

$global:emailc = $Anmeldename + "@strato.com"


$objListboxpraddresse =  New-Object System.Windows.Forms.Combobox
$objListboxpraddresse.Location = New-Object System.Drawing.Size(290,350) 
$objListboxpraddresse.Size = New-Object System.Drawing.Size(200,10) 
[void] $objListboxpraddresse.Items.Add($global:emailc)
[void] $objListboxpraddresse.Items.Add($global:email)



$objListboxpraddresse.Height = 70
$objForm.Controls.Add($objListboxpraddresse) 

}

function changeu
{
## user neustarten

$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$wshell.Popup("Hiermit erstellen sie ein User mit einem anderen User ")

$global:credi = Get-Credential
Start-Job -ScriptBlock { . NewUser } -Credential $credi

}

function createconfig ## button erstellen
{
   if (Test-Path $confpath2)
    {
     $global:StatusF = "Config ist bereits erstellt"
    }


        else
        {
        cd $confpath
        mkdir "Psadmin"
        cd $confpath2
        new-item -name config.pa -type "file"
        $global:StatusF = "Config wurde erstellt"
        }

$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$wshell.Popup($global:StatusF)
}

function settings #alpha
{
$objFormsett = New-Object System.Windows.Forms.Form
$objFormsett.StartPosition = "CenterScreen"
$objFormsett.Size = New-Object System.Drawing.Size(400,400)
$objFormsett.Text = "Settings"



$Createconfigbot = New-Object System.Windows.Forms.Button
$Createconfigbot.Location = New-Object System.Drawing.Size(10,30)
$Createconfigbot.Size = New-Object System.Drawing.Size(75,23)
$Createconfigbot.Text = "Create Config"
$Createconfigbot.Add_Click({. createconfig})  #start funktion
$objFormsett.Controls.Add($Createconfigbot) 

$saveconfigbot = New-Object System.Windows.Forms.Button
$saveconfigbot.Location = New-Object System.Drawing.Size(10,60)
$saveconfigbot.Size = New-Object System.Drawing.Size(75,23)
$saveconfigbot.Text = "Save Config"
$saveconfigbot.Add_Click({})  #start funktion
$objFormsett.Controls.Add($saveconfigbot) 

$loadeconfigbot = New-Object System.Windows.Forms.Button
$loadeconfigbot.Location = New-Object System.Drawing.Size(10,90)
$loadeconfigbot.Size = New-Object System.Drawing.Size(75,23)
$loadeconfigbot.Text = "Load Config"
$loadeconfigbot.Add_Click({ })  #start funktion
$objFormsett.Controls.Add($loadeconfigbot) 



[void] $objFormsett.ShowDialog()

}

function loadconfig
{
cd $confpath2

$global:sing = Get-Content config.pa

$global:creduser = $global:sing[0]
$global:passwo =  $globalsing[1]

}

function saveconfig
{
cd $confpath2
Set-Content -Path config.pa -Value "sds/n22"

}

function load
{
$search = $objListbox2.SelectedItem
$e=0
do{

    if($data[$e] -eq $search)
    { 
     $objTextBoxN.Text = $import[$e].Name
     $objTextBoxV.Text = $import[$e].Vorname
     $objTextBoxAb.Text = $import[$e].Abteilung
     $objTextBoxK.Text = $import[$e].Kostenstelle
     $objTextBoxPnu.Text = $import[$e]."Personal-Nummer"
     $objTextBoxB.Text = $import[$e].Kommentar
      $objCalendar.SelectionStart= $import[$e].Austritt

      



    # $objCalendar.SelectionStart = $import[$e].Austritt
     $e = $data.Count - 1
     . CrAUser
     . creatmail
     . ppath
     . praewindows
     . exdb
     }

$e++
}while($e -le $data.Count)

    }

function naut
{
    if (Test-path V:\Office_IT\Listen\)
    {
    $Zcheck++
    }
        else
        {
               $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
               $wshell.Popup("Algemeiner Fehler : Path nicht gefunden")
        }

}

function createfolder
{

new-item -path "\\na32401\users\" -name $Anmeldename -type directory 

$Right="Modify"
$path=$objTextBoxPp.text
$Principal="Stratobln\$Anmeldename"
$rule=New-Object System.Security.AccessControl.FileSystemAccessRule($Principal,$Right,3, 0,"Allow")
$acl = get-acl $path
$acl.SetAccessRule($rule)
set-acl $path $acl

}

function excel
{

$objListbox.SelectedItem


if ($objListbox.SelectedItem -eq $null)
{
$global:import = import-excel V:\Office_IT\Listen\Mitarbeiter-Eing.xlsx
}
    else
    {
    $global:import = import-excel $objListbox.SelectedItem
    }


$global:data = $import.Name  
$data.count


$e=0
do{

$sting = $data[$e]


$objListbox2.Items.Add($sting)
$e++


}while($e -le $data.count)



}

function umlaut
{


$username = $global:Anmeldename 


$1 = "ä"

$2 = "ö"

$3 = "ü"

$4 = "ae"

$5 = "oe"

$6 = "ue"




$username = $username -replace($1,$4)
$username = $username -replace($2,$5)
$username = $username -replace($3,$6)

$global:Anmeldename = $username


}

function CrAUser
{
$global:Anmeldename = $objTextBoxN.Text.ToLower()


. umlaut
. creatmail
. praewindows
. ppath
. exdb


$objTextBoxA.Text = $Anmeldename

$global:Anzeigename = $objTextBoxV.Text + " " + $objTextBoxN.Text

$objTextBoxAna.Text = $Anzeigename


}

function ppath
{
$objTextBoxPp.text = "\\na32401\users\" + $Anmeldename

}

function creexchange
{

$UserCredential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch2010-a.berlin.strato.de/PowerShell/ -Authentication Kerberos -Credential $UserCredential

Import-PSSession $Session

Enable-Mailbox  -Identity  $objTextBoxAn.Text -Database $Database 

Wait-Event -Timeout 10

if ($objTypeCheckboxAdress.CheckState -eq $true)
        {
        Set-Mailbox $Anmeldename -HiddenFromAddressListsEnabled $true -EmailAddressPolicyEnabled $false -EmailAddresses $global:email 
        }

        else
        {
        Set-Mailbox $Anmeldename -EmailAddressPolicyEnabled $false -EmailAddresses $global:email 
        }

    ##actuive sync
    if ($objTypeCheckboxAs.CheckState -eq $true)

            {

            Set-CASMailbox -Identity $Anmeldename -ActiveSyncEnabled $true
            
            #setze active sync
            }

                else{
                    Set-CASMailbox -Identity $Anmeldename -ActiveSyncEnabled $false
                }


                ##owa 
                if ($objTypeCheckboxowa.CheckState -eq $true)
                {
                Set-CASMailbox -Identity $Anmeldename  -OWAEnabled $true
                
            
                #setze owa
                }
                    else{
                    
                    Set-CASMailbox -Identity $Anmeldename  -OWAEnabled $false
                    }



}

$handler = {
if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) {
$objListbox.Items.Add($filename)
}
}
}

function NewUser
{
 #fehlt Position  account sperren läüft ab 


   # Set-QADUser -Identity $objTextBoxA.Text -Department $objTextBoxAb.Text -PasswordNeverExpires $objTypeCheckboxKlna.Checked -UserMustChangePassword $objTypeCheckboxKae.Checked -ProfilePath $objTextBoxPp.Text

    if($objTypeCheckboxam.Checked -eq $true )
         {
         #ablauf set


         $doto = $dattum.ToFileTime()
         New-QADUser -ParentContainer $objCombobox.SelectedItem -SamAccountName $objTextBoxA.Text -UserPrincipalName $objTextBoxA.Text -Name $global:Anzeigename  -UserPassword "Strato2014" -FirstName $objTextBoxV.Text -Title $objTextBoxPo.text -LastName $objTextBoxN.Text -Description $objTextBoxB.Text -Office $objTextBoxBu.Text -HomePhone $objTextBoxRu.Text -Email $objTextBoxEm.Text -LogonScript $objTextBoxAs.Text -Department $objTextBoxAb.Text -ObjectAttributes @{extensionAttribute3=$objTextBoxFn.Text;employeeID=$objTextBoxAV.Text;extensionAttribute4=$objTextBoxAk.Text;employeeType=$objTextBoxKgr.Text;} 
      $objTextBoxA.Text |  Set-QADUser -AccountExpires $dattum
     # $objTextBoxA.Text | Add-QADGroupMember
    
         }
             else
             {

                New-QADUser -ParentContainer $objCombobox.SelectedItem -SamAccountName $objTextBoxA.Text -UserPrincipalName $objTextBoxA.Text -Name $global:Anzeigename  -UserPassword "Strato2014" -FirstName $objTextBoxV.Text -Title $objTextBoxPo.text -LastName $objTextBoxN.Text -Description $objTextBoxB.Text -Office $objTextBoxBu.Text -HomePhone $objTextBoxRu.Text -Email $objTextBoxEm.Text -LogonScript $objTextBoxAs.Text -Department $objTextBoxAb.Text -ObjectAttributes @{extensionAttribute3=$objTextBoxFn.Text;employeeID=$objTextBoxAV.Text;extensionAttribute4=$objTextBoxAk.Text;employeeType=$objTextBoxKgr.Text;} 
 
             }


Wait-Event -Timeout 5

    if ($objTypeCheckboxExchange.Checked -eq $true)
    {
    . creexchange
    }
   
    . createfolder

}


function

function exdb
{


$global:FUser = $objTextBoxN.Text.Substring(0,1)

    switch ($FUser) 
        { 

    "A" {$global:Database = "Mailbox_DB1"}
    "B" {$Database = "Mailbox_DB1"}
    "C" {$Database = "Mailbox_DB1"}
    "D" {$Database = "Mailbox_DB1"}
    "E" {$Database = "Mailbox_DB1"}
    "F" {$Database = "Mailbox_DB1"}
    "G" {$Database = "Mailbox_DB1"}
    "H" {$Database = "Mailbox_DB1"}
    "I" {$Database = "Mailbox_DB1"}
    "J" {$Database = "Mailbox_DB1"}
    "K" {$Database = "Mailbox_DB1"}
    "L" {$Database = "Mailbox_DB1"}
    "M" {$Database = "Mailbox_DB1"}
    "N" {$Database = "Mailbox_DB2"}
    "O" {$Database = "Mailbox_DB2"}
    "Ö" {$Database = "Mailbox_DB2"}
    "P" {$Database = "Mailbox_DB2"}
    "Q" {$Database = "Mailbox_DB2"}
    "R" {$Database = "Mailbox_DB2"}
    "S" {$Database = "Mailbox_DB2"}
    "T" {$Database = "Mailbox_DB2"}
    "U" {$Database = "Mailbox_DB2"}
    "V" {$Database = "Mailbox_DB2"}
    "W" {$Database = "Mailbox_DB2"}
    "X" {$Database = "Mailbox_DB2"}
    "Y" {$Database = "Mailbox_DB2"}
    "Z" {$Database = "Mailbox_DB2"}
    default  {$Database = "Mailbox_DB2"}
      }


  $objTextBoxEDB.Text = $Database

}

function praewindows
{
$objTextBoxAn.text = "STRATOBLN\" + $Anmeldename
}

function check
{
$Zcheck = 0
add-pssnapin Quest.ActiveRoles.ADManagement
$usergive = Get-QADUser $objTextBoxA.Text
    if ($usergive -eq $null)
    {
        #all ok

        $Zcheck ++
        
    }
        else{

        $objLabelA.Text = "Anmeldename !!!"
        $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
        $wshell.Popup("Anmeldename schon vorhanden bitte änden")
    
             }

            if($objTypeCheckboxam.Checked -eq $objTypeCheckboxnie.Checked)
            {
                      $objLabelABL.Text = "Ablauf !!! "
                      $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
                      $wshell.Popup("Ablauf Einstellungen bitte prüfen")
            }
                else{
                        $Zcheck ++
                    }

       if($Zcheck -eq 2)
       {
           $NUserButton.Enabled = $true
           $NUserButton.Visible = $true
           $ChangeUButton.Enabled = $true
           $ChangeUButton.Visible = $true
           $global:dattum = $objCalendar.SelectionStart
           #$dattum.ToFileTime() 
           $objTypeCheckboxam.Text = "AM: $dattum "
           . creatmail
           . praewindows
           . ppath

       }
   }

function calender
{
$objCalendar = New-Object System.Windows.Forms.MonthCalendar 
$objCalendar.ShowTodayCircle = $False
$objCalendar.MaxSelectionCount = 1
$objCalendar.Location = New-Object System.Drawing.Size(500,400) 
$objForm.Controls.Add($objCalendar) 
#$objCalendar.SelectionStart

}

function egg
{
    if ($objTextBoxV.Text -eq "Progger")
    {
        $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
        $wshell.Popup("Programmiert von Björn Sommer ")
    }

}

#Size from Main Windows
$objForm = New-Object System.Windows.Forms.Form
$objForm.StartPosition = "CenterScreen"
$objForm.Size = New-Object System.Drawing.Size(1000,800)
$objForm.Text = "Create User"

$objForm.AllowDrop = $true
$objForm.Add_DragEnter($handler)
$objForm.Controls.Add($listBox)

$objForm.Icon = "V:\Office_IT\Entwicklung\Usermanager\Usermanager\WindowsFormsApplication2\strato-icon.ico"

$objForm.Add_Click({. egg})

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
$objLabelEm.Text = "Alias:"
$objForm.Controls.Add($objLabelEm)

$objLabelED = New-Object System.Windows.Forms.Label
$objLabelED.Location = New-Object System.Drawing.Size(10,260) 
$objLabelED.Size = New-Object System.Drawing.Size(100,20) 
$objLabelED.Text = "Exch Datenbank:"
$objForm.Controls.Add($objLabelED)

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

$objTextBoxEDB = New-Object System.Windows.Forms.TextBox 
$objTextBoxEDB.Location = New-Object System.Drawing.Size(110,260) 
$objTextBoxEDB.Size = New-Object System.Drawing.Size(150,20) 
$objForm.Controls.Add($objTextBoxEDB)

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
$objTextBoxAna.Add_Click({. CrAUser})
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


$objListboxGrp = New-Object System.Windows.Forms.Listbox 
$objListboxGrp.Location = New-Object System.Drawing.Size(290,430) 
$objListboxGrp.Size = New-Object System.Drawing.Size(180,20) 
$objListboxGrp.SelectionMode = "MultiExtended"

[void] $objListboxGrp.Items.Add("OI_OFFICE_IT_c")
[void] $objListboxGrp.Items.Add("strato_cno-it")
[void] $objListboxGrp.Items.Add("3. Gruppe")
[void] $objListboxGrp.Items.Add("4. Gruppe")
[void] $objListboxGrp.Items.Add("5. Gruppe")

$objListboxGrp.Height = 70
$objForm.Controls.Add($objListboxGrp) 


#Button
$global:NUserButton = New-Object System.Windows.Forms.Button
$NUserButton.Location = New-Object System.Drawing.Size(600,230)
$NUserButton.Size = New-Object System.Drawing.Size(75,23)
$NUserButton.Text = "Create User"
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
$CheckallButton.Text = "Load List"
$CheckallButton.Add_Click({. excel})  #start funktion
$objForm.Controls.Add($CheckallButton) 

$LoadllButton = New-Object System.Windows.Forms.Button
$LoadllButton.Location = New-Object System.Drawing.Size(700,600)
$LoadllButton.Size = New-Object System.Drawing.Size(75,23)
$LoadllButton.Text = "Load"
$LoadllButton.Add_Click({. load})  #start funktion
$objForm.Controls.Add($LoadllButton) 


$ChangeUButton = New-Object System.Windows.Forms.Button
$ChangeUButton.Location = New-Object System.Drawing.Size(600,270)
$ChangeUButton.Size = New-Object System.Drawing.Size(75,30)
$ChangeUButton.Text = "Create User mit User..."
$ChangeUButton.Enabled = $false
$ChangeUButton.Visible = $false
$ChangeUButton.Add_Click({. changeu})  #start funktion
#$objForm.Controls.Add($ChangeUButton) 

$SettingsButton = New-Object System.Windows.Forms.Button
$SettingsButton.Location = New-Object System.Drawing.Size(800,600)
$SettingsButton.Size = New-Object System.Drawing.Size(75,23)
$SettingsButton.Text = "Settings"
$SettingsButton.Add_Click({. settings})  #start funktion
$objForm.Controls.Add($SettingsButton) 



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
$objTypeCheckboxKna.Text = "Kennwort kann nicht geändert werden"
$objForm.Controls.Add($objTypeCheckboxKna)

$objTypeCheckboxExchange = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxExchange.Location = New-Object System.Drawing.Size(10,470) 
$objTypeCheckboxExchange.Size = New-Object System.Drawing.Size(200,20)
$objTypeCheckboxExchange.Text = "E- Postfach erstellen"
$objForm.Controls.Add($objTypeCheckboxExchange)

$objTypeCheckboxAdress = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxAdress.Location = New-Object System.Drawing.Size(10,500) 
$objTypeCheckboxAdress.Size = New-Object System.Drawing.Size(200,20)
$objTypeCheckboxAdress.Text = "Im Addressbuch verstecken"
$objForm.Controls.Add($objTypeCheckboxAdress)

$objTypeCheckboxAs = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxAs.Location = New-Object System.Drawing.Size(10,530) 
$objTypeCheckboxAs.Size = New-Object System.Drawing.Size(200,20)
$objTypeCheckboxAs.Text = "Active Sync"
$objForm.Controls.Add($objTypeCheckboxAs)

$objTypeCheckboxowa = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxowa.Location = New-Object System.Drawing.Size(10,560) 
$objTypeCheckboxowa.Size = New-Object System.Drawing.Size(200,20)
$objTypeCheckboxowa.Text = "OWA"
$objForm.Controls.Add($objTypeCheckboxowa)

$objLabelPada = New-Object System.Windows.Forms.Label
$objLabelPada.Location = New-Object System.Drawing.Size(290,330) 
$objLabelPada.Size = New-Object System.Drawing.Size(300,20) 
$objLabelPada.Text = "Primäre Adresse bestimmen:"
$objForm.Controls.Add($objLabelPada)

#$CheckboSekcreat = New-Object System.Windows.Forms.Checkbox 
#$CheckboSekcreat.Location = New-Object System.Drawing.Size(290.700) 
#$CheckboSekcreat.Size = New-Object System.Drawing.Size(200,20)
#$CheckboSekcreat.Text = "Seks"
#$objForm.Controls.Add($CheckboSekcreat)





#$objTypeCheckboxDead = New-Object System.Windows.Forms.Checkbox 
#$objTypeCheckboxDead.Location = New-Object System.Drawing.Size(300,380) 
#$objTypeCheckboxDead.Size = New-Object System.Drawing.Size(200,20)
#$objTypeCheckboxDead.Text = "Create %@strato.com"
#$objForm.Controls.Add($objTypeCheckboxDead)

$objLabelABL = New-Object System.Windows.Forms.Label
$objLabelABL.Location = New-Object System.Drawing.Size(500,300) 
$objLabelABL.Size = New-Object System.Drawing.Size(100,20) 
$objLabelABL.Text = "Ablauf"
$objForm.Controls.Add($objLabelABL)

$objTypeCheckboxnie = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxnie.Location = New-Object System.Drawing.Size(500,350) 
$objTypeCheckboxnie.Size = New-Object System.Drawing.Size(40,20)
$objTypeCheckboxnie.Text = "Nie"
$objForm.Controls.Add($objTypeCheckboxnie)

$objTypeCheckboxam = New-Object System.Windows.Forms.Checkbox 
$objTypeCheckboxam.Location = New-Object System.Drawing.Size(600,350) 
$objTypeCheckboxam.Size = New-Object System.Drawing.Size(100,20)
$objTypeCheckboxam.Text = "Am:"
$objForm.Controls.Add($objTypeCheckboxam)






$objLabelFeature = New-Object System.Windows.Forms.Label
$objLabelFeature.Location = New-Object System.Drawing.Size(10,580) 
$objLabelFeature.Size = New-Object System.Drawing.Size(400,40) 
$objLabelFeature.Text = "Bugs :  account sperren;; Config file " #
$objForm.Controls.Add($objLabelFeature)



########################




$objListbox = New-Object System.Windows.Forms.Listbox 
$objListbox.Location = New-Object System.Drawing.Size(700,300) 
$objListbox.Size = New-Object System.Drawing.Size(200,100) 
$objForm.Controls.Add($objListbox)


########################


$objListbox2 = New-Object System.Windows.Forms.Listbox 
$objListbox2.Location = New-Object System.Drawing.Size(10,650) 
$objListbox2.Size = New-Object System.Drawing.Size(950,100) 
$objForm.Controls.Add($objListbox2)

. calender



[void] $objForm.ShowDialog()