##Haupt Men:
#Funktion Verzeweigung zu User erstellen , User ändern

function create
{
powershell.exe .\Create-User.ps1
}

function edit
{
powershell.exe .\Edit-User.ps1
}




[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")







#Size from Main Windows
$objForm = New-Object System.Windows.Forms.Form
$objForm.StartPosition = "CenterScreen"
$objForm.Size = New-Object System.Drawing.Size(400,400)

#Buttons 1

$Create = New-Object System.Windows.Forms.Button
$Create.Location = New-Object System.Drawing.Size(75,120)
$Create.Size = New-Object System.Drawing.Size(75,23)
$Create.Text = "Create User"
$Create.Add_Click({. create})
$objForm.Controls.Add($Create) 

# Button 2

$Edit = New-Object System.Windows.Forms.Button
$Edit.Location = New-Object System.Drawing.Size(200,120)
$Edit.Size = New-Object System.Drawing.Size(75,23)
$Edit.Text = "Edit User"
$Edit.Add_Click({. edit})
$objForm.Controls.Add($Edit) 

# Button 3
$Edit = New-Object System.Windows.Forms.Button
$Edit.Location = New-Object System.Drawing.Size(325,120)
$Edit.Size = New-Object System.Drawing.Size(75,23)
$Edit.Text = "User Jail"
$Edit.Add_Click({$x="OK geklickt";$objForm.Close()})
$objForm.Controls.Add($Edit) 


[void] $objForm.ShowDialog()