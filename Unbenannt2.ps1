$objConnection = New-Object -comobject ADODB.Connection
$Connection =  New-Object -ComObject ADODB.Recordset

##Variables##
$date = get-date
$filelocation = "C:\cidcarnet.htm"
$filelocationcsv = "C:\cidcarnet.csv"

function Get-Tables
{
$Report =  @()
$objConnection.Open($ODBCName)
$objRS = $objConnection.Execute("Select Name from sysobjects where  type='U'")

while ($objRS.EOF -ne $True) {
foreach ($field in $objRS.Fields)  {
$hosts = "" | Select Name
$hosts."Name" = $field.Value
$Report  +=  $hosts

}
$objRS.MoveNext()
}
$objConnection.Close()
$Report  = $Report | sort Name
return $Report
}

function Get-Columns([string]$table)
{
$Report =  @()
$fsTable.Clear()
$fsTable.Columns.Clear()
$objConnection.Open($ODBCName)
$objColumns  = $objConnection.Execute("Select * from $table")

foreach ($column in  $objColumns.Fields) {
$hosts = "" | Select Name, Value
$hosts."Name" =  $column.Name
$hosts."Value" = $column.Value
$Report +=  $hosts
}
$fivalues = @()
$Report |  ForEach-Object{
$fsTable.Columns.Add($_.Name)
$fivalues += $_.Value
}
$fsTable.Rows.add($fivalues)
$dgDataGrid.DataSource =  $fsTable
$objConnection.Close()
}

function ExportToCSV
{
Clear-Content $filelocation
Clear-Content  $filelocationcsv

$selectedRows = @()
for($counter = $dgDataGrid.SelectedCells.Count;  $counter -ge 0; $counter--)
{
$cellIndex =  $dgDataGrid.SelectedCells[$counter].ColumnIndex
$selectedRows +=  $dgDataGrid.Columns[$cellIndex].Name
}
#write-host "1" + $selectedRows

#Tabellenwerte auslesen
$table =  $snTablesNameDrop.SelectedItem
$selectedRows | %{$columns +=  "$_,"}
$columns =  $columns.Substring(0,$columns.lastIndexOfAny(","))
$columns =  $columns.Replace(" ","")
$query = "SELECT " + $columns +" FROM  $table"
write-host  $query
$objConnection.Open($ODBCName)
$Connection.open($query,$objConnection,3,3)

$Report = @()
while (!$Connection.EOF) {
$hosts = "" | Select  $selectedRows

foreach ($column in $selectedRows)
{
$hosts."$column" =  $Connection.Fields.Item("$column").Value
}
$Report +=  $hosts
$Connection.MoveNext()
}

$Connection.Close()
$objConnection.Close()

Write-Host $Report

if($exportRadioBtn2.checked)
{
ConvertTo-Html -title "Check" -body  "&lt;H1&gt;Report&lt;/H1&gt;" -head "&lt;style&gt;body {  background-color:#EEEEEE; }
body,table,td,th { font-family:Tahoma;  color:Black; Font-Size:10pt }
th { font-weight:bold;  background-color:#CCCCCC; }
td { background-color:white;  }&lt;/style&gt;" | Out-File $filelocation
ConvertTo-Html -title "Check"  -body "&lt;H4&gt;&lt;/H4&gt;" | Out-File -Append  $filelocation
ConvertTo-Html -title "Check" -body "&lt;H4&gt;Date and  time&lt;/H4&gt;",$date | Out-File -Append $filelocation
$Report |  ConvertTo-Html -body "&lt;H2&gt;$table&lt;/H2&gt;" | Out-File -Append  $filelocation
}

if($exportRadioBtn1.checked)
{
$Report | Export-Csv -NoTypeInformation  -UseCulture $filelocationcsv
}
}

&nbsp;

function  generateForm
{
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
$form = new-object System.Windows.Forms.form

# Add Tables DropLable
$snTablesNamelableBox = new-object  System.Windows.Forms.Label
$snTablesNamelableBox.Location = new-object  System.Drawing.Size(10,5)
$snTablesNamelableBox.size = new-object  System.Drawing.Size(80,15)
$snTablesNamelableBox.Text =  "Tabellen:"
$form.Controls.Add($snTablesNamelableBox)

# Add Tables Drop Down
$snTablesNameDrop = new-object  System.Windows.Forms.ComboBox
$snTablesNameDrop.Location = new-object  System.Drawing.Size(10,20)
$snTablesNameDrop.Size = new-object  System.Drawing.Size(200,30)
get-tables |  ForEach-Object{$snTablesNameDrop.Items.Add($_.Name)}
$snTablesNameDrop.Add_SelectedValueChanged({Get-Columns($snTablesNameDrop.SelectedItem)})
$form.Controls.Add($snTablesNameDrop)

#Subtabelle
$fsTable = New-Object  System.Data.DataTable
$fsTable.TableName = "TablePreview"
# Add DataGrid View
$dgDataGrid = new-object  System.windows.forms.DataGridView
$dgDataGrid.Location = new-object  System.Drawing.Size(10,50)
$dgDataGrid.size = new-object  System.Drawing.Size(770,75)
$dgDataGrid.AutoSizeRowsMode =  "AllHeaders"
$form.Controls.Add($dgDataGrid)

#Export-Radiobuttons
$exportRadioBtn1 = New-Object  System.Windows.Forms.RadioButton
$exportRadioBtn1.text = "Export nach  csv"
$exportRadioBtn1.top = 145
$exportRadioBtn1.left =  160
$exportRadioBtn1.checked =  $true
$form.Controls.Add($exportRadioBtn1)

$exportRadioBtn2 = New-Object  System.Windows.Forms.RadioButton
$exportRadioBtn2.text = "Export nach  html"
$exportRadioBtn2.top = 170
$exportRadioBtn2.left =  160
$exportRadioBtn2.width = 200
$form.Controls.Add($exportRadioBtn2)

# Create Button and set text and location
$button_exit = New-Object  Windows.Forms.Button
$button_exit.text = "Exit"
$button_exit.Location =  New-Object Drawing.Point 15,245
$form.controls.add($button_exit)

$button_exit.add_click({
$form.Close()
})

# Create Button and set text and location
$button_csv = New-Object  Windows.Forms.Button
$button_csv.text = "Export"
$button_csv.width =  100
$button_csv.Location = New-Object Drawing.Point  15,145
$form.controls.add($button_csv)

$button_csv.add_click({
ExportToCSV
})

$form.Text = "PS-Sybase Control"
$form.size = new-object  System.Drawing.Size(800,300)
$form.autoscroll = $true
$form.Width =  800
$form.Height = 300
$form.FormBorderStyle =  "FixedSingle"
$form.ControlBox =  $false
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()

}

if ($Args.count -lt 1) {

'No Parameter!'
'Usage:'
'./ODBC-Connection.ps1  ODBC-DSNName'
''
}
else{
$ODBCName =  $args[0]
generateForm
}