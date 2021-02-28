[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
#
#Create Form
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Create package directory tree"
$objForm.Size = New-Object System.Drawing.Size(500,300) 
$objForm.StartPosition = "CenterScreen"
#
#Create Grid
$objGrid = New-Object System.Windows.Forms.DataGridView
$objGrid.Size = New-Object System.Drawing.Size(500,300)
$objGrid.ColumnCount = 2
$objGrid.RowCount = 7
$objGrid.RowHeadersVisible = $false
$objGrid.ColumnHeadersVisible = $true
$objGrid.AllowUserToResizeColumns = $false
$objGrid.AllowUserToResizeRows = $false
$objGrid.AllowUserToAddRows = $false
$objGrid.AllowUserToDeleteRows = $false
$objGrid
$objGrid.Columns[0].SortMode = "NotSortable"
$objGrid.Columns[1].SortMode = "NotSortable"
$objGrid.Columns[0].Name = "Property"
$objGrid.Columns[1].Name = "Value"
$objGrid.Columns[0].Width = 250
$objGrid.Columns[1].Width = 250
$objForm.Controls.Add($objGrid)




$objForm.Topmost = $True
$objForm.Topmost = $True
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()