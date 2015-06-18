#[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
#[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO.SqlDataType')

$serverName = "www.g2gmp.com"
$sheets = "Staff_Roles", "Vendors", "Vendor_Rates"
$provider = "Microsoft.ACE.OLEDB.12.0"
$extendedProperties = "Excel 12.0;HDR=YES;IMEX=1"

$databaseName = "ExcelImport"
$tableName = "ExcelSheet"
$filePath = "D:\Downloads\professional-services-data.xlsx"

$connectionString = "Provider=$provider;Data Source=$filepath;Extended Properties=`"$extendedProperties`";"
# Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Downloads\professional-services-data.xlsx;Extended Properties="Excel 12.0;HDR=YES;IMEX=1"

$conn = New-Object System.Data.OleDb.OleDbConnection($connectionString)
$conn.Open()
#$dataSet = New-Object System.Data.DataSet

$sheets | % {
	$query = "SELECT * FROM [$($_)`$]"
	$cmd = New-Object System.Data.OleDb.OleDbCommand($query, $conn)
	$dataAdapter = New-Object System.Data.OleDb.OleDbDataAdapter($cmd)
	$dataTable = New-Object System.Data.DataTable
	$dataAdapter.fill($dataTable)
	$listName = $_ -replace "_", ""
	$listTitle = $_ -replace "_", " "
	
	#create list if it doesn't exist
	$dataTable.Columns | % {
		# create columns on list
	}
	
	$dataTable.Rows | % {

	}
}




