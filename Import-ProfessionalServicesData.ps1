param(
	[Parameter(Mandatory = $true)]
	[string]$WebUrl,
    [Parameter(Mandatory = $false)]
    [string]$ExcelPath
	)

#region Set up variables
$sheets = "Staff_Roles", "Vendors", "Vendor_Rates"
$provider = "Microsoft.ACE.OLEDB.12.0"
$extendedProperties = "Excel 12.0;HDR=YES;IMEX=1"
#endregion

#region Functions
function PromptFor-File {
	param(
		[string]$Type = "Open",
		[string]$Title = "Select File",
		[string]$FileName = $null,
		[Hashtable]$FileTypes,
		[switch]$RestoreDirectory,
		[IO.DirectoryInfo]$InitialDirectory = $null
	)
	
	[Void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	if($FileTypes) {
		$FileTypes.Keys | % {
			$filter += $FileTypes[$_] + " Files|*.$_|"
		}
		$filter = $filter.TrimEnd("|")
	} else {
		$filter = "All Files|*.*"
	}
	
	switch($Type) {
		"Open" {
			$dialog = New-Object System.Windows.Forms.OpenFileDialog
			$dialog.Multiselect = $false
		}
		"Save" {
			$dialog = New-Object System.Windows.Forms.SaveFileDialog
		}
	}
	
	$dialog.FileName = $FileName
	$dialog.Title = $Title
	$dialog.Filter = $filter
	$dialog.RestoreDirectory = $RestoreDirectory
	$dialog.InitialDirectory = $InitialDirectory.Fullname
	$dialog.ShowHelp = $true
	if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
		return $dialog.FileName
	} else {
		return $null
	}
}

function New-SPList {
	param(
	[Parameter(Mandatory = $true)]
	[Microsoft.SharePoint.SPWeb]$Web,
	[Parameter(Mandatory = $true)]
	[string]$Name,
	[Parameter(Mandatory = $false)]
	[string]$Title,
	[Parameter(Mandatory = $true)]
	[System.Data.DataColumnCollection]$Columns
	)
	
	$listId = $Web.Lists.Add($Name, "Used By Professional Services", [Microsoft.SharePoint.SPListTemplateType]::GenericList)
	
	$list = $Web.Lists.GetList($listId, $true)
	
	if ($Title) {
		$list.Title = $Title
		$list.Update()
	}

	$Columns | % {
		$col = $_
		if ($col.Ordinal -eq 0 -and $Name -ne "VendorRates") {
			if ($col.ColumnName -ne "Title") {
				$titleField = $list.Fields | ? { $_.InternalName -eq "Title" }
				$titleField.Title = $col.ColumnName
				$titleField.Update()
			}
		} else {
            $colName = $col.ColumnName -replace " ", ""
			switch ($Name) {
                "StaffRoles" {
                    $newFieldName = $list.Fields.Add($colName, [Microsoft.SharePoint.SPFieldType]::Note, $false)
				}
                "VendorRates" {
                    if ($colName -match "year") {
                        $newFieldName = $list.Fields.Add($colName, [Microsoft.SharePoint.SPFieldType]::Currency, $true)
                    } else {
 
                        switch ($colName) {
                            "Role" {
                                $luList = $Web.Lists["Staff Roles"]
                            }
                            "Vendor" {
                                $luList = $Web.Lists["Vendors"]
                            }
                        }

                        $newFieldName = $list.Fields.AddLookup($colName, $luList.ID, "true")
                        $newField = $list.Fields[$newFieldName]
                        $newField.LookupField = $luList.Fields["Title"]
                        $newField.Update()
                    }
                }
                default {
                    $newFieldName = $list.Fields.Add($colName, [Microsoft.SharePoint.SPFieldType]::Text, $false)
                }
			}
			$list.Update()
			if ($colName -ne $col.ColumnName) {
				$newField = $list.Fields[$newFieldName];
				$newField.Title = $col.ColumnName
				$newField.Update()
			}
			$list.DefaultView.ViewFields.Add($newFieldName)
			$list.DefaultView.Update()
			$list.Update()
		}
	}

    if ($Name -eq "VendorRates") {
        $titleField = $list.Fields["Title"]
        $titleField.Hidden = $true
        $titleField.Update()
    }

	return $list
}

function Import-SPListContent {
    param(
    [Parameter(Mandatory = $true)]
    [Microsoft.SharePoint.SPList]$List,
    [Parameter(Mandatory = $true)]
    [System.Data.DataTable]$DataTable
    )

    switch ($List.Title) {
        "Vendor Rates" {
            $DataTable.Rows | % {
                $row = $_
                $newItem = $List.Items.Add()
                $DataTable.Columns | % {
                    $colName = $_.ColumnName -replace " ", ""
                    switch ($colName) {
                        "Role" {
                            $luList = $List.ParentWeb.Lists["Staff Roles"]
                            $query = New-Object Microsoft.SharePoint.SPQuery
                            $query.Query = "<Where><Eq><FieldRef Name=`"Title`" /><Value Type=`"Text`">{0}</Value></Eq></Where>" -f $row[$colName]
                            $luItem = $luList.GetItems($query)
                            if ($luItem) {
                                $newItem[$colName] = $luItem
                            }
                        }
                        "Vendor" {
                            $luList = $List.ParentWeb.Lists["Vendors"]
                            $query = New-Object Microsoft.SharePoint.SPQuery
                            $query.Query = "<Where><Eq><FieldRef Name=`"Title`" /><Value Type=`"Text`">{0}</Value></Eq></Where>" -f $row[$colName]
                            $luItem = $luList.GetItems($query)
                            if ($luItem) {
                                $newItem[$colName] = $luItem
                            }
                        }
                        default {
                            $newItem[$colName] = $row[$_.ColumnName]
                        }
                    }


                }
                $newItem.SystemUpdate()
            }
        }
        default {
            $DataTable.Rows | % {
                $row = $_
                $newItem = $List.Items.Add()
                $DataTable.Columns | % {
                    if ($_.Ordinal -eq 0) {
                        $newItem["Title"] = $row[0]
                    }
                    if ($_.Ordinal -ne 0 -and $_.ColumnName -eq "Title") {
                        $colName = "Title0"
                    } else {
                        $colName = $_.ColumnName -replace " ", ""
                    }
                    if ($row[$_.ColumnName] -ne "NULL") {
                        $newItem[$colName] = $row[$_.ColumnName]
                    }
                }
                $newItem.SystemUpdate()
            }
        }
    }
}

#endregion

Add-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

<#
$databaseName = "ExcelImport"
$tableName = "ExcelSheet"
#>

if (-not ($ExcelPath -and (Test-Path -LiteralPath $ExcelPath))) {
    $ExcelPath = PromptFor-File -FileName "professional-services-data.xlsx" -FileTypes @{ "xlsx" = "Open Document Excel" } -InitialDirectory $PWD.Path -Title "Select EXCEL data file"
}

$connectionString = "Provider=$provider;Data Source=$ExcelPath;Extended Properties=`"$extendedProperties`";"
# Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Downloads\professional-services-data.xlsx;Extended Properties="Excel 12.0;HDR=YES;IMEX=1"

$conn = New-Object System.Data.OleDb.OleDbConnection($connectionString)
$conn.Open()

# get the web...
$w = Get-SPWeb -Identity $WebUrl

$sheets | % {
	$query = "SELECT * FROM [$($_)`$]"
	$cmd = New-Object System.Data.OleDb.OleDbCommand($query, $conn)
	$dataAdapter = New-Object System.Data.OleDb.OleDbDataAdapter($cmd)
	$dataTable = New-Object System.Data.DataTable
	$rowCount = $dataAdapter.fill($dataTable)
	$listName = $_ -replace "_", ""
	$listTitle = $_ -replace "_", " "
	
	#create list if it doesn't exist
	$l = New-SPList -Web $w -Name $listName -Title $listTitle -Columns $dataTable.Columns
	
    Import-SPListContent -List $l -DataTable $dataTable
}



$conn.Close()




