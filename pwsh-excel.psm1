function Open-Excel {
<#
.SYNOPSIS
	This advanced function opens an instance of the Microsoft Excel application.

.DESCRIPTION
    The function opens an instance of Microsoft Excel but keeps it hidden unless the Visible parameter is used.

.PARAMETER Visible
    The parameter switch Visible when specified will make Excel visible on the desktop.

.EXAMPLE
    The example below returns the Excel COM object when used.

    Open-Excel [-Visible] [-DisplayAlerts] [-AskToUpdateLinks]

    PS C:\> $myObjExcel = Open-Excel

    or

    PS C:\> $myObjExcel = Open-Excel -Visible

.NOTES
    Author: Michael van Blijdesteijn
    Last Edit: 2019-01-19
    Version 1.0 - Open-Excel
#>
	[cmdletbinding()]
		Param (
            [Parameter(Mandatory = $false,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [Switch]$Visible = $false,
            [Parameter(Mandatory = $false,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [Switch]$DisplayAlerts = $false,
            [Parameter(Mandatory = $false,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [Switch]$AskToUpdateLinks = $false
        )
		Begin {
            # Create an Object Excel.Application using Com interface
            $objExcel = New-Object -ComObject Excel.Application
        }
		Process {
            # Disable the 'visible' property if not specified.
            $objExcel.Visible = $Visible
            # Disable the 'DisplayAlerts' property if not specified.
            $objExcel.DisplayAlerts = $DisplayAlerts
            # Disable the 'AskToUpdateLinks' property if not specified.
            $objExcel.AskToUpdateLinks = $AskToUpdateLinks
		}
		End {
            # Return the Excel COM object.
			Return $objExcel
		}
}

function Close-Excel {
    <#
    .SYNOPSIS
        This advanced function closes Excel ending all related objects.

    .DESCRIPTION
        The function closes the Excel and releases the COM Object, Workbook, and Worksheet, then cleans up the instance of Excel.

    .PARAMETER ObjExcel
        The mandatory parameter ObjExcel is the Excel COM Object passed to the function.

    .EXAMPLE
        The example below closes the excel instance defined by the COM Objects from the parameter section.

        Close-Excel -ObjExcel <PS Excel COM Object>

        PS C:\> Close-Excel -ObjExcel $myObjExcel

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Close-Excel
    #>
    [cmdletbinding()]
        Param (
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({$_.GetType().FullName -eq "Microsoft.Office.Interop.Excel.ApplicationClass"})]
                $ObjExcel)
        Begin {
            # Define a workbook array.
            $workbooks = @()
            # Define a worksheet array.
            $worksheets = @()
        }
        Process {
            ForEach ($workbook in $ObjExcel.Workbooks) {
                # Add the workbook COM objects to the workbook array.
                $workbooks += $workbook
                # Add the worksheet COM objects to the worksheet array.
                $worksheets += $workbook.Sheets.Item($workbook.ActiveSheet.Name)
                # Close the current worksheet.
                $workbook.Close($false)
            }
            # Quit the Excel Object.
            $ObjExcel.Quit()
        }
        End {
            # Release all the worksheet COM Ojbects.
            Foreach ($w in $worksheets) {
                [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($w)
            }
            # Release all the workbook COM Objects.
            Foreach ($w in $workbooks) {
                [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($w)
            }

            # Release the Excel COM Object.
            [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($ObjExcel)
            # Forces an immediate garbage collection of all generations.
            [System.GC]::Collect()
            # Suspends the current thread until the thread that is processing the queue of finalizers has emptied that queue.
            [System.GC]::WaitForPendingFinalizers()
        }
}

function Get-Workbook {
    <#
    .SYNOPSIS
        This advanced function creates returns a Microsoft Excel Workbook COM Object.

    .DESCRIPTION
        Given the Microsoft Excel COM Object and Path to the Excel file, the function retuns the Workbook COM Object.

    .PARAMETER ObjExcel
        The mandatory parameter ObjExcel is the Excel COM Object passed to the function.

    .PARAMETER Path
        The mandatory parameter Path is the location string of the Excel file. Relative and Absolute paths are supported.

    .EXAMPLE
        The example below returns the workbook COM object specified by Path.

        Get-Workbook -ObjExcel [-Path <String>]

        PS C:\> $wb = Get-Workbook -ObjExcel $myExcelObj -Path "C:\Excel.xlsx"

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Get-Workbook
    #>
    [cmdletbinding()]
        Param (
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({$_.GetType().FullName -eq "Microsoft.Office.Interop.Excel.ApplicationClass"})]
                $ObjExcel,
            [Parameter(Mandatory = $false,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({Test-Path $_})]
                [String]$Path)
        Begin {
            # If no path was specified, prompt for path until it has a value.
            if (-not $Path) {
                $Path = Read-FilePath -Title "Select Microsoft Excel Workbook to Import" -Extension xls,xlsx
                if (-not $Path) {Return "Error, Workbook not specified."}
            }
            # Check to make sure the file is either a xls or xlsx file.
            if ((Get-ChildItem -Path $Path).Extension -notmatch "xls") {
                Return {"File is not an excel file. Please select a valid .xls or .xlsx file."}
            }
            # Check to see if the path is relative or absolute. A rooted path is absolute.
            if (-not [System.IO.Path]::IsPathRooted($Path)) {
                # Resolve absolute path from relative path.
                $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            }
        }
        Process {
            # Open the Excel workbook found at location specified in the Path variable.
            $workbook = $ObjExcel.Workbooks.Open($Path)
        }
        End {
            # Return the workbook COM object.
            Return $workbook
        }
}

function Get-Worksheet {
    <#
    .SYNOPSIS
        This advanced function returns a named Microsoft Excel Worksheet.

    .DESCRIPTION
        This function returns the Worksheet COM Object specified by the Workbook and Sheetname.

    .PARAMETER Workbook
        The mandatory parameter Workbook is the workbook COM Object passed to the function.

    .PARAMETER Sheetname
        The mandatory parameter Sheetname is the name of the worksheet returned.

    .EXAMPLE
        The example below returns the named "Sheet1" worksheet COM Object.

        Get-Worksheet -Workbook <PS Excel Workbook COM Object> -SheetName <String>

        PS C:\> $ws = Get-Worksheet -Workbook $wb -SheetName "Sheet1"

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Get-Worksheet
    #>
    [cmdletbinding()]
        Param (
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({$_.GetType().IsCOMObject})]
                $Workbook,
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({($Workbook.Worksheets | Select-Object Name).Name -Contains $_})]
                [string]$SheetName)
        Begin {
            # Activate the current Excel workbook.
            $Workbook.Activate()
        }
        Process {
            # Get the worksheet COM object specified by the SheetName string variable.
            $worksheet = $Workbook.Sheets.Item($SheetName)
        }
        End {
            # Return the Excel worksheet COM object.
            Return $worksheet
        }
}

function Add-Worksheet {
    <#
    .SYNOPSIS
        This advanced function creates a new worksheet.

    .DESCRIPTION
        This function creates a new worksheet in the given workbook. if a Sheetname is specified it renames the
	new worksheet to that name.

    .PARAMETER ObjExcel
        The mandatory parameter ObjExcel is the Excel COM Object passed to the function.

    .PARAMETER Workbook
        The mandatory parameter Workbook is the workbook COM Object passed to the function.

    .PARAMETER Sheetname
        The optional parameter Sheetname is a string passed to the function to name the newly created worksheet.

    .EXAMPLE
        The example below creates a new worksheet named Data.

        Add-Worksheet -ObjExcel <PS Excel COM Object> -Workbook <PS Excel COM Workbook Object> [-SheetName <String>]

        PS C:\> Add-Worksheet -ObjExcel $myObjExcel -Workbook $wb -Sheetname "Data"

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Add-Worksheet
    #>
    [cmdletbinding()]
        Param (
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({$_.GetType().FullName -eq "Microsoft.Office.Interop.Excel.ApplicationClass"})]
                $ObjExcel,
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({$_.GetType().IsCOMObject})]
                $Workbook,
            [Parameter(Mandatory = $false,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({(Get-WorksheetNames -Workbook $Workbook) -NotContains $_})]
                [string]$SheetName)
        Begin {
            # http://www.planetcobalt.net/sdb/vba2psh.shtml
            $def = [Type]::Missing
            # Activate the current Excel workbook.
            $Workbook.Activate()
        }
        Process {
            # Add a single worksheet to the current workbook.
            $worksheet = $ObjExcel.Worksheets.Add($def,$def,1,$def)
            # If the SheetName variable is specified, rename the new worksheet.
            if ($SheetName) {
                $worksheet.Name = $SheetName
            }
        }
        End {
            # Return the updated Excel workbook COM object.
            Return $workbook
        }
}

function Add-Workbook {
    <#
    .SYNOPSIS
        This advanced function creates returns a Microsoft Excel Workbook COM Object.

    .DESCRIPTION
        Given the Microsoft Excel COM Object and Path to the Excel file, the function retuns the Workbook COM Object.

    .PARAMETER ObjExcel
        The mandatory parameter ObjExcel is needed to retrieve the Workbook COM Object.

    .EXAMPLE
        The example below returns the newly created Excel workbook COM Object.

        Add-Workbook -ObjExcel <PS Excel COM Object>

        PS C:\> Add-Workbook -ObjExcel $myExcelObj

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Add-Workbook
    #>
        [cmdletbinding()]
            Param (
                [Parameter(Mandatory = $true,
                    ValueFromPipeline = $true,
                    ValueFromPipelineByPropertyName = $true)]
                    [ValidateScript({$_.GetType().FullName -eq "Microsoft.Office.Interop.Excel.ApplicationClass"})]
                    $ObjExcel)
            Begin {}
            Process {
                # Add a new workbook to the current Excel COM object.
                $workbook = $ObjExcel.Workbooks.Add()
            }
            End {
                # Return the updated Excel workbook COM object.
                Return $workbook
            }
}

function Save-Workbook {
    <#
    .SYNOPSIS
        This advanced function saves the Microsoft Excel Workbook.

    .DESCRIPTION
        This advanced function saves the Microsoft Excel Workbook. if a Path is specified it does a SaveAs, otherwise
	it just saves the data.

    .PARAMETER Path
        The mandatory parameter Path is the location string of the Excel file.

    .PARAMETER Workbook
        The mandatory parameter Workbook is the workbook COM Object passed to the function.

    .EXAMPLE
        The example below Saves the workbook as C:\Excel.xlsx.

        Save-Workbook -Workbook <PS Excel COM Workbook Object> -Path <String>

        PS C:\> Save-Workbook -Workbook $wb -Path "C:\Excel.xlsx"

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Save-Workbook
    #>
        [cmdletbinding()]
            Param (
                [Parameter(Mandatory = $true,
                    ValueFromPipeline = $true,
                    ValueFromPipelineByPropertyName = $true)]
                    [ValidateScript({$_.GetType().IsCOMObject})]
                    $Workbook,
                [Parameter(Mandatory = $true,
                    ValueFromPipeline = $true,
                    ValueFromPipelineByPropertyName = $true)]
                    [String]$Path)
            Begin {
                # Add Excel namespace
                Add-Type -AssemblyName Microsoft.Office.Interop.Excel
                # Specify file format when saving excel - Open XML Workbook
                $xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook

                # Check to see if the path is relative or absolute. A rooted path is absolute.
                if (-not [System.IO.Path]::IsPathRooted($Path)) {
                    # Resolve absolute path from relative path.
                    $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
                    # Activate the current workbook.
                    $Workbook.Activate()
                }
            }
            Process {
                # If a path was specified proceed with a save as.
                if ($Path) {
                    $workbook.SaveAs($Path,$xlFixedFormat)
                }
                # Check if a path is indicated in the workbook object properties.
                elseif ($Workbook.Path) {
                    # Save the workbook to the path from the workbook object properties.
                    $Workbook.Save()
                }
                else {
                    # Write error to indicate a path must be specified if the workbook was created by this module and has not been previously saved.
                    Write-Error "Workbook has never been saved before, please provide a valid path."
                }
            }
            End {}
}

function Get-WorksheetUsedRange {
    <#
    .SYNOPSIS
        This advanced function returns the Column and Row of the used range in a Worksheet.

    .DESCRIPTION
        This advanced function returns a hashtable containing the last used column and last used row of a worksheet.

    .PARAMETER Worksheet
        The mandatory parameter Worksheet is the Excel worksheet com object passed to the function.

    .EXAMPLE
        The example below returns a hashtable containing the last used column and row of the referenced worksheet.

        Get-WorksheetUsedRange -Worksheet <PS Excel Worksheet Object>

        PS C:\> Get-WorksheetUsedRange -Worksheet $myWorksheet

    .NOTES
        There are several ways to get the used range in an Excel Worksheet. However, most of them will return areas
        in which formatting has been appied or changed. This method looks for the last column and row where a cell has a value.
        See https://blog.udemy.com/excel-vba-find/ for details.

        Author: Michael van Blijdesteijn
        Last Edit: 2019-02-26
        Version 1.0 - Get-WorksheetData
    #>
    [cmdletbinding()]
        Param (
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({$_.GetType().IsCOMObject})]
                $worksheet)
        Begin {
            # Define search parameters, see https://blog.udemy.com/excel-vba-find/ for details.
            # What (required): The only required parameter, What tells the Excel what to actually look for. This can be anything – string, integer, etc.).
            $What = "*"
            # After (optional): This specifies the cell after which the search is to begin. This must always be a single cell; you can’t use a range here.
            # If the after parameter isn’t specified, the search begins from the top-left corner of the cell range.
            $After = $worksheet.Range("A1")
            # LookIn (optional): This tells Excel what type of data to look in, such as xlFormulas.
            $LookIn = [Microsoft.Office.Interop.Excel.XlFindLookIn]::xlValues
            # LookAt (optional): This tells Excel whether to look at the whole set of data, or only a selected part. It can take two values: xlWhole and xlPart
            $LookAt = [Microsoft.Office.Interop.Excel.xllookat]::xlPart
            # SearchDirection(optional): This is used to specify whether Excel should search for the next or the previous matching value. You can use either xlNext
            # (to search for next matches) or xlPrevious (to search for previous matches).
            $XlSearchDirection = [Microsoft.Office.Interop.Excel.XlSearchDirection]::xlPrevious
            # MatchCase(optional): Self-explanatory; this tells Excel whether it should match case when doing the search or not. The default value is False.
            $MatchCase = $False
            # MatchByte(optional): This is used if you have installed double-type character set (DBCS). Understanding DBCS is beyond the scope of this tutorial.
            # Like MatchCase, this can also have two values: True or False, with default being False.
            $MatchByte = $False
            # SearchFormat(optional): This parameter is used when you want to select cells with a specified property. It is used in conjunction with the FindFormat
            # property. Say, you have a list of cells where one particular cell (or cell range) is in Italics. You could use the FindFormat property and set it to
            # Italics. If you later use the SearchFormat parameter in Find, it will select the Italicized cell.
            $SearchFormat = [Type]::Missing
            # Define an ordered hashtable.
            $hashtable = [ordered]@{}
        }
        Process {
            # Set the search order to be by columns.
            $SearchOrder = [Microsoft.Office.Interop.Excel.XlSearchOrder]::xlByColumns
            # Return the address of the last used column cell with data in it.
            $hashtable["Column"] = $worksheet.Cells.Find($What, $After, $LookIn, $LookAt, $SearchOrder, $XlSearchDirection, $MatchCase, $MatchByte, $SearchFormat).Column
            # Set the search order to be by rows.
            $SearchOrder = [Microsoft.Office.Interop.Excel.XlSearchOrder]::xlByRows
            # Return the address of the last used row cell with data in it.
            $hashtable["Row"] = $worksheet.Cells.Find($What, $After, $LookIn, $LookAt, $SearchOrder, $XlSearchDirection, $MatchCase, $MatchByte, $SearchFormat).Row
        }
        End {
            # Release the Excel Range COM Object.
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($After)
            # Return the result hashtable.
            Return $hashtable
        }
}

function Get-WorksheetData {
    <#
    .SYNOPSIS
        This advanced function creates an array of pscustom objects from an Microsoft Excel worksheet.

    .DESCRIPTION
        This advanced function creates an array of pscustom objects from an Microsoft Excel worksheet.
        The first row will be used as the object members and each additional row will form the object data for that member.

    .PARAMETER Worksheet
        The parameter Worksheet is the Excel worksheet com object passed to the function.

    .PARAMETER HashtableReturn
        The optional switch parameter HashtableReturn with default value False, causes the function to return an array of
    hashtables instead of an array of objects.

    .PARAMETER TrimHeaders
        The optional switch parameter TrimHeaders, removes whitespace from the column headers when creating the object or hashtable.

    .EXAMPLE
        The example below returns an array of custom objects using the first row as object parameter names and each
    additional row as object data.

        Get-WorksheetData -Worksheet <PS Excel Worksheet COM Object> [-HashtableReturn] [-TrimHeaders]

        PS C:\> Get-WorksheetData -Worksheet $myWorksheet

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-03-19
        Version 1.1 - Get-WorksheetData
    #>
    [cmdletbinding()]
        Param (
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({($_.UsedRange.SpecialCells(11).row -ge 2) -and $_.GetType().IsCOMObject})]
                $worksheet,
            [Parameter(Mandatory = $false,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [Switch]$HashtableReturn = $false,
            [Parameter(Mandatory = $false,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [Switch]$TrimHeaders = $false
        )
        Begin {
            $usedRange = Get-WorksheetUsedRange -worksheet $worksheet

            # Addressing in $worksheet.cells.item(Row,Column)
            # Get the Address of the last column on the worksheet.
            $lastColumnAddress = $workSheet.Cells.Item(1,$usedRange.Column).address()
            # Get the Address of the last row on the worksheet.
            $lastColumnRowAddress = $workSheet.Cells.Item($usedRange.Row,$usedRange.Column).address()
            # Get the values of the first row to use as object Properties. Replace "" with "" to convert to a one dimensional array.
            $headers = $workSheet.Range("A1",$lastColumnAddress).Value() -replace "",""
            # If $TrimHeaders is true, remove whitespce from the headers.
            # https://stackoverflow.com/questions/24355760/removing-spaces-from-a-variable-input-using-powershell-4-0
            # To remove all spaces at the beginning and end of the line, and replace all double-and-more-spaces or tab symbols to spacebar symbol.
            if ($TrimHeaders.IsPresent) {
                $headers = $headers -replace '(^\s+|\s+$)','' -replace '\s+',''
            }
            # Get the values of the remaining rows to use as object values.
            $data	= $workSheet.Range("A2",$lastColumnRowAddress).Value()
            # Define the return array.
            $returnArray = @()
        }
        Process {
            for ($i = 1; $i -lt $UsedRange.Row; $i++)
                {
                    # Define an Ordered hashtable.
                    $hashtable = [ordered]@{}
                    for ($j = 1; $j -le $UsedRange.Column; $j++)
                    {
                        # If there is more than one column.
                        if ($UsedRange.Column -ne 1) {
                            # Then add a key value to the current hashtable. Where the key (i.e. header) is in row 1 and column $j and the value (i.e. data) is in row $i and column $j.
                            $hashtable[$headers[$j-1]] = $data[$i,$j]
                        }
                        # If is only one column and there are more than two rows.
                        elseif ($UsedRange.Row -gt 2) {
                            # Then add a key value to the current hashtable. Where the key (i.e. header) is just the header (row 1, column 1) and the value is in row $i and column 1.
                            $hashtable[$headers] = $data[$i,1]
                        }
                        # If there is only there is only one column and two rows.
                        else {
                            # Then add a key value to the current hashtable. Where the key (i.e. header) is just the header (row 1, column 1) and the value is in row 2 and column 1.
                            $hashtable[$headers] = $data
                        }
                    }
                    # Add Worksheet NoteProperty Item to Hashtable.
                    $hashtable["WorkSheet"] = $workSheet.Name
                    # If the HashtableReturn switch has been selected, add the hashtable to the return array.
                    if ($HashtableReturn) {
                        $returnArray += $hashtable
                    }
                    else {
                        # If the HashtableReturn switch is $false (Default), convert the hashtable to a custom object and add it to the return array.
                        $returnArray += [pscustomobject]$hashtable
                    }
                }
        }
        End {
            # return the array of hashtables or custom objects.
            Return $returnArray
        }
}

function Set-WorksheetData {
    <#
    .SYNOPSIS
        This advanced function populates a Microsoft Excel Worksheet with data from an Array of custom objects or hashtables.

    .DESCRIPTION
        This advanced function populates a Microsoft Excel Worksheet with data from an Array of custom objects. The object
	members populates the first row of the sheet as header items. The object values are placed beneath the headers on
	each successive row.

    .PARAMETER Worksheet
        The mandatory parameter Worksheet is the Excel worksheet com object passed to the function.

    .PARAMETER InputArray
        The mandatory parameter InputArray is an Array of custom objects.

    .EXAMPLE
        The example below returns an array of custom objects using the first row as object parameter names and each additional
    row as object data.

        Set-WorksheetData -Worksheet <PS Excel Worksheet COM Object> -InputArray <PS Object Array>

        PS C:\> Set-WorksheetData -Worksheet $Worksheet -ImputArray $myObjectArray

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Set-WorksheetData
    #>
    [cmdletbinding()]
        Param (
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({$_.GetType().IsCOMObject})]
                $worksheet,
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                $InputArray)
        Begin {
            # Convert an input hashtables to pscustomobjects
            if ($InputArray[0] -is "Hashtable") {
                $InputArray = $InputArray | ForEach-Object {[pscustomobject]$_}
            }
        }
        Process {
            $properties = $InputArray[0].PSObject.Properties
            # Number of columns is equal to the header count.
            $columns = $properties.Name.Count
            # Number of rows is equal to the number of values devided by the number of headers.
            $rows = $InputArray.Count
            # Create a multidimenstional array sized number of rows by number of columns.
            $array = New-Object 'object[,]' $($rows + 1), $columns

            for ($i=0; $i -lt $rows; $i++) {
                $row = $i + 1
                for ($j=0; $j -lt $columns; $j++) {
                    if ($i -eq 0) {
                        $array[$i,$j] = $properties.Name[$j];
                    }
                    $array[$row,$j] = $InputArray[$i].$($properties.Name[$j])
                }
            }
            # Define the Excel worksheet range.
            $range = $Worksheet.ActiveSheet.Range($Worksheet.ActiveSheet.Cells(1,1), $Worksheet.ActiveSheet.Cells($($rows + 1),$columns))
            # Populate the worksheet using the Worksheet.Range function.
            $range.Value2 = $array
        }
        End {}
}

function Set-WorksheetName {
    <#
    .SYNOPSIS
        This advanced function sets the name of the given worksheet.

    .DESCRIPTION
        This advanced function sets the name of the given worksheet.

    .PARAMETER Worksheet
        The mandatory parameter Worksheet is the Excel worksheet com object passed to the function.

    .EXAMPLE
        The example below renames the worksheet to Data unless that name is already in use.

        Set-WorksheetName -Worksheet <PS Excel Worksheet COM Object> -SheetName <String>

        PS C:\> Set-WorksheetName -Worksheet $myWorksheet -SheetName "Data"

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Set-WorksheetName
    #>
    [cmdletbinding()]
        Param (
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({($_.UsedRange.SpecialCells(11).row -ge 2) -and $_.GetType().IsCOMObject})]
                $worksheet,
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({(Get-WorksheetNames -Workbook $Workbook) -NotContains $_})]
                [string]$SheetName)
        Begin {}
        Process {
            # Set the current worksheet name to the value of the SheetName string variable.
            $worksheet.Name = $SheetName
        }
        End {}
}

function Get-WorksheetNames {
    <#
    .SYNOPSIS
        This advanced function returns a list of all worksheets in a workbook.

    .DESCRIPTION
        This advanced function returns an array of strings of all worksheets in a workbook.

    .PARAMETER Workbook
        The mandatory parameter Workbook is the Excel workbook com object passed to the function.

    .EXAMPLE
        The example below renames the worksheet to Data unless that name is already in use.

        Get-WorksheetNames -Workbook <PS Excel Workbook COM Object>

        PS C:\> Get-WorksheetNames -Workbook $myWorkbook

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-23
        Version 1.0 - Get-WorksheetNames
    #>
    [cmdletbinding()]
        Param (
            [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [ValidateScript({$_.GetType().IsCOMObject})]
            $Workbook)
        Begin {
            # Activate the current workbook.
            $Workbook.Activate()
        }
        Process {
            # Get the names of all worksheets in the current active workbook COM object.
            $names = ($Workbook.Worksheets | Select-Object Name).Name
        }
        End {
            # Return the worksheet names as an array of strings.
            Return $names
        }
}

function ConvertPSObjectToHashtable {
    <#
    .SYNOPSIS
        This advanced function returns a hashtable converted from a PSObject.

    .DESCRIPTION
        This advanced function returns a hashtable converted from a PSObject and will return work with nested PSObjects.

    .PARAMETER InputObject
        The mandatory parameter InputObject is a PSObject.

    .EXAMPLE
        The example below returns a hashtable created from the myPSObject PSObject.

        ConvertPSObjectToHashtable -InputObject <PSObject>

        PS C:\> $myNewHash = ConvertPSObjectToHashtable -InputObject $myPSObject

    .NOTES
        Author: Dave Wyatt - https://stackoverflow.com/questions/3740128/pscustomobject-to-hashtable
    #>

    param (
        [Parameter(ValueFromPipeline)]
        $InputObject)

    process
    {
        # If the inputObject is empty, return $null.
        if ($null -eq $InputObject) { return $null }

        # IF the InputObject can be iterated through and is not a string.
        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string])
        {
            # Call this function recursively for each object in InputObjects.
            $collection = @(
                foreach ($object in $InputObject) { ConvertPSObjectToHashtable $object }
            )

            Write-Output -NoEnumerate $collection
        }
        # If the InputObject is already an Object.
        elseif ($InputObject -is [psobject])
        {
            # Define an hashtable called hash.
            $hash = @{}

            # Iterate through all the properties in the PSObject.
            foreach ($property in $InputObject.PSObject.Properties)
            {
                # Add a key value pair to the hashtable and call the ConvertPSObjectToHashtable function on the property value.
                $hash[$property.Name] = ConvertPSObjectToHashtable $property.Value
            }

            # Return the hashtable.
            $hash
        }
        else
        {
            # Return the InputObject.
            $InputObject
        }
    }
}

function Export-Yaml {
    <#
    .SYNOPSIS
        This advanced function exports a Hashtable or PSObject to a Yaml file.

    .DESCRIPTION
        This advanced function exports a hashtable or PSObject to a Yaml file

    .PARAMETER InputObject
        The mandatory parameter InputObject is a hashtable or PSObject.

    .PARAMETER Path
        The mandatory parameter Path is the location string of the Yaml file.

    .EXAMPLE
        The example below returns a hashtable created from the myPSObject PSObject.

        Export-Yaml -InputObject <PSObject> -Path <String>

        PS C:\> Export-Yaml -InputObject $myHastable -FilePath "C:\myYamlFile.yml"

        or

        PS C:\> Export-Yaml -InputObject $myPSObject -FilePath "C:\myYamlFile.yml"

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-03-01
        Version 1.0 - Export-Yaml
    #>
    param (
		[Parameter(Mandatory=$true, Position=0)]
		    $InputObject,
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [String]$Path)
    begin {
        # Check to see if the path is relative or absolute. A rooted path is absolute.
        if (-not [System.IO.Path]::IsPathRooted($Path)) {
            # Resolve absolute path from relative path.
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            $Workbook.Activate()
        }
        # Install powershell-yaml if not already installed.
        if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Confirm:$false -Force
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
            Install-Module -Name powershell-yaml -AllowClobber -Confirm:$false
        }
        # Import the powershell-yaml module.
        Import-Module powershell-yaml
    }
    process {
        # Convert the InputObject to Yaml and save it to the Path location with overwrite.
        $InputObject | ConvertTo-Yaml | Set-Content -Path $Path -Force
    }
    end {}
}

function Export-Json {
    <#
    .SYNOPSIS
        This advanced function exports a hashtable or PSObject to a Json file.

    .DESCRIPTION
        This advanced function exports a hashtable or PSObject to a Json file

    .PARAMETER InputObject
        The mandatory parameter InputObject is a hashtable or PSObject.

    .PARAMETER Path
        The mandatory parameter Path is the location string of the Json file.

    .EXAMPLE
        The example below returns a hashtable created from the myPSObject PSObject.

        Export-Json -InputObject <PSObject> -Path <String>

        PS C:\> Export-Json -InputObject $myHastable -FilePath "C:\myJsonFile.json"

        or

        PS C:\> Export-Json -InputObject $myPSObject -FilePath "C:\myJsonFile.json"

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-03-01
        Version 1.0 - Export-Json
    #>
    param (
		[Parameter(Mandatory=$true, Position=0)]
		    $InputObject,
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [String]$Path)
    begin {
        # Check to see if the path is relative or absolute. A rooted path is absolute.
        if (-not [System.IO.Path]::IsPathRooted($Path)) {
            # Resolve absolute path from relative path.
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        }
    }
    process {
        # Convert the InputObject to Json and save it to the Path location with overwrite.
        $InputObject | ConvertTo-Json | Set-Content -Path $Path -Force
    }
    end {}
}

function Import-Json {
    <#
    .SYNOPSIS
        This advanced function imports a Json file and returns a PSCustomObject.

    .DESCRIPTION
        This advanced function imports a Json file and returns a PSCustomObject.

    .PARAMETER Path
        The mandatory parameter Path is the location string of the Json file.

    .EXAMPLE
        The example below returns a pscustomobject created from the contents of C:\myJasonFile.json.

        Import-Json -Path <String>

        PS C:\> Import-Json -Path "C:\myJsonFile.json"

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-03-01
        Version 1.0 - Import-Json
    #>
    param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [String]$Path)
    begin {
        if (-not [System.IO.Path]::IsPathRooted($Path)) {
            # Resolve absolute path from relative path.
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        }
    }
    process {
        # Load the raw content from the Path provided file and convert it from Json.
        $InputObject = Get-Content -Raw -Path $Path | ConvertFrom-Json
    }
    end {
        # Return the result set as an array of PSCustom Objects.
        Return $InputObject
    }
}

function Import-Yaml {
    <#
    .SYNOPSIS
        This advanced function imports a Yaml file and returns a PSCustomObject.

    .DESCRIPTION
        This advanced function imports a Yaml file and returns a PSCustomObject.

    .PARAMETER Path
        The mandatory parameter Path is the location string of the Yaml file.

    .EXAMPLE
        The example below returns a pscustomobject created from the contents of C:\myYamlFile.yml.

        Import-Yaml -Path <String>

        PS C:\> Import-Yaml -Path "C:\myYamlFile.yml"

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-03-01
        Version 1.0 - Import-Yaml
    #>
    param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [String]$Path)
    begin {
        if (-not [System.IO.Path]::IsPathRooted($Path)) {
            # Resolve absolute path from relative path.
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        }
        # Install powershell-yaml if not already installed.
        if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Confirm:$false -Force
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
            Install-Module -Name powershell-yaml -AllowClobber -Confirm:$false
        }
        # Import the powershell-yaml module.
        Import-Module powershell-yaml
    }
    process {
        # Load the raw content from the provided path and convert it from Yaml to Json and then from Json to an Array of Custom Objects.
        $InputObject = [pscustomobject](Get-Content -Raw -Path $Path | ConvertFrom-Yaml | ConvertTo-Json | ConvertFrom-Json)
    }
    end {
        # Return the result array of custom objects.
        Return $InputObject
    }
}

function Import-ExcelData {
    <#
    .SYNOPSIS
    	This function extracts all excel worksheet data and returns a hashtable of custom objects.

    .DESCRIPTION
    	This function imports Microsoft Excel worksheets and puts the data in to a hashtable of pscustom objects. The hashtable
    	keys are the names of the Excel worksheets with spaces omitted. The function imports data from all worksheets. It does not
    	validate that the data started in cell A1 and is in format of regular rows and columns, which is required to load the data.

    .PARAMETER Path
        The optional parameter Path accepts a path string to the excel file. The string can be either the absolute or relative path.

    .PARAMETER Exclude
        The optional parameter Exclude accepts a comma separated list of strings of worksheets to exclude from loading.

    .PARAMETER HashtableReturn
        The optional switch parameter HashtableReturn directs if the return array will contain hashtables or pscustom objects.

    .PARAMETER TrimHeaders
        The optional switch parameter TrimHeaders, removes whitespace from the column headers when creating the object or hashtable.

    .EXAMPLE
        The example below shows the command line use with Parameters.

        Import-ExcelData [-Path <String>] [-Exclude <String>,<String>,...] [-HashtableReturn] [-TrimHeaders]

        PS C:\> Import-ExcelData -Path "C:\temp\myExcel.xlsx"

    	or

        PS C:\> Import-ExcelData -Path "C:\temp\myExcel.xlsx" -Exclude "sheet2","sheet3"

    .NOTES

        Author: Michael van Blijdesteijn
        Last Edit: 2019-03-18
        Version 1.0 - Import-ExcelData
    #>

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [ValidateScript({Test-Path $_})]
            [String]$Path,
    	[Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
    		[ValidateNotNullOrEmpty()]
    		[String[]]$Exclude,
    	[Parameter(Mandatory = $false,
    		ValueFromPipeline = $true,
    		ValueFromPipelineByPropertyName = $true)]
            [Switch]$HashtableReturn = $false,
        [Parameter(Mandatory = $false,
    		ValueFromPipeline = $true,
    		ValueFromPipelineByPropertyName = $true)]
    		[Switch]$TrimHeaders = $false
    )

    # If no path was specified, prompt for path until it has a value.
    if (-not $Path) {
        Try {
            $Path = Read-FilePath -Title "Select Microsoft Excel Workbook to Import" -Extension xls,xlsx -ErrorAction Stop
        }
        Catch {
            Return "Path not specified."
        }
    }
    # Check to see if the path is relative or absolute. A rooted path is absolute.
    if (-not [System.IO.Path]::IsPathRooted($Path)) {
    	# Resolve absolute path from relative path.
    	$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    }

    # Check to make sure the file is either a xls or xlsx file.
    if ((Get-ChildItem -Path $Path).Extension -notmatch "xls") {
        Return {"File is not an excel file. Please select a valid .xls or .xlsx file."}
    }

    # Create Microsoft Excel COM Object.
    $obj = Open-Excel

    # Load Microsoft Excel Workbook from location Path.
    $wb = Get-Workbook -ObjExcel $obj -Path $Path

    # Get all Excel worksheet names.
    $ws = Get-WorksheetNames -Workbook $wb

    # Declare the data array.
    $data = @()

    $ws | ForEach-Object {
    	If ($HashtableReturn) {
    		# Add each worksheet's hashtable objects to the data array.
    		$data += Get-WorksheetData -Worksheet $(Get-Worksheet -Workbook $wb -SheetName $_) -HashtableReturn:$true -TrimHeaders:$TrimHeaders.IsPresent
    	}
    	else {
    		# Add each worksheet's pscustom objects to the data array.
    		$data += Get-WorksheetData -Worksheet $(Get-Worksheet -Workbook $wb -SheetName $_) -TrimHeaders:$TrimHeaders.IsPresent
    	}
    }

    # Close Excel.
    Close-Excel -ObjExcel $obj

    # Declare an ordered hashtable.
    $ReturnSet = [Ordered]@{}

    # Add all the pscustom objects from a worksheet to the hashtable with the key equal to the worksheet name.
    # Exclude worksheets that were specified in the Exclude parameter.
    ForEach ($name in $($ws | Where-Object {$Exclude -NotContains $_})) {
    	$ReturnSet[$name.replace(" ","")] = $data | Where-Object {$_.WorkSheet -eq $name}
    }

    # Return the hashtable of custom objects.
    Return $ReturnSet

}

function Read-FilePath {
    <#
    .SYNOPSIS
    	This function opens a gui window dialog to navigate to an excel file.

    .DESCRIPTION
    	This function opens a gui window dialog to navigate to an excel file and returns the path.

    .PARAMETER Title
        The mandatory parameter Title, is a string that appears on the navigation window.

    .PARAMETER Extension
        The optional parameter Extension, is a string array that filters the file extensions to allow selection of.

    .EXAMPLE
        The example below shows the command line use with Parameters.

        Read-FilePath -Title <String> -Extension <String[]>

        PS C:\> Read-FilePath -Title "Select a file to upload" -Extension exe,msi,intunewin

    .NOTES

        Author: Michael van Blijdesteijn
        Last Edit: 2019-08-06
        Version 1.0 - Read-FilePath
    #>

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [String]$Title,
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [String[]]$Extension
    )
    # https://docs.microsoft.com/en-us/previous-versions/windows/silverlight/dotnet-windows-silverlight/cc189944(v%3dvs.95)

    Add-Type -AssemblyName System.Windows.Forms
    $topform = New-Object System.Windows.Forms.Form
	$topform.Topmost = $true
    $topform.MinimizeBox = $true

    $openFileDialog = New-Object windows.forms.openfiledialog
    $openFileDialog.title = $Title
    $openFileDialog.InitialDirectory = $pwd.path
    if ($Extension) {
        $openFileDialog.filter = "File types ($(($Extension -join "; *.").Insert(0,"*.")))|$(($Extension -join ";*.").Insert(0,"*."))"
    }
    $openFileDialog.ShowHelp = $false
    $openFileDialog.ShowDialog($topform) | Out-Null

    if ($openFileDialog.FileName -eq "") {
        Return $null
    }
    else {
        Return $openFileDialog.FileName
    }
}

# Export the functions above.
Export-ModuleMember -Function 'Add-*'
Export-ModuleMember -Function 'Close-*'
Export-ModuleMember -Function 'Export-*'
Export-ModuleMember -Function 'Get-*'
Export-ModuleMember -Function 'Import-*'
Export-ModuleMember -Function 'Open-*'
Export-ModuleMember -Function 'Read-*'
Export-ModuleMember -Function 'Set-*'
Export-ModuleMember -Function 'Save-*'
