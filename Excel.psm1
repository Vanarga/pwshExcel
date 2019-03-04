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
            # Disable the 'visible' property so the Document won't open in excel
            $objExcel.Visible = $Visible
            $objExcel.DisplayAlerts = $DisplayAlerts
            $objExcel.AskToUpdateLinks = $AskToUpdateLinks
		}
		End {
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

    .PARAMETER Workbook
        The mandatory parameter Workbook is the workbook COM Object passed to the function.

    .EXAMPLE
        The example below closes the excel instance defined by the COM Objects from the parameter section.
        PS C:\> Close-Excel -ObjExcel $myObjExcel -Workbook $wb

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
            $workbooks = @()
            $worksheets = @()
        }
        Process {
            ForEach ($workbook in $ObjExcel.Workbooks) {
                $workbooks += $workbook
                $worksheets += $workbook.Sheets.Item($workbook.ActiveSheet.Name)
                $workbook.Close($false)
            }
            $ObjExcel.Quit()
        }
        End {

            Foreach ($w in $worksheets) {
                [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($w)
            }
            Foreach ($w in $workbooks) {
                [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($w)
            }

            [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($ObjExcel)
            [System.GC]::Collect()
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
            [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true)]
                [ValidateScript({Test-Path $_})]
                [String]$Path)
        Begin {
            If (-not [System.IO.Path]::IsPathRooted($Path)) {
                $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            }
        }
        Process {
            $workbook = $ObjExcel.Workbooks.Open($Path)
        }
        End {
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
            $Workbook.Activate()
        }
        Process {
            $worksheet = $Workbook.Sheets.Item($SheetName)
        }
        End {
            Return $worksheet
        }
}

function Add-Worksheet {
    <#
    .SYNOPSIS
        This advanced function creates a new worksheet.

    .DESCRIPTION
        This function creates a new worksheet in the given workbook. If a Sheetname is specified it renames the  
	new worksheet to that name.

    .PARAMETER ObjExcel
        The mandatory parameter ObjExcel is the Excel COM Object passed to the function.

    .PARAMETER Workbook
        The mandatory parameter Workbook is the workbook COM Object passed to the function.

    .PARAMETER Sheetname
        The parameter Sheetname is a string passed to the function to name the newly created worksheet.

    .EXAMPLE
        The example below creates a new worksheet named Data.
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
            $Workbook.Activate()
        }
        Process {
            $worksheet = $ObjExcel.Worksheets.Add($def,$def,1,$def)
            If ($SheetName) {
                $worksheet.Name = $SheetName
            }
        }
        End {
            Return $workbook
        }
}

function Add-Workbook {
    <#
    .SYNOPSIS
        This advanced function creates returns a Microsoft Excel Workbook COM Object.

    .DESCRIPTION
        Given the Microsoft Excel COM Object and Path to the Excel file, the function retuns the Workbook COM Object.

    .PARAMETER Path
        The mandatory parameter Path is the location string of the Excel file.

    .PARAMETER ObjExcel
        The mandatory parameter ObjExcel is needed to retrieve the Workbook COM Object.

    .EXAMPLE
        The example below returns the newly created Excel workbook COM Object.
        PS C:\> Add-Workbook -ObjExcel $myExcelObj -Path "C:\Excel.xlsx"

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
                $workbook = $ObjExcel.Workbooks.Add()
            }
            End {
                Return $workbook
            }
}

function Save-Workbook {
    <#
    .SYNOPSIS
        This advanced function saves the Microsoft Excel Workbook.

    .DESCRIPTION
        This advanced function saves the Microsoft Excel Workbook. If a Path is specified it does a SaveAs, otherwise
	it just saves the data.

    .PARAMETER Path
        The mandatory parameter Path is the location string of the Excel file.

    .PARAMETER Workbook
        The mandatory parameter Workbook is the workbook COM Object passed to the function.

    .EXAMPLE
        The example below Saves the workbook as C:\Excel.xlsx.
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
                Add-Type -AssemblyName Microsoft.Office.Interop.Excel
                $xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook

                If (-not [System.IO.Path]::IsPathRooted($Path)) {
                    $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
                    $Workbook.Activate()
                }
            }
            Process {
                If ($Path) {
                    $workbook.SaveAs($Path,$xlFixedFormat)
                }
                else {
                    $Workbook.Save()
                }
            }
            End {}
}

function Get-WorksheetUsedRange {
    <#
    .SYNOPSIS
        This advanced function returns the Column and Row of the used range in a Worksheet.

    .DESCRIPTION
        This advanced function returns a hashtable containing the last used column and last used row of a worksheet..

    .PARAMETER Worksheet
        The parameter Worksheet is the Excel worksheet com object passed to the function.

    .EXAMPLE
        The example below returns a hashtable containing the last used column and row of the referenced worksheet.
        PS C:\> Get-WorksheetUsedRange $Worksheet

    .NOTES
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
            $What = "*"
            $After = $worksheet.Range("A1")
            $LookIn = [Microsoft.Office.Interop.Excel.XlFindLookIn]::xlValues
            $LookAt = [Microsoft.Office.Interop.Excel.xllookat]::xlPart
            $XlSearchDirection = [Microsoft.Office.Interop.Excel.XlSearchDirection]::xlPrevious
            $MatchCase = $False
            $MatchByte = $False
            $SearchFormat = [Type]::Missing
            $hashtable = [ordered]@{}
        }
        Process {
            $SearchOrder = [Microsoft.Office.Interop.Excel.XlSearchOrder]::xlByColumns
            $hastable["Column"] = $worksheet.Cells.Find($What, $After, $LookIn, $LookAt, $SearchOrder, $XlSearchDirection, $MatchCase, $MatchByte, $SearchFormat).Column
            $SearchOrder = [Microsoft.Office.Interop.Excel.XlSearchOrder]::xlByRows
            $hashtable["Row"] = $worksheet.Cells.Find($What, $After, $LookIn, $LookAt, $SearchOrder, $XlSearchDirection, $MatchCase, $MatchByte, $SearchFormat).Row
        }
        End {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($After)
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
        The switch parameter HashtableReturn with default value False, causes the function to return an array of 
	hashtables instead of an array of objects.

    .EXAMPLE
        The example below returns an array of custom objects using the first row as object parameter names and each 
	additional row as object data.
        PS C:\> Get-WorksheetData $Worksheet

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Get-WorksheetData

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
                [Switch]$HashtableReturn = $false)
        Begin {
            $usedRange = Get-WorksheetUsedRange -worksheet $worksheet

            # Addressing in $worksheet.cells.item(Row,Column)
            $lastColumnAddress		= $workSheet.Cells.Item(1,$usedRange.Column).address()
            $lastColumnRowAddress	= $workSheet.Cells.Item($usedRange.Row,$usedRange.Column).address()
            $header	= $workSheet.Range("A1",$lastColumnAddress).Value()
            $data	= $workSheet.Range("A2",$lastColumnRowAddress).Value()
            $hashtable = [ordered]@{}
            $returnArray = @()
        }
        Process {
            for ($i = 1; $i -lt $lastRow; $i++)
                {
                    for ($j = 1; $j -le $lastColumn; $j++)
                    {
                        If ($lastColumn -ne 1) {
                            $hashtable[$header[1,$j]] = $data[$i,$j]
                        }
                        elseif ($lastRow -gt 2) {
                            $hashtable[$header] = $data[$i,$j]
                        }
                        else {
                            $hashtable[$header] = $data
                        }
                    }
                    # Add Worksheet NoteProperty Item to Hashtable.
                    $hashtable["WorkSheet"] = $workSheet.Name
                    If ($HashtableReturn) {
                        $returnArray += $hashtable
                    }
                    else {
                        $returnArray += [pscustomobject]$hashtable
                    }
                    $hashtable.clear()
                }
        }
        End {
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
        The parameter Worksheet is the Excel worksheet com object passed to the function.

    .PARAMETER ImputArray
        The parameter ImputArray is an Array of custom objects.

    .PARAMETER HashtableReturn
        The switch parameter HashtableReturn with default value False, causes the function to return an array of hashtables 
	instead of an array of objects.

    .EXAMPLE
        The example below returns an array of custom objects using the first row as object parameter names and each additional 
	row as object data.
        PS C:\> Set-WorksheetData $Worksheet

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
            If ($InputArray[0] -is "Hashtable") {
                $InputArray = $InputArray | ForEach-Object {[pscustomobject]$_}
            }
            $myStack = new-object system.collections.stack
            $headers = $InputArray[0].PSObject.Properties.Name
	        $values  = $InputArray | ForEach-Object {$_.psobject.properties.value}
        }
        Process {
	        If ($headers.count -gt 1) {
	        	$values[($values.length - 1)..0] | ForEach-Object {$myStack.Push($_)}
	        	$headers[($headers.length - 1)..0] | ForEach-Object {$myStack.Push($_)}
	        }
	        Else {
	        	$values	 | ForEach-Object {$myStack.Push($_)}
	        	$headers | ForEach-Object {$myStack.Push($_)}
	        }
	        $columns = $headers.Count
	        $rows = $values.Count/$headers.Count + 1
            $array = New-Object 'object[,]' $rows, $columns
	        For ($i=0;$i -lt $rows;$i++)
	        	{
	        		For ($j = 0; $j -lt $columns; $j++) {
	        			$array[$i,$j] = $myStack.Pop()
	        		}
	        	}
	        $range = $Worksheet.Range($Worksheet.Cells(1,1), $Worksheet.Cells($rows,$columns))
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
        The parameter Worksheet is the Excel worksheet com object passed to the function.

    .EXAMPLE
        The example below renames the worksheet to Data unless that name is already in use.
        PS C:\> Set-WorksheetName -Worksheet $ws -SheetName "Data"

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
        The parameter Workbook is the Excel workbook com object passed to the function.

    .EXAMPLE
        The example below renames the worksheet to Data unless that name is already in use.
        PS C:\> Get-WorksheetNames -Workbook $wb

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
            $Workbook.Activate()
        }
        Process {
            $names = ($Workbook.Worksheets | Select-Object Name).Name
        }
        End {
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
        The parameter InputObject is a PSObject.

    .EXAMPLE
        The example below returns a hashtable created from the myPSObject PSObject.
        PS C:\> $myNewHash = ConvertPSObjectToHashtable $myPSObject

    .NOTES
        Author: Dave Wyatt - https://stackoverflow.com/questions/3740128/pscustomobject-to-hashtable
    #>

    param (
        [Parameter(ValueFromPipeline)]
        $InputObject)

    process
    {
        if ($null -eq $InputObject) { return $null }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string])
        {
            $collection = @(
                foreach ($object in $InputObject) { ConvertPSObjectToHashtable $object }
            )

            Write-Output -NoEnumerate $collection
        }
        elseif ($InputObject -is [psobject])
        {
            $hash = @{}

            foreach ($property in $InputObject.PSObject.Properties)
            {
                $hash[$property.Name] = ConvertPSObjectToHashtable $property.Value
            }

            $hash
        }
        else
        {
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
        The parameter InputObject is a hashtable or PSObject.

    .PARAMETER Path
        The mandatory parameter Path is the location string of the Yaml file.

    .EXAMPLE
        The example below returns a hashtable created from the myPSObject PSObject.
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
        If (-not [System.IO.Path]::IsPathRooted($Path)) {
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            $Workbook.Activate()
        }
        if (-not (get-module -Name powershell-yaml)) {
            import-module powershell-yaml}
    }
    process {
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
        The parameter InputObject is a hashtable or PSObject.

    .PARAMETER Path
        The mandatory parameter Path is the location string of the Json file.

    .EXAMPLE
        The example below returns a hashtable created from the myPSObject PSObject.
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
        If (-not [System.IO.Path]::IsPathRooted($Path)) {
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            $Workbook.Activate()
        }
    }
    process {
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
        PS C:\> Export-Json -FilePath "C:\myJsonFile.json"

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
    begin {}
    process {
        $InputObject = Get-Content -Raw -Path $Path | ConvertFrom-Json
    }
    end {
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
        PS C:\> Export-Json -FilePath "C:\myYamlFile.yml"

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
        if (-not (get-module -Name powershell-yaml)) {
            import-module powershell-yaml}
    }
    process {
        $InputObject = [pscustomobject](Get-Content -Raw -Path $Path | ConvertFrom-Yaml | ConvertTo-Json | ConvertFrom-Json)
    }
    end {
        Return $InputObject
    }
}

Export-ModuleMember -Function 'Add-*'
Export-ModuleMember -Function 'Close-*'
Export-ModuleMember -Function 'Export-*'
Export-ModuleMember -Function 'Get-*'
Export-ModuleMember -Function 'Import-*'
Export-ModuleMember -Function 'Open-*'
Export-ModuleMember -Function 'Set-*'
Export-ModuleMember -Function 'Save-*'
