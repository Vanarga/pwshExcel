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

    [CmdletBinding ()]
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