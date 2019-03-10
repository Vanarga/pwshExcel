<#
.SYNOPSIS
	This script extracts all excel worksheet data and returns a hashtable of custom objects.

.DESCRIPTION
	This script imports Microsoft Excel worksheets and puts the data in to a hashtable of pscustom objects. The hashtable
	keys are the names of the Excel worksheets with spaces omitted. The script imports data from all worksheets. It does not
	validate that the data started in cell A1 and is in format of regular rows and columns, which is required to load the data.

.PARAMETER Path
    The mandatory parameter Path accepts a path string to the excel file. The string can be either the absolute or relative path.
	
.PARAMETER Omit
    The optional parameter Omit accepts a comma separated list of strings of worksheets to omit from loading.

.EXAMPLE
    The example below shows the command line use with Parameters.
    PS C:\> .\Load from excel.ps1 -Path "C:\temp\myExcel.xlsx"
	
	or
	
    PS C:\> .\Load from excel.ps1 -Path "C:\temp\myExcel.xlsx" -Omit "sheet2","sheet3"

.NOTES
	The script requires the PSExcel module downloadable from: https://github.com/Vanarga/PowerShell-Excel

    Author: Michael van Blijdesteijn
    Last Edit: 2019-03-010
    Version 1.0 - Load from excel
#>

[cmdletbinding()]
Param (
    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({Test-Path $_})]
        [String]$Path,
	[Parameter(Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
        [[String[]]$Omit
)

# Check to see if the path is relative or absolute. A rooted path is absolute.
if (-not [System.IO.Path]::IsPathRooted($Path)) {
	# Resolve absolute path from relative path.
	$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

# Import-Module PSExcel.
try {Import-Module PSExcel -erroraction stop}

Catch {Write-host "module PSExcel not present. Please install from https://github.com/Vanarga/PowerShell-Excel"}

# Create Microsoft Excel COM Object.
$obj = Open-Excel

# Load Microsoft Excel Workbook from location Path.
$wb = Get-Workbook -ObjExcel $obj -Path $Path

# Get all Excel worksheet names.
$ws = Get-WorksheetNames -Workbook $wb

# Declare the data array.
$data = @()

# Add each worksheet's pscustom objects to the data array.
$ws | ForEach-Object {
    $data += Get-WorksheetData -worksheet $(Get-Worksheet -Workbook $wb -SheetName $_)
}

# Close Excel.
Close-Excel -ObjExcel $obj

# Declare an ordered hashtable.
$ReturnSet = [ordered]@{}

# Add all the pscustom objects from a worksheet to the hashtable with the key equal to the worksheet name.
# Omit worksheets that were specified in the Omit parameter.
Foreach ($name in ${$ws | Where-Object {$omit -notcontains $_}}) {
	$ReturnSet[$name.replace(" ","")] = $data | Where-Object {$_.WorkSheet -eq $name}
}

# Return the hashtable of custom objects.
Return $ReturnSet
