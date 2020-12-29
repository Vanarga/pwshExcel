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
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [ValidateScript({$_.GetType().FullName -eq "Microsoft.Office.Interop.Excel.ApplicationClass" -or $_.Name -eq "Microsoft Excel"})]
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