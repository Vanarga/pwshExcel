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
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [ValidateScript({$_.GetType().IsCOMObject})]
            $Workbook,
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [String]$Path
    )
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