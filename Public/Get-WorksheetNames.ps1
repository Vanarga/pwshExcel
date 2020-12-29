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
    [CmdletBinding ()]
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