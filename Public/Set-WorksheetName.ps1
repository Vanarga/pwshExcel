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
    [CmdletBinding ()]
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
            [string]$SheetName
    )
    Begin {}
    Process {
        # Set the current worksheet name to the value of the SheetName string variable.
        $worksheet.Name = $SheetName
    }
    End {}
}