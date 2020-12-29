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