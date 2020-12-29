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
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [ValidateScript({$_.GetType().FullName -eq "Microsoft.Office.Interop.Excel.ApplicationClass" -or $_.Name -eq "Microsoft Excel"})]
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
            [string]$SheetName
    )
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