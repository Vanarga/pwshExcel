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
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [ValidateScript({$_.GetType().FullName -eq "Microsoft.Office.Interop.Excel.ApplicationClass" -or $_.Name -eq "Microsoft Excel"})]
            $ObjExcel
    )
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