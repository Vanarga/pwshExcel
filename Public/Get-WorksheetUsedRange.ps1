function Get-WorksheetUsedRange {
    <#
    .SYNOPSIS
        This advanced function returns the Column and Row of the used range in a Worksheet.

    .DESCRIPTION
        This advanced function returns a hashtable containing the last used column and last used row of a worksheet.

    .PARAMETER Worksheet
        The mandatory parameter Worksheet is the Excel worksheet com object passed to the function.

    .EXAMPLE
        The example below returns a hashtable containing the last used column and row of the referenced worksheet.

        Get-WorksheetUsedRange -Worksheet <PS Excel Worksheet Object>

        PS C:\> Get-WorksheetUsedRange -Worksheet $myWorksheet

    .NOTES
        There are several ways to get the used range in an Excel Worksheet. However, most of them will return areas
        in which formatting has been appied or changed. This method looks for the last column and row where a cell has a value.
        See https://blog.udemy.com/excel-vba-find/ for details.

        Author: Michael van Blijdesteijn
        Last Edit: 2019-02-26
        Version 1.0 - Get-WorksheetData
    #>
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [ValidateScript({$_.GetType().IsCOMObject})]
            $worksheet
    )
    Begin {
        # Define search parameters, see https://blog.udemy.com/excel-vba-find/ for details.
        # What (required): The only required parameter, What tells the Excel what to actually look for. This can be anything – string, integer, etc.).
        $What = "*"
        # After (optional): This specifies the cell after which the search is to begin. This must always be a single cell; you can’t use a range here.
        # If the after parameter isn’t specified, the search begins from the top-left corner of the cell range.
        $After = $worksheet.Range("A1")
        # LookIn (optional): This tells Excel what type of data to look in, such as xlFormulas.
        $LookIn = [Microsoft.Office.Interop.Excel.XlFindLookIn]::xlValues
        # LookAt (optional): This tells Excel whether to look at the whole set of data, or only a selected part. It can take two values: xlWhole and xlPart
        $LookAt = [Microsoft.Office.Interop.Excel.xllookat]::xlPart
        # SearchDirection(optional): This is used to specify whether Excel should search for the next or the previous matching value. You can use either xlNext
        # (to search for next matches) or xlPrevious (to search for previous matches).
        $XlSearchDirection = [Microsoft.Office.Interop.Excel.XlSearchDirection]::xlPrevious
        # MatchCase(optional): Self-explanatory; this tells Excel whether it should match case when doing the search or not. The default value is False.
        $MatchCase = $False
        # MatchByte(optional): This is used if you have installed double-type character set (DBCS). Understanding DBCS is beyond the scope of this tutorial.
        # Like MatchCase, this can also have two values: True or False, with default being False.
        $MatchByte = $False
        # SearchFormat(optional): This parameter is used when you want to select cells with a specified property. It is used in conjunction with the FindFormat
        # property. Say, you have a list of cells where one particular cell (or cell range) is in Italics. You could use the FindFormat property and set it to
        # Italics. If you later use the SearchFormat parameter in Find, it will select the Italicized cell.
        $SearchFormat = [Type]::Missing
        # Define an ordered hashtable.
        $hashtable = [ordered]@{}
    }
    Process {
        # Set the search order to be by columns.
        $SearchOrder = [Microsoft.Office.Interop.Excel.XlSearchOrder]::xlByColumns
        # Return the address of the last used column cell with data in it.
        $hashtable["Column"] = $worksheet.Cells.Find($What, $After, $LookIn, $LookAt, $SearchOrder, $XlSearchDirection, $MatchCase, $MatchByte, $SearchFormat).Column
        # Set the search order to be by rows.
        $SearchOrder = [Microsoft.Office.Interop.Excel.XlSearchOrder]::xlByRows
        # Return the address of the last used row cell with data in it.
        $hashtable["Row"] = $worksheet.Cells.Find($What, $After, $LookIn, $LookAt, $SearchOrder, $XlSearchDirection, $MatchCase, $MatchByte, $SearchFormat).Row
    }
    End {
        # Release the Excel Range COM Object.
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($After)
        # Return the result hashtable.
        Return $hashtable
    }
}
