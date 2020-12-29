function Close-Excel {
    <#
    .SYNOPSIS
        This advanced function closes Excel ending all related objects.

    .DESCRIPTION
        The function closes the Excel and releases the COM Object, Workbook, and Worksheet, then cleans up the instance of Excel.

    .PARAMETER ObjExcel
        The mandatory parameter ObjExcel is the Excel COM Object passed to the function.

    .EXAMPLE
        The example below closes the excel instance defined by the COM Objects from the parameter section.

        Close-Excel -ObjExcel <PS Excel COM Object>

        PS C:\> Close-Excel -ObjExcel $myObjExcel

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Close-Excel
    #>
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [ValidateScript({$_.GetType().FullName -eq "Microsoft.Office.Interop.Excel.ApplicationClass" -or $_.Name -eq "Microsoft Excel"})]
            $ObjExcel
    )
    Begin {
        # Define a workbook array.
        $workbooks = @()
        # Define a worksheet array.
        $worksheets = @()
    }
    Process {
        ForEach ($workbook in $ObjExcel.Workbooks) {
            # Add the workbook COM objects to the workbook array.
            $workbooks += $workbook
            # Add the worksheet COM objects to the worksheet array.
            $worksheets += $workbook.Sheets.Item($workbook.ActiveSheet.Name)
            # Close the current worksheet.
            $workbook.Close($false)
        }
        # Quit the Excel Object.
        $ObjExcel.Quit()
    }
    End {
        # Release all the worksheet COM Ojbects.
        Foreach ($w in $worksheets) {
            [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($w)
        }
        # Release all the workbook COM Objects.
        Foreach ($w in $workbooks) {
            [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($w)
        }
        # Release the Excel COM Object.
        [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($ObjExcel)
        # Forces an immediate garbage collection of all generations.
        [System.GC]::Collect()
        # Suspends the current thread until the thread that is processing the queue of finalizers has emptied that queue.
        [System.GC]::WaitForPendingFinalizers()
    }
}