function Get-WorksheetData {
    <#
    .SYNOPSIS
        This advanced function creates an array of pscustom objects from an Microsoft Excel worksheet.

    .DESCRIPTION
        This advanced function creates an array of pscustom objects from an Microsoft Excel worksheet.
        The first row will be used as the object members and each additional row will form the object data for that member.

    .PARAMETER Worksheet
        The parameter Worksheet is the Excel worksheet com object passed to the function.

    .PARAMETER HashtableReturn
        The optional switch parameter HashtableReturn with default value False, causes the function to return an array of
    hashtables instead of an array of objects.

    .PARAMETER TrimHeaders
        The optional switch parameter TrimHeaders, removes whitespace from the column headers when creating the object or hashtable.

    .EXAMPLE
        The example below returns an array of custom objects using the first row as object parameter names and each
    additional row as object data.

        Get-WorksheetData -Worksheet <PS Excel Worksheet COM Object> [-HashtableReturn] [-TrimHeaders]

        PS C:\> Get-WorksheetData -Worksheet $myWorksheet

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-03-19
        Version 1.1 - Get-WorksheetData
    #>
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [ValidateScript({($_.UsedRange.SpecialCells(11).row -ge 2) -and $_.GetType().IsCOMObject})]
            $worksheet,
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [Switch]$HashtableReturn = $false,
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [Switch]$TrimHeaders = $false
    )
    Begin {
        $usedRange = Get-WorksheetUsedRange -worksheet $worksheet
        # Addressing in $worksheet.cells.item(Row,Column)
        # Get the Address of the last column on the worksheet.
        $lastColumnAddress = $workSheet.Cells.Item(1,$usedRange.Column).address()
        # Get the Address of the last row on the worksheet.
        $lastColumnRowAddress = $workSheet.Cells.Item($usedRange.Row,$usedRange.Column).address()
        # Get the values of the first row to use as object Properties. Replace "" with "" to convert to a one dimensional array.
        $headers = $workSheet.Range("A1",$lastColumnAddress).Value() -replace "",""
        # If $TrimHeaders is true, remove whitespce from the headers.
        # https://stackoverflow.com/questions/24355760/removing-spaces-from-a-variable-input-using-powershell-4-0
        # To remove all spaces at the beginning and end of the line, and replace all double-and-more-spaces or tab symbols to spacebar symbol.
        if ($TrimHeaders.IsPresent) {
            $headers = $headers -replace '(^\s+|\s+$)','' -replace '\s+',''
        }
        # Get the values of the remaining rows to use as object values.
        $data	= $workSheet.Range("A2",$lastColumnRowAddress).Value()
        # Define the return array.
        $returnArray = @()
    }
    Process {
        for ($i = 1; $i -lt $UsedRange.Row; $i++)
            {
                # Define an Ordered hashtable.
                $hashtable = [ordered]@{}
                for ($j = 1; $j -le $UsedRange.Column; $j++)
                {
                    # If there is more than one column.
                    if ($UsedRange.Column -ne 1) {
                        # Then add a key value to the current hashtable. Where the key (i.e. header) is in row 1 and column $j and the value (i.e. data) is in row $i and column $j.
                        $hashtable[$headers[$j-1]] = $data[$i,$j]
                    }
                    # If is only one column and there are more than two rows.
                    elseif ($UsedRange.Row -gt 2) {
                        # Then add a key value to the current hashtable. Where the key (i.e. header) is just the header (row 1, column 1) and the value is in row $i and column 1.
                        $hashtable[$headers] = $data[$i,1]
                    }
                    # If there is only there is only one column and two rows.
                    else {
                        # Then add a key value to the current hashtable. Where the key (i.e. header) is just the header (row 1, column 1) and the value is in row 2 and column 1.
                        $hashtable[$headers] = $data
                    }
                }
                # Add Worksheet NoteProperty Item to Hashtable.
                $hashtable["WorkSheet"] = $workSheet.Name
                # If the HashtableReturn switch has been selected, add the hashtable to the return array.
                if ($HashtableReturn) {
                    $returnArray += $hashtable
                }
                else {
                    # If the HashtableReturn switch is $false (Default), convert the hashtable to a custom object and add it to the return array.
                    $returnArray += [pscustomobject]$hashtable
                }
            }
    }
    End {
        # return the array of hashtables or custom objects.
        Return $returnArray
    }
}