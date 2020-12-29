function Set-WorksheetData {
    <#
    .SYNOPSIS
        This advanced function populates a Microsoft Excel Worksheet with data from an Array of custom objects or hashtables.

    .DESCRIPTION
        This advanced function populates a Microsoft Excel Worksheet with data from an Array of custom objects. The object
	members populates the first row of the sheet as header items. The object values are placed beneath the headers on
	each successive row.

    .PARAMETER Worksheet
        The mandatory parameter Worksheet is the Excel worksheet com object passed to the function.

    .PARAMETER InputArray
        The mandatory parameter InputArray is an Array of custom objects.

    .EXAMPLE
        The example below returns an array of custom objects using the first row as object parameter names and each additional
    row as object data.

        Set-WorksheetData -Worksheet <PS Excel Worksheet COM Object> -InputArray <PS Object Array>

        PS C:\> Set-WorksheetData -Worksheet $Worksheet -ImputArray $myObjectArray

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Set-WorksheetData
    #>
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [ValidateScript({$_.GetType().IsCOMObject})]
            $worksheet,
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            $InputArray
    )
    Begin {
        # Convert an input hashtables to pscustomobjects
        if ($InputArray[0] -is "Hashtable") {
            $InputArray = $InputArray | ForEach-Object {[pscustomobject]$_}
        }
    }
    Process {
        $properties = $InputArray[0].PSObject.Properties
        # Number of columns is equal to the header count.
        $columns = $properties.Name.Count
        # Number of rows is equal to the number of values devided by the number of headers.
        $rows = $InputArray.Count
        # Create a multidimenstional array sized number of rows by number of columns.
        $array = New-Object -TypeName 'object[,]' $($rows + 1), $columns
        for ($i=0; $i -lt $rows; $i++) {
            $row = $i + 1
            for ($j=0; $j -lt $columns; $j++) {
                if ($i -eq 0) {
                    $array[$i,$j] = $properties.Name[$j];
                }
                $array[$row,$j] = $InputArray[$i].$($properties.Name[$j])
            }
        }
        # Define the Excel worksheet range.
        $range = $Worksheet.Range($Worksheet.Cells(1,1), $Worksheet.Cells($($rows + 1),$columns))
        # Populate the worksheet using the Worksheet.Range function.
        $range.Value2 = $array
    }
    End {}
}