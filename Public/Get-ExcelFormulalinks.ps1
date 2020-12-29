function Get-ExcelFormulalinks {
    <#
    .SYNOPSIS
        This function finds formulas with external file references in Excel worksheets.

    .DESCRIPTION
        This function extracts the formulas with external file references from an Excel worksheet. It looks for formulas that contain
        ".xls" or ".xlsx" and returns them in an array of objects that represent the cells with those formulas. It also accepts an array
        of strings to filter for specific values if other search criteria are needed. If the AllStrings switch is provided, it will search
        for references that contain all values specified, otherwise it will filter on any number of value matches.

    .PARAMETER SearchString
        The optional parameter SearchString accepts a comma-separated list of strings to filter for. Values should be entered with
        single quotes.

    .PARAMETER Worksheet
        The mandatory parameter Worksheet accepts the Worksheet Object to find links in.

    .PARAMETER AllStrings
        The optional parameter AllStrings, changes the search logic from OR (match one or more terms) to AND (Match all terms).

    .PARAMETER Silent
        The optional parameter Silent, suppresses all messages the function returns.

    .EXAMPLE
        The example below shows the command line use with Parameters.

        Get-ExcelFormulalinks -Worksheet <Worksheet Object> [-SearchString <String>] [-AllStrings] [-Silent]

        PS C:\> Get-ExcelFormulalinks -Worksheet $worksheet -SearchString '$C$514','.xls','General' -AllStrings -Silent

    .NOTES
        https://stackoverflow.com/questions/9662669/powershell-how-to-retrieve-formula-without-iterating-through-all-row-and-column
        https://docs.microsoft.com/en-us/office/vba/api/excel.xlcelltype
        http://excelmatters.com/2015/06/02/super-secret-specialcells/
        Author: Michael van Blijdesteijn
        Last Edit: 2019-12-20
        Version 1.0 - Get-ExcelFormulalinks
    #>
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [String[]]$SearchString = @(".xlsx",".xls"),
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            $Worksheet,
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [Switch]$AllStrings,
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [Switch]$Silent
    )
    # Replace all non-alphanumeric characters with a '\' followed by that character e.g. $ becomes \$
    $filterString = $SearchString -replace "[^(a-zA-Z0-9)]","\$&"
    # Declare the output array.
    $output = @()
    # Try to get the cells with formulas in them.
    Try {
        # (XlCellType constant, XlSpecialCellsValue constant) - Return the Range of cell with formula
        $cellsWithFormula = $Worksheet.Cells.SpecialCells(-4123,23) | Where-Object {$_.Formula}
        # If not silent, send worksheet Formula count to standard output.
        if (-not $Silent.IsPresent) {
            Write-Host "**** Found formulas: $($cellsWithFormula.count)" -ForegroundColor Yellow
        }
        # If there are formulas do, find the ones with external references.
        if ($cellsWithFormula) {
            # Set index to 1. This is to keep track of which formula is being worked on.
            $i = 1
            # Using ForEach instead of For. This helps to enumerate the formulas returned by the COM object. The array cannot be indexed with For - Object[$i]
            ForEach ($cell in $cellsWithFormula) {
                # Process if the AllStrings Parameter is set. Used when the link is a match if it matches all provided SearchStrings.
                if ($AllStrings.IsPresent) {
                    # Match formula to each String in the filterString Array. Then check to see if all Boolean result values are true.
                    # If all values are true then it matches all search strings.
                	if (($filterString | ForEach-Object {$cell.formula -match $_}) -notcontains $false) {
                        #  If not silent, send formula number and cell information to standard outp
                        if (-not $Silent.IsPresent) {
                            Write-Host "[$i] Cell (Row,Column): ($($cell.Row),$($cell.Column)):`r`nFormula: " -NoNewline
                            Write-Host "$($cell.formula)`r`n" -ForegroundColor Green
                        }
                        # Add the Cell to the output array.
                        $output += $cell
                        # Increment the formula reference counter.
                        $i++
                    }
                # Process the formula to match any match (one or more). The -join '|' makes the match a regex OR.
                } elseif ($cell.formula -match ($filterString -join '|')) {
                    #  If not silent, send formula number and cell information to standard output.
                    if (-not $Silent.IsPresent) {
                        Write-Host "[$i] Cell (Row,Column): ($($cell.Row),$($cell.Column)):`r`nFormula: " -NoNewline
                        Write-Host "$($cell.formula)`r`n" -ForegroundColor Green
                    }
                    # Add the Cell to the output array.
                    $output += $cell
                    # Increment the link counter.
                    $i++
                }
            }
        }
    }
    # If there are no Formulas the write to standard output.
    Catch {
        if (-not $Silent.IsPresent) {
            Write-Host "No formulas found.`r`n" -ForegroundColor Red
        }
    }
    if (-not $Silent.IsPresent) {
        Write-Host "**** Found formulas with external links: $($output.count)" -ForegroundColor Yellow
    }
    # Return the output array of cells with formulas that contain external references or that match the search.
    Return $output
}