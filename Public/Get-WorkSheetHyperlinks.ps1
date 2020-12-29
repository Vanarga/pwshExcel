function Get-WorkSheetHyperlinks {
    <#
    .SYNOPSIS
    	This function finds hyperlinks in Excel worksheets.

    .DESCRIPTION
        This function extracts the hyperlinks from an Excel worksheet. It looks at Addresses and SubAddresses and returns an array
        of objects that represent the cells with those hyperlinks. It also accepts an array of strings to filter for specific values.
        If the AllStrings switch is provided, it will search for links that contain all values specified, otherwise it will filter on
        any number of value matches.

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

        Get-WorkSheetHyperlinks -Worksheet <Worksheet Object> [-SearchString <String>] [-AllStrings] [-Silent]

        PS C:\> Get-WorkSheetHyperlinks -Worksheet $worksheet -SearchString '$C$514','.xls','General' -AllStrings -Silent

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-12-20
        Version 1.0 - Get-WorkSheetHyperlinks
    #>
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [String[]]$SearchString,
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
    # If not silent, send worksheet information to standard output.
    if (-not $Silent.IsPresent) {
        Write-Host -Object "Processing worksheet: " -NoNewline
        Write-Host -Object "$($Worksheet.name)" -ForegroundColor Magenta
    }
    # Declare the output array.
    $output = @()
    # Check if the Worksheet contains hyperlinks.
    if ($Worksheet.Hyperlinks.count -gt 0) {
        # If not silent, send worksheet hyperlink count to standard output.
        if (-not $Silent.IsPresent) {
            Write-Host -Object "**** Found Hyperlinks: $($Worksheet.Hyperlinks.count)" -ForegroundColor Yellow
        }
        # Set link counter to 1.
        $i = 1
        # Iterate through all links and filter if relevant.
        ForEach ($link in $Worksheet.Hyperlinks) {
            # Address and SubAddress are mutually exclusive, so adding the strings together results in just the string that exists.
            $Address = $link.Address + $link.SubAddress
            # Process if the AllStrings Parameter is set. Used when the link is a match if it matches all provided SearchStrings.
            if ($AllStrings.IsPresent) {
                # Match Address to each String in the filterString Array. Then check to see if all Boolean result values are true.
                # If all values are true then it matches all search strings.
                if (($filterString | ForEach-Object {$Address -match $_}) -notcontains $false) {
                    #  If not silent, send Hyperlink number and cell information to standard output.
                    if (-not $Silent.IsPresent) {
                        Write-Host -Object "[$i] Cell (Row,Column): ($($link.range.Row),$($link.range.Column)):`r`nFormula: " -NoNewline
                        Write-Host -Object "$Address`r`n" -ForegroundColor Green
                    }
                    # Add the Cell to the output array.
                    $output += $link
                    # Increment the link counter.
                    $i++
                }
            # Process the hyperlink to match any match (one or more). The -join '|' makes the match a regex OR.
            } elseif ($Address -match ($filterString -join '|')) {
                #  If not silent, send Hyperlink number and cell information to standard output.
                if (-not $Silent.IsPresent) {
                    Write-Host -Object "[$i] Cell (Row,Column): ($($link.range.Row),$($link.range.Column)):`r`nFormula: " -NoNewline
                    Write-Host -Object "$Address`r`n" -ForegroundColor Green
                }
                # Add the Cell to the output array.
                $output += $link
                # Increment the link counter.
                $i++
            }
        }
    } else {
        # If the worksheet contains no links, output to standard output, unless silent switch parameter was specified.
        if (-not $Silent.IsPresent) {
            Write-Host -Object "No Hyperlinks found." -ForegroundColor Red
        }
    }
    # Return the output array of cells with hyperlinks that match the search.
    Return $output
}