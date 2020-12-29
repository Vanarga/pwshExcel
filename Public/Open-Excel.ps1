function Open-Excel {
    <#
    .SYNOPSIS
        This advanced function opens an instance of the Microsoft Excel application.

    .DESCRIPTION
        The function opens an instance of Microsoft Excel but keeps it hidden unless the Visible parameter is used.

    .PARAMETER Visible
        The parameter switch Visible when specified will make Excel visible on the desktop.

    .EXAMPLE
        The example below returns the Excel COM object when used.

        Open-Excel [-Visible] [-DisplayAlerts] [-AskToUpdateLinks]

        PS C:\> $myObjExcel = Open-Excel

        or

        PS C:\> $myObjExcel = Open-Excel -Visible

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2019-01-19
        Version 1.0 - Open-Excel
    #>
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [Switch]$Visible = $false,
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [Switch]$DisplayAlerts = $false,
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
            [Switch]$AskToUpdateLinks = $false
    )
    Begin {
        # Create an Object Excel.Application using Com interface
        $objExcel = New-Object -ComObject Excel.Application
    }
    Process {
        # Disable the 'visible' property if not specified.
        $objExcel.Visible = $Visible
        # Disable the 'DisplayAlerts' property if not specified.
        $objExcel.DisplayAlerts = $DisplayAlerts
        # Disable the 'AskToUpdateLinks' property if not specified.
        $objExcel.AskToUpdateLinks = $AskToUpdateLinks
    }
    End {
        # Return the Excel COM object.
        Return $objExcel
    }
}