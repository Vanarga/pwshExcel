Function Get-ExcelWorkBookLinks {
    <#
    .SYNOPSIS
    	This function finds hyperlinks and formulas with external references.

    .DESCRIPTION
        This function returns a hastable with two arrays of custom objects. One array contains all the Excel (xls or xlsx) files with
        Hyperlinks and Formulas with external references. The other array contains the Excel files that could not be opened due to being
        secured with a password.

    .PARAMETER Path
        The mandatory Path parameter is the string of the folder path the function will recursively search for Excel files to process.
        The path can be either relative or absolute.

    .PARAMETER Log
        The optional Log switch parameter enables detailed logging to standard output.

    .PARAMETER Exportpath
        The optional Path string parameter is the string of the folder path the function will write the current working data in two csv
        files. The file names will be passwordprotected.csv and links.csv.

    .PARAMETER Skip
        The optional Skip integer parameter is accepts the number of files to skip while searching for links in a specific folder heirarchy.
        This enables the function to resume where it left off.

    .EXAMPLE
        The example below shows the command line use with Parameters.

        Get-ExcelWorkBookLinks -Path <Folder Path String or Path Object> [-Log:$true] [-ExportPath <Folder path String>] [-Skip <Integer>]

        PS C:\> Get-ExcelWorkBookLinks -Path C:\Temp

        or

        PS C:\> Get-ExcelWorkBookLinks -Path C:\Temp -Log:$true -ExportPath C:\Temp -Skip 26

    .NOTES
        Author: Michael van Blijdesteijn
        Last Edit: 2020-1-2
        Version 1.0.1 - Get-ExcelWorkBookLinks
    #>
    [CmdletBinding ()]
    Param (
        [Parameter(Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        $Path,
        [Parameter(Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [Switch]$Log,
        [Parameter(Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        $ExportPath,
        [Parameter(Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        $Skip
    )
    # Check to see if the path is relative or absolute. A rooted path is absolute.
    if (-not [System.IO.Path]::IsPathRooted($Path)) {
        # Resolve absolute path from relative path.
        $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    }
    # Declare the Excel COM object.
    $obj = Open-Excel
    # Declare the array for files that could be processed.
    $Enumerated = @()
    # Declare the array for files that were password protected.
    $protected = @()
    # Get the Excel files in the folder path.
    $files = Get-ChildItem -Path $Path -Include "*.xls","*.xlsx" -Recurse
    # Write the number of files found to the standard output.
    Write-Host -Object "$($files.count) Excel workbooks found." -ForegroundColor Magenta

    # Lets you skip a certain number of files. (For safety, always start)
    if ($Skip) {
        $files = $files | Select-Object -Skip $Skip
    }

    # Creates a CSV file to list password protected files and places it on the desktop of current user and adds a header row
    New-Item -Path $ExportPath -ItemType File -Name passwordprotected.CSV
    Add-content -Path "$ExportPath\links.csv" -Value "FileName,DirectoryName,Path,HyperLinks,ExternalFormulaLinks"

    # Creates a CSV file to list all excel files for links and formulas and adds a header row
    New-Item -Path $ExportPath -ItemType File -Name links.CSV
    Add-content -Path "$ExportPath\passwordprotected.csv" -Value "FileName,DirectoryName,Path,PasswordProtected"

    # Set the file counter.
    if ($Skip) {
        $i = $Skip
    } else {
        $i = 1
    }

    # Iterate through all the files.
    ForEach ($file in $files) {
        # Set the hyperlink counter to zero.
        $numhyperlinks = 0
        # Set the external formula reference to zero.
        $numfomulalinks = 0
        # Write what the file being worked on to the standard output.
        Write-Host -Object "[$i]: Processing file: " -NoNewline -ForegroundColor White
        Write-Host -Object $file -ForegroundColor Green
        # https://stackoverflow.com/questions/36555094/powershell-test-for-excel-password-protection
        $sig = [Byte[]] (0x50,0x4b,0x03,0x04)
        $bytes = Get-Content $file.fullname -Encoding Byte -Total 4
        if (@(Compare-Object $sig $bytes -sync 0).length -eq 0) {
          # process unencrypted file
          $workbook = Get-Workbook -ObjExcel $obj -Path $file.FullName
        } else {
            Write-Host -Object "$file is protected by password, skipping..." -ForegroundColor Yellow
            $PasswordLocked = [PSCustomObject]@{
                FileName = $file.Name
                DirectoryName = $file.DirectoryName
                Path = $file.FullName
                PasswordProtected = $true
            }
            # Add the custom object to the protected array.
            $protected += $PasswordLocked

            # Write the output to the passwordprotected file.
            Add-Content -Path "$ExportPath\passwordprotected.csv" -Value "$($file.Name),$($file.DirectoryName),$($file.FullName),True"

            # Increment the number of files iterated through.
            $i++
            # Continue to the next item in the ForEach loop. Skip everything below this.
            continue
        }

        # Get the list of Worksheet names from the workbook.
        $worksheetnames = Get-WorksheetNames -Workbook $workbook

        # Iterate through the worksheet names and get the worksheets.
        ForEach ($sheet in $worksheetnames) {
            $worksheet = Get-Worksheet -Workbook $workbook -SheetName $sheet
            # Get the number of hyperlinks in the worksheet.
            $numhyperlinks += (Get-WorkSheetHyperlinks -Worksheet $worksheet -Silent:$(-not $Log.IsPresent)).Count
            # Get the number of formulas with external references in the worksheet.
            $numfomulalinks += (Get-ExcelFormulalinks -Worksheet $worksheet -Silent:$(-not $Log.IsPresent)).Count
        }
        # Create a pscustom object for the workbook with details of the number of hyperlinks and formulas with external references.
        $links = [PSCustomObject]@{
            FileName = $file.Name
            DirectoryName = $file.DirectoryName
            Path = $file.FullName
            HyperLinks = $numhyperlinks
            ExternalFormulaLinks = $numfomulalinks
        }
        # Add the custom object to the Enumerated array.
        $Enumerated += $links
        # Close the Excel workbook.
        $workbook.Close($false)

        # Write the output to the links file.
        Add-Content -Path "$ExportPath\links.csv" -Value "$($file.FileName), $($file.DirectoryName), $($file.FullName), $numhyperlinks, $numformulalinks"

        # Increment the number of files iterated through.
        $i++
    }
    # Close Excel.
    Close-Excel $obj
    # Create the output hastable.
    $output = @{
        output = $Enumerated
        protected = $protected
    }
    # Return the output hastable.
    Return $output
}