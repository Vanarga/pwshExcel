# Readme #
## Project: PowerShell Excel Module ##

This repo contains the Excel PowerShell module for manipulating Microsoft Excel Workbooks, Worksheets, etc. The main purpose of the module is to allow data retrieval from Excel for use in automation scripts and configuration scripts.

## Installation ##  
Create a folder called PSExcel in **C:\Windows\System32\WindowsPowerShell\v1.0\Modules\**  
Copy files **PSExcel.psd1** and **PSExcel.psm1** to C:\Windows\System32\WindowsPowerShell\v1.0\Modules\PSExcel  

## Functions ##
1. Open-Excel
2. Close-Excel
3. Get-Workbook
4. Get-Worksheet
5. Add-Worksheet
6. Add-Workbook
7. Save-Workbook
8. Get-WorksheetUsedRange
9. Get-WorksheetData
10. Set-WorksheetData
11. Set-WorksheetName
12. Get-WorksheetNames
13. ConvertPSObjectToHashtable
14. Export-Yaml
15. Export-Json
16. Import-Jason
17. Import-Yaml
18. Import-ExcelData
19. Read-ExcelPath

## Function Help ##
1. **Open-Excel** - Creates a new excel COM object.  

**.DESCRIPTION**  
    The function opens an instance of Microsoft Excel but keeps it hidden unless the Visible parameter is used.  

**.PARAMETER** - Visible  
    The parameter switch Visible when specified will make Excel visible on the desktop.  

**.PARAMETER** - DisplayAlerts  
    The parameter switch DisplayAlerts when specified will make Excel Display Alerts if any are triggered.  

**.PARAMETER** - AskToUpdateLinks  
    The parameter switch AskToUpdateLinks when specified will make Excel prompt to Update Links.  

**.EXAMPLE**  
    The example below returns the Excel COM object when used.  
```
    The example below returns the Excel COM object when used.

    Open-Excel [-Visible] [-DisplayAlerts] [-AskToUpdateLinks]

    PS C:\> $myObjExcel = Open-Excel

    or

    PS C:\> $myObjExcel = Open-Excel -Visible
```
2. **Close-Excel** - Close Excel and release COM objects.

    **.DESCRIPTION**  
        The function closes the Excel and releases the COM Object, Workbook, and Worksheet, then cleans up the instance of Excel.  

    **.PARAMETER** - ObjExcel  
        The mandatory parameter ObjExcel is the Excel COM Object passed to the function.  

    **.PARAMETER** - Workbook  
        The mandatory parameter Workbook is the workbook COM Object passed to the function.  

    **.EXAMPLE**  
        The example below closes the excel instance defined by the COM Objects from the parameter section.  
```
        Close-Excel -ObjExcel <PS Excel COM Object>

        PS C:\> Close-Excel -ObjExcel $myObjExcel
```
3. **Get-Workbook**

    **.DESCRIPTION**  
        Given the Microsoft Excel COM Object and Path to the Excel file, the function retuns the Workbook COM Object.  

    **.PARAMETER** - ObjExcel  
        The mandatory parameter ObjExcel is the Excel COM Object passed to the function.  

    **.PARAMETER** - Path  
        The optional parameter Path is the location string of the Excel file. Relative and Absolute paths are supported.  

    **.EXAMPLE**  
        The example below returns the workbook COM object specified by Path.  
```
        Get-Workbook -ObjExcel [-Path <String>]

        PS C:\> $wb = Get-Workbook -ObjExcel $myExcelObj -Path "C:\Excel.xlsx"
```
4. **Get-Worksheet**

    **.DESCRIPTION**  
        This function returns the Worksheet COM Object specified by the Workbook and Sheetname.  

    **.PARAMETER** - Workbook  
        The mandatory parameter Workbook is the workbook COM Object passed to the function.  

    **.PARAMETER** - Sheetname  
        The mandatory parameter Sheetname is the name of the worksheet returned.  

    **.EXAMPLE**  
        The example below returns the named "Sheet1" worksheet COM Object.  
```
        Get-Worksheet -Workbook <PS Excel Workbook COM Object> -SheetName <String>

        PS C:\> $ws = Get-Worksheet -Workbook $wb -SheetName "Sheet1"
```
5. **Add-Worksheet**

    **.DESCRIPTION**  
        This function creates a new worksheet in the given workbook. If a Sheetname is specified it renames the new worksheet to that name.  

    **.PARAMETER** - ObjExcel  
        The mandatory parameter ObjExcel is the Excel COM Object passed to the function.  

    **.PARAMETER** - Workbook  
        The mandatory parameter Workbook is the workbook COM Object passed to the function.  

    **.PARAMETER** - Sheetname  
        The optional parameter Sheetname is a string passed to the function to name the newly created worksheet.  

    **.EXAMPLE**  
        The example below creates a new worksheet named Data.  
```
        Add-Worksheet -ObjExcel <PS Excel COM Object> -Workbook <PS Excel COM Workbook Object> [-SheetName <String>]

        PS C:\> Add-Worksheet -ObjExcel $myObjExcel -Workbook $wb -Sheetname "Data"
```
6. **Add-Workbook**

    **.DESCRIPTION**  
        Given the Microsoft Excel COM Object and Path to the Excel file, the function retuns the Workbook COM Object.  

    **.PARAMETER** - ObjExcel  
        The mandatory parameter ObjExcel is needed to retrieve the Workbook COM Object.  

    **.EXAMPLE**  
        The example below returns the newly created Excel workbook COM Object.  
```
        Add-Workbook -ObjExcel <PS Excel COM Object>

        PS C:\> Add-Workbook -ObjExcel $myExcelObj
```
7. **Save-Workbook**

    **.DESCRIPTION**  
        This advanced function saves the Microsoft Excel Workbook. If a Path is specified it does a SaveAs, otherwise it just saves the data.  

    **.PARAMETER** - Path  
        The mandatory parameter Path is the location string of the Excel file.  

    **.PARAMETER** - Workbook  
        The mandatory parameter Workbook is the workbook COM Object passed to the function.  

    **.EXAMPLE**  
        The example below Saves the workbook as C:\Excel.xlsx.  
```
        Save-Workbook -Workbook <PS Excel COM Workbook Object> -Path <String>

        PS C:\> Save-Workbook -Workbook $wb -Path "C:\Excel.xlsx"
```
8. **Get-WorksheetUsedRange**

    **.DESCRIPTION**  
        This advanced function returns a hashtable containing the last used column and last used row of a worksheet..  

    **.PARAMETER** - Worksheet  
        The mandatory parameter Worksheet is the Excel worksheet com object passed to the function.  

    **.EXAMPLE**  
        The example below returns a hashtable containing the last used column and row of the referenced worksheet.  
```
        Get-WorksheetUsedRange -Worksheet <PS Excel Worksheet Object>

        PS C:\> Get-WorksheetUsedRange -Worksheet $myWorksheet
```
9. **Get-WorksheetData**

    **.DESCRIPTION**  
        This advanced function creates an array of pscustom objects from an Microsoft Excel worksheet.  
        The first row will be used as the object members and each additional row will form the object data for that member.  

    **.PARAMETER** - Worksheet  
        The mandatory parameter Worksheet is the Excel worksheet com object passed to the function.  

    **.PARAMETER** - HashtableReturn  
        The optional switch parameter HashtableReturn with default value False, causes the function to return an array of hashtables instead of an array of objects.  

    **.PARAMETER** - TrimHeaders  
        The optional switch parameter TrimHeaders, removes whitespace from the column headers when creating the object or hashtable.  

    **.EXAMPLE**  
        The example below returns an array of custom objects using the first row as object parameter names and each additional row as object data.  
```
        Get-WorksheetData -Worksheet <PS Excel Worksheet COM Object> [-HashtableReturn] [-TrimHeaders]

        PS C:\> Get-WorksheetData -Worksheet $myWorksheet
```
10. **Set-WorksheetData**

    **.DESCRIPTION**  
        This advanced function populates a Microsoft Excel Worksheet with data from an Array of custom objects. The object members populates the first row of the sheet as header items.  
        The object values are placed beneath the headers on each successive row.  

    **.PARAMETER** - Worksheet  
        The mandatory parameter Worksheet is the Excel worksheet com object passed to the function.  

    **.PARAMETER** - ImputArray  
        The mandatory parameter ImputArray is an Array of custom objects.  

    **.EXAMPLE**  
        The example below returns an array of custom objects using the first row as object parameter names and each additional row as object data.  
```
        Set-WorksheetData -Worksheet <PS Excel Worksheet COM Object> -InputArray <PS Object Array>

        PS C:\> Set-WorksheetData -Worksheet $Worksheet -ImputArray $myObjectArray
```
11. **Set-WorksheetName**

    **.DESCRIPTION**  
        This Advance Function sets the name of the given worksheet.  

    **.PARAMETER** - Worksheet  
        The mandatory parameter Worksheet is the Excel worksheet com object passed to the function.  

    **.EXAMPLE**  
        The example below renames the worksheet to Data unless that name is already in use.  
```
        Set-WorksheetName -Worksheet <PS Excel Worksheet COM Object> -SheetName <String>

        PS C:\> Set-WorksheetName -Worksheet $myWorksheet -SheetName "Data"
```
12. **Get-WorksheetNames**

    **.DESCRIPTION**  
        This Advance Function returns an array of strings of all worksheets in a workbook.  

    **.PARAMETER** - Workbook  
        The mandatory parameter Workbook is the Excel workbook com object passed to the function.  

    **.EXAMPLE**  
        The example below renames the worksheet to Data unless that name is already in use.  
```
        Get-WorksheetNames -Workbook <PS Excel Workbook COM Object>

        PS C:\> Get-WorksheetNames -Workbook $myWorkbook
```
13. **ConvertPSObjectToHashtable**

    **.DESCRIPTION**  
        This Advance Function returns a Hashtable converted from a PSObject and will return work with nested PSObjects.  

    **.PARAMETER** - InputObject  
        The mandatory parameter InputObject is a PSObject.  

    **.EXAMPLE**  
        The example below returns a hashtable created from the myPSObject PSObject.  
```
        ConvertPSObjectToHashtable -InputObject <PSObject>

        PS C:\> $myNewHash = ConvertPSObjectToHashtable -InputObject $myPSObject
```
14. **Export-Yaml**

    **.DESCRIPTION**  
        This Advanced Function Exports a Hashtable or PSObject to a Yaml file  

    **.PARAMETER** - InputObject  
        The mandatory parameter InputObject is a Hashtable or PSObject.  

    **.PARAMETER** - Path  
        The mandatory parameter Path is the location string of the Yaml file.  

    **.EXAMPLE**  
        The example below returns a hashtable created from the myPSObject PSObject.  
```
        Export-Yaml -InputObject <PSObject> -Path <String>

        PS C:\> Export-Yaml -InputObject $myHastable -FilePath "C:\myYamlFile.yml"

        or

        PS C:\> Export-Yaml -InputObject $myPSObject -FilePath "C:\myYamlFile.yml"
```
15. **Export-Json**

    **.DESCRIPTION**  
        This Advanced Function Exports a Hashtable or PSObject to a Json file  

    **.PARAMETER** - InputObject  
        The mandatory parameter InputObject is a Hashtable or PSObject.  

    **.PARAMETER** - Path  
        The mandatory parameter Path is the location string of the Json file.  

    **.EXAMPLE**  
        The example below returns a hashtable created from the myPSObject PSObject.  
```
        Export-Json -InputObject <PSObject> -Path <String>

        PS C:\> Export-Json -InputObject $myHastable -FilePath "C:\myJsonFile.json"

        or

        PS C:\> Export-Json -InputObject $myPSObject -FilePath "C:\myJsonFile.json"
```
16. **Import-Jason**

    **.DESCRIPTION**  
        This Advance Function imports a Json file and returns a PSCustomObject.  

    **.PARAMETER** - Path  
        The mandatory parameter Path is the location string of the Json file.  

    **.EXAMPLE**  
        The example below returns a pscustomobject created from the contents of C:\myJasonFile.json.  
```
        Import-Json -Path <String>

        PS C:\> Import-Json -Path "C:\myJsonFile.json"
```
17. **Import-Yaml**

    **.DESCRIPTION**  
        This Advance Function imports a Yaml file and returns a PSCustomObject.  

    **.PARAMETER** - Path  
        The mandatory parameter Path is the location string of the Yaml file.  

    **.EXAMPLE**  
        The example below returns a pscustomobject created from the contents of C:\myYamlFile.yml.  
```
        Import-Yaml -Path <String>

        PS C:\> Import-Yaml -Path "C:\myYamlFile.yml"
```  

18. **Import-ExcelData**  

    **.DESCRIPTION**  
    	This function imports Microsoft Excel worksheets and puts the data in to a hashtable of pscustom objects. The hashtable keys are the names of the Excel worksheets with spaces omitted. The function imports data from all worksheets. It does not validate that the data started in cell A1 and is in format of regular rows and columns, which is required to load the data.  

    **.PARAMETER** - Path  
        The optional parameter Path accepts a path string to the excel file. The string can be either the absolute or relative path.

    **.PARAMETER** - Exclude  
        The optional parameter Exclude accepts a comma separated list of strings of worksheets to exclude from loading.

    **.PARAMETER** - HashtableReturn  
        The optional switch parameter HashtableReturn directs if the return array will contain hashtables or pscustom objects.  

    **.PARAMETER** - TrimHeaders  
        The optional switch parameter TrimHeaders, removes whitespace from the column headers when creating the object or hashtable.  

    **.EXAMPLE**  
        The example below shows the command line use with Parameters.
```
        Import-ExcelData [-Path <String>] [-Exclude <String>,<String>,...] [-HashtableReturn] [-TrimHeaders]

        PS C:\> Import-ExcelData -Path "C:\temp\myExcel.xlsx"

    	or

        PS C:\> Import-ExcelData -Path "C:\temp\myExcel.xlsx" -Exclude "sheet2","sheet3"

```

18. **Read-ExcelPath**  

    **.DESCRIPTION**  
        This function opens a gui window dialog to navigate to an excel file and returns the path.

    **.PARAMETER** - Title  
        The mandatory parameter Title, is a string that appears on the navigation window.

    **.EXAMPLE**  
        The example below shows the command line use with Parameters.
```
        ReadExcelPath -Title <String>

        PS C:\> Read-ExcelPath -Title "Select Microsoft Excel Workbook to Import"
```

## Working Example ##  

Here is an Excel workbook with two worksheets (Virtual Machines and Virtual Networks). The code below will parse all the worksheets. It creates one object per row starting with row two. Where the property names are taken from row one of the column the data is in. The objects are addeded to the object array and returned.

####Sheet Name: Virtual Machines  

1 Hostname  |  Instance Size  |  Internal IP
--------  |  -------------  |  -----------
2 VM01  |  Standard_D4  |  10.10.1.10  
3 VM02  |  Standard_D4  |  10.10.1.11  
4 VM03  |  Standard_D4  |  10.10.1.12  
5 VM04  |  Standard_D4  |  10.10.1.13  

####Sheet Name: Virtual Networks  

1 Name  |  Resource Group  |  Location  |  Vnet Address Prefix  |  Deploy  
----  |  --------------  |  --------  |  -------------------  |  ------  
2 VNET01  |  RG-VNT-01  |  eastus  |  10.10.1.0/20  |  TRUE  
3 VNET01  |  RG-VNT-01  |  northcentralus  |  10.11.1.0/20  |  TRUE  

So for the Worksheet Virtual Machines, the script will create four objects with properties Hostname, Instance Size and Internal IP. The data for each object is taken from each consecutive row starting with row two.  

###Object 1:  

**Property:         Value**  
Hostname:           VM01  
Instance Size:      Standard_D4  
Internal IP:        10.10.1.10  

###Object 2:  

**Property:         Value**  
Hostname:           VM02  
Instance Size:      Standard_D4  
Internal IP:        10.10.1.11  

```  
$data = Import-ExcelData -Path "C:\Excel.xlsx"  

Write-Output $data | Out-String  

```  
