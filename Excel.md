# Readme #
## Project: PowerShell Excel Module ##

This repo contains the Excel PowerShell module for manipulating Microsoft Excel Workbooks, Worksheets, etc. The main purpose of the module is to allow data retrieval from Excel for use in automation scripts and configuration scripts.

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


## Function Help ##
1. **Open-Excel** - Creates a new excel COM object.  

**.DESCRIPTION**  
    The function opens an instance of Microsoft Excel but keeps it hidden unless the Visible parameter is used.  

**.PARAMETER** - Visible  
    The parameter switch Visible when specified will make Excel visible on the desktop.  

**.EXAMPLE**  
    The example below returns the Excel COM object when used.  
```
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
        PS C:\> Close-Excel -ObjExcel $myObjExcel -Workbook $wb  
```
3. **Get-Workbook**

    **.DESCRIPTION**  
        Given the Microsoft Excel COM Object and Path to the Excel file, the function retuns the Workbook COM Object.  

    **.PARAMETER** - ObjExcel  
        The mandatory parameter ObjExcel is the Excel COM Object passed to the function.  

    **.PARAMETER** - Path  
        The mandatory parameter Path is the location string of the Excel file. Relative and Absolute paths are supported.  

    **.EXAMPLE**  
        The example below returns the workbook COM object specified by Path.  
```
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
        The parameter Sheetname is a string passed to the function to name the newly created worksheet.  

    **.EXAMPLE**  
        The example below creates a new worksheet named Data.  
```
        PS C:\> Add-Worksheet -ObjExcel $myObjExcel -Workbook $wb -Sheetname "Data"  
```
6. **Add-Workbook**

    **.DESCRIPTION**  
        Given the Microsoft Excel COM Object and Path to the Excel file, the function retuns the Workbook COM Object.  

    **.PARAMETER** - Path  
        The mandatory parameter Path is the location string of the Excel file.  

    **.PARAMETER** - ObjExcel  
        The mandatory parameter ObjExcel is needed to retrieve the Workbook COM Object.  

    **.EXAMPLE**  
        The example below returns the newly created Excel workbook COM Object.  
```        
        PS C:\> Add-Workbook -ObjExcel $myExcelObj -Path "C:\Excel.xlsx"  
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
        PS C:\> Save-Workbook -Workbook $wb -Path "C:\Excel.xlsx"  
```
8. **Get-WorksheetUsedRange**

    **.DESCRIPTION**  
        This advanced function returns a hashtable containing the last used column and last used row of a worksheet..  

    **.PARAMETER** - Worksheet  
        The parameter Worksheet is the Excel worksheet com object passed to the function.  

    **.EXAMPLE**  
        The example below returns a hashtable containing the last used column and row of the referenced worksheet.  
```        
        PS C:\> Get-WorksheetUsedRange $Worksheet  
```
9. **Get-WorksheetData**

    **.DESCRIPTION**  
        This advanced function creates an array of pscustom objects from an Microsoft Excel worksheet.  
        The first row will be used as the object members and each additional row will form the object data for that member.  

    **.PARAMETER** - Worksheet  
        The parameter Worksheet is the Excel worksheet com object passed to the function.  

    **.PARAMETER** - HashtableReturn  
        The switch parameter HashtableReturn with default value False, causes the function to return an array of hashtables instead of an array of objects.  

    **.EXAMPLE**  
        The example below returns an array of custom objects using the first row as object parameter names and each additional row as object data.  
```        
        PS C:\> Get-WorksheetData $Worksheet  
```
10. **Set-WorksheetData**

    **.DESCRIPTION**  
        This advanced function populates a Microsoft Excel Worksheet with data from an Array of custom objects. The object members populates the first row of the sheet as header items.  
        The object values are placed beneath the headers on each successive row.  

    **.PARAMETER** - Worksheet  
        The parameter Worksheet is the Excel worksheet com object passed to the function.  

    **.PARAMETER** - ImputArray  
        The parameter ImputArray is an Array of custom objects.  

    **.PARAMETER** - HashtableReturn  
        The switch parameter HashtableReturn with default value False, causes the function to return an array of hashtables instead of an array of objects.  

    **.EXAMPLE**  
        The example below returns an array of custom objects using the first row as object parameter names and each additional row as object data.  
```        
        PS C:\> Set-WorksheetData $Worksheet  
```
11. **Set-WorksheetName**

    **.DESCRIPTION**  
        This Advance Function sets the name of the given worksheet.  

    **.PARAMETER** - Worksheet  
        The parameter Worksheet is the Excel worksheet com object passed to the function.  

    **.EXAMPLE**  
        The example below renames the worksheet to Data unless that name is already in use.  
```        
        PS C:\> Set-WorksheetName -Worksheet $ws -SheetName "Data"  
```
12. **Get-WorksheetNames**

    **.DESCRIPTION**  
        This Advance Function returns an array of strings of all worksheets in a workbook.  

    **.PARAMETER** - Workbook  
        The parameter Workbook is the Excel workbook com object passed to the function.  

    **.EXAMPLE**  
        The example below renames the worksheet to Data unless that name is already in use.  
```        
        PS C:\> Get-WorksheetNames -Workbook $wb  
```
13. **ConvertPSObjectToHashtable**

    **.DESCRIPTION**  
        This Advance Function returns a Hashtable converted from a PSObject and will return work with nested PSObjects.  

    **.PARAMETER** - InputObject  
        The parameter InputObject is a PSObject.  

    **.EXAMPLE**  
        The example below returns a hashtable created from the myPSObject PSObject.  
```        
        PS C:\> $myNewHash = ConvertPSObjectToHashtable $myPSObject  
```
14. **Export-Yaml**

    **.DESCRIPTION**  
        This Advanced Function Exports a Hashtable or PSObject to a Yaml file  

    **.PARAMETER** - InputObject  
        The parameter InputObject is a Hashtable or PSObject.  

    **.PARAMETER** - Path  
        The mandatory parameter Path is the location string of the Yaml file.  

    **.EXAMPLE**  
        The example below returns a hashtable created from the myPSObject PSObject.  
```        
        PS C:\> Export-Yaml -InputObject $myHastable -FilePath "C:\myYamlFile.yml"  

        or  

        PS C:\> Export-Yaml -InputObject $myPSObject -FilePath "C:\myYamlFile.yml"  
```
15. **Export-Json**

    **.DESCRIPTION**  
        This Advanced Function Exports a Hashtable or PSObject to a Json file  

    **.PARAMETER** - InputObject  
        The parameter InputObject is a Hashtable or PSObject.  

    **.PARAMETER** - Path  
        The mandatory parameter Path is the location string of the Json file.  

    **.EXAMPLE**  
        The example below returns a hashtable created from the myPSObject PSObject.  
```        
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
        PS C:\> Export-Json -FilePath "C:\myJsonFile.json"  
```
17. **Import-Yaml**

    **.DESCRIPTION**  
        This Advance Function imports a Yaml file and returns a PSCustomObject.  

    **.PARAMETER** - Path  
        The mandatory parameter Path is the location string of the Yaml file.  

    **.EXAMPLE**  
        The example below returns a pscustomobject created from the contents of C:\myYamlFile.yml.  
```        
        PS C:\> Export-Json -FilePath "C:\myYamlFile.yml"  
```       
