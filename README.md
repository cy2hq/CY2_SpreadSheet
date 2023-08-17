# CY2 SpreadSheet

Read and write Excel spreadsheets from PeopleSoft.

This is a PeopleCode wrapper around Apache POI library that ships with PeopleSoft. Works with PeopleTools 8.55 or later.

The code is in a two files [(CY2_SPREADSHEET:Workbook.pcode)](CY2_SPREADSHEET:Workbook.pcode) and [(CY2_SPREADSHEET:StreamingWorkbook.pcode)](CY2_SPREADSHEET:StreamingWorkbook.pcode), or you can import the Application Designer package in the CY2_SPREADSHEET directory. Note: The StreamingWorkbork is only for creating Excel files (usually large ones), and the API is the same except the read methods have been removed. Feel free to change the application package name to suite your naming conventions, but please keep the copyright notice in the code.

# [Documentation](Documentation.md)

# Examples

## Creating a spreadsheet

```
import CY2_SPREADSHEET:Workbook;

Local CY2_SPREADSHEET:Workbook &workbook = create CY2_SPREADSHEET:Workbook(&fullPath);
&workbook.SetCellString(1, 1, "test");
&workbook.SetCellNumber(1, 2, 1);
&workbook.SetCellNumber(1, 3, 2.2);
&workbook.SetCellNumberFormat(1, 4, 3.3, "00.000");
&workbook.SetCellFormula(2, 1, "=b1+c1");
&workbook.SetCellFormula(2, 2, "c1+d1");
....
&workbook.Save();
```

## Reading data from a spreadsheet

```
import CY2_SPREADSHEET:Workbook;

Local CY2_SPREADSHEET:Workbook &workbook = create CY2_SPREADSHEET:Workbook(&fullPath);

Local string &string = &workbook.GetCellString(1, 1);

Local number &number;
Local boolean &success = &workbook.GetCellNumber(1, 2, &number);

Local boolean &boolean;
&success = &workbook.GetCellBoolean(2, 3, &boolean);

Local DateTime &datetime;
&success = &workbook.GetCellDateTime(4, 1, &datetime);
```
