# CY2 SpreadSheet
Read and write Excel spreadsheets from PeopleSoft

The PeopleSoft delivered PSSpredSheet is lacking in features (cannot read data) and a little buggy.  This is a PeopleCode wrapper around Apache POI library that ships with PeopleSoft. Works with PeopleTools 8.55 or later.

# Known Issues
1. Timezone issues
2. Only writes XLSX (but do care about XLS anymore?)
3. Needs better/real tests

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
