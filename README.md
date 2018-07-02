# CY2 SpreadSheet
Read and write Excel spreadsheets from PeopleSoft

The PeopleSoft delivered PSSpredSheet is a little lacking in features (cannot read data) and a little buggy.  This a pure PeopleCode project that uses the Apache POI libraries that ship with PeopleSoft. Works with PeopleTools 8.55 or later.

# Known Issues
1. Timezone issues
2. Many missing methods
3. Only writes XLSX (but do care about XLS anymore?)
4. Needs real tests

# [Documentation](Documentation.md)
# Examples
## Creating a spreadsheet
```
import CY2_SPREADSHEET:Workbook;

Local CY2_SPREADSHEET:Workbook &sheet = create CY2_SPREADSHEET:Workbook(&path);
&sheet.SetCellString(1, 1, "test");
&sheet.SetCellNumber(1, 2, 1);
&sheet.SetCellNumber(1, 3, 2.2);
&sheet.SetCellNumberFormat(1, 4, 3.3, "00.000");
&sheet.SetCellFormula(2, 1, "=b1+c1");
&sheet.SetCellFormula(2, 2, "c1+d1");

&sheet.Save();
```

## Reading data from a spreadsheet
```import CY2_SPREADSHEET:Workbook;

Local CY2_SPREADSHEET:Workbook &sheet = create CY2_SPREADSHEET:Workbook(&path);

Local string &string = &sheet.GetCellString(1, 1);

Local number &number;
Local boolean &success = &sheet.GetCellNumber(1, 2, &number);

Local boolean &boolean;
&success = &sheet.GetCellBoolean(2, 3, &boolean);

Local DateTime &datetime;
&success = &sheet.GetCellDateTime(4, 1, &datetime);
```
