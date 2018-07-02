# Documentation

#### `method Workbook(&p_filePath As string)`

Workbook constructor

 * **Parameters:** `p_file` — The full path and file name of the Excel file to open or create.
 * **Exceptions:**  — exception is thrown if the file path does not exist or cannot be created.

#### `method CreateSheet(&name As string)`

Creates a worksheet in the workbook with the given name and makes it the active sheet. A sheet with that name already exists, it becomes active. Sheets are modified to meet Excel sheet name requirements.

 * **Parameters:** `name` — The name of the worksheet to create

#### `method Save()`

Saves the Excel file and creates the file if it doesn't exist.

#### `method SetCellString(&row As integer, &column As integer, &string As string)`

Set the value of the given cell to the given string value. If the row or column number is not positive, nothing is done.

 * **Parameters:**
   * `row` — The row number of the cell to set. It must be a positive integer.
   * `column` — The column number of the cell to set. It must be a positive integer.
   * `string` — The string value to set the cell to.

#### `method SetCellNumber(&row As integer, &column As integer, &value As number)`

Set the value of the given cell to the given numeric value. If the row or column number is not positive, nothing is done.

 * **Parameters:**
   * `row` — The row number of the cell to set. It must be a positive integer.
   * `column` — The column number of the cell to set. It must be a positive integer.
   * `value` — The number to set the cell to.

#### `method SetCellNumberFormat(&row As integer, &column As integer, &value As number, &format As string)`

Set the value of the given cell to the given numeric value with the given Excel format. If the row or column number is not positive, nothing is done.

 * **Parameters:**
   * `row` — The row number of the cell to set. It must be a positive integer.
   * `column` — The column number of the cell to set. It must be a positive integer.
   * `number` — The number to set the cell to.
   * `format` — The Excel format to apply to the number.

#### `method SetCellDateTime(&row As integer, &column As integer, &dateTime As datetime)`

Set the value of the given cell to the given date value. The DateTime is formatted as dd/mm/yyyy hh:mm:ss . If the row or column number is not positive, nothing is done.

 * **Parameters:**
   * `row` — The row number of the cell to set. It must be a positive integer.
   * `column` — The column number of the cell to set. It must be a positive integer.
   * `dateTime` — The DateTime value to set the cell to.

#### `method SetCellDate(&row As integer, &column As integer, &date As date)`

Set the value of the given cell to the given date value. The Date is formatted as dd/mm/yyyy. If the row or column number is not positive, nothing is done.

 * **Parameters:**
   * `row` — The row number of the cell to set. It must be a positive integer.
   * `column` — The column number of the cell to set. It must be a positive integer.
   * `dateTime` — The Date value to set the cell to

#### `method SetCellBoolean(&row As integer, &column As integer, &boolean As boolean)`

Set the value of the given cell to the given Boolean value.

 * **Parameters:**
   * `row` — The row number of the cell to set. It must be a positive integer.
   * `column` — The column number of the cell to set. It must be a positive integer.
   * `dateTime` — The Boolean value to set the cell to

#### `method SetCellFormula(&row As integer, &column As integer, &formula As string)`

Set the value of the given cell to the given formula value. If the row or column number is not positive, nothing is done.

 * **Parameters:**
   * `row` — The row number of the cell to set. It must be a positive integer.
   * `column` — The column number of the cell to set. It must be a positive integer.
   * `formula` — The formula value to set the cell to

#### `method SetCellHyperlink(&row As integer, &column As integer, &text As string, &url As string)`

Set the value of the given cell to the given hyperlink. If the row or column number is not positive, nothing is done.

 * **Parameters:**
   * `row` — The row number of the cell to set. It must be a positive integer.
   * `column` — The column number of the cell to set. It must be a positive integer.
   * `text` — The text value of the cell.
   * `column` — The URL the link points to.

#### `method GetCellString(&row As integer, &column As integer) Returns string`

Returns the value of the requested cell as a string. If the cell does not exist (or an invalid cell location is given), an empty string is returned.

 * **Parameters:**
   * `row` — The row number of the cell. It must be a positive integer.
   * `column` — The column number of the cell. It must be a positive integer.
* **Returns:** String - the string value of the cell. If the cell is not a string cell, the value is converted into a string.

#### `method GetCellNumber(&row As integer, &column As integer, &value As number out) Returns boolean`

Retrieves the numeric value of a cell and places its value into the out parameter &value.

 * **Parameters:**
   * `row` — The row number of the cell. It must be a positive integer.
   * `column` — The column number of the cell. It must be a positive integer.
   * `value` — The numeric value of a cell if it contains a numeric value or if cell value can be converted to a number.
   
 * **Returns:** Boolean - a numeric value was successfully retrieved.

#### `method GetCellBoolean(&row As integer, &column As integer, &value As boolean out) Returns boolean`

Retrieves the Boolean value of a cell and places its value into the out parameter &value.

 * **Parameters:**
   * `row` — The row number of the cell. It must be a positive integer.
   * `column` — The column number of the cell. It must be a positive integer.
   * `value` — The Boolean value of a cell if it contains a Boolean value or if cell value can be converted to a Boolean.
   
* **Returns:** Boolean - a numeric value was successfully retrieved.

#### `method GetCellDateTime(&row As integer, &column As integer, &value As datetime out) Returns boolean`

Retrieves the DateTime value of a cell and places its value into the out parameter &value.

 * **Parameters:**
   * `row` — The row number of the cell. It must be a positive integer.
   * `column` — The column number of the cell. It must be a positive integer.
   
 * **Returns:** Boolean - a DateTime value was successfully retrieved.
 
#### `method RenameSheet(&oldName As string, &newName As string)`

Rename a worksheet. If the requested worksheet does not exist, nothing happend.

 * **Parameters:**
   * `oldName` — The name of the old worksheet to rename
   * `newName` — The new name for the worksheet
   
#### `method RenameCurrentSheet(&newName As string)`

Rename the current worksheet.

 * **Parameters:**
   * `newName` — The new name for the worksheet
   
#### `method SelectSheet(&name As string)`

The requested worksheet becomes the current worksheet. If the requested worksheet does not exist, nothing happend.

 * **Parameters:**
   * `name` — The name of the worksheet to select  
