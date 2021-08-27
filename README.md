# vba-utils

## Hang on a minute because I need to write more here.

These files are a set of functions I wrote to make VBA in Excel less terrible. I hope these functions will make VBA less terrible for you, too.

In time, I'll write up some documentation for these and post that here. For now, just know that the most useful of these files are `zListObjects.bas` and `zCollections.bas`. I'm stashing the other files here because they exist but I will comb through these and I may destroy some of the other ones. 

The functions in `zListObjects` and `zCollections` are fairly well-tested and should perform decently well. No guarantees, though. Log an issue or PR if you would like to make these better.

## Installation

From the VBA editor in Excel, create a new module and paste in the functions you'd like to use. 

## Usage

These functions will be available in the public scope within your VBA project. Call the functions from any module in your project. Here is an example call of the the `getListObject` function in any module:

```
Sub main() 
  Dim tblSales as ListObject
  Set tblSales = getListObject("Sales") 'gets a table named "Sales" in ThisWorkbook.
  Debug.print(tblSales.ListRows.count) 'prints number of rows in sales table.
End Sub
```
## Documentation

### zListObjects.bas
---
**getListObject**

`getListObject(String, [ Workbook ]) As ListObject`

*Purpose*<br/>
Searches a workbook for a table with provided name and returns the table as ListObject.

*Arguments*<br/>
`String` Table name to get<br>
Optional `Workbook` to search within; defaults to ThisWorkbook.

*Returns*<br/>`ListObject`

---
**columnExists**

`columnExists(ListObject, String) As Boolean`

*Purpose*<br/>
Determines if a column exists in a table.

*Arguments*<br/>
`ListObject` table to check <br>
`String` column name to search for <br>

*Returns*<br/>`Boolean`

---
**getTableValue**

`getTableValue(ListObject, String, Variant, String)`

*Purpose*<br/>
Performs a lookup in a table: searches for a value in one field, returns the value on the same record in another field. Returns value at first occurrence top-to-bottom.

*Arguments*<br/>
`ListObject` table to search <br>
`String` column name to search in <br>
`Variant` value to search for in first column <br>
`String` column name to retrieve value from <br>

*Returns*<br>
`Variant` value from table

---
**setTableValue**

`setTableValue(ListObject, String, Variant, String, Variant)`

*Purpose*<br/>
Finds a value in one column and changes the value in a different column on the same record. Modifies record of first occurrence, top-to-bottom.

*Arguments*<br/>
`ListObject` table to search <br>
`String` column name to search in <br>
`Variant` value to search for in first column <br>
`String` column name to set new value in <br>
`Variant` new value to set <br>

*Returns*<br>
No return value.

---
**listColumnIsNumeric**

`listColumnIsNumeric(ListObject, String)`

*Purpose*<br/>
Determines if all values in a ListColumn are numeric.

*Arguments*<br/>
`ListObject` table to search <br>
`String` name of ListColumn to analyze <br>

*Returns*<br>
`Boolean`

---
**makeListObject**

`makeListObject(Worksheet, String)`

*Purpose*<br/>
Creates a table with a given name on a specific worksheet starting in cell A1. Assumes that the data is unbroken in Column A and Row 1 and begins in A1. Does nothing if table already exists.

*Arguments*<br/>
`Worksheet` sheet with data for table <br>
`String` name of new table <br>

*Returns*<br>
`ListObject`

---
**mergeTables**

`mergeTables(ListObject, ListObject, [ Boolean ])`

*Purpose*<br/>
Appends the data from all columns in one table to the end of the matching columns of another table. Optionally adds new columns if columns in the source table are not found.

*Arguments*<br/>
`ListObject` table to read from <br>
`ListObject` table to append value into <br>
Optional `Boolean` create new columns if column not found.

*Returns*<br>
No return value.

---
**deleteRowsByValue**

`deleteRowsByValue(ListObject, String, Variant)`

*Purpose*<br/>
Filters a ListObject to a specified value and deletes the **entire sheet row** at those values.
**This is not a safe function if there is anything other than the table on the table's worksheet.**

*Arguments*<br/>
`ListObject` table to delete rows from <br>
`String` column with value to delete <br>
`Variant` value in records to delete <br>

*Returns*<br>
No return value.

