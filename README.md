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

`getListObject(tblName As String, Optional wbk As Workbook) As ListObject`

*Purpose*<br/>
Searches a workbook for a table with provided name and returns the table as ListObject.

*Arguments*<br/>
`String` Table name to get<br>
Optional `Workbook` to search within; defaults to ThisWorkbook.

*Returns*<br/>`ListObject`

---
**columnExists**

`columnExists(tbl As ListObject, colName As String) As Boolean`

*Purpose*<br/>
Determines if a column exists in a table.

*Arguments*<br/>
`ListObject` table to check <br>
`String` column name to search for <br>

*Returns*<br/>`Boolean`

---
**getTableValue**

`getTableValue(tbl As ListObject, fieldSearch As String, itemSearch As Variant, fieldGet As String)`

*Purpose*<br/>
Performs a lookup in a table: searches for a value in one field, returns the value on the same record in another field. Returns value at first occurrence top-to-bottom.

*Arguments*<br/>
`ListObject` table to search <br>
`String` column name to search in <br>
`Variant` value to search for in first column <br>
`String` column name to retrieve value from <br>

*Returns*<br>
`Variant` value from table
