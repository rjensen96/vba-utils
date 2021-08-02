Attribute VB_Name = "zListObjects"
Function deleteRowsByValue(tbl As ListObject, colName As String, val As Variant)
    Dim colChange As ListColumn
    Set colChange = tbl.ListColumns(colName)
    
    'I am aware that loops exist, but looping through ListRows and deleting one by one is THE PITS.
    'IT CAN TAKE ABSOLUTELY FOREVER....LIKE, MINUTES IN LARGE DATA SETS.
    'so just filter it out, delete it, and move on.
    
    'This function only safe on sheets with one table... not the greatest
    
    tbl.Range.AutoFilter field:=colChange.Index, Criteria1:=val
    tbl.DataBodyRange.EntireRow.Delete
    tbl.Range.AutoFilter field:=colChange.Index
    
End Function

Function columnExists(tbl As ListObject, colName As String) As Boolean
    Dim col As ListColumn
    On Error GoTo jump
    Set col = tbl.ListColumns(colName)
    columnExists = True
jump:
End Function

Function getTableValue(tbl As ListObject, fieldSearch As String, itemSearch As Variant, fieldGet As String) 'this function, I think, could be used universally.

    'TODO 8/12/2020 -
    'make this more efficient by reading the listcolumns into arrays, then looping through those in memory.

    Dim rng As Range
    Dim colSearch As ListColumn
    Dim colGet As ListColumn

    Dim n As Integer
    
    On Error GoTo jump
    Set colSearch = tbl.ListColumns(fieldSearch)
    Set colGet = tbl.ListColumns(fieldGet)
    
    For n = 1 To tbl.ListRows.Count
        If colSearch.DataBodyRange(n, 1).Value = itemSearch Then
            getTableValue = colGet.DataBodyRange(n, 1).Value
            Exit For
        End If
    Next n
    
    Exit Function
    
jump:
    getTableValue = Null
    
End Function

Function setTableValue(tbl As ListObject, fieldSearch As String, itemSearch As Variant, fieldSet As String, valSet As Variant) 'this function, I think, could be used universally.

    Dim rng As Range
    Dim colSearch As ListColumn
    Dim colSet As ListColumn

    Dim n As Integer
    
    On Error GoTo jump
    Set colSearch = tbl.ListColumns(fieldSearch)
    Set colSet = tbl.ListColumns(fieldSet)
    
    'note - this sets ALL found values; does not stop on first occurrence like getTableValue.
    For n = 1 To tbl.ListRows.Count
        If colSearch.DataBodyRange(n, 1).Value = itemSearch Then
            colSet.DataBodyRange(n, 1).Value = valSet
        End If
    Next n
    
jump:
    
End Function

Function listColumnIsNumeric(lo As ListObject, colName As String) As Boolean
    Dim lc As ListColumn
    Dim arVals() As Variant
    Dim itm As Variant
    Dim rtn As Boolean
    
    Set lc = lo.ListColumns(colName)
    
    If Not IsEmpty(lc.DataBodyRange.Value) Then
        If lc.DataBodyRange.Cells.Count > 1 Then
            arVals() = lc.DataBodyRange.Value
            rtn = True 'default, now disprove it.
            
            For Each itm In arVals()
                If Not IsNumeric(itm) Then
                    rtn = False
                    Exit For
                End If
            Next itm
        ElseIf IsNumeric(lc.DataBodyRange.Value) Then 'if table has just one row, arVals() fails so we have to use this instead
            rtn = True
        End If
        
    End If
    
    listColumnIsNumeric = rtn

End Function

Function makeListObject(sht As Worksheet, tblName As String) As ListObject
    
    'REQUIRED: DATA MUST START IN CELL A1!!!
    
    Dim rngTable As Range
    Dim loRtn As ListObject
    
    If sht.Range("A1").Value = "" Then 'no data found
        Set makeListObject = Nothing
        Exit Function
    End If
    
    Set rngTable = sht.Range("A1")
    Set rngTable = Range(rngTable, rngTable.End(xlToRight))
    Set rngTable = Range(rngTable, rngTable.End(xlDown))
    
    On Error Resume Next
    Set loRtn = rngTable.ListObject 'in case of second run, dont duplicate the table.
    On Error GoTo 0
    
    If loRtn Is Nothing Then
        Set loRtn = sht.ListObjects.Add(xlSrcRange, rngTable, , xlYes)
    End If
    
    loRtn.TableStyle = ""
    loRtn.Name = tblName
    
    Set makeListObject = loRtn
    
End Function

Function getListObject(tblName As String, Optional wbk As Workbook) As ListObject
    'WHY THIS, YOU MIGHT ASK?
    'because ListObject is a child of Worksheet, not workbook
    'DESPITE THE FACT THAT ALL LISTOBJECTS MUST HAVE UNIQUE NAMES.
    'YOU HAVE TO LOOP THROUGH ALL THE SHEETS JUST TO FIND SOMETHING THAT IS ALREADY UNIQUE.
    'so I'm solving that problem permanently.
    
    Dim sht As Worksheet
    Dim loReturn As ListObject
    
    If wbk Is Nothing Then
        Set wbk = ThisWorkbook
    End If
    
    If tblName = "" Then
        Set getListObject = Nothing
        Exit Function
    End If
    
    On Error Resume Next
    For Each sht In wbk.Sheets
        Set loReturn = sht.ListObjects(tblName)
        If Not loReturn Is Nothing Then 'WE FOUND IT.
            Exit For
        End If
    Next sht
    On Error GoTo 0
    
    Set getListObject = loReturn 'either is a table, or Nothing.
    
End Function

Function mergeTables(loSrc As ListObject, loDest As ListObject, Optional NewColIfNotFound As Boolean)

    'nested loops to bring in data for each matching column.
    'option to create new column in loDest if not found in source.
    'THIS DOES NOT HAVE OPTIMIZATION BUILT-IN...turn of screenupdating/calculation OUTSIDE this function. this is just the function.
    
    Dim colSrc As ListColumn
    Dim colDest As ListColumn
    Dim rngPost As Range
    Dim x, y As Integer
    
    x = loDest.ListRows.Count
    
    For Each colDest In loDest.ListColumns
        For Each colSrc In loSrc.ListColumns
            If colSrc.Name = colDest.Name Then 'copy data
                Set rngPost = colDest.DataBodyRange(x + 1, 1)
                Set rngPost = rngPost.Resize(loSrc.ListRows.Count, 1)
                rngPost.Value = colSrc.DataBodyRange.Value
            End If
        Next colSrc
    Next colDest
    
    If NewColIfNotFound = True Then 'make new columns for unique columns in loSrc and post data
        For Each colSrc In loSrc.ListColumns
            Set colDest = Nothing
            
            On Error Resume Next
            Set colDest = loDest.ListColumns(colSrc.Name)
            On Error GoTo 0
            
            If colDest Is Nothing Then
                Set colDest = loDest.ListColumns.Add
                colDest.Name = colSrc.Name
                Set rngPost = colDest.DataBodyRange(x + 1, 1)
                Set rngPost = rngPost.Resize(loSrc.ListRows.Count, 1)
                rngPost.Value = colSrc.DataBodyRange.Value
            End If
        Next colSrc
    End If

End Function

