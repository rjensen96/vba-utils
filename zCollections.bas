Attribute VB_Name = "zCollections"
Function getCollectionFromCSV(str As String)
    Dim col As New Collection
    Dim itm As Variant
    Dim ar() As String
    
    ar() = Split(str, ",")
    
    For Each itm In ar()
        col.Add itm, CStr(itm)
    Next itm
    
    Set getCollectionFromCSV = col
    
End Function

Function collectionToCSV(col As Collection) As String
    'ONLY WORKS IF ITEMS ARE NON-OBJECTS.
    Dim itm As Variant
    Dim str As String
    
    If col.Count > 0 Then
        For Each itm In col
            str = str & itm & ","
        Next itm
    End If
    
    If str <> "" Then
        str = Left(str, Len(str) - 1) 'remove last ","
    End If
    
    collectionToCSV = str
    
End Function

Function getUniqueCollection(loGet As ListObject, col As String, Optional sortType As String) As Collection

    Dim liCol As ListColumn
    Dim colVals As New Collection
    Dim arVals() As Variant
    Dim itm As Variant
    Dim srcItm As Variant
    Dim rng As Range
    Dim bKeep As Boolean
    
    Set liCol = loGet.ListColumns(col)
    
    If Not IsEmpty(liCol.DataBodyRange.Value) Then
    
        arVals() = liCol.DataBodyRange.Value
        
        For Each srcItm In arVals
            bKeep = True
            
            If KeyExists(colVals, CStr(srcItm)) = False Then
                colVals.Add (srcItm), CStr(srcItm)
            End If
            
        Next srcItm
        
        'sort if user wants it sorted.
        If UCase(sortType) = "ASCENDING" Then
            Set colVals = sortCollection(colVals)
        ElseIf UCase(sortType) = "DESCENDING" Then
            Set colVals = sortCollection(colVals, True)
        End If
    
    End If
    
    Set getUniqueCollection = colVals

End Function

Function tableToDictionary(lo As ListObject, keyField As String) As Dictionary
    Dim rw As ListRow
    Dim dTable As New Dictionary
    Dim k As String
    
    For Each rw In lo.ListRows
        k = rw.Range(1, lo.ListColumns(keyField).Index).Value
        dTable(k) = rw.Range.Value
    Next rw
    
    Set tableToDictionary = dTable
    
End Function

Function sortCollection(colSort As Collection, Optional descending As Boolean) As Collection

    'default to ascending sort.
    
    Dim i, j As Long
    Dim tmp As Variant
    
    If colSort.Count > 1 Then
        
        For i = 1 To colSort.Count - 1
            For j = i + 1 To colSort.Count
                If Not descending Then
                    If colSort(i) > colSort(j) Then
                        tmp = colSort(j)
                        colSort.Remove j
                        colSort.Add tmp, CStr(tmp), i
                    End If
                Else
                    If colSort(i) < colSort(j) Then
                        tmp = colSort(j)
                        colSort.Remove j
                        colSort.Add tmp, CStr(tmp), i
                    End If
                End If
            Next j
        Next i
        
    End If
    
    Set sortCollection = colSort
    
End Function

Function addUniqueToCollection(col As Collection, val As Variant) As Collection
    Dim itm As Variant
    Dim bKeep As Boolean
    
    bKeep = True
    
    For Each itm In col
        bKeep = True
        If itm = val Then
            bKeep = False
            Exit For
        End If
    Next itm
    
    If bKeep = True Then
        col.Add (val)
    End If
    
    Set addUniqueToCollection = col
    
End Function

Function KeyExists(coll As Collection, key As String) As Boolean
' https://excelmacromastery.com/

    On Error GoTo EH
    IsObject (coll.Item(key))
    KeyExists = True
EH:

End Function

Function InCollection(coll As Collection, itm As Variant) As Boolean
    'Works for strings, ints, etc.
    'NOT TESTED WITH OBJECTS
    
    Dim val As Variant
    Dim rtn As Boolean
    
    rtn = False
    For Each val In coll
        If val = itm Then
            rtn = True
        End If
    Next val
    
    InCollection = rtn
    
End Function
