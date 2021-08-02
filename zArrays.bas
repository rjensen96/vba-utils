Attribute VB_Name = "zArrays"
Function arContains(inArr() As Variant, item As String) As Boolean
    'for ONE-DIMENSIONAL ARRAYS ONLY.
    
    Dim i As Integer
    Dim rtn As Boolean
    
    rtn = False
    
    For i = LBound(inArr) To UBound(inArr)
        If inArr(i) = item Then
            rtn = True
            Exit For
        End If
    Next i
    
    arContains = rtn
    
End Function

Function addUniqueToArray(arAdd() As Variant, val As String) As Variant
    Dim str As Variant
    
    For Each str In arAdd()
        If str = val Then
            GoTo jump
        End If
    Next str

    ReDim Preserve arAdd(UBound(arAdd()) + 1)
    arAdd(UBound(arAdd())) = val
    
jump:
    addUnique = arAdd()
    
End Function

