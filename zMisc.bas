Attribute VB_Name = "zMisc"
Function ToCSV(arVals() As Variant) As String
    Dim val As Variant
    Dim str As String
    For Each val In arVals()
        str = str & val & ","
    Next val
    'trim off last lingering ","
    str = Left(str, Len(str) - 1)
    ToCSV = str
End Function

Function sheetExists(sheetName As String, Optional inWbk As Workbook) As Boolean
    
    Dim sht As Worksheet
    Dim returnMe As Boolean
    
    If inWbk Is Nothing Then
        Set inWbk = ThisWorkbook
    End If
    
    returnMe = False
    
    For Each sht In inWbk.Sheets
        If sht.Name = sheetName Then
            returnMe = True
            Exit For
        End If
    Next sht
 
    sheetExists = returnMe

End Function
