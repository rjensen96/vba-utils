Attribute VB_Name = "Navigation"
Sub jumpToSheet() 'jumps to a sheet. workbook is a huge pain to navigate by hand.
Attribute jumpToSheet.VB_ProcData.VB_Invoke_Func = "J\n14"

    Dim shtName As String
    Dim sht As Worksheet
    Dim wb As Workbook
    Dim found As Boolean
    Dim msg As Variant
    
    Set wb = ActiveWorkbook
    shtName = Application.InputBox("Enter sheet name", "Jump to Sheet")
    
    found = False
    
    If Left(shtName, 1) = "*" Then
        shtName = Right(shtName, Len(shtName) - 1)
        For Each sht In Sheets
            If UCase(Right(sht.Name, Len(shtName))) = UCase(shtName) Then
                found = True
                Exit For
            End If
        Next sht
    ElseIf Right(shtName, 1) = "*" Then
        shtName = Left(shtName, Len(shtName) - 1)
        For Each sht In Sheets
            If UCase(Left(sht.Name, Len(shtName))) = UCase(shtName) Then
                found = True
                Exit For
            End If
        Next sht
    Else 'no wildcard
        For Each sht In Sheets
            If UCase(sht.Name) = UCase(shtName) Then
                found = True
                Exit For
            End If
        Next sht
    End If
    
    'Set sht = wb.Sheets(shtName)
    If found = True Then
        sht.Activate
    Else
        msg = MsgBox(Prompt:="Sheet not found. Please search again.", Title:="Search Failed") = vbOK
    End If
    
End Sub


