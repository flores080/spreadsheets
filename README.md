# Move Row to another Sheet by Cell value

```visualbasic
Sub MoveRowBasedOnCellValue()
'Updated by Extendoffice 2017/11/10
    Dim xRg As Range
    Dim xCell As Range
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim source As String
    Dim destiny As String
    Dim keyWord As String
    Dim keyWordRange As String
    keyWord = "CSYDR"
    source = "PROPUESTA HORARIO"
    destiny = keyWord & " (2)"
    keyWordRange = "H7:H"
    I = Worksheets(source).UsedRange.Rows.Count
    J = 7
    If J = 1 Then
    If Application.WorksheetFunction.CountA(Worksheets(destiny).UsedRange) = 0 Then J = 0
    End If
    Set xRg = Worksheets(source).Range(keyWordRange & I)
    On Error Resume Next
    Application.ScreenUpdating = True
    For K = 1 To xRg.Count
        If CStr(xRg(K).Value) = keyWord Then
            xRg(K).EntireRow.Copy Destination:=Worksheets(destiny).Range("A" & J)
            J = J + 1
        End If
    Next
    Application.ScreenUpdating = True
End Sub
```
