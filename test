Sub GenerateBigData()
    Dim rowCount As Long: rowCount = 1000000
    Dim colCount As Long: colCount = 20
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Dim c As Long, r As Long, i As Long, txt As String, chars As String
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    For c = 1 To colCount: ws.Cells(1, c).Value = "Col_" & c: Next c
    For r = 2 To rowCount
        For c = 1 To colCount
            txt = ""
            For i = 1 To 10
                txt = txt & Mid(chars, Int((Len(chars) * Rnd) + 1), 1)
            Next i
            ws.Cells(r, c).Value = txt
        Next c
        If r Mod 1000 = 0 Then DoEvents
    Next r
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "✅ Готово!"
End Sub
