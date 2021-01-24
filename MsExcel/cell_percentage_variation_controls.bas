Sub changePercentage()

Application.EnableEvents = False
Dim n As String: n = ActiveSheet.Shapes(Application.Caller).Name
Dim S As Shape: Set S = ActiveSheet.Shapes(n)
Dim C As Range: Set C = Cells(S.TopLeftCell.Row, S.TopLeftCell.Column)
S.LockAspectRatio = msoFalse
S.Width = C.Height * 0.9
S.Height = S.Width
S.LockAspectRatio = msoTrue

If InStr(1, n, "Plus") Then
    C.Value = C.Value + 1
    S.Left = C.Left + C.Width * 0.9 - S.Width
    S.Top = C.Top
ElseIf InStr(1, n, "Minus") Then
    C.Value = C.Value - 1
    S.Left = C.Left + C.Width * 0.1
    S.Top = C.Top
End If
  
'ThisWorkbook.RefreshAll
Call updateAllPvt("tree")
Application.EnableEvents = True

End Sub