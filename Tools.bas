Attribute VB_Name = "Tools"
Private Sub CreateSheet()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(After:= _
             ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.Name = "Temp"
    ActiveWorkbook.Sheets("Temp").Visible = xlSheetVeryHidden
End Sub
Private Sub DeleteSheet()
    ActiveWorkbook.Sheets("Temp").Visible = xlSheetVisible
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("Temp").Delete
    Application.DisplayAlerts = True
End Sub

Public Function Col_Letter(lngCol As Long) As String
    Dim varr
    varr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = varr(0)
End Function


