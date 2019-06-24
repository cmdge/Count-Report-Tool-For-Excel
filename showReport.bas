Attribute VB_Name = "showReport"
Sub object_Click()

 If WorksheetFunction.CountA(ActiveSheet.UsedRange) = 0 And ActiveSheet.Shapes.Count = 0 Then
        MsgBox "Sheet is empty", vbInformation, "KuTools For Excel"
    Else
        ReportTool.Show
    End If
    
End Sub
    





 

