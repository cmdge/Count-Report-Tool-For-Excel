Attribute VB_Name = "Module2"
Option Explicit

Sub Macro1()

    With Worksheets("sheet1")
        .Range("A1:B27").AdvancedFilter Action:=xlFilterCopy, _
                             CopyToRange:=.Range("H1:I1"), Unique:=True
    End With

End Sub

Sub uniKue()
    Dim i As Long, N As Long, s As String, r As Range
    Dim offs As Variant
    N = Cells(Rows.count, "A").End(xlUp).Row
    For i = 2 To N
        Cells(i, 5) = Cells(i, 1) & " " & Cells(i, 2)
        Cells(i, 6) = Cells(i, 5)
    Next i

    
    Range("F:F").RemoveDuplicates Columns:=1, Header:=xlNo
    For Each r In Range("F:F").SpecialCells(2).Offset(, 1)
        r.Formula = "=COUNTIF(E:E," & r.Offset(, -1).Address & ")"
    Next r
End Sub


Sub mema()
Dim objCellCount As Object
Set objCellCount = CreateObject("scripting.dictionary")
Dim i As Long, N As Long, s As String, r As Range
Dim celldt As Variant
Dim k As Variant
Dim strTmp As String
Dim Val As Variant
    N = Cells(Rows.count, "A").End(xlUp).Row
    For i = 2 To N
        Cells(i, 5) = Cells(i, 1) & " " & Cells(i, 2)
        celldt = Cells(i, 5)
    Debug.Print "celldt: " & celldt
    Next i
    
    For Each k In celldt
        strTmp = Trim(k)
        If Len(strTmp) > 0 Then objCellCount(strTmp) = objCellCount(strTmp) + 1
    Next k
    
    For Each Val In objCellCount.keys
        Debug.Print Val & " " & objCellCount(Val)
    Next Val
End Sub


Sub UniqueCustomerVehiclesCodeList()
  Dim X, vRws
  Dim objDict As Object
  Dim lngRow As Long, lngLastRow

  Set objDict = CreateObject("Scripting.Dictionary")
  objDict.CompareMode = 1
  With Sheets("Sheet3")
    lngLastRow = .Range("K" & .Rows.count).End(xlUp).Row
    vRws = Evaluate("row(2:" & lngLastRow & ")")
    X = Application.Index(.Cells, vRws, Array(11, 12, 13))  '11, 12, 13 are columns K, L, M
  End With
  For lngRow = 1 To UBound(X, 1)
    objDict(X(lngRow, 1) & "|" & X(lngRow, 2) & "|" & X(lngRow, 3)) = 1
  Next
  With Sheets("Sheet4")
    .UsedRange.Resize(, 3).ClearContents
    With .Range("A1:A" & objDict.count)
      .Value = Application.Transpose(objDict.keys)
      .TextToColumns Destination:=.Cells(1), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                     ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
                     Other:=True, OtherChar:="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
                     TrailingMinusNumbers:=True
    End With
  End With
End Sub

