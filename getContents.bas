Attribute VB_Name = "getContents"
Public strColLetter As String
Public varOffsetAddress As Variant
Public varLastCell As Variant
Public ctlCaption As String
Public ctl As Control
Public Result As Variant

Sub getCount(m_chckBox As MSForms.CheckBox)
    
    Dim rngCheckBox As Range
    Dim varCheckBoxAddress As Variant
    Dim intCheckBoxCol As Long, intOffsetCol As Long
    Dim varCellValues As Variant
    Dim objCellCount As Object
    Dim rngRange1 As Range, rngCell1 As Range
    Dim strTmp As String
    
    
    ReportTool.ListBox1.Clear
    
    ' Gets the cell address of the column header (check box caption)
        
    Set rngCheckBox = Cells.Find(What:=m_chckBox.Caption, LookIn:=xlFormulas, LookAt _
                                :=xlPart, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                                False, SearchFormat:=False)

    ' Stores the check box range address into a variable
    varCheckBoxAddress = rngCheckBox.Address(False, False)
    
    ' Gets the column of the check box range address
    intCheckBoxCol = Range(varCheckBoxAddress).Column

    ' converts column number to column letter
    ' it is needed to find And get The address Of first visible cell of the filtered data
    strColLetter = Col_Letter(intCheckBoxCol)
    
End Sub

Sub getResult(strColLetter As String, varOffsetAddress As Variant, varLastCell As Variant)

    ' Gets the address of the first visibile cell of the filtered data
    With ActiveSheet.AutoFilter.Range
        varOffsetAddress = Range(strColLetter & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Address(False, False)
    End With
        
    ' Gets the column number of the address of the first visibile cell of the filtered data
    intOffsetCol = Range(varOffsetAddress).Column
          
    ' Gets the cell address of the last non-blank cell in a column
     varLastCell = Cells(Rows.Count, intOffsetCol).End(xlUp).Address(False, False)

    
    'Get the unique distinct values plus its count
    Set objCellCount = CreateObject("scripting.dictionary")
    Set rngRange1 = Range(varOffsetAddress, varLastCell)
    
    For Each rngCell1 In rngRange1
        rngcelval = rngCell1.Value
        strTmp = Trim(rngCell1.Value)
        If Len(strTmp) > 0 Then objCellCount(strTmp) = objCellCount(strTmp) + 1
        
    Next rngCell1

    For Each varCellValues In objCellCount.keys
         'ReportTool.ListBox1.AddItem varCellValues & objCellCount(varCellValues)
         With ReportTool.ListBox1
            .AddItem
            .list(i, 0) = varCellValues
            .list(i, 2) = objCellCount(varCellValues)
            i = i + 1
        End With
    Next varCellValues

    ActiveSheet.ShowAllData

End Sub
Sub getCountUnfiltered(m_chckBox As MSForms.CheckBox)

    Dim rngCheckBox As Range
    Dim varCheckBoxAddress As Variant
    Dim varLastCell As Variant
    Dim intCheckBoxCol As Long
    Dim strTmp As String
    Dim rngRange1 As Range
    Dim rngCell1 As Range
    Dim objCellCount As Object
    Dim varCellValues As Variant
    
    
    ' Gets the cell address of the column header
    Set rngCheckBox = Cells.Find(What:=m_chckBox.Caption, LookIn:=xlFormulas, LookAt _
                                :=xlPart, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                                False, SearchFormat:=False)

    ' Stores the check box range address into a variable
    varCheckBoxAddress = rngCheckBox.Offset(1, 0).Address(False, False)
    
    ' Gets the column of the cell address
    intCheckBoxCol = Range(varCheckBoxAddress).Column
          
    ' Gets the cell address of the last non-blank cell in a column
    varLastCell = Cells(Rows.Count, intCheckBoxCol).End(xlUp).Address(False, False)
        
         
    ' Get the unique distinct values plus its count
    Set objCellCount = CreateObject("scripting.dictionary")
    Set rngRange1 = Range(varCheckBoxAddress, varLastCell)
    
    For Each rngCell1 In rngRange1
        strTmp = Trim(rngCell1.Value)
        If Len(strTmp) > 0 Then objCellCount(strTmp) = objCellCount(strTmp) + 1
    Next rngCell1
    
    For Each varCellValues In objCellCount.keys
        MsgBox varCellValues & " " & objCellCount(varCellValues)
    Next varCellValues

End Sub

Sub getCOuntOfTwoColumn(ctlCaption As String)

    Dim wsActive As Worksheet
    Dim varFirstVisibile1 As Variant, varFirstVisibile2 As Variant, varLastCell1 As Variant, varLastCell2 As Variant
    Dim intLastRow1 As Long, intLastRow2 As Long
    Dim i As Long, j As Long, k As Long, l As Long
    Dim rngRange2 As Range, rngRange3 As Range
    Dim strTwoValues As String
    Dim varSpltTwoValues As Variant
    
    Dim varLastCell3
    Dim rngCheckBox1 As Range, rngCheckBox2 As Range
    Dim varCheckBoxAddress1 As Variant, varCheckBoxAddress2 As Variant
    Dim intCheckBoxCol1 As Long, intCheckBoxCol2 As Long
    Dim strColLetter1 As String, strColLetter2 As String
    
    ' Creates a temporary sheet. It is needed to process data to have the combination of the selected checkboxes
    
    Application.Run "Tools.CreateSheet"
    Set wsTemp = ActiveWorkbook.Sheets("Temp")
    
    Result = Split(ctlCaption, "~")
    varFirstCol = Result(1)
    varSecCol = Result(2)
    
    Set rngCheckBox1 = Cells.Find(What:=varFirstCol, LookIn:=xlFormulas, LookAt _
                                :=xlPart, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                                False, SearchFormat:=False)
    Set rngCheckBox2 = Cells.Find(What:=varSecCol, LookIn:=xlFormulas, LookAt _
                                :=xlPart, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                                False, SearchFormat:=False)

    ' Stores the check box range address into a variable
    varCheckBoxAddress1 = rngCheckBox1.Address(False, False)
    varCheckBoxAddress2 = rngCheckBox2.Address(False, False)
    
    ' Gets the column of the check box address
    intCheckBoxCol1 = Range(varCheckBoxAddress1).Column
    intCheckBoxCol2 = Range(varCheckBoxAddress2).Column
    
    ' Gets the column letter of the check box address
    strColLetter1 = Col_Letter(intCheckBoxCol1)
    strColLetter2 = Col_Letter(intCheckBoxCol2)
    
    ' Gets the first visible cell address
    With ActiveSheet.AutoFilter.Range
        varFirstVisible1 = Range(strColLetter1 & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Address(False, False)
        varFirstVisible2 = Range(strColLetter2 & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Address(False, False)
    End With
        

    varLastCell1 = Cells(Rows.Count, intCheckBoxCol1).End(xlUp).Address(False, False)
    varLastCell2 = Cells(Rows.Count, intCheckBoxCol2).End(xlUp).Address(False, False)
    
    ' Copy the selected checkbox's data to the hidden sheet
    ActiveSheet.Range(varFirstVisible1, varLastCell1).SpecialCells(xlCellTypeVisible).Copy
    wsTemp.Cells(1, 1).PasteSpecial
    ActiveSheet.Range(varFirstVisible2, varLastCell2).SpecialCells(xlCellTypeVisible).Copy
    wsTemp.Cells(1, 2).PasteSpecial
    
    intLastRow1 = Sheets("Temp").Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Combines the 2 cell values and paste it to another column
    For i = 1 To intLastRow1
        Sheets("Temp").Cells(i, 3) = Sheets("Temp").Cells(i, 1) & "~" & Sheets("Temp").Cells(i, 2)
        Sheets("Temp").Cells(i, 4) = Sheets("Temp").Cells(i, 3)
    Next i

    ' Removes duplicates
    Sheets("Temp").Range("D:D").RemoveDuplicates Columns:=1, Header:=xlYes
    
    ' Counts the occurrences of cell values
    For Each rngRange2 In Sheets("Temp").Range("D:D").SpecialCells(2).Offset(, 1)
        rngRange2.Formula = "=COUNTIF(C:C," & rngRange2.Offset(, -1).Address & ")"
    Next rngRange2
    
    varLastCell3 = Sheets("Temp").Cells(Rows.Count, 4).End(xlUp).Address(False, False)
    intLastRow2 = Sheets("Temp").Cells(Rows.Count, 4).End(xlUp).Row
    
    Set rngRange3 = Sheets("Temp").Range("D1", varLastCell3)
    
    ' Splits the combined/ concatenated cell values
    For Each cell In rngRange3
    
        strTwoValues = cell.Value
        varSpltTwoValues = Split(strTwoValues, "~")
        j = j + 1
        Sheets("Temp").Cells(j, 6) = varSpltTwoValues(0)
        Sheets("Temp").Cells(j, 7) = varSpltTwoValues(1)
        
     Next cell
     
     ' Displays the value in a listbox
    For k = 1 To intLastRow2
        With ReportTool.ListBox1
            .ColumnWidths = "80;80;25;"
            .AddItem
            .list(l, 0) = Sheets("Temp").Cells(k, 6)
            .list(l, 1) = Sheets("Temp").Cells(k, 7)
            .list(l, 2) = Sheets("Temp").Cells(k, 5)
            l = l + 1
        End With
    Next k
    
    Sheets("Temp").Range("A:A,B:B,C:C,D:D,E:E,F:F,G:G").EntireColumn.Delete
    Application.Run "Tools.DeleteSheet"
    ActiveSheet.ShowAllData
    
    End Sub


