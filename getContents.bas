Attribute VB_Name = "getContents"
Public strColLetter As String
Public varOffsetAddress As Variant
Public varLastCell As Variant



Sub getCount(m_chckBox As MSForms.checkbox)
    
    Dim rngCheckBox As Range
    Dim varCheckBoxAddress As Variant
    Dim intCheckBoxCol As Long
    'Dim varLastCell As Variant
    'Dim strColLetter As String
    'Dim varOffsetAddress As Variant
    Dim intOffsetCol As Long
    Dim varCellValues As Variant
    Dim objCellCount As Object
    Dim rngRange1 As Range
    Dim rngCell1 As Range
    Dim strTmp As String
    
    
    
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
    
    ' gets the address of the first visibile cell of the filtered data
    'With ActiveSheet.AutoFilter.range
        'varOffsetAddress = range(strColLetter & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Address(False, False)
    'End With
        
    ' gets the column number of the address of the first visibile cell of the filtered data
    'intOffsetCol = range(varOffsetAddress).Column
          
    ' Gets the cell address of the last non-blank cell in a column
     'varLastCell = Cells(Rows.count, intOffsetCol).End(xlUp).Address(False, False)

    'Get the unique distinct values plus its count
    'Set objCellCount = CreateObject("scripting.dictionary")
    'Set rngRange1 = range(varOffsetAddress, varLastCell)
    
    'For Each rngCell1 In rngRange1
        'strTmp = Trim(rngCell1.Value)
        'If Len(strTmp) > 0 Then objCellCount(strTmp) = objCellCount(strTmp) + 1
    'Next rngCell1
         
    'For Each varCellValues In objCellCount.keys
        'Debug.Print varCellValues & " " & objCellCount(varCellValues)
    'Next varCellValues
    
    'addToCollection
    

End Sub

Sub getResult(strColLetter As String, varOffsetAddress As Variant, varLastCell As Variant)

' gets the address of the first visibile cell of the filtered data
    With ActiveSheet.AutoFilter.Range
        varOffsetAddress = Range(strColLetter & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Address(False, False)
    End With
        
    ' gets the column number of the address of the first visibile cell of the filtered data
    intOffsetCol = Range(varOffsetAddress).Column
          
    ' Gets the cell address of the last non-blank cell in a column
     varLastCell = Cells(Rows.count, intOffsetCol).End(xlUp).Address(False, False)


'Get the unique distinct values plus its count
    Set objCellCount = CreateObject("scripting.dictionary")
    Set rngRange1 = Range(varOffsetAddress, varLastCell)
    
    For Each rngCell1 In rngRange1
        strTmp = Trim(rngCell1.Value)
        If Len(strTmp) > 0 Then objCellCount(strTmp) = objCellCount(strTmp) + 1
        
    Next rngCell1

    Debug.Print "rngCell1: " & rngCell1
    Debug.Print "strTmp: " & strTmp
    Debug.Print "objCellCount(strTmp): " & objCellCount(strTmp)

    For Each varCellValues In objCellCount.keys
         'ReportTool.ListBox1.AddItem varCellValues & objCellCount(varCellValues)
         With ReportTool.ListBox1
            .AddItem
            .List(i, 0) = varCellValues
            .List(i, 1) = objCellCount(varCellValues)
            i = i + 1
        End With
    Next varCellValues

    ActiveSheet.ShowAllData

End Sub
Sub getCountUnfiltered(m_chckBox As MSForms.checkbox)

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
    varLastCell = Cells(Rows.count, intCheckBoxCol).End(xlUp).Address(False, False)
        
         
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




