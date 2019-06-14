Attribute VB_Name = "Module1"
Sub FirstVisibleCell()


Dim cbRange As Range
    'Dim cBoxAddress As Variant
    'Dim lastCellValue As Variant
    'Dim tmp As String
    Dim cBox As Long
    Dim offseet As Variant
    
    ' Gets the cell address of the column header
        
    Set cbRange = Cells.Find(What:="Column1", LookIn:=xlFormulas, LookAt _
                :=xlPart, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)

    ' Stores the check box range address into a variable
    cBoxAddress = cbRange.Offset(1, 0).Address(False, False)
    ' Gets the column of the cell address
    
    cBox = Range(cBoxAddress).Column
    Debug.Print cBox
    k = Col_Letter(cBox)
    Debug.Print k
    With Worksheets("Sheet1").AutoFilter.Range
        offseet = Range(k & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Address(False, False)
    End With
    Debug.Print offseet
End Sub



Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
