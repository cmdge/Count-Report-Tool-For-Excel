Attribute VB_Name = "filterDates"
Public stDate As Variant
Public enDate As Variant


Sub filterDate(stDate As Variant, enDate As Variant)
    
    Dim rngDateRange As Range
    Dim intDateRangeCol As Long
    Dim intLastCol As Long
    Dim intLastRow As Long
    Dim startMonth As Variant
    Dim startYear As Variant
    Dim endMonth As Variant
    Dim endYear As Variant

    
    
    ' Selecting all cells
    'intLastCol = ActiveSheet.range("A1").End(xlToRight).Column
    'intLastRow = ActiveSheet.Cells(65536, intLastCol).End(xlUp).Row

    ' Gets the cell address of the cell with a word "Date"
    Set rngDateRange = Cells.Find(What:="Date", LookIn:=xlFormulas, LookAt _
                :=xlPart, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False)
                
    intDateRangeCol = rngDateRange.Column
    
    
    startMonth = Month(stDate)
    startYear = Year(stDate)
    endMonth = Month(enDate)
    endYear = Year(enDate)
    

    ldatefrom = DateSerial(startYear, startMonth, 1)
    ldateto = DateSerial(endYear, endMonth + 1, 0)
    
    
    'Filter data between dates
    ActiveSheet.Range("A1", ActiveSheet.Cells).AutoFilter Field:=intDateRangeCol, _
    Criteria1:=">=" & ldatefrom, Operator:=xlAnd, Criteria2:="<=" & ldateto
    
    Debug.Print ThisMonth1 & ThisYear1; ThisMonth2 & ThisYear2


End Sub




