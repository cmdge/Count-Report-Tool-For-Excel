Attribute VB_Name = "Dates"
Public arrDate As Variant
    

Function FnMonthName()

    Dim strDate

   Dim strResult

    strDate = "15-07-2013"

   strResult = "Full Month Name of the " & strDate & " is -> " & MonthName(Month(strDate)) & vbCrLf

   strResult = strResult & "Abbriviated Month Name of the " & strDate & " is -> " & MonthName(Month(strDate), True)

   MsgBox strResult

End Function


Sub GetDates()

    Dim dRange As Range
    Dim dRangeAddress As Variant
    Dim lastCellValue As Variant
    Dim r As Range
    Dim c As Range
    Dim dateVal As Variant
    Dim cellValues As Variant
    
    ' Gets the cell address of the word "Date"
    Set dRange = Cells.Find(What:="Date", LookIn:=xlFormulas, LookAt _
                :=xlPart, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)
    
    
    dRangeAddress = dRange.Offset(1, 0).Address(False, False)
    dr = Range(dRangeAddress).Column

    lastCellValue = Cells(Rows.count, dr).End(xlUp).Address(False, False)
    
    
    
    ' Get the unique distinct values plus its count
    Set r = Range(dRangeAddress, lastCellValue)
    Set cellCount = CreateObject("scripting.dictionary")
    
    For Each c In r
       
        dateVal = c.Value
        Debug.Print dateVal
        mnthName = Format(dateVal, "mmmm-yy")
        
        tmp = Trim(mnthName)
       If Len(tmp) > 0 Then cellCount(tmp) = cellCount(tmp) + 1
      'Debug.Print "tmp: " & tmp
    Next c
    
    For Each cellValues In cellCount.keys
        Debug.Print cellValues
        
        ' Add to Combo Box
       ' ReportTool.ComboBox1.AddItem cellValues
    Next cellValues
    
    
    
    
   ' With Report
           ' With .ComboBox1
              '  .Clear
               ' .List = arrDate(cellValues)
               ' .Style = fmStyleDropDownList
           ' End With
           ' .Show
       ' End With
'Unload Report
  


        
        
       
        

    
    
    
    
End Sub

  
