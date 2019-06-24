VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportTool 
   Caption         =   "Report"
   ClientHeight    =   6108
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11064
   OleObjectBlob   =   "ReportTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private eventHandlerCollection As New Collection



'Create check box dynamically
Private Sub UserForm_Initialize()

    Dim intcurRow As Long, intLastColumn  As Long
    Dim i As Long
    Dim chkbox As MSForms.CheckBox
    
    intcurRow = 1 'Set row index

    intLastColumn = Worksheets("Sheet1").Cells(intcurRow, Columns.Count).End(xlToLeft).Column 'Find the last non-blank cell in row 1
    

    'Displays the check box with its appropriate check box caption
    For i = 1 To intLastColumn
        Set chkbox = Me.Controls.Add("Forms.CheckBox.1", "CheckBox" & i)
        chkbox.Caption = Worksheets("Sheet1").Cells(intcurRow, i).Value
        
        'Sets the position of check box
        chkbox.Left = 25
        chkbox.Top = 46 + ((i - 1) * 20)
    Next i

    'Individual click events to dynamic checkboxes on userform

    Dim chckBoxEventHandler As CheckBoxEventHandler
    Dim c As Control
    
    For Each c In ReportTool.Controls
        If TypeName(c) = "CheckBox" Then
            'Create event handler instance
            Set chckBoxEventHandler = New CheckBoxEventHandler
            'Assign it reference to a checkbox
            chckBoxEventHandler.AssignCheckBox c
            'Store the event handler in the userform's collection,
            eventHandlerCollection.Add chckBoxEventHandler
        End If
    Next
    
''''''''''''''''''''''''''''''START AND END DATE COMBO BOX'''''''''''''''''''''''''''
   
    Dim MyDate As Date
    Dim l As Integer
    MyDate = Date
    For l = 1 To 100
        Me.ComboBox1.AddItem Format(MyDate, "mmm yy")
        Me.ComboBox2.AddItem Format(MyDate, "mmm yy")
        MyDate = DateAdd("m", 1, MyDate)
    
    Next
    
End Sub

Private Sub ComboBox1_AfterUpdate()

    stDate = ComboBox1.Value
    
End Sub

Private Sub ComboBox2_AfterUpdate()
    enDate = ComboBox2.Value
End Sub

Private Sub CommandButton1_Click()
    
    
   
    Dim j As Long
    ctlCaption = ""
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is MSForms.CheckBox Then
            If Me.Controls(ctl.Name).Value = True Then
                ctlCaption = ctlCaption & "~" & ctl.Caption
                j = j + 1
            End If
        End If
    Next
    ctlCaption = ctlCaption
    
    If ComboBox1 = vbNullString Or ComboBox2 = vbNullString Or j = 0 Then
        MsgBox "All Fields are Required!"
        ComboBox1.SetFocus
    Else


        'If ticked checkbox is = 1
        If j = 1 Then
            Call filterDate(stDate, enDate)
            ReportTool.ListBox1.Clear
            Call getResult(strColLetter, varOffsetAddress, varLastCell)
            
        End If

        'If ticked checkbox is = 2
        If j = 2 Then
            Call filterDate(stDate, enDate)
            ReportTool.ListBox1.Clear
            Call getCOuntOfTwoColumn(ctlCaption)
        End If
        
        'If ticked checkbox is > 2
        If j > 2 Then
            MsgBox "Number of check box exceed!"
        End If
        
    End If
        
End Sub
