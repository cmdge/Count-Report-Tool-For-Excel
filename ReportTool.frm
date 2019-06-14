VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportTool 
   Caption         =   "Report"
   ClientHeight    =   6600
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16968
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

Private Sub ListBox1_Click()
End Sub

'Create check box dynamically
Private Sub UserForm_Initialize()

    Dim curRow      As Long
    Dim LastColumn  As Long
    Dim i           As Long
    Dim chkbox      As MSForms.checkbox
    Dim d As Variant
    
    curRow = 1 'Set row index

    LastColumn = Worksheets("Sheet1").Cells(curRow, Columns.count).End(xlToLeft).Column 'Find the last non-blank cell in row 1
    

    'Displays the check box with its appropriate check box caption
    For i = 1 To LastColumn
        Set chkbox = Me.Controls.Add("Forms.CheckBox.1", "CheckBox_" & i)
        chkbox.Caption = Worksheets("Sheet1").Cells(curRow, i).Value
        
        'Sets the position of check box
        chkbox.Left = 25
        chkbox.Top = 46 + ((i - 1) * 20)
    Next i
    
    'Debug.Print

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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End Sub
Private Sub ComboBox1_Change()
'Unload ReportTool
End Sub
Private Sub ComboBox1_AfterUpdate()
    stDate = ComboBox1.Value

End Sub
Private Sub ComboBox2_Change()

End Sub
Private Sub ComboBox2_AfterUpdate()
    enDate = ComboBox2.Value
End Sub



Private Sub CommandButton1_Click()

    Call filterDate(stDate, enDate)
    Call getResult(strColLetter, varOffsetAddress, varLastCell)

    
End Sub
