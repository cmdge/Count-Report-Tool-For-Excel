VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckBoxEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This will store a reference to a checkbox and enable handling its events.
Public WithEvents m_chckBox As MSForms.checkbox
Attribute m_chckBox.VB_VarHelpID = -1

' Method to assign a reference to a checkbox to your event handler
Public Sub AssignCheckBox(c As MSForms.checkbox)
    Set m_chckBox = c
    
End Sub

' Private sub to execute something on the event
Private Sub m_chckBox_Click()



    If m_chckBox = True Then
        
        
        'If ActiveSheet.FilterMode = True Then
            Call getCount(m_chckBox)
        'Else
            'Call getCountUnfiltered(m_chckBox)
        'End If
        
        
        'Call count
    
    Else
    
    '    Call removeFromCollection(m_chckBox)
        
    End If
    
     
End Sub



