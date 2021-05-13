VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TagLabelFunctionality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFunctionality

Public WithEvents labelEvents As MSForms.Label
Attribute labelEvents.VB_VarHelpID = -1

Private Sub labelEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    
    If Button = 2 Then
        'labelEvents_Terminate
    End If
    
End Sub

Private Sub labelEvents_Terminate()

    FilesearchForm.TagsEntered.Remove FilesearchForm.TagsEntered.Item(labelEvents.Caption)
    FilesearchForm.FilterProjectsDatabase
    labelEvents.Left = 10000
    Set labelEvents = Nothing
    

End Sub
