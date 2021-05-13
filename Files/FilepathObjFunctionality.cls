VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilepathObjFunctionality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ObsoleteCallStatement
Option Explicit

Implements IFunctionality

Public WithEvents labelEvents As MSForms.Label
Attribute labelEvents.VB_VarHelpID = -1

Private Sub labelEvents_Click()

    Dim str As String
    str = Replace(labelEvents.Caption, "Testing.projData", vbNullString)
    Call Shell("explorer.exe" & " " & str, vbNormalFocus)

End Sub

Private Sub labelEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    
    If Button = 2 Then
        'Shell labelEvents.caption, vbNormalNoFocus
    End If

End Sub
