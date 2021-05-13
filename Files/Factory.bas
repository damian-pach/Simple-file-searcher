Attribute VB_Name = "Factory"
Option Explicit

Public Function Vector2(x As Integer, y As Integer) As Vector2

    Dim v2 As Vector2
    Set v2 = New Vector2

    v2.SetProperties x, y
    
    Set Vector2 = v2

End Function

Public Function ProjectInfo(Name As String, path As String, Tags() As String) As ProjectInfo

    Dim i As Integer
    Dim project As ProjectInfo
    Set project = New ProjectInfo
    
    project.ProjectName = Name
    project.path = path
    
    For i = 0 To Manager.ArraySize(Tags)
        project.AddTag (Tags(i))
    Next
    
    Set ProjectInfo = project

End Function

