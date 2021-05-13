VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FilesearchForm 
   Caption         =   "File searcher"
   ClientHeight    =   5196
   ClientLeft      =   0
   ClientTop       =   216
   ClientWidth     =   9552.001
   OleObjectBlob   =   "FilesearchForm.frx":0000
End
Attribute VB_Name = "FilesearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@IgnoreModule ImplicitActiveSheetReference, ObsoleteCallStatement
Option Explicit

'--------------'
'--- To do: ---'
'------------------------------------------------------'
'--- Wartoœci liczbowe w voidach zast¹piæ zmiennymi ---'
'------------------------------------------------------------------------------'
'--- Przerobiæ wszystkie arraye na kolekcje, s³abo siê operuje na arrayach. ---'
'------------------------------------------------------------------------------'

'---------------------------------------------------------------------------------------------------------------------------------'
'--- No idea czy dobre, zwykle jest sugerowane ¿eby nie robiæ Dim as New dla customowych klas ------------------------------------'
'--- Nie tykaæ póki nie ogarnê niewysypuj¹cej siê alternatywy --------------------------------------------------------------------'
'--- problem jest z podawaniem arraya customowej klasy (w domysle z zaimplementowanym interfejsem) jako argumentu dla funkcji, ---'
'--- wiêc póki nie rozwi¹¿e siê tego problemu musi zostaæ tak jak jest -----------------------------------------------------------'
'--- 20.11.20 --------------------------------------------------------------------------------------------------------------------'
'--- Problem chyba do rozwi¹zania, problem z podawaniem arraya wynika³ z wykorzystania lub nie keywordu Call przed wywo³aniem ----'
'--- Do póŸniejszej analizy jeœli bêdzie konieczna. ------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------------------------------'
Private TagLabelFunctionality() As New TagLabelFunctionality
Private FilepathFunctionality() As New FilepathObjFunctionality

'---------------------------------------------------------------------------'
'--- Wprowadzone poprzez input tagi, wyœwietlane w oknie poni¿ej inputu. ---'
'---------------------------------------------------------------------------'
Public TagsEntered As Collection

'-----------------------------------------------'
'--- Zapewnia, ¿e wyœwietla poprawnie tekst. ---'
'-----------------------------------------------'
Private strModifier As StringModifier

'-----------------------------------------------------------'
'--- Do manipulacji projektami znajduj¹cymi siê w bazie. ---'
'-----------------------------------------------------------'
Private DatabaseProjects() As ProjectInfo

'----------------------------------------------------------------------'
'--- Zestaw danych po filtracji, wyœwietlany w oknie ProjectsShown. ---'
'----------------------------------------------------------------------'
Private ProjectsToDisplay() As ProjectInfo

'-----------------------------------------------'
'--- Do zape³niania listy w dropdown liœcie. ---'
'-----------------------------------------------'
Private TagsInProjects As Collection

'-----------------------------------------------------------'
'--- Do filtrowania listy wszystkich istniej¹cych tagów ----'
'--- Gdy user wpisze czêœæ taga. ---------------------------'
'-----------------------------------------------------------'
Private FiteredTagsInProjects As Collection

Private Sub UserForm_Initialize()
    
    Application.Visible = False
    
    Set TagsEntered = New Collection
    Set TagsInProjects = New Collection
    Set FiteredTagsInProjects = New Collection
    
    Set strModifier = New StringModifier
    
    On Error Resume Next
    Call GetDataFromDatabase
    Call DisplayData(DatabaseProjects)
    Set FiteredTagsInProjects = TagsInProjects

End Sub

Private Sub UserForm_Terminate()
    
    Application.Visible = True
    ActiveWorkbook.Save
    Application.Quit
    
End Sub

Private Sub DataExtractionButton_Click()

    Dim i As Integer
    For i = 0 To Manager.ArraySize(DatabaseProjects)
        DataImporter.CreateFile DatabaseProjects(i)
    Next

End Sub

Private Sub DiscScanButton_Click()
    
    Dim answer As VbMsgBoxResult
    
    answer = MsgBox("Delete imported files from disc?", vbYesNo)
    If answer = vbYes Then
        TempCheckBox.Value = True
    Else
        TempCheckBox.Value = False
    End If
    
    Call DataImporter.CheckForData
    Call GetDataFromDatabase
    Call DisplayData(DatabaseProjects)

End Sub

Private Sub TagsListBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Dim textInBox As String
    textInBox = TagsListBox.Text
    
    If KeyCode = 13 Then
        
        '------Empty textbox check------'
        If textInBox = vbNullString Then Exit Sub
        
        '------Existsing tag check------'
        If Not Manager.IsInCollection(TagsEntered, textInBox) Then
            
            Dim labelInit As Control
            Set labelInit = AddLabel(textInBox)
            
            Dim tagPosition As Vector2
            Set tagPosition = LabelPositionInFrame(TagsWindow, labelInit, TagsEntered)
            
            If tagPosition.y > TagsWindow.Top + TagsWindow.Height Then
                'Call AddExpansionButton   <--- do implementacji
            Else
                AddNewLabel textInBox, tagPosition, TagsEntered, "TagLabel", "TagLabelFunctionality"
            End If
            
            Set labelInit = Nothing

            Call FilterProjectsDatabase
                                              
            '---Do implementacji sortowanie bazy projektów po wstawieniu tagu.---'
            
            TagsListBox.Text = vbNullString
            
        Else

            TagsListBox.Text = vbNullString
            MsgBox strModifier.KeepPolishLetters("Tag already has been added")
            
        End If
        
        KeyCode = 0
                       
    End If
    
    Call TagsListBox.DropDown
    
'---------------------------------------------------------'
'--- Filtrowanie listy dostêpnych tagów, do zrobienia. ---'
'---------------------------------------------------------'

    '    'If KeyCode >= 65 And KeyCode <= 90 Then
'
'        Set FiteredTagsInProjects = New Collection
'
'        textInBox = textInBox & ChrW(KeyCode)
'        Debug.Print KeyCode
'        Debug.Print textInBox
'
'        Dim i As Integer
'        For i = 1 To TagsInProjects.Count
'
'            If UCase(Left(TagsInProjects(i), Len(textInBox))) = UCase(textInBox) Then
'
'                FiteredTagsInProjects.Add TagsInProjects(i)
'
'            End If
'
'        Next
'
'        For i = TagsListBox.ListCount - 1 To 0 Step -1
'
'            TagsListBox.RemoveItem (i)
'
'        Next
'
'        For i = 1 To FiteredTagsInProjects.Count
'
'            TagsListBox.AddItem (FiteredTagsInProjects(i))
'
'        Next
'
'        TagsListBox.ListRows = TagsListBox.ListCount + 1
'
'
'    'End If

End Sub

Private Sub AddNewLabel(captionText As String, position As Vector2, labelCollection As Collection, collectionName As String, Optional functionality As String)

    If functionality = vbNullString Then functionality = "Default"
    
    Dim Label As Control
    Set Label = Controls.Add("Forms.Label.1", collectionName & labelCollection.Count, True)
    
    With Label
        .Caption = captionText
        .BackColor = vbWhite
        .Width = 500
        .AutoSize = True
        .ZOrder fmZOrderFront
        .Left = position.x
        .Top = position.y
    End With
    
    labelCollection.Add Label, Label.Caption
    
    Select Case functionality
        
        Case "TagLabelFunctionality"
        
            ReDim Preserve TagLabelFunctionality(1 To labelCollection.Count)
            Set TagLabelFunctionality(labelCollection.Count).labelEvents = Label
            
        Case "FilepathFunctionality"
            
            ReDim Preserve FilepathFunctionality(1 To labelCollection.Count)
            Set FilepathFunctionality(labelCollection.Count).labelEvents = Label
    
    End Select

End Sub


Private Sub GetDataFromDatabase()

    Dim i As Integer: i = 1
    Dim k As Integer: k = 0
    
    On Error Resume Next
    Do While Cells(i, 1) <> vbNullString
    
        Dim projectTags() As String
        k = 0
        
    '--- Przy implementacji extra informacji zamiast vbNullString mo¿na umieœciæ tekst ---'
    '--- sygnalizuj¹cy zmianê typu danych, np. ProjectType, Wykonawcy, EndOfFile. --------'
    '--- Wymaga zmiany DataImportera. ----------------------------------------------------'
        Do Until Cells(i, k + 3) = vbNullString
            ReDim Preserve projectTags(0 To k)
            projectTags(k) = (Cells(i, k + 3))
            Dim bl As Boolean: bl = Manager.IsInCollection(TagsInProjects, projectTags(k))
            If bl = False Then
                TagsInProjects.Add projectTags(k), projectTags(k)
            End If
            k = k + 1
        Loop
        
        ReDim Preserve DatabaseProjects(0 To i - 1)
        Set DatabaseProjects(i - 1) = Factory.ProjectInfo(Cells(i, 2), Cells(i, 1), projectTags)
        
        i = i + 1
        
    Loop
    
    Call SortCollection(TagsInProjects)
    
    For i = 1 To TagsInProjects.Count
    
        TagsListBox.AddItem (TagsInProjects(i))
        
    Next
    
End Sub

'-------------------------------------------------------------------------------------------------'
'--- Zast¹piæ wyœwietlanie w 3-ciej kolumnie Listboksu wyœwietlaniem w customowym inspektorze, ---'
'--- gdzie bêdzie mo¿na wyœwietlaæ wszystkie tagi. -----------------------------------------------'
'-------------------------------------------------------------------------------------------------'
Public Sub DisplayData(data() As ProjectInfo)

    Dim i As Integer
    Dim tempTagString As String
    Dim j As Integer
    Dim dataLength As Integer: dataLength = Manager.ArraySize(data)
    
    Dim projectsCount As Integer
    For projectsCount = ProjectsShown.ListCount - 1 To 0 Step -1
        ProjectsShown.RemoveItem (projectsCount)
    Next
    
    If data(0) Is Nothing Then Exit Sub
    
    For i = 0 To dataLength
    
        With ProjectsShown
        
            .AddItem
            .List(i, 0) = data(i).ProjectName
            .List(i, 1) = data(i).path
            
            '---tymczasowo dodane, powinno pokazywaæ tagi w postaci labeli---'
            tempTagString = vbNullString
            
            For j = 0 To data(i).TagsAmount - 1
            
                tempTagString = tempTagString & " " & data(i).Tag(j)
                
            Next
    
            .List(i, 2) = tempTagString
            
        End With
        
    'dodanie labeli do projektu zamiast wyœwietlania jednego stringu
    'docelowo ma byæ póŸniej repozycjonowywanie labeli w zale¿noœci od wprowadzonych tagów
    'w pierwszej kolejnoœci pasuj¹ce do tych wprowadzonych przez Usera, wyt³uszczone, potem reszta

'    For j = 0 To data(i).TagsAmount - 1
'        AddNewLabel(data(i).Tag(j),LabelPositionInFrame(customFrame,customLabel,customCollection),customCollection, "customCollection","RBMFunctionality")
'    Next
'
    Next

    'to te¿ zast¹piæ zewnêtrznym stringiem
    ProjectsShown.ColumnWidths = "4.5 cm;1.5 cm;6 cm"
    
End Sub

'-----------------------------------------------------------------------------------------'
'---Zamiast podawania ca³ej kolekcji mo¿na podaæ ostatni element (label) tej kolekcji. ---'
'-----------------------------------------------------------------------------------------'
Private Function LabelPositionInFrame(frame As Control, Label As Control, labelCollection As Collection) As Vector2

    Dim expectedX As Integer
    Dim expectedY As Integer
    
    
    On Error GoTo FramePos
    Dim lastLabel As Control
    Set lastLabel = labelCollection.Item(labelCollection.Count)
    expectedX = lastLabel.Left + lastLabel.Width + 10
    expectedY = lastLabel.Top
    
    If expectedX + Label.Width + 10 > frame.Left + frame.Width Then
        Set LabelPositionInFrame = Factory.Vector2(frame.Left + 10, lastLabel.Top + lastLabel.Height + 10)
    Else
        Set LabelPositionInFrame = Factory.Vector2(expectedX, lastLabel.Top)
    End If
    
    Exit Function
    
    If False Then
        
FramePos:
        Set LabelPositionInFrame = Factory.Vector2(frame.Left + 10, frame.Top + 10)
        Exit Function
        
    End If

End Function

'---Koniecznie dodaæ funkcjonalnoœæ RemoveLabel---'
Private Function AddLabel(captionText As String) As Control

    Dim tempLabel As Control
    Set tempLabel = Controls.Add("Forms.Label.1", "tempLabel", True)
    tempLabel.Caption = captionText
    tempLabel.Width = 500
    tempLabel.AutoSize = True
    tempLabel.Tag = 5
    
    '------------------------------------------------------'
    '---trzeba to zast¹piæ usuwaniem bo s³aby workaround---'
    '------------------------------------------------------'
    tempLabel.Left = 100000
    
    Set AddLabel = tempLabel
    
    'Call RemoveLabel("tempLabel")
    
End Function

'------------------------------------------------------------------------'
'--- Funkcja powoduje b³¹d (od strony Windowsa?). -----------------------'
'--- Excel siê wysypuje przy próbie usuniêcia labelu za drugim razem. ---'
'------------------------------------------------------------------------'
'Public Sub RemoveLabel(labelName As String)
'
'
'    Dim fra As MSForms.UserForm
'    Dim lbl As MSForms.label
'    Dim i As Long
'
'    Set fra = Me
'
'    For i = fra.Controls.Count - 1 To 0 Step -1
'
'        On Error Resume Next
'        Set lbl = fra.Controls(i)
'        'On Error GoTo 0
'
'        If Not lbl Is Nothing Then
'            If lbl.Name = labelName Then
'                fra.Controls.Remove i
'                Set lbl = Nothing
'            End If
'        End If
'
'    Next i
'
'
'End Sub



'-------------'
'---Dzia³a!---'
'-------------'
Public Sub FilterProjectsDatabase()
    
    If TagsEntered.Count = 0 Then

        ReDim ProjectsToDisplay(0 To UBound(DatabaseProjects))
        ProjectsToDisplay = DatabaseProjects
        Call DisplayData(ProjectsToDisplay)
        Exit Sub

    End If
    
    Dim i As Integer: i = 0
    Dim projCounter As Integer
    Dim tagCounter As Integer
    ReDim ProjectsToDisplay(0 To 0)
    
    If AnyTagsRequiredButton.Value Then
    
        Dim projectHasTag As Boolean

        For projCounter = 0 To UBound(DatabaseProjects)

            projectHasTag = False

            For tagCounter = 1 To TagsEntered.Count

                If DatabaseProjects(projCounter).HasTag(TagsEntered(tagCounter)) And Not projectHasTag Then

                    ReDim Preserve ProjectsToDisplay(0 To i)
                    Set ProjectsToDisplay(i) = DatabaseProjects(projCounter)
                    i = i + 1
                    projectHasTag = True

                End If

            Next

        Next

    Else
    
        Dim amountOfTagsInProject As Integer

        For projCounter = 0 To UBound(DatabaseProjects)

            amountOfTagsInProject = 0
            
            For tagCounter = 1 To TagsEntered.Count

                If DatabaseProjects(projCounter).HasTag(TagsEntered(tagCounter)) Then

                    amountOfTagsInProject = amountOfTagsInProject + 1
                    
                    If amountOfTagsInProject = TagsEntered.Count Then
                    
                        ReDim Preserve ProjectsToDisplay(0 To i)
                        Set ProjectsToDisplay(i) = DatabaseProjects(projCounter)
                        i = i + 1
                    
                    End If

                End If

            Next

        Next

    End If

    Call DisplayData(ProjectsToDisplay)

End Sub

'----------------------------------------------------------'
'--- Okropne, do ogarniêcia czemu siê nie da w normalny ---'
'--- sposób usuwaæ instancji klas. ------------------------'
'----------------------------------------------------------'
Private Sub DeleteLastEntryButton_Click()

    On Error Resume Next
    TagsEntered(TagsEntered.Count).Left = 100000
    TagsEntered.Remove (TagsEntered.Count)
    Call DisplayData(DatabaseProjects)
    Call FilterProjectsDatabase

End Sub

Private Sub ProjectsShown_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = 13 Then
        
        Dim i As Integer
        
        For i = 0 To ProjectsShown.ListCount - 1
        
            If (ProjectsShown.Selected(i)) Then
                
                Dim str As String
                str = ProjectsShown.List(i, 1)
                Call Shell("explorer.exe " & str, vbNormalFocus)
                Exit Sub
            
            End If
            
        Next
        
    End If

End Sub

Private Sub ProjectsShown_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
    Dim i As Integer
    
    For i = 0 To ProjectsShown.ListCount - 1
    
        If (ProjectsShown.Selected(i)) Then
            
            Dim str As String
            str = ProjectsShown.List(i, 1)
            Call Shell("explorer.exe " & str, vbNormalFocus)
            Exit Sub
        
        End If
        
    Next
        
End Sub

'znaki specjalne zawsze na koñcu, za du¿o roboty z tym by by³o
Sub SortCollection(collectionToSort As Collection)

    Dim i As Long
    Dim j As Long
    Dim tempItem As Variant

    For i = 1 To collectionToSort.Count - 1
    
        For j = i + 1 To collectionToSort.Count
        
            If collectionToSort(i) > collectionToSort(j) Then

                tempItem = collectionToSort(j)

                collectionToSort.Remove j
                collectionToSort.Add tempItem, tempItem, i
                
            End If
            
        Next j
        
    Next i

End Sub

Private Sub AllTagsRequiredButton_Click()
    
    Call FilterProjectsDatabase
    Call DisplayData(ProjectsToDisplay)
    
End Sub

Private Sub AnyTagsRequiredButton_Click()

    Call FilterProjectsDatabase
    Call DisplayData(ProjectsToDisplay)

End Sub

Public Function IsInDatabase(project As ProjectInfo) As Boolean

    Dim i As Integer
    IsInDatabase = False
    
    For i = 0 To UBound(DatabaseProjects)
    
        If project.ProjectName = DatabaseProjects(i).ProjectName Then
        
            IsInDatabase = True
            Exit Function
            
        End If
        
    Next
    
End Function

