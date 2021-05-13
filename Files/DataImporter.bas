Attribute VB_Name = "DataImporter"
Option Explicit

'--------------------------------------------------------------'
'--- Keywordy w notatniku oddzielaj¹ce poszczególne sekcje. ---'
'--------------------------------------------------------------'
Const nameKeyword As String = "Project name"
Const tagKeyword As String = "Tags"
Const additionalContent As String = "Additional content"

'--------------------------------------------'
'--- Rozszerzenie pliku; inne ni¿ '.txt'. ---'
'--------------------------------------------'
Const FileType As String = ".projData"

Private keywordDictionary As Dictionary

Private DataEntryNumber As Integer

Public Sub CheckForData()
    
    DataEntryNumber = 0
    Call Manager.SetManualCalculation
    Call LoopAllSubFolders(ActiveWorkbook.path)
    Call Manager.SetAutomaticCalculation
    
End Sub

'------------------------------------------------------------------------------------------------------------------------'
'--- Rozszerzyc o pomijanie folderów, w których byly juz robione pliki z informacjami. Pary string (filepath) + bool. ---'
'--- Przemyœleæ sposób uwzglêdniania folderów i podfolderów. Tylko g³ówne pliki? A co z wieloetapowymi projektami? ------'
'------------------------------------------------------------------------------------------------------------------------'
Public Sub LoopAllSubFolders(ByVal folderPath As String)

    Dim FileSystem As Object
    Dim HostFolder As String

    HostFolder = folderPath

    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    DoFolder FileSystem.GetFolder(HostFolder)
    
End Sub

Sub DoFolder(Folder)

    Dim SubFolder
    Dim FileName As String
    
    For Each SubFolder In Folder.SubFolders
    
        DoFolder SubFolder
        
    Next
    
    Dim File
    
    For Each File In Folder.Files
    
                If CheckExtension(File.Name, FileType) = True Then

                    DataEntryNumber = DataEntryNumber + 1
                    Call GetDataFromFile(Folder & "\", File.Name)

'--------------------------------------------------------------------------------------------------------------------'
'--- Dodanie do listy œcie¿ek lokalizacji plików, zeby pozniej zrobic osobna funkcje na zbiorcze usuwanie plikow. ---'
'--- zamiast tego co jest w tej chwili ponizej, wykorzystujacego slaby if. ------------------------------------------'
'--------------------------------------------------------------------------------------------------------------------'
                    If FilesearchForm.TempCheckBox.Value Then Kill (Folder & "\" & File.Name)
                    
                End If
    Next

End Sub

'    Dim FileName As String
'    Dim fullFilePath As String
'    Dim FoldersAmount As Long
'    Dim folders() As String
'    Dim i As Long
'
'    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
'    FileName = Dir(folderPath & "*.*", vbDirectory)
'
'    Do While Len(FileName) <> 0
'
'        If Left(FileName, 1) <> "." Then
'
'            fullFilePath = folderPath & FileName
'            'do debugowania
'
'            On Error GoTo NextStatement
'            On Error Resume Next
'            If (GetAttr(fullFilePath) And vbDirectory) = vbDirectory Then
'
'                If InStr(1, folderPath, "#recycle") Then
'                    GoTo NextStatement
'                End If
'
'                ReDim Preserve folders(0 To FoldersAmount) As String
'                folders(FoldersAmount) = fullFilePath
'                FoldersAmount = FoldersAmount + 1
'
'            Else
''
'                If CheckExtension(FileName, FileType) = True Then
'
'                    DataEntryNumber = DataEntryNumber + 1
'                    Call GetDataFromFile(folderPath, FileName)
'
''------------------------------------------------------------------------------------------------------------------'
''--- Dodanie do listy miejsc gdzie sa te pliki, zeby pozniej zrobic osobna funkcje na zbiorcze usuwanie plikow. ---'
''--- zamiast tego co jest w tej chwili ponizej, wykorzystujacego slaby if. ----------------------------------------'
''------------------------------------------------------------------------------------------------------------------'
'                    If FilesearchForm.TempCheckBox.Value Then Kill (folderPath & FileName)
''
''                End If
'
'            End If
'
'        End If
'
'NextStatement:
'
'        FileName = Dir()
'
'    Loop
'
'    For i = 0 To FoldersAmount - 1
'
'        Call LoopAllSubFolders(folders(i))
'
'    Next i
'
'End Sub

Sub GetDataFromFile(folderPath As String, FileName As String)

    Dim project As ProjectInfo
    Set project = New ProjectInfo
    Dim textLine As String
    Dim currentKeyword As String
    Debug.Print folderPath & FileName
    Open folderPath & FileName For Input As #1
    
    Do Until EOF(1)
    
        Line Input #1, textLine
        
        If GetKeywordFromFile(textLine) Then
        
            currentKeyword = textLine
            GoTo EndOfLoop
            
        End If
        
        If textLine = vbNullString Then GoTo EndOfLoop
    
        Select Case currentKeyword
            
            Case nameKeyword
                project.ProjectName = textLine
            
            Case tagKeyword
                project.AddTag (textLine)
                
            Case additionalContent
                project.AddAdditionalInfo (textLine)
                
        End Select
        
EndOfLoop:

    Loop
        
    Close #1
    
    project.path = folderPath
    
    Call SaveProjectInDatabase(project)
        
End Sub

Private Function CheckExtension(FileName As String, desiredFileType As String) As Boolean

    If Right$(FileName, Len(desiredFileType)) = desiredFileType Then
        
        CheckExtension = True
        Exit Function

    End If
    
    CheckExtension = False
    
End Function

'-------------------------------------------------------------'
'--- Poprawiæ funkcjê na dodawanie do istniej¹cych danych. ---'
'-------------------------------------------------------------'
Sub SaveProjectInDatabase(project As ProjectInfo)
    
    If CheckForExistingData = True Then
    
        If FilesearchForm.IsInDatabase(project) Then
            Exit Sub
        End If
        
    End If
        
    Cells(DataEntryNumber, 1) = project.path
    Cells(DataEntryNumber, 2) = project.ProjectName
    
    Dim i As Integer
    
    For i = 0 To project.TagsAmount - 1
    
        Cells(DataEntryNumber, 3 + i) = project.Tag(i)
        
    Next
    
End Sub

Public Sub CreateFile(project As ProjectInfo)

    Dim fileSysObj As Object
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    
    Dim obj_DataFile As Object
    Set obj_DataFile = fileSysObj.CreateTextFile(project.path & project.ProjectName & FileType)
    
    obj_DataFile.WriteLine nameKeyword
    obj_DataFile.WriteLine project.ProjectName
    obj_DataFile.WriteLine
    obj_DataFile.WriteLine tagKeyword
    
    Dim i As Integer
    For i = 0 To project.TagsAmount - 1
        obj_DataFile.WriteLine project.Tag(i)
    Next
    
    obj_DataFile.WriteLine
    obj_DataFile.WriteLine additionalContent
    
    obj_DataFile.Close

End Sub

Public Function CheckForExistingData() As Boolean
    
    CheckForExistingData = False
    Dim endOfDatabase As Boolean: endOfDatabase = False
    Dim i As Integer: i = 1
    
    Do Until endOfDatabase = True
    
        If Cells(i, 1) = vbNullString Then
            endOfDatabase = True
        Else
            CheckForExistingData = True
            i = i + 1
        End If
        
    Loop
    
    DataEntryNumber = i

End Function


Public Function GetKeywordFromFile(currentLine As String) As Boolean
    
    On Error GoTo e_CreateKeywordDictionary
    
e_GetKeywordFromFile:

    If keywordDictionary.Exists(currentLine) Then
        GetKeywordFromFile = True
    Else
        GetKeywordFromFile = False
    End If
    
    Exit Function
    
e_CreateKeywordDictionary:

    If keywordDictionary Is Nothing Then Call CreateKeywordDictionary
    
    GoTo e_GetKeywordFromFile
        
End Function

'----------------------------------------------------------------------'
'--- Tu maj¹ zostaæ dodane pozosta³e keywordy które bêd¹ w plikach. ---'
'----------------------------------------------------------------------'
Private Sub CreateKeywordDictionary()

    Set keywordDictionary = New Dictionary
    
    With keywordDictionary
        .Add nameKeyword, nameKeyword
        .Add tagKeyword, tagKeyword
        .Add additionalContent, additionalContent
    End With

End Sub
