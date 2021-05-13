VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ObsoleteCallStatement
Option Explicit

'---Nazwa projektu, w którym ten plik siê znajduje---'
Private m_projectName As String

'---Nazwa pliku i pozosta³e dane, który ma zostaæ otwarty przy doubleclicku---'
Private m_targetFileName As String
Private m_Tags() As String
Private m_AdditionalInfo() As String
'---rodzaj pliku, ktory ma zostac otwarty. Docelowo do filtracji danych po typie pliku---'
Private m_fileType As String

'---Œcie¿ka do folderu, w którym znajduje siê plik---'
Private m_path As String

'od 19.11.20
'---nazwa pliku tekstowego, w ktorym sa zawarte informacje o raporcie/projekcie/pliku---'
'---Œcie¿ka docelowo taka sama jak w przypadku pliku, który ma zostaæ otwarty---'
Private m_dataFileName As String

'-----------------------------------------------'
'---Rodzaj pliku - excel, raport, oferta etc.---'
'-----------------------------------------------'
Public Property Let FileType(v_fileType As String)

    m_fileType = v_fileType

End Property

Public Property Get FileType() As String

    FileType = m_fileType

End Property


'---------------------------------------------------------------'
'---Nazwa pliku z danymi projektu (z rozszerzeniem .projData)---'
'---------------------------------------------------------------'
Public Property Let DataFileName(v_fileName As String)

    m_dataFileName = v_fileName

End Property

Public Property Get DataFileName() As String

    DataFileName = m_dataFileName

End Property

'------------------------------------------'
'---Nazwa pliku, który ma zostaæ otwarty---'
'------------------------------------------'
Public Property Let TargetFileName(v_targetFileName As String)

    m_targetFileName = v_targetFileName
    Call SetFileType

End Property

Public Property Get TargetFileName() As String

    TargetFileName = m_targetFileName

End Property

'--------------------'
'---Nazwa projektu---'
'--------------------'
Public Property Let ProjectName(v_projectName As String)

    m_projectName = v_projectName

End Property

Public Property Get ProjectName() As String

    ProjectName = m_projectName

End Property


'-------------------------------------'
'---Œcie¿ka do lokalizacji projektu---'
'-------------------------------------'
Public Property Let path(v_filePath As String)

    m_path = v_filePath

End Property

Public Property Get path() As String

    path = m_path

End Property

'----------'
'---Tagi---'
'----------'
Public Property Let Tags(tagsData() As String)
    
    Dim i As Integer
    
    ReDim Preserve m_Tags(0 To UBound(tagsData))
    For i = 0 To Manager.ArraySize(tagsData)
    
        m_Tags(i) = tagsData(i)
        
    Next

End Property

Public Property Get Tags() As String()

    Tags = m_Tags

End Property

Public Function TagsAmount() As Integer

    TagsAmount = Manager.ArraySize(m_Tags) + 1
    
End Function

Public Function Tag(i As Integer) As String

    On Error GoTo TagError
    Tag = m_Tags(i)
    
    Exit Function
    
TagError:
    Tag = "Out of bound exception: Tried to access index " & i & " of " & TagsAmount() & "-element array. (0 to " & TagsAmount() - 1 & ")"

End Function

Public Sub AddTag(tagData As String)

    If m_Tags(UBound(m_Tags)) = vbNullString Then
        m_Tags(UBound(m_Tags)) = tagData
        Exit Sub
    End If
    ReDim Preserve m_Tags(Manager.ArraySize(m_Tags) + 1)
    m_Tags(UBound(m_Tags)) = tagData

End Sub

Public Function HasTag(tagName As String) As Boolean

    Dim m_tag As Variant
    HasTag = False
    
    For Each m_tag In Tags
        If m_tag = tagName Then HasTag = True
    Next

End Function

'Public Sub AddTags(tagsData() As String)
'
'    Dim i As Integer
'    Dim initialUbound As Integer: initialUbound = UBound(m_Tags)
'
'    ReDim Preserve m_Tags(0 To UBound(tagsData) + UBound(m_Tags) + 1)
'    For i = 0 To UBound(tagsData)
'
'        m_Tags(i + initialUbound + 1) = tagsData(i)
'
'    Next
'
'End Sub

'wprowadzic to do komorek w excelu tez


Public Sub AddAdditionalInfo(info As String)

    If m_AdditionalInfo(0) = vbNullString Then
        m_AdditionalInfo(0) = info
        Exit Sub
    End If

    ReDim Preserve m_AdditionalInfo(Manager.ArraySize(m_AdditionalInfo) + 1)
    m_AdditionalInfo(Manager.ArraySize(m_AdditionalInfo)) = info

End Sub

Private Sub Class_Initialize()
    
    m_projectName = "Default Project"
    
    ReDim m_Tags(0)
    ReDim m_AdditionalInfo(0)

End Sub


Private Sub SetFileType()

'----------------------------------------------------------------------------------'
'---chamsko, ale na chwilê obecn¹ nie mam wiêkszych pomys³ów na to jak to zrobiæ---'
'----------------------------------------------------------------------------------'
    Dim ExtensionLibrary As Dictionary
    Set ExtensionLibrary = New Dictionary
    With ExtensionLibrary
        .Add ".xlsx", "Excel"
        .Add ".xlsm", "Excel"
        .Add ".docx", "Word"
        .Add ".pptx", "PowerPoint"
        .Add ".pdf", "PDF"
    End With
    
    
    Dim fileTypeExtension As String
    Dim i As Integer
    For i = 1 To Len(m_targetFileName)
        If Mid$(m_targetFileName, Len(m_targetFileName) - i, 1) = "." Then
            fileTypeExtension = Right$(m_targetFileName, i + 1)
            m_fileType = ExtensionLibrary(fileTypeExtension)
            Exit Sub
        End If
    Next
    
    m_fileType = "Undefined"

End Sub
