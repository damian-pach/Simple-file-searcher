VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringModifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DefaultPL As String
Private m_UnicodePL As String

Private m_defPLarr() As String
Private m_UniPLarr() As String


Private Sub Class_Initialize()
    
    m_DefaultPL = "Í,Û,π,ú,≥,ø,ü,Ê,Ò, ,”,•,å,£,Ø,è,∆,—"
    m_UnicodePL = "&H119,&HF3,&H105,&H15B,&H142,&H17C,&H17A,&H107,&H144,&H118,&HD3,&H104,&H15A,&H141,&H17B,&H179,&H106,&H143"
    
    
    m_defPLarr() = Split(m_DefaultPL, ",")
    m_UniPLarr() = Split(m_UnicodePL, ",")
    
End Sub

Public Function KeepPolishLetters(inputString As String) As String

    Dim outputString As String: outputString = inputString
    Dim i As Integer
    
    For i = 1 To UBound(m_defPLarr)
    
        outputString = Replace(outputString, m_defPLarr(i), ChrW$(m_UniPLarr(i)))
    
    Next i
    
    KeepPolishLetters = outputString

End Function
