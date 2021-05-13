Attribute VB_Name = "Manager"
Option Explicit

Public Sub SetManualCalculation()

    Application.ScreenUpdating = False
    Application.calculation = xlCalculationManual

End Sub

Public Sub SetAutomaticCalculation()

    Application.ScreenUpdating = True
    Application.calculation = xlCalculationAutomatic

End Sub

Public Function ArraySize(arr As Variant) As Integer

    On Error GoTo ErrorStatement
    ArraySize = UBound(arr) - LBound(arr)
    Exit Function
    
ErrorStatement:
    Debug.Print ("Passed argument for ArraySize function is not an array. Returning 0")
    ArraySize = 0
    Exit Function

End Function

Public Function IsInCollection(container As Collection, itemKey As Variant) As Boolean

    Dim obj As Variant
    
    On Error GoTo err
    
    IsInCollection = True
    obj = container(itemKey)
    Exit Function
        
err:
    IsInCollection = False
        
End Function

Public Sub InitializeForm()

    FilesearchForm.Show

End Sub
