Attribute VB_Name = "BaseData_Util"
Option Explicit

Public Sub TurnOffStuff()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
End Sub

Public Sub TurnOnStuff()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Sub Range_End_Method()
    'Finds the last non-blank cell in a single row or column
    
    Dim lRow        As Long
    Dim lCol        As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.count).End(xlToLeft).Column
    
    MsgBox "Last Row: " & lRow & vbNewLine & _
           "Last Column: " & lCol
End Sub

Public Function lastRow() As Long
    Dim lRow        As Long
    lRow = Cells(Rows.count, 1).End(xlUp).Row
    
    lastRow = lRow
End Function

'Public Function Contains(Col As Collection, key As Variant) As Boolean
'    On Error Resume Next
'    Col (key)                                    ' Just try it. If it fails, Err.Number will be nonzero.
'    Contains = (Err.number = 0)
'    Err.Clear
'End Function

Function Contains(objCollection As Object, strName As String) As Boolean
    Dim o           As Object
    On Error Resume Next
    
    Set o = objCollection(strName)
    Contains = (Err.Number = 0)
    Err.Clear
End Function

Function RemoveDupesDict(MyArray As Variant) As Variant
    'DESCRIPTION: Removes duplicates from your array using the dictionary method.
    'NOTES: (1.a) You must add a reference to the Microsoft Scripting Runtime library via
    ' the Tools > References menu.
    ' (1.b) This is necessary because I use Early Binding in this function.
    ' Early Binding greatly enhances the speed of the function.
    ' (2) The scripting dictionary will not work on the Mac OS.
    'SOURCE: https://wellsr.com
    '-----------------------------------------------------------------------
    Dim i           As Long
    Dim d           As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    
    With d
        For i = LBound(MyArray) To UBound(MyArray)
            If IsMissing(MyArray(i)) = False Then
                .Item(MyArray(i)) = 1
            End If
        Next
        RemoveDupesDict = .Keys
    End With
End Function

Public Function GetLength(a As Variant) As Integer
    ' if array is empty return 0
    ' else return number of array item
    
    If IsEmpty(a) Then
        GetLength = 0
    Else
        GetLength = UBound(a) - LBound(a) + 1
    End If
End Function

Public Function getUnique(ByRef array_tabcolor As Variant) As Variant
    ' remove duplicated item in array
    ' and return unique array value
    
    Dim array_size  As Variant
    Dim new_array   As Variant
    
    new_array = RemoveDupesDict(array_tabcolor)
    getUnique = new_array
End Function

Function ConvertToLongInteger(ByVal stValue As String) As Long
    On Error GoTo ConversionFailureHandler
    ConvertToLongInteger = CLng(stValue)        'TRY to convert to an Integer value
    Exit Function        'If we reach this point, then we succeeded so exit
    
ConversionFailureHandler:
    'IF we've reached this point, then we did not succeed in conversion
    'If the error is type-mismatch, clear the error and return numeric 0 from the function
    'Otherwise, disable the error handler, and re-run the code to allow the system to
    'display the error
    If Err.Number = 13 Then        'error # 13 is Type mismatch
        Err.Clear
        ConvertToLongInteger = 0
        Exit Function
    Else
        On Error GoTo 0
        Resume
    End If
End Function


