Option Explicit

Sub BackupData()

    Sheets("main").Select
    Sheets("main").Copy Before:=Sheets(1)
    ActiveWindow.SmallScroll Down:=-15
    
    
    ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
        
    Columns("R:X").Select
    Selection.Delete Shift:=xlToLeft
    Range("Q10").Select
    

    On Error GoTo Catch
    ActiveSheet.name = Range("'main'!$S$8").Value
    Range("B2").Value = Range("'main'!$S$8").Value & " Data, -- " & Now()
    SetRandomSheetTabColor
    
Catch:
    Exit Sub
    
End Sub


Sub SetRandomSheetTabColor()

    ' Define an array of RGB colors
    Dim Colors(1 To 10) As Long
    
    Colors(1) = RGB(192, 0, 0) ' Red
    Colors(2) = RGB(255, 165, 0) ' Orange
    Colors(3) = RGB(255, 255, 0) ' Yellow
    Colors(4) = RGB(0, 176, 80) ' Green
    Colors(5) = RGB(0, 112, 192) ' Blue
    Colors(6) = RGB(112, 48, 160) ' Purple
    Colors(7) = RGB(128, 128, 128) ' Gray
    Colors(8) = RGB(255, 192, 203) ' Pink
    Colors(9) = RGB(128, 64, 64) ' brown
    Colors(10) = RGB(64, 224, 208) ' turquoise

    ' Generate a random number to select a color from the array
    Randomize
    Dim RandomIndex As Integer
    RandomIndex = Int((UBound(Colors) + 1) * Rnd)
    
    ' Set the sheet tab color to the randomly selected color
    ActiveSheet.Tab.Color = Colors(RandomIndex)

End Sub


Sub ShiftUp()

    Range("B7:N35").Select
    Selection.Copy
    
    Range("B6").Select
    ActiveSheet.PasteSpecial Format:=3, link:=1, DisplayAsIcon:=False, IconFileName:=False
    
    Range("B35:N35").Select
    Selection.ClearContents
    
End Sub


Sub CopySingleData()

    Dim i As Integer
    
    For i = 0 To 12
        Cells(35, i + 2).Value = Sheets("main").Cells(40, i + 2).Value
    Next i

End Sub


Sub ShiftNewYear()

    Dim nYear As Integer

    nYear = Year(Now()) - 30
    
    If Range("B6").Value = nYear Then
        Exit Sub
    End If
    
    Call ShiftUp
    Call get_single_data

End Sub



Function get_currentarea_code() As Integer

    Dim name As String
    Dim nArea As Integer
    Dim tbl As ListObject
    
    On Error GoTo Process
    
    Set tbl = Sheets("Code").ListObjects("tblCode")
    name = ActiveSheet.name
   
    nArea = Application.WorksheetFunction.VLookup(name, tbl.Range, 2, False)
    get_currentarea_code = nArea
    Exit Function
    
Process:
    nArea = 0
    get_currentarea_code = nArea
        
End Function



