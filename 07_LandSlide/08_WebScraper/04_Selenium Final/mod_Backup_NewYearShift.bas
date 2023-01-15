Attribute VB_Name = "mod_Backup_NewYearShift"
Option Explicit

Sub BackupData()

    Sheets("main").Select
    Sheets("main").Copy Before:=Sheets(3)
    ActiveWindow.SmallScroll Down:=-15
    
    
    ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
        
    Columns("R:X").Select
    Selection.Delete Shift:=xlToLeft
    Range("Q10").Select
    

    On Error GoTo Catch
    ActiveSheet.name = Range("'main'!$S$8").Value
    Range("B2").Value = Range("'main'!$S$8").Value & " Data, -- " & Now()
Catch:
    Exit Sub
    
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


Function get_Single_data_bySeleniumXpath(ByVal nArea As Integer) As String()
 
    Dim i As Long
    Dim bot As New ChromeDriver
    Dim td As Selenium.WebElement
    
    Dim url As String
    Dim SearchString As String
    Dim out(0 To 12) As String
    Dim nYear As Integer
    
 
    nYear = Year(Now()) - 1
    url = "https://www.weather.go.kr/weather/climate/past_table.jsp?stn=" & nArea & "&yy=" & nYear & "&obs=21&x=22&y=12"
    
    With bot
        .AddArgument "--headless"   ''This is the fix
        .Get url
    End With
    
    out(0) = nYear
    
    For i = 2 To 13
        SearchString = "//*[@id=""content_weather""]/table/tbody/tr[32]/td[" & i & "]"
        Set td = bot.FindElementByXPath(SearchString)
        out(i - 1) = CStr(td.text)
    Next i
    
    get_Single_data_bySeleniumXpath = out
    
    bot.Close
    Set bot = Nothing
    
End Function



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


Sub get_single_data()

    Dim resOut() As String
    Dim i, j, k  As Integer
    
    Dim nArea, nYear As Integer
       
    nYear = Year(Now()) - 1
    nArea = get_currentarea_code()
    
    If nArea = 0 Then
        MsgBox " Area Code Does not match ... "
        Exit Sub
    End If
    
    resOut = get_Single_data_bySeleniumXpath(nArea)
     
    For i = 0 To 12
        Cells(35, i + 2).Value = resOut(i)
    Next i
    
    Call ChangeFormat
    Call delete_ignore_error
End Sub


