Attribute VB_Name = "Objects"
Sub SimpleObjectCreated()
Workbooks("MicrosoftExcelProgrammingChapter04IntroducingTheExcelObjectModelObjects.xlsm").Worksheets("Sheet1").Activate
Dim Title As Range
Set Title = Range("B1")
'Set Title = ActiveSheet.Range("B1") 'No need for ActiveSheet.
    Title.Value = "Sales"
    Title.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Title.Borders(xlEdgeBottom).Weight = xlThick
    Title.Font.Bold = True
    Title.Font.Color = RGB(0, 0, 255)
    Title.HorizontalAlignment = xlRight
End Sub

Sub SimpleObjectCreatedWithWithStatement()
Workbooks("MicrosoftExcelProgrammingChapter04IntroducingTheExcelObjectModelObjects.xlsm").Worksheets("Sheet1").Activate
Dim Title As Range
Set Title = Range("C1")
With Title
    .Value = "Quantity"
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).Weight = xlThick
    .HorizontalAlignment = xlRight
    With .Font
        .Bold = True
        .Color = RGB(0, 0, 255)
     End With
End With
End Sub

Sub UsingAnObjectMethod()
Workbooks("MicrosoftExcelProgrammingChapter04IntroducingTheExcelObjectModelObjects.xlsm").Worksheets("Sheet1").Activate
Dim DeleteRange As Range
Set DeleteRange = Range("D4", "E59")
DeleteRange.Select
DeleteRange.Interior.Color = rgbLightGoldenrodYellow
DeleteRange.Value = "Delete Me"
DeleteRange.Delete (xlShiftToLeft)
End Sub
