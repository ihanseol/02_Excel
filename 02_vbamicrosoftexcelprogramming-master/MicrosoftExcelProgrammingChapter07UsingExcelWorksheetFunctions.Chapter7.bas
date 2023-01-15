Attribute VB_Name = "Chapter7"
Sub UseWorksheetFunction()
    Worksheets("Sheet1").Activate
    
    Range("A1").Select
    Range("A15:A19").Font.Bold = True
    Range("A15").Value = WorksheetFunction.Sum(Range("A5:A14"))
    Range("A16").Value = WorksheetFunction.Max(Range("A5:A14"))
    Range("A17").Value = WorksheetFunction.Min(Range("A5:A14"))
    Range("A18").Value = WorksheetFunction.Average(Range("A5:A14"))
    Range("A19").Value = Round(WorksheetFunction.Average(Range("A5:A14")))
    Range("A20").Value = Round(WorksheetFunction.Average(Range("A5:A14"), 4))
End Sub

Sub DisplayInputBox()
    Dim Choice As String
    Workbooks("MicrosoftExcelProgrammingChapter07UsingExcelWorksheetFunctions.xlsm").Worksheets("Sheet1").Activate
    Choice = InputBox("Choose Sum, Max, or Min.", "Title here", "Sum")
    
    If Choice = "Sum" Then
        Range("B15").Value = WorksheetFunction.Sum(Range("A5:A14"))
    ElseIf Choice = "Max" Then
        Range("B15").Value = WorksheetFunction.Max(Range("A5:A14"))
    ElseIf Choice = "Min" Then
        Range("B15").Value = WorksheetFunction.Min(Range("A5:A14"))
    Else
        Range("B15").Value = "Invalid choice"
    End If
End Sub

Sub ShowDate()
    Workbooks("MicrosoftExcelProgrammingChapter07UsingExcelWorksheetFunctions.xlsm").Worksheets("Sheet1").Activate
    Range("A1").Select
    Range("C5").Value = Date
    Range("C6").Value = Now
    Range("C7").Value = Time
End Sub

Sub CalculateTwoDifferentDates()
    Workbooks("MicrosoftExcelProgrammingChapter07UsingExcelWorksheetFunctions.xlsm").Worksheets("Sheet1").Activate
    Range("A1").Select
    Dim Date1 As Date
    Dim Date2 As Date
    'Date1 = Range("D5").Value
    Date1 = Now
    Date2 = Range("D6").Value
    Range("D7").Value = DateDiff("d", Date1, Date2)
    Range("D8").Value = DateDiff("ww", Date1, Date2)
End Sub

Sub NumberFormat()
    Workbooks("MicrosoftExcelProgrammingChapter07UsingExcelWorksheetFunctions.xlsm").Worksheets("Sheet1").Activate
    Range("A1").Select
    Range("G5").Value = FormatNumber(Range("F5"), 2)
    Range("G6").Value = FormatCurrency(Range("F6"))
    Range("G7").Value = FormatCurrency(Range("F7"), 2, vbFalse, vbTrue)
    Range("G8").Value = FormatPercent(Range("F8"), 2)
End Sub

Sub ChangeCase()
    Workbooks("MicrosoftExcelProgrammingChapter07UsingExcelWorksheetFunctions.xlsm").Worksheets("Sheet1").Activate
    Range("A1").Select
    Range("B22").Value = UCase(Range("A22").Value)
    Range("c22").Value = LCase(Range("A22").Value)
End Sub

Sub PortionOfString()
    Workbooks("MicrosoftExcelProgrammingChapter07UsingExcelWorksheetFunctions.xlsm").Worksheets("Sheet1").Activate
    Range("A1").Select
    
    Dim sentence As String
    sentence = Range("A23").Value
    Range("E23").Value = Left(sentence, 11)
    Range("E24").Value = Mid(sentence, 17, 8)
    Range("E25").Value = Right(sentence, 7)
End Sub

