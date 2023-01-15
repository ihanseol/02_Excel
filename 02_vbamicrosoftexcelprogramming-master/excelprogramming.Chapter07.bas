Attribute VB_Name = "Chapter07"
Sub worksheetfunctionexercise()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("7").Activate
    'Use the WorksheetFunction property to use an Excel worksheet function in VBA; _
    e.g. max(), min(), average()
    Dim salesrange As Range
    Set salesrange = Range("A2", Range("A2").End(xlDown))
    Range("A13") = worksheetfunction.Min(salesrange)
    Range("A14") = worksheetfunction.Max(salesrange)
    Range("A15") = worksheetfunction.Average(salesrange)
End Sub

Sub messagebox()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("7").Activate
    Dim salesrange As Range
    Set salesrange = Range("A2", Range("A2").End(xlDown))
    mbPrompt = "Calculate Total Sales?"
    mbButton = vbYesNo + vbQuestion
    mbTitle = "Calculate the Sales Total"
    answer = MsgBox(mbPrompt, mbButton, mbTitle)
    If answer = 6 Then
        MsgBox Format(worksheetfunction.Sum(salesrange), "#,##0")
    ElseIf answer = 7 Then
        MsgBox "no!  tough, the sum is " & Format(worksheetfunction.Sum(salesrange), "#,##0")
    End If
End Sub

Sub datetime()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("7").Activate
    Range("C1").Value = Date                     'print today's date 4/17/2018
    Range("C2").Value = Time                     'print current time 7:22:26 PM
    Range("C3").Value = Format(Time, "h:m")      'print current time 19:26
    Range("C4").Value = Format(Time, "hh:mm AM/PM") 'print current time 7:30 PM

    Range("E8").Value = DateDiff("yyyy", Range("C8").Value, Range("D8").Value) 'print years difference
    Range("E9").Value = DateDiff("m", Range("C8").Value, Range("D8").Value) 'print months difference
    Range("E10").Value = DateDiff("d", Range("C8").Value, Range("D8").Value) 'print days difference
    Range("E11").Value = DateDiff("ww", Range("C8").Value, Range("D8").Value) 'print weeks difference
    Range("E12").Value = DateDiff("h", Range("C8").Value, Range("D8").Value) 'print hours difference

    Range("E14").Value = FormatDateTime(Range("C8").Value, vbGeneralDate) 'print 4/17/2018
    Range("E15").Value = FormatDateTime(Range("C8").Value, vbLongDate) 'print Tuesday, April 17, 2018
    Range("E16").Value = FormatDateTime(Range("C8").Value, vbShortDate) 'print 4/17/2018
    Range("E17").Value = FormatDateTime(Range("C8").Value, vbLongTime) 'print 12:00:00 AM
End Sub

Sub formatnumericstring()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("7").Activate
    Range("E19").Value = FormatCurrency(Range("C19").Value, 0) 'print $12,545
    Range("E20").Value = FormatNumber(Range("C20").Value, 0) 'print 12,545
    Range("E21").Value = Format(Range("C21").Value, "$#,##0;($#,##0)") 'print -$12,545 aligned right
    Range("E22").Value = Format(Range("C22").Value, "#,###") 'print 12,545
    Range("E24").Value = UCase(Range("C24").Value) 'print HELLO
    Range("E25").Value = LCase(Range("C25").Value) 'print hello
End Sub

Sub leftmidrightstring()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("7").Activate
    Dim word As String
    Range("C27:C29").Value = "thequickbrownfox"
    Range("E27").Value = Left(Range("C27").Value, 5) 'print thequ
    Range("E28").Value = Right(Range("C28").Value, 5) 'print wnfox
    Range("E29").Value = Mid(Range("C29").Value, 5, 3) 'print uic
End Sub


