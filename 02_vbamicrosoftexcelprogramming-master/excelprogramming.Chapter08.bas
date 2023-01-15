Attribute VB_Name = "Chapter08"
Sub breakpointrunlocalswindow()
    'Also add Watches Debug-->Add Watch
    'Quck Watch Debug-->Quick Watch such as on a line with a variable
    Application.Workbooks("excelprogramming.xlsm").Worksheets("8").Activate
    Dim number1, number2, answer As Integer
    number1 = Range("A1").Value
    number2 = Range("A2").Value
    answer = number1 + number2
    Range("A3").Value = answer
End Sub

Sub watchrunwatcheswindow()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("8").Activate
    Dim watch1 As Integer
    Dim cell As Range
    watch1 = 40
    For Each cell In Range("B1:B10")
        cell.Value = watch1
        watch1 = watch1 + 5                      'set quick watch here
    Next cell
End Sub

