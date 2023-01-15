Attribute VB_Name = "Module2"
Option Explicit

Sub OptionButton5_Click()
CompleteWeekday
End Sub
Sub OptionButton6_Click()
CompleteWeekday
End Sub
Sub Button7_Click()
'/// WeekDay
Dim dDate As Date, rLink As Range
    Set rPopulation = Range("A1").CurrentRegion.Columns(1)
    Set rPopulation = rPopulation.Offset(1).Resize(rPopulation.Rows.Count - 1)
    
    For Each rX In rPopulation.Cells
        rX.Offset(, 11) = WeekdayName(Weekday(rX.Offset(, 6)))
    Next rX
    
End Sub
