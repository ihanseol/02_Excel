Attribute VB_Name = "Module1"
Option Explicit
Public rPopulation As Range, rX As Range, sDate As String
Sub RangeClearContents()

    Set rPopulation = Range("A1").CurrentRegion.Columns(1)
    Set rPopulation = rPopulation.Offset(1).Resize(rPopulation.Rows.Count - 1)

    With rPopulation
        .Offset(, 6).ClearContents
        .Offset(, 8).ClearContents
        .Offset(, 10).ClearContents
    End With
End Sub

Sub Button1_Click()
Dim rPopulation As Range, rX As Range, sDate As String
    Set rPopulation = Range("A1").CurrentRegion.Columns(1)
    Set rPopulation = rPopulation.Offset(1).Resize(rPopulation.Rows.Count - 1)
    
    Application.Calculation = xlCalculationManual
    RangeClearContents
    
    For Each rX In rPopulation.Cells
        sDate = rX.Offset(, 5)
        rX.Offset(, 6) = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Right(sDate, 2))
    Next rX
    
    For Each rX In rPopulation.Cells
        sDate = Replace(rX.Offset(, 7), ".", "")
        rX.Offset(, 8) = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Right(sDate, 2))
    Next rX
    
    
    For Each rX In rPopulation.Cells
        sDate = Replace(rX.Offset(, 9), ".", "")
        rX.Offset(, 10) = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Right(sDate, 2))
    Next rX
    Application.Calculation = xlCalculationAutomatic
End Sub

Function ConvertToDate(rX As Range)
Dim sX As String
    sX = Replace(rX, ".", "")
    ConvertToDate = DateSerial(Left(sX, 4), Mid(sX, 5, 2), Right(sX, 2))
End Function
Sub DateRange(rPopulation As Range)
Dim rX As Range
    For Each rX In rPopulation.Cells
        rX.Offset(, 1) = ConvertToDate(rX)
    Next rX
End Sub
Sub Button2_Click()
Dim rPopulation As Range, rX As Range, sDate As String
    Set rPopulation = Range("A1").CurrentRegion.Columns(1)
    Set rPopulation = rPopulation.Offset(1).Resize(rPopulation.Rows.Count - 1)
    
    Application.Calculation = xlManual
    RangeClearContents
    
    DateRange rPopulation.Offset(, 5)
    DateRange rPopulation.Offset(, 7)
    DateRange rPopulation.Offset(, 9)
    
    Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub CompleteWeekday()
'/// WeekDay
Dim dDate As Date, rLink As Range
    Set rPopulation = Range("A1").CurrentRegion.Columns(1)
    Set rPopulation = rPopulation.Offset(1).Resize(rPopulation.Rows.Count - 1)
    Set rLink = Range("CellLink")
    
    For Each rX In rPopulation.Cells
        dDate = rX.Offset(, 6)
        rX.Offset(, 11) = ConvertToWeekday(dDate, CInt(rLink))
    Next rX
    
End Sub

Function ConvertToWeekday(dDate As Date, iType As Integer)
Select Case iType
    Case 1
        ConvertToWeekday = Choose(Weekday(dDate), "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    Case 2
        ConvertToWeekday = Choose(Weekday(dDate), "일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일")
End Select
End Function
