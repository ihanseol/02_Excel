Attribute VB_Name = "Arrays"
Sub DeclareArray()
Dim Month(1 To 3) As String
    Month(1) = "Jan"
    Month(2) = "Feb"
    Month(3) = "Mar"
    Cells(1, 1).Value = Month(1)
    Cells(2, 1).Value = Month(2)
    Cells(3, 1).Value = Month(3)
End Sub

Sub DeclareArrayAtZero()
Dim Names(0 To 5) As String
Names(0) = "George"
Names(1) = "Robert"
Names(2) = "Cathy"
Names(3) = "Leslie"
Names(4) = "Todd"
Names(5) = "Janelle"
counter = 0
Do Until counter = 6
    Range("B" & counter + 1).Value = Names(counter)
    counter = counter + 1
Loop
End Sub
