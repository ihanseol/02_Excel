Attribute VB_Name = "Chapter05usedefined"
Type Course
    CourseName As String
    Unit As String
    numberofstudents As Integer
End Type

Sub userdefinedtypesexample()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("5").Activate
    Dim course1 As Course
    Dim course2 As Course
    course1.CourseName = "math"
    course1.Unit = "calculus"
    course1.numberofstudents = 40
    Range("A10").Value = course1.CourseName
    Range("A11").Value = course1.Unit
    Range("A12").Value = course1.numberofstudents

    course2.CourseName = "business"
    course2.Unit = "accounting"
    course2.numberofstudents = 50
    Range("A13").Value = course2.CourseName
    Range("A14").Value = course2.Unit
    Range("A15").Value = course2.numberofstudents
End Sub

