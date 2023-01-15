Attribute VB_Name = "Chapter05"
Sub declarearray()
    'arrayname(50) no lower bound index count starts at 0 and ends at 50
    'if you want all arrays start lower bound 1 type Option Base 1 before declare arrays
    Application.Workbooks("excelprogramming.xlsm").Worksheets("5").Activate
    'better than dim month1 as String, dim month2 as String, dim month3 as String
    Dim month(1 To 3) As String
    month(1) = "Jan"
    month(2) = "Feb"
    month(3) = "Mar"
    Range("A10").Value = month(1)
    Cells(11, 1).Value = month(2)
    [a12].Value = month(3)
    'skipped multidimensional array page 78-79.
End Sub

