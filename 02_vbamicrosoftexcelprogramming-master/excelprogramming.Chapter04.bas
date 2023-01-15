Attribute VB_Name = "Chapter04"
Sub createanobject()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("4").Activate
    'Application Object, Workbook Object, Worksheet Object _
    Chart Object, Range Object Dialog Object
    Dim title As Range
    Set title = Range("A1:E1")
    With title
        .MergeCells = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    title.Value = "title variable Range A1 to E1"
End Sub

Sub somecellproperties()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("4").Activate
    Dim cell As Range
    Dim hellocells As Range
    Set cell = Range("A5")
    Set hellocells = Range("D4:E8")
    With cell
        .Value = "Cell A5"
        .Font.Bold = True
        .Font.Color = RGB(255, 0, 0)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeBottom).Color = RGB(0, 255, 0)
    End With
    hellocells.Value = "hello d4, hello e8"
    Range("A8:e8").Value = "going up"
    Range("7:7").Delete (xlShiftToUp)
End Sub


