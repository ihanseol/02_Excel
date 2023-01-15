Attribute VB_Name = "copyrangecellsvalues"
Sub copyrangecells()
    Workbooks("excelmacromasteryemailtips.xlsm").Worksheets("copynames").Activate
    Range("A1").Select
    Dim markscreatearray As Variant
    
    markscreatearray = Worksheets("Names").Range("A1:A20").Value
    Worksheets("copyNames").Range("A1:A20").Value = markscreatearray
End Sub

