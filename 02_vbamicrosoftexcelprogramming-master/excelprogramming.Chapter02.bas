Attribute VB_Name = "Chapter02"
Sub assignvalue()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("2").Activate
    Range("A1").Value = 100
End Sub

Sub assignnumbers()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("2").Activate
    
    Dim x As Integer
    
    Range("A1:A10").clearcontents
    For x = 1 To 10 Step 1
        Cells(x, 1).Value = x
    Next x
End Sub
