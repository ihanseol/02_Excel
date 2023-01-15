Attribute VB_Name = "currentregion"
Sub copycells()
    Workbooks("excelmacromasteryemailtips.xlsm").Sheets("currentregion").Activate
    Dim rg As Range, row As Range
    
    Set rg = Range("B2").currentregion
    
    Range("B15").Select
    
    For Each row In rg.Rows
        ActiveCell.Value = row.Cells(1, 1)
        ActiveCell.Offset(1, 0).Select
    Next
End Sub

Sub copycellsnoheader()
    Workbooks("excelmacromasteryemailtips.xlsm").Sheets("currentregion").Activate
    Dim rg As Range, row As Range
    
    Set rg = Range("B2").currentregion
    Range("B15").Select
    
    For i = 2 To rg.Rows.Count
        ActiveCell.Value = rg.Cells(i, 1)
        ActiveCell.Offset(1, 0).Select
    Next
End Sub

