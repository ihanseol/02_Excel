Attribute VB_Name = "colornamesAblue"
Sub colornamesblue()
    Workbooks("excelmacromasteryemailtips.xlsm").Sheets("Names").Activate
    Range("A1:A20").ClearFormats
    For Each cell In Range("A1:A20")
        If Left(cell.Value, 1) = "A" Then
            cell.Font.Color = rgbBlue
        End If
    Next cell
End Sub

Sub colornamesbluevariables()
    'Workbooks("excelmacromasteryemailtips.xlsm").Sheets("Names").Activate
    ThisWorkbook.Worksheets("Names").Activate    'also works.  Using thisworkbook VBA code works _
                                                 even if someone changes the filename or workbook name
    Dim setrange As Range
    Set setrange = Range("A1:A20")
    setrange.ClearFormats                        'clear all formats
    For Each cell In setrange
        If Left(cell.Value, 1) = "A" Then
            cell.Font.Color = rgbBlue
        End If
    Next cell
End Sub

Sub findlastrow()
    Workbooks("excelmacromasteryemailtips.xlsm").Sheets("Names").Activate
    Dim lastrow As Long, lastcolumn As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).row
    lastcolumn = Cells(1, Columns.Count).End(xlToLeft).Column
    MsgBox lastrow & " " & lastcolumn
    Range("A5").currentregion.Select
End Sub


