Attribute VB_Name = "Chapter11"
Sub copyandpaste()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("11").Activate
    Range("A1").Select
    Selection.copy
    Range("B1").Select
    ActiveSheet.Paste
    
    Application.CutCopyMode = False              'remove dancing ants Range("A1")
    Range("A3").copy Range("B3")                 'Copy and paste Range("A3") to Range("B3")
    
    'Go To Sheet 11copy.
    Sheets("11copy").Select                      'Worksheets("11copy").Activate also works
    'Copy and paste Sheet 11 Range("A1") to Sheet 11copy Range("B1")
    Sheets("11").Range("A1").copy Sheets("11copy").Range("B1") 'print hello
    Worksheets("11copy").Range("B5").Select

    'copy a column of cells from worksheets 11 to worksheets 11copy
    Worksheets("11").Select
    Worksheets("11").Range("A5", Range("A5").End(xlDown)).copy Worksheets("11copy").Range("A5")
End Sub

Sub movearange()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("11").Activate
    Range("D1:D10").Value = 1
    For i = 5 To 12 Step 1:
        Range("E" & i).Value = i
    Next i
    Range("D1:D10").Cut Range("H1")
    Range("E5", Range("E5").End(xlDown)).Cut Range("I1")
End Sub

Sub copycurrentregion()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("11").Activate
    'Range("A15").CurrentRegion.Select
    Range("A15").CurrentRegion.copy Sheets("11copy").Range("A15")
    Worksheets("11copy").Select
    'Range copied is a table
    'Range("tablename[#All]").copy Sheets("11").Range("A40")
End Sub

Sub activateworksheets()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("11").Activate
    Range("A1").Select
    Worksheets("10").Activate
    'same as
    Sheets("10").Select
    Sheets("10").Range("E14").Select
End Sub

Sub selectingvariousranges()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("11").Activate
    'Range(ActiveCell, ActiveCell.End(xlDown)).select
    Range("E15", Range("E15").End(xlDown)).Select
    Range("F1").Select
    Range("F19").Select
    ActiveCell.CurrentRegion.Select              'all cells around Range("F19") are selected
    ActiveCell.CurrentRegion.Font.Bold = True
End Sub

Sub nextemptycell()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("11").Activate
    For i = 1 To 10 Step 1
        If i = 8 Then
            Range("J" & i).Value = ""
        Else
            Range("J" & i).Value = i
        End If
    Next i
    'Range("B2").End(xlDown).Offset(1, 0).Select
    'FYI Range("J1", Range("J1").End(xlDown)).Select
    Range("J1").End(xlDown).Offset(1, 0).Select
    userentry = InputBox("Enter a number")       'ask user for a number
    ActiveCell.Value = userentry                 'put user number in the active cell
End Sub

Sub countselectedcellsrowscolumns()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("11").Activate
    Range("J1", Range("J1").End(xlDown)).Select
    Range("K7").Value = Selection.Count          'print 7
    Range("B12").Value = Range("A5:A12").Count   'print 8
    Range("F:I").Select                          'highlight column F to column I
    Range("G1").Value = Selection.Columns.Count  'print 4
    Range("24:30").Select                        'highlight row 24 to row 30
    Range("A26").Value = Selection.Rows.Count    'print 7
    'The Count property uses the Long data type, so the largest value that it can store is 2,147,483,647.
    'CountLarge uses the Double data type, which can handle values up to 1.79+E^308.
End Sub

Sub loopselectedcells()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("11").Activate
    Range("A30:M30").Select
    For Each cell In Selection
        If cell.Value = "" Then
            cell.Value = "yes"
            cell.Interior.Color = RGB(255, 0, 0)
        End If
    Next cell
    Range("A30:M30").Delete
End Sub

Sub deleteemptyrowsdoesntwork()
    'sub can't delete two empty cells together
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("11").Activate
    Dim area As Range
    Set area = Range("A29:A40")
    Application.ScreenUpdating = False
    For Each cell In area:
        cell.Select
        If cell.Value = "" Then
            cell.EntireRow.Delete
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Sub celltype()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("11").Activate
    Range("E23").Value = IsNumeric(Range("E22"))
    Range("F23").Value = IsNumeric(Range("F22"))
    Range("G23").Value = IsNumeric(Range("G22"))
    Range("H23").Value = Application.IsText(Range("H22"))
    Range("I23").Value = IsDate(Range("I22"))
    Range("J23").Value = Application.IsLogical(Range("J22"))
    Range("K23").Value = IsEmpty(Range("K22"))
    Range("L23").Value = InStr(1, Range("L22"), "like", vbTextCompare) 'print 3 position like is found
End Sub

