Attribute VB_Name = "Chapter10"
Sub addworksheet()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("10").Activate
    ThisWorkbook.Sheets.Add before:=Sheets("1"), count:=1, Type:=xlWorksheet
    ThisWorkbook.Sheets.Add after:=Sheets("1"), count:=1, Type:=xlWorksheet
    'add worksheet after third position or add worksheet fourth position
    ThisWorkbook.Sheets.Add after:=Sheets(3), count:=1, Type:=xlWorksheet
    'add worksheet last
    ThisWorkbook.Sheets.Add after:=Sheets(Sheets.count), count:=1, Type:=xlWorksheet
    'a different way add worksheet Workbooks("excelprogramming.xlsm").Worksheets.Add after:=Sheets(Sheets.count)
    'a different way add worksheet Worksheets.Add after:=Sheets(Sheets.count)
End Sub

Sub deleteworksheet()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("10").Activate
    ThisWorkbook.Sheets.Add before:=Sheets(1), count:=1, Type:=xlWorksheet
    'delete the first worksheet
    Sheets(1).Delete
    'Worksheets(1).Delete 'also works
    ThisWorkbook.Sheets.Add after:=Sheets(Sheets.count), count:=1, Type:=xlWorksheet
    'delete the last worksheet
    Sheets(Sheets.count).Delete

    'delete a worksheet name
    'Sheets("worksheetname").Delete

    'deactivate the alerts then reactivate the alerts
    'Application.DisplayAlerts = False
    'Application.DisplayAlerts = True
End Sub

Sub moveworksheet()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("10").Activate
    ThisWorkbook.Sheets.Add before:=Sheets(1), count:=1, Type:=xlWorksheet
    'move first worksheet to last position
    Sheets(1).Move after:=Sheets(Sheets.count)
    'delete worksheet last position
    Sheets(Sheets.count).Delete

    'temporary move sheet "7"
    Sheets("7").Move after:=Sheets("2")
    'move sheet "7" back to position before sheet "8"
    Sheets("7").Move before:=Sheets("8")
End Sub

Sub copyworksheet()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("10").Activate
    ThisWorkbook.Sheets.Add after:=Sheets(Sheets.count), count:=1, Type:=xlWorksheet
    Range("A1").Value = "Copy the worksheet"
    ActiveSheet.Copy before:=Sheets(1)
    'a different way Sheets(1).Copy after:=Sheets(3)
    'a different way copy fifth worksheet before sheet named "9" Worksheets(5).Copy before:=Sheets("9")
End Sub

Sub hideworksheet()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("10").Activate
    Sheets(Sheets.count).Visible = False
    Sheets(Sheets.count).Visible = True
    Worksheets(5).Visible = False
    Worksheets(5).Visible = True
    Sheets("3").Visible = False
    Sheets("3").Visible = True
    Worksheets("5").Visible = False
    Worksheets("5").Visible = True
End Sub

Sub renameworksheet()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("10").Activate
    'Workbooks("excelprogramming.xlsm").Worksheets.Add after:=Sheets(Sheets.count)
    Worksheets.Add after:=Sheets(Sheets.count)
    Sheets(Sheets.count).Name = "Rename sheet"
    Sheets(Sheets.count).Delete
    Sheets("10").Name = "1010"
    Sheets("1010").Name = "10"
End Sub

Sub printworksheet()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("10").Activate
    Range("A1").Value = "print"
    ActiveSheet.PageSetup.Orientation = xlLandscape
    ActiveSheet.PageSetup.PrintArea = "$A$1:$D$13"
    ActiveSheet.PrintOut Copies:=1, preview:=True, ActivePrinter:="Adobe PDF"
    'if you don't set ActivePrinter, then VBA uses default
End Sub

