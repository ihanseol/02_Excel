Attribute VB_Name = "worksheet"
Sub quickguideworksheets()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("worksheet").Activate
    Worksheets("worksheet").Select
    Worksheets(2).Select                         'access worksheet by position from left
    Worksheets(Worksheets.count).Select          'access right most worksheet
    'ActiveSheet Access the active worksheet
    'Dim declareworksheet As Worksheet 'error message
    'Set declareworksheet = Worksheets("worksheet")
    Worksheets.Add after:=Worksheets(Worksheets.count) 'insert new worksheet at most right
    Worksheets(Worksheets.count).Delete          'delete worksheet at most right
    Worksheets("rangecells").Copy after:=Worksheets(Worksheets.count) 'copy specific worksheet at most right
    Worksheets(Worksheets.count).Delete          'delete worksheet at most right
    Worksheets.Add after:=Worksheets(Worksheets.count) 'insert new worksheet at most right
    Worksheets(Worksheets.count).Activate
    Worksheets(Worksheets.count).name = "rename last worksheet" 'rename last worksheet name
    Worksheets(Worksheets.count).Delete          'delete worksheet at most right
    Worksheets("worksheet").Select
    Worksheets("worksheet").Visible = xlSheetHidden 'hide worksheet
    Worksheets("worksheet").Visible = xlSheetVisible 'show worksheet
    Worksheets("worksheet").Select
End Sub

Sub printworksheetsname()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("worksheet").Activate
    Dim i As Double
    For i = 1 To Worksheets.count
        Cells(i, 1).Value = Worksheets(i).name
    Next i
End Sub

Sub activesheetworksheets()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("worksheet").Activate
    'The ActiveSheet object refers to the worksheet currently active.  Use ActiveSheet _
    if you have a specific need to refer to the worksheet which is active.
    ActiveSheet.Range("B1") = "b1"
    'same as
    Range("B1") = "b11"
End Sub

Sub declareworksheetobject()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("worksheet").Activate
    'Dim sht As Worksheet 'RM error message
    'Set sht = Worksheets("worksheet")
    'sht.Range("B2") = "declare worksheet"
End Sub

Sub nutshellworksheet()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("worksheet").Activate
    'use worksheet currently active
    ActiveSheet.Range("b2").Value = "use ActiveSheet on worksheet currently active"

    'instructor suggests using Code Name to access another worksheet in same workbook or Excel file
    Sheet1.Select                                'access the first worksheet left most
    Worksheets(Worksheets.count).Activate

    'if worksheet is in a different workbook, get the workbook and get the worksheet
    Dim wk As Workbook
    Set wk = Workbooks.Open("C:\Docs\Accounts.xlsx", ReadOnly:=True)
    Dim sh As Worksheet
    Set sh = wk.Worksheets("Sheet1")
End Sub

Sub adddeletemultipleworksheets()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("worksheet").Activate
    Dim i As Double
    For i = 1 To 3 Step 1
        Worksheets.Add after:=Worksheets(Worksheets.count)
        ActiveSheet.name = "name" & i
    Next i
    For i = 1 To 3 Step 1
        Worksheets(Worksheets.count).Delete
    Next i
    Worksheets.Add after:=Worksheets(Worksheets.count)
    ActiveSheet.name = "temp delete"
    Worksheets("temp delete").Delete
End Sub

Sub addmultipleworksheetsfromtable()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("worksheet").Activate
    Dim i As Double
    For i = 15 To 17
        Worksheets.Add after:=Worksheets(Worksheets.count)
        ActiveSheet.name = Worksheets("worksheet").Range("A" & i)
    Next i
End Sub

Sub forloopeachworksheet()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("worksheet").Activate
    Dim i As Double
    'create three new worksheets with names
    For i = 1 To 3 Step 1
        Worksheets.Add after:=Worksheets(Worksheets.count)
        ActiveSheet.name = "name" & i
    Next i
    'write Hello World Range("A1") on all three new worksheets
    For i = 1 To 3 Step 1
        Worksheets("name" & i).Select
        Worksheets("name" & i).Range("A1").Value = "Hello World"
    Next i
End Sub

Sub printallworksheetsopenworkbooks()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("worksheet").Activate
    Dim openworkbooks As Workbook
    Dim i As Long
    Range("D1").Select
    For Each openworkbooks In Workbooks
        For i = 1 To openworkbooks.Worksheets.count
            ActiveCell.Value = openworkbooks.name & " " & openworkbooks.Worksheets(i).name
            ActiveCell.offset(1, 0).Select
        Next i
    Next openworkbooks
End Sub


