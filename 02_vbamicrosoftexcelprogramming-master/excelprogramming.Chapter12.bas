Attribute VB_Name = "Chapter12"
Sub cutpaste()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("12").Activate
    Range("E10:G15").clearcontents
    Range("E10:G15").ColumnWidth = 7.65
    Range("A1").Value = "Cut me"
    Range("A1").Cut Range("C1")                  'move Cut me from A1 to C1
    n = 1
    For Each cell In Range("A10:C15")
        cell.Value = n
        n = n + 1
    Next
    Range("A10:C15").Cut Range("E10:G15")        'move numbers and paste Range E10:G15
    'also works
    'Range("A10:C15").Cut Range("E10") 'move numbers and paste left corner Range E10
    Range("E10:G15").Columns.AutoFit
End Sub

Sub copypaste()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("12").Activate
    Range("A20:C25").clearcontents
    Range("D1").clearcontents
    'Range("E20:G25").clearcontents
    'also works
    Range("E20", Range("E20").End(xlDown).End(xlToRight)).clearcontents
    Range("B1").Value = "Copy me"
    Range("B1").Copy Range("D1")
    n = 1
    For Each cell In Range("A20:C25")
        cell.Value = n
        cell.Font.Bold = False
        n = n + 1
    Next
    Range("A20:C25").Copy Range("E20:G25")       'move numbers and paste Range E20:G25
    'also works
    'Range("A20:C25").Copy Range("E20") 'move numbers and paste left corner Range E20
    Range("E20:G25").Columns.AutoFit
End Sub

Sub bonuscolorindex()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("12").Activate
    'examples
    'Range("E21").Font.ColorIndex = (1) 'Black
    'Range("E21").Border.ColorIndex = (3) 'Red
    'Range("E21").Interior.ColorIndex = (5) 'Blue
    For Each cell In Range("E20:G25")
        cell.Font.ColorIndex = cell.Value
    Next
    'Page 199 lists colors 1-16.  1 is black, 2 is white, 5 is blue, 6 is yellow, 9 is brown.
    'There are 56 colors ColorIndex.
End Sub

Sub copypastespecial()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("12").Activate
    n = 1
    For Each cell In Range("A20:C25")
        cell.Value = n
        cell.Font.Bold = True
        n = n + 1
    Next
    Range("A20", Range("A20").End(xlDown).End(xlToRight)).Copy
    Range("A30").PasteSpecial                    'default is Paste:=xlPasteAll start at Cell A30
    Range("E30").PasteSpecial xlPasteValues      'not bold start at Cell E30
    Application.CutCopyMode = False              'stop copy
    'Page 201 other Paste Special parameters include xlPasteColumnWidths, xlPasteFormulas
End Sub

Sub copytosheets()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("12odd").Activate
    Range("A1").Select
    Range("A1", Range("A1").End(xlDown)).clearcontents
    Application.Workbooks("excelprogramming.xlsm").Worksheets("12even").Activate
    Range("A1").Select
    Range("A1", Range("A1").End(xlDown)).clearcontents
    Application.Workbooks("excelprogramming.xlsm").Worksheets("12").Activate
    Range("J1").Select
    'n = 1
    'Do While n < 21
    '    Range("J" & n).Value = n
    '    n = n + 1
    'Loop
    n = 1
    Do While n < 21
        If Cells(n, 10).Value Mod 2 = 0 Then
            Cells(n, 10).Copy
            Application.Workbooks("excelprogramming.xlsm").Worksheets("12even").Activate
            ActiveCell.PasteSpecial
            ActiveCell.offset(1, 0).Select
        Else
            Cells(n, 10).Copy
            Application.Workbooks("excelprogramming.xlsm").Worksheets("12odd").Activate
            ActiveCell.PasteSpecial
            ActiveCell.offset(1, 0).Select
        End If
        n = n + 1
        Application.Workbooks("excelprogramming.xlsm").Worksheets("12").Activate
    Loop
    Application.CutCopyMode = False              'stop copy
End Sub

Sub findword()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("12").Activate
    Range("N1").Select
    'On Error Resume Next 'continue processing if an error occurs
    Range("L1:L5").Find(What:="12odd", LookIn:=xlValues, LookAt:=xlPart, _
                        searchorder:=xlByRows, MatchCase:=False).Activate
    'RM:  highlights Cell L3 only.  It doesn't check all cells L1:L5.
End Sub

Sub findreplaceword()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("12").Activate
    Range("L7:L12").Replace What:="region1", replacement:="north", _
                            LookAt:=xlWhole, MatchCase:=False
End Sub

Sub findreplaceformatting()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("12").Activate
    With Application.ReplaceFormat.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 8
        .ColorIndex = 10
    End With
    Range("L14:L19").Replace What:="region1", replacement:="region11", _
                             ReplaceFormat:=True, LookAt:=xlWhole, MatchCase:=False
End Sub

Sub findreplaceformattingifstatement()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("12").Activate
    'RM:  I believe .replace and replaceformat are unnecessary.  Just change the font and value.
    For Each cell In Range("L21:L26")
        If cell.Value = "region1" Then
            With Application.ReplaceFormat.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 8
                .ColorIndex = 10
            End With
            cell.Replace What:="region1", replacement:="region11", _
                         ReplaceFormat:=True, LookAt:=xlWhole, MatchCase:=False
        ElseIf cell.Value = "region2" Then
            With Application.ReplaceFormat.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 18
                .ColorIndex = 3
            End With
            cell.Replace What:="region2", replacement:="region22", _
                         ReplaceFormat:=True, LookAt:=xlWhole, MatchCase:=False
        End If
    Next
End Sub


