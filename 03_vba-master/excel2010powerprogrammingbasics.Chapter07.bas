Attribute VB_Name = "Chapter07"
Sub test()
    'Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Activate
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("7").Activate
    'Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("7").Range("B1").Select
    Sum = 1 + 1
    Range("A1").Value = Sum                      'print 2
    Range("A2").Value = addtwo(2, 2)             'print 4
    Range("A4").Value = "Hello"
    Range("A4").ClearContents
    Range("A17").Value = "break a lengthy instruction into two or more lines. end line with a space" _
                       & " followed by an underscore character and press Enter and continue on the next line."
    Message = "Is you name " & Application.UserName & "?" 'print Is your name Mar?
    answer = MsgBox(Message, vbYesNo)
    Range("B18").Value = answer                  'print 6 for Yes or 7 for No
    If answer = vbNo Then
        Range("A18").Value = "Oh, never mind"
    Else
        Range("A18").Value = "I must be clairvoyant"
    End If
End Sub

Function addtwo(a, b)
    addtwo = a + b
End Function

Sub formatcells()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("7").Activate
    Range("A5:A9").Select                        'Cells A5:A9 Highlighted
    Selection.Style = "Comma"
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    Range("A6").Value = "test cell A6"
    'better using With-End With statement
    Range("B5:B9").Select                        'Cells B5:B9 Highlighted
    With Selection
        .Style = "comma"
        .Font.Bold = True
        .Font.Italic = True
    End With
    Range("B7").Value = "test cell B7"
    'better using With-End With statement
    With Range("C5:C9")
        .Style = "Comma"
        .Font.Bold = True
        .Font.Italic = True
    End With
    Range("C8").Value = "test cell C8"
End Sub

Sub applicationhierarchyp167()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets(1).Activate
    Range("E11").Value = 5
    Range("E5").Value = "Clear me here"
    Range("E6").Value = "Clear me here everything "
    Range("E1").Value = "Examples of Application object hierarchy: Application.Workbooks, " _
                      & "Application.Windows, Application.AddIns, Workbooks.Worksheets, Workbooks.Charts, " _
                      & "Workbooks.Names"
    Range("E2").Value = "Range object has a property called Value page 169"
    Range("E3") = "Range object default property is Value Range(""E3"") = ... page 170"
    Range("E4").Value = "Range object has a method which is an action.  e.g. Range(""E5"").Clear"
    Range("E5").Clear                            'Clear cell contenets
    Range("E6").ClearContents                    'Clear everything
    Range("A1").copy Range("E8")                 'Copy A1 to E8
    Range("E8").AddComment ("Read the comment")
    Range("E8").Comment.Visible = True           'Show comment
    Range("E8").Comment.Shape.Fill.ForeColor.RGB = RGB(0, 255, 0) 'Change background color
    Range("E8").Comment.Shape.TextFrame.Characters.Font.ColorIndex = 5 'Change font color
    Range("E8").Comment.Visible = False          'Hide comment
    Range("E8").Comment.Delete                   'Delete comment
    Range("E9").Value = "check delete"
    Range("E9").Delete                           'Note delete moves cells up
    Range("E10").Clear
End Sub

Sub activesp180()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("7").Activate
    'Place cursor at cell H6
    ActiveCell.Value = "I'm at cell h6"          'print I'm at cell h6 at H6
    Range("H7").Value = ActiveSheet.Name         'print 7
    Range("H8").Value = ActiveWorkbook.Name      'print excel2010powerprogrammingbasics.xlsm
    Range("H9:H11").Select
    Selection.Value = "Copy all"                 'print Copy All cells H9:H11
    Range("H12").Value = ActiveWorkbook.FullName 'print G:\Raymond\Excel Files 2GB Backup 010718\ _
                                                 & VBA Macros Round Two\excel2010powerprogrammingbasics.xlsm
    ActiveCell.Offset(0, 1).Select
End Sub

Sub rangep182()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("7").Activate
    Worksheets("7").Range("A25").Value = 12.3
    Range("input").Value = "B25 is named input"  'I named cell b25 as input
    'To refer to merged cells, you can reference the entire merged range or just the upper-left cell
    'within the merged range. For example, if a worksheet contains four cells merged into one (A1,
    'B1, A2, and B1), reference the merged cells using either of the following expressions:
    'Range(“A1:B2?
    'Range (“A1?
    Range("A26:E26, G26:G29").Value = 3
    Range("A27, A29, A31") = 4
    Cells(34, 1).Value = "Row 34, Column 1"
End Sub


