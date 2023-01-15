Attribute VB_Name = "Chapter13"
Sub ExcelFileSearch()
    'provide the file extension and the folder, VBA programs returns the file name, memory size, and last modified
    Dim srchExt     As Variant, srchDir As Variant
    Dim i           As Long, j As Long, strName As String
    Dim varArr(1 To 1048576, 1 To 3) As Variant
    Dim strFileFullName As String
    Dim ws          As Worksheet
    Dim fso         As Object
    Let srchExt = Application.InputBox("Please Enter File Extension", _
                                       "Info Request")
    If srchExt = False And Not TypeName(srchExt) = "String" Then
        Exit Sub
    End If
    Let srchDir = BrowseForFolderShell
    If srchDir = False And Not TypeName(srchDir) = "String" Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets.Add(Sheets(1))
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("FileSearch Results").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    ws.Name = "FileSearch Results"
    Let strName = Dir$(srchDir & "\*" & srchExt)
    Do While strName <> vbNullString
        Let i = i + 1
        Let strFileFullName = srchDir & strName
        Let varArr(i, 1) = strFileFullName
        Let varArr(i, 2) = FileLen(strFileFullName) \ 1024
        Let varArr(i, 3) = FileDateTime(strFileFullName)
        Let strName = Dir$()
    Loop
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call recurseSubFolders(fso.GetFolder(srchDir), varArr(), i, CStr(srchExt))
    Set fso = Nothing
    ThisWorkbook.Windows(1).DisplayHeadings = False
    With ws
        If i > 0 Then
            .Range("A2").Resize(i, UBound(varArr, 2)).Value = varArr
            For j = 1 To i
                .Hyperlinks.Add anchor:=.Cells(j + 1, 1), Address:=varArr(j, 1)
            Next
        End If
        .Range(.Cells(1, 4), .Cells(1, .Columns.Count)).EntireColumn.Hidden = _
                                                                            True
        .Range(.Cells(.Rows.Count, 1).End(xlUp)(2), _
               .Cells(.Rows.Count, 1)).EntireRow.Hidden = True
        With .Range("A1:C1")
            .Value = Array("Full Name", "Kilobytes", "Last Modified")
            .Font.Underline = xlUnderlineStyleSingle
            .EntireColumn.AutoFit
            .HorizontalAlignment = xlCenter
        End With
    End With
    Application.ScreenUpdating = True
End Sub

Private Sub recurseSubFolders(ByRef Folder As Object, _
                              ByRef varArr() As Variant, _
                              ByRef i As Long, _
                              ByRef srchExt As String)
    Dim SubFolder   As Object
    Dim strName     As String, strFileFullName As String
    For Each SubFolder In Folder.SubFolders
        Let strName = Dir$(SubFolder.Path & "\*" & srchExt)
        Do While strName <> vbNullString
            Let i = i + 1
            Let strFileFullName = SubFolder.Path & "\" & strName
            Let varArr(i, 1) = strFileFullName
            Let varArr(i, 2) = FileLen(strFileFullName) \ 1024
            Let varArr(i, 3) = FileDateTime(strFileFullName)
            Let strName = Dir$()
        Loop
        If i > 1048576 Then Exit Sub
        Call recurseSubFolders(SubFolder, varArr(), i, srchExt)
    Next
End Sub

Private Function BrowseForFolderShell() As Variant
    Dim objShell    As Object, objFolder As Object
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "Please Select a folder", _
                                             0, "C:\")
    If Not objFolder Is Nothing Then
        On Error Resume Next
        If IsError(objFolder.Items.Item.Path) Then
            BrowseForFolderShell = CStr(objFolder)
        Else
            On Error GoTo 0
            If Len(objFolder.Items.Item.Path) > 3 Then
                BrowseForFolderShell = objFolder.Items.Item.Path & _
                                       Application.PathSeparator
            Else
                BrowseForFolderShell = objFolder.Items.Item.Path
            End If
        End If
    Else
        BrowseForFolderShell = False
    End If
    Set objFolder = Nothing: Set objShell = Nothing
End Function

'Option Base 1
Sub OpenLargeCSVFast()
    Dim buf(1 To 16384) As Variant
    Dim i           As Long
    'Change the file location and name here
    Const strFilePath As String = "C:\temp\Sales.CSV"
    Dim strRenamedPath As String
    strRenamedPath = Split(strFilePath, ".")(0) & "txt"
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    'Setting an array for FieldInfo to open CSV
    For i = 1 To 16384
        buf(i) = Array(i, 2)
    Next
    Name strFilePath As strRenamedPath
    Workbooks.OpenText Filename:=strRenamedPath, DataType:=xlDelimited, _
                       Comma:=True, FieldInfo:=buf
    Erase buf
    ActiveSheet.UsedRange.Copy ThisWorkbook.Sheets(1).Range("A1")
    ActiveWorkbook.Close False
    Kill strRenamedPath
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
End Sub

Sub SplitWorkbook()
    Dim ws          As Worksheet
    Dim DisplayStatusBar As Boolean
    DisplayStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Sheets
        Dim NewFileName As String
        Application.StatusBar = ThisWorkbook.Sheets.Count & _
                                " Remaining Sheets"
        If ThisWorkbook.Sheets.Count <> 1 Then
            NewFileName = ThisWorkbook.Path & "\" & ws.Name & ".xlsm" _
                          'Macro-Enabled
            ' NewFileName = ThisWorkbook.Path & "\" & ws.Name & ".xlsx" _
            'Not Macro-Enabled
            ws.Copy
            ActiveWorkbook.Sheets(1).Name = "Sheet1"
            ActiveWorkbook.SaveAs Filename:=NewFileName, _
                                  FileFormat:=xlOpenXMLWorkbookMacroEnabled
            ' ActiveWorkbook.SaveAs Filename:=NewFileName, _
            FileFormat:=xlOpenXMLWorkbook
            ActiveWorkbook.Close SaveChanges:=False
        Else
            NewFileName = ThisWorkbook.Path & "\" & ws.Name & ".xlsm"
            ' NewFileName = ThisWorkbook.Path & "\" & ws.Name & ".xlsx"
            ws.Name = "Sheet1"
        End If
    Next
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.DisplayStatusBar = DisplayStatusBar
    Application.ScreenUpdating = True
End Sub

Sub CombineWorkbooks()
    Dim CurFile     As String, DirLoc As String
    Dim DestWB      As Workbook
    Dim ws          As Object                    'allows for different sheet types
    'DirLoc = ThisWorkbook.Path & "\tst\" 'location of files
    'CurFile = Dir(DirLoc & "*.xls*")
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set DestWB = Workbooks.Add(xlWorksheet)
    Do While CurFile <> vbNullString
        Dim OrigWB  As Workbook
        Set OrigWB = Workbooks.Open(Filename:=DirLoc & CurFile, _
                                    ReadOnly:=True)
        ' Limit to valid sheet names and removes .xls*
        CurFile = Left(Left(CurFile, Len(CurFile) - 5), 29)
        For Each ws In OrigWB.Sheets
            ws.Copy After:=DestWB.Sheets(DestWB.Sheets.Count)
            If OrigWB.Sheets.Count > 1 Then
                DestWB.Sheets(DestWB.Sheets.Count).Name = CurFile & ws.Index
            Else
                DestWB.Sheets(DestWB.Sheets.Count).Name = CurFile
            End If
        Next
        OrigWB.Close SaveChanges:=False
        CurFile = Dir
    Loop
    Application.DisplayAlerts = False
    DestWB.Sheets(1).Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Set DestWB = Nothing
End Sub

Sub Filter_NewSheet()
    Dim wbBook      As Workbook
    Dim wsSheet     As Worksheet
    Dim rnStart     As Range, rnData As Range
    Dim i           As Long
    Set wbBook = ThisWorkbook
    Set wsSheet = wbBook.Worksheets("Sheet1")
    With wsSheet
        'Make sure that the first row contains headings.
        Set rnStart = .Range("A2")
        Set rnData = .Range(.Range("A2"), .Cells(.Rows.Count, 3).End(xlUp))
    End With
    Application.ScreenUpdating = True
    For i = 1 To 5
        'Here we filter the data with the first criterion.
        rnStart.AutoFilter Field:=1, Criteria1:="AA" & i
        'Copy the filtered list
        rnData.SpecialCells(xlCellTypeVisible).Copy
        'Add a new worksheet to the active workbook.
        Worksheets.Add Before:=wsSheet
        'Name the added new worksheets.
        ActiveSheet.Name = "AA" & i
        'Paste the filtered list.
        Range("A2").PasteSpecial xlPasteValues
    Next i
    'Reset the list to its original status.
    rnStart.AutoFilter Field:=1
    With Application
        'Reset the clipboard.
        .CutCopyMode = False
        .ScreenUpdating = False
    End With
End Sub

Sub CriteriaRange_Copy()
    Dim Table       As ListObject
    Dim SortColumn  As ListColumn
    Dim CriteriaColumn As ListColumn
    Dim FoundRange  As Range
    Dim TargetSheet As Worksheet
    Dim HeaderVisible As Boolean
    Set Table = ActiveSheet.ListObjects(1)       ' Set as desired
    HeaderVisible = Table.ShowHeaders
    Table.ShowHeaders = True
    On Error GoTo RemoveColumns
    Set SortColumn = Table.ListColumns.Add(Table.ListColumns.Count + 1)
    Set CriteriaColumn = Table.ListColumns.Add _
                         (Table.ListColumns.Count + 1)
    On Error GoTo 0
    'Add a column to keep track of the original order of the records
    SortColumn.Name = " Sort"
    CriteriaColumn.Name = " Criteria"
    SortColumn.DataBodyRange.Formula = "=ROW(A1)"
    SortColumn.DataBodyRange.Value = SortColumn.DataBodyRange.Value
    'add the formula to mark the desired records
    'the records not wanted will have errors
    CriteriaColumn.DataBodyRange.Formula = "=1/(([@Units]<10)*([@Cost]<5))"
    CriteriaColumn.DataBodyRange.Value = CriteriaColumn.DataBodyRange.Value
    Table.Range.Sort Key1:=CriteriaColumn.Range(1, 1), _
        Order1:=xlAscending, Header:=xlYes
    On Error Resume Next
    Set FoundRange = Intersect(Table.Range, CriteriaColumn.DataBodyRange. _
                                           SpecialCells(xlCellTypeConstants, xlNumbers).EntireRow)
    On Error GoTo 0
    If Not FoundRange Is Nothing Then
        Set TargetSheet = ThisWorkbook.Worksheets.Add(After:=ActiveSheet)
        FoundRange(1, 1).Offset(-1, 0).Resize(FoundRange.Rows.Count + 1, _
                                              FoundRange.Columns.Count - 2).Copy
        TargetSheet.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
    End If
    Table.Range.Sort Key1:=SortColumn.Range(1, 1), Order1:=xlAscending, _
        Header:=xlYes
RemoveColumns:
    If Not SortColumn Is Nothing Then SortColumn.Delete
    If Not CriteriaColumn Is Nothing Then CriteriaColumn.Delete
    Table.ShowHeaders = HeaderVisible
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Debug.Print Target

    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("13").Activate
    If Target.Column > 2 Or Target.Cells.Count > 1 Then Exit Sub
    If Application.IsNumber(Target.Value) = False Then
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
        MsgBox "Numbers only please."
        Exit Sub
    End If
    Select Case Target.Column
    Case 1
        If Target.Value > Target.Offset(0, 1).Value Then
            Application.EnableEvents = False
            Application.Undo
            Application.EnableEvents = True
            MsgBox "Value in column A may Not be larger than value " & _
                   "in column B."
            Exit Sub
        End If
    Case 2
        If Target.Value < Target.Offset(0, -1).Value Then
            Application.EnableEvents = False
            Application.Undo
            Application.EnableEvents = True
            MsgBox "Value in column B may Not be smaller " & _
                   "than value in column A."
            Exit Sub
        End If
    End Select
    Dim x           As Long
    x = Target.Row
    Dim z           As String
    z = Range("B" & x).Value - Range("A" & x).Value
    With Range("C" & x)
        .Formula = "=IF(RC[-1]<=RC[-2],REPT(""n"",RC[-1])&" & _
                                                            "REPT(""n"",RC[-2]-RC[-1]),REPT(""n"",RC[-2])&" & _
                                                            "REPT(""o"",RC[-1]-RC[-2]))"
        .Value = .Value
        .Font.Name = "Wingdings"
        .Font.ColorIndex = 1
        .Font.Size = 10
        If Len(Range("A" & x)) <> 0 Then
            .Characters(1, (.Characters.Count - z)).Font.ColorIndex = 3
            .Characters(1, (.Characters.Count - z)).Font.Size = 12
        End If
    End With
End Sub


