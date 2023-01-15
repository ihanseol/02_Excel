Attribute VB_Name = "Module_Find"

'https://powerspreadsheets.com/excel-vba-find/#VBA-Code-to-Find-(Cell-with)-Value-in-Cell-Range
'https://excelmacromastery.com/excel-vba-find/



Function FindValueInCellRange(MyRange As Range, MyValue As Variant) As String
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 2 arguments: MyRange and MyValue
        '(2) Finds a value passed as argument (MyValue) in a cell range passed as argument (MyRange)
        '(3) Returns the address (as an A1-style relative reference) of the first cell in the cell range (MyRange) where the value (MyValue) is found
     
    With MyRange
        FindValueInCellRange = .Find(What:=MyValue, After:=.Cells(.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End With
     
End Function


Function FindValueInTable(MyWorksheetName As String, MyValue As Variant, Optional MyTableIndex As Long = 1) As String
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 3 arguments: MyWorksheetName, MyValue and MyTableIndex
        '(2) Finds a value passed as argument (MyValue) in an Excel Table stored in a worksheet whose name is passed as argument (MyWorksheetName). The index number of the Excel Table is either:
            '(1) Passed as an argument (MyTableIndex); or
            '(2) Assumed to be 1 (if MyTableIndex is omitted)
        '(3) Returns the address (as an A1-style relative reference) of the first cell in the Excel Table (stored in the MyWorksheetName worksheet and whose index is MyTableIndex) where the value (MyValue) is found
     
    With ThisWorkbook.Worksheets(MyWorksheetName).ListObjects(MyTableIndex).DataBodyRange
        FindValueInTable = .Find(What:=MyValue, After:=.Cells(.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End With
     
End Function


Function FindValueInColumn(MyColumn As Range, MyValue As Variant) As String
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 2 arguments: MyColumn and MyValue
        '(2) Finds a value passed as argument (MyValue) in a column passed as argument (MyColumn)
        '(3) Returns the address (as an A1-style relative reference) of the first cell in the column (MyColumn) where the value (MyValue) is found
     
    With MyColumn
        FindValueInColumn = .Find(What:=MyValue, After:=.Cells(.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End With
     
End Function


Function FindValueInTableColumn(MyWorksheetName As String, MyColumnIndex As Long, MyValue As Variant, Optional MyTableIndex As Long = 1) As String
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 4 arguments: MyWorksheetName, MyColumnIndex, MyValue and MyTableIndex
        '(2) Finds a value passed as argument (MyValue) in an Excel Table column, where:
            '(1) The table column's index is passed as argument (MyColumnIndex); and
            '(2) The Excel Table is stored in a worksheet whose name is passed as argument (MyWorksheetName). The index number of the Excel Table is either:
                '(1) Passed as an argument (MyTableIndex); or
                '(2) Assumed to be 1 (if MyTableIndex is omitted)
        '(3) Returns the address (as an A1-style relative reference) of the first cell in the applicable Excel Table column where the value (MyValue) is found
     
    With ThisWorkbook.Worksheets(MyWorksheetName).ListObjects(MyTableIndex).ListColumns(MyColumnIndex).DataBodyRange
        FindValueInTableColumn = .Find(What:=MyValue, After:=.Cells(.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End With
     
End Function

Function FindMinimumValueInCellRange(MyRange As Range) As Double
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 1 argument: MyRange
        '(2) Finds the minimum value in the cell range passed as argument (MyRange)
     
    FindMinimumValueInCellRange = Application.Min(MyRange)
 
End Function


Function FindStringInCellRange(MyRange As Range, MyString As Variant) As String
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 2 arguments: MyRange and MyString
        '(2) Finds a string passed as argument (MyString) in a cell range passed as argument (MyRange). The search is case-insensitive
        '(3) Returns the address (as an A1-style relative reference) of the first cell in the cell range (MyRange) where the string (MyString) is found
     
    With MyRange
        FindStringInCellRange = .Find(What:=MyString, After:=.Cells(.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End With
     
End Function

Function FindStringInColumn(MyColumn As Range, MyString As Variant) As String
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 2 arguments: MyColumn and MyString
        '(2) Finds a string passed as argument (MyString) in a column passed as argument (MyColumn). The search is case-insensitive
        '(3) Returns the address (as an A1-style relative reference) of the first cell in the column (MyColumn) where the string (MyString) is found
     
    With MyColumn
        FindStringInColumn = .Find(What:=MyString, After:=.Cells(.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End With
     
End Function


Function FindStringInCell(MyCell As Range, MyString As Variant, Optional MyStartingPosition As Variant = 1) As Variant
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts three arguments: MyCell, MyString and MyStartingPosition (optional argument with a default value of 1)
        '(2) Finds a string passed as argument (MyString) in the value/string stored in a cell passed as argument (MyCell)
        '(3) Returns the following (as applicable):
            'If MyString is not found in the value/string stored in MyCell: The string "String not found in cell"
            'If MyString is found in the value/string stored in MyCell: The position of the first occurrence of MyString in the value/string stored in MyCell
     
    'Obtain position of first occurrence of MyString in the value/string stored in MyCell
    FindStringInCell = InStr(MyStartingPosition, MyCell.Value, MyString, vbBinaryCompare)
     
    'If MyString is not found in the value/string stored in MyCell, return the string "String not found in cell"
    If FindStringInCell = 0 Then FindStringInCell = "String not found in cell"
 
End Function

Function FindTextInString(MyString As Variant, MyText As Variant, Optional MyStartingPosition As Variant = 1) As Variant
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts three arguments: MyString, MyText and MyStartingPosition (optional argument with a default value of 1)
        '(2) Finds text (a string) passed as argument (MyText) in a string passed as argument (MyString)
        '(3) Returns the following (as applicable):
            'If MyText is not found in MyString: The string "Text not found in string"
            'If MyText is found in MyString: The position of the first occurrence of MyText in MyString
     
    'Obtain position of first occurrence of MyText in MyString
    FindTextInString = InStr(MyStartingPosition, MyString, MyText, vbBinaryCompare)
     
    'If MyText is not found in MyString, return the string "Text not found in string"
    If FindTextInString = 0 Then FindTextInString = "Text not found in string"
 
End Function

Function FindCharacterInString(MyString As Variant, MyCharacter As Variant, Optional MyStartingPosition As Variant = 1) As Variant
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts three arguments: MyString, MyCharacter and MyStartingPosition (optional argument with a default value of 1)
        '(2) Finds a character passed as argument (MyCharacter) in a string passed as argument (MyString)
        '(3) Returns the following (as applicable):
            'If MyCharacter is not found in MyString: The string "Character not found in string"
            'If MyCharacter is found in MyString: The position of the first occurrence of MyCharacter in MyString
     
    'Obtain position of first occurrence of MyCharacter in MyString
    FindCharacterInString = InStr(MyStartingPosition, MyString, MyCharacter, vbBinaryCompare)
     
    'If MyCharacter is not found in MyString, return the string "Character not found in string"
    If FindCharacterInString = 0 Then FindCharacterInString = "Character not found in string"
 
End Function

Function FindColumnWithSpecificHeader(MyRange As Range, MyHeader As Variant) As Long
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 2 arguments: MyRange and MyHeader
        '(2) Finds a header passed as argument (MyHeader) in the first row (the header row) of a cell range passed as argument (MyRange). The search is case-insensitive
        '(3) Returns the number of the column containing the first cell in the header row where the header (MyHeader) is found
     
    With MyRange.Rows(1)
        FindColumnWithSpecificHeader = .Find(What:=MyHeader, After:=.Cells(.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
    End With
 
End Function

Sub FindNextAll()
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This procedure:
        '(1) Finds all cells whose value is 10 in cells A6 to H30 of the "Find Next All" worksheet in this workbook
        '(2) Sets the found cells' interior/fill color to light green
     
    'Declare variable to hold/represent searched value
    Dim MyValue As Long
     
    'Declare variable to hold/represent address of first cell where searched value is found
    Dim FirstFoundCellAddress As String
     
    'Declare object variable to hold/represent cell range where search takes place
    Dim MyRange As Range
     
    'Declare object variable to hold/represent cell where searched value is found
    Dim FoundCell As Range
     
    'Specify searched value
    MyValue = 10
     
    'Identify cell range where search takes place
    Set MyRange = ThisWorkbook.Worksheets("Find Next All").Range("A6:H30")
     
    'Find first cell where searched value is found
    With MyRange
        Set FoundCell = .Find(What:=MyValue, After:=.Cells(.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    End With
     
    'Test whether searched value is found in cell range where search takes place
    If Not FoundCell Is Nothing Then
         
        'Store address of first cell where searched value is found
        FirstFoundCellAddress = FoundCell.Address
         
        Do
             
            'Set interior/fill color of applicable cell where searched value is found to light green
            FoundCell.Interior.Color = RGB(63, 189, 133)
             
            'Find next cell where searched value is found
            Set FoundCell = MyRange.FindNext(After:=FoundCell)
         
        'Loop until address of current cell where searched value is found is equal to address of first cell where searched value was found
        Loop Until FoundCell.Address = FirstFoundCellAddress
     
    End If
     
End Sub




Function FindLastRow(MyRange As Range) As Long
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 1 argument: MyRange
        '(2) Tests whether MyRange is empty
        '(3) If MyRange is empty, returns 0 as the number of the last row with data in MyRange
        '(4) If MyRange is not empty:
            '(1) Finds the last row with data in MyRange by searching for the last cell with any character sequence
            '(2) Returns the number of the last row with data in MyRange
         
    'Test if MyRange is empty
    If Application.CountA(MyRange) = 0 Then
 
        'If MyRange is empty, assign 0 to FindLastRow
        FindLastRow = 0
 
    Else
     
        'If MyRange isn't empty, find the last cell with any character sequence by:
            '(1) Searching for the previous match;
            '(2) Across rows
        FindLastRow = MyRange.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
 
    End If
     
End Function

Function FindLastColumn(MyWorksheetName As String) As Long
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 1 argument: MyWorksheetName
        '(2) Tests whether the worksheet named MyWorksheetName is empty
        '(3) If the worksheet named MyWorksheetName is empty, returns 0 as the number of the last column with data in the worksheet
        '(4) If the worksheet named MyWorksheetName is not empty:
            '(1) Finds the last column with data in the worksheet by searching for the last cell with any character sequence
            '(2) Returns the number of the last column with data in the worksheet
     
    'Declare object variable to hold/represent all cells in the worksheet named MyWorksheetName
    Dim MyRange As Range
     
    'Identify all cells in the worksheet named MyWorksheetName
    Set MyRange = ThisWorkbook.Worksheets(MyWorksheetName).Cells
     
    'Test if MyRange is empty
    If Application.CountA(MyRange) = 0 Then
 
        'If MyRange is empty, assign 0 to FindLastColumn
        FindLastColumn = 0
 
    Else
     
        'If MyRange isn't empty, find the last cell with any character sequence by:
            '(1) Searching for the previous match;
            '(2) Across columns
        FindLastColumn = MyRange.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
 
    End If
     
End Function

Function FindLastNonEmptyCellColumn(MyColumn As String) As String
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 1 argument: MyColumn
        '(2) Finds the last non empty cell in the column whose letter is passed as argument (MyColumn) in the worksheet where the UDF is used
        '(3) Returns the address (as an A1-style relative reference) of the last non empty cell found in the column whose letter is passed as argument (MyColumn) in the worksheet where the UDF is used
    With Application.Caller.Parent
        FindLastNonEmptyCellColumn = .Range(MyColumn & .Rows.Count).End(xlUp).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End With
 
End Function

Sub FindBlankCells()
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This procedure:
        '(1) Finds all blank cells in cells A6 to H30 of the "Find Blank Cells" worksheet in this workbook
        '(2) Sets the found cells' interior/fill color to light green
     
    'Declare object variable to hold/represent blank cells
    Dim MyBlankCells As Range
         
    'Enable error-handling
    On Error Resume Next
     
    'Identify blank cells in searched cell range
    Set MyBlankCells = ThisWorkbook.Worksheets("Find Blank Cells").Range("A6:H30").SpecialCells(xlCellTypeBlanks)
     
    'Disable error-handling
    On Error GoTo 0
     
    'Test whether blank cells were found in searched cell range
    If Not MyBlankCells Is Nothing Then
         
        'Set interior/fill color of blank cells found to light green
        MyBlankCells.Interior.Color = RGB(63, 189, 133)
         
    End If
 
End Sub


Function FindNextEmptyCellRange(MyRange As Range) As String
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 1 argument: MyRange
        '(2) Loops through each individual cell in the cell range (MyRange) and tests whether the applicable cell is empty
        '(3) Returns the following (as applicable):
            'If there are empty cells in MyRange: The address (as an A1-style relative reference) of the first empty cell
            'If there are no empty cells in MyRange: The string "No empty cells found in cell range"
     
    'Declare object variable to iterate/loop through all cells in the cell range (MyRange)
    Dim iCell As Range
     
    'Loop through each cell in the cell range (MyRange)
    For Each iCell In MyRange
         
        'If the current cell is empty:
            '(1) Return the current cell's address (as an A1-style relative reference)
            '(2) Exit the For Each... Next loop
        If IsEmpty(iCell) Then
            FindNextEmptyCellRange = iCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            Exit For
        End If
         
    Next iCell
     
    'If no empty cells are found in the cell range (MyRange), return the string "No empty cells found in cell range"
    If FindNextEmptyCellRange = "" Then FindNextEmptyCellRange = "No empty cells found in cell range"
 
End Function


Function FindNextEmptyCellColumn(MySourceCell As Range) As String
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 1 argument: MySourceCell
        '(2) Finds the next empty/blank cell in the applicable column and after the cell passed as argument (MySourceCell)
        '(3) Returns the address (as an R1C1-style absolute reference) of the next empty/blank cell in the applicable column and after the cell passed as argument (MySourceCell)
     
    With MySourceCell(1)
         
        'If first cell in cell range passed as argument (MySourceCell) is empty, obtain/return that cell's address
        If IsEmpty(MySourceCell(1)) Then
            FindNextEmptyCellColumn = .Address(ReferenceStyle:=xlR1C1)
         
        'If cell below first cell in cell range passed as argument (MySourceCell) is empty, obtain/return that cell's address
        ElseIf IsEmpty(.Offset(1, 0)) Then
            FindNextEmptyCellColumn = .Offset(1, 0).Address(ReferenceStyle:=xlR1C1)
         
        'Otherwise:
            '(1) Find the next empty/blank cell in the applicable column and after the first cell in cell range passed as argument (MySourceCell)
            '(2) Obtain/return the applicable cell's address
        Else
            FindNextEmptyCellColumn = .End(xlDown).Offset(1, 0).Address(ReferenceStyle:=xlR1C1)
        End If
         
    End With
 
End Function


Function FindValueInArray(MyArray As Variant, MyValue As Variant) As Variant
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 2 arguments: MyArray and MyValue
        '(2) Loops through each element in the array (MyArray) and tests whether the applicable element is equal to the searched value (MyValue)
        '(3) Returns the following (as applicable):
            'If MyValue is not found in MyArray: The string "Value not found in array"
            'If MyValue is found in MyArray: The position (index number) of MyValue in MyArray
     
    'Declare variable to represent loop counter
    Dim iCounter As Long
 
    'Specify default/fallback value/string to be returned ("Value not found in array") if MyValue is not found in MyArray
    FindValueInArray = "Value not found in array"
 
    'Loop through each element in the array (MyArray)
    For iCounter = LBound(MyArray) To UBound(MyArray)
     
        'Test if the current array element is equal to MyValue
        If MyArray(iCounter) = MyValue Then
             
            'If the current array element is equal to MyValue:
                '(1) Return the position (index number) of the current array element
                '(2) Exit the For... Next loop
            FindValueInArray = iCounter
            Exit For
             
        End If
         
    Next iCounter
 
End Function


Sub FindValueInArrayCaller()
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This procedure:
        '(1) Declares an array and fills it with values (0, 10, 20, ..., 90)
        '(2) Calls the FindValueInArray UDF to find the number "50" in the array
        '(3) Displays a message box with the value/string returned by the FindValueInArray UDF
     
    'Declare a zero-based array with 10 elements of the Long data type
    Dim MySearchedArray(0 To 9) As Long
     
    'Assign values to array elements
    MySearchedArray(0) = 0
    MySearchedArray(1) = 10
    MySearchedArray(2) = 20
    MySearchedArray(3) = 30
    MySearchedArray(4) = 40
    MySearchedArray(5) = 50
    MySearchedArray(6) = 60
    MySearchedArray(7) = 70
    MySearchedArray(8) = 80
    MySearchedArray(9) = 90
     
    'Do the following:
        '(1) Call the FindValueInArray UDF and search for the number "50" in MySearchedArray
        '(2) Display a message box with the value/string returned by the FindValueInArray UDF
    MsgBox FindValueInArray(MySearchedArray, 50)
     
 
End Sub
























