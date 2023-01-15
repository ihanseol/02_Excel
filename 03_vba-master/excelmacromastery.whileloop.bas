Attribute VB_Name = "whileloop"
Sub whileloop()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("whileloop").Activate
    'Quick guide
    'Do While . . . Loop runs 0 or more times while true.  Do While result="Correct" . . . Loop
    'Do . . . Loop While runs 1 or more times while true.  Do . . . Loop While result="Correct"
    'Do Until . . . Loop runs 0 or more times until true.  Do Until result <> "Correct" Loop
    'Do . . . Until Loop runs 1 or more times until true.  Do . . . Loop Until result <> "Correct"
    'Exit the Do Loop
    'Do While i < 10 i = GetTotal if i < 0 then Exit Do End If Loop
    i = 1
    Do While i < 11
        Range("A" & i).Value = i
        i = i + 1
    Loop
    i = 1
    Do Until i = 11
        Range("B" & i) = i
        i = i + 1
    Loop
    i = 1
    Do
        Range("C" & i) = i
        i = i + 1
    Loop While i < 11
    i = 1
    Do
        Range("D" & i) = i
        i = i + 1
    Loop Until i = 11

    Dim item As String
    Range("A12").Select
    Do
        item = InputBox("Please Enter Item.  Type nothing and press OK to quit.")
        ActiveCell.Value = item
        ActiveCell.offset(1, 0).Select
    Loop While item <> ""
    Range("A12").CurrentRegion.Select
    'Range("A12").CurrentRegion.ClearContents
    'Range("A12").CurrentRegion.Delete
End Sub

Sub alwaysatleastoneinput()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("whileloop").Activate
    Dim oneplusinput As String
    
    Range("B12").Select
    Do While oneplusinput <> "q"
        oneplusinput = InputBox("Enter an item.  Press q to quit")
        ActiveCell.Value = oneplusinput
        ActiveCell.offset(1, 0).Select
    Loop
    
    Range("B12").CurrentRegion.ClearContents
    Range("C12").Select
    Do
        oneplusinput = InputBox("Enter an item.  Press q to quit")
        ActiveCell.Value = oneplusinput
        ActiveCell.offset(1, 0).Select
    Loop While oneplusinput <> "q"
    
    Range("C12").CurrentRegion.Select
End Sub

Sub maybeoneinput()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("whileloop").Activate
    Dim mayinput As String

    'Do While mayinput <> "q" doesn't loop because mayinput equals q
    mayinput = "q"
    Range("D12").Select
    Do While mayinput <> "q"
        mayinput = InputBox("Enter an item.  Press q to quit")
        ActiveCell = mayinput
        ActiveCell.offset(1, 0).Select
    Loop

    'Do Loop While mayinput <> "q" loops because mayinput equals q is checked after one loop
    mayinput = "q"
    Range("D12").Select
    Do
        mayinput = InputBox("Enter an item for second maybeoneinput() loop.  Press q to quit")
        ActiveCell = mayinput
        ActiveCell.offset(1, 0).Select
    Loop While mayinput <> "q"
    Range("D12").CurrentRegion.Select
End Sub

Sub whileuntil()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("whileloop").Activate
    Dim stopcommand As String
    
    Do Until stopcommand = "q"
        stopcommand = InputBox("Please enter item for do until loop 1")
    Loop
    stopcommand = ""                             'RM:  reset stopcommand variable for loop 2
    
    Do While stopcommand <> "q"
        stopcommand = InputBox("Please enter item for do while loop 2")
    Loop
    
    Do
        stopcommand = InputBox("Please enter item for do loop until 3")
    Loop Until stopcommand = "q"
    
    Do
        stopcommand = InputBox("Please enter item for do loop while 4")
    Loop While stopcommand <> "q"
End Sub

Sub loopsisnull()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("whileloop").Activate
    Range("A1").Select
    
    Do While Not ActiveCell.Value = ""
        MsgBox ActiveCell.Value
        ActiveCell.offset(1, 0).Select
    Loop
    'RM:  Nothing and Not Nothing are used for objects.  Can't do Do While Not ActiveCell.Value Is Nothing
End Sub

Sub exitdo()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("whileloop").Activate
    Range("B1").Select
    
    Do While ActiveCell.Value <> ""
        If ActiveCell.Value = 6 Then
            Exit Do
        Else
            MsgBox ActiveCell.Value
        End If
        ActiveCell.offset(1, 0).Select
    Loop
    MsgBox "Loop Is Exited"
End Sub

Sub exitdorangeobject()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("whileloop").Activate
    Dim startcell As Range
    Set startcell = Range("C1")
    startcell.Select
    
    Do While ActiveCell.Value <> ""
        If ActiveCell.Value = 6 Then
            Exit Do
        Else
            MsgBox ActiveCell.Value
        End If
        ActiveCell.offset(1, 0).Select
    Loop
    MsgBox "Loop Is Exited"
End Sub

Sub longwaysumnumbers()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("whileloop").Activate
    Dim total As Long, startcell As Range
    
    Set startcell = Range("D1")
    startcell.Select
    total = startcell.Value
    
    Do While ActiveCell.Value <> ""
        total = total + ActiveCell.Value
        ActiveCell.offset(1, 0).Select
    Loop
    startcell.End(xlDown).offset(1, 0).Value = total
End Sub

