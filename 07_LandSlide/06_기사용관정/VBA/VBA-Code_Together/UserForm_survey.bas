
Public Enum LC_COMBOBOX
    lcDAEJEON = 1
    lcJIYEOL = 2
End Enum

Public IS_FIRST_LOAD As Boolean

Private Sub OptionButton_DAEJEON_Click()

    If IS_FIRST_LOAD Then
        Call LoadComboBox
        IS_FIRST_LOAD = False
    Else
        Call LoadComboBox
        ComboBox_AREA.Value = "default"
        IS_FIRST_LOAD = False
    End If
    
End Sub


Private Sub OptionButton_JIYEOL_Click()
    
    If IS_FIRST_LOAD Then
        Call LoadComboBox
        IS_FIRST_LOAD = False
    Else
        Call LoadComboBox
        ComboBox_AREA.Value = "default"
        IS_FIRST_LOAD = False
    End If
    
End Sub

Sub PutDataToASheet(ByVal sh As String, ByVal table As String, ByVal area As String, SurveyData As Variant)
    Dim tbl As ListObject
    Dim cell As Range
    Dim i As Integer: i = 1
    
    Set tbl = Sheets(sh).ListObjects(table)
        
    For Each cell In tbl.ListColumns(area).DataBodyRange.Cells
        cell.Value = SurveyData(i)
        i = i + 1
    Next cell
    
End Sub

Function getSurveyData() As Variant
    Dim values() As Variant
    Dim i As Integer: i = 1
    Dim ctl As Control

    ReDim values(1 To 23)
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            values(i) = ctrl.Value
            i = i + 1
        End If
    Next ctrl
    
    getSurveyData = values
    
End Function



Private Sub CommandButton_Insert_Click()

    Dim values As Variant
    Dim area As String
    
    values = getSurveyData()

    area = ComboBox_AREA.Value
    
    If area = "" Then
        area = Default
    End If
    
    
    If OptionButton_JIYEOL.Value Then
        Call PutDataToASheet("ref1", "tableJIYEOL", area, values)
    Else
        Call PutDataToASheet("ref", "tableCNU", area, values)
    End If
    
    Call PutText(area)
    
    Unload Me
    
End Sub


Sub PutText(area As String)
    Dim TextBox_AREA As MSForms.TextBox

    Set TextBox_AREA = TextBoxFind("TextBox_AREA")
    TextBox_AREA.Value = area
End Sub

' ***************************************************************
' UserForm_Survey
'
' ***************************************************************



Private Sub CommandButton_LOAD_Click()
    Call LoadSurveyData(ComboBox_AREA.Value)
End Sub


Private Sub ComboBox_AREA_Change()

End Sub

Sub Initialize_Setting()
    Dim TextBox_AREA As MSForms.TextBox

    Set TextBox_AREA = TextBoxFind("TextBox_AREA")
    Debug.Print "TextBox_AREA.Value", "'" & TextBox_AREA.Value & "'"
    
    If is_Jiyeol(TextBox_AREA.Value) Then
        OptionButton_JIYEOL.Value = True
    Else
        OptionButton_DAEJEON.Value = True
    End If
    
    ' Call LoadComboBox
    ' OptionButton.Value = True set is triggered clicked event
    
    ComboBox_AREA.Value = TextBox_AREA.Value
    LoadSurveyData (TextBox_AREA.Value)
    
End Sub


Sub LoadSurveyData(area As String)
    Dim tbl As ListObject
    Dim values() As Variant
    
    
    If OptionButton_JIYEOL.Value Then
        Set tbl = Sheets("ref1").ListObjects("tableJIYEOL")
    Else
        Set tbl = Sheets("ref").ListObjects("tableCNU")
    End If
    
    
    If area = "" Then
        area = Default
    End If
    
    values = tbl.ListColumns(area).DataBodyRange.Value
    Dim i As Integer: i = 1
        
    Dim ctl As Control

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ' MsgBox "Found a TextBox with the name: " & ctrl.NAME
            ctrl.Value = values(i, 1)
            i = i + 1
        End If
    Next ctrl
    
End Sub


Sub LoadComboBox()
    Dim tbl As ListObject
    Dim tableNAME, shNAME As String
    Dim headerRowArray() As Variant
    
    ComboBox_AREA.Clear
    
    If OptionButton_JIYEOL.Value Then
        tableNAME = "tableJIYEOL"
        shNAME = "ref1"
    Else
        tableNAME = "tableCNU"
        shNAME = "ref"
    End If
    
    Set tbl = Sheets(shNAME).ListObjects(tableNAME)

    headerRowArray = tbl.HeaderRowRange.Value
    
    Dim i As Integer
    Dim isFirst As Boolean: isFirst = True
    
    
    For i = LBound(headerRowArray, 2) To UBound(headerRowArray, 2)
        If isFirst Then
            isFirst = False
            GoTo NEXT_LOOP
        End If
        
        ComboBox_AREA.AddItem headerRowArray(1, i)
        
NEXT_LOOP:
    Next i
End Sub



Function is_Jiyeol(ByVal area As String) As Boolean

    Dim tbl As ListObject
    Dim headerRowArray() As Variant
    
    Set tbl = Sheets("ref1").ListObjects("tableJIYEOL")
    
    headerRowArray = tbl.HeaderRowRange.Value
    
    Dim i As Integer
    
    For i = LBound(headerRowArray, 2) To UBound(headerRowArray, 2)
        If headerRowArray(1, i) = area Then
            is_Jiyeol = True
            Exit Function
        End If
    Next i
    
    is_Jiyeol = False

End Function


Private Sub UserForm_Initialize()
    Dim i As Integer

    IS_FIRST_LOAD = True
    Call Initialize_Setting
    
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub

Private Sub CommandButton_Cancel_Click()
    Unload Me
End Sub



Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub


