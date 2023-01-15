Attribute VB_Name = "BaseData_DataJOjung"
Option Explicit

Dim ColorValue(1 To 20) As Long

Public Sub InitialSetColorValue()
    
    ColorValue(1) = RGB(192, 0, 0)
    ColorValue(2) = RGB(255, 0, 0)
    ColorValue(3) = RGB(255, 192, 0)
    ColorValue(4) = RGB(255, 255, 0)
    ColorValue(5) = RGB(146, 208, 80)
    ColorValue(6) = RGB(0, 176, 80)
    ColorValue(7) = RGB(0, 176, 240)
    ColorValue(8) = RGB(0, 112, 192)
    ColorValue(9) = RGB(0, 32, 96)
    ColorValue(10) = RGB(112, 48, 160)
    
    ColorValue(11) = RGB(192 + 10, 10, 0)
    ColorValue(12) = RGB(255, 0 + 10, 0)
    ColorValue(13) = RGB(255, 192 + 10, 0)
    ColorValue(14) = RGB(255, 255, 10)
    ColorValue(15) = RGB(146 + 10, 208 + 10, 80 + 10)
    ColorValue(16) = RGB(0 + 10, 176 + 10, 80)
    ColorValue(17) = RGB(0 + 10, 176 + 10, 240 + 10)
    ColorValue(18) = RGB(0 + 10, 112 + 10, 192)
    ColorValue(19) = RGB(0 + 10, 32 + 10, 96)
    ColorValue(20) = RGB(112, 48 + 10, 160 + 10)
    
End Sub

Private Sub initialize_wellstyle()
    
    Range("C3:C22").Select
    Selection.NumberFormat = "General"
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    
    With Selection.Font
        .Name = "∏º¿∫ ∞ÌµÒ"
        .Size = 10
        .ThemeColor = xlThemeColorLight1
    End With
    
    Range("E19:G19").Select
    With Selection.Font
        .Name = "∏º¿∫ ∞ÌµÒ"
        .Size = 12
        .ThemeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("E21:G21").Select
    With Selection.Font
        .Name = "∏º¿∫ ∞ÌµÒ"
        .Size = 12
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("B25:K29").Select
    With Selection.Font
        .Name = "∏º¿∫ ∞ÌµÒ"
        .Size = 11
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("d23").Select
    
End Sub

Private Sub change_font_size()
    
    Range("J25").Select
    Selection.Font.Size = 10
    Range("F26").Select
    Selection.Font.Size = 10
    
End Sub

Public Sub make_wellstyle()
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    n_sheets = sheets_count()
    
    Call TurnOffStuff
    
    For i = 1 To n_sheets
        Worksheets(CStr(i)).Activate
        Call initialize_wellstyle
        Call change_font_size
    Next i
    
    Call TurnOnStuff
    
End Sub

Private Sub JojungData(ByVal nsheet As Integer)
    
    Dim nselect     As String
    
    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
    Range("F21").Activate
    
    nsheet = nsheet + 3
    '=Well!D7
    nselect = Mid(Range("c2").Formula, 8)
    
    'Debug.Print Mid(Range("c2").Formula, 8) & ":" & nselect
    
    Selection.Replace What:=nselect, Replacement:=CStr(nsheet), LookAt:=xlPart, _
                      SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                      ReplaceFormat:=False
    
    Range("E21").Select
    Range("E21").Formula = "=Well!" & Cells(nsheet, "I").Address
    
End Sub

Private Sub SetMyTabColor(ByVal index As Integer)
    
    If Sheets("Well").SingleColor.value Then
        With ActiveWorkbook.Sheets(CStr(index)).Tab
            .Color = 192
            .TintAndShade = 0
        End With
    Else
        
        With ActiveWorkbook.Sheets(CStr(index)).Tab
            .Color = ColorValue(index)
            .TintAndShade = 0
        End With
        
    End If
    
End Sub

'∞¢∞¢¿« Ω¨∆Æ∏¶ º¯»∏«œ∏Èº≠, ºø¿« ¬¸¡∂∞™¿ª ê¨√ﬂæÓ¡ÿ¥Ÿ.
'
Public Sub JojungSheetData()
    
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        Cells(i + 3, "A").value = "W" & i
    Next i
    
    For i = 1 To n_sheets
        Sheets(CStr(i)).Activate
        Call JojungData(i)
        Call SetMyTabColor(i)
    Next i
    
End Sub

