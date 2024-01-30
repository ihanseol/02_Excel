Attribute VB_Name = "water_q"
' ***************************************************************
' water_q
'
' ***************************************************************


Public SS(1 To 5, 1 To 2) As Double
Public AA(1 To 6, 1 To 2) As Double

Public SS_CITY As Double
Public ISIT_FIRST As Boolean

Public Enum SS_VALUE
    svGAJUNG = 1
    svILBAN = 2
    svSCHOOL = 3
    svGONGDONG = 4
    svMAEUL = 5
End Enum

Public Enum AA_VALUE
    avJEONJAK = 1
    avDAPJAK = 2
    avWONYE = 3
    avCOW = 4
    avPIG = 5
    avCHICKEN = 6
End Enum

Function CheckBoxFind(objNAME As String) As MSForms.CheckBox
    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim myCheckBox As MSForms.CheckBox
    
    
    Set ws = ThisWorkbook.Worksheets("ss")
    Set myCheckBox = Nothing
    
    For Each obj In ws.OLEObjects
        If TypeOf obj.Object Is MSForms.CheckBox Then
            If obj.Name = objNAME Then
                Set myCheckBox = obj.Object
                Exit For
            End If
        End If
    Next obj
    
    If Not (myCheckBox Is Nothing) Then
        ' found
        Set CheckBoxFind = myCheckBox
    Else
        ' not found
        Set CheckBoxFind = Nothing
    End If
End Function

Function ComboBoxFind(objNAME As String) As MSForms.ComboBox
    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim myComboBox As MSForms.ComboBox
    
    
    Set ws = ThisWorkbook.Worksheets("ss")
    Set myComboBox = Nothing
    
    For Each obj In ws.OLEObjects
        If TypeOf obj.Object Is MSForms.ComboBox Then
            If obj.Name = objNAME Then
                Set myComboBox = obj.Object
                Exit For
            End If
        End If
    Next obj
    
    If Not (myComboBox Is Nothing) Then
        ' found
        Set ComboBoxFind = myComboBox
    Else
        ' not found
        Set ComboBoxFind = Nothing
    End If
End Function


Function TextBoxFind(objNAME As String) As MSForms.TextBox
    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim myTextBox As MSForms.TextBox
    
    
    Set ws = ThisWorkbook.Worksheets("ss")
    Set myTextBox = Nothing
    
    For Each obj In ws.OLEObjects
        If TypeOf obj.Object Is MSForms.TextBox Then
            If obj.Name = objNAME Then
                Set myTextBox = obj.Object
                Exit For
            End If
        End If
    Next obj
    
    If Not (myTextBox Is Nothing) Then
        ' found
        Set TextBoxFind = myTextBox
    Else
        ' not found
        Set TextBoxFind = Nothing
    End If
End Function



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


Sub initialize()
    Dim TextBox_AREA As MSForms.TextBox

    Set TextBox_AREA = TextBoxFind("TextBox_AREA")
        
    If is_Jiyeol(TextBox_AREA.Value) Then
        Call initialize_JIYEOL(TextBox_AREA.Value)
    Else
        Call initialize_CNU(TextBox_AREA.Value)
    End If
       
End Sub


Private Function lastRowByKey(cell As String) As Long
    lastRowByKey = Range(cell).End(xlDown).row
End Function


' 물량계산
Sub ComputeQ()
    Dim i As Integer
    Dim lastRow As Long

    Call initialize
    
    Sheets("ss").Activate
    lastRow = lastRowByKey("A1")
    
    For i = 2 To lastRow
        Cells(i, "L").Value = ss_water(Range("I" & CStr(i)).Value, Range("K" & CStr(i)).Value, 100)
    Next i
    
    Sheets("aa").Activate
    lastRow = lastRowByKey("A1")
    
    For i = 2 To lastRow
        Cells(i, "L").Value = aa_water(Range("I" & CStr(i)).Value, Range("K" & CStr(i)).Value, 100)
    Next i
End Sub


Function ss_water(ByVal qhp As Integer, ByVal strPurpose As String, Optional ByVal npopulation As Integer = 60) As Double
    '지열 냉난방
    If CheckSubstring(strPurpose, "냉") Then
        ss_water = qhp * 0.01
        Exit Function
    End If
    
    ' 일반용
    If CheckSubstring(strPurpose, "일") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    
    ' 가정용
    If CheckSubstring(strPurpose, "가") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    ' 기타
    If CheckSubstring(strPurpose, "기") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    ' 농생활겸용
    If CheckSubstring(strPurpose, "농") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' 청소용
    If CheckSubstring(strPurpose, "청") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    '간이상수도
    If CheckSubstring(strPurpose, "상") Then
        ss_water = Round(SS(svMAEUL, 1) + npopulation * SS(svMAEUL, 2), 2)
        Exit Function
    End If
    
    ' 공사용
    If CheckSubstring(strPurpose, "공사") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' 공동주택용
    If CheckSubstring(strPurpose, "공동") Then
        ss_water = Round(SS(svGONGDONG, 1) + npopulation * SS(svGONGDONG, 2), 2)
        Exit Function
    End If
        
    ' 민방위용
    If CheckSubstring(strPurpose, "민방") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' 학교용
    If CheckSubstring(strPurpose, "학교") Then
        ss_water = Round(SS(svSCHOOL, 1) + npopulation * SS(svSCHOOL, 2), 2)
        Exit Function
    End If
    
   ss_water = 900
End Function




Function aa_water(qhp As Integer, strPurpose As String, Optional ByVal nhead As Integer = 30) As Double
    'nhead - 축산업의 두수 ....

    ' 전작용
    If CheckSubstring(strPurpose, "전") Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    ' 답작용
    If CheckSubstring(strPurpose, "답") Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    
    ' 원예용
    If CheckSubstring(strPurpose, "원") Then
        aa_water = Round(AA(avWONYE, 1) + qhp * AA(avWONYE, 2), 2)
        Exit Function
    End If
    
    ' 농생활겸용
    If CheckSubstring(strPurpose, "농") Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    ' 양계장용
    If CheckSubstring(strPurpose, "양") Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    '축산용
    If CheckSubstring(strPurpose, "축") Then
        aa_water = Round(AA(avCOW, 1) + nhead * AA(avCOW, 2), 2)
        Exit Function
    End If
    
    ' 기타
    If CheckSubstring(strPurpose, "기타") Then
        aa_water = Round(AA(avDAPJAK, 1) + nhead * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
   aa_water = 900
End Function










