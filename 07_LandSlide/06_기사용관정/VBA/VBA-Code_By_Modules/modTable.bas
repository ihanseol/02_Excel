Attribute VB_Name = "modTable"
' ***************************************************************
' modTable
'
' ***************************************************************


Option Explicit

Sub test_tableindex()
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("tableCNU")
    
    Dim values() As Variant
    values = tbl.ListColumns("nonsan").DataBodyRange.Value
    
    Dim i As Long
    For i = 1 To UBound(values, 1)
        ActiveSheet.Cells(29, Chr(Asc("A") + i)).Value = values(i, 1)
    Next i
End Sub


Sub initialize_CNU(area As String)
    Dim tbl As ListObject
    Set tbl = Sheets("ref").ListObjects("tableCNU")
    
    Dim values() As Variant
    
    If area = "" Then
        area = "default"
    End If
    
    values = tbl.ListColumns(area).DataBodyRange.Value
         
    '전라남도, 목포시, 2020 환경부 지하수업무수행지침
    SS(svGAJUNG, 1) = values(1, 1)
    SS(svGAJUNG, 2) = values(2, 1)
    SS_CITY = values(3, 1)
    
    SS(svILBAN, 1) = values(4, 1)
    SS(svILBAN, 2) = values(5, 1)
    
    SS(svSCHOOL, 1) = values(6, 1)
    SS(svSCHOOL, 2) = values(7, 1)
    
    SS(svGONGDONG, 1) = values(8, 1)
    SS(svGONGDONG, 2) = values(9, 1)
    
    SS(svMAEUL, 1) = values(10, 1)
    SS(svMAEUL, 2) = values(11, 1)
    
'----------------------------------------

    AA(avJEONJAK, 1) = values(12, 1)
    AA(avJEONJAK, 2) = values(13, 1)
    
    AA(avDAPJAK, 1) = values(14, 1)
    AA(avDAPJAK, 2) = values(15, 1)
    
    AA(avWONYE, 1) = values(16, 1)
    AA(avWONYE, 2) = values(17, 1)
    
    AA(avCOW, 1) = values(18, 1)
    AA(avCOW, 2) = values(19, 1)
    
    AA(avPIG, 1) = values(20, 1)
    AA(avPIG, 2) = values(21, 1)
    
    AA(avCHICKEN, 1) = values(22, 1)
    AA(avCHICKEN, 2) = values(23, 1)
    
End Sub


Sub initialize_JIYEOL(area As String)
    Dim tbl As ListObject
    Set tbl = Sheets("ref1").ListObjects("tableJIYEOL")
    
    Dim values() As Variant
    
    
    If (area = "") Then
        area = "default"
    End If
    
    values = tbl.ListColumns(area).DataBodyRange.Value
           
    '전라남도, 목포시, 2020 환경부 지하수업무수행지침
    SS(svGAJUNG, 1) = values(1, 1)
    SS(svGAJUNG, 2) = values(2, 1)
    SS_CITY = values(3, 1)
    
    SS(svILBAN, 1) = values(4, 1)
    SS(svILBAN, 2) = values(5, 1)
    
    SS(svSCHOOL, 1) = values(6, 1)
    SS(svSCHOOL, 2) = values(7, 1)
    
    SS(svGONGDONG, 1) = values(8, 1)
    SS(svGONGDONG, 2) = values(9, 1)
    
    SS(svMAEUL, 1) = values(10, 1)
    SS(svMAEUL, 2) = values(11, 1)
    
'----------------------------------------

    AA(avJEONJAK, 1) = values(12, 1)
    AA(avJEONJAK, 2) = values(13, 1)
    
    AA(avDAPJAK, 1) = values(14, 1)
    AA(avDAPJAK, 2) = values(15, 1)
    
    AA(avWONYE, 1) = values(16, 1)
    AA(avWONYE, 2) = values(17, 1)
    
    AA(avCOW, 1) = values(18, 1)
    AA(avCOW, 2) = values(19, 1)
    
    AA(avPIG, 1) = values(20, 1)
    AA(avPIG, 2) = values(21, 1)
    
    AA(avCHICKEN, 1) = values(22, 1)
    AA(avCHICKEN, 2) = values(23, 1)
End Sub




