Attribute VB_Name = "mod_MakeFieldList"
Option Explicit

' 이곳에다가 기본적인 설정값을 세팅해준다.
' 파일이름과, 조사일같은것들을 ...

Const EXPORT_DATE As String = "2023-12-20"
Const EXPORT_ADDR_HEADER As String = "서울특별시 "
Const EXPORT_FILE_NAME As String = "d:\05_Send\iyong_template.xlsx"
        
' 1인 1일당 급수량, 엑셀파일을 보고 검사
' 서울특별시 강북구
Const ONEMAN_WATER_SUPPLY As Double = 265.16
        
Public Enum ALLOW_TYPE_VALUE
     at_HEOGA = 0
     at_SINGO = 1
End Enum


Sub delay(ti As Integer)
    Application.Wait Now + TimeSerial(0, 0, ti)
End Sub


Sub MakeFieldList()
    Call make_datamid
    Call Delete_Outside_Boundary
    Call ExportData
End Sub


Sub ExportData()
    ' data_mid 에서, 중간과정으로 만들어진 데이타를 불러와서, 파이썬처리용 엑셀쉬트를 만든다.
    Call Make_DataOut
    Call ExportCurrentWorksheet("data_out")
End Sub

Sub ExportCurrentWorksheet(sh As String)
    Dim filePath As String
    
    If Not ActivateSheet(sh) Then
        Debug.Print "ActivateSheet Error, maybe sheet does not exist ...."
        Exit Sub
    End If
        
    'filePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx")
    ' filePath = "d:\05_Send\aaa.xlsx"
    
    filePath = EXPORT_FILE_NAME
    
    If VarType(filePath) = vbString Then
    
        If Dir(filePath) <> "" Then
            ' Delete the file
            Kill filePath
    
'            If MsgBox("The file " & filePath & " already exists. Do you want to overwrite it?", _
'                      vbQuestion + vbYesNo, "Confirm Overwrite") = vbNo Then
'                Exit Sub
'            End If
        End If
    
    
        If Sheets(sh).Visible = False Then
            Sheets(sh).Visible = True
        End If
        
        Sheets(sh).Activate
        ActiveSheet.Copy
        ActiveWorkbook.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook, ConflictResolution:=xlLocalSessionChanges
        ActiveWorkbook.Close savechanges:=False
    End If
End Sub


Private Sub DeleteFile(filePath As String)
    ' Check if the file exists before attempting to delete
    If Dir(filePath) <> "" Then
        ' Delete the file
        Kill filePath
        ' MsgBox "File deleted successfully.", vbInformation
    Else
        ' MsgBox "File not found.", vbExclamation
    End If
End Sub



Function ActivateSheet(sh As String) As Boolean
    On Error GoTo ErrorHandler
    Sheets(sh).Activate
    ActivateSheet = True
    Exit Function
    
ErrorHandler:
'    MsgBox "An error occurred while trying to activate the sheet." & vbNewLine & _
'           "Please check that the sheet name is correct and try again.", _
'           vbExclamation, "Error"

    ActivateSheet = False
End Function

Sub Make_DataOut()
    Dim str_, address, id, purpose As String
    Dim allowType, i, lastRow  As Integer
    Dim simdo, diameter, hp, capacity, tochool, Q As Double
    Dim setting As String
    
    Dim ag_start, ag_end, ag_year  As String
    Dim sayong_gagu, sayong_ingu, sayong_ilin_geupsoo As Double
    Dim usage_day, usage_month, usage_year As Double
    
    str_ = ChrW(&H2714)
    
    
    If Not Sheets("data_mid").Visible Then
        Sheets("data_mid").Visible = True
    End If
    
    Sheets("data_mid").Activate
    
    Call initialize
    lastRow = getlastrow()
    
    For i = 2 To lastRow
    
        Call GetDataFromSheet(i, id, address, allowType, simdo, diameter, hp, capacity, tochool, purpose, Q)
        
        If allowType = at_HEOGA Then
            setting = setting & "b,"
            ' 허가시설
        Else
            setting = setting & "c,"
            ' 신고시설
        End If
        
'       충적관정인지, 암반관정인지를 검사해서 추가해줌 ...
'       If (diameter >= 150) And (hp >= 1#) Then
'            setting = setting & "aq,"
'       Else
'            setting = setting & "ap,"
'       End If

        setting = setting & IIf(diameter >= 150 And hp >= 1#, "aq,", "ap,")

       
        Select Case LCase(Left(id, 1))
            Case "s"
                setting = setting & "f,"
                setting = setting & SS_StringCheck(purpose)
                setting = setting & SS_PublicCheck(purpose)
            
            Case "a"
                setting = setting & "u,"
                setting = setting & AA_StringCheck(purpose)
                
                If allowType = at_HEOGA Then
                    setting = setting & "ab,"
                Else
                    setting = setting & AA_PublicCheck(purpose)
                End If
                                            
            Case "i"
                setting = setting & "o,"
                setting = setting & II_StringCheck(purpose)
                setting = setting & II_PublicCheck(purpose)
                
                
        End Select
        
        ' 음용수 인가 , 먹을수있는 물인가 ?
        If IsDrinking(purpose) Then
            setting = setting & "ah,"
        Else
            setting = setting & "ai,"
        End If
        
        
        
        ' ad = 연중사용
        Select Case LCase(Left(id, 1))
            Case "s"
                setting = setting & "ad,"
                ag_start = "1"
                ag_end = "12"
                ag_year = "365"
            
            Case "a"
                '농업용 : 3 ~ 11월까지
                ag_start = "3"
                ag_end = "11"
                ag_year = "274"
            
            
            Case "i"
                ' 공업용 - 연중사용
                setting = setting & "ad,"
                ag_start = "1"
                ag_end = "12"
                ag_year = "365"
                
        End Select
        
        
        '음용수, 사용가구, 사용인구, 일인급수량이 결정됨
        If IsDrinking(purpose) Then
                 ' 용도가, 가정용일때 ...
                 If CheckSubstring(purpose, "가정") Then
                        sayong_gagu = 1
                        sayong_ingu = SS_CITY
                        sayong_ilin_geupsoo = Q / SS_CITY
                 End If
                
                 ' https://kosis.kr/statHtml/statHtml.do?orgId=110&tblId=DT_11001N_2013_A055
                 ' 용도가 간이상수도 일때 ...
                 If CheckSubstring(purpose, "간이") Then
                        sayong_gagu = 30
                        sayong_ingu = 90
                        sayong_ilin_geupsoo = ONEMAN_WATER_SUPPLY
                End If
        Else
            sayong_gagu = 0
            sayong_ingu = 0
            sayong_ilin_geupsoo = 0
        End If
                
         
        ' 일사용량 계산
        usage_day = Q
        usage_month = Q * 29
        
        If LCase(Left(id, 1)) = "s" Then
            usage_year = usage_month * 12
        Else
            usage_year = usage_month * 8
        End If
        
        
        '허가공 -  av,aw,ay,az,ba,
        
        ' 관정현황 체크
        Select Case LCase(Left(id, 1))
            Case "s"
                If allowType = at_SINGO Then ' 신고시설이면
                    If CheckSubstring(purpose, "일반") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "간이") Then setting = setting & "av,aw,ax,ay,az,ba,"
                    If CheckSubstring(purpose, "공동") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "민방") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "학교") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "청소") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "공사") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "겸용") Then setting = setting & "av,aw,ay,"
                Else ' 허가시설이면
                    setting = setting & "av,aw,ax,ay,az,ba"
                End If
            
            Case "a"
                If allowType = at_SINGO Then ' 신고시설이면
                    If CheckSubstring(purpose, "전작") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "답작") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "원예") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "겸용") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "양어장") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "축산") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "기타") Then setting = setting & "aw,ay,"
                Else ' 허가시설이면
                    setting = setting & "av,aw,ax,ay,az,ba"
                End If
            
            
            Case "i"
                ' 공업용 - 연중사용
                setting = setting & "ad,"
                If allowType = at_SINGO Then
                    ' 신고시설이면
                    setting = setting & "aw,ay,"
                    
                Else
                    ' 허가시설이면
                    setting = setting & "av,aw,ax,ay,az,ba"
                End If
                
        End Select
        
        
        
        
        Debug.Print "**********************************"
        Debug.Print setting
        
        Call PutDataSheetOut(i, setting, address, simdo, diameter, hp, capacity, tochool, Q, ag_start, ag_end, ag_year, _
                             sayong_gagu, sayong_ingu, sayong_ilin_geupsoo, usage_day, usage_month, usage_year)
        
        
        setting = ""
    
    Next i

' =INDEX(itable[value], MATCH("d1", itable[key], 0))

End Sub

Sub PutDataSheetOut(ii As Variant, setting As Variant, address As Variant, simdo As Variant, diameter As Variant, hp As Variant, _
                    capacity As Variant, tochool As Variant, Q As Variant, _
                    ag_start As Variant, ag_end As Variant, ag_year As Variant, _
                    sayong_gagu As Variant, sayong_ingu As Variant, sayong_ilin_geupsoo As Variant, _
                    usage_day As Variant, usage_month As Variant, usage_year As Variant)

    Dim out() As String
    Dim i As Integer
    Dim index, str_, setting_1 As String
    
    Sheets("data_out").Activate
    
    With Range("A" & CStr(ii) & ":BB" & CStr(ii))
        .Value = " "
    End With

    str_ = ChrW(&H2714)
    
    
    setting_1 = DeepCopyString(CStr(setting))
    
    out = FilterString(setting_1)
    
    For i = LBound(out) To UBound(out)
        index = out(i)
        Sheets("data_out").Cells(ii, index).Value = str_
    Next i
    
    '  myString = Format(myDate, "yyyy-mm-dd")
    Sheets("data_out").Cells(ii, "a").Value = " " & Format(EXPORT_DATE, "yyyy-mm-dd") & "."
    Sheets("data_out").Cells(ii, "e").Value = address
    Sheets("data_out").Cells(ii, "ar").Value = simdo
    Sheets("data_out").Cells(ii, "as").Value = diameter
    Sheets("data_out").Cells(ii, "at").Value = hp
    Sheets("data_out").Cells(ii, "au").Value = capacity
    Sheets("data_out").Cells(ii, "av").Value = tochool
    
    Sheets("data_out").Cells(ii, "ae").Value = ag_start
    Sheets("data_out").Cells(ii, "af").Value = ag_end
    Sheets("data_out").Cells(ii, "ag").Value = ag_year
    
    ' 음용수 일때만, 사용가구, 사용인구, 1인급수 세팅
    If Sheets("data_out").Cells(ii, "ah").Value = ChrW(&H2714) Then
        Sheets("data_out").Cells(ii, "aj").Value = CStr(Format(sayong_gagu, "0.00"))
        Sheets("data_out").Cells(ii, "ak").Value = CStr(Format(sayong_ingu, "0.00"))
        Sheets("data_out").Cells(ii, "al").Value = CStr(Format(sayong_ilin_geupsoo, "0.00"))
    End If
    
    Sheets("data_out").Cells(ii, "am").Value = CStr(Format(usage_day, "0.00"))
    Sheets("data_out").Cells(ii, "an").Value = CStr(Format(usage_month, "#,##0"))
    Sheets("data_out").Cells(ii, "ao").Value = CStr(Format(usage_year, "#,##0"))
    

End Sub
                             
                          
' GetDataFromSheet(i, id, address, allowType, simdo, diameter, hp, capacity, tochool, purpose, Q)
Sub GetDataFromSheet(i As Variant, id As Variant, address As Variant, allowType As Variant, _
                     simdo As Variant, diameter As Variant, hp As Variant, capacity As Variant, tochool As Variant, _
                     purpose As Variant, Q As Variant)
    
    id = Sheets("data_mid").Cells(i, "a").Value
    address = Sheets("data_mid").Cells(i, "b").Value
    allowType = Sheets("data_mid").Cells(i, "c").Value
    simdo = Sheets("data_mid").Cells(i, "d").Value
    diameter = Sheets("data_mid").Cells(i, "e").Value
    hp = Sheets("data_mid").Cells(i, "f").Value
    capacity = Sheets("data_mid").Cells(i, "g").Value
    tochool = Sheets("data_mid").Cells(i, "h").Value
    purpose = Sheets("data_mid").Cells(i, "i").Value
    Q = Sheets("data_mid").Cells(i, "j").Value
    
End Sub


Function getlastrow() As Integer
    ' ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    getlastrow = ActiveSheet.Range("A3333").End(xlUp).row
End Function


' 2024-1-11 , modify last cell check
' using cell reference SUM_SS, SUM_AA, SUM_II

Sub LastRowFindAll(row_ss As Variant, row_aa As Variant, row_ii As Variant)
    
    If Range("SUM_SS").Value = 0 Then
        row_ss = 0
    Else
        Sheets("ss").Activate
        row_ss = getlastrow() - 1
    End If
           
    If Range("SUM_AA").Value = 0 Then
        row_aa = 0
    Else
        Sheets("aa").Activate
        row_aa = getlastrow() - 1
    End If
      
    
    If Range("SUM_II").Value = 0 Then
        row_ii = 0
        Exit Sub
    Else
        Sheets("ii").Activate
        row_ii = getlastrow() - 1
    End If
    
End Sub

Sub EraseSheetData()
    Worksheets("data_mid").Range("A2:J1000").Delete
    Worksheets("data_out").Range("A2:BD1000").Delete
End Sub


' allowType = 1 - 신고공
' allowType = 0 - 허가공
Public Sub make_datamid()
    Dim i, j, row_end As Integer
    Dim newAddress, id, purpose As String
    Dim allowType As Integer
    Dim well_data(1 To 5) As Double
    Dim Q As Double
    Dim boundary As String
    Dim row_ss, row_aa, row_ii As Integer
    
    Call LastRowFindAll(row_ss, row_aa, row_ii)
    Call EraseSheetData
    
    Sheets("ss").Activate
    ' Debug.Print row_end
    For i = 1 To row_ss
        id = Cells(i + 1, "a").Value
        ' 주소헤더, 지역에 따라 값을 다시 설정해주어야 한다.
        newAddress = EXPORT_ADDR_HEADER & Cells(i + 1, "c") & " " & Cells(i + 1, "d") & " " & Cells(i + 1, "e") & " , " & id
        
        If Cells(i + 1, "b").Value = "신고공" Then
            allowType = 1
        Else
            allowType = 0
        End If
        
        ' Debug.Print allowType, newAddress
        
        For j = 1 To 5
            well_data(j) = Cells(i + 1, Chr(Asc("f") + j - 1)).Value
        Next j
        
        purpose = Cells(i + 1, "k").Value
        Q = Cells(i + 1, "l").Value
        boundary = Cells(i + 1, "s").Value
        
        Call putdata(i, id, newAddress, allowType, well_data, purpose, Q, boundary)
    Next i
    
    
    Sheets("aa").Activate
    ' Debug.Print row_end
    For i = 1 To row_aa
    
        id = Cells(i + 1, "a").Value
        newAddress = EXPORT_ADDR_HEADER & Cells(i + 1, "c") & " " & Cells(i + 1, "d") & " " & Cells(i + 1, "e") & " , " & id
        
        If Cells(i + 1, "b").Value = "신고공" Then
            allowType = 1
        Else
            allowType = 0
        End If
        
        ' Debug.Print allowType, newAddress
        
        For j = 1 To 5
            well_data(j) = Cells(i + 1, Chr(Asc("f") + j - 1)).Value
        Next j
        
        purpose = Cells(i + 1, "k").Value
        Q = Cells(i + 1, "l").Value
        boundary = Cells(i + 1, "s").Value
        
        Call putdata(i + row_ss, id, newAddress, allowType, well_data, purpose, Q, boundary)
    Next i
    
    Sheets("ii").Activate
    ' Debug.Print row_end
    
    For i = 1 To row_ii
    
        id = Cells(i + 1, "a").Value
        newAddress = EXPORT_ADDR_HEADER & Cells(i + 1, "c") & " " & Cells(i + 1, "d") & " " & Cells(i + 1, "e") & " , " & id
        
        If Cells(i + 1, "b").Value = "신고공" Then
            allowType = 1
        Else
            allowType = 0
        End If
        
        ' Debug.Print allowType, newAddress
        
        For j = 1 To 5
            well_data(j) = Cells(i + 1, Chr(Asc("f") + j - 1)).Value
        Next j
        
        purpose = Cells(i + 1, "k").Value
        Q = Cells(i + 1, "l").Value
        boundary = Cells(i + 1, "s").Value
        
        Call putdata(i + row_ss + row_aa, id, newAddress, allowType, well_data, purpose, Q, boundary)
    Next i
    
End Sub

' 2024-1-11
' delete outside boundary

Private Sub Delete_Outside_Boundary()

    Dim row_ss, row_aa, row_ii As Integer
    Dim i, j As Integer
        
    j = 2
    Sheets("data_mid").Activate
    
    For i = 1 To getlastrow()
        
        If Cells(j, "K").Value = "O" Then
            j = j + 1
        Else
            Rows(j & ":" & j).Select
            Selection.Delete Shift:=xlUp
        End If
    
    Next i

End Sub

Sub putdata(i As Variant, id As Variant, newAddress As Variant, allowType As Variant, well_data As Variant, _
            purpose As Variant, Q As Variant, boundary As Variant)
        
    Sheets("data_mid").Cells(i + 1, "a").Value = id
    Sheets("data_mid").Cells(i + 1, "b").Value = newAddress
    Sheets("data_mid").Cells(i + 1, "c").Value = allowType
    Sheets("data_mid").Cells(i + 1, "d").Value = well_data(1)
    Sheets("data_mid").Cells(i + 1, "e").Value = well_data(2)
    Sheets("data_mid").Cells(i + 1, "f").Value = well_data(3)
    Sheets("data_mid").Cells(i + 1, "g").Value = well_data(4)
    Sheets("data_mid").Cells(i + 1, "h").Value = well_data(5)
    Sheets("data_mid").Cells(i + 1, "i").Value = purpose
    Sheets("data_mid").Cells(i + 1, "j").Value = Q
    Sheets("data_mid").Cells(i + 1, "k").Value = boundary
    
End Sub













