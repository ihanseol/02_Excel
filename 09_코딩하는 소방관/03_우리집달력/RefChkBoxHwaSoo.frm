VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RefChkBoxHwaSoo 
   Caption         =   "Check Anything"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "RefChkBoxHwaSoo.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "RefChkBoxHwaSoo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

    Dim DBSheet As Worksheet
    Dim w() As String
    Dim Singlew As Variant
    Dim i As Integer
    Dim j As Integer
    Dim r As Integer
    Dim a As Integer
    Dim InputYear As Integer
    Dim v() As String
    Dim CalendarData() As Variant
    Dim dataMonth As Integer
    Dim SingleRange As Range
    Dim dDate As Date
    Dim eMon, eDay As Integer '음력 달,월 저장
    Dim strDate As String
    
    '달력 초기화
    Call ClearCalendar

    Set DBSheet = Sheets("DB")
    InputYear = year(Date)

    
    '체크박스 인덱스를 배열에 삽입
    For i = 0 To Me.Controls.Count - 1
        If Me.Controls(i) Then
            ReDim Preserve w(j)
            w(j) = i
            j = j + 1
        End If
    Next
    
    '체크박스 선택 안 하면 매크로 종료
    If j <= 1 Then GoTo j

    '체크한 데이터 중 날짜의 행을 확인하여 j에 대입
    j = 0

    For Each Singlew In w
        If DBSheet.Cells(1, CInt(Singlew)).Value = "일자" Then
            j = Singlew
        End If
    Next
    
    '날짜열을 선택하지 않으면 매크로 종료
    If j = 0 Then
        MsgBox "날짜가 있는 열을 선택하세요", vbCritical
        GoTo j
    End If
    
    '선택한 열을 순환하며 달력에 출력할 데이터를 배열에 삽입
    r = DBSheet.Range("A1").CurrentRegion.Rows.Count
    ReDim v(UBound(w) - 1)
    ReDim CalendarData(r - 2, 1)
    For i = 2 To r
        For Each Singlew In w
            If Singlew <> j Then
                v(a) = DBSheet.Cells(i, CInt(Singlew)).Value
                a = a + 1
            End If
        Next
        a = 0
        
        '출력할 데이터와 날짜를 분리하여 배열에 삽입
        CalendarData(i - 2, 0) = Join(v, ", ")
        
        If DBSheet.Cells(i, "D").Value = "음력" Then
            dDate = dependOnMonthType(DBSheet.Cells(i, CInt(j)).Value)
            eMon = Month(dDate)
            eDay = Day(dDate)
            
            InputYear = IIf(eMon = 12, year(Date) - 1, year(Date))
            strDate = Lun2Sol(InputYear, eMon, eDay, isYoonDal2(dDate))
            
            CalendarData(i - 2, 1) = getYMD(strDate)
        Else
            dDate = dependOnMonthType(DBSheet.Cells(i, CInt(j)).Value)
            CalendarData(i - 2, 1) = dDate
        End If
    Next
    
    '데이터 배열을 순환하여 달력에 데이터 출력
    For i = 0 To r - 2
        dDate = dependOnMonthType(CalendarData(i, 1))
        dataMonth = Month(dDate) + 2
        
        'dataMonth = Month(CalendarData(i, 1)) + 2
        
        For Each SingleRange In Sheets(dataMonth).UsedRange
            'If SingleRange = CalendarData(i, 1) Then
             If SingleRange = dDate Then
                With SingleRange.Offset(1)
                    If .Value <> "" Then
                        .Value = .Value & vbLf & CalendarData(i, 0)
                    Else
                        .Value = CalendarData(i, 0)
                    End If
                    .HorizontalAlignment = xlLeft
                    .Font.Color = vbBlack
                End With
            End If
        Next
        
    Next

j:
    Unload Me
    
End Sub


Function isYoonDal(d As Date) As Boolean
    Dim arr As Variant
    Dim yy, mm As Integer
    Dim thisYear As Integer
    Dim i, j As Integer
    
    thisYear = year(Date)
    yy = year(d)
    mm = Month(d)
    
    arr = [{2023, 2 ; 2025,6 ; 2028,5 ; 2031,3 ; 2033,11 ; 2036,6 ; 2039,5 ; 2042,2 ; 2044,7 ; 2047,5 ; 2050,3 ; 2052,8 ; 2055,6 ; 2058,4 ; 2061,3}]
    
    For i = 1 To UBound(arr, 1)
        If yy < arr(i, 1) Then
            isYoonDal = False
            Exit Function
        End If
    
        If yy = arr(i, 1) Then
              If mm = arr(i, 2) Then
                isYoonDal = True
                Exit Function
              Else
                isYoonDal = False
                Exit Function
              End If
        End If
    Next i
End Function


Function isYoonDal2(d As Date) As Boolean
    Dim arr As Variant
    Dim yy, mm As Integer
    Dim thisYear As Integer
    Dim i, j As Integer
    Dim dict As New Scripting.Dictionary
    
    thisYear = year(Date)
    yy = year(d)
    mm = Month(d)
    
    arr = [{ 2023, 2 ; 2025,6 ; 2028,5 ; 2031,3 ; 2033,11 ; 2036,6 ; 2039,5 ; 2042,2 ; 2044,7 ; 2047,5 ; 2050,3 ; 2052,8 ; 2055,6 ; 2058,4 ; 2061,3}]
    
    For i = 1 To UBound(arr, 1)
        dict.Add arr(i, 1), arr(i, 2)
    Next i
    
    If dict.Exists(yy) Then
        If mm = dict(yy) Then
          isYoonDal2 = True
        Else
          isYoonDal2 = False
        End If
    Else
        isYoonDal2 = False
    End If
   
End Function





Function getYMD(strDate As String) As Date

    Dim Match As Object
    Dim i As Integer
    Dim InputYear As Integer
    
    InputYear = year(Date)
    
    With CreateObject("VBscript.regexp")    '정규식 생성
    
        .Global = True  '모든 숫자 파악
        .Pattern = "\d+"    '숫자만
        
        If .test(strDate) = True Then
            Set Match = .Execute(strDate)
             getYMD = DateSerial(CInt(Match(0)), CInt(Match(1)), CInt(Match(2)))
        Else
            getYMD = Date
        End If
        
    End With


End Function


Function dependOnMonthType(ByVal d As Variant) As Date
    
    'Debug.Print TypeName(d)
    
    If TypeName(d) = "String" Then
        dependOnMonthType = makeDateFromString(d)
    Else
        dependOnMonthType = d
    End If

End Function

Function makeDateFromString(tmp As Variant) As Date
 
    Dim Match As Object
    Dim i As Integer
    Dim InputYear As Integer
    
    InputYear = year(Date)
    
    With CreateObject("VBscript.regexp")    '정규식 생성
    
        .Global = True  '모든 숫자 파악
        .Pattern = "\d+"    '숫자만
        
        If .test(tmp) = True Then
            Set Match = .Execute(tmp)
             makeDateFromString = DateSerial(InputYear, CInt(Match(0)), CInt(Match(1)))
        Else
            makeDateFromString = Date
        End If
        
    End With
    
End Function
 



Private Sub UserForm_Initialize()

    Dim i As Integer
    Dim OptionList
    Dim chkBox As MSForms.CheckBox
    Dim btn As CommandButton
    Dim DBSheet As Worksheet
    
    '달력에 출력을 원하는 제목을 체크박스로 나타내기
    Set DBSheet = Sheets("DB")
    OptionList = DBSheet.Range("A1").CurrentRegion.Rows(1)
    
    For i = 1 To UBound(OptionList, 2)
    
        Set chkBox = Me.Controls.Add("Forms.CheckBox.1", "CheckBox_" & i)
        
        With chkBox
            .Caption = OptionList(1, i)
            .Left = 5
            .Top = 5 + ((i - 1) * 20)
            
            Me.Width = .Width
            Me.Height = .Height * (i + 2)
        End With
    
    Next i
    
    '확인버튼 생성
    Set btn = Me.CommandButton1
    
    With btn
        .Caption = "확인"
        .Top = Me.Height - .Height + (0.5 * chkBox.Height) + 10
        .Left = (Me.Width * 0.5) - (.Width * 0.5)
        Me.Height = Me.Height + .Height + (0.5 * chkBox.Height) + 20
    End With
    
End Sub
