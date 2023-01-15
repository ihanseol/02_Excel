VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RefChkBox 
   Caption         =   "Check Anything"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "RefChkBox.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "RefChkBox"
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
    Dim v() As String
    Dim CalendarData() As Variant
    Dim dataMonth As Integer
    Dim SingleRange As Range

    '달력 초기화
    Call ClearCalendar

    Set DBSheet = Sheets("DB")
    
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
        If TypeName(DBSheet.Cells(2, CInt(Singlew)).Value) = "Date" Then
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
        CalendarData(i - 2, 1) = DBSheet.Cells(i, CInt(j)).Value
    Next
    
    '데이터 배열을 순환하여 달력에 데이터 출력
    For i = 0 To r - 2
        dataMonth = Month(CalendarData(i, 1)) + 2
        
        For Each SingleRange In Sheets(dataMonth).UsedRange
            If SingleRange = CalendarData(i, 1) Then
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
