VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShtSelect 
   Caption         =   "시트를 선택하세요."
   ClientHeight    =   3180
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   2205
   OleObjectBlob   =   "frmShtSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShtSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSubmit_Click()

Dim WB As Workbook
Dim WS As Worksheet: Dim NewWS As Worksheet
Dim rng As Range
Dim i As Long: i = 0: Dim j As Long: j = 2
Dim endCol As Long: Dim endRow As Long
Dim strWS As String

Set WB = ThisWorkbook

'// 리스트박스 선택여부 확인
If isListBoxSelected(Me.lstSheet) = False Then MsgBox "시트를 선택하세요": Exit Sub


'// 시트병합 시트 존재여부를 확인합니다.
For Each WS In WB.Worksheets
    If WS.Name = "시트병합" Then MsgBox "시트병합 시트가 존재합니다.": Exit Sub
Next

Set NewWS = WB.Worksheets.Add(before:=WB.Worksheets(1))
NewWS.Name = "시트병합"
With NewWS
    .Cells(1, 1) = "매장"
    .Cells(1, 2) = "날짜"
    .Cells(1, 3) = "이름"
    .Cells(1, 4) = "출근시간"
    .Cells(1, 5) = "퇴근시간"
End With

'// 리스트박스에서 선택된 시트명 받아오기
For i = 0 To Me.lstSheet.ListCount - 1
    If Me.lstSheet.Selected(i) = True Then
        strWS = Me.lstSheet.List(i)
        Set WS = WB.Worksheets(strWS)
        With WS
            endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            endRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
            rng.Copy NewWS.Cells(j, 1)
            j = j + rng.Rows.Count
        End With
    End If
Next

'// 시트병합 완료 안내
MsgBox "시트병합이 완료되었습니다."
Unload Me

End Sub

Private Sub UserForm_Initialize()

Dim WB As Workbook
Dim WS As Worksheet

Set WB = ThisWorkbook

For Each WS In WB.Worksheets
    Me.lstSheet.AddItem WS.Name
Next

End Sub

Function isListBoxSelected(ListBox As MSForms.ListBox) As Boolean

Dim i As Long

For i = 0 To ListBox.ListCount - 1
    If ListBox.Selected(i) Then isListBoxSelected = True: Exit Function
Next

isListBoxSelected = False

End Function
