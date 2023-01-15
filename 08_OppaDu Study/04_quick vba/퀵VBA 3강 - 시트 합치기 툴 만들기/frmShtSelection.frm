VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShtSelection 
   Caption         =   "시트를 선택하세요."
   ClientHeight    =   3216
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   2910
   OleObjectBlob   =   "frmShtSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShtSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSubmit_Click()

Dim WB As Workbook
Dim WS As Worksheet: Dim newWS As Worksheet
Dim Rng As Range
Dim i As Long: Dim j As Long: j = 2
Dim endCol As Long: Dim endRow As Long

Set WB = ThisWorkbook

'// isListBoxSelected 함수에 대한 자세한 설명은 아래 링크를 참고하세요.
'// https://oppadu.com/엑셀-사전/vba-리스트박스-값-선택여부/
If isListBoxSelected(Me.lstSheet) = False Then MsgBox "시트를 선택하세요.": Exit Sub

For Each WS In WB.Worksheets
    If WS.Name = "시트병합" Then MsgBox "시트병합 시트가 존재합니다. 시트병합 시트의 이름을 변경하시거나 삭제하신 뒤 다시 실행해주세요.": Exit Sub
Next

Set newWS = WB.Worksheets.Add(before:=WB.Worksheets(1))
newWS.Name = "시트병합"
With newWS
    .Cells(1, 1) = "매장"
    .Cells(1, 2) = "날짜"
    .Cells(1, 3) = "이름"
    .Cells(1, 4) = "출근시간"
    .Cells(1, 5) = "퇴근시간"
End With

For i = 0 To Me.lstSheet.ListCount - 1
    If Me.lstSheet.Selected(i) = True Then
        Set WS = WB.Worksheets(Me.lstSheet.List(i))
        With WS
            '// 시트의 마지막 행/열 받아오기에 대한 자세한 설명은 아래 링크를 참고하세요.
            '// https://oppadu.com/엑셀-사전/엑셀-vba-마지막-셀-찾기-마지막-행-찾기/
            endRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            Set Rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
            Rng.Copy newWS.Cells(j, 1)
            j = j + Rng.Rows.Count
        End With
    End If
Next

MsgBox "시트 병합이 완료되었습니다."
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
