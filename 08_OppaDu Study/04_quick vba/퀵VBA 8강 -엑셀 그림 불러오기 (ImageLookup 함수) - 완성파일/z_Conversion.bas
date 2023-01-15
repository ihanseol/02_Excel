Attribute VB_Name = "z_Conversion"
Option Explicit

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ cvRng 함수
'▶ 사용자지정함수에서 범위로 인수를 받을 시 사용합니다. 만약 인수가 범위로 입력되었을 경우, 범위에 입력된 값을 반환합니다.
'▶ 인수 설명
'_____________TargetRng     : 값을 반환할 범위 또는 그외 값입니다.
'###############################################################
Function cvRng(TargetRng)

If TypeName(TargetRng) = "Range" Then
    cvRng = TargetRng.Value
Else
    cvRng = TargetRng
End If

End Function
