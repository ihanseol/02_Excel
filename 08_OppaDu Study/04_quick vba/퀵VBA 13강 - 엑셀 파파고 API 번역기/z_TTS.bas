Attribute VB_Name = "z_TTS"
'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ VocSpeak 함수
'▶ 단어를 특정 언어로 읽어주는 텍스트 음성 변환 함수입니다.
'▶ 인수 설명
'_____________Voca                  : 음성 변환 할 텍스트입니다.
'_____________Language          : 텍스트를 읽을 대상 언어입니다. 사용중인 PC에 해당 언어팩이 설치되어 있어야 합니다.
'###############################################################
 
Public Voc As Object
 
Sub VocSpeak(Optional Voca As Variant, Optional Language As String, Optional blnStop As Boolean = False)
 
' 변수 생성
If IsMissing(Voca) Then Voca = ""
Set Voc = CreateObject("SAPI.SpVoice")

' 가용한 음성 변환 목록을 하나씩 돌아가며 사용 언어와 일치하는 항목이 있는지 확인
For i = 0 To Voc.GetVoices.Count - 1
    Set Voc.Voice = Voc.GetVoices.Item(i)
    If InStr(1, Voc.Voice.GetDescription, Language) Then GoTo Speak
Next
 
'일치하는 항목이 없을 경우 안내메시지 띄우고 함수 종료
MsgBox "음성 변환 할 언어가 PC에 설치되어 있지 않습니다.", vbInformation, "오빠두엑셀 - 오류안내"
 
Exit Sub
 

' 가용한 음성 변환 항목이 있을 시 음성 변환 후 명령문을 종료합니다.
Speak:
Voc.Speak Voca, 1
If blnStop = True Then Voc.Speak "", 0

End Sub

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ GetVocList 함수
'▶ 사용 가능한 음성 변환 목록을 배열로 반홯납니다.
'▶ 인수 설명
'_____________ShowMsgBox : True일 경우 사용가능한 음성 변환 목록을 메시지박스로 출력합니다.
'###############################################################
Function GetVocList(Optional ShowMsg As Boolean = False) As Variant
 
' 변수 생성
Dim Voc As Object
Dim vaReturn As Variant
Dim v As Variant: Dim s As String
Set Voc = CreateObject("SAPI.SpVoice")
 
' 사용 가능한 음성 변환 목록 개수 넓이로 배열 생성
ReDim vaReturn(0 To Voc.GetVoices.Count - 1)
 
' 음성 이름을 배열에 추가
For i = 0 To Voc.GetVoices.Count - 1
    Set Voc.Voice = Voc.GetVoices.Item(i)
    vaReturn(i) = Voc.Voice.GetDescription
Next
 
' 메세지 출력여부 True일 경우 메세지 출력
If ShowMsg = True Then: For Each v In vaReturn: s = s & v & vbNewLine: Next: MsgBox s
 
' 결과값 반환
GetVocList = vaReturn
 
End Function
