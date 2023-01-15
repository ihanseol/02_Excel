Attribute VB_Name = "KakaoTalk_SMS"
Option Explicit

#If VBA7 Then
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-findwindowa
Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
                                                        (ByVal lpClassName As String, _
                                                        ByVal lpWindowName As String) As Long
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-findwindowexa
Private Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" _
                                                            (ByVal hwndParent As Long, _
                                                            ByVal hwndChildAfter As Long, _
                                                            ByVal lpszClass As String, _
                                                            ByVal lpszWindow As String) As Long
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessagea
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal wMsg As Long, _
                                                            ByVal wParam As Long, _
                                                            ByRef lParam As Any) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-postmessagea
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal wMsg As Long, _
                                                            ByVal wParam As Long, _
                                                            ByRef lParam As Any) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-keybd_event
Private Declare PtrSafe Sub keybd_event Lib "user32.dll" _
                                                            (ByVal bVk As Byte, _
                                                            ByVal bScan As Byte, _
                                                            ByVal dwFlags As Long, _
                                                            ByVal dwExtraInfo As Long)
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getkeystate
Private Declare PtrSafe Function GetKeyState Lib "user32" ( _
                                                            ByVal nVirtKey As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindow
Public Declare PtrSafe Function GetWindow _
                                                        Lib "user32" _
                                                        (ByVal hwnd As Long, _
                                                        ByVal wCmd As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getclassname
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal lpClassName As String, _
                                                            ByVal nMaxCount As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowtexta
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal lpString As String, _
                                                            ByVal cch As Long) As Long
#Else
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-findwindowa
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
                                                        (ByVal lpClassName As String, _
                                                        ByVal lpWindowName As String) As Long
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-findwindowexa
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" _
                                                            (ByVal hwndParent As Long, _
                                                            ByVal hwndChildAfter As Long, _
                                                            ByVal lpszClass As String, _
                                                            ByVal lpszWindow As String) As Long
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessagea
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal wMsg As Long, _
                                                            ByVal wParam As Long, _
                                                            ByRef lParam As Any) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-postmessagea
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal wMsg As Long, _
                                                            ByVal wParam As Long, _
                                                            ByRef lParam As Any) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-keybd_event
Private Declare Sub keybd_event Lib "user32.dll" _
                                                            (ByVal bVk As Byte, _
                                                            ByVal bScan As Byte, _
                                                            ByVal dwFlags As Long, _
                                                            ByVal dwExtraInfo As Long)
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getkeystate
Private Declare Function GetKeyState Lib "user32" ( _
                                                            ByVal nVirtKey As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindow
Public Declare Function GetWindow _
                                                        Lib "user32" _
                                                        (ByVal hwnd As Long, _
                                                        ByVal wCmd As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getclassname
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal lpClassName As String, _
                                                            ByVal nMaxCount As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowtexta
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal lpString As String, _
                                                            ByVal cch As Long) As Long
#End If

'######################################
' 공통 상수 선언
'######################################
' 가상키 ASCII 코드
Public Const WM_SETTEXT = &HC
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11
Public Const VK_ESC = &H1B
Public Const KEYEVENTF_KEYUP As Long = 2
Public Const VK_A = &H41
Public Const VK_V = &H56
Public Const VK_C = &H43
Public Const VK_DOWN = &H28
Public Const VK_UP = &H26

Sub Test()

'▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶
SendKakao "받을사람(친구이름)", "보낼메시지"

End Sub

'######################################
' 친구에게 카카오톡 메세지를 전송합니다.
'######################################
Function SendKakao(Target As String, Message As String, Optional ChatRoomSearch As Boolean = False, Optional iDelay As Long = 1) As Boolean

' 변수 선언
Dim hwnd_RichEdit As Long       ' 채팅입력창 hwnd

' 친구 채팅입력창 hwnd 찾기
hwnd_RichEdit = FindRecepientHwnd(Target, ChatRoomSearch, iDelay)
'▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶

'보낼 대상의 채팅창 없을 경우 안내메세지 출력 후 종료
If hwnd_RichEdit = 0 Then MsgBox "보낼 대상의 카카오톡 대화창을 찾을 수 없습니다.": Exit Function

' 메세지 보내기
Send_TextMsg Message, hwnd_RichEdit

End Function

'######################################
' 카카오톡 메세지를 전송합니다.
'######################################
Sub Send_TextMsg(Message As String, hwnd_RichEdit As Long)

' 대화상대 채팅 입력창 hWnd 에 메세지 입력
Call SendMessage(hwnd_RichEdit, WM_SETTEXT, 0, ByVal Message)
'▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶
' 사용자 Ctrl 키 입력여부 확인
If IsCtrlKeyDown = False Then
    ' Ctrl 키 미입력 시, 메세지 전송
    Call PostMessage(hwnd_RichEdit, WM_KEYDOWN, VK_RETURN, 0)
Else
    ' Ctrl 키 입력중일 경우, 강제로 Ctrl 키 올림 -> 메세지 전송 -> Ctrl 키 재입력
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    Call PostMessage(hwnd_RichEdit, WM_KEYDOWN, VK_RETURN, 0)
    keybd_event VK_CONTROL, 0, 0, 0
End If
    
End Sub

'######################################
' 보낼 대상의 hWnd 값을 찾습니다.
' 보낼 대상의 채팅창이 없을 경우 False를 반환합니다.
'######################################
Function FindRecepientHwnd(Target As String, ChatRoomSearch As Boolean, iDelay As Long) As Long

' 변수 선언
Dim dStart As Double
Dim hwnd_KakaoTalk As Long   ' 친구 채팅창 hwnd
Dim hwnd_RichEdit As Long       ' 채팅입력창 hwnd

' 보낼대상의 카톡창 실행
ActiveChat Target, iDelay, ChatRoomSearch
'▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶

' 보낼대상의 카톡창 Hwnd 찾기
hwnd_KakaoTalk = FindWindow(vbNullString, Target)
'▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶

dStart = Now
While hwnd_KakaoTalk = 0
    hwnd_KakaoTalk = FindWindow(vbNullString, Target)
    ' 1초 지날때까지 창을 못찾으면 함수 결과로 False 반환 후 종료
    If DateDiff("s", dStart, Now) > 1 Then FindRecepientHwnd = 0: Exit Function
Wend

' 친구의 채팅입력창 hWnd 찾기
hwnd_RichEdit = FindWindowEx(hwnd_KakaoTalk, 0, "RichEdit50W", vbNullString)  '카카오톡 버전에 따른 RichEdit ClassName 차이
If hwnd_RichEdit = 0 Then hwnd_RichEdit = FindWindowEx(hwnd_KakaoTalk, 0, "RichEdit20W", vbNullString)

FindRecepientHwnd = hwnd_RichEdit

End Function

'############################
' 친구 채팅창이 열려있지 않을 경우 채팅창을 검색 후 활성화합니다.
'############################
Function ActiveChat(Target As String, iDelay As Long, ChatRoomSearch As Boolean)

' 변수 설정
Dim hWndMain As Long: Dim hWndChild1 As Long: Dim hWndChild2 As Long: Dim hWndEdit As Long
Dim i As Long

' 대화상대 창 이미 열려있을 시 명령문 종료
If FindWindow(vbNullString, Target) > 0 Then Exit Function
'▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶

' 카카오톡 메인 hWnd 검색 [▶▶▶▶▶F5키]
hWndMain = FindHwndEVA
' 카카오톡 메인 창 없으면 카톡 실행 안됨 -> 함수 False 반환 후 명령문 종료
If hWndMain = 0 Then ActiveChat = False: Exit Function

' 메인 [▶▶▶▶▶F5키]
hWndChild1 = FindWindowEx(hWndMain, 0, "EVA_ChildWindow", vbNullString)
'▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶
' 연락처 목록 검색 -> EVA_Window 첫번째 항목
hWndChild2 = FindWindowEx(hWndChild1, 0, "EVA_Window", vbNullString)
' 채팅창 목록 검색 -> EVA_Window 두번째 항목
If ChatRoomSearch = True Then hWndChild2 = FindWindowEx(hWndChild1, hWndChild2, "EVA_Window", vbNullString)
' EVA_Window의 대화상대 검색창
hWndEdit = FindWindowEx(hWndChild2, 0, "Edit", vbNullString)
  
' 검색창에 대화상대 복사/붙여넣기
Call SendMessage(hWndEdit, WM_SETTEXT, 0, ByVal Target): MyDelay iDelay
' 채팅창 목록 검색일 경우 윗쪽방향키 눌러서 첫번째 항목 활성화
If ChatRoomSearch = True Then Call PostMessage(hWndEdit, WM_KEYDOWN, VK_UP, 0): MyDelay iDelay
' 엔터키로 채팅창 열기
Call PostMessage(hWndEdit, WM_KEYDOWN, VK_RETURN, 0): MyDelay iDelay

End Function

'############################
' 카카오톡 실행여부 확인 후, 실행중일 시 메인창의 hWnd 를 반환합니다.
'############################
Private Function FindHwndEVA() As Long

' 변수 설정
Dim hwnd As Long: Dim lngT As Long: Dim strT As String

' 현재 Desktop에서 실행중인 첫번째 프로그램 hWnd
hwnd = FindWindowEx(0, 0, vbNullString, vbNullString)

' 모든 hWnd 를 돌아가며 검색
While hwnd <> 0
    ' hWnd를 돌아가며 ClassName을 확인
    strT = String(100, Chr(0))
    lngT = GetClassName(hwnd, strT, 100)
    ' hWnd 의 ClassName 에 "EVA_Window_DblClk"이 포함될 경우
    If InStr(1, Left(strT, lngT), "EVA_Window_Dblclk") > 0 Then
        ' 해당 hWnd의 채팅창이름을 받아옵니다.
        strT = String(100, Chr(0))
        lngT = GetWindowText(hwnd, strT, 100)
        ' 채팅창 이름이 "카카오톡" 또는 "KakaoTalk"(영문OS 사용시)일 경우, hWnd 을 함수 결과로 반환 후 종료
        If InStr(1, Left(strT, lngT), "카카오톡") > 0 Or InStr(1, Left(strT, lngT), "KakaoTalk") > 0 Then FindHwndEVA = hwnd: Exit Function
    End If
hwnd = FindWindowEx(0, hwnd, vbNullString, vbNullString)
Wend

End Function

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ MyDelay 함수
'▶ 임시 지연 코드
'▶ ______________iDelay    : 지연 속도 조절 (클수록 많은 지연, 1 ≒ 4ms)
'###############################################################
Sub MyDelay(Optional iDelay As Long = 1)

Dim i As Long
For i = 1 To iDelay * 10000000:        i = i + 1:        Next

End Sub

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ IsCtrlKeyDown 함수
'▶ 키보드 Ctrl 키 누름여부를 확인합니다.
'▶ 인수 설명
'_____________LeftRightKey : 왼쪽/오른쪽 Ctrl 키 누름여부를 정합니다.
' 1 : 왼쪽 Ctrl키 입력시 TRUE
' 2 : 오른쪽 Ctrl키 입력시 TRUE
' 3 : 양쪽 Ctrl키 동시 입력시 TRUE
' 0 : 둘 중 하나라도 입력시 TRUE
'###############################################################

Private Function IsCtrlKeyDown(Optional LeftRightKey As Long = 0) As Boolean

Const VK_LCTRL = &HA2: Const VK_RCTRL = &HA3: Const KEY_MASK As Integer = &HFF80

Dim Result As Long

Select Case LeftRightKey
    '왼쪽 CTRL 키 입력여부 확인
    Case 1:        Result = GetKeyState(VK_LCTRL) And KEY_MASK
    '오른쪽 CTRL 키 입력여부 확인
    Case 2:        Result = GetKeyState(VK_RCTRL) And KEY_MASK
    '양쪽 CTRL 키 동시 입력여부 확인
    Case 3:        Result = GetKeyState(VK_LCTRL) And GetKeyState(VK_RCTRL) And KEY_MASK
    'CTRL 키 둘중 하나의 입력여부 확인
    Case Else:    Result = GetKeyState(vbKeyControl) And KEY_MASK
End Select

IsCtrlKeyDown = CBool(Result)

End Function


