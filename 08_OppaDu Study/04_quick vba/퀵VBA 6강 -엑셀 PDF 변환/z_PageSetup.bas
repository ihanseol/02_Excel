Attribute VB_Name = "z_PageSetup"
Public Enum ePrintMargin
    xlNone = 0
    xlNarrow = 1
    xlNormal = 2
    xlWide = 3
End Enum

Public Enum ePaperSize
    xlA4 = 9
    xlA3 = 8
    xlLetter = 1
    xlA5 = 11
End Enum

Function getPrintMargin(eValue As ePrintMargin) As Variant

'// 설정된 eNum 값으로 페이지 여백설정을 위한 값을 배열로 나열합니다.

Select Case eValue
    Case 0
        getPrintMargin = Array(0.05, 0.05, 0.05, 0.05, 0.1, 0.1)
    Case 1
        getPrintMargin = Array(0.25, 0.25, 0.75, 0.75, 0.3, 0.3)
    Case 2
        getPrintMargin = Array(0.7, 0.7, 0.75, 0.75, 0.3, 0.3)
    Case 3
        getPrintMargin = Array(1, 1, 1, 1, 0.5, 0.5)
End Select

End Function

Sub Page_Setup(WS As Worksheet, Optional LHead As String = "", Optional RHead As String = "&D / &T", _
                Optional LFoot As String = "본 페이지의 무단복제를 금합니다.", Optional RFoot As String = "&P / &N 페이지", _
                Optional eMargin As ePrintMargin = xlNarrow, _
                Optional HFit As Boolean = True, Optional VFit As Boolean = False, _
                Optional HCenter As Boolean = True, Optional VCenter As Boolean = False, _
                Optional eOrient As XlPageOrientation = xlPortrait, Optional eSize As ePaperSize = xlA4)

Dim pSetup As String
Dim varMargin As Variant
Dim lngOrient As Integer

'// 인쇄설정 업데이트 중단 (속도증가)
Application.PrintCommunication = False

'// 인쇄여백값을 받아옵니다.
varMargin = getPrintMargin(eMargin)

'// 인쇄용지 방향을 설정합니다.
If eOrient = xlPortrait Then
    lngOrient = 1
Else
    lngOrient = 2
End If

'// ExecuteExcel4Macro 의 Page.Setup 명령문 실행을 위한 문구를 입력합니다.
Head = """&L" & LHead & "&R" & RHead & """"     '// 페이지 머릿말입니다.
Foot = """&L" & LFoot & "&R" & RFoot & """"     '// 페이지 꼬릿말입니다.
pLeft = varMargin(0)                            '// 왼쪽여백
pRight = varMargin(1)                           '// 오른쪽여백
Top = varMargin(2)                              '// 윗여백
Bot = varMargin(3)                              '// 아래여백
Head_margin = varMargin(4)                      '// 머릿말여백
Foot_margin = varMargin(5)                      '// 꼬릿말여백
Hdng = 0                                        '// 행/열반복 출력여부 0 = 반복출력안함 1 = 반복출력
Grid = False                                    '// 눈금선출력여부
Notes = False                                   '// 메모출력여부
H_cntr = HCenter                                '// 가운데정렬
V_cntr = VCenter                                '// 중앙정렬
Orient = lngOrient                              '// 문서방향, 1 = 세로 2 = 가로
Paper_size = eSize                              '// 용지크기
Pg_num = 1                                      '// 페이지 시작번호
Pg_order = 1                                    '// 페이지번호 순서, 1 = 위-아래-우 2 = 좌-우-아래
Quality = ""                                    '// 인쇄품질 (dot-per-inch로 입력) (공백 = 자동)
bw_cells = False                                '// 흑백인쇄여부, TRUE = 글자/테두리 검정,배경 흰색 FALSE = 색깔
pScale = 100                                    '// 축소/확대비율 또는 TRUE (Fit to Page)

'// 여백을 없음으로 설정할 경우 머릿말/꼬릿말을 삭제하여 인쇄영역과 겹치지 않도록 합니다.
If eMargin = xlNone Then
    Head = """"""
    Foot = """"""
End If


'// ExecuteExcel4Macro 명령문을 실행합니다.
pSetup = "PAGE.SETUP(" & Head & ", " & Foot & ", " & pLeft & ", " & pRight & ", " & Top & ", " & Bot & ", "
pSetup = pSetup & Hdng & ", " & Grid & "," & H_cntr & "," & V_cntr & "," & Orient & ","
pSetup = pSetup & Paper_size & "," & pScale & ","
pSetup = pSetup & Pg_num & "," & Pg_order & "," & bw_cells & "," & Quality & ","
pSetup = pSetup & Head_margin & "," & Foot_margin & "," & Notes & ")"


Application.ExecuteExcel4Macro pSetup

'// ExecuteExcel4Macro에서는 '한 페이지에 행/열 맞추기' 기능이 지원되지 않습니다.
'// 따라서 시트의 PageSetup 속성으로 '페이지 행/열 맞추기 기능을 설정합니다.
With WS.PageSetup
    If HFit = True Then
        .FitToPagesWide = 1
    Else
        .FitToPagesWide = False
    End If
    
    If VFit = True Then
        .FitToPagesTall = 1
    Else
        .FitToPagesTall = False
    End If
End With

'// 인쇄설정 업데이트
Application.PrintCommunication = True

End Sub

