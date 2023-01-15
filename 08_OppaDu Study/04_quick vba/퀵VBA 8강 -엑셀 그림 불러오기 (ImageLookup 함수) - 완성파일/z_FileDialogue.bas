Attribute VB_Name = "z_FileDialogue"
Option Explicit

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ vbFileSearch 함수
'▶ 지정한 폴더에 특정 확장자를 가진 파일명을 유사일치/정확히 일치로 검색합니다.
'▶ 인수 설명
'_____________FileName      : 검색할 파일명입니다.
'_____________LookIn        : 조회할 폴더입니다.
'_____________Extension     : 특정 확장자만 조회합니다. (쉼표(,)로 구분)
'_____________ExactMatch    : 정확히일치 검색 여부입니다.
'_____________withPath      : True일 경우 결과값에 폴더경로를 출력합니다.
'▶ 사용된 기타 사용자지정함수
'_____________SplitFileExt 함수
'_____________IsInArray 함수
'_____________ListFiles 함수
'###############################################################
Function vbFileSearch(FileName As String, _
                        LookIn As String, _
                        Optional ExactMatch As Boolean = True, _
                        Optional Extension As String = "", _
                        Optional withPath As Boolean = True) As String

Dim sFullName As Variant
Dim sExts As Variant: Dim sExt As Variant
Dim vaArr As Variant: Dim vaRtn As Variant
Dim i As Long: Dim j As Long

vbFileSearch = "-1"
If Right(LookIn, 1) <> "\" Then LookIn = LookIn & "\"

vaArr = SplitFileExt(ListFiles(LookIn, False))
vaRtn = IsInArray(FileName, vaArr, ExactMatch, rtnArrayValue, 0)

If TypeName(vaRtn) = "Variant()" Then
    If Extension = "" Then
        If withPath = True Then vbFileSearch = LookIn & vaRtn(0, 0) & vaRtn(0, 1) Else: vbFileSearch = vaRtn(0, 0) & vaRtn(0, 1)
    Else
        sExts = Split(Extension, ",")
        For Each sExt In sExts
            For i = LBound(vaRtn) To UBound(vaRtn)
                If StrComp(CStr(vaRtn(i, 1)), CStr("." & Trim(sExt)), vbTextCompare) = 0 Then
                    If withPath = True Then
                        vbFileSearch = LookIn & vaRtn(i, 0) & vaRtn(i, 1)
                        Exit Function
                    Else
                        vbFileSearch = vaRtn(i, 0) & vaRtn(i, 1)
                        Exit Function
                    End If
                End If
            Next
        Next
        vbFileSearch = "-1"
    End If
Else
    vbFileSearch = vaRtn
End If

End Function

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ ListFiles 함수
'▶ 선택한 폴더의 파일목록을 배열로 반환합니다.
'▶ 인수 설명
'_____________sPath     : 파일목록을 출력할 폴더입니다.
'_____________withPath  : 폴더경로를 같이 출력할 여부를 결정합니다.
'###############################################################

Function ListFiles(sPath As String, Optional withPath As Boolean = False)

'// 각 변수를 생성합니다.
Dim arr As Variant
Dim i As Integer
Dim oFSO As Object: Dim oFolder As Object: Dim oFiles As Object: Dim oFile As Object

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(sPath)
Set oFiles = oFolder.Files

'// 변수를 설정합니다.
i = 1
If oFiles.Count = 0 Then Exit Function
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"

'// 폴더에 파일이 한개라도 존재시 배열 생성합니다.
ReDim arr(1 To oFiles.Count)

'// 각 파일을 돌아가며 Arr 배열로 반환합니다.
For Each oFile In oFiles
    If withPath = False Then arr(i) = oFile.Name Else: arr(i) = sPath & oFile.Name
    i = i + 1
Next

'// 배열을 결과값으로 출력합니다.
ListFiles = arr

End Function


'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ SplitFileExt 함수
'▶ 배열 또는 String으로 받아온 확장자가 포함된 파일경로를 파일이름과 확장자로 분리합니다. (2차원 배열, 0 - 파일명, 1 - 확장자)
'▶ 인수 설명
'_____________Files     : 확장자를 분리할 배열 또는 파일명입니다.
'###############################################################
Function SplitFileExt(Files As Variant) As Variant

Dim i As Long
Dim sFile As String
Dim vaArr As Variant

On Error GoTo ErrHandler:

'// 파일 타입을 확인합니다. 배열 또는 기타 문자열일 경우
If TypeName(Files) = "Variant()" Then
    '// 입력 형식이 배열일 경우
    ReDim vaArr(LBound(Files) To UBound(Files), 0 To 1)
    For i = LBound(Files) To UBound(Files)
        sFile = Files(i)
        vaArr(i, 0) = Left(sFile, InStrRev(sFile, ".") - 1)
        vaArr(i, 1) = Right(sFile, Len(sFile) - InStrRev(sFile, ".") + 1)
    Next
Else
    '// 입력 형식이 기타 문자열일 경우
    ReDim vaArr(0, 1)
    vaArr(0, 0) = Left(Files, InStrRev(Files, ".") - 1)
    vaArr(0, 1) = Right(Files, Len(Files) - InStrRev(Files, ".") + 1)
End If

SplitFileExt = vaArr


Exit Function

ErrHandler:
MsgBox "올바른 파일경로 또는 파일명을 입력하세요." & sFile
End

End Function
