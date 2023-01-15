Attribute VB_Name = "Module1"
Option Explicit

'###############################################################
'오빠두엑셀 VBA 사용자 지정함수 (https://www.oppadu.com)
'▶ ImageLookup 함수 (워크시트 함수)
'▶ 폴더경로와 그림파일의 파일이름을 지정하여 '정확히일치' 또는 '유사일치'로 해당셀에 이미지를 출력합니다. (메모활용)
'▶ 인수 설명
'_____________ImageName     : 검색할 이미지파일의 이름 (확장자 포함 X)
'_____________FolderPath    : 이미지를 검색할 폴더경로
'_____________ExactMatch    : 그림파일 이름 정확히 일치 검색여부
'_____________NAValue       : 그림파일 없을 시 출력할 오류 메세지
'_____________ShowImageName : 그림파일 경로 출력여부
'_____________ShowBorder    : 그림파일의 테두리 출력여부
'_____________Extension     : 특정 그림파일 형식만 검색할경우, 원하는 확장자 지정
'▶ 사용된 기타 사용자지정함수
'_____________cvRng 함수
'_____________vbFileSearch 함수
'▶ 그외 참고사항
'###############################################################

Function ImageLookup(ImageName, _
                        Optional FolderPath = "", _
                        Optional ExactMatch As Boolean = True, _
                        Optional NAvalue As String = "-", _
                        Optional ShowImageName As Boolean = False, _
                        Optional ShowBorder As Boolean = False, _
                        Optional Extension = "png, jpeg, jpg, gif")
                        
                        
Dim rng As Range
Dim myComment As Comment
Dim sName As String: Dim FullPath As String

'// 함수가 입력된 셀
Set rng = Application.Caller
sName = CStr(cvRng(ImageName))

'// 함수가 입력된 셀의 메모 삭제
rng.ClearComments

'// 참조할 파일이름이 빈칸일 경우 오류메세지 반환
If sName = "" Then ImageLookup = NAvalue: Exit Function

On Error GoTo EmptyFolder:
'// 폴더경로 빈칸일 경우 해당 워크시트 폴더경로 반환
If FolderPath = "" Then FolderPath = rng.Parent.Parent.Path Else: FolderPath = cvRng(FolderPath)
On Error GoTo 0

'// 폴더경로안에 파일 검색
FullPath = vbFileSearch(sName, CStr(FolderPath), ExactMatch, CStr(Extension))

'// 폴더안에 파일 존재시 함수가 입력된 셀에 메모삽입 후 이미지 출력
If FullPath <> "-1" Then
    Set myComment = rng.AddComment
    
    With myComment
        .Visible = True

        If sName = "" Or ShowImageName = False Then
            .Text Text:=" "
        ElseIf ShowImageName = True Then
            .Text Text:=FullPath
        End If
    
        With .Shape
                .Left = rng.Left + 3
                .Top = rng.Top + 3
    
    
                .Width = rng.MergeArea.Width - 6
                .Height = rng.MergeArea.Height - 6
                
                .Fill.UserPicture FullPath
                .Line.ForeColor.RGB = rng.Interior.Color
                
                If ShowBorder = True Then .Line.Visible = msoTrue Else: .Line.Visible = msoFalse
        End With
    End With
    ImageLookup = ""
Else
'// 폴더안에 파일 없을 경우 오류메세지 출력
    ImageLookup = NAvalue
End If

Exit Function

EmptyFolder:
ImageLookup = NAvalue: Exit Function

End Function

