Attribute VB_Name = "z_Array"
Option Explicit

Public Enum xlArrayReturnType
    rtnSequence = 1
    rtnValue = 2
    rtnArrayValue = 3
End Enum

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ IsInArray 함수
'▶ 배열안에 선택한 값의 존재여부를 확인합니다. 4가지 값으로 반환 가능합니다. (순번, 값, 배열)
'▶ 인수 설명
'_____________FindValue     : 배열에서 찾을 값입니다.
'_____________vaArray       : 값을 검색할 배열입니다.
'_____________ExactMatch    : 정확히일치/유사일치 검색여부입니다.
'_____________returnType    : 반환형식을 결정합니다. (Public Enum)
'_____________Dimension     : 값을 검색할 배열의 차원입니다.
'_____________vbCompare     : 대소문자 구분 여부입니다. (vbTextCompare : 대소문자구분X, vbBinaryCompare : 대소문자구분)
'▶ 사용된 기타 사용자지정함수
'_____________ArrayDimension 함수
'▶ 그외 참고사항
'_____________xlArrayReturnType PublicEnum이 사용되었습니다.
'###############################################################

Function IsInArray(FindValue As Variant, _
                    vaArray As Variant, _
                    Optional ExactMatch As Boolean = True, _
                    Optional returnType As xlArrayReturnType = rtnValue, _
                    Optional Dimension As Integer = 0, _
                    Optional vbCompare As VbCompareMethod = vbTextCompare) As Variant

Dim i As Long: Dim j As Long
Dim dicArr As Object: Dim dicKey As Variant: Dim dicKeys As Variant
Dim rtnArr As Variant
Dim ArrDim As Integer

ArrDim = ArrayDimension(vaArray)
Set dicArr = CreateObject("Scripting.Dictionary")

'// IsInArray 초기값을 설정합니다.
IsInArray = -1

'// 정확히일치일 경우
If ExactMatch = True Then
        '// 값이 있을시 ReturnType에 따라 반환합니다. (정확히 일치이므로 결과는 무조건 1개만 반환)
        If ArrDim = 1 Then
        '// 배열이 1차원일 경우
            For i = LBound(vaArray) To UBound(vaArray)
                If StrComp(FindValue, vaArray(i), vbCompare) = 0 Then
                    If returnType = rtnValue Or returnType = rtnArrayValue Then IsInArray = vaArray(i): Exit For
                    If returnType = rtnSequence Then IsInArray = i: Exit For
                End If
            Next i
        Else
        '// 배열이 2차원일 경우
           For i = LBound(vaArray) To UBound(vaArray)
                If StrComp(FindValue, vaArray(i, Dimension), vbCompare) = 0 Then
                    If returnType = rtnValue Then IsInArray = vaArray(i, Dimension): Exit For
                    If returnType = rtnSequence Then IsInArray = i: Exit For
                    If returnType = rtnArrayValue Then dicArr.Add i, i: Exit For
                    Exit For
                End If
            Next i
            
            '// Dictionary에 값이 하나라도 존재시
            If dicArr.Count > 0 Then
                '배열 Redim
                ReDim rtnArr(0, 0 To ArrDim - 1)
                    '// Dictionary 에서 받아온 각 값을 배열로 옮깁니다
                    For Each dicKey In dicArr.keys
                        For j = LBound(vaArray, 2) To UBound(vaArray, 2)
                            rtnArr(0, j) = vaArray(dicKey, j)
                        Next
                    Next
                '// 배열로 결과를 출력합니다
                IsInArray = rtnArr
            End If
        End If
   
'// 유사일치일 경우
Else
    If ArrDim = 1 Then
        '// 배열이 1차원일 경우
        For i = LBound(vaArray) To UBound(vaArray)
        '// 값이 있을시 ReturnType에 따라 결과를 반환합니다
            If InStr(1, vaArray(i), FindValue) > 0 Then
                If returnType = rtnValue Then IsInArray = vaArray(i): Exit For
                If returnType = rtnSequence Then IsInArray = i: Exit For
                If returnType = rtnArrayValue Then dicArr.Add i, i
            End If
        Next i
        
        '// 유사일치 값이 하나라도 존재시
        If dicArr.Count > 0 Then
            '배열 Redim
            ReDim rtnArr(0 To dicArr.Count - 1)
            i = 0
                '// Dictionary 에서 받아온 각 값을 배열로 옮깁니다
                For Each dicKey In dicArr.keys
                        rtnArr(i) = vaArray(dicKey)
                    i = i + 1
                Next
            '// 배열로 결과를 출력합니다
            IsInArray = rtnArr
        End If
    Else
        '// 배열이 2차원일 경우
        For i = LBound(vaArray) To UBound(vaArray)
        '// 값이 있을시 ReturnType에 따라 결과를 반환합니다
            If InStr(1, vaArray(i, Dimension), FindValue) > 0 Then
                If returnType = rtnValue Then IsInArray = vaArray(i, Dimension): Exit For
                If returnType = rtnSequence Then IsInArray = i: Exit For
                If returnType = rtnArrayValue Then dicArr.Add i, i
            End If
        Next i
        
        '// 유사일치 값이 하나라도 존재시
        If dicArr.Count > 0 Then
            '배열 Redim
            ReDim rtnArr(0 To dicArr.Count - 1, 0 To ArrDim - 1)
            i = 0
                '// Dictionary 에서 받아온 각 값을 배열로 옮깁니다
                For Each dicKey In dicArr.keys
                    For j = LBound(vaArray, 2) To UBound(vaArray, 2)
                        rtnArr(i, j) = vaArray(dicKey, j)
                    Next
                    i = i + 1
                Next
            '// 배열로 결과를 출력합니다
            IsInArray = rtnArr
        End If
        
        End If
End If
  
End Function


'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ ArrayDimension 함수
'▶ 배열의 차원수를 반환합니다.
'▶ 인수 설명
'_____________vaArray     : 차원을 검토할 배열을 입력합니다.
'###############################################################
Function ArrayDimension(vaArray As Variant) As Integer

Dim i As Integer: Dim x As Integer

On Error Resume Next

Do
    i = i + 1
    x = UBound(vaArray, i)
Loop Until Err.Number <> 0

Err.Clear

ArrayDimension = i - 1

End Function

