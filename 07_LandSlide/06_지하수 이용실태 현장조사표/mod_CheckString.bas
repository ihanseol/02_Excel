Attribute VB_Name = "mod_CheckString"
Option Explicit


Function CheckSubstring(str As String, chk As String) As Boolean
    
    If InStr(str, chk) > 0 Then
        ' The string contains "chk"
        CheckSubstring = True
    Else
        ' The string does not contain "chk"
        CheckSubstring = False
    End If
End Function

Function SS_StringCheck(str As String) As String
    
    ' 가정용 - 사설
    If CheckSubstring(str, "가정") Then SS_StringCheck = "g,"
    
    ' 일반용 - 사설
    If CheckSubstring(str, "일반") Then SS_StringCheck = "h,"
    
    ' 학교용 - 공공
    If CheckSubstring(str, "학교") Then SS_StringCheck = "i,"
        
    ' 민방위용 - 공공
    If CheckSubstring(str, "민방") Then SS_StringCheck = "j,"
    
    ' 공동주택용 - 사설
    If CheckSubstring(str, "공동") Then SS_StringCheck = "k,"
    
    ' 간이상수도 - 공공
    If CheckSubstring(str, "간이") Then SS_StringCheck = "l,"
    
    ' 농생활겸용 - 사설
    If CheckSubstring(str, "겸용") Then SS_StringCheck = "m,"
    
    ' 기타 - 사설
    If CheckSubstring(str, "기타") Then SS_StringCheck = "n,"

End Function

Function AA_StringCheck(str As String) As String
    
    ' 농업용은 전부 사설, 이중 허가공 - 공공
    If CheckSubstring(str, "전작") Then AA_StringCheck = "v,"
    If CheckSubstring(str, "답작") Then AA_StringCheck = "w,"
    If CheckSubstring(str, "원예") Then AA_StringCheck = "x,"
    
    If CheckSubstring(str, "축산") Then AA_StringCheck = "y,"
    If CheckSubstring(str, "양어") Then AA_StringCheck = "z,"
    If CheckSubstring(str, "기타") Then AA_StringCheck = "aa,"
    
End Function


Function II_StringCheck(str As String) As String
    
    ' 극가, 지방, 농공 - 공공
    If CheckSubstring(str, "국가") Then II_StringCheck = "p,"
    If CheckSubstring(str, "지방") Then II_StringCheck = "q,"
    If CheckSubstring(str, "농공") Then II_StringCheck = "r,"
    
    ' 자유입지, 기타 - 사설
    If CheckSubstring(str, "자유입지") Then II_StringCheck = "s,"
    If CheckSubstring(str, "기타") Then II_StringCheck = "t,"

End Function



Function SS_PublicCheck(str As String) As String
    
    ' 가정용 - 사설
    If CheckSubstring(str, "가정") Then SS_PublicCheck = "ab,"
    
    ' 일반용 - 사설
    If CheckSubstring(str, "일반") Then SS_PublicCheck = "ac,"
    
    ' 학교용 - 공공
    If CheckSubstring(str, "학교") Then SS_PublicCheck = "ab,"
        
    ' 민방위용 - 공공
    If CheckSubstring(str, "민방") Then SS_PublicCheck = "ab,"
    
    ' 공동주택용 - 사설
    If CheckSubstring(str, "공동") Then SS_PublicCheck = "ac,"
    
    ' 간이상수도 - 공공
    If CheckSubstring(str, "간이") Then SS_PublicCheck = "ab,"
    
    ' 농생활겸용 - 사설
    If CheckSubstring(str, "겸용") Then SS_PublicCheck = "ac,"
    
    ' 기타 - 사설
    If CheckSubstring(str, "기타") Then SS_PublicCheck = "ac,"

End Function

Function AA_PublicCheck(str As String) As String
    
    ' 농업용은 전부 사설, 이중 허가공 - 공공
    If CheckSubstring(str, "전작") Then AA_PublicCheck = "ac,"
    If CheckSubstring(str, "답작") Then AA_PublicCheck = "ac,"
    If CheckSubstring(str, "원예") Then AA_PublicCheck = "ac,"
    
    If CheckSubstring(str, "축산") Then AA_PublicCheck = "ac,"
    If CheckSubstring(str, "양어") Then AA_PublicCheck = "ac,"
    If CheckSubstring(str, "기타") Then AA_PublicCheck = "ac,"
    
End Function


Function II_PublicCheck(str As String) As String
    
    ' 극가, 지방, 농공 - 공공
    If CheckSubstring(str, "국가") Then II_PublicCheck = "ab,"
    If CheckSubstring(str, "지방") Then II_PublicCheck = "ab,"
    If CheckSubstring(str, "농공") Then II_PublicCheck = "ab,"
    
    ' 자유입지, 기타 - 사설
    If CheckSubstring(str, "자유입지") Then II_PublicCheck = "ac,"
    If CheckSubstring(str, "기타") Then II_PublicCheck = "ac,"

End Function


