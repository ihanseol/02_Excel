Attribute VB_Name = "modOpenAPI"
Sub TransPose30Year()

    Dim i, j As Integer
    Dim i1, i2 As Integer
    Dim sYear, eYear As Integer
    
    
    Range("C1").Select
    Selection.End(xlDown).Select
    
    eYear = Year(Now()) - 1
    sYear = eYear - 29
    
       
    For i = 1 To 30
    
        i1 = 12 * (i - 1) + 9
        i2 = i1 + 11
        
        Range("C" & CStr(i1) & ":C" & CStr(i2)).Select
        Selection.Copy
        
        j = i + 8
        Range("G" & CStr(j)).Select
        
        Range("F" & CStr(j)).Value = sYear + i - 1
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
            
    Next i
        
        
End Sub

Sub test()
    On Error GoTo Retry
        WebDriver.FindElementById ("element")
    Exit Sub

    Dim i As Integer
: Retry
        WebDriver.Wait (500)
        i = i + 1
        If i = 20 Then
            On Error GoTo 0
        End If
    Resume
End Sub





Sub CallOpenAPI()
    Dim strURL As String
    Dim strResult As String
    
    Dim objHttp As New WinHttpRequest
    
    strURL = "Open API 주소를 입력하세요"
    objHttp.Open "GET", strURL, False
    objHttp.send
    
    
    If objHttp.Status = 200 Then '성공했을 경우
        strResult = objHttp.responseText
        
        'XML로 연결
        Dim objXml As MSXML2.DOMDocument60
        Set objXml = New DOMDocument60
        objXml.LoadXML (strResult)
        
        '노드 연결
        Dim nodeList As IXMLDOMNodeList
        Dim nodeRow As IXMLDOMNode
        Dim nodeCell As IXMLDOMNode
        Dim nRowCount As Integer
        Dim nCellCount As Integer
        
        Set nodeList = objXml.SelectNodes("/response/fields/field")
        
        nRowCount = Range("A60000").End(xlUp).Row
        For Each nodeRow In nodeList
        nRowCount = nRowCount + 1
        
        nCellCount = 0
        For Each nodeCell In nodeRow.ChildNodes
        nCellCount = nCellCount + 1
        '엑셀에 값 반영
        Cells(nRowCount, nCellCount).Value = nodeCell.Text
        Next nodeCell
        
        Next nodeRow
    
    Else
        MsgBox "접속에 에러가 발생했습니다"
    
    End If
End Sub





