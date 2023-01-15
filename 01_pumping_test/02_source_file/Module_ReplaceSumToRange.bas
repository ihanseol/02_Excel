Attribute VB_Name = "Module_ReplaceSumToRange"
Option Explicit

Function CleanString(strIn As String) As String

    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    
    With objRegex
     .Global = True
     .Pattern = "[^\d]+"
      CleanString = .Replace(strIn, vbNullString)
    End With
    
End Function



'Extract start and end number in range
'Exam : "A12:A15"
'12,15

Function ExtractStartEnd_FromRange(rng As Range) As String()

    Dim cellValue As String
    Dim colonParam As Integer
        
    Dim retValue(2) As String
   

    cellValue = rng.Address
    cellValue = Replace(cellValue, "$", "")
    
    colonParam = InStr(cellValue, ":")
    
    retValue(0) = CleanString(Mid(cellValue, 1, colonParam - 1))
    retValue(1) = CleanString(Mid(cellValue, colonParam + 1, Len(cellValue)))
    
    ExtractStartEnd_FromRange = retValue
    
End Function



Function Extract_Ranges_From_Formula(rng As Range) As String()
    
    Dim rCell As Range
    Dim cellValue As String
    
    Dim openingParen As Integer
    Dim closingParen As Integer
    Dim colonParam As Integer
    
    Dim retValue() As String
    Dim i As Long: i = 0
        
       
    ReDim retValue(0 To rng.Count - 1, 0 To 1)
    
    
    For Each rCell In rng
        cellValue = rCell.Formula
    
        openingParen = InStr(cellValue, "(")
        colonParam = InStr(cellValue, ":")
        closingParen = InStr(cellValue, ")")
        
        retValue(i, 0) = Mid(cellValue, openingParen + 1, colonParam - openingParen - 1)
        retValue(i, 1) = Mid(cellValue, colonParam + 1, closingParen - colonParam - 1)
            
        i = i + 1
    Next rCell
    
    Extract_Ranges_From_Formula = retValue

End Function


Private Sub ExtractTest()

    Dim rng As Range
    Dim ret As Variant
    
    Set rng = Range("E25")
    Set rng = Union(rng, Range("G25"))
    Set rng = Union(rng, Range("I25"))
    
    
    ret = Extract_Ranges_From_Formula(rng)
    Debug.Print ret(0, 0)

End Sub


' Uses Range.Find to get a range of all find results within a worksheet
' Same as Find All from search dialog box
'
Function FindAll(rng As Range, What As Variant, Optional LookIn As XlFindLookIn = xlValues, Optional LookAt As XlLookAt = xlWhole, Optional SearchOrder As XlSearchOrder = xlByColumns, Optional SearchDirection As XlSearchDirection = xlNext, Optional MatchCase As Boolean = False, Optional MatchByte As Boolean = False, Optional SearchFormat As Boolean = False) As Range
    Dim SearchResult As Range
    Dim firstMatch As String
    With rng
        Set SearchResult = .Find(What, , LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat)
        If Not SearchResult Is Nothing Then
            firstMatch = SearchResult.Address
            Do
                If FindAll Is Nothing Then
                    Set FindAll = SearchResult
                Else
                    Set FindAll = Union(FindAll, SearchResult)
                End If
                Set SearchResult = .FindNext(SearchResult)
            Loop While Not SearchResult Is Nothing And SearchResult.Address <> firstMatch
        End If
    End With
End Function


Function RangeToStringArray(myRange As Range) As String()
    ReDim strArray(myRange.Cells.Count - 1) As String
    Dim idx As Long
    Dim c As Range
    
    For Each c In myRange
        strArray(idx) = c.Text
        idx = idx + 1
    Next c

    RangeToStringArray = strArray
End Function


Function SerializeRange(theRange As Excel.Range) As String()
    Dim cell As Range
    Dim values() As String
    Dim i As Integer
    
    i = 0
    ReDim values(theRange.Cells.Count)
    
    For Each cell In theRange
           values(i) = cell.Address
           i = i + 1
    Next cell
    
    SerializeRange = values
    
End Function

Private Sub serialize_test()
    Dim ar As Variant
    
    ar = SerializeRange(Range("C11:C14"))
    Debug.Print ar(0)
 End Sub



 Sub ReplaceSUMtoEachCell()

    Dim ws As Worksheet
    Dim iList As Range, iName As Range
        
    Dim strRange As Variant
    Dim strResult As String
    Dim i As Integer
    
    Dim cellValue, Value As String
    Dim openingParen As Integer
    Dim closingParen As Integer

    Set ws = ThisWorkbook.ActiveSheet
    Set iList = FindAll(ws.UsedRange, "SUM", xlFormulas, xlPart)
    
    For Each iName In iList
    
        cellValue = iName.Formula
    
        openingParen = InStr(cellValue, "(")
        closingParen = InStr(cellValue, ")")
    
    
        Value = Mid(cellValue, openingParen + 1, closingParen - openingParen - 1)
            
        strResult = "="
        strRange = SerializeRange(Range(Value))
        
        For i = 0 To UBound(strRange)
            strResult = strResult & strRange(i) & "+"
        Next i
        
        strResult = Left(strResult, Len(strResult) - 2)
               
        strResult = Replace(strResult, "$", "")
        iName.Formula = strResult

    Next iName
    
End Sub



    
    
    
    

