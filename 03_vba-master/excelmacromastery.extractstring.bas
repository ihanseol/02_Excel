Attribute VB_Name = "extractstring"
Sub extractvbacode()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("extractstring").Activate
    Dim brownfox As String
    brownfox = "The quick brown fox jumped over the lazy dog"
    Range("A1:A8").Value = brownfox
    Range("B1").Value = Left(Range("A1").Value, 7) 'print The qui
    Range("B2").Value = Right(Range("A2").Value, 7) 'print azy dog
    Range("B3").Value = Mid(Range("A3").Value, 11, 5) 'print brown
    Range("B4").Value = Split(Range("A4").Value, " ")(0) 'print The
    Range("B5").Value = Split(Range("A5").Value, " ")(1) 'print quick
    Range("B6").Value = Split(Range("A6").Value, " ")(2) 'print brown
    Range("B7").Value = Split(Range("A7").Value, " ") 'print The
    Dim v, lastname As Variant
    v = Split(Range("A8").Value, " ")
    Range("A10").Value = v                       'print The
    lastname = v(UBound(v))
    Range("A11").Value = lastname                'print dog
End Sub

Sub instring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("extractstring").Activate
    'InStr(start position, find text, find what in text) or InStr(find text, find what in text)
    If InStr(Range("A13").Value, "Henry") > 0 Then
        Range("B13").Value = "Found Henry"
    End If
End Sub

Sub instringfirstandlastname()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("extractstring").Activate
    For i = 15 To 24
        Range("B" & i).Value = Left(Range("A" & i), InStr(Range("A" & i).Value, " ") - 1)
        Range("C" & i).Value = Right(Range("A" & i), Len(Range("A" & i)) - InStr(Range("A" & i).Value, " "))
    Next i
End Sub

Sub instringmiddlename()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("extractstring").Activate
    Dim threepartname As String
    Dim firstpart, secondpart As Long
    threepartname = "John Henry Smith"
    firstpart = InStr(threepartname, " ") + 1
    secondpart = InStr(firstpart, threepartname, " ")
    count = secondpart - firstpart
    MsgBox Mid(threepartname, firstpart, count)

    'check valid filename ...AA...1234....pdf
    Dim e As Range
    For Each e In Range("A40:A44")
        If InStr(e, "1234") > InStr(e, "AA") And Right(e, 4) = ".pdf" Then
            e.offset(0, 1) = "Valid"
        Else
            e.offset(0, 1) = "Invalid"
        End If
    Next e
    'VBA has Pattern Matching
    Dim f As Range
    Dim pattern As String
    pattern = "*AA*1234*.pdf"
    For Each f In Range("A40:A44")
        f.offset(0, 2).Value = f Like pattern    'print TRUE or FALSE
    Next f
End Sub

Sub splitfunction()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("extractstring").Activate
    'simple example
    Dim henry As String
    henry = "John Henry Smith"
    Range("A26") = Split(henry, " ")(0)          'print John
    Range("A27") = Split(henry, " ")(1)          'print Henry
    Range("A28") = Split(henry, " ")(2)          'print Smith
    Range("B26") = Split(henry, ",")(0)          'print John Henry Smith
    henry = "John,Henry,Smith"
    Range("B27") = Split(henry, ",")(0)          'print John
    Range("B28") = Split(henry, ",")(1)          'print Henry
    Range("B29") = Split(henry, ",")(2)          'print Smith
    'array variable
    Dim henry2, arr() As String
    henry2 = "John Henry Smith"
    arr = Split(henry2, " ")
    Range("C26") = arr(0)                        'print John
    Range("C27") = arr(1)                        'print Henry
    Range("C28") = arr(2)                        'print Smith

    'extract numbers between two underscores
    'FYI:  excel formula =MID(A31,FIND("_",A31,1)+1,FIND("_",A31,FIND("_",A31,1)+1)-(FIND("_",A31,1))-1)*1
    Range("B31") = Split(Range("A31"), "_")(1)   'print 23476
    Range("B32") = Split(Range("A32"), "_")(1)   'print 987
    Range("B33") = Split(Range("A33"), "_")(1)   'print 12223
    'easier using for loop in a table
    Dim c As Range
    For Each c In Range("A31:A33")
        c.offset(0, 2).Value = Split(c, "_")(1)  'print the numbers two cells to the right
    Next c

    'extract IP Address Range and validate between a range
    'FYI:  excel formula =IF(AND(MID(A35,FIND(" ",A35,1)+5,2)*1>=16,MID(A35,FIND(" ",A35,1)+5,2)*1<=31),"Valid","Invalid")
    Range("B35") = Split(Range("A35"), ".")(1)   'print 16
    Range("B36") = Split(Range("A36"), ".")(1)   'print 25
    Range("B37") = Split(Range("A37"), ".")(1)   'print 14
    Range("B38") = Split(Range("A38"), ".")(1)   'print 32
    'easier using for loop in a table
    Dim d As Range
    For Each d In Range("A35:A38")
        If Split(d, ".")(1) >= 16 And Split(d, ".")(1) <= 31 Then
            d.offset(0, 2).Value = "Valid"       'print the results two cells to the right
        Else
            d.offset(0, 2).Value = "Invalid"     'print the results two cells to the right
        End If
    Next d
End Sub

