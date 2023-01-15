Attribute VB_Name = "stringvba"
Sub stringquickguide()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    'Quick Guide to String Functions RM:  There's a lot.
    'Append two or more strings: Format or "&"
    'Build a string from an array: Join
    'Compare - normal: StrComp or "="
    'Compare - pattern: Like
    'Convert to a string: CStr, Str
    'Convert string to date: Simple: CDate; Advanced: Format
    'Convert string to number: Simple: CLng, CInt, CDbl, Val; Advanced: Format
    'Convert to unicode, wide, narrow: StrConv
    'Convert to upper/lower case: StrConv, UCase, LCase
    'Extract part of a string: Left, Right, Mid
    'Format a string: Format
    'Find characters in a string: InStr, InStrRev
    'Generate a string: String
    'Get length of a string: Len
    'Remove blanks: LTrim, RTrim, Trim
    'Replace part of a string: Replace
    'Reverse a string: StrReverse
    'Parse string to array: Split
End Sub

Sub extractstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    Dim name As String
    name = "John Thomas Smith"
    Range("A4:A6") = name
    Range("C4") = Left(Range("A4"), 4)           'print John
    Range("C5") = Right(Range("A5"), 5)          'print Smith
    Range("C6") = Mid(Range("A6"), 6, 6)         'print Thomas
End Sub

Sub searchstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    'InStr and InStrRev are VBA functions used to search through strings for a substring.
    'If the search string is found then the position(from the start of the check string) of the _
    search string is returned.
    'If the search string is not found then zero is returned. If either string is null _
    then null is returned.
    Dim name As String
    name = "John Thomas Smith"
    Range("A8:A10") = name
    Range("C8") = InStr(Range("A8"), "n")        'print 4
    Range("C9") = InStr(Range("A8"), "z")        'print 0
    Range("C10") = InStr(3, Range("A8"), "o")    'print 8.  Start search at position 3

    'The InStrRev function is the same as InStr except that it searches from the end of the string.
    'The position returned is the position from the start.
    name = "John Thomas Smith"
    Range("A12:A16") = name
    Range("C12") = InStr(Range("A12"), "J")      'print 1
    Range("C13") = InStrRev(Range("A13"), "J")   'print 1
    Range("C14") = InStr(Range("A14"), "h")      'print 3
    Range("C15") = InStrRev(Range("A15"), "h")   'print 17
    Range("C16") = InStrRev(Range("A16"), "h", 8) 'print 7.  Starts at position 8 as end of string
End Sub

Sub trimstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    Dim name As String
    name = "     John Smith     "
    Debug.Print (LTrim(name))
    Debug.Print (RTrim(name))
    Debug.Print (Trim(name))
End Sub

Sub lenstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    Dim name As String
    name = "John Thomas Smith"
    Debug.Print Len(name)                        'print 17
    Debug.Print StrReverse(name)                 'print htimS samohT nhoJ
End Sub

Sub comparestring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    name = "John Thomas Smith"
    Range("A18:A21") = name
    Range("C18").Value = StrComp(Range("A18"), name, vbTextCompare) 'print 0 or True
    Range("C19").Value = StrComp(Range("A19"), ("JOhn Thomas Smith"), vbTextCompare) 'print 0 or True
    Range("C20").Value = StrComp(Range("A20"), ("JOhn Thomas Smith"), vbBinaryCompare) 'print 1 or False
    Range("C21").Value = StrComp(Range("A21"), ("john thomas smith"), vbBinaryCompare) 'print -1 or False
    '0 strings match, -1 string1 less than string2, 1 string1 greater than string2, null either string null
    'vbTextCompare is case insensitive.
    'vbBinaryCompare is case sensitive.
    'You can use the Option Compare setting instead of having to use this parameter each time.
    'Option Compare is set at the top of a Module.
    'Option Compare Text: makes vbTextCompare the default Compare argument.
    'Option Compare Binary: Makes vbBinaryCompare the default Compare argument.  Default _
    if no Option Compare is set.

    'You can also use the equals sign to compare strings.  Returns True or False.  Can't Compare.
    'Also not equal sign <>
    name = "John Thomas Smith"
    Range("A23:A25") = name
    Range("C23") = Range("A23") = "John Thomas Smith" 'print TRUE
    Range("C24") = Range("A24") = UCase("John Thomas Smith") 'print FALSE
    Range("C25") = Range("A25") = Null           'print *nothing* or Null
End Sub

Sub patternstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    ' True abc first, not def second, any single character third, any single number four, anything fifth
    Debug.Print 1; "apY6X" Like "[abc][!def]?#X*"

    ' True - any combination of chars after x is valid
    Debug.Print 2; "apY6Xsf34FAD" Like "[abc][!def]?#X*"

    ' False - char d not in [abc]
    Debug.Print 3; "dpY6X" Like "[abc][!def]?#X*"

    ' False - 2nd char e is in [def]
    Debug.Print 4; "aeY6X" Like "[abc][!def]?#X*"

    ' False - A at position 4 is not a digit
    Debug.Print 5; "apYAX" Like "[abc][!def]?#X*"

    ' False - char at position 5 must be X
    Debug.Print 1; "apY6Z" Like "[abc][!def]?#X*"
End Sub

Sub replacestring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    ' Replaces all the question marks with(?) with semi colons(;)
    Debug.Print Replace("A?B?C?D?E", "?", ";")
    ' Replace Smith with Jones
    Debug.Print Replace("Peter Smith,Ann Smith", "Smith", "Jones")
    ' Replace AX with AB
    Debug.Print Replace("ACD AXC BAX", "AX", "AB")
    ' Replaces first question mark only
    Debug.Print Replace("A?B?C?D?E", "?", ";", count:=1)
    ' Replaces first three question marks
    Debug.Print Replace("A?B?C?D?E", "?", ";", count:=3)
    ' Use original string from position 4
    Debug.Print Replace("A?B?C?D?E", "?", ";", Start:=4) 'print ;C;D;E
    ' Use original string from position 8
    Debug.Print Replace("AA?B?C?D?E", "?", ";", Start:=8) 'print D;E
    ' No item replaced but still only returns last 2 characters
    Debug.Print Replace("ABCD", "X", "Y", Start:=3) 'print CD
    ' Replace capital A's only
    Debug.Print Replace("AaAa", "A", "X", Compare:=vbBinaryCompare) 'print XaXa
    ' Replace All A's
    Debug.Print Replace("AaAa", "A", "X", Compare:=vbTextCompare) 'print XXXX
    'vbTextCompare is case insensitive.
    'vbBinaryCompare is case sensitive.

    'multiple replace text write multiple replace VBA code
    Dim newString As String
    ' Replace A with X
    newString = Replace("ABCD ABDN", "A", "X")
    Debug.Print newString                        'print XBCD XBDN
    ' Now replace B with Y in new string
    newString = Replace(newString, "B", "Y")
    Debug.Print newString                        'print XYCD XYDN
    'next multiple replace text
    newString = Replace(Replace("ABCD ABDN", "A", "X"), "B", "Y")
    Debug.Print newString                        'print XYCD XYDN
End Sub

Sub convertstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    Range("B27") = Str(Range("A27").Value)       'print 12.99 as a number
    Range("B28") = Str(12.99)                    'print 12.99 as a number
End Sub

Sub valstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    'The value function convert numeric parts of a string to the correct number type.
    'The Val function converts the ..rst numbers it meets. Once it meets letters in a string it stops.
    'If there are only letters then it returns zero as the value.
    ' Prints 45
    Debug.Print Val("45 New Street")
    ' Prints 45
    Debug.Print Val("     45 New Street")
    ' Prints 0
    Debug.Print Val("New Street 45")
    ' Prints 12
    Debug.Print Val("12 f 34")
End Sub

Sub stringstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    'The String function is used to generate a string of repeated characters.
    'The first argument is the number of times to repeat it, the second argument is the character.
    ' Prints: AAAAA
    Debug.Print String(5, "A")
    ' Prints: >>>>>
    Debug.Print String(5, 62)
    ' Prints: (((ABC)))
    Debug.Print String(3, "(") & "ABC" & String(3, ")")
End Sub

Sub convertcasestring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    'Convert Case upper case convert case lower case convert proper case
    Dim s As String
    s = "Mary had a little lamb"
    Range("A29").Value = UCase(s)                'print MARY HAD A LITTLE LAMB
    Range("B29").Value = StrConv(s, vbUpperCase) 'print MARY HAD A LITTLE LAMB
    Range("A30").Value = LCase(s)                'print mary had a little lamb
    Range("B30").Value = StrConv(s, vbLowerCase) 'print mary had a little lamb
    Range("A31").Value = StrConv(s, vbProperCase) 'print Mary Had A Little Lamb
End Sub

Sub splitstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    Dim arr() As String
    Dim name As Variant
    arr = Split(Range("A33"), ",")
    Range("A34").Select
    For Each name In arr
        ActiveCell.Value = name
        ActiveCell.offset(1, 0).Select
    Next
End Sub

Sub joinstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    Dim beatlessong(0 To 5) As String
    Range("A34").Select
    For n = 0 To 5
        beatlessong(n) = ActiveCell.Value
        ActiveCell.offset(1, 0).Select
    Next
    Range("A40").Value = Join(beatlessong, ",")
End Sub

Sub formatstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("string").Activate
    Dim datestring As String
    datestring = "12/31/2015 10:15:45"
    ' Prints: 12 31 15
    Debug.Print Format(datestring, "MM DD YY")
    ' Prints: Thu 31 Dec 2015
    Debug.Print Format(datestring, "DDD DD MMM YYYY")
    ' Prints: Thursday 31 December 2015
    Debug.Print Format(datestring, "DDDD DD MMMM YYYY")
    ' Prints: 10:15
    Debug.Print Format(datestring, "HH:MM")
    ' Prints: 10:15:45 AM
    Debug.Print Format(datestring, "HH:MM:SS AM/PM")
    ' Prints: 50.00%
    Debug.Print Format(0.5, "0.00%")
    ' Prints: 023.45
    Debug.Print Format(23.45, "00#.00")
    ' Prints: 23,000
    Debug.Print Format(23000, "##,000")
    ' Prints: 023,000
    Debug.Print Format(23000, "0##,000")
    Range("A42").Value = Format(23000, "0##,000") 'print 23,000.  Not 023,000
    ' Prints: $23.99
    Debug.Print Format(23.99, "$#0.00")
End Sub


