Attribute VB_Name = "Chapter03"
Public title As String
Public startingsalary As Long
Public company As String

Sub rangepractice()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("3").Activate
    Range("A1").Value = 100
    [B2].Value = 500
    Range("D1:D8, E10:E20, A22:E22").Value = "Multiple cells one range code"
    Range("J1:J20").Value = "column J add 1 to 20 lower memory size"
    Range("A25:Z25").Value = "row 25 with quotes 25:25 A to Z lower memory size"
    Range("A27:Z28").Value = "rows 27 28 with quotes 27:28 A to Z lower memory size"
End Sub

Sub clearcontents()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("3").Activate
    Range("A:K").clearcontents
    Range("25:25").clearcontents
    Range("27:28").clearcontents
End Sub

Sub variablepractice()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("3").Activate
    Dim lastname As String
    Dim salary As Long
    Dim datehired As Date
    Dim anythinggoes As Variant
    Dim candrive As Boolean
    Dim toolongnumber As Double
    'title, startingsalary, company public variables
    title = "president"
    startingsalary = 1000000
    company = "acme"
    lastname = "Smith"
    datehired = "1/1/2010"
    [a29].Value = title
    [A30].Value = startingsalary
    [a31].Value = company
    Range("A32").Value = lastname
    Cells(33, 1).Value = datehired
End Sub

Sub divisions()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("3").Activate
    Range("A9").Value = 10 / 3                   'print 3.333...
    Range("A10").Value = 10 \ 3                  'print 3
    Range("A11").Value = 10 Mod 3                'print 1 the remainder
End Sub

Sub commonvbaconstants()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("3").Activate
    'carriage return
    Range("A15").Value = "First line" & vbCrLf & "second line"
    'tab which isn't working
    Range("A16").Value = "First word" & vbTab & "second word"
End Sub

