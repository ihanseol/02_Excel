VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Sheets("Aggregate2").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
    Dim fname As String
    Dim nofWell, i As Integer
    Dim Q() As Double '¾ç¼ö·®
    Dim Natural() As Double
    Dim Stable() As Double
    Dim Recover() As Double
    
    Dim Radius() As Double
    Dim deltaS() As Double
    Dim daeSoo() As Double
    Dim T() As Double
    Dim S() As Double
    
    Sheets("Well").Select
    Sheets("Well").Range("A30").Select
    Selection.End(xlUp).Select
    
    nofWell = GetNumeric2(Selection.value)
    
    Sheets("Aggregate2").Select
    
    ReDim Q(1 To nofWell)
    ReDim Natural(1 To nofWell)
    ReDim Stable(1 To nofWell)
    ReDim Recover(1 To nofWell)
    
    ReDim Radius(1 To nofWell)
    ReDim deltaS(1 To nofWell)
    ReDim daeSoo(1 To nofWell)
    ReDim T(1 To nofWell)
    ReDim S(1 To nofWell)
    
    For i = 1 To nofWell
        fname = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fname) Then
            MsgBox "Please open the yangsoo data ! " & fname
            Exit Sub
        End If
        
        Q(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c4").value
        Natural(i) = Workbooks(fname).Worksheets("SkinFactor").Range("i4").value
        Stable(i) = Workbooks(fname).Worksheets("SkinFactor").Range("i6").value
        Recover(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c10").value
        
        Radius(i) = Workbooks(fname).Worksheets("SkinFactor").Range("e4").value
        
        deltaS(i) = Workbooks(fname).Worksheets("SkinFactor").Range("b4").value
        daeSoo(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c16").value
        T(i) = Workbooks(fname).Worksheets("SkinFactor").Range("d5").value
        S(i) = Workbooks(fname).Worksheets("SkinFactor").Range("e10").value
        
    Next i

    Call WriteWellData(Q, Natural, Stable, Recover, Radius, deltaS, daeSoo, T, S, nofWell)
    Application.CutCopyMode = False
End Sub


Sub WriteWellData(Q As Variant, Natural As Variant, Stable As Variant, Recover As Variant, Radius As Variant, deltaS As Variant, daeSoo As Variant, T As Variant, S As Variant, ByVal nofWell As Integer)
    
    Dim i As Integer
    
    For i = 1 To nofWell
        Range("D" & (i + 2)).value = "W-" & i
        
        Range("e" & (i + 2)).value = Q(i)
        Range("l" & (i + 2)).value = Q(i)
        
        Range("f" & (i + 2)).value = Natural(i)
        Range("g" & (i + 2)).value = Stable(i)
        Range("h" & (i + 2)).value = Stable(i) - Natural(i)
        
        Range("i" & (i + 2)).value = Radius(i)
        Range("j" & (i + 2)).value = deltaS(i)
        
        Range("m" & (i + 2)).value = Radius(i)
        Range("n" & (i + 2)).value = Radius(i)
        Range("o" & (i + 2)).value = daeSoo(i)
        Range("p" & (i + 2)).value = T(i)
        Range("q" & (i + 2)).value = S(i)
        
        Range("s" & (i + 2)).value = Stable(i)
        Range("t" & (i + 2)).value = Recover(i)
        Range("u" & (i + 2)).value = Stable(i) - Recover(i)
        
        
    Next i
   
End Sub

Private Sub CommandButton3_Click()
    Range("d3:u29").Select
    Selection.ClearContents
    
    Range("B24").Select
    Application.CutCopyMode = False
End Sub


