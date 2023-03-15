VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()
    Sheets("aggWhpa").Visible = False
    Sheets("Well").Select
End Sub


Function getDirectionFromWell(i) As Integer

    If Sheets(CStr(i)).Range("k12").Font.Bold Then
        getDirectionFromWell = Sheets(CStr(i)).Range("k12").value
    Else
        getDirectionFromWell = Sheets(CStr(i)).Range("l12").value
    End If

End Function



Private Sub CommandButton2_Click()
' Collect Data

    Dim fname As String
    Dim nofWell, i As Integer
    
    Dim Q() As Double
    Dim daeSoo() As Double
    
    Dim T1() As Double
    Dim S1() As Double
    
    Dim direction() As Integer
    Dim gradient() As Double
    
    
    
    nofWell = GetNumberOfWell()
    If ActiveSheet.name <> "aggWhpa" Then Sheets("aggWhpa").Select
    
    ReDim Q(1 To nofWell) As Double
    ReDim daeSoo(1 To nofWell) As Double
    
    ReDim T1(1 To nofWell) As Double
    ReDim S1(1 To nofWell) As Double
    
    ReDim direction(1 To nofWell) As Integer
    ReDim gradient(1 To nofWell) As Double
      

    ' --------------------------------------------------------------------------------------
    
    For i = 1 To nofWell
        
        Sheets(CStr(i)).Select
        
        Q(i) = Sheets(CStr(i)).Range("c16").value
        daeSoo(i) = Sheets(CStr(i)).Range("c14").value
        
        T1(i) = Sheets(CStr(i)).Range("e7").value
        S1(i) = Sheets(CStr(i)).Range("g7").value
        
        direction(i) = getDirectionFromWell(i)
        gradient(i) = Sheets(CStr(i)).Range("k18").value
        
    Next i


    Sheets("aggWhpa").Select
    Call WriteWellData(Q, daeSoo, T1, S1, direction, gradient, nofWell)
    
    Range("a1").Select
    Application.CutCopyMode = False
    
End Sub



Private Sub WriteWellData(Q As Variant, daeSoo As Variant, T1 As Variant, S1 As Variant, direction As Variant, gradient As Variant, nofWell As Variant)
    Dim i As Integer
    Dim t_sum As Double
    Dim daesoo_sum As Double
    Dim gradient_sum As Double
    Dim direction_sum As Double
    
    t_sum = 0
    daesoo_sum = 0
    gradient_sum = 0
    direction_sum = 0
    
    With Range("c4:o17")
        .value = ""
    End With
            
    
    
    For i = 1 To nofWell
    
        Cells(3 + i, "c").value = "W-" & CStr(i)
        
        Cells(3 + i, "e").value = Q(i)
        Cells(3 + i, "f").value = T1(i)
        t_sum = t_sum + T1(i)
        
        Cells(3 + i, "i").value = daeSoo(i)
        daesoo_sum = daesoo_sum + daeSoo(i)
        
        Cells(3 + i, "k").value = direction(i)
        direction_sum = direction_sum + direction(i)
        
        Cells(3 + i, "m").value = gradient(i)
        gradient_sum = gradient_sum + gradient(i)
        
        Cells(4, "d").value = "5년"
    
    Next i
    
    
    Cells(4, "g").value = Format(Round(t_sum / nofWell, 4), "###0.0000")
    merge_cells ("g")
    merge_cells ("d")
    
    Cells(4, "j").value = Round(daesoo_sum / nofWell, 1)
    merge_cells ("j")
    
    Cells(4, "l").value = Round(direction_sum / nofWell, 1)
    merge_cells ("l")
    
    Cells(4, "n").value = Format(Round(gradient_sum / nofWell, 4), "###0.0000")
    merge_cells ("n")
    
    Cells(4, "o").value = "무경계조건"
    merge_cells ("o")
    
    Cells(4, "h").value = 0.03
    merge_cells ("h")
    
End Sub



Sub merge_cells(cel As String)

    Range(cel & "4:" & cel & "8").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
End Sub

















