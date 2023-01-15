VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shStepTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call findAnswer_StepTest
End Sub

Private Sub CommandButton2_Click()
    Call check_StepTest
End Sub


Private Sub CommandButton3_Click()
    Dim mm As Integer
    Dim dtLongTerm, dtStempTime As Date
    
    mm = ComboBox1.Value
    dtLongTerm = Sheets("LongTest").Range("c10").Value
    
    dtStepTime = dtLongTerm - mm / 1440
    Range("c12").Value = dtStepTime
End Sub

Private Sub CommandButton4_Click()
    Dim dtToday, ntime, nDate As Date
    
'    dtToday = Date
'    ntime = TimeSerial(10, 0, 0)
'    nDate = dtToday + ntime
'
'    Range("c12").Value = nDate
    
    UserFormTS2.Show
End Sub

Private Sub Worksheet_Activate()
    Dim arr() As Variant
    Dim i As Integer
    
    arr = Array(250, 260, 270, 300, 360, 370, 380, 390, 420, 480, 490, 500, 510, 540, 600)
        
    If (ActiveSheet.name = "Step") Then Exit Sub
    
    ComboBox1.Clear
    For i = LBound(arr) To UBound(arr)
        ComboBox1.AddItem (arr(i))
    Next i
    ComboBox1.Value = arr(UBound(arr))
End Sub


