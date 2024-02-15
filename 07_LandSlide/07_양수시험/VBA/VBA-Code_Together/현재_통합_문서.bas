
Private Sub Workbook_Open()
      
    'Sheet6.Activate
    sh01_StepSelect.name = "Step.Select"
    
    'Sheet7.Activate
    sh02_JanggiSelect.name = "Janggi.Select"
    
    'Sheet71.Activate
    sh03_RecoverSelect.name = "Recover.Select"
       
       
    With shW_LongTEST.ComboBox1
        .AddItem "60"
        .AddItem "75"
        .AddItem "90"
        .AddItem "105"
        .AddItem "120"
        .AddItem "140"
        .AddItem "160"
        .AddItem "180"
        .AddItem "240"
        .AddItem "300"
        .AddItem "360"
        .AddItem "420"
        .AddItem "480"
        .AddItem "540"
        .AddItem "600"
        .AddItem "660"
        .AddItem "720"
        .AddItem "780"
        .AddItem "840"
        .AddItem "900"
        .AddItem "960"
        .AddItem "1020"
        .AddItem "1080"
        .AddItem "1140"
        .AddItem "1200"
        .AddItem "1260"
        .AddItem "1320"
        .AddItem "1380"
        .AddItem "1440"
        .AddItem "1500"
    End With
   
    Call initDictionary
    ' Call GotoTopPosition
    
End Sub


'Private Sub GotoTopPosition()
'
'    Dim sht As Worksheet
'
'    For Each sht In Application.Worksheets
'        sht.Activate
'        Application.GoTo Reference:=Range("a1"), Scroll:=True
'    Next sht
'
'    shInput.Activate
'
'End Sub
'
