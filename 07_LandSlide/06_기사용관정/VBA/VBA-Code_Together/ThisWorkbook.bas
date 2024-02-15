' ***************************************************************
' ThisWorkbook
'
' ***************************************************************


Private Sub Workbook_Open()
    Sheets("ss").Activate
    ISIT_FIRST = True
    
     Call clearRowA
     
End Sub

