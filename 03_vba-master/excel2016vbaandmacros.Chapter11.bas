Attribute VB_Name = "Chapter11"
Sub loopdata()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("11").Activate
    'get count of rows
    finalrow = Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox finalrow
    'highlight USA rows green using cells and Resize
    For presentrow = 2 To finalrow
        If Cells(presentrow, 5) = "USA" Then
            Cells(presentrow, 1).Resize(1, 8).Interior.Color = RGB(0, 255, 0)
        End If
    Next presentrow
    'highlight USA rows red using Range and .End
    For presentrow = 2 To finalrow
        If Range("E" & presentrow).Value = "USA" Then
            Range("A" & presentrow, Range("A" & presentrow).End(xlToRight)).Interior.Color = RGB(255, 0, 0)
        End If
    Next presentrow
End Sub


Sub deleterows()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("11").Activate
    'If you needed to delete records, you would need to be careful to run the loop from the bottom
    'of the data set to the top, using code like this:
    'get count of rows
    finalrow = Cells(Rows.Count, 1).End(xlUp).Row
    For n = finalrow To 2 Step -1
        If Cells(n, 5) = "USA" Then
            Rows(n).Delete
        End If
    Next n
    
    'delete row based on date
    For n = finalrow To 2 Step -1
        If Cells(n, 4) >= #1/1/2016# And Cells(n, 4) <= #12/31/2016# Then
            Rows(n).Delete
        End If
    Next n
    
End Sub


Sub keeprowsordeleterows()
    finalrow = Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox finalrow
    'keep rows with song
    For n = finalrow To 1 Step -1
        If InStr(1, Cells(n, 1), "song") = 0 Then
            Rows(n).Delete
        End If
    Next n
    'delete rows with song
    For n = finalrow To 1 Step -1
        If InStr(1, Cells(n, 1), "song") > 0 Then
            Rows(n).Delete
        End If
    Next n
End Sub


Sub keeprowssongofthedaytwitter()
    'Application.Workbooks("temptweets.xlsm").Worksheets("tweets").Activate
    Application.ScreenUpdating = False
    finalrow = Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox finalrow
    'keep rows with song and keep header row
    For n = finalrow To 2 Step -1
        'If InStr(1, cells(n, 6), "#mysongoftheday") = 0 Then
        If InStr(1, Cells(n, 6), "#raymondshistorybook") = 0 Then
            Rows(n).Delete
        End If
    Next n
    Application.ScreenUpdating = True
End Sub
