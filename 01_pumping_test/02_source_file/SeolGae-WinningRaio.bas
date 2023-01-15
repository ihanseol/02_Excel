Attribute VB_Name = "WinningRaio"
Option Explicit



Sub copy_naeyeokseo()
    
    Sheets("내역서").Select
    Sheets("내역서").Copy Before:=Sheets(4)

End Sub

Sub clear_debug_window()

   Debug.Print String(65535, vbCr)
    
End Sub
    
    
    
Sub Apply_WinningRatio_methodV1()

    Dim lastRow_i, lastRow_k, lastRow_m, LASTROW(1 To 3)
    Dim i, j, ip As Long
    Dim source, target As String
    
    
    Call copy_naeyeokseo
    
    lastRow_i = ActiveSheet.Range("I" & Rows.Count).End(xlUp).row
    LASTROW(1) = lastRow_i
    
    lastRow_k = ActiveSheet.Range("K" & Rows.Count).End(xlUp).row
    LASTROW(2) = lastRow_k
    
    lastRow_m = ActiveSheet.Range("M" & Rows.Count).End(xlUp).row
    LASTROW(3) = lastRow_m
 
    ip = MainModule.getFreezingRow()
    Range("i" & ip).End(xlDown).Select
    ip = Selection.row
    
    Range("Q" & ip & ":" & "V" & WorksheetFunction.Max(lastRow_i, lastRow_k, lastRow_m)).ClearContents
    
       
    'backup original value
    For j = 0 To 5 Step 2
        source = Chr(Asc("i") + j)
        target = Chr(Asc("q") + j)
        
        For i = ip To LASTROW((j / 2) + 1)
           
               Range(target & i).Value = Range(source & i).Value
                    
        Next i
    Next j
          
    'apply ratio value into original cell
    For j = 0 To 5 Step 2
    
        source = Chr(Asc("i") + j)
        target = Chr(Asc("q") + j)
        
        For i = ip To LASTROW((j / 2) + 1)
             If Range(target & i).Value Then
                    Range(source & i).Formula = "=rounddown(" & target & i & "*$P$6,0)"
             End If
        Next i
    Next j

    
End Sub

    
    
