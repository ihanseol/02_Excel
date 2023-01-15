Attribute VB_Name = "Chapter08"
Sub arrays()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("8").Activate
    'if you want arrays to start at 1, then being subprocedure with Option Base 1
    Dim myarray(2)
    myarray(0) = "first array of three"
    myarray(1) = "second array of three"
    myarray(2) = "third array of three"
    Range("A1").Value = myarray(2)
End Sub
Sub arraysstartcount1()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("8").Activate
    Dim myarray(1 To 3)
    'dim myarray(100 to 300) is valid
    myarray(1) = "first array of three"
    myarray(2) = "second array of three"
    myarray(3) = "third array of three"
    Range("A1").Value = myarray(2)
End Sub
Sub multidimensionalarray()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("8").Activate
    Dim myarray(1 To 10, 1 To 20)
    myarray(1, 1) = 10
    myarray(1, 2) = 20
    myarray(2, 1) = 20
    myarray(10, 20) = 100
End Sub
