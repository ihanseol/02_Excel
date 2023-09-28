Sub GenerateColorTemplate1()
    Dim ColorValue(1 To 10) As Long
    Dim i As Integer
    
    ' Define the RGB values for each color
    ColorValue(1) = RGB(192, 0, 0) ' Red
    ColorValue(2) = RGB(255, 165, 0) ' Orange
    ColorValue(3) = RGB(255, 255, 0) ' Yellow
    ColorValue(4) = RGB(0, 176, 80) ' Green
    ColorValue(5) = RGB(0, 112, 192) ' Blue
    ColorValue(6) = RGB(112, 48, 160) ' Purple
    ColorValue(7) = RGB(128, 128, 128) ' Gray
    ColorValue(8)=RGB (255 ,192 ,203 )' Pink 
     colorvalue (9)=rgb (128 ,64 ,64 )' brown 
     colorvalue (10)=rgb (64 ,224 ,208)' turquoise 

    
   For i = LBound(ColorValue()) To UBound(ColorValue())
        Cells(i +1 , "A").Interior.Color =Colorvalue(i)
        Cells(i +1 ,"B")= "Color" & i
        
   Next i

End Sub




Sub GenerateColorTemplate2()
    Dim ColorValue(1 To 33) As Long
    Dim i As Integer
    
    ' Define the RGB values for each color
    ' The following are just examples, replace with actual desired RGB values.
    
    For i = LBound(ColorValue) To UBound(ColorValue)
        ColorValue(i) = RGB((i * 7) Mod 256, (i * 13) Mod 256, (i * 19) Mod 256)
    Next i
    
   For i = LBound(ColorValue()) To UBound(ColorValue())
        Cells(i +1 , "A").Interior.Color =Colorvalue(i)
        Cells(i +1 ,"B")= "Color" & i
        
   Next i

End Sub