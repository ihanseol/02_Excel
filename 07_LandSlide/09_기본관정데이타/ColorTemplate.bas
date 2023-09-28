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




Sub SetRandomSheetTabColor()

    ' Define an array of RGB colors
    Dim Colors() As Variant
    
    Colors = Array(RGB(255, 0, 0), RGB(0, 255, 0), RGB(0, 0, 255), _
    RGB(255, 255, 0), RGB(255, 0, 255), RGB(0, 255, 255), RGB(128, 0, 128), _
    RGB(255, 165, 0), RGB(0, 128, 0), RGB(128, 128, 0), RGB(128, 0, 0), _
    RGB(0, 128, 128), RGB(255, 192, 203), RGB(0, 255, 127), RGB(255, 215, 0), _
    RGB(173, 255, 47), RGB(255, 69, 0), RGB(70, 130, 180), RGB(240, 230, 140), RGB(0, 0, 128))
    
    ' Generate a random number to select a color from the array
    Randomize
    Dim RandomIndex As Integer
    RandomIndex = Int((UBound(Colors) + 1) * Rnd)
    
    ' Set the sheet tab color to the randomly selected color
    ActiveSheet.Tab.Color = Colors(RandomIndex)

End Sub



Dim ColorValue(1 To 20) As Long

Public Sub InitialSetColorValue()
    ColorValue(1) = RGB(192, 0, 0)
    ColorValue(2) = RGB(255, 0, 0)
    ColorValue(3) = RGB(255, 192, 0)
    ColorValue(4) = RGB(255, 255, 0)
    ColorValue(5) = RGB(146, 208, 80)
    ColorValue(6) = RGB(0, 176, 80)
    ColorValue(7) = RGB(0, 176, 240)
    ColorValue(8) = RGB(0, 112, 192)
    ColorValue(9) = RGB(0, 32, 96)
    ColorValue(10) = RGB(112, 48, 160)
    
    ColorValue(11) = RGB(192 + 10, 10, 0)
    ColorValue(12) = RGB(255, 0 + 10, 0)
    ColorValue(13) = RGB(255, 192 + 10, 0)
    ColorValue(14) = RGB(255, 255, 10)
    ColorValue(15) = RGB(146 + 10, 208 + 10, 80 + 10)
    ColorValue(16) = RGB(0 + 10, 176 + 10, 80)
    ColorValue(17) = RGB(0 + 10, 176 + 10, 240 + 10)
    ColorValue(18) = RGB(0 + 10, 112 + 10, 192)
    ColorValue(19) = RGB(0 + 10, 32 + 10, 96)
    ColorValue(20) = RGB(112, 48 + 10, 160 + 10)
End Sub



