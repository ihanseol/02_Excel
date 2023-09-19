Attribute VB_Name = "modDataChungNam"
'***********************
'Year 2023
'***********************
'data_GEUMSAN
'data_BORYUNG
'data_DAEJEON
'data_BUYEO
'data_SEOSAN
'data_CHEONAN
'data_CHEUNGJU
'***********************
'data_HONGSUNG
'data_SEJONG
'***********************

Function data_TEMP() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    data_TEMP = myArray

End Function



Function data_HONGSUNG() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    data_HONGSUNG = myArray

End Function



Function data_CHEUNGJU() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1993
    myArray(1, 2) = 7.3
    myArray(1, 3) = 75.4
    myArray(1, 4) = 33.1
    myArray(1, 5) = 54.6
    myArray(1, 6) = 127.2
    myArray(1, 7) = 118.8
    myArray(1, 8) = 254
    myArray(1, 9) = 378.1
    myArray(1, 10) = 126.6
    myArray(1, 11) = 39.6
    myArray(1, 12) = 66.9
    myArray(1, 13) = 20.2
    
    myArray(2, 1) = 1994
    myArray(2, 2) = 13.3
    myArray(2, 3) = 12.8
    myArray(2, 4) = 54.2
    myArray(2, 5) = 21.3
    myArray(2, 6) = 108.8
    myArray(2, 7) = 140.5
    myArray(2, 8) = 85.5
    myArray(2, 9) = 318.5
    myArray(2, 10) = 48.1
    myArray(2, 11) = 160.1
    myArray(2, 12) = 29.6
    myArray(2, 13) = 19.3
    
    myArray(3, 1) = 1995
    myArray(3, 2) = 21.5
    myArray(3, 3) = 14
    myArray(3, 4) = 34.4
    myArray(3, 5) = 64
    myArray(3, 6) = 70.7
    myArray(3, 7) = 30.9
    myArray(3, 8) = 204.9
    myArray(3, 9) = 835.4
    myArray(3, 10) = 17.5
    myArray(3, 11) = 22.6
    myArray(3, 12) = 20.3
    myArray(3, 13) = 3.6
    
    myArray(4, 1) = 1996
    myArray(4, 2) = 27.9
    myArray(4, 3) = 4.2
    myArray(4, 4) = 98.4
    myArray(4, 5) = 28.6
    myArray(4, 6) = 36.8
    myArray(4, 7) = 255.8
    myArray(4, 8) = 170.5
    myArray(4, 9) = 128.6
    myArray(4, 10) = 11.2
    myArray(4, 11) = 67.1
    myArray(4, 12) = 77.2
    myArray(4, 13) = 22.5
    
    myArray(5, 1) = 1997
    myArray(5, 2) = 12.9
    myArray(5, 3) = 39.1
    myArray(5, 4) = 31.6
    myArray(5, 5) = 58.5
    myArray(5, 6) = 179.1
    myArray(5, 7) = 210.3
    myArray(5, 8) = 425.5
    myArray(5, 9) = 211.1
    myArray(5, 10) = 55.5
    myArray(5, 11) = 8.4
    myArray(5, 12) = 180.3
    myArray(5, 13) = 44.3
    
    myArray(6, 1) = 1998
    myArray(6, 2) = 22
    myArray(6, 3) = 28.9
    myArray(6, 4) = 30.9
    myArray(6, 5) = 153.1
    myArray(6, 6) = 92.8
    myArray(6, 7) = 247
    myArray(6, 8) = 253
    myArray(6, 9) = 460.6
    myArray(6, 10) = 225.9
    myArray(6, 11) = 74.2
    myArray(6, 12) = 44.7
    myArray(6, 13) = 7.1
    
    myArray(7, 1) = 1999
    myArray(7, 2) = 1.6
    myArray(7, 3) = 3.6
    myArray(7, 4) = 54.1
    myArray(7, 5) = 91.4
    myArray(7, 6) = 102.4
    myArray(7, 7) = 191.1
    myArray(7, 8) = 122.4
    myArray(7, 9) = 197.4
    myArray(7, 10) = 281.3
    myArray(7, 11) = 252.4
    myArray(7, 12) = 15.4
    myArray(7, 13) = 13.4
    
    myArray(8, 1) = 2000
    myArray(8, 2) = 38.7
    myArray(8, 3) = 1.3
    myArray(8, 4) = 10.4
    myArray(8, 5) = 56.1
    myArray(8, 6) = 42.1
    myArray(8, 7) = 185.7
    myArray(8, 8) = 300
    myArray(8, 9) = 390.4
    myArray(8, 10) = 244.6
    myArray(8, 11) = 32.1
    myArray(8, 12) = 37.3
    myArray(8, 13) = 18.9
    
    myArray(9, 1) = 2001
    myArray(9, 2) = 56.9
    myArray(9, 3) = 50.3
    myArray(9, 4) = 11.3
    myArray(9, 5) = 12.7
    myArray(9, 6) = 14.3
    myArray(9, 7) = 217.5
    myArray(9, 8) = 171.5
    myArray(9, 9) = 135.5
    myArray(9, 10) = 11.8
    myArray(9, 11) = 75.9
    myArray(9, 12) = 6.9
    myArray(9, 13) = 19.5
    
    myArray(10, 1) = 2002
    myArray(10, 2) = 58.7
    myArray(10, 3) = 9
    myArray(10, 4) = 25.9
    myArray(10, 5) = 132
    myArray(10, 6) = 106.9
    myArray(10, 7) = 57.9
    myArray(10, 8) = 186.2
    myArray(10, 9) = 482.4
    myArray(10, 10) = 90.5
    myArray(10, 11) = 58
    myArray(10, 12) = 26.3
    myArray(10, 13) = 48
    
    myArray(11, 1) = 2003
    myArray(11, 2) = 16.2
    myArray(11, 3) = 45
    myArray(11, 4) = 38.9
    myArray(11, 5) = 192.7
    myArray(11, 6) = 113.5
    myArray(11, 7) = 186
    myArray(11, 8) = 467.2
    myArray(11, 9) = 293.9
    myArray(11, 10) = 150.6
    myArray(11, 11) = 32.5
    myArray(11, 12) = 33.1
    myArray(11, 13) = 12.2
    
    myArray(12, 1) = 2004
    myArray(12, 2) = 12.5
    myArray(12, 3) = 42.3
    myArray(12, 4) = 67.3
    myArray(12, 5) = 61
    myArray(12, 6) = 121.8
    myArray(12, 7) = 421.5
    myArray(12, 8) = 318.9
    myArray(12, 9) = 247.6
    myArray(12, 10) = 139
    myArray(12, 11) = 2
    myArray(12, 12) = 34
    myArray(12, 13) = 38
    
    myArray(13, 1) = 2005
    myArray(13, 2) = 4.6
    myArray(13, 3) = 13.8
    myArray(13, 4) = 36.8
    myArray(13, 5) = 66.1
    myArray(13, 6) = 50.7
    myArray(13, 7) = 170
    myArray(13, 8) = 373.1
    myArray(13, 9) = 334.7
    myArray(13, 10) = 295.5
    myArray(13, 11) = 54.6
    myArray(13, 12) = 16
    myArray(13, 13) = 11.3
    
    myArray(14, 1) = 2006
    myArray(14, 2) = 20
    myArray(14, 3) = 28.9
    myArray(14, 4) = 8.2
    myArray(14, 5) = 89.3
    myArray(14, 6) = 119.4
    myArray(14, 7) = 115.5
    myArray(14, 8) = 508
    myArray(14, 9) = 52
    myArray(14, 10) = 18.4
    myArray(14, 11) = 21.3
    myArray(14, 12) = 83.4
    myArray(14, 13) = 16.7
    
    myArray(15, 1) = 2007
    myArray(15, 2) = 11.2
    myArray(15, 3) = 33.3
    myArray(15, 4) = 103.2
    myArray(15, 5) = 35.8
    myArray(15, 6) = 145.5
    myArray(15, 7) = 81.2
    myArray(15, 8) = 273.2
    myArray(15, 9) = 385.5
    myArray(15, 10) = 391.4
    myArray(15, 11) = 43.5
    myArray(15, 12) = 8.8
    myArray(15, 13) = 21.9
    
    myArray(16, 1) = 2008
    myArray(16, 2) = 29
    myArray(16, 3) = 7.7
    myArray(16, 4) = 29.4
    myArray(16, 5) = 27
    myArray(16, 6) = 64.5
    myArray(16, 7) = 112
    myArray(16, 8) = 296.6
    myArray(16, 9) = 195.5
    myArray(16, 10) = 92.6
    myArray(16, 11) = 13.1
    myArray(16, 12) = 10.5
    myArray(16, 13) = 14.4
    
    myArray(17, 1) = 2009
    myArray(17, 2) = 17.8
    myArray(17, 3) = 13.1
    myArray(17, 4) = 54.9
    myArray(17, 5) = 30.4
    myArray(17, 6) = 109.6
    myArray(17, 7) = 77.2
    myArray(17, 8) = 345.7
    myArray(17, 9) = 187.5
    myArray(17, 10) = 49.5
    myArray(17, 11) = 49.5
    myArray(17, 12) = 43.9
    myArray(17, 13) = 40.7
    
    myArray(18, 1) = 2010
    myArray(18, 2) = 37.8
    myArray(18, 3) = 69.2
    myArray(18, 4) = 99.8
    myArray(18, 5) = 70.5
    myArray(18, 6) = 110
    myArray(18, 7) = 42.6
    myArray(18, 8) = 224.1
    myArray(18, 9) = 433.2
    myArray(18, 10) = 278.6
    myArray(18, 11) = 17.1
    myArray(18, 12) = 15.7
    myArray(18, 13) = 23.8
    
    myArray(19, 1) = 2011
    myArray(19, 2) = 4.5
    myArray(19, 3) = 43.2
    myArray(19, 4) = 23.5
    myArray(19, 5) = 111.2
    myArray(19, 6) = 116.2
    myArray(19, 7) = 360.7
    myArray(19, 8) = 531.9
    myArray(19, 9) = 290.2
    myArray(19, 10) = 182.5
    myArray(19, 11) = 34.5
    myArray(19, 12) = 92.6
    myArray(19, 13) = 14.6
    
    myArray(20, 1) = 2012
    myArray(20, 2) = 17.8
    myArray(20, 3) = 3.7
    myArray(20, 4) = 65.1
    myArray(20, 5) = 106.8
    myArray(20, 6) = 31.2
    myArray(20, 7) = 93.7
    myArray(20, 8) = 257.4
    myArray(20, 9) = 479.5
    myArray(20, 10) = 162.5
    myArray(20, 11) = 61.2
    myArray(20, 12) = 52.1
    myArray(20, 13) = 56.6
    
    myArray(21, 1) = 2013
    myArray(21, 2) = 30.5
    myArray(21, 3) = 33.2
    myArray(21, 4) = 46.8
    myArray(21, 5) = 65
    myArray(21, 6) = 97.9
    myArray(21, 7) = 229.9
    myArray(21, 8) = 253.6
    myArray(21, 9) = 183.9
    myArray(21, 10) = 162.6
    myArray(21, 11) = 25
    myArray(21, 12) = 75
    myArray(21, 13) = 37.3
    
    myArray(22, 1) = 2014
    myArray(22, 2) = 5.9
    myArray(22, 3) = 6.8
    myArray(22, 4) = 51.1
    myArray(22, 5) = 43.7
    myArray(22, 6) = 35
    myArray(22, 7) = 92.6
    myArray(22, 8) = 125.1
    myArray(22, 9) = 197.5
    myArray(22, 10) = 147.5
    myArray(22, 11) = 151.1
    myArray(22, 12) = 24.8
    myArray(22, 13) = 32.6
    
    myArray(23, 1) = 2015
    myArray(23, 2) = 16
    myArray(23, 3) = 26.5
    myArray(23, 4) = 44.1
    myArray(23, 5) = 109.1
    myArray(23, 6) = 24.4
    myArray(23, 7) = 83.3
    myArray(23, 8) = 141.4
    myArray(23, 9) = 54.3
    myArray(23, 10) = 20.1
    myArray(23, 11) = 90.5
    myArray(23, 12) = 107.5
    myArray(23, 13) = 39.7
    
    myArray(24, 1) = 2016
    myArray(24, 2) = 5.7
    myArray(24, 3) = 45.5
    myArray(24, 4) = 13.2
    myArray(24, 5) = 132.1
    myArray(24, 6) = 84.4
    myArray(24, 7) = 39.9
    myArray(24, 8) = 320
    myArray(24, 9) = 69
    myArray(24, 10) = 78.1
    myArray(24, 11) = 83.6
    myArray(24, 12) = 26.4
    myArray(24, 13) = 40.1
    
    myArray(25, 1) = 2017
    myArray(25, 2) = 12
    myArray(25, 3) = 38.7
    myArray(25, 4) = 8.9
    myArray(25, 5) = 61.7
    myArray(25, 6) = 11.9
    myArray(25, 7) = 17.5
    myArray(25, 8) = 789.1
    myArray(25, 9) = 225.2
    myArray(25, 10) = 78.3
    myArray(25, 11) = 23.1
    myArray(25, 12) = 13.7
    myArray(25, 13) = 21.1
    
    myArray(26, 1) = 2018
    myArray(26, 2) = 17.6
    myArray(26, 3) = 30.6
    myArray(26, 4) = 81.7
    myArray(26, 5) = 133
    myArray(26, 6) = 92
    myArray(26, 7) = 63.3
    myArray(26, 8) = 324.9
    myArray(26, 9) = 247.9
    myArray(26, 10) = 204
    myArray(26, 11) = 112.2
    myArray(26, 12) = 45.9
    myArray(26, 13) = 28.5
    
    myArray(27, 1) = 2019
    myArray(27, 2) = 0.1
    myArray(27, 3) = 23
    myArray(27, 4) = 20.3
    myArray(27, 5) = 60.8
    myArray(27, 6) = 20.3
    myArray(27, 7) = 82.5
    myArray(27, 8) = 204.8
    myArray(27, 9) = 80.5
    myArray(27, 10) = 155.1
    myArray(27, 11) = 84.3
    myArray(27, 12) = 104.9
    myArray(27, 13) = 20.1
    
    myArray(28, 1) = 2020
    myArray(28, 2) = 62
    myArray(28, 3) = 62.7
    myArray(28, 4) = 22.9
    myArray(28, 5) = 15.7
    myArray(28, 6) = 65.3
    myArray(28, 7) = 145.9
    myArray(28, 8) = 386.6
    myArray(28, 9) = 385.8
    myArray(28, 10) = 160.6
    myArray(28, 11) = 5.8
    myArray(28, 12) = 41
    myArray(28, 13) = 4.3
    
    myArray(29, 1) = 2021
    myArray(29, 2) = 12.7
    myArray(29, 3) = 7.5
    myArray(29, 4) = 76.6
    myArray(29, 5) = 46.4
    myArray(29, 6) = 136.4
    myArray(29, 7) = 75.4
    myArray(29, 8) = 138.1
    myArray(29, 9) = 233.1
    myArray(29, 10) = 185
    myArray(29, 11) = 29.4
    myArray(29, 12) = 57.3
    myArray(29, 13) = 3.7
    
    myArray(30, 1) = 2022
    myArray(30, 2) = 1.4
    myArray(30, 3) = 2.4
    myArray(30, 4) = 59
    myArray(30, 5) = 45.2
    myArray(30, 6) = 9.1
    myArray(30, 7) = 129.6
    myArray(30, 8) = 171.7
    myArray(30, 9) = 519.4
    myArray(30, 10) = 116
    myArray(30, 11) = 105.9
    myArray(30, 12) = 56.7
    myArray(30, 13) = 20

    data_CHEUNGJU = myArray

End Function

Function data_GEUMSAN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1993
    myArray(1, 2) = 13.3
    myArray(1, 3) = 72.5
    myArray(1, 4) = 44.5
    myArray(1, 5) = 37
    myArray(1, 6) = 132
    myArray(1, 7) = 212
    myArray(1, 8) = 302.5
    myArray(1, 9) = 281
    myArray(1, 10) = 100.5
    myArray(1, 11) = 49.5
    myArray(1, 12) = 94.5
    myArray(1, 13) = 20.5
    
    myArray(2, 1) = 1994
    myArray(2, 2) = 18.7
    myArray(2, 3) = 12.5
    myArray(2, 4) = 35
    myArray(2, 5) = 22.8
    myArray(2, 6) = 115
    myArray(2, 7) = 89
    myArray(2, 8) = 118.9
    myArray(2, 9) = 196
    myArray(2, 10) = 20
    myArray(2, 11) = 99
    myArray(2, 12) = 20.5
    myArray(2, 13) = 21.1
    
    myArray(3, 1) = 1995
    myArray(3, 2) = 23.2
    myArray(3, 3) = 17.1
    myArray(3, 4) = 46.9
    myArray(3, 5) = 65.5
    myArray(3, 6) = 35.5
    myArray(3, 7) = 54
    myArray(3, 8) = 83.5
    myArray(3, 9) = 579.5
    myArray(3, 10) = 47.5
    myArray(3, 11) = 23.5
    myArray(3, 12) = 31
    myArray(3, 13) = 4.6
    
    myArray(4, 1) = 1996
    myArray(4, 2) = 25.4
    myArray(4, 3) = 2.9
    myArray(4, 4) = 123
    myArray(4, 5) = 42.5
    myArray(4, 6) = 37.5
    myArray(4, 7) = 546
    myArray(4, 8) = 174
    myArray(4, 9) = 130
    myArray(4, 10) = 12.5
    myArray(4, 11) = 75.5
    myArray(4, 12) = 89.8
    myArray(4, 13) = 43
    
    myArray(5, 1) = 1997
    myArray(5, 2) = 21.3
    myArray(5, 3) = 48.2
    myArray(5, 4) = 34
    myArray(5, 5) = 58
    myArray(5, 6) = 170.5
    myArray(5, 7) = 238.5
    myArray(5, 8) = 444.5
    myArray(5, 9) = 246.5
    myArray(5, 10) = 89
    myArray(5, 11) = 9
    myArray(5, 12) = 160
    myArray(5, 13) = 49
    
    myArray(6, 1) = 1998
    myArray(6, 2) = 38.4
    myArray(6, 3) = 53.9
    myArray(6, 4) = 25.6
    myArray(6, 5) = 177.5
    myArray(6, 6) = 98.5
    myArray(6, 7) = 278.5
    myArray(6, 8) = 184
    myArray(6, 9) = 520
    myArray(6, 10) = 237.3
    myArray(6, 11) = 49
    myArray(6, 12) = 46.1
    myArray(6, 13) = 6.8
    
    myArray(7, 1) = 1999
    myArray(7, 2) = 5.3
    myArray(7, 3) = 22.9
    myArray(7, 4) = 73
    myArray(7, 5) = 91.5
    myArray(7, 6) = 117.5
    myArray(7, 7) = 198
    myArray(7, 8) = 114.5
    myArray(7, 9) = 167.5
    myArray(7, 10) = 289.5
    myArray(7, 11) = 125
    myArray(7, 12) = 16.4
    myArray(7, 13) = 10.3
    
    myArray(8, 1) = 2000
    myArray(8, 2) = 36.2
    myArray(8, 3) = 2.9
    myArray(8, 4) = 24.5
    myArray(8, 5) = 73.7
    myArray(8, 6) = 29
    myArray(8, 7) = 244.5
    myArray(8, 8) = 344
    myArray(8, 9) = 372
    myArray(8, 10) = 223
    myArray(8, 11) = 34.5
    myArray(8, 12) = 42
    myArray(8, 13) = 6.5
    
    myArray(9, 1) = 2001
    myArray(9, 2) = 63.2
    myArray(9, 3) = 76.5
    myArray(9, 4) = 17
    myArray(9, 5) = 22.5
    myArray(9, 6) = 22.5
    myArray(9, 7) = 212.5
    myArray(9, 8) = 203
    myArray(9, 9) = 43
    myArray(9, 10) = 87
    myArray(9, 11) = 96
    myArray(9, 12) = 12
    myArray(9, 13) = 24.1
    
    myArray(10, 1) = 2002
    myArray(10, 2) = 71.5
    myArray(10, 3) = 7.7
    myArray(10, 4) = 52
    myArray(10, 5) = 149.5
    myArray(10, 6) = 127.5
    myArray(10, 7) = 57
    myArray(10, 8) = 139.5
    myArray(10, 9) = 551
    myArray(10, 10) = 98.5
    myArray(10, 11) = 55.5
    myArray(10, 12) = 23.2
    myArray(10, 13) = 57.8
    
    myArray(11, 1) = 2003
    myArray(11, 2) = 22.4
    myArray(11, 3) = 66
    myArray(11, 4) = 44
    myArray(11, 5) = 202.5
    myArray(11, 6) = 164
    myArray(11, 7) = 138
    myArray(11, 8) = 575
    myArray(11, 9) = 280.5
    myArray(11, 10) = 192
    myArray(11, 11) = 22.5
    myArray(11, 12) = 42.5
    myArray(11, 13) = 17
    
    myArray(12, 1) = 2004
    myArray(12, 2) = 11.2
    myArray(12, 3) = 27.3
    myArray(12, 4) = 33
    myArray(12, 5) = 75.5
    myArray(12, 6) = 90.5
    myArray(12, 7) = 323.5
    myArray(12, 8) = 406
    myArray(12, 9) = 330.5
    myArray(12, 10) = 126
    myArray(12, 11) = 2.5
    myArray(12, 12) = 43
    myArray(12, 13) = 34.5
    
    myArray(13, 1) = 2005
    myArray(13, 2) = 9.4
    myArray(13, 3) = 34
    myArray(13, 4) = 51
    myArray(13, 5) = 31.5
    myArray(13, 6) = 65.5
    myArray(13, 7) = 191
    myArray(13, 8) = 411.5
    myArray(13, 9) = 387
    myArray(13, 10) = 118
    myArray(13, 11) = 23
    myArray(13, 12) = 30.5
    myArray(13, 13) = 22.6
    
    myArray(14, 1) = 2006
    myArray(14, 2) = 28
    myArray(14, 3) = 41.1
    myArray(14, 4) = 8.4
    myArray(14, 5) = 112
    myArray(14, 6) = 93.5
    myArray(14, 7) = 73
    myArray(14, 8) = 681.5
    myArray(14, 9) = 118
    myArray(14, 10) = 40.5
    myArray(14, 11) = 54
    myArray(14, 12) = 71
    myArray(14, 13) = 28.9
    
    myArray(15, 1) = 2007
    myArray(15, 2) = 13.7
    myArray(15, 3) = 57
    myArray(15, 4) = 129
    myArray(15, 5) = 27.5
    myArray(15, 6) = 104
    myArray(15, 7) = 180
    myArray(15, 8) = 252
    myArray(15, 9) = 343.5
    myArray(15, 10) = 398.5
    myArray(15, 11) = 32
    myArray(15, 12) = 13.5
    myArray(15, 13) = 35.4
    
    myArray(16, 1) = 2008
    myArray(16, 2) = 32.4
    myArray(16, 3) = 6.1
    myArray(16, 4) = 28.3
    myArray(16, 5) = 37.6
    myArray(16, 6) = 84.5
    myArray(16, 7) = 190.5
    myArray(16, 8) = 202
    myArray(16, 9) = 210
    myArray(16, 10) = 35.9
    myArray(16, 11) = 40.1
    myArray(16, 12) = 15.1
    myArray(16, 13) = 19.7
    
    myArray(17, 1) = 2009
    myArray(17, 2) = 13.2
    myArray(17, 3) = 41.5
    myArray(17, 4) = 43
    myArray(17, 5) = 36
    myArray(17, 6) = 120.3
    myArray(17, 7) = 116.3
    myArray(17, 8) = 515.5
    myArray(17, 9) = 97
    myArray(17, 10) = 54.5
    myArray(17, 11) = 24
    myArray(17, 12) = 29
    myArray(17, 13) = 38.3
    
    myArray(18, 1) = 2010
    myArray(18, 2) = 33.6
    myArray(18, 3) = 74.5
    myArray(18, 4) = 83.8
    myArray(18, 5) = 73.7
    myArray(18, 6) = 114.5
    myArray(18, 7) = 62.5
    myArray(18, 8) = 278.5
    myArray(18, 9) = 495.6
    myArray(18, 10) = 110.3
    myArray(18, 11) = 20.2
    myArray(18, 12) = 20
    myArray(18, 13) = 36.5
    
    myArray(19, 1) = 2011
    myArray(19, 2) = 2.2
    myArray(19, 3) = 63.5
    myArray(19, 4) = 21.5
    myArray(19, 5) = 132.9
    myArray(19, 6) = 130.6
    myArray(19, 7) = 237.8
    myArray(19, 8) = 571.2
    myArray(19, 9) = 403
    myArray(19, 10) = 77.8
    myArray(19, 11) = 52.2
    myArray(19, 12) = 98
    myArray(19, 13) = 7.8
    
    myArray(20, 1) = 2012
    myArray(20, 2) = 23.7
    myArray(20, 3) = 1.1
    myArray(20, 4) = 83.6
    myArray(20, 5) = 75.9
    myArray(20, 6) = 21.7
    myArray(20, 7) = 115.7
    myArray(20, 8) = 239.2
    myArray(20, 9) = 497.5
    myArray(20, 10) = 219.5
    myArray(20, 11) = 46.6
    myArray(20, 12) = 47.3
    myArray(20, 13) = 62.7
    
    myArray(21, 1) = 2013
    myArray(21, 2) = 37
    myArray(21, 3) = 43.8
    myArray(21, 4) = 64.6
    myArray(21, 5) = 86.4
    myArray(21, 6) = 79.5
    myArray(21, 7) = 117.7
    myArray(21, 8) = 216.9
    myArray(21, 9) = 159.5
    myArray(21, 10) = 80.8
    myArray(21, 11) = 32.6
    myArray(21, 12) = 53.9
    myArray(21, 13) = 24.1
    
    myArray(22, 1) = 2014
    myArray(22, 2) = 4.1
    myArray(22, 3) = 2.7
    myArray(22, 4) = 97.9
    myArray(22, 5) = 88.7
    myArray(22, 6) = 26
    myArray(22, 7) = 45.6
    myArray(22, 8) = 105.8
    myArray(22, 9) = 426.4
    myArray(22, 10) = 91.2
    myArray(22, 11) = 141.2
    myArray(22, 12) = 70.1
    myArray(22, 13) = 31.3
    
    myArray(23, 1) = 2015
    myArray(23, 2) = 37.6
    myArray(23, 3) = 23.4
    myArray(23, 4) = 46.6
    myArray(23, 5) = 93.5
    myArray(23, 6) = 29.5
    myArray(23, 7) = 143.7
    myArray(23, 8) = 162.3
    myArray(23, 9) = 83.6
    myArray(23, 10) = 18.6
    myArray(23, 11) = 93.5
    myArray(23, 12) = 109.6
    myArray(23, 13) = 35.7
    
    myArray(24, 1) = 2016
    myArray(24, 2) = 11.1
    myArray(24, 3) = 46
    myArray(24, 4) = 54.5
    myArray(24, 5) = 171.7
    myArray(24, 6) = 70.5
    myArray(24, 7) = 87.4
    myArray(24, 8) = 377.9
    myArray(24, 9) = 105.6
    myArray(24, 10) = 160.9
    myArray(24, 11) = 157.2
    myArray(24, 12) = 33.2
    myArray(24, 13) = 49.6
    
    myArray(25, 1) = 2017
    myArray(25, 2) = 13.6
    myArray(25, 3) = 54.6
    myArray(25, 4) = 29.8
    myArray(25, 5) = 76.1
    myArray(25, 6) = 31.8
    myArray(25, 7) = 48.3
    myArray(25, 8) = 305.5
    myArray(25, 9) = 222.3
    myArray(25, 10) = 105.6
    myArray(25, 11) = 35.1
    myArray(25, 12) = 15.6
    myArray(25, 13) = 29.3
    
    myArray(26, 1) = 2018
    myArray(26, 2) = 25.7
    myArray(26, 3) = 28.1
    myArray(26, 4) = 91.5
    myArray(26, 5) = 142.4
    myArray(26, 6) = 110.4
    myArray(26, 7) = 104.3
    myArray(26, 8) = 163.5
    myArray(26, 9) = 410.4
    myArray(26, 10) = 135.2
    myArray(26, 11) = 112.6
    myArray(26, 12) = 45.5
    myArray(26, 13) = 27.6
    
    myArray(27, 1) = 2019
    myArray(27, 2) = 6.4
    myArray(27, 3) = 41.5
    myArray(27, 4) = 33
    myArray(27, 5) = 93
    myArray(27, 6) = 44.2
    myArray(27, 7) = 101
    myArray(27, 8) = 141.1
    myArray(27, 9) = 105.8
    myArray(27, 10) = 236.4
    myArray(27, 11) = 99.3
    myArray(27, 12) = 47.9
    myArray(27, 13) = 33
    
    myArray(28, 1) = 2020
    myArray(28, 2) = 80.8
    myArray(28, 3) = 83.9
    myArray(28, 4) = 20.5
    myArray(28, 5) = 35.6
    myArray(28, 6) = 80.5
    myArray(28, 7) = 234
    myArray(28, 8) = 628
    myArray(28, 9) = 373.4
    myArray(28, 10) = 167.2
    myArray(28, 11) = 4.1
    myArray(28, 12) = 41.9
    myArray(28, 13) = 8.3
    
    myArray(29, 1) = 2021
    myArray(29, 2) = 23.5
    myArray(29, 3) = 19.3
    myArray(29, 4) = 88
    myArray(29, 5) = 39.3
    myArray(29, 6) = 162.7
    myArray(29, 7) = 105.6
    myArray(29, 8) = 300.8
    myArray(29, 9) = 297.2
    myArray(29, 10) = 151.9
    myArray(29, 11) = 44
    myArray(29, 12) = 50.7
    myArray(29, 13) = 7.1
    
    myArray(30, 1) = 2022
    myArray(30, 2) = 1.5
    myArray(30, 3) = 4.1
    myArray(30, 4) = 80.6
    myArray(30, 5) = 63.3
    myArray(30, 6) = 4.7
    myArray(30, 7) = 145.4
    myArray(30, 8) = 183.7
    myArray(30, 9) = 265.7
    myArray(30, 10) = 68.2
    myArray(30, 11) = 59.3
    myArray(30, 12) = 54.2
    myArray(30, 13) = 18

    data_GEUMSAN = myArray

End Function



Function data_DAEJEON() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1993
    myArray(1, 2) = 8
    myArray(1, 3) = 82.9
    myArray(1, 4) = 40.1
    myArray(1, 5) = 64.8
    myArray(1, 6) = 154.9
    myArray(1, 7) = 222.4
    myArray(1, 8) = 295.6
    myArray(1, 9) = 364.3
    myArray(1, 10) = 142.4
    myArray(1, 11) = 39.3
    myArray(1, 12) = 93.7
    myArray(1, 13) = 24.7
    
    myArray(2, 1) = 1994
    myArray(2, 2) = 17.9
    myArray(2, 3) = 16.8
    myArray(2, 4) = 46.5
    myArray(2, 5) = 38.7
    myArray(2, 6) = 138.4
    myArray(2, 7) = 115.1
    myArray(2, 8) = 105.3
    myArray(2, 9) = 145.9
    myArray(2, 10) = 37.9
    myArray(2, 11) = 145.3
    myArray(2, 12) = 24.3
    myArray(2, 13) = 25.8
    
    myArray(3, 1) = 1995
    myArray(3, 2) = 23.5
    myArray(3, 3) = 16.9
    myArray(3, 4) = 33.8
    myArray(3, 5) = 54.7
    myArray(3, 6) = 62.2
    myArray(3, 7) = 33.6
    myArray(3, 8) = 155.4
    myArray(3, 9) = 641.9
    myArray(3, 10) = 53.4
    myArray(3, 11) = 36
    myArray(3, 12) = 17.5
    myArray(3, 13) = 7.3
    
    myArray(4, 1) = 1996
    myArray(4, 2) = 32.7
    myArray(4, 3) = 4.4
    myArray(4, 4) = 138
    myArray(4, 5) = 49.8
    myArray(4, 6) = 62.9
    myArray(4, 7) = 411.4
    myArray(4, 8) = 257.7
    myArray(4, 9) = 114.4
    myArray(4, 10) = 11.4
    myArray(4, 11) = 90.8
    myArray(4, 12) = 77.1
    myArray(4, 13) = 28.6
    
    myArray(5, 1) = 1997
    myArray(5, 2) = 15.6
    myArray(5, 3) = 51.1
    myArray(5, 4) = 37.1
    myArray(5, 5) = 55.4
    myArray(5, 6) = 200.9
    myArray(5, 7) = 267.5
    myArray(5, 8) = 424.2
    myArray(5, 9) = 463.5
    myArray(5, 10) = 30.2
    myArray(5, 11) = 7.7
    myArray(5, 12) = 168.2
    myArray(5, 13) = 44.5
    
    myArray(6, 1) = 1998
    myArray(6, 2) = 33.3
    myArray(6, 3) = 36.3
    myArray(6, 4) = 31.1
    myArray(6, 5) = 154.3
    myArray(6, 6) = 119.5
    myArray(6, 7) = 297.2
    myArray(6, 8) = 256.1
    myArray(6, 9) = 781.7
    myArray(6, 10) = 254.7
    myArray(6, 11) = 71.5
    myArray(6, 12) = 31.6
    myArray(6, 13) = 2.7
    
    myArray(7, 1) = 1999
    myArray(7, 2) = 1.8
    myArray(7, 3) = 12.2
    myArray(7, 4) = 79.4
    myArray(7, 5) = 103
    myArray(7, 6) = 116.8
    myArray(7, 7) = 245.7
    myArray(7, 8) = 137.8
    myArray(7, 9) = 203
    myArray(7, 10) = 359.5
    myArray(7, 11) = 171.6
    myArray(7, 12) = 16.5
    myArray(7, 13) = 7.9
    
    myArray(8, 1) = 2000
    myArray(8, 2) = 27.5
    myArray(8, 3) = 4.1
    myArray(8, 4) = 17.8
    myArray(8, 5) = 67.8
    myArray(8, 6) = 54.3
    myArray(8, 7) = 238.3
    myArray(8, 8) = 470.1
    myArray(8, 9) = 473.6
    myArray(8, 10) = 263.2
    myArray(8, 11) = 24.6
    myArray(8, 12) = 44.6
    myArray(8, 13) = 21.6
    
    myArray(9, 1) = 2001
    myArray(9, 2) = 61.2
    myArray(9, 3) = 70
    myArray(9, 4) = 16
    myArray(9, 5) = 20.4
    myArray(9, 6) = 30.2
    myArray(9, 7) = 234.2
    myArray(9, 8) = 171
    myArray(9, 9) = 78.1
    myArray(9, 10) = 25.2
    myArray(9, 11) = 91.2
    myArray(9, 12) = 10.8
    myArray(9, 13) = 20.4
    
    myArray(10, 1) = 2002
    myArray(10, 2) = 92.1
    myArray(10, 3) = 12
    myArray(10, 4) = 33.5
    myArray(10, 5) = 155.5
    myArray(10, 6) = 130.5
    myArray(10, 7) = 55.4
    myArray(10, 8) = 149.1
    myArray(10, 9) = 538.8
    myArray(10, 10) = 77
    myArray(10, 11) = 67.8
    myArray(10, 12) = 24
    myArray(10, 13) = 43
    
    myArray(11, 1) = 2003
    myArray(11, 2) = 11.2
    myArray(11, 3) = 59.2
    myArray(11, 4) = 44.2
    myArray(11, 5) = 217.5
    myArray(11, 6) = 119.5
    myArray(11, 7) = 186.4
    myArray(11, 8) = 576.3
    myArray(11, 9) = 254.9
    myArray(11, 10) = 208.5
    myArray(11, 11) = 21.5
    myArray(11, 12) = 32.6
    myArray(11, 13) = 17.1
    
    myArray(12, 1) = 2004
    myArray(12, 2) = 10.9
    myArray(12, 3) = 30.6
    myArray(12, 4) = 83.2
    myArray(12, 5) = 73.1
    myArray(12, 6) = 109
    myArray(12, 7) = 383.5
    myArray(12, 8) = 391
    myArray(12, 9) = 198.3
    myArray(12, 10) = 133.7
    myArray(12, 11) = 5
    myArray(12, 12) = 37.1
    myArray(12, 13) = 41.1
    
    myArray(13, 1) = 2005
    myArray(13, 2) = 6
    myArray(13, 3) = 37.5
    myArray(13, 4) = 38.8
    myArray(13, 5) = 48.5
    myArray(13, 6) = 60.5
    myArray(13, 7) = 209.6
    myArray(13, 8) = 463.3
    myArray(13, 9) = 499.5
    myArray(13, 10) = 226.4
    myArray(13, 11) = 30.5
    myArray(13, 12) = 20.3
    myArray(13, 13) = 15.2
    
    myArray(14, 1) = 2006
    myArray(14, 2) = 31.2
    myArray(14, 3) = 33.1
    myArray(14, 4) = 8.1
    myArray(14, 5) = 94.2
    myArray(14, 6) = 119.7
    myArray(14, 7) = 131
    myArray(14, 8) = 531
    myArray(14, 9) = 113.6
    myArray(14, 10) = 24.1
    myArray(14, 11) = 19.3
    myArray(14, 12) = 60
    myArray(14, 13) = 29.9
    
    myArray(15, 1) = 2007
    myArray(15, 2) = 14
    myArray(15, 3) = 45
    myArray(15, 4) = 117.5
    myArray(15, 5) = 28.6
    myArray(15, 6) = 130.1
    myArray(15, 7) = 133
    myArray(15, 8) = 275.7
    myArray(15, 9) = 373
    myArray(15, 10) = 549.9
    myArray(15, 11) = 47.4
    myArray(15, 12) = 9.8
    myArray(15, 13) = 26.9
    
    myArray(16, 1) = 2008
    myArray(16, 2) = 45.3
    myArray(16, 3) = 9.1
    myArray(16, 4) = 29.1
    myArray(16, 5) = 34.4
    myArray(16, 6) = 59.2
    myArray(16, 7) = 148.3
    myArray(16, 8) = 253.4
    myArray(16, 9) = 325.2
    myArray(16, 10) = 85.5
    myArray(16, 11) = 19.6
    myArray(16, 12) = 12.1
    myArray(16, 13) = 16.4
    
    myArray(17, 1) = 2009
    myArray(17, 2) = 15.4
    myArray(17, 3) = 27.5
    myArray(17, 4) = 60.3
    myArray(17, 5) = 34.5
    myArray(17, 6) = 124.3
    myArray(17, 7) = 87.3
    myArray(17, 8) = 429.2
    myArray(17, 9) = 148.3
    myArray(17, 10) = 46
    myArray(17, 11) = 24.7
    myArray(17, 12) = 54.7
    myArray(17, 13) = 38.2
    
    myArray(18, 1) = 2010
    myArray(18, 2) = 46.4
    myArray(18, 3) = 81.5
    myArray(18, 4) = 100.1
    myArray(18, 5) = 88.5
    myArray(18, 6) = 117.6
    myArray(18, 7) = 65.6
    myArray(18, 8) = 223.1
    myArray(18, 9) = 376.4
    myArray(18, 10) = 250.5
    myArray(18, 11) = 17.9
    myArray(18, 12) = 16.4
    myArray(18, 13) = 35.7
    
    myArray(19, 1) = 2011
    myArray(19, 2) = 4
    myArray(19, 3) = 44.8
    myArray(19, 4) = 19
    myArray(19, 5) = 71
    myArray(19, 6) = 162
    myArray(19, 7) = 391.6
    myArray(19, 8) = 587.3
    myArray(19, 9) = 420.3
    myArray(19, 10) = 91.7
    myArray(19, 11) = 37
    myArray(19, 12) = 103.2
    myArray(19, 13) = 11.5
    
    myArray(20, 1) = 2012
    myArray(20, 2) = 16.4
    myArray(20, 3) = 2.5
    myArray(20, 4) = 54.6
    myArray(20, 5) = 66.2
    myArray(20, 6) = 24
    myArray(20, 7) = 57.8
    myArray(20, 8) = 277.6
    myArray(20, 9) = 463.6
    myArray(20, 10) = 242.5
    myArray(20, 11) = 81.3
    myArray(20, 12) = 58.4
    myArray(20, 13) = 64.6
    
    myArray(21, 1) = 2013
    myArray(21, 2) = 46.2
    myArray(21, 3) = 54.2
    myArray(21, 4) = 52.8
    myArray(21, 5) = 86.8
    myArray(21, 6) = 110.4
    myArray(21, 7) = 162.6
    myArray(21, 8) = 218.7
    myArray(21, 9) = 126.6
    myArray(21, 10) = 146.4
    myArray(21, 11) = 19.6
    myArray(21, 12) = 63.1
    myArray(21, 13) = 32.8
    
    myArray(22, 1) = 2014
    myArray(22, 2) = 6.5
    myArray(22, 3) = 8.5
    myArray(22, 4) = 67.2
    myArray(22, 5) = 59.4
    myArray(22, 6) = 49.7
    myArray(22, 7) = 143.7
    myArray(22, 8) = 177.2
    myArray(22, 9) = 240.9
    myArray(22, 10) = 118
    myArray(22, 11) = 169.4
    myArray(22, 12) = 40.7
    myArray(22, 13) = 36.5
    
    myArray(23, 1) = 2015
    myArray(23, 2) = 31.5
    myArray(23, 3) = 27
    myArray(23, 4) = 44.7
    myArray(23, 5) = 95.2
    myArray(23, 6) = 28.9
    myArray(23, 7) = 119.8
    myArray(23, 8) = 145.6
    myArray(23, 9) = 51.6
    myArray(23, 10) = 18.5
    myArray(23, 11) = 94.1
    myArray(23, 12) = 126.1
    myArray(23, 13) = 39.7
    
    myArray(24, 1) = 2016
    myArray(24, 2) = 11.6
    myArray(24, 3) = 45.8
    myArray(24, 4) = 40.3
    myArray(24, 5) = 154.9
    myArray(24, 6) = 85.1
    myArray(24, 7) = 62.5
    myArray(24, 8) = 367.9
    myArray(24, 9) = 57.4
    myArray(24, 10) = 196
    myArray(24, 11) = 122.6
    myArray(24, 12) = 29.5
    myArray(24, 13) = 54.8
    
    myArray(25, 1) = 2017
    myArray(25, 2) = 15
    myArray(25, 3) = 42
    myArray(25, 4) = 11.6
    myArray(25, 5) = 77.7
    myArray(25, 6) = 29.3
    myArray(25, 7) = 35.3
    myArray(25, 8) = 434.5
    myArray(25, 9) = 293.8
    myArray(25, 10) = 111.4
    myArray(25, 11) = 28.3
    myArray(25, 12) = 15.1
    myArray(25, 13) = 33.5
    
    myArray(26, 1) = 2018
    myArray(26, 2) = 23.9
    myArray(26, 3) = 40.5
    myArray(26, 4) = 108.4
    myArray(26, 5) = 155.3
    myArray(26, 6) = 95.9
    myArray(26, 7) = 115.8
    myArray(26, 8) = 226.9
    myArray(26, 9) = 408.6
    myArray(26, 10) = 149.4
    myArray(26, 11) = 133.9
    myArray(26, 12) = 49.8
    myArray(26, 13) = 33.7
    
    myArray(27, 1) = 2019
    myArray(27, 2) = 1.7
    myArray(27, 3) = 46.3
    myArray(27, 4) = 33.7
    myArray(27, 5) = 91.6
    myArray(27, 6) = 35.6
    myArray(27, 7) = 77.9
    myArray(27, 8) = 199
    myArray(27, 9) = 104.3
    myArray(27, 10) = 167
    myArray(27, 11) = 106.1
    myArray(27, 12) = 94
    myArray(27, 13) = 27
    
    myArray(28, 1) = 2020
    myArray(28, 2) = 78.5
    myArray(28, 3) = 91.2
    myArray(28, 4) = 24.4
    myArray(28, 5) = 17.8
    myArray(28, 6) = 80.4
    myArray(28, 7) = 192.5
    myArray(28, 8) = 544.9
    myArray(28, 9) = 361.6
    myArray(28, 10) = 173.6
    myArray(28, 11) = 3.2
    myArray(28, 12) = 41.8
    myArray(28, 13) = 4.1
    
    myArray(29, 1) = 2021
    myArray(29, 2) = 23.6
    myArray(29, 3) = 14.1
    myArray(29, 4) = 95.2
    myArray(29, 5) = 47.4
    myArray(29, 6) = 134.2
    myArray(29, 7) = 105.9
    myArray(29, 8) = 151.8
    myArray(29, 9) = 289.2
    myArray(29, 10) = 161.2
    myArray(29, 11) = 40.8
    myArray(29, 12) = 41.7
    myArray(29, 13) = 4.4
    
    myArray(30, 1) = 2022
    myArray(30, 2) = 1.2
    myArray(30, 3) = 1.4
    myArray(30, 4) = 74
    myArray(30, 5) = 69.7
    myArray(30, 6) = 8.1
    myArray(30, 7) = 117.6
    myArray(30, 8) = 195
    myArray(30, 9) = 496.1
    myArray(30, 10) = 90.2
    myArray(30, 11) = 89.3
    myArray(30, 12) = 45.8
    myArray(30, 13) = 14.7



    
    data_DAEJEON = myArray

End Function


Function data_BUYEO() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1993
    myArray(1, 2) = 4.8
    myArray(1, 3) = 73
    myArray(1, 4) = 25.5
    myArray(1, 5) = 34.5
    myArray(1, 6) = 116
    myArray(1, 7) = 260
    myArray(1, 8) = 273.5
    myArray(1, 9) = 284
    myArray(1, 10) = 98.5
    myArray(1, 11) = 34
    myArray(1, 12) = 89.5
    myArray(1, 13) = 21.9
    
    myArray(2, 1) = 1994
    myArray(2, 2) = 18.5
    myArray(2, 3) = 14.4
    myArray(2, 4) = 54.4
    myArray(2, 5) = 41
    myArray(2, 6) = 137
    myArray(2, 7) = 137.5
    myArray(2, 8) = 96.5
    myArray(2, 9) = 286.5
    myArray(2, 10) = 34
    myArray(2, 11) = 162
    myArray(2, 12) = 23
    myArray(2, 13) = 25
    
    myArray(3, 1) = 1995
    myArray(3, 2) = 22.6
    myArray(3, 3) = 23.5
    myArray(3, 4) = 24.4
    myArray(3, 5) = 62
    myArray(3, 6) = 59.5
    myArray(3, 7) = 34.5
    myArray(3, 8) = 171.5
    myArray(3, 9) = 839
    myArray(3, 10) = 46.5
    myArray(3, 11) = 22
    myArray(3, 12) = 15
    myArray(3, 13) = 5.7
    
    myArray(4, 1) = 1996
    myArray(4, 2) = 26.4
    myArray(4, 3) = 2.8
    myArray(4, 4) = 131
    myArray(4, 5) = 45
    myArray(4, 6) = 33
    myArray(4, 7) = 289
    myArray(4, 8) = 235
    myArray(4, 9) = 67
    myArray(4, 10) = 16
    myArray(4, 11) = 90.5
    myArray(4, 12) = 76
    myArray(4, 13) = 35
    
    myArray(5, 1) = 1997
    myArray(5, 2) = 9
    myArray(5, 3) = 54.9
    myArray(5, 4) = 44
    myArray(5, 5) = 70
    myArray(5, 6) = 229.5
    myArray(5, 7) = 236.5
    myArray(5, 8) = 404.5
    myArray(5, 9) = 263
    myArray(5, 10) = 24.5
    myArray(5, 11) = 8
    myArray(5, 12) = 219.5
    myArray(5, 13) = 39.5
    
    myArray(6, 1) = 1998
    myArray(6, 2) = 40.6
    myArray(6, 3) = 47
    myArray(6, 4) = 45
    myArray(6, 5) = 200.5
    myArray(6, 6) = 130.5
    myArray(6, 7) = 324
    myArray(6, 8) = 323
    myArray(6, 9) = 451.3
    myArray(6, 10) = 313.1
    myArray(6, 11) = 75.5
    myArray(6, 12) = 46.3
    myArray(6, 13) = 3.5
    
    myArray(7, 1) = 1999
    myArray(7, 2) = 3.5
    myArray(7, 3) = 10
    myArray(7, 4) = 75.7
    myArray(7, 5) = 92.5
    myArray(7, 6) = 127.5
    myArray(7, 7) = 203
    myArray(7, 8) = 149
    myArray(7, 9) = 119.5
    myArray(7, 10) = 426
    myArray(7, 11) = 290
    myArray(7, 12) = 15.5
    myArray(7, 13) = 17.4
    
    myArray(8, 1) = 2000
    myArray(8, 2) = 41.4
    myArray(8, 3) = 2.3
    myArray(8, 4) = 14.1
    myArray(8, 5) = 62
    myArray(8, 6) = 40
    myArray(8, 7) = 248.5
    myArray(8, 8) = 248.5
    myArray(8, 9) = 543
    myArray(8, 10) = 238.5
    myArray(8, 11) = 39
    myArray(8, 12) = 29.5
    myArray(8, 13) = 13.8
    
    myArray(9, 1) = 2001
    myArray(9, 2) = 65
    myArray(9, 3) = 69.5
    myArray(9, 4) = 9.8
    myArray(9, 5) = 25
    myArray(9, 6) = 23.5
    myArray(9, 7) = 132
    myArray(9, 8) = 216
    myArray(9, 9) = 98
    myArray(9, 10) = 10.5
    myArray(9, 11) = 76.5
    myArray(9, 12) = 10.5
    myArray(9, 13) = 16.3
    
    myArray(10, 1) = 2002
    myArray(10, 2) = 72.3
    myArray(10, 3) = 6
    myArray(10, 4) = 32.5
    myArray(10, 5) = 142.5
    myArray(10, 6) = 159
    myArray(10, 7) = 70.5
    myArray(10, 8) = 208
    myArray(10, 9) = 358.5
    myArray(10, 10) = 57.5
    myArray(10, 11) = 78.5
    myArray(10, 12) = 31.5
    myArray(10, 13) = 57.2
    
    myArray(11, 1) = 2003
    myArray(11, 2) = 24.2
    myArray(11, 3) = 59
    myArray(11, 4) = 52
    myArray(11, 5) = 208.5
    myArray(11, 6) = 144.5
    myArray(11, 7) = 228
    myArray(11, 8) = 626.5
    myArray(11, 9) = 202
    myArray(11, 10) = 167.5
    myArray(11, 11) = 24.5
    myArray(11, 12) = 29.5
    myArray(11, 13) = 13.8
    
    myArray(12, 1) = 2004
    myArray(12, 2) = 18.1
    myArray(12, 3) = 26.2
    myArray(12, 4) = 63.1
    myArray(12, 5) = 73.5
    myArray(12, 6) = 109
    myArray(12, 7) = 388
    myArray(12, 8) = 296
    myArray(12, 9) = 249
    myArray(12, 10) = 176.5
    myArray(12, 11) = 1
    myArray(12, 12) = 50.5
    myArray(12, 13) = 43
    
    myArray(13, 1) = 2005
    myArray(13, 2) = 6
    myArray(13, 3) = 39
    myArray(13, 4) = 26.5
    myArray(13, 5) = 75
    myArray(13, 6) = 65.5
    myArray(13, 7) = 186
    myArray(13, 8) = 448.5
    myArray(13, 9) = 381.5
    myArray(13, 10) = 225.5
    myArray(13, 11) = 30.5
    myArray(13, 12) = 21
    myArray(13, 13) = 22
    
    myArray(14, 1) = 2006
    myArray(14, 2) = 30.2
    myArray(14, 3) = 29.5
    myArray(14, 4) = 7.8
    myArray(14, 5) = 99
    myArray(14, 6) = 81.5
    myArray(14, 7) = 111
    myArray(14, 8) = 503
    myArray(14, 9) = 83.5
    myArray(14, 10) = 37.5
    myArray(14, 11) = 15
    myArray(14, 12) = 51
    myArray(14, 13) = 27.5
    
    myArray(15, 1) = 2007
    myArray(15, 2) = 21.8
    myArray(15, 3) = 47.8
    myArray(15, 4) = 159
    myArray(15, 5) = 28
    myArray(15, 6) = 104
    myArray(15, 7) = 101
    myArray(15, 8) = 286
    myArray(15, 9) = 319.5
    myArray(15, 10) = 502.5
    myArray(15, 11) = 37
    myArray(15, 12) = 13
    myArray(15, 13) = 31.7
    
    myArray(16, 1) = 2008
    myArray(16, 2) = 39.6
    myArray(16, 3) = 11.2
    myArray(16, 4) = 42.2
    myArray(16, 5) = 38.8
    myArray(16, 6) = 51.6
    myArray(16, 7) = 260
    myArray(16, 8) = 194.3
    myArray(16, 9) = 154
    myArray(16, 10) = 48.8
    myArray(16, 11) = 24.1
    myArray(16, 12) = 14.1
    myArray(16, 13) = 23.4
    
    myArray(17, 1) = 2009
    myArray(17, 2) = 10.6
    myArray(17, 3) = 23.6
    myArray(17, 4) = 63.9
    myArray(17, 5) = 51
    myArray(17, 6) = 135.5
    myArray(17, 7) = 113.2
    myArray(17, 8) = 408
    myArray(17, 9) = 140.2
    myArray(17, 10) = 30.5
    myArray(17, 11) = 23.7
    myArray(17, 12) = 54.5
    myArray(17, 13) = 34.9
    
    myArray(18, 1) = 2010
    myArray(18, 2) = 37.1
    myArray(18, 3) = 89.5
    myArray(18, 4) = 94.9
    myArray(18, 5) = 69.6
    myArray(18, 6) = 140.7
    myArray(18, 7) = 36.1
    myArray(18, 8) = 262.1
    myArray(18, 9) = 431.1
    myArray(18, 10) = 149.8
    myArray(18, 11) = 17.8
    myArray(18, 12) = 18.6
    myArray(18, 13) = 31
    
    myArray(19, 1) = 2011
    myArray(19, 2) = 3.7
    myArray(19, 3) = 60.7
    myArray(19, 4) = 16
    myArray(19, 5) = 70
    myArray(19, 6) = 111.2
    myArray(19, 7) = 316
    myArray(19, 8) = 599.6
    myArray(19, 9) = 618.1
    myArray(19, 10) = 104.2
    myArray(19, 11) = 26.6
    myArray(19, 12) = 81.6
    myArray(19, 13) = 7
    
    myArray(20, 1) = 2012
    myArray(20, 2) = 16
    myArray(20, 3) = 3.2
    myArray(20, 4) = 60.2
    myArray(20, 5) = 109.3
    myArray(20, 6) = 19.5
    myArray(20, 7) = 71.3
    myArray(20, 8) = 302.9
    myArray(20, 9) = 573.3
    myArray(20, 10) = 186.2
    myArray(20, 11) = 83
    myArray(20, 12) = 60.7
    myArray(20, 13) = 60.2
    
    myArray(21, 1) = 2013
    myArray(21, 2) = 45.4
    myArray(21, 3) = 58.7
    myArray(21, 4) = 50.3
    myArray(21, 5) = 93.7
    myArray(21, 6) = 159
    myArray(21, 7) = 151.7
    myArray(21, 8) = 240.4
    myArray(21, 9) = 119.5
    myArray(21, 10) = 184.8
    myArray(21, 11) = 17.5
    myArray(21, 12) = 79.4
    myArray(21, 13) = 35.9
    
    myArray(22, 1) = 2014
    myArray(22, 2) = 2.2
    myArray(22, 3) = 15.3
    myArray(22, 4) = 69.3
    myArray(22, 5) = 94.1
    myArray(22, 6) = 61.5
    myArray(22, 7) = 77.8
    myArray(22, 8) = 174.7
    myArray(22, 9) = 225.1
    myArray(22, 10) = 157.5
    myArray(22, 11) = 170.5
    myArray(22, 12) = 42.4
    myArray(22, 13) = 51.7
    
    myArray(23, 1) = 2015
    myArray(23, 2) = 35.4
    myArray(23, 3) = 35.6
    myArray(23, 4) = 42.4
    myArray(23, 5) = 99.5
    myArray(23, 6) = 53.5
    myArray(23, 7) = 92.7
    myArray(23, 8) = 119.9
    myArray(23, 9) = 56.9
    myArray(23, 10) = 22
    myArray(23, 11) = 104
    myArray(23, 12) = 130
    myArray(23, 13) = 56.9
    
    myArray(24, 1) = 2016
    myArray(24, 2) = 6.6
    myArray(24, 3) = 59.6
    myArray(24, 4) = 19
    myArray(24, 5) = 164.6
    myArray(24, 6) = 121.6
    myArray(24, 7) = 49.4
    myArray(24, 8) = 341.1
    myArray(24, 9) = 33.4
    myArray(24, 10) = 133.7
    myArray(24, 11) = 120.1
    myArray(24, 12) = 17.1
    myArray(24, 13) = 63.1
    
    myArray(25, 1) = 2017
    myArray(25, 2) = 16
    myArray(25, 3) = 28.5
    myArray(25, 4) = 8.8
    myArray(25, 5) = 78.4
    myArray(25, 6) = 35.8
    myArray(25, 7) = 51.4
    myArray(25, 8) = 326.7
    myArray(25, 9) = 358.5
    myArray(25, 10) = 97.1
    myArray(25, 11) = 51.9
    myArray(25, 12) = 22.8
    myArray(25, 13) = 36.1
    
    myArray(26, 1) = 2018
    myArray(26, 2) = 25
    myArray(26, 3) = 43.1
    myArray(26, 4) = 99.3
    myArray(26, 5) = 156.5
    myArray(26, 6) = 116.1
    myArray(26, 7) = 107.1
    myArray(26, 8) = 278.8
    myArray(26, 9) = 277
    myArray(26, 10) = 98.3
    myArray(26, 11) = 159.2
    myArray(26, 12) = 66
    myArray(26, 13) = 31.5
    
    myArray(27, 1) = 2019
    myArray(27, 2) = 0.5
    myArray(27, 3) = 37.6
    myArray(27, 4) = 35
    myArray(27, 5) = 73.7
    myArray(27, 6) = 44.3
    myArray(27, 7) = 59.9
    myArray(27, 8) = 216.7
    myArray(27, 9) = 102.1
    myArray(27, 10) = 191.9
    myArray(27, 11) = 85.6
    myArray(27, 12) = 113.5
    myArray(27, 13) = 31.2
    
    myArray(28, 1) = 2020
    myArray(28, 2) = 79.6
    myArray(28, 3) = 92.4
    myArray(28, 4) = 19.3
    myArray(28, 5) = 17.7
    myArray(28, 6) = 108.5
    myArray(28, 7) = 188.4
    myArray(28, 8) = 492.6
    myArray(28, 9) = 367.8
    myArray(28, 10) = 208.9
    myArray(28, 11) = 4.4
    myArray(28, 12) = 41.8
    myArray(28, 13) = 3.4
    
    myArray(29, 1) = 2021
    myArray(29, 2) = 32.1
    myArray(29, 3) = 18.1
    myArray(29, 4) = 95.7
    myArray(29, 5) = 42.3
    myArray(29, 6) = 136.9
    myArray(29, 7) = 76.9
    myArray(29, 8) = 187.7
    myArray(29, 9) = 227.6
    myArray(29, 10) = 187.1
    myArray(29, 11) = 36.9
    myArray(29, 12) = 73.4
    myArray(29, 13) = 8.7
    
    myArray(30, 1) = 2022
    myArray(30, 2) = 3.5
    myArray(30, 3) = 2.5
    myArray(30, 4) = 76.1
    myArray(30, 5) = 62.6
    myArray(30, 6) = 4
    myArray(30, 7) = 123.4
    myArray(30, 8) = 168.5
    myArray(30, 9) = 615.6
    myArray(30, 10) = 87
    myArray(30, 11) = 103.7
    myArray(30, 12) = 36.4
    myArray(30, 13) = 17.8

    data_BUYEO = myArray

End Function


Function data_SEOSAN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1993
    myArray(1, 2) = 9.2
    myArray(1, 3) = 68.2
    myArray(1, 4) = 27.4
    myArray(1, 5) = 45.2
    myArray(1, 6) = 50.7
    myArray(1, 7) = 151.8
    myArray(1, 8) = 393.1
    myArray(1, 9) = 95
    myArray(1, 10) = 81.7
    myArray(1, 11) = 31.5
    myArray(1, 12) = 105.1
    myArray(1, 13) = 34.7
    
    myArray(2, 1) = 1994
    myArray(2, 2) = 14.6
    myArray(2, 3) = 5.9
    myArray(2, 4) = 65.6
    myArray(2, 5) = 32.4
    myArray(2, 6) = 156
    myArray(2, 7) = 167.8
    myArray(2, 8) = 107.1
    myArray(2, 9) = 309.7
    myArray(2, 10) = 99.2
    myArray(2, 11) = 216.3
    myArray(2, 12) = 23.3
    myArray(2, 13) = 36.6
    
    myArray(3, 1) = 1995
    myArray(3, 2) = 22.7
    myArray(3, 3) = 7.2
    myArray(3, 4) = 37.3
    myArray(3, 5) = 48.2
    myArray(3, 6) = 67.1
    myArray(3, 7) = 24.5
    myArray(3, 8) = 144.1
    myArray(3, 9) = 992.7
    myArray(3, 10) = 20.2
    myArray(3, 11) = 19.3
    myArray(3, 12) = 49.9
    myArray(3, 13) = 15.1
    
    myArray(4, 1) = 1996
    myArray(4, 2) = 29.1
    myArray(4, 3) = 5.7
    myArray(4, 4) = 115.1
    myArray(4, 5) = 48.1
    myArray(4, 6) = 20
    myArray(4, 7) = 179.2
    myArray(4, 8) = 152.8
    myArray(4, 9) = 74.1
    myArray(4, 10) = 6.4
    myArray(4, 11) = 92.2
    myArray(4, 12) = 72.1
    myArray(4, 13) = 35.3
    
    myArray(5, 1) = 1997
    myArray(5, 2) = 20.5
    myArray(5, 3) = 32.5
    myArray(5, 4) = 29.6
    myArray(5, 5) = 69.5
    myArray(5, 6) = 232.8
    myArray(5, 7) = 204.4
    myArray(5, 8) = 298.7
    myArray(5, 9) = 87.2
    myArray(5, 10) = 16.1
    myArray(5, 11) = 8.7
    myArray(5, 12) = 116.7
    myArray(5, 13) = 40.2
    
    myArray(6, 1) = 1998
    myArray(6, 2) = 40.1
    myArray(6, 3) = 54.2
    myArray(6, 4) = 35
    myArray(6, 5) = 160.6
    myArray(6, 6) = 95.5
    myArray(6, 7) = 281.7
    myArray(6, 8) = 295.6
    myArray(6, 9) = 491.8
    myArray(6, 10) = 168
    myArray(6, 11) = 24.3
    myArray(6, 12) = 55.6
    myArray(6, 13) = 9.2
    
    myArray(7, 1) = 1999
    myArray(7, 2) = 8
    myArray(7, 3) = 7.8
    myArray(7, 4) = 59.9
    myArray(7, 5) = 90.1
    myArray(7, 6) = 178.8
    myArray(7, 7) = 105.1
    myArray(7, 8) = 175.6
    myArray(7, 9) = 497.4
    myArray(7, 10) = 532.6
    myArray(7, 11) = 111.3
    myArray(7, 12) = 36.6
    myArray(7, 13) = 23.4
    
    myArray(8, 1) = 2000
    myArray(8, 2) = 63
    myArray(8, 3) = 2.9
    myArray(8, 4) = 3.7
    myArray(8, 5) = 38.1
    myArray(8, 6) = 62.1
    myArray(8, 7) = 204.4
    myArray(8, 8) = 60.8
    myArray(8, 9) = 608.1
    myArray(8, 10) = 298.1
    myArray(8, 11) = 34.4
    myArray(8, 12) = 24.8
    myArray(8, 13) = 24.4
    
    myArray(9, 1) = 2001
    myArray(9, 2) = 66.9
    myArray(9, 3) = 40.4
    myArray(9, 4) = 12.7
    myArray(9, 5) = 18.7
    myArray(9, 6) = 17.8
    myArray(9, 7) = 200.2
    myArray(9, 8) = 402
    myArray(9, 9) = 136.6
    myArray(9, 10) = 15
    myArray(9, 11) = 47.5
    myArray(9, 12) = 8.2
    myArray(9, 13) = 20.8
    
    myArray(10, 1) = 2002
    myArray(10, 2) = 22.5
    myArray(10, 3) = 7
    myArray(10, 4) = 29.3
    myArray(10, 5) = 179.5
    myArray(10, 6) = 177.3
    myArray(10, 7) = 60.8
    myArray(10, 8) = 296.1
    myArray(10, 9) = 428.2
    myArray(10, 10) = 57.5
    myArray(10, 11) = 78.3
    myArray(10, 12) = 36.3
    myArray(10, 13) = 14.8
    
    myArray(11, 1) = 2003
    myArray(11, 2) = 20.9
    myArray(11, 3) = 39
    myArray(11, 4) = 22.5
    myArray(11, 5) = 180
    myArray(11, 6) = 105.5
    myArray(11, 7) = 221.8
    myArray(11, 8) = 290.2
    myArray(11, 9) = 257.9
    myArray(11, 10) = 201.9
    myArray(11, 11) = 23
    myArray(11, 12) = 53.6
    myArray(11, 13) = 17.1
    
    myArray(12, 1) = 2004
    myArray(12, 2) = 27.3
    myArray(12, 3) = 26.3
    myArray(12, 4) = 15.7
    myArray(12, 5) = 80.2
    myArray(12, 6) = 140.3
    myArray(12, 7) = 211.1
    myArray(12, 8) = 321.9
    myArray(12, 9) = 131.2
    myArray(12, 10) = 282.6
    myArray(12, 11) = 1.8
    myArray(12, 12) = 70.5
    myArray(12, 13) = 32
    
    myArray(13, 1) = 2005
    myArray(13, 2) = 10.4
    myArray(13, 3) = 34
    myArray(13, 4) = 36.1
    myArray(13, 5) = 77.2
    myArray(13, 6) = 56.1
    myArray(13, 7) = 147
    myArray(13, 8) = 386.1
    myArray(13, 9) = 270.5
    myArray(13, 10) = 228.7
    myArray(13, 11) = 30.9
    myArray(13, 12) = 19.6
    myArray(13, 13) = 37.6
    
    myArray(14, 1) = 2006
    myArray(14, 2) = 29.7
    myArray(14, 3) = 18.9
    myArray(14, 4) = 5
    myArray(14, 5) = 77.3
    myArray(14, 6) = 133.5
    myArray(14, 7) = 226.8
    myArray(14, 8) = 494.5
    myArray(14, 9) = 58.2
    myArray(14, 10) = 10.1
    myArray(14, 11) = 10.5
    myArray(14, 12) = 55
    myArray(14, 13) = 19.7
    
    myArray(15, 1) = 2007
    myArray(15, 2) = 13
    myArray(15, 3) = 25.5
    myArray(15, 4) = 127.2
    myArray(15, 5) = 28.1
    myArray(15, 6) = 108.5
    myArray(15, 7) = 123.5
    myArray(15, 8) = 257
    myArray(15, 9) = 414.6
    myArray(15, 10) = 305.8
    myArray(15, 11) = 30.7
    myArray(15, 12) = 14.4
    myArray(15, 13) = 22.8
    
    myArray(16, 1) = 2008
    myArray(16, 2) = 15
    myArray(16, 3) = 7
    myArray(16, 4) = 26
    myArray(16, 5) = 46.1
    myArray(16, 6) = 88.5
    myArray(16, 7) = 118.1
    myArray(16, 8) = 335.5
    myArray(16, 9) = 114.2
    myArray(16, 10) = 62.7
    myArray(16, 11) = 34
    myArray(16, 12) = 34.6
    myArray(16, 13) = 27.9
    
    myArray(17, 1) = 2009
    myArray(17, 2) = 15.2
    myArray(17, 3) = 26.5
    myArray(17, 4) = 67
    myArray(17, 5) = 43
    myArray(17, 6) = 117.9
    myArray(17, 7) = 74.9
    myArray(17, 8) = 364.9
    myArray(17, 9) = 196.3
    myArray(17, 10) = 16
    myArray(17, 11) = 49.2
    myArray(17, 12) = 59.1
    myArray(17, 13) = 44.3
    
    myArray(18, 1) = 2010
    myArray(18, 2) = 55.5
    myArray(18, 3) = 58.4
    myArray(18, 4) = 79.2
    myArray(18, 5) = 52.2
    myArray(18, 6) = 168
    myArray(18, 7) = 94.9
    myArray(18, 8) = 447.1
    myArray(18, 9) = 707
    myArray(18, 10) = 402
    myArray(18, 11) = 29.1
    myArray(18, 12) = 12
    myArray(18, 13) = 36.4
    
    myArray(19, 1) = 2011
    myArray(19, 2) = 8.8
    myArray(19, 3) = 55.8
    myArray(19, 4) = 34.5
    myArray(19, 5) = 96.2
    myArray(19, 6) = 107.9
    myArray(19, 7) = 462.6
    myArray(19, 8) = 656.5
    myArray(19, 9) = 151.2
    myArray(19, 10) = 50.3
    myArray(19, 11) = 18.1
    myArray(19, 12) = 48.9
    myArray(19, 13) = 13.6
    
    myArray(20, 1) = 2012
    myArray(20, 2) = 15.1
    myArray(20, 3) = 2.4
    myArray(20, 4) = 41.6
    myArray(20, 5) = 113.5
    myArray(20, 6) = 14.5
    myArray(20, 7) = 91.1
    myArray(20, 8) = 266.8
    myArray(20, 9) = 647.9
    myArray(20, 10) = 201.5
    myArray(20, 11) = 100.7
    myArray(20, 12) = 82.1
    myArray(20, 13) = 65.4
    
    myArray(21, 1) = 2013
    myArray(21, 2) = 36.8
    myArray(21, 3) = 64.8
    myArray(21, 4) = 60.8
    myArray(21, 5) = 61.8
    myArray(21, 6) = 114.9
    myArray(21, 7) = 94.4
    myArray(21, 8) = 213.8
    myArray(21, 9) = 120.6
    myArray(21, 10) = 147.4
    myArray(21, 11) = 5.7
    myArray(21, 12) = 64.9
    myArray(21, 13) = 32.8
    
    myArray(22, 1) = 2014
    myArray(22, 2) = 7
    myArray(22, 3) = 17
    myArray(22, 4) = 31.2
    myArray(22, 5) = 85.6
    myArray(22, 6) = 52.7
    myArray(22, 7) = 69.3
    myArray(22, 8) = 151.7
    myArray(22, 9) = 242.3
    myArray(22, 10) = 106.7
    myArray(22, 11) = 117.2
    myArray(22, 12) = 37.8
    myArray(22, 13) = 81.6
    
    myArray(23, 1) = 2015
    myArray(23, 2) = 20.7
    myArray(23, 3) = 23.1
    myArray(23, 4) = 20.6
    myArray(23, 5) = 116.8
    myArray(23, 6) = 40.6
    myArray(23, 7) = 64.1
    myArray(23, 8) = 158.5
    myArray(23, 9) = 63.1
    myArray(23, 10) = 15.1
    myArray(23, 11) = 73.1
    myArray(23, 12) = 156.6
    myArray(23, 13) = 63.6
    
    myArray(24, 1) = 2016
    myArray(24, 2) = 21.9
    myArray(24, 3) = 61.7
    myArray(24, 4) = 24.3
    myArray(24, 5) = 87
    myArray(24, 6) = 153.7
    myArray(24, 7) = 36.8
    myArray(24, 8) = 295.6
    myArray(24, 9) = 34
    myArray(24, 10) = 53.1
    myArray(24, 11) = 73.8
    myArray(24, 12) = 17.5
    myArray(24, 13) = 62.7
    
    myArray(25, 1) = 2017
    myArray(25, 2) = 21.3
    myArray(25, 3) = 31.4
    myArray(25, 4) = 4.8
    myArray(25, 5) = 38.9
    myArray(25, 6) = 27.9
    myArray(25, 7) = 23.3
    myArray(25, 8) = 327.8
    myArray(25, 9) = 231.3
    myArray(25, 10) = 37.6
    myArray(25, 11) = 25.5
    myArray(25, 12) = 24.7
    myArray(25, 13) = 35.9
    
    myArray(26, 1) = 2018
    myArray(26, 2) = 21
    myArray(26, 3) = 40.5
    myArray(26, 4) = 76.6
    myArray(26, 5) = 132.8
    myArray(26, 6) = 147.7
    myArray(26, 7) = 162.3
    myArray(26, 8) = 152.9
    myArray(26, 9) = 156.8
    myArray(26, 10) = 82.7
    myArray(26, 11) = 153.2
    myArray(26, 12) = 73.9
    myArray(26, 13) = 26.8
    
    myArray(27, 1) = 2019
    myArray(27, 2) = 1.1
    myArray(27, 3) = 30.2
    myArray(27, 4) = 43.7
    myArray(27, 5) = 43.9
    myArray(27, 6) = 20.1
    myArray(27, 7) = 56.3
    myArray(27, 8) = 174.5
    myArray(27, 9) = 121.1
    myArray(27, 10) = 181.1
    myArray(27, 11) = 81
    myArray(27, 12) = 124.6
    myArray(27, 13) = 37.4
    
    myArray(28, 1) = 2020
    myArray(28, 2) = 46
    myArray(28, 3) = 72.3
    myArray(28, 4) = 23
    myArray(28, 5) = 20.7
    myArray(28, 6) = 101.3
    myArray(28, 7) = 144
    myArray(28, 8) = 329.4
    myArray(28, 9) = 400
    myArray(28, 10) = 257.7
    myArray(28, 11) = 12.6
    myArray(28, 12) = 72
    myArray(28, 13) = 9.7
    
    myArray(29, 1) = 2021
    myArray(29, 2) = 25.3
    myArray(29, 3) = 9.6
    myArray(29, 4) = 112.8
    myArray(29, 5) = 110.6
    myArray(29, 6) = 132.3
    myArray(29, 7) = 70.9
    myArray(29, 8) = 121.6
    myArray(29, 9) = 217.8
    myArray(29, 10) = 206
    myArray(29, 11) = 55.9
    myArray(29, 12) = 126.2
    myArray(29, 13) = 18.3
    
    myArray(30, 1) = 2022
    myArray(30, 2) = 8.6
    myArray(30, 3) = 4.7
    myArray(30, 4) = 72.1
    myArray(30, 5) = 52.2
    myArray(30, 6) = 2.9
    myArray(30, 7) = 352.4
    myArray(30, 8) = 178.4
    myArray(30, 9) = 468.7
    myArray(30, 10) = 165.9
    myArray(30, 11) = 160
    myArray(30, 12) = 72.9
    myArray(30, 13) = 31.9
    
    
        
        
        data_SEOSAN = myArray
    
    End Function
    
    
    
    Function data_CHEONAN() As Variant
    
        Dim myArray() As Variant
        ReDim myArray(1 To 30, 1 To 13)
        
        myArray(1, 1) = 1993
    myArray(1, 2) = 3.2
    myArray(1, 3) = 68.6
    myArray(1, 4) = 29.8
    myArray(1, 5) = 31
    myArray(1, 6) = 59
    myArray(1, 7) = 139
    myArray(1, 8) = 323.5
    myArray(1, 9) = 163.5
    myArray(1, 10) = 122.5
    myArray(1, 11) = 48
    myArray(1, 12) = 66.9
    myArray(1, 13) = 25.7
    
    myArray(2, 1) = 1994
    myArray(2, 2) = 9.1
    myArray(2, 3) = 10.1
    myArray(2, 4) = 39.5
    myArray(2, 5) = 13.5
    myArray(2, 6) = 106.5
    myArray(2, 7) = 160.5
    myArray(2, 8) = 98
    myArray(2, 9) = 418
    myArray(2, 10) = 52
    myArray(2, 11) = 220
    myArray(2, 12) = 22
    myArray(2, 13) = 21
    
    myArray(3, 1) = 1995
    myArray(3, 2) = 19
    myArray(3, 3) = 8.2
    myArray(3, 4) = 25.3
    myArray(3, 5) = 47
    myArray(3, 6) = 48
    myArray(3, 7) = 14.5
    myArray(3, 8) = 239.9
    myArray(3, 9) = 1082.5
    myArray(3, 10) = 29
    myArray(3, 11) = 23.5
    myArray(3, 12) = 40.2
    myArray(3, 13) = 8.9
    
    myArray(4, 1) = 1996
    myArray(4, 2) = 29.5
    myArray(4, 3) = 10.2
    myArray(4, 4) = 115
    myArray(4, 5) = 54.5
    myArray(4, 6) = 19
    myArray(4, 7) = 237
    myArray(4, 8) = 177.5
    myArray(4, 9) = 116.5
    myArray(4, 10) = 8
    myArray(4, 11) = 102.5
    myArray(4, 12) = 71.6
    myArray(4, 13) = 26.2
    
    myArray(5, 1) = 1997
    myArray(5, 2) = 10.7
    myArray(5, 3) = 44.1
    myArray(5, 4) = 30
    myArray(5, 5) = 66.5
    myArray(5, 6) = 211
    myArray(5, 7) = 191.5
    myArray(5, 8) = 305
    myArray(5, 9) = 175.5
    myArray(5, 10) = 14.5
    myArray(5, 11) = 23
    myArray(5, 12) = 153.5
    myArray(5, 13) = 43.5
    
    myArray(6, 1) = 1998
    myArray(6, 2) = 20.4
    myArray(6, 3) = 27.9
    myArray(6, 4) = 29.5
    myArray(6, 5) = 120.5
    myArray(6, 6) = 85
    myArray(6, 7) = 219.5
    myArray(6, 8) = 277
    myArray(6, 9) = 408.1
    myArray(6, 10) = 283
    myArray(6, 11) = 51.5
    myArray(6, 12) = 52.8
    myArray(6, 13) = 8.5
    
    myArray(7, 1) = 1999
    myArray(7, 2) = 2.7
    myArray(7, 3) = 2.8
    myArray(7, 4) = 46.5
    myArray(7, 5) = 88.5
    myArray(7, 6) = 121.5
    myArray(7, 7) = 163.7
    myArray(7, 8) = 138.5
    myArray(7, 9) = 313.5
    myArray(7, 10) = 324.5
    myArray(7, 11) = 134.5
    myArray(7, 12) = 16.5
    myArray(7, 13) = 11.9
    
    myArray(8, 1) = 2000
    myArray(8, 2) = 52.3
    myArray(8, 3) = 2.7
    myArray(8, 4) = 7.1
    myArray(8, 5) = 36
    myArray(8, 6) = 36
    myArray(8, 7) = 181
    myArray(8, 8) = 83
    myArray(8, 9) = 636
    myArray(8, 10) = 287.5
    myArray(8, 11) = 32
    myArray(8, 12) = 32
    myArray(8, 13) = 22.5
    
    myArray(9, 1) = 2001
    myArray(9, 2) = 43.5
    myArray(9, 3) = 44
    myArray(9, 4) = 16.5
    myArray(9, 5) = 19
    myArray(9, 6) = 15
    myArray(9, 7) = 227.5
    myArray(9, 8) = 178
    myArray(9, 9) = 194.5
    myArray(9, 10) = 12
    myArray(9, 11) = 63.5
    myArray(9, 12) = 6.3
    myArray(9, 13) = 18.4
    
    myArray(10, 1) = 2002
    myArray(10, 2) = 45.3
    myArray(10, 3) = 6
    myArray(10, 4) = 25.5
    myArray(10, 5) = 128
    myArray(10, 6) = 104
    myArray(10, 7) = 54
    myArray(10, 8) = 229.5
    myArray(10, 9) = 481.5
    myArray(10, 10) = 57
    myArray(10, 11) = 91.5
    myArray(10, 12) = 42.1
    myArray(10, 13) = 48.1
    
    myArray(11, 1) = 2003
    myArray(11, 2) = 18.6
    myArray(11, 3) = 44
    myArray(11, 4) = 38.1
    myArray(11, 5) = 172.3
    myArray(11, 6) = 106
    myArray(11, 7) = 178.6
    myArray(11, 8) = 381.2
    myArray(11, 9) = 334.6
    myArray(11, 10) = 264.2
    myArray(11, 11) = 27
    myArray(11, 12) = 46.7
    myArray(11, 13) = 17
    
    myArray(12, 1) = 2004
    myArray(12, 2) = 16.4
    myArray(12, 3) = 21.3
    myArray(12, 4) = 21.5
    myArray(12, 5) = 67.5
    myArray(12, 6) = 127.6
    myArray(12, 7) = 235
    myArray(12, 8) = 365.2
    myArray(12, 9) = 229.3
    myArray(12, 10) = 189
    myArray(12, 11) = 4.5
    myArray(12, 12) = 53
    myArray(12, 13) = 33
    
    myArray(13, 1) = 2005
    myArray(13, 2) = 3
    myArray(13, 3) = 29.8
    myArray(13, 4) = 37
    myArray(13, 5) = 53.7
    myArray(13, 6) = 48
    myArray(13, 7) = 183
    myArray(13, 8) = 313.8
    myArray(13, 9) = 202
    myArray(13, 10) = 377.5
    myArray(13, 11) = 26.7
    myArray(13, 12) = 23.5
    myArray(13, 13) = 11.3
    
    myArray(14, 1) = 2006
    myArray(14, 2) = 25.2
    myArray(14, 3) = 18.5
    myArray(14, 4) = 6.1
    myArray(14, 5) = 78.6
    myArray(14, 6) = 79
    myArray(14, 7) = 120
    myArray(14, 8) = 535.1
    myArray(14, 9) = 63.5
    myArray(14, 10) = 22.2
    myArray(14, 11) = 21.6
    myArray(14, 12) = 56.3
    myArray(14, 13) = 17.2
    
    myArray(15, 1) = 2007
    myArray(15, 2) = 9.4
    myArray(15, 3) = 34.1
    myArray(15, 4) = 108.3
    myArray(15, 5) = 35.3
    myArray(15, 6) = 126.2
    myArray(15, 7) = 106.7
    myArray(15, 8) = 215.6
    myArray(15, 9) = 470.6
    myArray(15, 10) = 353.3
    myArray(15, 11) = 43.4
    myArray(15, 12) = 15.6
    myArray(15, 13) = 43.9
    
    myArray(16, 1) = 2008
    myArray(16, 2) = 17.5
    myArray(16, 3) = 11.1
    myArray(16, 4) = 40.1
    myArray(16, 5) = 34
    myArray(16, 6) = 62.6
    myArray(16, 7) = 126.7
    myArray(16, 8) = 287.2
    myArray(16, 9) = 138.8
    myArray(16, 10) = 89.3
    myArray(16, 11) = 30.4
    myArray(16, 12) = 16.6
    myArray(16, 13) = 15.8
    
    myArray(17, 1) = 2009
    myArray(17, 2) = 13.3
    myArray(17, 3) = 16
    myArray(17, 4) = 51.6
    myArray(17, 5) = 30.6
    myArray(17, 6) = 112.6
    myArray(17, 7) = 55.6
    myArray(17, 8) = 335.8
    myArray(17, 9) = 212.3
    myArray(17, 10) = 30.8
    myArray(17, 11) = 61.1
    myArray(17, 12) = 39.7
    myArray(17, 13) = 40.5
    
    myArray(18, 1) = 2010
    myArray(18, 2) = 40.7
    myArray(18, 3) = 50.4
    myArray(18, 4) = 73.8
    myArray(18, 5) = 61
    myArray(18, 6) = 84
    myArray(18, 7) = 37
    myArray(18, 8) = 171
    myArray(18, 9) = 486.1
    myArray(18, 10) = 316.9
    myArray(18, 11) = 19.4
    myArray(18, 12) = 13.5
    myArray(18, 13) = 24.5
    
    myArray(19, 1) = 2011
    myArray(19, 2) = 7.9
    myArray(19, 3) = 31
    myArray(19, 4) = 26.5
    myArray(19, 5) = 133.2
    myArray(19, 6) = 103.3
    myArray(19, 7) = 374.6
    myArray(19, 8) = 645.1
    myArray(19, 9) = 268.2
    myArray(19, 10) = 153.2
    myArray(19, 11) = 26.5
    myArray(19, 12) = 65.8
    myArray(19, 13) = 10.5
    
    myArray(20, 1) = 2012
    myArray(20, 2) = 14.5
    myArray(20, 3) = 2.3
    myArray(20, 4) = 44.9
    myArray(20, 5) = 81.6
    myArray(20, 6) = 16.8
    myArray(20, 7) = 75.1
    myArray(20, 8) = 252.5
    myArray(20, 9) = 483.7
    myArray(20, 10) = 190.1
    myArray(20, 11) = 66.6
    myArray(20, 12) = 52.6
    myArray(20, 13) = 56
    
    myArray(21, 1) = 2013
    myArray(21, 2) = 28.5
    myArray(21, 3) = 35.2
    myArray(21, 4) = 40
    myArray(21, 5) = 56.3
    myArray(21, 6) = 123.5
    myArray(21, 7) = 102.1
    myArray(21, 8) = 308.2
    myArray(21, 9) = 173.6
    myArray(21, 10) = 117.5
    myArray(21, 11) = 12.2
    myArray(21, 12) = 58.2
    myArray(21, 13) = 40.3
    
    myArray(22, 1) = 2014
    myArray(22, 2) = 4.9
    myArray(22, 3) = 14.7
    myArray(22, 4) = 40.9
    myArray(22, 5) = 62.1
    myArray(22, 6) = 34.6
    myArray(22, 7) = 73.9
    myArray(22, 8) = 239
    myArray(22, 9) = 218.7
    myArray(22, 10) = 144
    myArray(22, 11) = 119.5
    myArray(22, 12) = 28.9
    myArray(22, 13) = 38.9
    
    myArray(23, 1) = 2015
    myArray(23, 2) = 12.7
    myArray(23, 3) = 21.5
    myArray(23, 4) = 23.3
    myArray(23, 5) = 87.6
    myArray(23, 6) = 27.5
    myArray(23, 7) = 86
    myArray(23, 8) = 136.8
    myArray(23, 9) = 64.2
    myArray(23, 10) = 29
    myArray(23, 11) = 69
    myArray(23, 12) = 128.6
    myArray(23, 13) = 41.8
    
    myArray(24, 1) = 2016
    myArray(24, 2) = 8
    myArray(24, 3) = 43.6
    myArray(24, 4) = 16.5
    myArray(24, 5) = 118.3
    myArray(24, 6) = 107.2
    myArray(24, 7) = 36.2
    myArray(24, 8) = 364.3
    myArray(24, 9) = 82
    myArray(24, 10) = 55
    myArray(24, 11) = 95.9
    myArray(24, 12) = 33.5
    myArray(24, 13) = 44.3
    
    myArray(25, 1) = 2017
    myArray(25, 2) = 13.9
    myArray(25, 3) = 32.2
    myArray(25, 4) = 6.5
    myArray(25, 5) = 42.9
    myArray(25, 6) = 14.3
    myArray(25, 7) = 15.6
    myArray(25, 8) = 788.1
    myArray(25, 9) = 291.5
    myArray(25, 10) = 43.3
    myArray(25, 11) = 14.1
    myArray(25, 12) = 23.8
    myArray(25, 13) = 18.8
    
    myArray(26, 1) = 2018
    myArray(26, 2) = 14
    myArray(26, 3) = 31.6
    myArray(26, 4) = 62.2
    myArray(26, 5) = 117
    myArray(26, 6) = 82.7
    myArray(26, 7) = 88.9
    myArray(26, 8) = 185.8
    myArray(26, 9) = 282.7
    myArray(26, 10) = 124.6
    myArray(26, 11) = 99.8
    myArray(26, 12) = 48.3
    myArray(26, 13) = 25.8
    
    myArray(27, 1) = 2019
    myArray(27, 2) = 0.6
    myArray(27, 3) = 25.5
    myArray(27, 4) = 26.9
    myArray(27, 5) = 43.9
    myArray(27, 6) = 15.1
    myArray(27, 7) = 84.9
    myArray(27, 8) = 234.7
    myArray(27, 9) = 90.7
    myArray(27, 10) = 102.8
    myArray(27, 11) = 81.9
    myArray(27, 12) = 120.6
    myArray(27, 13) = 18
    
    myArray(28, 1) = 2020
    myArray(28, 2) = 59.7
    myArray(28, 3) = 63.1
    myArray(28, 4) = 21.7
    myArray(28, 5) = 15.1
    myArray(28, 6) = 86.4
    myArray(28, 7) = 121.9
    myArray(28, 8) = 356.4
    myArray(28, 9) = 481.7
    myArray(28, 10) = 167.2
    myArray(28, 11) = 18.9
    myArray(28, 12) = 45.9
    myArray(28, 13) = 5.5
    
    myArray(29, 1) = 2021
    myArray(29, 2) = 17.8
    myArray(29, 3) = 9.2
    myArray(29, 4) = 75.3
    myArray(29, 5) = 54.7
    myArray(29, 6) = 135.9
    myArray(29, 7) = 44.8
    myArray(29, 8) = 117.6
    myArray(29, 9) = 230
    myArray(29, 10) = 250.8
    myArray(29, 11) = 49.5
    myArray(29, 12) = 67.9
    myArray(29, 13) = 5.4
    
    myArray(30, 1) = 2022
    myArray(30, 2) = 3.3
    myArray(30, 3) = 3.3
    myArray(30, 4) = 57.6
    myArray(30, 5) = 51.6
    myArray(30, 6) = 9.8
    myArray(30, 7) = 168
    myArray(30, 8) = 117
    myArray(30, 9) = 366.6
    myArray(30, 10) = 133.3
    myArray(30, 11) = 98.2
    myArray(30, 12) = 43.2
    myArray(30, 13) = 28.8
    
    data_CHEONAN = myArray

End Function





Function data_BORYUNG() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1993
    myArray(1, 2) = 4.6
    myArray(1, 3) = 66.7
    myArray(1, 4) = 21.6
    myArray(1, 5) = 32.6
    myArray(1, 6) = 68.5
    myArray(1, 7) = 225.5
    myArray(1, 8) = 241.5
    myArray(1, 9) = 206.5
    myArray(1, 10) = 141.5
    myArray(1, 11) = 32
    myArray(1, 12) = 87.4
    myArray(1, 13) = 33.7
    
    myArray(2, 1) = 1994
    myArray(2, 2) = 17.9
    myArray(2, 3) = 6
    myArray(2, 4) = 58.2
    myArray(2, 5) = 51.5
    myArray(2, 6) = 135.5
    myArray(2, 7) = 207
    myArray(2, 8) = 137
    myArray(2, 9) = 443.5
    myArray(2, 10) = 21
    myArray(2, 11) = 155
    myArray(2, 12) = 17.5
    myArray(2, 13) = 18.9
    
    myArray(3, 1) = 1995
    myArray(3, 2) = 15.7
    myArray(3, 3) = 11
    myArray(3, 4) = 19.6
    myArray(3, 5) = 65.5
    myArray(3, 6) = 49.5
    myArray(3, 7) = 26.5
    myArray(3, 8) = 144.5
    myArray(3, 9) = 996.5
    myArray(3, 10) = 70.5
    myArray(3, 11) = 24.5
    myArray(3, 12) = 23
    myArray(3, 13) = 12.7
    
    myArray(4, 1) = 1996
    myArray(4, 2) = 33.4
    myArray(4, 3) = 6.8
    myArray(4, 4) = 104.5
    myArray(4, 5) = 34
    myArray(4, 6) = 22.5
    myArray(4, 7) = 235
    myArray(4, 8) = 192.5
    myArray(4, 9) = 44.5
    myArray(4, 10) = 14
    myArray(4, 11) = 106.5
    myArray(4, 12) = 74.2
    myArray(4, 13) = 31.7
    
    myArray(5, 1) = 1997
    myArray(5, 2) = 15.1
    myArray(5, 3) = 38.4
    myArray(5, 4) = 30.5
    myArray(5, 5) = 57.5
    myArray(5, 6) = 203
    myArray(5, 7) = 272
    myArray(5, 8) = 353.5
    myArray(5, 9) = 211.5
    myArray(5, 10) = 23
    myArray(5, 11) = 10
    myArray(5, 12) = 193.5
    myArray(5, 13) = 34.3
    
    myArray(6, 1) = 1998
    myArray(6, 2) = 29.9
    myArray(6, 3) = 40.2
    myArray(6, 4) = 30.5
    myArray(6, 5) = 138
    myArray(6, 6) = 100
    myArray(6, 7) = 209.5
    myArray(6, 8) = 263
    myArray(6, 9) = 341.7
    myArray(6, 10) = 150.3
    myArray(6, 11) = 61
    myArray(6, 12) = 29.3
    myArray(6, 13) = 3.8
    
    myArray(7, 1) = 1999
    myArray(7, 2) = 7.9
    myArray(7, 3) = 9.5
    myArray(7, 4) = 71
    myArray(7, 5) = 88.5
    myArray(7, 6) = 124.5
    myArray(7, 7) = 192.5
    myArray(7, 8) = 98
    myArray(7, 9) = 180
    myArray(7, 10) = 292.5
    myArray(7, 11) = 169
    myArray(7, 12) = 24.9
    myArray(7, 13) = 25.8
    
    myArray(8, 1) = 2000
    myArray(8, 2) = 42.1
    myArray(8, 3) = 3.2
    myArray(8, 4) = 7
    myArray(8, 5) = 35
    myArray(8, 6) = 53.5
    myArray(8, 7) = 159.5
    myArray(8, 8) = 155
    myArray(8, 9) = 701.5
    myArray(8, 10) = 241
    myArray(8, 11) = 46
    myArray(8, 12) = 39.5
    myArray(8, 13) = 32.1
    
    myArray(9, 1) = 2001
    myArray(9, 2) = 73.3
    myArray(9, 3) = 46
    myArray(9, 4) = 15.9
    myArray(9, 5) = 26
    myArray(9, 6) = 17
    myArray(9, 7) = 129
    myArray(9, 8) = 286.5
    myArray(9, 9) = 170
    myArray(9, 10) = 10
    myArray(9, 11) = 85
    myArray(9, 12) = 13
    myArray(9, 13) = 32
    
    myArray(10, 1) = 2002
    myArray(10, 2) = 50.8
    myArray(10, 3) = 5.5
    myArray(10, 4) = 32
    myArray(10, 5) = 169
    myArray(10, 6) = 155.5
    myArray(10, 7) = 72
    myArray(10, 8) = 217.5
    myArray(10, 9) = 477
    myArray(10, 10) = 27
    myArray(10, 11) = 134
    myArray(10, 12) = 61.1
    myArray(10, 13) = 51.8
    
    myArray(11, 1) = 2003
    myArray(11, 2) = 30.7
    myArray(11, 3) = 44.5
    myArray(11, 4) = 39.5
    myArray(11, 5) = 168.5
    myArray(11, 6) = 78.5
    myArray(11, 7) = 153
    myArray(11, 8) = 309.5
    myArray(11, 9) = 310
    myArray(11, 10) = 128
    myArray(11, 11) = 23
    myArray(11, 12) = 45.5
    myArray(11, 13) = 13
    
    myArray(12, 1) = 2004
    myArray(12, 2) = 22.1
    myArray(12, 3) = 28.5
    myArray(12, 4) = 45.7
    myArray(12, 5) = 58
    myArray(12, 6) = 105.5
    myArray(12, 7) = 234.5
    myArray(12, 8) = 263.5
    myArray(12, 9) = 164
    myArray(12, 10) = 195
    myArray(12, 11) = 4
    myArray(12, 12) = 56.5
    myArray(12, 13) = 38.9
    
    myArray(13, 1) = 2005
    myArray(13, 2) = 5.8
    myArray(13, 3) = 35.8
    myArray(13, 4) = 30
    myArray(13, 5) = 73.5
    myArray(13, 6) = 48.5
    myArray(13, 7) = 156
    myArray(13, 8) = 260.5
    myArray(13, 9) = 291.5
    myArray(13, 10) = 282.5
    myArray(13, 11) = 21
    myArray(13, 12) = 18
    myArray(13, 13) = 43.4
    
    myArray(14, 1) = 2006
    myArray(14, 2) = 27
    myArray(14, 3) = 25.9
    myArray(14, 4) = 10.6
    myArray(14, 5) = 81.5
    myArray(14, 6) = 94.5
    myArray(14, 7) = 114.5
    myArray(14, 8) = 321
    myArray(14, 9) = 21.5
    myArray(14, 10) = 23.5
    myArray(14, 11) = 24.5
    myArray(14, 12) = 61.5
    myArray(14, 13) = 25.4
    
    myArray(15, 1) = 2007
    myArray(15, 2) = 23.4
    myArray(15, 3) = 29.8
    myArray(15, 4) = 102
    myArray(15, 5) = 29.5
    myArray(15, 6) = 79
    myArray(15, 7) = 85
    myArray(15, 8) = 214
    myArray(15, 9) = 239.5
    myArray(15, 10) = 384
    myArray(15, 11) = 59
    myArray(15, 12) = 17.5
    myArray(15, 13) = 33.1
    
    myArray(16, 1) = 2008
    myArray(16, 2) = 20.9
    myArray(16, 3) = 10.8
    myArray(16, 4) = 48.2
    myArray(16, 5) = 40.5
    myArray(16, 6) = 78.9
    myArray(16, 7) = 101.3
    myArray(16, 8) = 257.2
    myArray(16, 9) = 119.5
    myArray(16, 10) = 46.9
    myArray(16, 11) = 26.7
    myArray(16, 12) = 37.6
    myArray(16, 13) = 25
    
    myArray(17, 1) = 2009
    myArray(17, 2) = 18.5
    myArray(17, 3) = 23.3
    myArray(17, 4) = 55.1
    myArray(17, 5) = 41.5
    myArray(17, 6) = 154.5
    myArray(17, 7) = 115.1
    myArray(17, 8) = 320.9
    myArray(17, 9) = 176.6
    myArray(17, 10) = 25.5
    myArray(17, 11) = 39.5
    myArray(17, 12) = 52.9
    myArray(17, 13) = 58
    
    myArray(18, 1) = 2010
    myArray(18, 2) = 30.1
    myArray(18, 3) = 73.5
    myArray(18, 4) = 75.9
    myArray(18, 5) = 58.5
    myArray(18, 6) = 122.8
    myArray(18, 7) = 60.8
    myArray(18, 8) = 396.5
    myArray(18, 9) = 402.7
    myArray(18, 10) = 213.1
    myArray(18, 11) = 19.2
    myArray(18, 12) = 16.3
    myArray(18, 13) = 32.9
    
    myArray(19, 1) = 2011
    myArray(19, 2) = 11.1
    myArray(19, 3) = 37.5
    myArray(19, 4) = 18
    myArray(19, 5) = 72.1
    myArray(19, 6) = 115.3
    myArray(19, 7) = 318
    myArray(19, 8) = 723.1
    myArray(19, 9) = 289.4
    myArray(19, 10) = 70.8
    myArray(19, 11) = 13.9
    myArray(19, 12) = 61.3
    myArray(19, 13) = 12.5
    
    myArray(20, 1) = 2012
    myArray(20, 2) = 24.2
    myArray(20, 3) = 9.2
    myArray(20, 4) = 45
    myArray(20, 5) = 68.9
    myArray(20, 6) = 14.6
    myArray(20, 7) = 76.8
    myArray(20, 8) = 231.1
    myArray(20, 9) = 450.2
    myArray(20, 10) = 207.7
    myArray(20, 11) = 65
    myArray(20, 12) = 61.1
    myArray(20, 13) = 65.2
    
    myArray(21, 1) = 2013
    myArray(21, 2) = 28.4
    myArray(21, 3) = 40.7
    myArray(21, 4) = 53.4
    myArray(21, 5) = 68.2
    myArray(21, 6) = 116.6
    myArray(21, 7) = 159.9
    myArray(21, 8) = 267.5
    myArray(21, 9) = 214.6
    myArray(21, 10) = 320
    myArray(21, 11) = 10.9
    myArray(21, 12) = 81.1
    myArray(21, 13) = 26.4
    
    myArray(22, 1) = 2014
    myArray(22, 2) = 3.4
    myArray(22, 3) = 20.5
    myArray(22, 4) = 56.3
    myArray(22, 5) = 70
    myArray(22, 6) = 47.1
    myArray(22, 7) = 125.8
    myArray(22, 8) = 104
    myArray(22, 9) = 168.5
    myArray(22, 10) = 152
    myArray(22, 11) = 156
    myArray(22, 12) = 39.9
    myArray(22, 13) = 66.6
    
    myArray(23, 1) = 2015
    myArray(23, 2) = 29.9
    myArray(23, 3) = 23.4
    myArray(23, 4) = 30.9
    myArray(23, 5) = 129.7
    myArray(23, 6) = 38.8
    myArray(23, 7) = 83.9
    myArray(23, 8) = 94.7
    myArray(23, 9) = 30.2
    myArray(23, 10) = 13.3
    myArray(23, 11) = 90
    myArray(23, 12) = 155.6
    myArray(23, 13) = 65
    
    myArray(24, 1) = 2016
    myArray(24, 2) = 7.8
    myArray(24, 3) = 54.2
    myArray(24, 4) = 18.7
    myArray(24, 5) = 105.1
    myArray(24, 6) = 146.5
    myArray(24, 7) = 23.7
    myArray(24, 8) = 200.2
    myArray(24, 9) = 5.1
    myArray(24, 10) = 73.4
    myArray(24, 11) = 108
    myArray(24, 12) = 5.6
    myArray(24, 13) = 44.5
    
    myArray(25, 1) = 2017
    myArray(25, 2) = 14.8
    myArray(25, 3) = 30.2
    myArray(25, 4) = 14.4
    myArray(25, 5) = 57.6
    myArray(25, 6) = 58.9
    myArray(25, 7) = 21.1
    myArray(25, 8) = 278.1
    myArray(25, 9) = 210
    myArray(25, 10) = 90
    myArray(25, 11) = 26.6
    myArray(25, 12) = 15.9
    myArray(25, 13) = 38.6
    
    myArray(26, 1) = 2018
    myArray(26, 2) = 15
    myArray(26, 3) = 33.6
    myArray(26, 4) = 92
    myArray(26, 5) = 128.1
    myArray(26, 6) = 104.5
    myArray(26, 7) = 71
    myArray(26, 8) = 262.7
    myArray(26, 9) = 239.6
    myArray(26, 10) = 158.2
    myArray(26, 11) = 154.7
    myArray(26, 12) = 46.7
    myArray(26, 13) = 31.1
    
    myArray(27, 1) = 2019
    myArray(27, 2) = 1.9
    myArray(27, 3) = 17.8
    myArray(27, 4) = 18.2
    myArray(27, 5) = 71.9
    myArray(27, 6) = 31.3
    myArray(27, 7) = 56
    myArray(27, 8) = 149
    myArray(27, 9) = 131.3
    myArray(27, 10) = 118.7
    myArray(27, 11) = 63.9
    myArray(27, 12) = 130.6
    myArray(27, 13) = 31.3
    
    myArray(28, 1) = 2020
    myArray(28, 2) = 49.4
    myArray(28, 3) = 75.3
    myArray(28, 4) = 22.8
    myArray(28, 5) = 16.5
    myArray(28, 6) = 92.4
    myArray(28, 7) = 139.7
    myArray(28, 8) = 345.9
    myArray(28, 9) = 321.5
    myArray(28, 10) = 177.1
    myArray(28, 11) = 16.2
    myArray(28, 12) = 35.4
    myArray(28, 13) = 9.7
    
    myArray(29, 1) = 2021
    myArray(29, 2) = 32
    myArray(29, 3) = 18.7
    myArray(29, 4) = 76.1
    myArray(29, 5) = 43.4
    myArray(29, 6) = 110
    myArray(29, 7) = 55
    myArray(29, 8) = 131.3
    myArray(29, 9) = 253.7
    myArray(29, 10) = 215.9
    myArray(29, 11) = 39.6
    myArray(29, 12) = 117.8
    myArray(29, 13) = 14.4
    
    myArray(30, 1) = 2022
    myArray(30, 2) = 8.4
    myArray(30, 3) = 5.3
    myArray(30, 4) = 60.9
    myArray(30, 5) = 34.8
    myArray(30, 6) = 5.7
    myArray(30, 7) = 225
    myArray(30, 8) = 119.9
    myArray(30, 9) = 637.1
    myArray(30, 10) = 102
    myArray(30, 11) = 112
    myArray(30, 12) = 23.3
    myArray(30, 13) = 14
    
    data_BORYUNG = myArray
    

End Function
