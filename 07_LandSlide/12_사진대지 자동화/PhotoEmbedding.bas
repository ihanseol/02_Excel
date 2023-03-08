Attribute VB_Name = "PhotoEmbedding"

Function ConvertPointToCm(ByVal pnt As Double) As Double
    ConvertPointToCm = pnt * 0.03527778
End Function


Function ConvertCmToPoint(ByVal cm As Double) As Double
    ConvertCmToPoint = cm * 28.34646
End Function


'Sub InsertImage_LinkedPicture()
'
'
'    Dim fd As FileDialog
'    Set fd = Application.FileDialog(msoFileDialogFilePicker)
'    fd.AllowMultiSelect = False
'
'    If fd.Show = -1 Then
'        Dim selectedFile As String
'        selectedFile = fd.SelectedItems(1)
'
'        Dim area As Range
'        ' Set area = Application.InputBox(prompt:="Select the area where you want to insert the image.", Type:=8)
'
'        Set area = MakeRangeFromIndex(FindPageIndex(ActiveCell.Row))
'
'        If Not area Is Nothing Then
'
'            Dim Image As Picture
'            Set Image = ActiveSheet.Pictures.Insert(selectedFile)
'
'            ' Resize image
'            Image.ShapeRange.LockAspectRatio = msoTrue
'
'            Image.Height = ConvertCmToPoint(8)    ' 8 cm * 28.35 points/cm
'            Image.Width = ConvertCmToPoint(10.7)  ' 10.7 cm * 28.35 points/cm
'
'            ' Position image
'            Dim centerX As Double
'            Dim centerY As Double
'
'            centerX = area.Left + (area.Width / 2) - (Image.Width / 2)
'            centerY = area.Top + (area.Height / 2) - (Image.Height / 2)
'            Image.Top = centerY
'            Image.Left = centerX
'
'        End If
'    End If
'
'
'    Image.Select
'    Selection.Placement = xlMoveAndSize
'
'End Sub



Sub InsertImage_EmbedPicture()
Attribute InsertImage_EmbedPicture.VB_ProcData.VB_Invoke_Func = "e\n14"

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = False
    
    If fd.Show = -1 Then
        Dim selectedFile As String
        selectedFile = fd.SelectedItems(1)
        
        Dim area As Range
        ' Set area = Application.InputBox(prompt:="Select the area where you want to insert the image.", Type:=8)
        
        Set area = MakeRangeFromIndex(FindPageIndex(ActiveCell.Row))
        
        If Not area Is Nothing Then
        
            Dim Image As Shape
            Set Image = ActiveSheet.Shapes.AddPicture(Filename:=selectedFile, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
            Left:=area.Left, Top:=area.Top, Width:=-1, Height:=-1)
            
            ' Resize image
            Image.LockAspectRatio = msoTrue
            
            Image.Height = ConvertCmToPoint(8)    ' 8 cm * 28.35 points/cm
            Image.Width = ConvertCmToPoint(10.7)  ' 10.7 cm * 28.35 points/cm
            
            ' Position image
            Dim centerX As Double
            Dim centerY As Double
            
            centerX = area.Left + (area.Width / 2) - (Image.Width / 2)
            centerY = area.Top + (area.Height / 2) - (Image.Height / 2)
            Image.Top = centerY
            Image.Left = centerX
            
        End If
    End If
    
End Sub




Public Function GetFilePath()
    Set myFile = Application.FileDialog(msoFileDialogOpen)
    With myFile
    .Title = "Choose File"
    .AllowMultiSelect = False
    If .Show <> -1 Then
        Exit Function
    End If
    
    FileSelected = .SelectedItems(1)
    End With
    
    GetFilePath = FileSelected
End Function
    
    

Function GetInitialPositionArray() As Integer()
    Dim InitialPosition() As Integer
    Dim page, volume, index As Integer
    Const PAGE_LIMIT As Integer = 12
    
    volume = PAGE_LIMIT * 2 + 1
    ReDim InitialPosition(0 To volume)
    
    For page = 0 To PAGE_LIMIT
        index = page * 2
        InitialPosition(index) = 3 + page * 24
        index = index + 1
        InitialPosition(index) = 14 + page * 24
    Next page
    
    GetInitialPositionArray = InitialPosition
End Function


Function FindPageIndex(iRow As Integer) As Integer
    Dim InitialPosition() As Integer
    Dim i As Integer
    
    InitialPosition = GetInitialPositionArray()
    
    ' Print the contents of the InitialPosition array
    ' iRow = ActiveCell.Row
       
    For i = LBound(InitialPosition) To UBound(InitialPosition)
    
        If InitialPosition(i + 1) > iRow Then
             index = i
             Exit For
        End If
    
        Debug.Print "InitialPosition(" & i & ") = " & InitialPosition(i)
    Next i
    
    FindPageIndex = index
End Function


Function MakeRangeFromIndex(index As Integer) As Range
    Dim InitialPosition() As Integer
    Dim h1, h2 As Integer
    Dim rng As Range
    
    InitialPosition = GetInitialPositionArray()
    h1 = InitialPosition(index)
    h2 = h1 + 8
    
    Set rng = Range("$A$" & CStr(h1) & ":$J" & CStr(h2))
    Set MakeRangeFromIndex = rng
End Function


Sub DeleteAllImages_LinkedPictures()
Attribute DeleteAllImages_LinkedPictures.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim shp As Shape
        
    For Each shp In ActiveSheet.Shapes
        Debug.Print shp.Name & " is a " & shp.Type & " shape"
        
        If shp.Type = msoLinkedPicture Then
            shp.Select
            Selection.Delete
        ElseIf shp.Type = msoAutoShape Then
            Debug.Print shp.Name & " is an AutoShape"
        End If
    Next shp
End Sub

Sub DeleteAllImages_EmbededPictures()
Attribute DeleteAllImages_EmbededPictures.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim shp As Shape
    
    ' Loop through all shapes on the active worksheet
    For Each shp In ActiveSheet.Shapes
        ' Check if shape is an image
        Debug.Print shp.Name & " is a " & shp.Type & " shape"
        If shp.Type = msoPicture Then
            ' Select and delete the image
            shp.Select
            Selection.Delete
        End If
    Next shp
End Sub



'Sub GetAreaDetails()
'
'    Dim area As Range
'    Set area = Application.InputBox(prompt:="Please select a cell range.", Type:=8)
'
'    If Not area Is Nothing Then
'        Dim leftPosition As Double
'        Dim topPosition As Double
'        Dim widthSize As Double
'        Dim heightSize As Double
'
'        leftPosition = area.Left
'        topPosition = area.Top
'        widthSize = area.Width
'        heightSize = area.Height
'
'        MsgBox "Area Left Position: " & leftPosition & vbCrLf & _
'               "Area Top Position: " & topPosition & vbCrLf & _
'               "Area Width Size: " & widthSize & vbCrLf & _
'               "Area Height Size: " & heightSize
'    End If
'
'End Sub
'


'Sub AddToInitialPositionWithPageNumber()
'    Dim InitialPosition() As Integer
'    Dim page, i, volume  As Integer
'    Dim index As Integer
'    Const PAGE_LIMIT As Integer = 12
'
'
'    ' InitialPosition = Array(3, 14,/ 27, 38,/ 51, 62,/ 75, 86,/ 99, 110,/ 123, 134,/ 147, 158,/ 171, 182,/ 195, 206)
'    ' Print contents of array to immediate window
'
'    volume = PAGE_LIMIT * 2 + 1
'    ReDim InitialPosition(0 To volume)
'
'    For page = 0 To PAGE_LIMIT
'        InitialPosition(page * 2) = 3 + page * 24
'        InitialPosition(page * 2 + 1) = 14 + page * 24
'    Next page
'
'     For i = LBound(InitialPosition) To UBound(InitialPosition)
'        Debug.Print "InitialPosition(" & i & ") = " & InitialPosition(i)
'    Next i
'End Sub
'
'
'Sub AddToInitialPosition()
'    Dim InitialPosition() As Integer
'
'    InitialPosition = Array(3, 14, 27, 38, 51, 62, 75, 86, 99, 110, 123, 134, 147, 158, 171, 182, 195, 206)
'
'    ' Print contents of array to immediate window
'    Dim i As Integer
'    For i = LBound(InitialPosition) To UBound(InitialPosition)
'        Debug.Print "InitialPosition(" & i & ") = " & InitialPosition(i)
'    Next i
'End Sub


    
