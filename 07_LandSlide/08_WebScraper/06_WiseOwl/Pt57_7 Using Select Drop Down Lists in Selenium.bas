Attribute VB_Name = "Module1"
Option Explicit

Private cd As Selenium.ChromeDriver

Sub NvidiaSelect()

    Dim ddl1 As Selenium.SelectElement, ddl2 As Selenium.SelectElement, ddl3 As Selenium.SelectElement
    Dim op1 As Selenium.WebElement, op2 As Selenium.WebElement, op3 As Selenium.WebElement
    Dim ws As Worksheet
    Dim r As Integer
    Dim FileURL As String
    
    Set cd = New Selenium.ChromeDriver
    
    cd.Start
    cd.Get "https://www.nvidia.co.uk/Download/index.aspx?lang=en-uk"
    
    Set ddl1 = cd.FindElementByCss("#selProductSeriesType").AsSelect
    Set ddl2 = cd.FindElementByCss("#selProductSeries").AsSelect
    Set ddl3 = cd.FindElementByCss("#selProductFamily").AsSelect
    
    ddl1.SelectByText ActiveCell.Offset(0, -2).End(xlUp).Value
    ddl2.SelectByText ActiveCell.Offset(0, -1).End(xlUp).Value
    ddl3.SelectByText ActiveCell.Value
    
    cd.FindElementByCss("#imgSearch").Click
    cd.FindElementByCss("#imgDwnldBtn").Click
    
    FileURL = cd.FindElementByCss("#mainContent > table > tbody > tr > td > a").Attribute("href")
    FileURL = "https:" & FileURL
    Debug.Print FileURL
    
    DownloadFile FileURL, Environ("UserProfile") & "\Downloads"
    
'    Set ws = ThisWorkbook.Worksheets.Add
'
'    For Each op1 In ddl1.Options
'        'Debug.Print op1.Text, op1.Value
'
'        ddl1.SelectByText op1.Text
'
'        r = r + 1
'        ws.Cells(r, 1).Value = op1.Text
'
'        For Each op2 In ddl2.Options
'            'Debug.Print , op2.Text, op2.Value
'
'            If op2.Value <> "All" Then
'                ddl2.SelectByText op2.Text
'
'                r = r + 1
'                ws.Cells(r, 2).Value = op2.Text
'
'                If cd.FindElementByCss("#selProductFamily").IsDisplayed Then
'
'                    For Each op3 In ddl3.Options
'                        'Debug.Print , , op3.Text, op3.Value
'
'                        r = r + 1
'                        ws.Cells(r, 3).Value = op3.Text
'
'                    Next op3
'
'                End If
'
'            End If
'
'        Next op2
'    Next op1
    
End Sub

Sub AmazonSelect()

    Dim ddl As Selenium.SelectElement
    Dim op As Selenium.WebElement
    Dim srs As Selenium.WebElements
    
    Set cd = New Selenium.ChromeDriver
    
    cd.Start
    cd.Get "https://www.amazon.co.uk/"
    
    cd.FindElementByCss("#sp-cc-accept").Click
    
    Set ddl = cd.FindElementByCss("#searchDropdownBox").AsSelect
    
'    For Each op In ddl.Options
'        Debug.Print op.Text, op.Value
'    Next op
    
    'ddl.SelectByValue "search-alias=stripbooks"
    ddl.SelectByText "Books"
    
    cd.FindElementByCss("#twotabsearchtextbox").SendKeys "owl"
    cd.FindElementByCss("#nav-search-submit-button").Click
    
    Set srs = cd.FindElementsByCss(".s-result-item")
    
    Debug.Print srs.Count
    
    srs.Text.ToExcel ThisWorkbook.Worksheets.Add.Range("A1")
    
End Sub

