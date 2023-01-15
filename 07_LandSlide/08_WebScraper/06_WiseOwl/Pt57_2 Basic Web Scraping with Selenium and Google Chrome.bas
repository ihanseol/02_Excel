Attribute VB_Name = "Module1"
Option Explicit

Private ch As Selenium.ChromeDriver

Sub ScrapeWiseOwlVideos()

    Dim FindBy As New Selenium.By
    Dim ResultSections As Selenium.WebElements
    Dim ResultSection As Selenium.WebElement
    Dim VideoTable As Selenium.WebElement
    
    If SearchSheet.Range("C2").Value = "" Then
        Exit Sub
    End If
    
    Set ch = New Selenium.ChromeDriver
    
    ch.Start baseUrl:="https://www.wiseowl.co.uk"
    ch.Get "/"
    
    If Not ch.IsElementPresent(FindBy.Name("what"), 3000) Then
        ch.Quit
        MsgBox "Could not find search input box", vbExclamation
        Exit Sub
    End If
    
    ch.FindElementByName("what").SendKeys SearchSheet.Range("C2").Value
    
    If Not ch.IsElementPresent(FindBy.Class("search__submit"), 3000) Then
        ch.Quit
        MsgBox "Could not find submit button", vbExclamation
        Exit Sub
    End If
    
    ch.FindElementByClass("search__submit").Click
    
    Set ResultSections = ch.FindElementsByClass("woFormAccordionPart")
    
    If ResultSections.Count = 0 Then
        ch.Quit
        MsgBox "Nothing at all was found", vbExclamation
        Exit Sub
    End If
    
    For Each ResultSection In ResultSections
        If ResultSection.Text Like "Video tutorials (*" Then
            'Debug.Print ResultSection.Text
            ResultSection.Click
            Set VideoTable = ResultSection.FindElementByTag("table")
            Exit For
        End If
    Next ResultSection
    
    If VideoTable Is Nothing Then
        ch.Quit
        MsgBox "No videos were found", vbExclamation
        Exit Sub
    End If
    
    'VideoTable.AsTable.ToExcel ThisWorkbook.Worksheets.Add.Range("A1")
    
    ProcessTable VideoTable
    
End Sub

Sub ProcessTable(TableToProcess As Selenium.WebElement)

    Dim AllRows As Selenium.WebElements
    Dim SingleRow As Selenium.WebElement
    Dim AllRowCells As Selenium.WebElements
    Dim SingleCell As Selenium.WebElement
    Dim OutputSheet As Worksheet
    Dim RowNum As Long, ColNum As Long
    Dim TargetCell As Range
    Dim VideoLinks As Selenium.WebElements
    
    Application.ScreenUpdating = False
    
    Set OutputSheet = ThisWorkbook.Worksheets.Add
    Set AllRows = TableToProcess.FindElementsByTag("tr")
    
    For Each SingleRow In AllRows
        
        RowNum = RowNum + 1
        
        Set AllRowCells = SingleRow.FindElementsByTag("td")
        
        If AllRowCells.Count = 0 Then
            Set AllRowCells = SingleRow.FindElementsByTag("th")
        End If
        
        For Each SingleCell In AllRowCells
            
            ColNum = ColNum + 1
            
            Set TargetCell = OutputSheet.Cells(RowNum, ColNum)
            
            TargetCell.Value = SingleCell.Text
            
            Set VideoLinks = SingleCell.FindElementsByTag("a")
            
            If VideoLinks.Count > 0 Then
                'Debug.Print VideoLinks(1).Attribute("href")
                TargetCell.Hyperlinks.Add TargetCell, VideoLinks(1).Attribute("href")
            End If
            
        Next SingleCell
        
        ColNum = 0
        
    Next SingleRow
    
    OutputSheet.Range("A1").CurrentRegion.EntireColumn.AutoFit
    OutputSheet.ListObjects.Add Source:=OutputSheet.Range("A1").CurrentRegion
    
    Application.ScreenUpdating = True
    
End Sub
