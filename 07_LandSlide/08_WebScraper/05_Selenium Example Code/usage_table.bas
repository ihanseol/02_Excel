Attribute VB_Name = "usage_table"
Private Assert As New Selenium.Assert


Private Sub Scrap_Table()
  Dim driver As New FirefoxDriver
  driver.Get "http://stats.nba.com/league/player/#!/"
  
  driver.FindElementByCss("table.table") _
        .AsTable _
        .ToExcel Map:="(e) => e.firstChild.textContent.trim()"
        
  driver.Quit
End Sub


Private Sub Handle_Table()
  Dim driver As New FirefoxDriver
  driver.Get "http://the-internet.herokuapp.com/tables"
  
  'Print all cells from the second column
  Dim ele As WebElement
  For Each ele In driver.FindElementsByCss("#table1 tbody tr td:nth-child(2)")
      Debug.Print ele.Text
  Next
  
  driver.Quit
End Sub

Private Sub Handle_Table2()
  Dim driver As New FirefoxDriver
  driver.Get "http://the-internet.herokuapp.com/tables"
  
  Dim tbl As TableElement
  Set tbl = driver.FindElementByCss("#table1").AsTable
  
  'Print all cells
  Dim data(): data = tbl.data
  For c = 1 To UBound(data, 1)
    For r = 1 To UBound(data, 1)
      Debug.Print data(r, 2)
    Next
    Debug.Print Empty
  Next
  
  driver.Quit
End Sub
