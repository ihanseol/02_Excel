Attribute VB_Name = "example_amazon"

Const css_input = "#twotabsearchtextbox"
Const css_spinner = "#centerBelowPlusspacer"
Const css_titles = "#s-results-list-atf a.s-access-detail-page"
Const css_bt_next = "#pagnNextLink #pagnNextString"

Private Sub Search_Amazon()
  Const search_input = "drum carrot"
  
  Dim driver As New ChromeDriver
  driver.Get "http://www.amazon.com/s"
  
  ' type in the field and submit
  driver.FindElementByCss(css_input) _
        .SendKeys(search_input) _
        .Submit
  
  Dim By As New By
  Do
    ' wait for the loading bar to disapear
    driver.WaitNotElement By.css(css_spinner), 7000
    
    ' handle results
    For Each ele In driver.FindElementsByCss(css_titles)
      Debug.Print ele.Text
    Next
    
    ' click next page or exit the loop if the link is not present
    Set bt_next = driver.FindElementByCss(css_bt_next, timeout:=0, Raise:=False)
    If bt_next Is Nothing Then _
      Exit Do
    bt_next.Click
  Loop
End Sub

