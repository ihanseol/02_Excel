Attribute VB_Name = "usage_dom"

Private Sub Get_DOM_1()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"

  Dim html As Object
  Set html = CreateObject("htmlfile")
  html.Open
  html.Write driver.PageSource()
  html.Close
  
  Debug.Print html.body.innerText
  driver.Quit
End Sub


Private Sub Get_DOM_2()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"

  Dim html As Object
  Set html = CreateObject("htmlfile")
  html.body.innerHTML = driver.ExecuteScript("return document.body.innerHTML;")
  
  Debug.Print html.body.innerText
  driver.Quit
End Sub
