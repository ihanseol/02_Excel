Attribute VB_Name = "usage_javascript"

''
' Executes a piece of Javascript on the page.
''
Private Sub Execute_Script()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  Dim title
  title = driver.ExecuteScript("return document.title;")
  Debug.Assert "Wikipedia, the free encyclopedia" = title
  
  driver.Quit
End Sub

''
' Executes a piece of Javascript on a web element.
' The web element is the context itself which is "this".
''
Private Sub Execute_Script_On_Element()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  Dim name
  name = driver.FindElementById("searchInput") _
               .ExecuteScript("return this.name;")
                
  Debug.Assert "search" = name
  
  driver.Quit
End Sub

''
' Executes a piece of Javascript on a collection of web elements.
' The web element is the context itself which is "this".
''
Private Sub Execute_Script_On_Elements()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  Dim links
  Set links = driver.FindElementsByTag("a") _
                    .ExecuteScript("return this.href;")
  
  driver.Quit
End Sub

''
' Executes an asynchronous piece of Javascript.
' The script returns once "callback" is called.
''
Private Sub Execute_Script_Async()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  Dim response
  response = driver.ExecuteAsyncScript( _
    "var r = new XMLHttpRequest();" & _
    "r.onreadystatechange = function(){" & _
    " if(r.readyState == XMLHttpRequest.DONE)" & _
    "  callback(this.responseText);" & _
    "};" & _
    "r.open('GET', 'wiki/Euro');" & _
    "r.send();")
    
  Debug.Print response
  driver.Quit
End Sub

