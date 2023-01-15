Attribute VB_Name = "usage_scroll"

Private Sub Scroll_Element_To_Center()
  Dim drv As New FirefoxDriver
  drv.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  drv.FindElementByCss("#mp-other").ExecuteScript _
    "this.scrollIntoView(true);" & _
    "window.scrollBy(0, -(window.innerHeight - this.clientHeight) / 2);"
  
  Debug.Assert 0
  drv.Quit
End Sub


''
' Finds all the scrollable ancestor and scroll them vertically by the provided amout of pixels
' @element {WebElement}  Web element in a scrollable container or window
' @y {Long} Amount of pixels to vertically scroll
''
Private Function ScrollVertically(element As WebElement, y As Long)
  Const JS_SCROLL_Y = _
    "var y = arguments[0];" & _
    "for (var e=this; e; e=e.parentElement) {" & _
    "  var yy = e.scrollHeight - e.clientHeight;" & _
    "  if (yy === 0) continue;" & _
    "  yy = y < 0 ? Math.max(y, -e.scrollTop) : Math.min(y, yy - e.scrollTop);" & _
    "  if(yy === 0) continue;" & _
    "  e.scrollTop += yy;" & _
    "  if ((y -= yy) == 0) return;" & _
    "}" & _
    "window.scrollBy(0, y);"
    
  element.ExecuteScript JS_SCROLL_Y, y
End Function


''
' Scrolls an element in the center of the view
' @element {WebElement}  Web element in a scrollable container
''
Private Function ScrollIntoViewCenter(element As WebElement)
  Const JS_SCROLL_CENTER = _
    "this.scrollIntoView(true);" & _
    "var y = (window.innerHeight - this.offsetHeight) / 2;" & _
    "if (y < 1) return;" & _
    "for (var e=this; e; e=e.parentElement) {" & _
    "  if (e.scrollTop == 0) continue;" & _
    "  var yy = Math.min(e.scrollTop, y);" & _
    "  e.scrollTop -= yy;" & _
    "  if ((y -= yy) < 1) return;" & _
    "}" & _
    "window.scrollBy(0, -y);"
  
  element.ExecuteScript JS_SCROLL_CENTER
End Function

