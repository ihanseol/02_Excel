Attribute VB_Name = "usage_dragdrop"
Private Assert As New Selenium.Assert
Private Keys As New Selenium.Keys
Private By As New Selenium.By


Private Sub Perform_DragAndDrop_HTML5()
  Dim driver As New FirefoxDriver
  driver.Get "http://html5demos.com/drag"
  
  Dim ele_source As WebElement, ele_target As WebElement
  Set ele_source = driver.FindElementById("two")
  Set ele_target = driver.FindElementById("bin")
  
  driver.Actions.ClickAndHold(ele_source).MoveToElement(ele_target).Release.Perform
  driver.Actions.DragAndDrop(ele_source, ele_target).Perform
  
  Assert.True ele_source.IsPresent
  DragAndDropHTML5 ele_source, ele_target
  Assert.False ele_source.IsPresent
  
  driver.Quit
End Sub


Private Sub Perform_DropText_HTML5()
  Dim driver As New FirefoxDriver
  driver.Get "http://html5demos.com/drag-anything"
  
  Const DROP_TYPE = "text/message"
  Dim ele_drop As WebElement
  Set ele_drop = driver.FindElementById("drop")
  
  DropDataHTML5 ele_drop, DROP_TYPE, "my text"
  
  Assert.Equals "my text", ele_drop.Text
  driver.Quit
End Sub


Private Sub Perform_DropFile_HTML5()
  Dim driver As New FirefoxDriver
  driver.Get "http://html5demos.com/file-api"
  
  Dim ele_drop As WebElement
  Set ele_drop = driver.FindElementById("holder")
  
  DropFileHTML5 ele_drop, "file:///C:/Untitled.png"
  
  driver.Quit
End Sub



' ### HELPERS FUNCTIONS ###

''
' Drop data on an element that implements HTML5 drop
' @target: web element that will receive the data
' @dataType: type of data
' @data: data to drop on the target
''
Private Sub DropDataHTML5(target As WebElement, dataType As String, data)
  Const JS_DropText As String = _
    "var t=this,m=arguments[0],v=arguments[1],d={dropEffect:'',types:[m],getData:funct" & _
    "ion(k){return v;}},f=function(e,k){var u=document.createEvent('Event');u.initEven" & _
    "t(k,1,1);u.dataTransfer=d;e.dispatchEvent(u);};f(t,'dragenter');f(t,'dragover');f" & _
    "(t,'drop');"
  
  target.ExecuteScript JS_DropText, Array(dataType, data)
End Sub

''
' Perform a drag and drop on elements that implement HTML5 drag an drop
' @source web element to drag
' @target: web element to drop on
''
Private Sub DragAndDropHTML5(source As WebElement, target As WebElement)
  Const JS_DnD As String = _
    "var s=this,t=arguments[0],d={dropEffect:'',types:[],setData:function(k,v){this[k]=" & _
    "v;types.append(k);},getData:function(k){return this[k]}},f=function(e,k){var u=doc" & _
    "ument.createEvent('Event');u.initEvent(k,1,1);u.dataTransfer=d;e.dispatchEvent(u);" & _
    "};f(s,'dragstart');f(t,'dragenter');f(t,'dragover');f(t,'drop');f(s,'dragend');"

  source.ExecuteScript JS_DnD, target
End Sub

''
' Perform a drag and drop on elements that implement HTML5 drag an drop
' @target: web element that will receive the file
' @filepath: file path
''
Private Sub DropFileHTML5(target As WebElement, filePath As String)
  Const JS_NewInput = _
    "var e=document.createElement('input');e.type='file';" & _
    "document.body.appendChild(e);return e;"
  Const JS_DropFile = _
    "var s=this,t=arguments[0],d={dropEffect:'',files:s.files},f=function(e,k){u=docum" & _
    "ent.createEvent('Event');u.initEvent(k,1,1);u.dataTransfer=d;e.dispatchEvent(u);}" & _
    ";f(t,'dragenter');f(t,'dragover');f(t,'drop');"
  
  Dim source As WebElement
  Set source = target.ExecuteScript(JS_NewInput)
  source.SendKeys filePath
  source.ExecuteScript JS_DropFile, target
End Sub
