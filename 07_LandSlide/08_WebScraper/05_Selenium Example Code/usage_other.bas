Attribute VB_Name = "usage_other"

Private Const JS_GET_TIMINGS As String = _
  "var t=performance.timing; return [  " & _
  " t.responseEnd - t.navigationStart, " & _
  " t.loadEventEnd - t.responseEnd ];  "

Private Const JS_BUILD_CSS = _
  "var e=this,p=[],h=function(a,s){if(!a||!s)return 0;for(var i=0,l=a.length;i<l;i" & _
  "++){if(s.indexOf(a[i])==-1)return 0};return 1;};for(;e&&e.nodeType==1&&e.nodeNa" & _
  "me!='HTML';e=e.parentNode){if(e.id){p.unshift('#'+e.id);break;}var i=1,u=1,t=e." & _
  "localName,c=e.className&&e.className.split(/[\s,]+/g);for(var s=e.previousSibli" & _
  "ng;s;s=s.previousSibling){if(s.nodeType!=10&&s.nodeName==e.nodeName){if(h(c,s.c" & _
  "lassName))c=null;u=0;++i;}}for(var s=e.nextSibling;s;s=s.nextSibling){if(s.node" & _
  "Name==e.nodeName){if(h(c,s.className))c=null;u=0;}}p.unshift(u?t:(t+(c?'.'+c.jo" & _
  "in('.'):':nth-child('+i+')')));}return p.join(' > ');"

Private Const JS_BUILD_XPATH = _
  "var e=this,p=[];for(;e&&e.nodeType==1&&e.nodeName!='HTML';e=e.parentNode){if(e." & _
  "id){p.unshift('*[@id=\''+e.id+'\']');break;}var i=1,u=1,t=e.localName,c=e.class" & _
  "Name;for(var s=e.previousSibling;s;s=s.previousSibling){if(s.nodeType!=10&&s.no" & _
  "deName==e.nodeName){if(c==s.className)c=null;u=0;++i;}}for(var s=e.nextSibling;" & _
  "s;s=s.nextSibling){if(s.nodeName==e.nodeName){if(c==s.className)c=null;u=0;}}p." & _
  "unshift(u?t:(t+(c?'[@class=\''+c+'\']':'['+i+']')));}return '//'+p.join('/');"

Private Const JS_LIST_ATTRIBUTES As String = _
  "var d={}, a=this.attributes;  " & _
  "for(var i=0; i<a.length; i++) " & _
  "  d[a[i].name]=a[i].value;    " & _
  "return d;"


Private Sub Get_Performance_Timing()
  'https://developer.mozilla.org/en/docs/Web/API/Navigation_timing_API
  'http://www.html5rocks.com/en/tutorials/webperformance/basics
  'http://www.w3.org/TR/navigation-timing
  
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  times = driver.ExecuteScript(JS_GET_TIMINGS).values
  Debug.Print "Timing:"
  For Each T In times
    Debug.Print T
  Next
  
  driver.Quit
End Sub


Private Sub Build_XPath_Locators()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"

  Set elements = driver.FindElementsByCss("*")
  Set list_xpath = elements.ExecuteScript(JS_BUILD_XPATH)
  For Each txt In list_xpath
    Debug.Print txt
  Next

  driver.Quit
End Sub


Private Sub Build_CSS_Locators()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"

  Set elements = driver.FindElementsByCss("*")
  Set list_css = elements.ExecuteScript(JS_BUILD_CSS)
  For Each txt In list_css
    Debug.Print txt
  Next

  driver.Quit
End Sub

Private Sub Build_CSS_Locators2()
  Dim driver As New FirefoxDriver
  driver.Get "http://form.timeform.betfair.com/daypage?date=20150816"
  
  Dim css As String
  css = driver.FindElementByLinkText("Southwell") _
              .ExecuteScript(JS_BUILD_CSS)
  Debug.Print css

  driver.Quit
End Sub

Private Sub List_Element_Attributes()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  Set element = driver.FindElementById("mw-content-text")
  Set atts = element.ExecuteScript(JS_LIST_ATTRIBUTES)
  For Each att In atts
    Debug.Print att.key & "  " & att.Value
  Next
  
  driver.Quit
End Sub



