Attribute VB_Name = "usage_pdf"
Private Assert As New Assert


Private Sub Handle_PDF_Chrome()
  Dim driver As New ChromeDriver
  driver.Get "http://static.mozilla.com/moco/en-US/pdf/mozilla_privacypolicy.pdf"
  
  ' Return the first line using the pugin API (asynchronous).
  Const JS_READ_PDF_FIRST_LINE_CHROME As String = _
    "addEventListener('message',function(e){" & _
    " if(e.data.type=='getSelectedTextReply'){" & _
    "  var txt=e.data.selectedText;" & _
    "  callback(txt && txt.match(/^.+$/m)[0]);" & _
    " }" & _
    "});" & _
    "plugin.postMessage({type:'initialize'},'*');" & _
    "plugin.postMessage({type:'selectAll'},'*');" & _
    "plugin.postMessage({type:'getSelectedText'},'*');"
    
  ' Assert the first line
  Dim firstline
  firstline = driver.ExecuteAsyncScript(JS_READ_PDF_FIRST_LINE_CHROME)
  Assert.Equals "Websites Privacy Policy", firstline
  
  driver.Quit
End Sub


Private Sub Handle_PDF_FF()
  
  Const JS_WAIT_PDF_RENDERED_FIREFOX As String = _
    "(function fn(){" & _
    "  var pdf=PDFViewerApplication.pdfViewer;" & _
    "  if(pdf && pdf.onePageRendered){ " & _
    "    pdf.onePageRendered.then(callback);" & _
    "  }else{" & _
    "    window.setTimeout(fn,60);" & _
    "  }" & _
    "})();"
  
  Dim driver As New FirefoxDriver, Assert As New Assert
  driver.Get "http://static.mozilla.com/moco/en-US/pdf/mozilla_privacypolicy.pdf"
  
  ' Wait for the first page to be rendered
  driver.ExecuteAsyncScript JS_WAIT_PDF_RENDERED_FIREFOX
  
  ' Get the text from the first line
  firstline = driver.FindElementByCss("#pageContainer1 > .textLayer > div:nth-child(1)").Text
  
  ' Assert the first line
  Assert.Equals "Websites Privacy Policy", firstline
  
  driver.Quit
End Sub


Private Sub Handle_PDF_FF_advanced()
  
  ' Javascript to read the PDF with the pugin API (asynchronous).
  Const JS_READ_PDF_FIREFOX As String = _
    "(function fn(){" & _
    "  var pdf=PDFViewerApplication.pdfViewer;" & _
    "  if(pdf && pdf.onePageRendered){ " & _
    "    pdf.onePageRendered.then(function(){" & _
    "      var pdftxt=[],cnt=pdf.pagesCount,cpt=cnt;" & _
    "      for(var p=1; p<=cnt; p++){" & _
    "        pdf.pdfDocument.getPage(p).then(function(page){" & _
    "          page.getTextContent().then(function(content){" & _
    "            var lines=[], items=content.items;" & _
    "            for(var i=0, len=items.length; i<len; i++)" & _
    "              lines.push(items[i].str);" & _
    "            pdftxt[page.pageIndex]=lines.join();" & _
    "            if(--cpt < 1) callback(pdftxt);" & _
    "          });" & _
    "        });" & _
    "      }" & _
    "    });" & _
    "  }else{" & _
    "    setTimeout(fn, 60);" & _
    "  }" & _
    "})();"
  
  Dim driver As New FirefoxDriver
  driver.Get "http://static.mozilla.com/moco/en-US/pdf/mozilla_privacypolicy.pdf"
  
  Set pdfPages = driver.ExecuteAsyncScript(JS_READ_PDF_FIREFOX)
  Debug.Print pdfPages(1)  ' Print first page
  
  driver.Quit
End Sub



