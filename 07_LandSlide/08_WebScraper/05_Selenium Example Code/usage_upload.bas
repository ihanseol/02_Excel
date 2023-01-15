Attribute VB_Name = "usage_upload"

Private Sub Upload_File_FF()
  Dim file As String
  file = ThisWorkbook.Path & "\mozilla_privacypolicy.pdf"
  
  Dim driver As New FirefoxDriver
  driver.Get "http://the-internet.herokuapp.com/upload"
  
  'Upload the file
  driver.FindElementById("file-upload").SendKeys(file).Submit
  
  'Stop the browser
  driver.Quit
End Sub


Private Sub Upload_File_IE()
  Dim file As String
  file = ThisWorkbook.Path & "\mozilla_privacypolicy.pdf"
  
  Dim driver As New IEDriver
  driver.Get "http://the-internet.herokuapp.com/upload"
  
  'Upload the file
  driver.FindElementById("file-upload").SendKeys(file).Submit
  
  'Stop the browser
  driver.Quit
End Sub
