Attribute VB_Name = "usage_download"

Private Declare PtrSafe Function FindWindowExA Lib "user32.dll" ( _
  ByVal hwndParent As LongPtr, _
  ByVal hwndChildAfter As LongPtr, _
  ByVal lpszClass As String, _
  ByVal lpszWindow As String) As Long

Private Declare PtrSafe Function PostMessageA Lib "user32.dll" ( _
  ByVal hwnd As LongPtr, _
  ByVal wMsg As LongPtr, _
  ByVal wParam As LongPtr, _
  ByVal lParama As LongPtr) As Long

Private Declare PtrSafe Function GetWindowLongA Lib "user32.dll" ( _
  ByVal hwnd As LongPtr, ByVal nIndex As Integer) As Long

Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc.dll" ( _
  ByVal hwnd As LongPtr, _
  ByVal dwId As Long, _
  ByRef riid As Any, _
  ByRef ppvObject As IAccessible) As Long
                                       
Private Declare PtrSafe Function AccessibleChildren Lib "oleacc.dll" ( _
  ByVal paccContainer As IAccessible, _
  ByVal iChildStart As Long, _
  ByVal cChildren As Long, _
  ByRef rgvarChildren As Variant, _
  ByRef pcObtained As Long) As Long




''
' Downloads the link defined in the href attribute of a web element
''
Private Sub Usage_Download_StaticLink()
  Dim driver As New IEDriver, ele As WebElement
  driver.Get "https://www.mozilla.org/en-US/foundation/documents"
  
  Set ele = driver.FindElementByLinkText("IRS Form 872-C")
  Download_StaticLink ele, ThisWorkbook.Path & "\irs-form-872-c_1.pdf"
  
  driver.Quit
End Sub


''
' Downloads a file with IE and waits for completion
''
Private Sub Download_File_IE()
  Dim driver As New IEDriver, ele As WebElement
  driver.Get "https://www.mozilla.org/en-US/foundation/documents"
  
  Dim filePath As String
  driver.FindElementByLinkText("IRS Form 872-C").ExecuteScript "this.click()"
  filePath = DownloadFileSyncIE(ThisWorkbook.Path)
  
  driver.Quit
End Sub


''
' Downloads a file with IE without waiting for completion
''
Private Sub Download_File_Asynchrone_IE()
  Dim driver As New IEDriver, ele As WebElement
  driver.Get "https://www.mozilla.org/en-US/foundation/documents"
  
  ' Init the file waiter
  WaitNewFile ThisWorkbook.Path & "\*.pdf"
  
  driver.FindElementByLinkText("IRS Form 872-C").ExecuteScript "this.click()"
  DownloadFileAsyncIE ThisWorkbook.Path
  
  ' Waits for a new file
  file = WaitNewFile()
  
  Debug.Assert 0
  driver.Quit
End Sub

''
' Sets the download folder with Firefox
''
Private Sub Download_File_Firefox()
  Dim driver As New FirefoxDriver, file As String
  
  'Set the preferences specific to Firefox
  driver.SetPreference "browser.download.folderList", 2
  driver.SetPreference "browser.download.dir", ThisWorkbook.Path
  driver.SetPreference "browser.helperApps.neverAsk.saveToDisk", "application/pdf"
  driver.SetPreference "pdfjs.disabled", True
  
  ' Init the file waiter
  WaitNewFile ThisWorkbook.Path & "\*.pdf"
  
  ' Open the file for download
  driver.Get "https://www.mozilla.org/en-US/foundation/documents"
  driver.FindElementByLinkText("IRS Form 872-C").Click
  
  ' Waits for a new file
  file = WaitNewFile()
  
  'Stop the browser
  driver.Quit
End Sub

''
' Sets the download folder with Chrome
''
Private Sub Download_File_Chrome()
  Dim driver As New ChromeDriver, file As String
  
  'Set the preferences specific to Chrome
  driver.SetPreference "download.default_directory", ThisWorkbook.Path
  driver.SetPreference "download.directory_upgrade", True
  driver.SetPreference "download.prompt_for_download", False
  driver.SetPreference "plugins.plugins_disabled", Array("Chrome PDF Viewer")
  
  ' Init the file waiter
  WaitNewFile ThisWorkbook.Path & "\*.pdf"
  
  'Open the file for download
  driver.Get "https://www.mozilla.org/en-US/foundation/documents"
  driver.FindElementByLinkText("IRS Form 872-C").Click
  
  ' Waits for a new file
  file = WaitNewFile()
  
  'Stop the browser
  Debug.Assert 0
  driver.Quit
End Sub





' ### HELPERS FUNCTIONS ###

''
' Saves the file pointed by the href attribute : <a href="/doc.pdf">Document</a>
' @element {WebElement} Web element with the href link
' @save_as {String}  Path were the file is to be saved
''
Private Sub Download_StaticLink(element As WebElement, save_as As String)
  ' Extract the data to build the request (link, user-agent, language, cookie)
  Dim info As Selenium.Dictionary
  Set info = element.ExecuteScript("return {" & _
    "link: this.href," & _
    "agent: navigator.userAgent," & _
    "lang: navigator.userLanguage," & _
    "cookie: document.cookie };")
  
  ' Send the request
  Static xhr As Object
  If xhr Is Nothing Then Set xhr = CreateObject("Msxml2.ServerXMLHTTP.6.0")
  xhr.Open "GET", info("link")
  xhr.setRequestHeader "User-Agent", info("agent")
  xhr.setRequestHeader "Accept-Language", info("lang")
  xhr.setRequestHeader "Cookie", info("cookie")
  xhr.Send
  If (xhr.Status \ 100) - 2 Then Err.Raise 5, , xhr.Status & " " & xhr.StatusText
  
  ' Save the response to a file
  Static bin As Object
  If bin Is Nothing Then Set bin = CreateObject("ADODB.Stream")
  If Len(Dir$(save_as)) Then Kill save_as
  bin.Open
  bin.Type = 1
  bin.Write xhr.ResponseBody
  bin.Position = 0
  bin.SaveToFile save_as
  bin.Close
End Sub


''
' Waits for a new file to be created in a folder
' @folder {String}  Folder where the file will be created
' Usage:
'   WaitNewFile "C:\download\*.pdf"
'   ' The new file is created here
'   filename = WaitNewFile()
''
Public Function WaitNewFile(Optional target As String) As String
  Static files As Collection, filter$
  Dim file$, file_path$, i&
  If Len(target) Then
    ' Initialize the list of files and return
    filter = target
    Set files = New Collection
    file = Dir(filter, vbNormal)
    Do While Len(file)
      files.Add Empty, file
      file = Dir
    Loop
    Exit Function
  End If
  
  ' Waits for a file that is not in the list
  On Error GoTo WaitReady
  Do
    file = Dir(filter, vbNormal)
    Do While Len(file)
      files.Item file
      file = Dir
    Loop
    For i = 0 To 3000: DoEvents: Next
  Loop
  
WaitReady:
  ' Waits for the size to be superior to 0 and try to rename it
  file_path = Left$(filter, InStrRev(filter, "\")) & file
  Do
    If FileLen(file_path) Then
      On Error Resume Next
      Name file_path As file_path
      If Err = 0 Then Exit Do
    End If
    For i = 0 To 3000: DoEvents: Next
  Loop
  files.Add Empty, file
  WaitNewFile = file_path
End Function


''
' Saves the file from the download dialogue, waits for completion and returns the path
' @save_as: folder or file path
''
Private Function DownloadFileSyncIE(ByVal save_as As String) As String
  Const dl_key = "HKCU\Software\Microsoft\Internet Explorer\Main\Default Download Directory"
  
  Static shl As Object, Waiter As New Waiter
  If shl Is Nothing Then
    Set shl = CreateObject("WScript.Shell")
    shl.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Main\NotifyDownloadComplete", "no", "REG_SZ"
    shl.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_FILEDOWNLOAD\iexplore.exe", 0, "REG_DWORD"
  End If
  
  Dim ie_hwnd, frm_hwnd, endtime#, i&, n&, folder_bak$, file_name$
  
  ' wait for the download dialogue (IEFrame/Frame Notification Bar/DirectUIHWND)
  ie_hwnd = FindWindowExA(0, 0, "IEFrame", vbNullString)
  endtime = Now + 5000 / 86400#
  Do
    frm_hwnd = FindWindowExA(ie_hwnd, 0, "Frame Notification Bar", vbNullString)
    If frm_hwnd Then
      If GetWindowLongA(frm_hwnd, -16) And &H10000000 Then Exit Do ' If visible
    End If
    If Now > endtime Then Err.Raise 5, , "Failed to find the download dialog"
    Waiter.Wait 100
  Loop
  
  ' save the download folder path and create a temporary folder
  tmp_dir = Environ$("TEMP") & "\dl-ie-4f521"
  On Error Resume Next
  folder_bak = shl.RegRead(dl_key)
  MkDir tmp_dir
  Kill tmp_dir & "\*"
  On Error GoTo 0
  
  ' set the download folder in the registry
  shl.RegWrite dl_key, tmp_dir, "REG_SZ"
  
  ' send the shortcut for Save (Alt + S)
  Waiter.Wait 500
  PostMessageA ie_hwnd, &H104&, &H12, &H20000001  'WM_SYSKEYDOWN, VK_MENU
  PostMessageA ie_hwnd, &H104&, &H53, &H20000001  'WM_SYSKEYDOWN, S
  PostMessageA ie_hwnd, &H105&, &H53, &HC0000001  'WM_SYSKEYUP, S
  PostMessageA ie_hwnd, &H101&, &H12, &HC0000001  'WM_KEYUP, VK_MENU
  
  ' wait for the file to be downloaded
  Do
    Waiter.Wait 100
    file_name = VBA.Dir$(tmp_dir & "\*")
  Loop While InStr(Len(file_name) - 8, file_name, ".partial") Or Len(file_name) = 0
  
  ' restore the download folder in the registry
  If folder_bak = Empty Then
    shl.RegDelete dl_key
  Else
    shl.RegWrite dl_key, folder_bak, "REG_SZ"
  End If
  
  ' delete existing file
  If Len(VBA.Dir$(save_as, vbNormal)) Then Kill save_as
  If Len(VBA.Dir$(save_as, vbDirectory)) Then
    save_as = save_as & "\" & file_name
    If Len(VBA.Dir$(save_as, vbNormal)) Then Kill save_as
  End If
  
  ' move the file to the provided path
  Name tmp_dir & "\" & file_name As save_as
  DownloadFileSyncIE = save_as
End Function

''
' Saves the file from the download dialogue without waiting for completion
' @folder: download folder
''
Private Sub DownloadFileAsyncIE(folder As String)
  Const timeout = 5000, bt_save = "Save", bt_close = "Close"
  Const dl_key = "HKCU\Software\Microsoft\Internet Explorer\Main\Default Download Directory"
  
  Static shl As Object, Waiter As New Waiter
  If shl Is Nothing Then
    Set shl = CreateObject("WScript.Shell")
    shl.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Main\NotifyDownloadComplete", "no", "REG_SZ"
    shl.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_FILEDOWNLOAD\iexplore.exe", 0, "REG_DWORD"
  End If
  
  ' wait for the download dialog (IEFrame/Frame Notification Bar/DirectUIHWND)
  Dim ie_hwnd, frm_hwnd, dlg_hwnd, endtime#, folder_bak$, i&
  ie_hwnd = FindWindowExA(0, 0, "IEFrame", vbNullString)
  endtime = Now + timeout / 86400#
  Do
    frm_hwnd = FindWindowExA(ie_hwnd, 0, "Frame Notification Bar", vbNullString)
    If frm_hwnd Then
      If GetWindowLongA(frm_hwnd, -16) And &H10000000 Then Exit Do  ' If visible
    End If
    If Now > endtime Then Err.Raise 5, , "Failed to find the download dialog"
    For i = 1 To 5000: DoEvents: Next
  Loop
  
  ' get the save button
  Dim acc As IAccessible, bt As IAccessible
  dlg_hwnd = FindWindowExA(frm_hwnd, 0, "DirectUIHWND", vbNullString)
  Set acc = acc_from_window(dlg_hwnd)
  Set bt = acc_find_button(acc, bt_save)
  If bt Is Nothing Then Err.Raise 5, , "Failed to find the Save button"
  
  ' save and set the download folder in the registry
  On Error Resume Next
  folder_bak = shl.RegRead(dl_key)
  On Error GoTo 0
  shl.RegWrite dl_key, folder, "REG_SZ"
  
  ' click on Save
  Waiter.Wait 500
  bt.accDoDefaultAction 0&
  Waiter.Wait 100
  
  ' restore the download folder in the registry
  If folder_bak = Empty Then
    shl.RegDelete dl_key
  Else
    shl.RegWrite dl_key, folder_bak, "REG_SZ"
  End If
End Sub

  
Private Function acc_from_window(hwnd) As IAccessible
  Dim iid&(0 To 3)
  iid(0) = &H618736E0 ' IAccessible interface
  iid(1) = &H11CF3C3D
  iid(2) = &HAA000C81
  iid(3) = &H719B3800
  AccessibleObjectFromWindow hwnd, 0&, iid(0), acc_from_window
End Function

Private Function acc_find_button(ByVal acc As IAccessible, name$) As IAccessible
  If acc.accName(0&) Like name Then
    Set acc_find_button = acc
  ElseIf acc.accChildCount > 0 Then
    Dim children(0 To 20), count&, i&
    AccessibleChildren acc, 0, acc.accChildCount, children(0), count
    For i = 0 To count - 1
      If VBA.IsObject(children(i)) Then
        Set acc_find_button = acc_find_button(children(i), name)
        If Not acc_find_button Is Nothing Then Exit For
      End If
    Next
  End If
End Function

