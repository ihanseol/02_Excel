Attribute VB_Name = "m_DownloadFiles"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As LongPtr, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As LongPtr, _
        ByVal lpfnCB As LongPtr) As LongPtr
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As Long, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As Long, _
        ByVal lpfnCB As Long) As Long
#End If

Sub DownloadFile(FileURL As String, DestinationFolder As String)

    Dim DestinationFile As String
    
    If Dir(DestinationFolder, vbDirectory) = "" Then
        Debug.Print "Destination folder does not exist"
        Exit Sub
    End If
    
    DestinationFile = Mid(FileURL, InStrRev(FileURL, "/") + 1)
    DestinationFile = Replace(DestinationFile, "%20", " ")
    DestinationFile = DestinationFolder & "\" & DestinationFile
    
    If URLDownloadToFile(0, FileURL, DestinationFile, 0, 0) <> 0 Then
        Debug.Print "File download not started"
    End If
    
End Sub















