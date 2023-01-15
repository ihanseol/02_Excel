Attribute VB_Name = "utils"
'
' Function to directly call a .Net Dll without registration
'
' Edit the NETDLL function to make it return the Selenium.dll library
'
' Include these references if Office 32bits:
' c:\Windows\Microsoft.NET\Framework\v2.0.50727\mscoree.tlb
' c:\Windows\Microsoft.NET\Framework\v2.0.50727\mscorlib.tlb
'
' Include these references if Office 64bits:
' C:\Windows\Microsoft.NET\Framework64\v2.0.50727\mscoree.tlb
' C:\Windows\Microsoft.NET\Framework64\v2.0.50727\mscorlib.tlb
'

Private Function NETDLL() As String
    NETDLL = Environ("USERPROFILE") & "\AppData\Local\SeleniumBasic\Selenium.dll"
End Function

Function CreateObject2(typeName As String) As Object
    Static domain As mscorlib.AppDomain
    If domain Is Nothing Then
        Dim host As New mscoree.CorRuntimeHost
        host.Start
        host.GetDefaultDomain domain
    End If
    Set CreateObject2 = domain.CreateInstanceFrom(NETDLL, typeName).Unwrap
End Function

