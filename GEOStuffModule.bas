Attribute VB_Name = "GEOStuffModule"
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Const IF_NO_CACHE_WRITE = &H4000000
Private Const BUFFER_LEN = 8192

Public Function GetUrlSource(sURL As String) As String
 Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
 Dim hInternet As Long, hSession As Long, lReturn As Long
'=================================================
'get the handle of the current internet connection
'=================================================
 hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
'=========================
'get the handle of the url
'=========================
 If hSession Then
  hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
 End If
'=================
'READ WEB RESPONSE
'=================
 If hInternet Then
 '=====================================
 'GET FIRST CHUNK OF DATA AND BUFFER IT
 '=====================================
  iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
  sData = sBuffer
 '==============================================
 'LOOP THROUGH THE BALANCE AND ADD TO THE BUFFER
 '==============================================
  Do While lReturn <> 0
   iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
   sData = sData + Mid(sBuffer, 1, lReturn)
  Loop
 End If
'==========================================
'CLOSE THE CONNECTION AND RETURN THE RESULT
'==========================================
 iResult = InternetCloseHandle(hInternet)
 GetUrlSource = sData
End Function
