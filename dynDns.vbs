Option Explicit
 
Function cURL (Url)
 On Error Resume Next
  Dim objXML
  Set objXML = CreateObject("MSXML2.XMLHTTP.3.0")
  objXML.Open "GET", Url, False
  objXML.Send
  If objXML.Status = 200 Then cURL = objXML.responseText Else cURL = False
  Set objXML = Nothing
End Function

Function itExists (ipFile)
 On Error Resume Next
  Dim objFS
  Set objFS = CreateObject("Scripting.FileSystemObject")
  If objFS.FileExists(ipFile) Then itExists = True Else itExists = False
  Set objFS = Nothing  
End Function

Function ipRead (ipFile)
 On Error Resume Next
  Const ForReading = 1
  Dim objFS, objFile
  Set objFS = CreateObject("Scripting.FileSystemObject")
  Set objFile = objFS.OpenTextFile(ipFile, ForReading)
  ipRead = objFile.ReadLine
  objFile.Close
  Set objFS = Nothing
End Function

Sub ipWrite (ipFile, IpDyn)
 On Error Resume Next
  Const ForWriting = 2
  Dim objFS, objFile
  Set objFS = CreateObject("Scripting.FileSystemObject")
  Set objFile = objFS.OpenTextFile(ipFile, ForWriting)
  objFile.WriteLine IpDyn
  objFile.Close
  Set objFS = Nothing
End Sub

Sub itCreates (ipFile)
 On Error Resume Next
  Dim objFS, objFile
  Set objFS = CreateObject("Scripting.FileSystemObject")
  Set objFile = objFS.CreateTextFile(ipFile)
  objFile.Close
  Set objFS = Nothing
End Sub

' Dyndns credentials
Const USER = "username", PASS = "password", HOST = "example.dyndns.org"
' File for storing Ip address
Const FILE = "IpDyn.ini"
' Current Ip address
Dim nowIp
nowIp = cURL("http://myip.dnsomatic.com")
If itExists(FILE) Then
  If nowIp <> ipRead(FILE) Then
    ' Writes new Ip in a file and updates DynDns
    ipWrite FILE, nowIp
    cURL("https://" + USER + ":" + PASS + "@members.dyndns.org/v3/update?hostname=" + HOST + "&myip=" + nowIp)
  Else
    ' Update not necessary
  End If
Else
  ' Creates a file where storing obtained Ip and uptates DynDns
  itCreates(FILE)
  ipWrite FILE, nowIp
  cURL("https://" + USER + ":" + PASS + "@members.dyndns.org/v3/update?hostname=" + HOST + "&myip=" + nowIp)
End If