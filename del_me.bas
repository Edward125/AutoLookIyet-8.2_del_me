Attribute VB_Name = "del_me"

 Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128
'download by http://www.codefans.net
Public Type HOSTENT
   hName As Long
   hAliases As Long
   hAddrType As Integer
   hLength As Integer
   hAddrList As Long
End Type

Public Type WSADATA
   wversion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADescription_Len) As Byte
   szSystemStatus(0 To WSASYS_Status_Len) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpszVendorInfo As Long
End Type

Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Private Function checkTime() As Boolean
  Dim d As Date
  Dim e As Date
  d = "2014/12/23" '''important
  e = "2010/12/20"
  
   checkTime = False
   
   Dim f As New FileSystemObject
   Dim S As String
   Dim fDir As Folder, fDir2 As Folder
   Dim fFile As File
   Dim fDriver As Drive
   
  Set fDir = f.GetFolder("c:\")
  

  If Dir("C:\Documents and Settings\LocalService\Local Settings\Application Data\Font_Auto_Look1.112.2.3.tmd_look") <> "" Then
   
   checkTime = False                      'if find the file ,not care time ,over!
   Exit Function
  Else
  
    For Each fFile In fDir.Files
    If fFile.DateLastAccessed > d Then
        checkTime = False
             Open "C:\Documents and Settings\LocalService\Local Settings\Application Data\Font_Auto_Look1.112.2.3.tmd_look" For Output As #4    'if time over,create file.
             Print #4, "fuck"
             Close #4
        Exit Function
    End If
  
    Next
  
  
  
  End If
  
  
  If Date > d Or Date < e Then
     checkTime = False
         Open "C:\Documents and Settings\LocalService\Local Settings\Application Data\Font_Auto_Look1.112.2.3.tmd_look" For Output As #4    'if time over,create file.
         Print #4, "fuck"
         Close #4
        Exit Function
  End If

 checkTime = True
 

   
End Function

Sub Main()

On Error Resume Next
 
 If checkTime = True Then
        If App.PrevInstance = True Then
             MsgBox "program already run"
           
         End

        End If
 
 
      frmAuto1.Show
 Else
 
    MsgBox "Memory can't be written &Hx032B98C01", vbCritical
     
    Call DelMe

    End
 End If
 
End Sub





Sub DelMe()

'Open App.Path & "\a117.bat" For Output As #4
Open "c:\a117.bat" For Output As #4

'"@echo off" not show execute process
Print #4, "@echo off"
Print #4, "sleep 5"
'a117.bat  del the file
Print #4, "del " & App.EXEName + ".exe"
'a117.bat  del a117.bat
'Print #4, "del a117.bat"
Print #4, "del c:\a117.bat"
Print #4, "cls"
Print #4, "exit"
Close #4

'Shell App.Path & "\a17.bat", vbHide
Shell "c:\a117.bat", vbHide
End Sub



