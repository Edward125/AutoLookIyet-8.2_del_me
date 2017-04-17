Attribute VB_Name = "Module1"
Public G As Integer
Public ViewText As String
Public LenViewText As Integer
Public ViewText2 As String
Public LenViewText2 As Integer
Public OpenView As Boolean
Public AnalogView As Boolean


Public BoardViewTrue As Boolean
Public BoardViewDevice As String

Public strTmpDVS(1 To 3) As String
Public strToolPath  As String
Public ListFind As Boolean
Public strReportPath  As String
Public strNetReportPath  As String
Public strBoardPath  As String
Public strNetBoardPath  As String
Public strNetName  As String
Public strName  As String
Public bFrmClose  As Boolean

Public strLoadFrom1Path  As String
Public strLoadFrom2Path  As String
Public strViewPath  As String
Public strUploadPath  As String
Public strFindOption As String

Public strNodeS  As String

Public strNodeI  As String
Public ListFindS As Boolean
Public ListFindI As Boolean
Public AnalogDeviceName As String
Public AnalogDeviceNameNet As String
Public FileTypeBBB As String
Public ReTestTimesOut As Integer
Public COM2_1 As String
 

Public Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpFn             As Long
   lParam           As Long
   iImage           As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const MAX_PATH = 260

Public Declare Function SHGetPathFromIDList _
   Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
   Alias "SHBrowseForFolderA" _
  (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Sub CoTaskMemFree Lib "ole32.dll" _
   (ByVal pv As Long)




