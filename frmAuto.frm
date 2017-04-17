VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAuto1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8400
   ClientLeft      =   9600
   ClientTop       =   1005
   ClientWidth     =   5385
   Icon            =   "frmAuto.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   5385
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   4080
      Top             =   2520
   End
   Begin VB.CommandButton cmdTop5 
      Caption         =   "ShowTop5"
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdSN 
      Caption         =   "SN"
      Height          =   255
      Left            =   4440
      TabIndex        =   31
      Top             =   5040
      Width           =   375
   End
   Begin VB.TextBox txtBasicCmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFFFC0&
      Caption         =   "com'VerTestjet"
      Height          =   255
      Left            =   1920
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "com'Testjet"
      Height          =   255
      Left            =   3360
      TabIndex        =   28
      Top             =   7800
      Width           =   1215
   End
   Begin VB.PictureBox PicCaption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "frmAuto.frx":0442
      ScaleHeight     =   720
      ScaleWidth      =   9600
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   9600
      Begin VB.PictureBox PicBorder 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   1200
         Picture         =   "frmAuto.frx":16C86
         ScaleHeight     =   150
         ScaleWidth      =   1050
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Clear"
      Height          =   255
      Left            =   4080
      TabIndex        =   25
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox l 
      Height          =   285
      Left            =   720
      TabIndex        =   24
      Text            =   """"
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Enabled         =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   7440
      Width           =   3735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFC0C0&
      Height          =   1230
      ItemData        =   "frmAuto.frx":17398
      Left            =   1320
      List            =   "frmAuto.frx":1739A
      TabIndex        =   18
      Top             =   6000
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5400
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SetPath"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   7800
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "NetWork"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DebugAnalog"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LockText"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5760
      Width           =   975
   End
   Begin VB.OptionButton otestjet 
      Caption         =   "Testjet"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   5040
      Width           =   855
   End
   Begin VB.OptionButton oanalog 
      Caption         =   "Analog"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
   Begin VB.OptionButton oshorts 
      Caption         =   "Short"
      Enabled         =   0   'False
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   5040
      Width           =   735
   End
   Begin VB.OptionButton oOpen 
      Caption         =   "Open"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2040
      Top             =   2400
   End
   Begin VB.CommandButton Command7 
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox ListDVS 
      BackColor       =   &H00FFC0C0&
      Height          =   1230
      ItemData        =   "frmAuto.frx":1739C
      Left            =   120
      List            =   "frmAuto.frx":1739E
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txtNGLog 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   0
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UnLockText"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label OpenViewText 
      Height          =   255
      Left            =   2400
      TabIndex        =   30
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "FileName:"
      Height          =   255
      Left            =   2040
      TabIndex        =   23
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "BasicCommand"
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "DeviceName:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "IPName:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   615
   End
   Begin VB.Menu File_ 
      Caption         =   "File"
      Begin VB.Menu OpenIyetFile 
         Caption         =   "SetIyetFile"
      End
      Begin VB.Menu OpenNetIyetFile 
         Caption         =   "SetNetIyetFile"
         Visible         =   0   'False
      End
      Begin VB.Menu OpenTestjetFile 
         Caption         =   "OpenTestjetFile"
         Visible         =   0   'False
      End
      Begin VB.Menu ww 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu set_ 
      Caption         =   "Set.."
      Begin VB.Menu SetPath 
         Caption         =   "SetPath"
      End
      Begin VB.Menu uuu 
         Caption         =   "-"
      End
      Begin VB.Menu ClearText 
         Caption         =   "ClearText"
      End
      Begin VB.Menu LockText 
         Caption         =   "LockText"
      End
      Begin VB.Menu www 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu NetWork 
         Caption         =   "NetWork"
         Visible         =   0   'False
      End
      Begin VB.Menu wwww 
         Caption         =   "-"
      End
      Begin VB.Menu DebugAnalog 
         Caption         =   "DebugAnalog"
      End
      Begin VB.Menu DebugTestJet 
         Caption         =   "Com""TestJet"""
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu ScreenViewFailDevice 
         Caption         =   "ScreenViewFailDevice"
      End
      Begin VB.Menu sss 
         Caption         =   "-"
      End
      Begin VB.Menu NgReportListTestjetPin 
         Caption         =   "NgReportListTestjetPin"
      End
      Begin VB.Menu uussss 
         Caption         =   "-"
      End
      Begin VB.Menu ShortsReTestOneTimes 
         Caption         =   "ShortsReTest 1 Times"
      End
      Begin VB.Menu bar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu BoardView_ 
         Caption         =   "OpenBoardView"
         Visible         =   0   'False
      End
      Begin VB.Menu CreateBoardView 
         Caption         =   "CreateBoardView"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Window_ 
      Caption         =   "Window"
      Begin VB.Menu MaxWindow 
         Caption         =   "MaxWindow"
      End
      Begin VB.Menu CenterScreen 
         Caption         =   "CenterScreen"
      End
   End
End
Attribute VB_Name = "frmAuto1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 窗口置前=========
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'-------------------


Private m_cN As cNeoCaption


'




Dim strDeviceTestTimes(5)
Dim intTime As Integer





Dim NotLoadFile As Boolean
Dim TestjetPinName(25) As String
Dim ListPinCont As Integer
Dim uu As Integer
Dim uu2 As Integer
Dim ListView As Boolean
Dim ReportTestjet As Boolean
Dim ShortsReTest_1_Times As Boolean
Dim StopRun As Boolean
Dim sBoardVersion As String
Dim sBoardVersionBasic As String
Dim strSN As String

Private Sub BoardView__Click()
BoardViewTrue = Not BoardViewTrue
If BoardViewTrue = True Then
    frmBoardView.Show
    BoardView_.Caption = "BoardViewTrue"
    'CloseBoardView.Enabled = True
   ' Call CenterScreen_Click
  Else
   Unload frmBoardView
    BoardView_.Caption = "BoardViewFalse"
     'Call CenterScreen_Click
End If
End Sub

Private Sub CenterScreen_Click()
'Windows Default
  Dim myval
  Static bTrue As Boolean
  If bTrue = False Then
       '窗口正常

      myval = SetWindowPos(frmAuto1.hwnd, -1, 0, 0, 0, 0, 3)
       CenterScreen.Caption = "Windows Default"
       bTrue = True
    Else
      ' 窗口置前
      myval = SetWindowPos(frmAuto1.hwnd, -2, 0, 0, 0, 0, 3)
      CenterScreen.Caption = "CenterScreen"
      bTrue = False
  End If
End Sub

Private Sub ClearText_Click()
Call Command10_Click
End Sub

Private Sub cmdSN_Click()
If strSN <> "" Then
   Clipboard.SetText strSN
End If
End Sub

Private Sub cmdTop5_Click()
Static bTrue_ As Boolean
If cmdTop5.Caption = "ShowTop5" Then
   bTrue_ = False
End If
If bTrue_ = False Then
  frmTop_5.Show
  cmdTop5.Caption = "HideTop5"
  bTrue_ = True
  Else
  bTrue_ = False
  frmTop_5.Hide
  cmdTop5.Caption = "ShowTop5"
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If oanalog.Value = True Then
 AnalogDeviceName = ListDVS.Text
 Unload frmAnalogDebug
 frmAnalogDebug.Show
 Else
   Unload frmAnalogDebug
  frmAnalogDebug.Show
End If
End Sub

Private Sub Command10_Click()
     FindTestFile = False
     ListDVS.Clear
    ' NotLoadFile = True
     txtNGLog.Text = ""
     Text1.Text = ""
     Text2.Text = ""
     intTestjetPin = 0
     oanalog.Enabled = False
     otestjet.Enabled = False
     oshorts.Enabled = False
     oOpen.Enabled = False
     intTestjetPin = 0
     Command8.Caption = ""
     Command8.Enabled = False
     List1.Clear
    ' Text4.Text = ""
     oanalog.ForeColor = &H80000012
     oshorts.ForeColor = &H80000012
     oOpen.ForeColor = &H80000012
     otestjet.ForeColor = &H80000012
    ' NotLoadFile = False
End Sub

Private Sub Command11_Click()
 Dim DK1orDB2testjet As String, strTmpModetestjet As String
 Dim intGG As Integer
 Dim strTmpModeTME As String
 Dim strTmpDIRTMP As String
 
On Error GoTo EX
If strToolPath = "" Then
   MsgBox "Board path is Null,please set!"
   Exit Sub
End If
 strDebugTestjet = Text1.Text
 'Command7.Caption = "get'" & Text1.Text & "'"
 DeviceU = Text2.Text
 
If strDebugTestjet = "" Or InStr(strDebugTestjet, "short") Then
  MsgBox "The board none testjet fail"
  Exit Sub
End If
 
strDebugTestjet = Replace(strDebugTestjet, "%", "%%")
DeviceU = Replace(Trim(DeviceU), "%", "%%")
  'strDebugTestjetDVS = "basic " & Text1.Text & "get " & "'" & strDebugTestjet & "'|findn " & "'" & DeviceU & "'" & Text1.Text
 strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strBoardPath & "'" & "|comp " & "'" & strDebugTestjet & "'|exit" & l.Text

' strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'" & l.Text
 Open strToolPath & "AutoLookLog\DebugTestjet.bat" For Output As #45
 Print #45, "cd " & strBoardPath
 Print #45, strDebugTestjetDVS
 Close #45
 strDebugTestjet = Replace(strDebugTestjet, "%%", "%")
 
 A = Shell(strToolPath & "AutoLookLog\DebugTestjet.bat", 0)
  S = "re-save|comp'" & LCase(Text1.Text) & "'|exit"
 Clipboard.SetText S
  
 Exit Sub

EX:
  
End Sub

Private Sub Command12_Click()
'Dim DelFile As String
'On Error Resume Next
'                      DelFile = left(strReportPath, Len(strReportPath) - 3)
'                      Open DelFile For Output As #32
'                        Print #32, "The board fail shorts!"
'                      Close #32

 Dim DK1orDB2testjet As String, strTmpModetestjet As String
 Dim intGG As Integer
 Dim strTmpModeTME As String
 Dim strTmpDIRTMP As String
 
On Error GoTo EX
If strToolPath = "" Then
   MsgBox "Board path is Null,please set!"
   Exit Sub
End If
 strDebugTestjet = Text1.Text
 
 DeviceU = Text2.Text
 
If strDebugTestjet = "" Or InStr(strDebugTestjet, "short") Then
  MsgBox "The board none testjet fail"
  Exit Sub
End If
 
strDebugTestjet = Replace(strDebugTestjet, "%", "%%")
DeviceU = Replace(Trim(DeviceU), "%", "%%")
  'strDebugTestjetDVS = "basic " & Text1.Text & "get " & "'" & strDebugTestjet & "'|findn " & "'" & DeviceU & "'" & Text1.Text
 strDebugTestjetTT = Replace(strDebugTestjet, "%%", "%")
 
 If Dir(strBoardPath & sBoardVersion & "\" & strDebugTestjetTT) <> "" Then
    '     strDebugTestjetDVS = "basic " & l.Text & "comp" & "'" & sBoardVersion & "\" & strDebugTestjet & " ';version ''|exit" & l.Text
          strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strBoardPath & "'" & "|comp" & "'" & sBoardVersion & "\" & strDebugTestjet & "';version ''|exit" & l.Text
          
          Open strToolPath & "AutoLookLog\DebugTestjet.bat" For Output As #45
         Print #45, "cd " & strBoardPath '& sBoardVersion & "\"
        Print #45, strDebugTestjetDVS
        Close #45
    
    Else
      strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strBoardPath & "'" & "|comp" & "'" & strDebugTestjet & "';version ''|exit" & l.Text
 
       Open strToolPath & "AutoLookLog\DebugTestjet.bat" For Output As #45
    Print #45, "cd " & strBoardPath
    Print #45, strDebugTestjetDVS
    Close #45
 End If
' strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'" & l.Text
 
 strDebugTestjet = Replace(strDebugTestjet, "%%", "%")
 
 A = Shell(strToolPath & "AutoLookLog\DebugTestjet.bat", 0)
 If sBoardVersion = "" Then
     If sBoardVersionBasic = "BASE" Then
         S = "re-save|comp" & l.Text & LCase(Text1.Text) & l.Text & "|exit"
       Else
         S = "re-save|comp" & l.Text & LCase(Text1.Text) & l.Text & ";version''|exit"
     End If
   Else
     S = "re-save|comp" & l.Text & sBoardVersion & "\" & LCase(Text1.Text) & l.Text & ";version''|exit"
 End If
 Clipboard.SetText S
  
 Exit Sub

EX:










End Sub

Private Sub Command2_Click()
If Command2.Caption = "LockText" Then
    Timer1.Enabled = False
   ' Command2.Enabled = False
    Command3.Enabled = True
    NotLoadFile = False
    Command2.Caption = "UnLock"
    LockText.Caption = "UnLock"
    txtNGLog.ForeColor = &H0&
  Else
  LockText.Caption = "LockText"
   Command2.Caption = "LockText"
 ' Command2.Enabled = True
   NotLoadFile = True
  Command3.Enabled = False
  Timer1.Enabled = True
  txtNGLog.ForeColor = &HFF&
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
    Open strReportPath For Output As #12
    Close #12
    Command2.Enabled = True
    Command3.Enabled = False
 Timer1.Enabled = True
End Sub

Private Sub Command4_Click()
frmAutoNet1.Show
NetWork.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command5_Click()
 
'If bFrmClose <> True Then
 Unload frmAnalogDebug
 Unload frmAnalogDebug_GaoJi
' Unload frmAnalogDebugnet
 'Unload frmAutoNet1
 Unload frmAllPath
 Unload frmView
 Unload frmView2
 Unload frmTop_5
' Unload frmCreate
'  Unload frmBoardView
 Unload Me
'End If
End Sub

Private Sub Command6_Click()
bFrmClose = True

frmAllPath.Show
Unload frmTop_5
Unload frmAuto1
End Sub

Private Sub Command7_Click()
' Dim DK1orDB2testjet As String, strTmpModetestjet As String
' Dim intGG As Integer
' Dim strTmpModeTME As String
' Dim strTmpDIRTMP As String
'
'On Error GoTo EX
'If strToolPath = "" Then
'   MsgBox "Board path is Null,please set!"
'   Exit Sub
'End If
' strDebugTestjet = Text1.Text
'
' DeviceU = Text2.Text
'
' If Command7.Cancel = "get""""" Then
'     Exit Sub
' End If
''If strDebugTestjet = "" Or InStr(strDebugTestjet, "short") Then
''  MsgBox "The board none testjet fail"
''  Exit Sub
''End If
'
'strDebugTestjet = Replace(strDebugTestjet, "%", "%%")
'DeviceU = Replace(Trim(DeviceU), "%", "%%")
'  'strDebugTestjetDVS = "basic " & Text1.Text & "get " & "'" & strDebugTestjet & "'|findn " & "'" & DeviceU & "'" & Text1.Text
' strDebugTestjetDVS = "basic " & L.Text & "msi " & "'" & strBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'|findn '" & DeviceU & "'" & L.Text
'
'' strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'" & l.Text
' Open strToolPath & "AutoLookLog\DebugTestjet.bat" For Output As #45
' Print #45, "cd " & strBoardPath
' Print #45, strDebugTestjetDVS
' Close #45
' strDebugTestjet = Replace(strDebugTestjet, "%%", "%")
'
' A = Shell(strToolPath & "AutoLookLog\DebugTestjet.bat", 0)
'  S = "com" & L.Text & LCase(Text1.Text) & L.Text
' Clipboard.SetText S
'
' Exit Sub
'
'EX:
End Sub


Private Sub Command8_Click()
 Dim DK1orDB2testjet As String, strTmpModetestjet As String
 Dim intGG As Integer
 Dim strTmpModeTME As String
 Dim strTmpDIRTMP As String
 Dim FindTestPindText As String
 
On Error GoTo EX
If strToolPath = "" Then
   MsgBox "Board path is Null,please set!"
   Exit Sub
End If
 strDebugTestjet = Text1.Text
 
 DeviceU = Text2.Text
 
If strDebugTestjet = "" Then
  MsgBox "The board none testjet fail"
  Exit Sub
End If
 
 
 FindTestPindText = TestjetPinName(ListPinCont)
 FindTestPindText = Replace(FindTestPindText, """", "'""'")
strDebugTestjet = Replace(strDebugTestjet, "%", "%%")
DeviceU = Replace(Trim(DeviceU), "%", "%%")
  'strDebugTestjetDVS = "basic " & Text1.Text & "get " & "'" & strDebugTestjet & "'|findn " & "'" & DeviceU & "'" & Text1.Text
  strDebugTestjetTT = Replace(strDebugTestjet, "%%", "%")

If Dir(strBoardPath & sBoardVersion & "\" & strDebugTestjetTT) <> "" And sBoardVersion <> "" Then '20100827
   strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strBoardPath & "'" & "|get " & "'" & sBoardVersion & "\" & strDebugTestjet & "'|findn '" & FindTestPindText & "'" & l.Text '20100827
    Open strToolPath & "AutoLookLog\DebugTestjet.bat" For Output As #45
     Print #45, "cd " & strBoardPath '& sBoardVersion & "\"
     Print #45, strDebugTestjetDVS
    Close #45
  Else '20100827
   strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'|findn '" & FindTestPindText & "'" & l.Text '20100827

 Open strToolPath & "AutoLookLog\DebugTestjet.bat" For Output As #45
  Print #45, "cd " & strBoardPath
 Print #45, strDebugTestjetDVS
 Close #45 '20100827


End If '20100827
' strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'" & l.Text
 

 strDebugTestjet = Replace(strDebugTestjet, "%%", "%")
 
 A = Shell(strToolPath & "AutoLookLog\DebugTestjet.bat", 0)
 'S = "comp" & L.Text & LCase(Text1.Text) & L.Text
If sBoardVersion = "" Then
     S = "re-save|comp" & l.Text & LCase(Text1.Text) & l.Text & "|exit"
  Else
     S = "re-save|comp" & l.Text & sBoardVersion & "\" & LCase(Text1.Text) & l.Text & ";version''|exit"
  End If
 Clipboard.SetText S
EX:

End Sub

Private Sub Command9_Click()
If Command9.Caption = "+" Then
    Me.Height = 9000 '8865
'    Me.Width = 4800
    Me.ScaleHeight = 8310 '8175
    
     Me.Width = 5505
'    Me.Width = 4800
     Me.ScaleWidth = 5385
     txtNGLog.Width = Me.ScaleWidth - 250
'    Me.ScaleWidth = 4680

 

    Command9.Caption = "-"
    MaxWindow.Caption = "MinWindow"
  Else
    Me.Height = 6240 '6030
'    Me.Width = 4800
    Me.ScaleHeight = 5550 '5340
'    Me.ScaleWidth = 4680
    Command9.Caption = "+"
    MaxWindow.Caption = "MaxWindow"
End If
End Sub

Private Sub CreateBoardView_Click()
If CreateBoardView.Caption <> "CreateBoardViewTrue" Then
   
  frmCreate.Show
   Unload Me
  CreateBoardView.Caption = "CreateBoardViewTrue"
  Else
  Unload frmCreate
  CreateBoardView.Caption = "CreateBoardViewFalse"
End If
End Sub

Private Sub DebugAnalog_Click()
Call Command1_Click
End Sub

Private Sub DebugTestJet_Click()
Call Command11_Click
End Sub

Private Sub Exit_Click()
 
If bFrmClose <> True Then
 Unload frmAnalogDebug
' Unload frmAnalogDebugnet
 'Unload frmAutoNet1
  Unload frmAnalogDebug_GaoJi
 Unload frmAllPath
 Unload frmView
 Unload frmView2
 Unload frmTop_5
' Unload frmCreate
'  Unload frmBoardView
 Unload Me
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

  If App.PrevInstance Then
     MsgBox "The application is already open", vbInformation, "Error"
'     Unload frmAnalogDebug
' Unload frmAnalogDebugnet
' Unload frmAutoNet1
' Unload frmAllPath
' Unload frmView
' Unload frmView2
'
'   Unload Me
   End
   Exit Sub
   
  End If

Timer2.Enabled = True

 '!----------------IP-NAME---------------------
    Dim l1 As String
    Dim l2 As Long
    Dim l3 As Long
    l2 = 255
    l1 = String$(l2, " ")
    '得到本机的名字。
    l3 = GetComputerName(l1, l2)
    getname = ""
    If l3 <> 0 Then
        getname = left(l1, l2)
    End If
    Text3.Text = getname
    'Socket初始化
    SocketsInitialize
    
 '!--------------------------------------
    
NgReportListTestjetPin.Caption = "NgReportListTestjetPin" & ReportTestjet
ScreenViewFailDevice.Caption = "ScreenViewFailDevice" & ListView

' 窗口置前
  Dim myval
'  myval = SetWindowPos(frmAuto1.hwnd, -1, 0, 0, 0, 0, 3)
'  CenterScreen.Caption = "Windows Default"

 '窗口正常
'myval = SetWindowPos(frmAuto1.hwnd, -2, 0, 0, 0, 0, 3)
'CenterScreen.Caption = "CenterScreen"



  
  
  
  
  
  
  
'
'
''
''
    Set m_cN = New cNeoCaption
   Skin Me, m_cN

''
'
''
''
''
''
'
'





Me.Height = 6240 '6030
'Me.Width = 4800
Me.ScaleHeight = 5550 '5340
'Me.ScaleWidth = 4680
NotLoadFile = True




Timer1.Enabled = True
Dim MyStr As String
strToolPath = App.path
If right(strToolPath, 1) <> "\" Then strToolPath = strToolPath & "\"
MkDir strToolPath & "AutoLookLog"
Open strToolPath & "AutoLookLog\NotDelete.sys" For Output As #77


'strReportPath = "C:\Documents and Settings\great\My Documents\pomona-a  retest070822.txt"
   If FileLen(strToolPath & "AutoLookLog\Path.ini") = 0 Then
     Kill strToolPath & "AutoLookLog\Path.ini"
   End If
   If Dir(strToolPath & "AutoLookLog\Path.ini") = "" Then
'      Open strToolPath & "AutoLookLog\Path.ini" For Output As #1
'         Print #1, "#IyetPath#:Null"
'         Print #1, "#NetIyetPath#:Null"
'         Print #1, "#BoardPath#:Null"
'         Print #1, "#NetBoardPath#:Null"
'         Print #1, "#Name#:Null"
'         Print #1, "#NetWorkName#:"
'      Close #1
     Else
      Open strToolPath & "AutoLookLog\Path.ini" For Input As #2
         Do Until EOF(2)
         DoEvents
          Line Input #2, MyStr
            MyStr = Trim(UCase(MyStr))
            If left(MyStr, 1) <> "!" Then
               If left(MyStr, 11) = "#IYETPATH#:" Then
                  strReportPath = right(MyStr, Len(MyStr) - 11)
               End If
               If left(MyStr, 12) = "#BOARDPATH#:" Then
                  strBoardPath = right(MyStr, Len(MyStr) - 12)
               End If
'               If left(MyStr, 14) = "#NETIYETPATH#:" Then
'                  strNetReportPath = right(MyStr, Len(MyStr) - 14)
'               End If
               If left(MyStr, 7) = "#NAME#:" Then
                  strName = right(MyStr, Len(MyStr) - 7)
               End If
'               If left(MyStr, 15) = "#NETBOARDPATH#:" Then
'                  strNetBoardPath = right(MyStr, Len(MyStr) - 15)
'               End If
'               If left(MyStr, 14) = "#NETWORKNAME#:" Then
'                  strNetName = right(MyStr, Len(MyStr) - 14)
'               End If
               
            End If
         Loop
      Close #2
   End If
  ' Text3.Text = strName
 If strReportPath = "" Or strReportPath = "NULL" Then
   frmAuto1.Caption = "please set path"
   Timer1.Enabled = False
   Else
   frmAuto1.Caption = UCase(Text3.Text)
   Timer1.Enabled = True
 End If
 If strBoardPath <> "" Then
    tmpBoardName = left(strBoardPath, Len(strBoardPath) - 1)
    Dim strFenPei() As String
    strFenPei = Split(tmpBoardName, "\")
    frmAuto1.Caption = UCase(Text3.Text) & " " & strFenPei(UBound(strFenPei))
 End If
 ShortsReTest_1_Times = True
 ShortsReTestOneTimes.Caption = "ShortsReTest1TimesTrue"
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Height > 9000 Then

    Me.Height = 9000 '8865
'    Me.Width = 4800
    Me.ScaleHeight = 8310 '8175
'    Me.ScaleWidth = 4680
End If
If Me.Width > 5505 Then
   Me.Width = 5505
   Me.ScaleWidth = 5385
End If
'If Me.Width < 5505 Then
' 'txtNGLog.Width = 4935
'  txtNGLog.Width = Me.ScaleWidth - 250
'   'txtNGLog.ScrollBars vbBoth
'End If
txtNGLog.Width = Me.ScaleWidth - 250
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SocketsCleanup
If bFrmClose <> True Then
 StopRun = True
 Unload frmAnalogDebug
 Unload frmAnalogDebug_GaoJi
' Unload frmAnalogDebugnet
 'Unload frmAutoNet1
 Unload frmAllPath
 Unload frmView
 Unload frmView2
' Unload frmCreate
 'Unload frmBoardView
 unloadTop = True
 Unload frmTop_5
 Unload Me
 Close #77
Set m_cN = Nothing
 Kill strToolPath & "AutoLookLog\NotDelete.sys"
 ' TerminateTask "AutoLookIyet"
' End
End If
End Sub

Private Sub List1_Click()
Command8.Caption = List1.Text
'MsgBox List1.ListCount
ListPinCont = List1.ListIndex
 
If Command8.Caption = "" Then
   Command8.Enabled = False
   Else
   Command8.Enabled = True
End If
End Sub

Private Sub List1_DblClick()
Call Command8_Click
End Sub

Private Sub ListDVS_DblClick()
If ListDVS.Text = "" Then Exit Sub

 Call Command1_Click
End Sub

Private Sub LockText_Click()
'If Command2.Caption = "LockText" Then
'    Timer1.Enabled = False
'   ' Command2.Enabled = False
'    Command3.Enabled = True
'    Command2.Caption = "UnLock"
'    LockText.Caption = "UnLock"
'  Else
'  LockText.Caption = "LockText"
'   Command2.Caption = "LockText"
' ' Command2.Enabled = True
'  Command3.Enabled = False
'  Timer1.Enabled = True
'
'End If
Call Command2_Click
End Sub

Private Sub MaxWindow_Click()
Call Command9_Click
End Sub

Private Sub NetWork_Click()
Call Command4_Click
End Sub

Private Sub NgReportListTestjetPin_Click()
ReportTestjet = Not ReportTestjet
NgReportListTestjetPin.Caption = "NgReportListTestjetPin" & ReportTestjet
End Sub

Private Sub OpenIyetFile_Click()
Call Command6_Click
'On Error GoTo EX
'
'  Me.CommonDialog1.Filter = "*.eee|*.eee|*.txt|*.txt|*.*|*.*"
'  Me.CommonDialog1.CancelError = True
'  Me.CommonDialog1.ShowOpen
'   If CommonDialog1.FileName = "" Then Exit Sub
'      If LCase(CommonDialog1.FileTitle) = "failure.txt" Then: MsgBox "The failure.txt is iyet test retry file ,please open failure.eee file !", vbCritical: Exit Sub
'
'   strReportPath = CommonDialog1.FileName
'      Open strToolPath & "AutoLookLog\Path.ini" For Output As #3
'         Print #3, "#IyetPath#:" & strReportPath
'         Print #3, "#NetIyetPath#:" & strNetReportPath
'         Print #3, "#BoardPath#:" & strBoardPath
'         Print #3, "#NetBoardPath#:" & strNetBoardPath
'         Print #3, "#Name#:"; strName
'         Print #3, "#NetWorkName#:" & strNetName
'     Close #3
'  Exit Sub
'EX:
End Sub


Private Sub OpenNetIyetFile_Click()
On Error GoTo EX

  Me.CommonDialog1.Filter = "*.eee|*.eee|*.txt|*.txt|*.*|*.*"
  Me.CommonDialog1.CancelError = True
  Me.CommonDialog1.ShowOpen
   If CommonDialog1.FileName = "" Then Exit Sub
      If LCase(CommonDialog1.FileTitle) = "failure.txt" Then: MsgBox "The failure.txt is iyet test retry file ,please open failure.eee file !", vbCritical: Exit Sub

   strNetReportPath = CommonDialog1.FileName
     Open strToolPath & "AutoLookLog\Path.ini" For Output As #3
         Print #3, "#IyetPath#:" & strReportPath
         Print #3, "#NetIyetPath#:" & strNetReportPath
         Print #3, "#BoardPath#:" & strBoardPath
         Print #3, "#NetBoardPath#:" & strNetBoardPath
         Print #3, "#Name#:"; strName
         Print #3, "#NetWorkName#:" & strNetName
     Close #3
  Exit Sub
   
  Exit Sub
EX:
End Sub

Private Sub OpenTestjetFile_Click()
'Call Command7_Click
End Sub

Private Sub ScreenViewFailDevice_Click()
ListView = Not ListView
ScreenViewFailDevice.Caption = "ScreenViewFailDevice" & ListView
End Sub

Private Sub SetPath_Click()
Call Command6_Click
End Sub

Private Sub ShortsReTestOneTimes_Click()
ShortsReTest_1_Times = Not ShortsReTest_1_Times
If ShortsReTest_1_Times = False Then
   ShortsReTestOneTimes.Caption = "ShortsReTest1TimesFalse"
  Else
   
  ShortsReTestOneTimes.Caption = "ShortsReTest1TimesTrue"
End If
End Sub

Private Sub text3_KeyUp(KeyCode As Integer, Shift As Integer)
Dim MyStr As String
Dim tmpStr As String
If KeyCode = 13 Then
   If Dir(strToolPath & "AutoLookLog\Path.ini") = "" Then
       frmAuto1.Caption = "please set path"
'      Open strToolPath & "AutoLookLog\Path.ini" For Output As #1
'         Print #1, "#IyetPath#:Null"
'         Print #1, "#NetIyetPath#:Null"
'         Print #1, "#BoardPath#:Null"
'         Print #1, "#NetBoardPath#:Null"
'         Print #1, "#Name#:Null"
'         Print #1, "#NetWorkName#:"
'      Close #1
     Else
      Open strToolPath & "AutoLookLog\Path.ini" For Input As #2
         Do Until EOF(2)
          Line Input #2, MyStr
            MyStr = Trim(UCase(MyStr))
            If left(MyStr, 1) <> "!" Then
               If left(MyStr, 7) = "#NAME#:" Then
                  strName = right(MyStr, Len(MyStr) - 7)
               End If
            End If
         Loop
      Close #2
   End If
   tmpStr = Trim(Text3.Text)
   Open strToolPath & "AutoLookLog\Path.ini" For Output As #3
         Print #3, "#IyetPath#:" & strReportPath
         Print #3, "#NetIyetPath#:" & strNetReportPath
         Print #3, "#BoardPath#:" & strBoardPath
         Print #3, "#NetBoardPath#:" & strNetBoardPath
         Print #3, "#Name#:"; tmpStr
         Print #3, "#NetWorkName#:" & strNetName
   Close #3
End If
End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
 
   Dim strCurrentDeviceName As String
   Dim MyStr As String
   Dim intTestjetPin As Integer
   Dim FindTestFile As Boolean
   Dim MyStr1 As String
   Dim i As Integer
   Dim G As Integer
   Dim OpenFormNode() As String
   Dim strOpenFormNodeName As String
   Dim strOpenToNodeName As String
    Dim TestJetView() As String
   Dim tmpStr As String
   Dim ReTestDataTime As String
   Dim OutFileText As String
   Dim strFenPei_Testjet() As String
   Dim strTestjetMSG As String
   Dim FindShorts As Boolean
   Dim FindOpen As Boolean
   Dim FindTestjet As Boolean
   Dim strShorts() As String
   Dim strOpen() As String
   Dim strShorts1() As String
   Dim strOpen1() As String
   Dim tmpShort As String
   Dim tmpOpen As String
   If StopRun = True Then Timer1.Enabled = False: Exit Sub
   ReTestDataTime = Format(Date, "YYMMDD")
 If Dir(strReportPath) = "" Then
    Open strReportPath For Output As #12
    Close #12
    Exit Sub
 End If
'  If Dir(strReportPath, vbDirectory) = vbDirectory Then
'    Open strReportPath For Output As #12
'    Close #12
'    Exit Sub
' End If
 
 i = 0
 TextLineLen = 0  '20200827
  If FileLen(strReportPath) > 0 Then
     FindTestFile = False
     ListDVS.Clear
     strSN = ""
     uu = 0
     uu2 = 0
     txtNGLog.Text = ""
     Text1.Text = ""
     Text2.Text = ""
     intTestjetPin = 0
     oanalog.Enabled = False
     otestjet.Enabled = False
     oshorts.Enabled = False
     oOpen.Enabled = False
     intTestjetPin = 0
     sBoardVersion = ""
     sBoardVersionBasic = ""
     Command8.Caption = ""
     Command8.Enabled = False
     List1.Clear
    ' Text4.Text = ""
     oanalog.ForeColor = &H80000012
     oshorts.ForeColor = &H80000012
     oOpen.ForeColor = &H80000012
     otestjet.ForeColor = &H80000012
     If strReportPath = "" Then Exit Sub
    Open strReportPath For Input As #1
       Do Until EOF(1)
       DoEvents
       If NotLoadFile = False Then
         Exit Do
         Exit Sub
       End If
          If TextLineLen > 500 Then Exit Do     '20200827
          Line Input #1, MyStr1
             MyStr = UCase(Trim(MyStr1))

            ' board version
               If InStr(MyStr, "BOARD VERSION:") <> 0 Then
                  sBoardVersion = Trim(right(MyStr, Len(MyStr) - 14))
                  frmAuto1.Caption = "board version is: " & sBoardVersion
               End If
               If UCase(sBoardVersion) = "**BASE**" Then
                   sBoardVersion = ""
                   sBoardVersionBasic = "BASE"
               End If
             'analog
            If InStr(MyStr, "HAS FAILED") <> 0 Then
                oanalog.Value = True
                oanalog.Enabled = True
                oanalog.ForeColor = &HFF&
                tmpStr = Trim(left(MyStr, Len(MyStr) - 10))
'               Open strToolPath & "AutoLookLog\RestTimes" & ReTestDataTime & ".dll" For Append As #2
'                  Print #2, "AnalogType:" & tmpStr
'               Close #2
               
               'Retest Times-------------------
               strCurrentDeviceName = tmpStr
                 Call Top_10(strCurrentDeviceName, "[Analog]")
               '-------------------------------
               For G = 0 To i
                 If ListDVS.List(G) = LCase(tmpStr) Then
                     tmpStr = ""
                 End If
               Next G
               If tmpStr <> "" Then
                  ListDVS.List(i) = LCase(tmpStr)
                  i = i + 1
               End If
            End If
            
            'shorts
             If InStr(MyStr, "SHORTS REPORT FOR ") <> 0 Then
                Text1.Text = Trim(right(MyStr, Len(MyStr) - 18))
                 Text1.Text = Replace(Text1.Text, """", "")
                 Text1.Text = LCase(Replace(Text1.Text, ".", ""))
                 FindShorts = True
                 FindOpen = True
             End If
             If FindShorts = True And FindOpen = True Then
                If left(Trim(MyStr), 7) = "SHORT #" Then
                    FindOpen = False
                    oshorts.Value = True
                    oshorts.Enabled = True
                    oshorts.ForeColor = &HFF&
                 'ShortsReTest_1_Times
                    If ShortsReTest_1_Times = True Then
                      DelFile = left(strReportPath, Len(strReportPath) - 3)
                      Open DelFile For Output As #32
                        Print #32, "The board fail shorts!"
                      Close #32
                    End If
                    
                End If
              'Open
                If left(Trim(MyStr), 6) = "OPEN #" Then
                    FindOpen = False
                    oOpen.Value = True
                    oOpen.Enabled = True
                    oOpen.ForeColor = &HFF&
                End If
          
             End If

             
             
             
             'testjet
             If InStr(MyStr, "TESTJET REPORT FOR ") <> 0 Then
                  Text1.Text = Trim(right(MyStr, Len(MyStr) - 18))
                  Text1.Text = Replace(Text1.Text, """", "")
                  Text1.Text = LCase(Replace(Text1.Text, ".", ""))
                  FindTestjet = True
                  otestjet.Value = True
                  otestjet.Enabled = True
                  otestjet.ForeColor = &HFF&
             End If

             If FindTestjet = True Then
               If left(MyStr, 6) = "OPEN #" And InStr(MyStr, "DEVICE") <> 0 Then
                  Text2.Text = right(MyStr, Len(MyStr) - (InStr(MyStr, "DEVICE") + 6))
               End If
               If left(MyStr, 3) = "PIN" And InStr(MyStr, "NODE") <> 0 And intTestjetPin <= 20 Then
                    'intTestjetPin list1 >20 then stop find testjet
                     
                      If Dir(strBoardPath & Text1.Text) <> "" Or Dir(strBoardPath & sBoardVersion & "\" & Text1.Text) <> "" Then
                         Dim TmpPinState As String
                         If Dir(strBoardPath & sBoardVersion & "\" & Text1.Text) <> "" Then '20100827
                            Open strBoardPath & sBoardVersion & "\" & Text1.Text For Input As #11  '20100827
                            Else '20100827
                            Open strBoardPath & Text1.Text For Input As #11   '20100827 '
                         End If '20100827
                            Do Until EOF(11)
'                               If inttextjetpin > 20 Then
'                                    FindTestFile = False
'                                    Exit Do
'                               End If

                               DoEvents
                               Line Input #11, TmpPinState
                               TmpPinState = Trim(TmpPinState)
                               If TmpPinState <> "" Then
                                  If left(TmpPinState, 1) <> "!" Then
                                     If left(TmpPinState, 7) = "device " And InStr(TmpPinState, LCase(Text2.Text)) <> 0 Then
                                         FindTestFile = True
                                     End If
                                     If FindTestFile = True Then
                                        Dim tmp As String
                                        tmp = " " & Trim(Mid(MyStr, InStr(MyStr, "PIN") + 4, InStr(MyStr, "NODE") - 5)) & "; "
                                         
                                          ' TmpPinState = LCase(TmpPinState)
                                          ' tmp = Trim(LCase(tmp))
                                           Debug.Print "-----------------"
                                           Debug.Print Replace(left(TmpPinState, 9), " ", "")
                                           Debug.Print InStr(LCase(TmpPinState), LCase(tmp))
                                           Debug.Print LCase(TmpPinState)
                                           Debug.Print LCase(tmp)
                                           TmpPinState1 = Replace(TmpPinState, """", "")
                                        If Replace(left(TmpPinState, 9), " ", "") = "testpins" And InStr(TmpPinState1, tmp) <> 0 Then
                                           List1.List(intTestjetPin) = LCase(Text2.Text) & ";" & Trim(left(TmpPinState1, InStr(TmpPinState1, "!") - 1))
                                           TestjetPinName(intTestjetPin) = TmpPinState
                                           intTestjetPin = intTestjetPin + 1
                                           FindTestFile = False
                                           Exit Do
                                        End If
                                        
                                        
                                        
                                       If left(TmpPinState, 3) = "end" Then
                                           FindTestFile = False
                                       End If
                                        
                                     End If
                                  End If
                               End If
                            Loop
                         Close #11
                      End If
               End If
             End If
             
 
              If otestjet.Value = True Then
                   If left(Trim(UCase(MyStr)), 3) = "PIN" Then
                      Debug.Print MyStr
                       STRL1 = MyStr
                    
                        Do
                            strMyStrtmp = Replace(STRL1, "  ", " ")
                            If STRL1 = strMyStrtmp Then Exit Do
                            STRL1 = strMyStrtmp
                        Loop
                      strFenPei_Testjet = Split(STRL1, " ")
                      strTestjetMSG = strFenPei_Testjet(0) & " " & strFenPei_Testjet(1) & "," & strFenPei_Testjet(3)
                      
                          Call Top_10(Text2.Text & "," & strTestjetMSG, "[Testjet]")
                      '
                   End If
              End If
             
'             If oshorts.Value = True Then
'                If left(Trim(MyStr), 6) = "FROM: " Then
'                    MyStr = Replace(MyStr, " ", ",")
'                    strShorts = Split(MyStr, " ")
'                   ' tmpShort = Trim(Mid(MyStr, 7, (InStr(MyStr, strShorts(UBound(strShorts)) + 1))))
'                    Open strToolPath & "AutoLookLog\RestTimes" & ReTestDataTime & ".dll" For Append As #2
'                       Print #2, "ShortsType:" & tmpShort
'                    Close #2
'                End If
'                If left(Trim(MyStr), 4) = "TO: " Then
'                    tmpShort = strShorts1(1) & strShorts1(2)
'                    Open strToolPath & "AutoLookLog\RestTimes" & ReTestDataTime & ".dll" For Append As #2
'                       Print #2, "ShortsType:" & tmpShort
'                    Close #2
'                End If
'             End If
'----------------------------------------------------------------------------------------



             If oOpen.Value = True Then
                If left(Trim(MyStr), 6) = "FROM: " Then
                     strL = MyStr
                    
                        Do
                            strMyStrtmp = Replace(strL, "  ", " ")
                            If strL = strMyStrtmp Then Exit Do
                            strL = strMyStrtmp
                        Loop

                    strOpen = Split(strL, " ")
 
                   strOpenFormNodeName = Trim(strOpen(1))

                End If
                If left(Trim(MyStr), 4) = "TO: " Then
                     strL = MyStr
                    
                        Do
                            strMyStrtmp = Replace(strL, "  ", " ")
                            If strL = strMyStrtmp Then Exit Do
                            strL = strMyStrtmp
                        Loop
                
                        
                       strOpen1 = Split(strL, " ")
                        
                    strOpenToNodeName = Trim(strOpen1(1))

                   Call Top_10(strOpenFormNodeName & "," & strOpenToNodeName, "[Open]")
                   strOpenFormNodeName = ""
                   strOpenToNodeName = ""

                End If
             End If
             strMyStrtmp = ""
             Erase strOpen
             Erase strOpen1
'---------------------------------------------------------------
'
'             If Left(MyStr, 10) = "------END," Then
'                FindOpen = False
'                FindShorts = False
'             End If
'
             If left(MyStr, 4) = "S/N:" Then
               strSN = left(Trim(right(MyStr, Len(MyStr) - 4)), 23)
                'Clipboard.SetText = Trim(right(MyStr, Len(MyStr) - 4))
             End If
             'boardview open
             If BoardViewTrue = True Then
                     If oOpen.Value = True And UCase(left(MyStr, 5)) = "FROM:" Then
                         OpenFormNode = Split(MyStr, " ")
                          For II = 2 To UBound(OpenFormNode)
                              If OpenFormNode(II) <> "" Then
                                 OpenViewText.Caption = OpenFormNode(II)
                                  Exit For
                              End If
                          Next
                          Debug.Print OpenFormNode(II)
                           If OpenViewText.Caption <> "" Then
                                 BoardViewDevice = OpenViewText.Caption
                                AnalogView = False
                                OpenView = True
                           End If
                    End If
              End If
             
             
             
             
            txtNGLog.Text = txtNGLog.Text + MyStr1 + Chr(13) + Chr(10)
            TextLineLen = TextLineLen + 1  '20200827
       Loop
  
    Close #1
    
    'FileCopy strReportPath, strToolPath & "AutoLookLog\RetestNG.dll"
    OutFileText = txtNGLog.Text
    strDate = Format(Date$, "yymmdd")
    Open strToolPath & "AutoLookLog\RetestNG" & strDate & ".tmp" For Append As #5
     Print #5, OutFileText
     Print #5, "[" & strReportPath & "]"
     ReTestTimesOut = ReTestTimesOut + 1
    Close #5
    Open strToolPath & "AutoLookLog\CurrNg.txt" For Output As #5
       Print #5, OutFileText
    Close #5
'    Open strToolPath & "AutoLookLog\NGLog.log" For Append As #6
'     Print #6, OutFileText
'    Close #6
    Open strReportPath For Output As #2
    Close #2
  End If
'    If ReTestTimesOut > 3 Then
'      ReTestTimesOut = 0
'      Kill strToolPath & "AutoLookLog\RetestNG.tmp"
'    End If
'View Device

If ListView = True Then
    If ListDVS.List(0) <> "" Then
       ViewText = UCase(ListDVS.List(0))
       If uu = 0 Then
         frmView.Show
         uu = uu + 1
       End If
    End If
    If ListDVS.List(1) <> "" Then
       ViewText2 = UCase(ListDVS.List(1))
       If uu2 = 0 Then
         frmView2.Show
         uu2 = uu2 + 1
       End If
    End If
    If List1.List(0) <> "" Then
      TestJetView = Split(List1.List(0), ";")
       
       ViewText = Replace(LCase(TestJetView(0) & TestJetView(1)), "test", "")
       ViewText = Replace(ViewText, "pins", "pin")
       If uu = 0 Then
         frmView.Show
         uu = uu + 1
       End If
    End If
    If List1.List(1) <> "" Then
      TestJetView = Split(List1.List(1), ";")
       
       ViewText2 = Replace(LCase(TestJetView(0) & TestJetView(1)), "test", "")
       ViewText2 = Replace(ViewText2, "pins", "pin")
       If uu2 = 0 Then
         frmView2.Show
         uu2 = uu2 + 1
       End If
    End If
    
End If
'Report testjet pin
If ReportTestjet = True Then
    If List1.List(0) <> "" Then
        Open strBoardPath & "ReportAddPin.dll" For Output As #11
            For u = 0 To intTestjetPin
              tmpSS = List1.List(u)
              tmpSS = Replace(tmpSS, "test ", "")
              tmpSS = Replace(tmpSS, "threshold ", "")
              Print #11, tmpSS
            Next u
        Close #11
    End If

    
    
    
End If

If BoardViewTrue = True Then
         If OpenView = False Then
           BoardViewDevice = ListDVS.Text
         End If
       If OpenViewText.Caption <> "" Then
                 BoardViewDevice = OpenViewText.Caption
                AnalogView = False
                OpenView = True
        End If


     
    If ListDVS.List(0) <> "" Then
        If ListDVS.Text = "" Then
            BoardViewDevice = ListDVS.List(0)
        End If
           '
    
       If InStr(BoardViewDevice, "%") <> 0 Then
          ' BoardViewDevice = Replace(left(BoardViewDevice, Len(BoardViewDevice) - (Len(BoardViewDevice) - InStr(BoardViewDevice, "%"))), "%", "")
         Else
         
           If InStr(BoardViewDevice, "_") <> 0 Then
                BoardViewDevice = Replace(left(BoardViewDevice, Len(BoardViewDevice) - (Len(BoardViewDevice) - InStr(BoardViewDevice, "_"))), "_", "")
               Else
                If ListDVS.Text = "" Then
                    BoardViewDevice = ListDVS.List(0)
                End If
           End If
         
           
       End If
       AnalogView = True
       OpenView = False
       OpenViewText.Caption = ""
    End If
    
    If Text2.Text <> "" Then
    
         BoardViewDevice = Text2.Text
           AnalogView = True
          OpenView = False
           ListDVS.Text = ""
          OpenViewText.Caption = ""
    End If
    
    
    
    
    
    
    
   Else
      BoardViewDevice = ""
End If






   Exit Sub
EX:
End Sub



Private Sub Timer2_Timer()
On Error Resume Next
 
Dim aa
Dim tt As Integer
aa = Format(Now, "hhmm")
strTimesPath = "C:\WINDOWS\system\Top10"
'If intTime = 3600 Then
   tt = Val(aa)
   If tt > 2030 And tt < 2035 Then
        For i = 0 To 24
         frmTop_5.Text1(i).Text = ""
        Next
        Kill strTimesPath & "\*.*"
   End If
   
  If tt > 830 And tt < 835 Then
        For i = 0 To 24
         frmTop_5.Text1(i).Text = ""
        Next
        Kill strTimesPath & "\*.*"
   End If
   
  ' intTime = 0
'End If
End Sub

Private Sub txtBasicCmd_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sMystr As String
Dim sCommand As String
 On Error Resume Next
 
 
If KeyCode = 13 Then
   txtBasicCmd.Text = Trim(LCase(txtBasicCmd.Text))
   Kill (strToolPath & "AutoLookLog\BasicCommand.bat")
   If txtBasicCmd.Text = "" Then Exit Sub
   sMystr = LCase(Trim(txtBasicCmd.Text))
   'fix grap
   
   'testplan
   If sMystr = "testplan" Then
      sCommand = "basic " & """" & "msi " & "'" & strBoardPath & "'" & "|get" & "'" & "testplan'" & """"

      
    Open strToolPath & "AutoLookLog\BasicCommand.bat" For Output As #45
         Print #45, "cd " & strBoardPath '& sBoardVersion & "\"
        Print #45, sCommand
    Close #45
   End If
   
   If sMystr = "ksh" Then
      ' sCommand = "start " & strBoardPath
      
    Open strToolPath & "AutoLookLog\BasicCommand.bat" For Output As #45
        ' Print #45, "cd " & strBoardPath '& sBoardVersion & "\"
        Print #45, "start ksh"
    Close #45
   End If
   
 
   
   'dir
   If sMystr = "dir" Then
      sCommand = "start " & strBoardPath
      
    Open strToolPath & "AutoLookLog\BasicCommand.bat" For Output As #45
        ' Print #45, "cd " & strBoardPath '& sBoardVersion & "\"
        Print #45, sCommand
    Close #45
   End If
   'testorder
    If sMystr = "testorder" Then
      sCommand = "basic " & """" & "msi " & "'" & strBoardPath & "'" & "|get" & "'" & "testorder'" & """"

      
    Open strToolPath & "AutoLookLog\BasicCommand.bat" For Output As #45
          Print #45, "cd " & strBoardPath '& sBoardVersion & "\"
        Print #45, sCommand
    Close #45
   End If
    If LCase(sMystr) <> "testorder" And LCase(sMystr) <> "testplan" And LCase(sMystr) <> "dir" And LCase(sMystr) <> "ksh" Then
      sCommand = "basic " & """" & "msi " & "'" & strBoardPath & "'" & "|" & txtBasicCmd.Text & """"

      sCommand = Replace(sCommand, "%", "%%")
    Open strToolPath & "AutoLookLog\BasicCommand.bat" For Output As #45
          Print #45, "cd " & strBoardPath '& sBoardVersion & "\"
        Print #45, sCommand
    Close #45
   End If
    A = Shell(strToolPath & "AutoLookLog\BasicCommand.bat", 0)
    txtBasicCmd.Text = ""
    sMystr = ""
    sCommand = ""
    A = ""
    txtBasicCmd.SetFocus
End If
End Sub

Private Sub txtNGLog_DblClick()
'Me.Height = 8865
''Me.Width = 4800
'Me.ScaleHeight = 8175
''Me.ScaleWidth = 4680
' Command9.Caption = "-"
Call Command9_Click
End Sub
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Private Sub Skin(f As Form, cN As cNeoCaption)
    cN.ActiveCaptionColor = &HFFFFFF
    cN.InActiveCaptionColor = &HC0C0C0
    cN.ActiveMenuColor = &H0&
    cN.ActiveMenuColorOver = &H0
    cN.InActiveMenuColor = &H0&
    cN.MenuBackgroundColor = RGB(207, 203, 207)
   ' cN.CaptionFont.Name = "MS Sans Serif"
   ' cN.CaptionFont.Size = 8
   ' cN.MenuFont.Name = "MS Sans Serif"
    'cN.MenuFont.Size = 8
    cN.Attach f, f.PicCaption.Picture, f.PicBorder.Picture, 19, 20, 90, 140, 240, 400
    f.BackColor = RGB(207, 203, 207)
End Sub
Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

Sub SocketsInitialize()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String
    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
    If iReturn <> 0 Then
        MsgBox "Winsock.dll is not responding."
        End
    End If
    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        MsgBox sMsg
        End
    End If
   
    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox sMsg
        End
    End If
End Sub

Sub SocketsCleanup()
    Dim lReturn As Long
    lReturn = WSACleanup()
    If lReturn <> 0 Then
        'Socket有错误
        MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
        End
    End If
End Sub
