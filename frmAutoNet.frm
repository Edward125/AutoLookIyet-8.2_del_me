VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAutoNet1 
   ClientHeight    =   8130
   ClientLeft      =   4035
   ClientTop       =   900
   ClientWidth     =   5295
   Icon            =   "frmAutoNet.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   5295
   Begin VB.CommandButton Command12 
      Caption         =   "com'VerTestjet"
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "ComTestjet"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   7800
      Width           =   1095
   End
   Begin VB.PictureBox PicCaption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   0
      Picture         =   "frmAutoNet.frx":0442
      ScaleHeight     =   720
      ScaleWidth      =   9600
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   9600
      Begin VB.PictureBox PicBorder 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   0
         Picture         =   "frmAutoNet.frx":16C86
         ScaleHeight     =   150
         ScaleWidth      =   1050
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Clear"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox l 
      Height          =   285
      Left            =   720
      TabIndex        =   21
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
      Picture         =   "frmAutoNet.frx":17398
      TabIndex        =   18
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Enabled         =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   7440
      Width           =   3735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFC0C0&
      Height          =   1425
      ItemData        =   "frmAutoNet.frx":177DA
      Left            =   1320
      List            =   "frmAutoNet.frx":177DC
      TabIndex        =   15
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
      TabIndex        =   12
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DebugAnalog"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LockText"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5760
      Width           =   735
   End
   Begin VB.OptionButton otestjet 
      Caption         =   "Testjet"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
   Begin VB.OptionButton oanalog 
      Caption         =   "Analog"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   5040
      Width           =   855
   End
   Begin VB.OptionButton oshorts 
      Caption         =   "Short"
      Enabled         =   0   'False
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   5040
      Width           =   735
   End
   Begin VB.OptionButton oOpen 
      Caption         =   "Open"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   2400
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UnLockText"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   5760
      Width           =   1815
   End
   Begin VB.ListBox ListDVS 
      BackColor       =   &H00FFC0C0&
      Height          =   1425
      ItemData        =   "frmAutoNet.frx":177DE
      Left            =   120
      List            =   "frmAutoNet.frx":177E0
      TabIndex        =   0
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
      TabIndex        =   14
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "FileName:"
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Debug Testjeet"
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "DeviceName:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "IPName:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   615
   End
   Begin VB.Menu File_ 
      Caption         =   "File"
      Begin VB.Menu OpenNetIyetFile 
         Caption         =   "SetNetIyetFile"
      End
      Begin VB.Menu OpenTestjetFile 
         Caption         =   "OpenTestjetFile"
      End
      Begin VB.Menu eee 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Closee"
      End
   End
   Begin VB.Menu set_ 
      Caption         =   "Set.."
      Begin VB.Menu ClearText 
         Caption         =   "ClearText"
      End
      Begin VB.Menu LockText 
         Caption         =   "LockText"
      End
      Begin VB.Menu rrr 
         Caption         =   "-"
      End
      Begin VB.Menu DebugAnalog 
         Caption         =   "DebugAnalog"
      End
      Begin VB.Menu DebugTestJet 
         Caption         =   "DebugTestJet"
      End
   End
   Begin VB.Menu Window_ 
      Caption         =   "Window"
      Begin VB.Menu MaxWindow 
         Caption         =   "MaxWindow"
      End
   End
End
Attribute VB_Name = "frmAutoNet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private m_cN As cNeoCaption

'================================

Dim NotLoadFile As Boolean


Dim TestjetPinName(25) As String
Dim ListPinCont As Integer
Dim StopRun As Boolean

Private Sub ClearText_Click()
Call Command10_Click
End Sub

Private Sub Command1_Click()
On Error Resume Next
If oanalog.Value = True Then
 AnalogDeviceNameNet = ListDVS.Text
 Unload frmAnalogDebugnet
 frmAnalogDebugnet.Show
 Else
   Unload frmAnalogDebugnet
  frmAnalogDebugnet.Show
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
 
 DeviceU = Text2.Text
 
If strDebugTestjet = "" Or InStr(strDebugTestjet, "short") Then
  MsgBox "The board none testjet fail"
  Exit Sub
End If
 
strDebugTestjet = Replace(strDebugTestjet, "%", "%%")
DeviceU = Replace(Trim(DeviceU), "%", "%%")
  'strDebugTestjetDVS = "basic " & Text1.Text & "get " & "'" & strDebugTestjet & "'|findn " & "'" & DeviceU & "'" & Text1.Text
 strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strNetBoardPath & "'" & "|com " & "'" & strDebugTestjet & "'|exit" & l.Text

' strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'" & l.Text
 Open strToolPath & "AutoLookLog\DebugTestjet.bat" For Output As #45
 Print #45, "cd " & strNetBoardPath
 Print #45, strDebugTestjetDVS
 Close #45
 strDebugTestjet = Replace(strDebugTestjet, "%%", "%")
 
 A = Shell(strToolPath & "AutoLookLog\DebugTestjet.bat", 0)
  S = "com" & l.Text & LCase(Text1.Text) & l.Text
 Clipboard.SetText S
  
 Exit Sub

EX:
End Sub

Private Sub Command12_Click()
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
 strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strNetBoardPath & "'" & "|com " & "'" & strDebugTestjet & "';version''|exit" & l.Text

' strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'" & l.Text
 Open strToolPath & "AutoLookLog\DebugTestjet.bat" For Output As #45
 Print #45, "cd " & strNetBoardPath
 Print #45, strDebugTestjetDVS
 Close #45
 strDebugTestjet = Replace(strDebugTestjet, "%%", "%")
 
 A = Shell(strToolPath & "AutoLookLog\DebugTestjet.bat", 0)
  S = "com" & l.Text & LCase(Text1.Text) & l.Text
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
  Else
  LockText.Caption = "LockText"
   Command2.Caption = "LockText"
 ' Command2.Enabled = True
   NotLoadFile = True
  Command3.Enabled = False
  Timer1.Enabled = True
  
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
'frmAutoNet1.Show
'Command4.Enabled = False
End Sub





Private Sub Command7_Click()
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
  If Command7.Cancel = "get""""" Then
     Exit Sub
 End If
'If strDebugTestjet = "" Then
'  MsgBox "The board none testjet fail"
'  Exit Sub
'End If
 
strDebugTestjet = Replace(strDebugTestjet, "%", "%%")
DeviceU = Replace(Trim(DeviceU), "%", "%%")
  'strDebugTestjetDVS = "basic " & Text1.Text & "get " & "'" & strDebugTestjet & "'|findn " & "'" & DeviceU & "'" & Text1.Text
 strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strNetBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'|findn '" & DeviceU & "'" & l.Text

' strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strnetBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'" & l.Text
 Open strToolPath & "AutoLookLog\DebugTestjet.bat" For Output As #45
 Print #45, "cd " & strBoardPath
 Print #45, strDebugTestjetDVS
 Close #45
 strDebugTestjet = Replace(strDebugTestjet, "%%", "%")
 
 A = Shell(strToolPath & "AutoLookLog\DebugTestjet.bat", 0)
  S = "com" & l.Text & LCase(Text1.Text) & l.Text
 Clipboard.SetText S
  
 Exit Sub

EX:
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
 strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strNetBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'|findn '" & FindTestPindText & "'" & l.Text

' strDebugTestjetDVS = "basic " & l.Text & "msi " & "'" & strnetBoardPath & "'" & "|get " & "'" & strDebugTestjet & "'" & l.Text
 Open strToolPath & "AutoLookLog\DebugTestjet.bat" For Output As #45
  Print #45, "cd " & strBoardPath
 Print #45, strDebugTestjetDVS
 Close #45
 strDebugTestjet = Replace(strDebugTestjet, "%%", "%")
 
 A = Shell(strToolPath & "AutoLookLog\DebugTestjet.bat", 0)
 S = "com" & l.Text & LCase(Text1.Text) & l.Text
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

Private Sub DebugAnalog_Click()
Call Command1_Click
End Sub

Private Sub Exit_Click()
Unload Me
End Sub



Private Sub Command5_Click()
frmAuto1.Command4.Enabled = True

Unload Me

End Sub



Private Sub Form_Load()
On Error Resume Next

    Set m_cN = New cNeoCaption
    Skin Me, m_cN



Me.Height = 6240 '6030
'Me.Width = 4800
Me.ScaleHeight = 5550 '5340
'Me.ScaleWidth = 4680
NotLoadFile = True
Timer1.Enabled = True
Dim MyStr As String

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
          Line Input #2, MyStr
            MyStr = Trim(UCase(MyStr))
            If left(MyStr, 1) <> "!" Then
               If left(MyStr, 14) = "#NETWORKNAME#:" Then
                  strNetName = right(MyStr, Len(MyStr) - 14)
               End If
               
            End If
         Loop
      Close #2
   End If
   Text3.Text = strNetName
 If strNetReportPath = "" Or strNetReportPath = "NULL" Then
   frmAutoNet.Caption = "please set path"
   Timer1.Enabled = False
   Else
   frmAutoNet1.Caption = UCase(strNetName)
   Timer1.Enabled = True
 End If
 
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
frmAuto1.Command4.Enabled = True
frmAuto1.NetWork.Enabled = True
StopRun = True
Unload Me
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
Call Command2_Click
End Sub

Private Sub MaxWindow_Click()
Call Command9_Click
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
EX:
End Sub



Private Sub OpenTestjetFile_Click()
Call Command7_Click
End Sub

Private Sub text3_KeyUp(KeyCode As Integer, Shift As Integer)
Dim MyStr As String
Dim tmpStr As String
If KeyCode = 13 Then
   If Dir(strToolPath & "AutoLookLog\Path.ini") = "" Then
       frmAutoNet1.Caption = "please set path"
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
               If left(MyStr, 14) = "#NETWORKNAME#:" Then
                  strNetName = right(MyStr, Len(MyStr) - 14)
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
         Print #3, "#Name#:"; strName
         Print #3, "#NetWorkName#:" & tmpStr
   Close #3
End If
End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
   Dim MyStr As String
   Dim intTestjetPin As Integer
   Dim FindTestFile As Boolean
   Dim MyStr1 As String
   Dim i As Integer
   Dim G As Integer
   Dim tmpStr As String
   Dim ReTestDataTime As String
   Dim OutFileText As String
   Dim FindShorts As Boolean
   Dim FindOpen As Boolean
   Dim FindTestjet As Boolean
   Dim strShorts() As String
   Dim strOpen() As String
   Dim strShorts1() As String
   Dim strOpen1() As String
   Dim tmpShort As String
   Dim tmpOpen As String
   ReTestDataTime = Format(Date, "YYMMDD")
   If StopRun = True Then Timer1.Enabled = False: Exit Sub
 If Dir(strNetReportPath) = "" Then
    Open strNetReportPath For Output As #12
    Close #12
    Exit Sub
 End If
 
'  If Dir(strNetReportPath, vbDirectory) = vbDirectory Then
'    Open strNetReportPath For Output As #12
'    Close #12
'    Exit Sub
' End If
 i = 0
  If FileLen(strNetReportPath) > 0 Then
     FindTestFile = False
     ListDVS.Clear
     
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
     If strNetReportPath = "" Then Exit Sub
    Open strNetReportPath For Input As #1
       Do Until EOF(1)
       DoEvents
       If NotLoadFile = False Then
         Exit Do
         Exit Sub
       End If
       
          Line Input #1, MyStr1
             MyStr = UCase(Trim(MyStr1))

             'analog
            If InStr(MyStr, "HAS FAILED") <> 0 Then
                oanalog.Value = True
                oanalog.Enabled = True
                oanalog.ForeColor = &HFF&
                tmpStr = Trim(left(MyStr, Len(MyStr) - 10))
'               Open strToolPath & "AutoLookLog\RestTimes" & ReTestDataTime & ".dll" For Append As #2
'                  Print #2, "AnalogType:" & tmpStr
'               Close #2
               
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

                      If Dir(strNetBoardPath & Text1.Text) <> "" Then
                         Dim TmpPinState As String
                         Open strNetBoardPath & Text1.Text For Input As #11
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
                                        
                                        TmpPinState1 = Replace(TmpPinState, """", "")
                                        If Replace(left(TmpPinState, 9), " ", "") = "testpins" And InStr(LCase(TmpPinState1), LCase(tmp)) Then
                                           List1.List(intTestjetPin) = LCase(Text2.Text) & ";" & Trim(left(TmpPinState1, InStr(TmpPinState1, "!") - 1))
                                           TestjetPinName(intTestjetPin) = TmpPinState
                                           intTestjetPin = intTestjetPin + 1
                                           FindTestFile = False
                                           Exit Do
                                        End If
'                                        If Left(TmpPinState, 3) = "end" Then
'
'                                        End If
                                        
                                     End If
                                  End If
                               End If
                            Loop
                         Close #11
                      End If
               End If
             End If
             
             
             
             
'             If FindShorts = True Then
'                If Left(Trim(MyStr), 6) = "FROM: " Then
'                    MyStr = Replace(MyStr, " ", ",")
'                    strShorts = Split(MyStr, " ")
'                   ' tmpShort = Trim(Mid(MyStr, 7, (InStr(MyStr, strShorts(UBound(strShorts)) + 1))))
'                    Open strToolPath & "AutoLookLog\RestTimes" & ReTestDataTime & ".dll" For Append As #2
'                       Print #2, "ShortsType:" & tmpShort
'                    Close #2
'                End If
'                If Left(Trim(MyStr), 4) = "TO: " Then
'                    tmpShort = strShorts1(1) & strShorts1(2)
'                    Open strToolPath & "AutoLookLog\RestTimes" & ReTestDataTime & ".dll" For Append As #2
'                       Print #2, "ShortsType:" & tmpShort
'                    Close #2
'                End If
'             End If
'             If FindOpen = True Then
'                If Left(Trim(MyStr), 6) = "FROM: " Then
'                    strOpen = Split(MyStr, " ")
'                    tmpOpen = Trim(strOpen(1) & strOpen(2))
'
'                    Open strToolPath & "AutoLookLog\RestTimes" & ReTestDataTime & ".dll" For Append As #2
'                       Print #2, "OpenType:" & tmpOpen
'                    Close #2
'                End If
'                If Left(Trim(MyStr), 4) = "TO: " Then
'                    strOpen1 = Split(MyStr, " ")
'                    tmpOpen = Trim(strOpen1(1) & strOpen1(2))
'                    Open strToolPath & "AutoLookLog\RestTimes" & ReTestDataTime & ".dll" For Append As #2
'                       Print #2, "OpenType:" & tmpOpen
'                    Close #2
'                End If
'             End If

'
'             If Left(MyStr, 10) = "------END," Then
'                FindOpen = False
'                FindShorts = False
'             End If
'
             If left(MyStr, 4) = "S/N:" Then
              'txtSN.Text = Trim(Right(MyStr, Len(MyStr) - 4))
             End If
             
             
            txtNGLog.Text = txtNGLog.Text + MyStr1 + Chr(13) + Chr(10)
       Loop
  
    Close #1
    
    'FileCopy strnetreportPath, strToolPath & "AutoLookLog\RetestNG.dll"
    OutFileText = txtNGLog.Text
    strDate = Format(Date$, "yymmdd")
    Open strToolPath & "AutoLookLog\RetestNG" & strDate & ".tmp" For Append As #5
     Print #5, OutFileText
     ReTestTimesOut = ReTestTimesOut + 1
    Close #5
'    Open strToolPath & "AutoLookLog\NGLog.log" For Append As #6
'     Print #6, OutFileText
'    Close #6
    Open strNetReportPath For Output As #2
    Close #2
  End If
'    If ReTestTimesOut > 3 Then
'      ReTestTimesOut = 0
'      Kill strToolPath & "AutoLookLog\RetestNG.tmp"
'    End If
 Command7.Caption = "get'" & Text1.Text & "'"
   Exit Sub
EX:
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
