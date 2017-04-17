VERSION 5.00
Begin VB.Form frmAnalogDebugnet 
   BackColor       =   &H80000006&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5700
   ClientLeft      =   465
   ClientTop       =   5385
   ClientWidth     =   10110
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   365
      Left            =   8760
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8535
      Begin VB.TextBox rtxtBox 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   15
         Top             =   600
         Width           =   8295
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   8295
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   480
         MaxLength       =   1
         TabIndex        =   13
         Text            =   """"
         Top             =   840
         Width           =   150
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000008&
      Height          =   5055
      Left            =   8760
      TabIndex        =   3
      Top             =   480
      Width           =   1215
      Begin VB.CommandButton Command11 
         BackColor       =   &H8000000C&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdBom 
         BackColor       =   &H8000000C&
         Caption         =   "Top"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdTop 
         BackColor       =   &H8000000C&
         Caption         =   "Top"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ComVer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H8000000C&
         Caption         =   "!Skip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdSkip 
         BackColor       =   &H8000000C&
         Caption         =   "!Skip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdInputOk 
         BackColor       =   &H8000000C&
         Caption         =   "InputOK"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCompile 
         BackColor       =   &H8000000C&
         Caption         =   "Com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000C&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000C&
         Caption         =   "ReLoad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAnalogDebugnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private trapUndo As Boolean           'flag to indicate whether actions should be trapped
Private UndoStack As New Collection   'collection of undo elements
Private RedoStack As New Collection   'collection of redo elements
Dim strAnalogPath As String
Dim intGG  As Integer
Dim strTmpDIRTMP As String
Dim strDVSName As String
Dim strAnalogFile As String
Dim MyCharA(5000) As String
Dim intChar As Integer
Dim TrueandFLS As Boolean
Dim TxtChange As Boolean

Dim AnalogDebugPath As String
Private Sub cmdCompile_Click()

Dim strDebugAnalogDVS As String
'strDebugAnalogDVS = "basic " & Text3.Text & "com " & "'analog\" & Text2.Text & "'|exit " & Text3.Text
On Error GoTo EX
  Dim strTmp As String
  Dim strD As String
  strD = Trim(Text2.Text)
  If strD = "" Then Exit Sub
 '  Me.Caption = "Please input device name!"
 
  strTmp = Dir(AnalogDebugPath & "analog\" & strD)
'   Me.Caption = "Not find the device file!"
  If strTmp = "" Then Exit Sub
 ' If Me.Caption = "Please input device name!" Or Me.Caption = "Not find the device file!" Then Me.Caption = ""

  strDebugAnalogDVS = "basic " & Text3.Text & "msi " & "'" & AnalogDebugPath & "'" & "|com " & "'analog\" & Trim(Text2.Text) & "'|exit " & Text3.Text
 strDebugAnalogDVS = Replace(strDebugAnalogDVS, "%", "%%")
 Open AnalogDebugPath & "DebugCompile.bat" For Output As #43
  Print #43, "cd " & AnalogDebugPath
 Print #43, strDebugAnalogDVS
 Close #43
 
 
cc = Shell(AnalogDebugPath & "DebugCompile.bat", 0)
 Exit Sub
EX:
 MsgBox "Creat DebugCompile.bat Error!"
End Sub

Private Sub cmdDel_Click()
'rtxtBox.SetFocus
'SendKeys "{BACKSPACE}"
'rtxtBox.SetFocus
Dim strDebugAnalogDVS As String
'strDebugAnalogDVS = "basic " & Text3.Text & "com " & "'analog\" & Text2.Text & "'|exit " & Text3.Text
On Error GoTo EX
  Dim strTmp As String
  Dim strD As String
  strD = Trim(Text2.Text)
  If strD = "" Then Exit Sub
 '  Me.Caption = "Please input device name!"
 
  strTmp = Dir(AnalogDebugPath & "analog\" & strD)
'   Me.Caption = "Not find the device file!"
  If strTmp = "" Then Exit Sub
 ' If Me.Caption = "Please input device name!" Or Me.Caption = "Not find the device file!" Then Me.Caption = ""

  strDebugAnalogDVS = "basic " & Text3.Text & "msi " & "'" & AnalogDebugPath & "'" & "|com " & "'analog\" & Trim(Text2.Text) & "';version''|exit " & Text3.Text
 strDebugAnalogDVS = Replace(strDebugAnalogDVS, "%", "%%")
 Open AnalogDebugPath & "DebugCompile.bat" For Output As #43
  Print #43, "cd " & AnalogDebugPath
 Print #43, strDebugAnalogDVS
 Close #43
 
 
cc = Shell(AnalogDebugPath & "DebugCompile.bat", 0)
 Exit Sub
EX:
 MsgBox "Creat DebugCompile.bat Error!"



End Sub





Private Sub cmdInputOk_Click()
rtxtBox.Text = ""

If Text2.Text = "" Then
 MsgBox "Please input debug test name !"
 Text2.SetFocus
 Exit Sub
 End If
 Dim strAnalogPatha  As String
 Dim strAnalogFilea As String
 Dim strDvs As String
 Dim MyCharB(5000) As String
 Dim MyChar As String
 intChar = 0
 
 strDvs = Text2.Text
strAnalogPatha = AnalogDebugPath & "analog\" & Trim(strDvs)


  On Error GoTo EX1
 strAnalogFilea = Dir(strAnalogPatha)
 If strAnalogFilea = "" Then
 MsgBox strAnalogPatha & " not find"
 Exit Sub
 Else
 'BACKUP analog
 If Dir(strAnalogPatha & ".bak") = "" Then
 FileCopy strAnalogPatha, strAnalogPatha & ".bak"
 End If
 
 Open strAnalogPatha For Input As #52
 Do Until EOF(52)
   
   Line Input #52, MyChar
   MyCharB(intChar) = MyChar
   intChar = intChar + 1
   If left(MyChar, Len("!!!!    2")) = "!!!!    2" Then
      Text1.Text = MyChar + Chr(13) + Chr(10)
      Else
        rtxtBox.Text = rtxtBox.Text + MyChar + Chr(13) + Chr(10)
   End If
Loop
 Close #52
 End If
  frmAnalogDebug.Caption = strAnalogPatha & "      " & intChar & " Line"
 
 If TrueandFLS = True Then
' rtxtBox.ForeColor = &HFF00FF
 TrueandFLS = False
 GoTo Color1
 
 End If
 If TrueandFLS = False Then
 'rtxtBox.ForeColor = &HFF0000
 TrueandFLS = True
 End If
Color1:
 Exit Sub
EX1:
 MsgBox "Open " & strAnalogPatha & " Error"
End Sub


Private Sub cmdSkip_Click()
Call Comment
rtxtBox.SetFocus
End Sub
Private Sub Comment()
Dim strSel As String
Dim arSel() As String
Dim i As Integer
Dim iStart As Long, iEnd As Long, iTmp As Long


With Me.rtxtBox
    iStart = .SelStart
    iEnd = .SelStart + .SelLength
    If iStart > iEnd Then
        iTmp = iEnd
        iEnd = iStart
        iStart = iTmp
    End If
    'iStart = getStartPos(iStart)
    'iEnd = getEndPos(iEnd)
    .SelStart = iStart
    .SelLength = iEnd - iStart
        
    strSel = .SelText
    arSel = Split(strSel, vbCrLf)
    For i = 0 To UBound(arSel)
        arSel(i) = "!" & arSel(i)
        iEnd = iEnd + 1
    Next i
    
    .SelText = Join(arSel, vbCrLf)
    .SelStart = iStart
    .SelLength = iEnd - iStart
End With

End Sub



Private Sub Command1_Click()
rtxtBox.Text = ""

Call Form_Load
End Sub

Private Sub Command11_Click()
If TxtChange = True Then
 strMsg = MsgBox("Create analog list ,Do you want to continue ?", 52, "Warning!")
If strMsg = vbYes Then
Call Command2_Click
Unload Me
Exit Sub
ElseIf strMsg = vbNo Then
Unload Me
Exit Sub
End If
Else
Unload Me
End If
End Sub

Private Sub Command2_Click()
Dim strSave As String
strSave = Text1.Text + rtxtBox.Text
On Error GoTo EX
If Text2.Text = "" Then Exit Sub
  AnalogDeviceNameNet = Text2.Text
 strAnalogPath = AnalogDebugPath & "analog\" & Trim(AnalogDeviceNameNet)
Open strAnalogPath For Output As #52
 Print #52, strSave
 Close #52
 MsgBox "Save " & strAnalogPath & " OK"
 TxtChange = False
 Exit Sub
EX:
 MsgBox "Save " & strAnalogPath & " Error"

End Sub





Private Sub Command6_Click()
Call UnComment
Me.rtxtBox.SetFocus
End Sub
Private Sub UnComment()
Dim strSel As String
Dim arSel() As String
Dim i As Integer
Dim iStart As Long, iEnd As Long, iTmp As Long


With Me.rtxtBox
    iStart = .SelStart
    iEnd = .SelStart + .SelLength
    If iStart > iEnd Then
        iTmp = iEnd
        iEnd = iStart
        iStart = iTmp
    End If
    'iStart = getStartPos(iStart)
    'iEnd = getEndPos(iEnd)
    .SelStart = iStart
    .SelLength = iEnd - iStart
        
    strSel = .SelText
    arSel = Split(strSel, vbCrLf)
    For i = 0 To UBound(arSel)
        If left(arSel(i), 1) = "!" Then
            arSel(i) = right(arSel(i), Len(arSel(i)) - 1)
            iEnd = iEnd - 1
        End If
    Next i
    
    .SelText = Join(arSel, vbCrLf)
    .SelStart = iStart
    .SelLength = iEnd - iStart
End With

End Sub

Private Sub Form_Load()
On Error Resume Next
AnalogDebugPath = strNetBoardPath
If Len(AnalogDeviceNameNet) = 0 Then Exit Sub
  
On Error GoTo EX1
Dim MyChar
intChar = 0





 Text2.Text = AnalogDeviceNameNet
 strAnalogPath = AnalogDebugPath & "analog\" & LCase(Trim(AnalogDeviceNameNet))

 On Error GoTo EX1
 strAnalogFile = Dir(strAnalogPath)
 If strAnalogFile = "" Then
 MsgBox strAnalogPath & " not find"
 'Unload Me
 Exit Sub
 Else
 'BACKUP analog
 If Dir(strAnalogPath & ".bak") = "" Then
 FileCopy strAnalogPath, strAnalogPath & ".bak"
 End If
 
 Open strAnalogPath For Input As #51
 Do Until EOF(51)
   Line Input #51, MyChar
   MyCharA(intChar) = MyChar
    intChar = intChar + 1
   If left(MyChar, Len("!!!!    2")) = "!!!!    2" Then
      Text1.Text = MyChar + Chr(13) + Chr(10)
      Else
   rtxtBox.Text = rtxtBox.Text + MyChar + Chr(13) + Chr(10)
   End If
Loop
 Close #51
 End If
 If Text1.Text = "" Then MsgBox "Not analog file!"
 frmAnalogDebug.Caption = strAnalogPath & "      " & intChar & " Line"

Exit Sub

EX1:
MsgBox Err.Description
End Sub


Private Sub Text1_Click()
rtxtBox.SetFocus
End Sub

Private Sub Text1_GotFocus()
rtxtBox.SetFocus
End Sub

Private Sub Text2_Click()
Command1.Enabled = False
cmdInputOk.Enabled = True
End Sub

Private Sub text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 Call cmdInputOk_Click
End If
End Sub






