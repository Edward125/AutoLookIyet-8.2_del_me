VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAllPath 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SetPah"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9945
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Local Computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9735
      Begin VB.Frame Frame3 
         Caption         =   "Board Path"
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   9495
         Begin VB.TextBox txtBoard 
            BackColor       =   &H00FFC0C0&
            Height          =   405
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   8175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Browse..."
            Height          =   375
            Left            =   8400
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Iyet Path"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9495
         Begin VB.TextBox txtReport 
            BackColor       =   &H00FFC0C0&
            Height          =   405
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   8175
         End
         Begin VB.CommandButton cmdOpenReport 
            Caption         =   "Open..."
            Height          =   375
            Left            =   8400
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      Caption         =   "NetWork Computer"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   9735
      Begin VB.Frame Frame2 
         Caption         =   "Net Iyet Path"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   9495
         Begin VB.TextBox txtNetReport 
            BackColor       =   &H00C0FFFF&
            Height          =   405
            Left            =   120
            TabIndex        =   15
            Text            =   "C:\A.EEE"
            Top             =   240
            Width           =   8175
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Open..."
            Height          =   375
            Left            =   8400
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Net Board Path"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   9495
         Begin VB.TextBox txtNetBoard 
            BackColor       =   &H00C0FFFF&
            Height          =   405
            Left            =   120
            TabIndex        =   12
            Text            =   "C:\"
            Top             =   240
            Width           =   8175
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Browse..."
            Height          =   375
            Left            =   8400
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   1440
      Picture         =   "frmAllPath.frx":0000
      Top             =   2280
      Width           =   4155
   End
End
Attribute VB_Name = "frmAllPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmAuto1.Show
 bFrmClose = False
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim MyStr As String
On Error GoTo EX
   If txtReport.Text = "" Then
      MsgBox "please set IYET report file"
      txtReport.SetFocus
      Exit Sub
   End If
   If txtBoard.Text = "" Then
      MsgBox "please set board path!"
      txtBoard.SetFocus
      Exit Sub
   End If
   
'   If txtNetReport.Text = "" Then
'      MsgBox "please set network iyet file"
'      txtNetReport.SetFocus
'      Exit Sub
'   End If
'   If txtNetBoard.Text = "" Then
'      MsgBox "please set network board path!"
'      txtNetBoard.SetFocus
'      Exit Sub
'   End If

       Open strToolPath & "AutoLookLog\Path.ini" For Output As #10
           MyStr = Trim(txtReport.Text)
          Print #10, "#IyetPath#:" & MyStr
'           MyStr = Trim(txtNetReport.Text)
'          Print #10, "#NetIyetPath#:" & MyStr
           MyStr = Trim(txtBoard.Text)
          If right(MyStr, 1) <> "\" Then MyStr = MyStr & "\"

          Print #10, "#BoardPath#:" & MyStr
            MyStr = Trim(txtNetBoard.Text)
            If right(MyStr, 1) <> "\" Then MyStr = MyStr & "\"
'          Print #10, "#NetBoardPath#:" & MyStr
          Print #10, "#Name#:" & strName
'          Print #10, "#NetWorkName#:" & strNetName
       Close #10
        'frmAuto1.Caption = UCase(strName)
        
 On Error Resume Next
strTimesPath = "C:\WINDOWS\system\Top10"
        For i = 0 To 24
           frmTop_5.Text1(i).Text = ""
        Next
Kill strTimesPath & "\*.*"
        
        
        
       MsgBox "Set OK!"
Call cmdCancel_Click
       
 Exit Sub
EX:
End Sub

Private Sub cmdOpenReport_Click()
On Error GoTo EX

  Me.CommonDialog1.Filter = "*.eee|*.eee|*.txt|*.txt|*.*|*.*"
  Me.CommonDialog1.CancelError = True
  Me.CommonDialog1.ShowOpen
   If CommonDialog1.FileName = "" Then Exit Sub
   If LCase(CommonDialog1.FileTitle) = "failure.txt" Then: MsgBox "The failure.txt is iyet test retry file ,please open failure.eee file !", vbCritical: Exit Sub
  tmpstrReportPath = strReportPath
  
  strReportPath = CommonDialog1.FileName
  If strReportPath <> tmpstrReportPath Then
      On Error Resume Next
      strTimesPath = "C:\WINDOWS\system\Top10"
            Kill strTimesPath & "\*.*"
  End If
  
  txtReport.Text = CommonDialog1.FileName
  txtBoard.Text = Replace(LCase(txtReport.Text), "iyet\failure.txt.eee", "")
  Exit Sub
EX:
End Sub

Private Sub Text1_Change()
Call cmdOpenReport_Click
End Sub

Private Sub Command1_Click()
On Error GoTo EX

   Me.CommonDialog1.Filter = "*.eee|*.eee|*.txt|*.txt|*.*|*.*"
  Me.CommonDialog1.CancelError = True
  Me.CommonDialog1.ShowOpen
   If CommonDialog1.FileName = "" Then Exit Sub
      If LCase(CommonDialog1.FileTitle) = "failure.txt" Then: MsgBox "The failure.txt is iyet test retry file ,please open failure.eee file !", vbCritical: Exit Sub

  strNetReportPath = CommonDialog1.FileName
  txtNetReport.Text = CommonDialog1.FileName
  
  txtNetBoard.Text = Replace(LCase(txtNetReport.Text), "iyet\failure.txt.eee", "")
  Exit Sub
EX:
End Sub


Private Sub Command2_Click()
    Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer

  bi.hOwner = Me.hwnd
 '指向相对于浏览器中的“根”文件夹的位置的元素标志列表，
 '如果本参数为NULL，则为桌面文件夹
  bi.pidlRoot = 0&
 '显示在浏览对话框中的消息
  bi.lpszTitle = "please select directory"
 '要返回的文件夹类型
  bi.ulFlags = BIF_RETURNONLYFSDIRS
 '显示浏览对话框
  pidl = SHBrowseForFolder(bi)
 '退出了浏览对话框，解析、显示用户选择的文件夹
  path = Space$(MAX_PATH)
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     txtBoard.Text = left(path, pos - 1)
  End If
  Call CoTaskMemFree(pidl)
End Sub

Private Sub Command3_Click()
    Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer

  bi.hOwner = Me.hwnd
 '指向相对于浏览器中的“根”文件夹的位置的元素标志列表，
 '如果本参数为NULL，则为桌面文件夹
  bi.pidlRoot = 0&
 '显示在浏览对话框中的消息
  bi.lpszTitle = "please select directory"
 '要返回的文件夹类型
  bi.ulFlags = BIF_RETURNONLYFSDIRS
 '显示浏览对话框
  pidl = SHBrowseForFolder(bi)
 '退出了浏览对话框，解析、显示用户选择的文件夹
  path = Space$(MAX_PATH)
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     txtNetBoard = left(path, pos - 1)
  End If
  Call CoTaskMemFree(pidl)

End Sub

Private Sub Form_Load()
On Error GoTo EX
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
 txtReport.Text = strReportPath
'txtNetReport = strNetReportPath
   txtBoard = strBoardPath
'    txtNetBoard = strNetBoardPath
EX:
End Sub



Private Sub Form_Unload(Cancel As Integer)
frmAuto1.Show
End Sub

Private Sub txtBoard_DblClick()
Call Command2_Click
End Sub

Private Sub txtNetBoard_DblClick()
Call Command3_Click
End Sub



Private Sub txtNetReport_DblClick()
Call Command1_Click
End Sub



Private Sub txtReport_DblClick()
Call cmdOpenReport_Click
End Sub
