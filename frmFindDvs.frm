VERSION 5.00
Begin VB.Form frmFindDvs 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Device"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4200
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Canecl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3975
      Begin VB.TextBox text3 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox text2 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox text1 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   0
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000012&
         Caption         =   "Device(3)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Device(2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Device(1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmFindDvs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'strFindDeviceName1 = Trim(text1.Text)
'strFindDeviceName1 = Trim(text2.Text)
'strFindDeviceName1 = Trim(text3.Text)
Me.Hide
End Sub

Private Sub Command2_Click()
Unload Me
text1.Text = ""
text2.Text = ""
text3.Text = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  SendKeys "{Tab}"
End If

End Sub

Private Sub Text1_GotFocus()
  '  Me.text1.SelStart = Len(Me.text1.Text)
End Sub

Private Sub text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   If text1.Text = "" Then
      text1.SetFocus
      Exit Sub
     Else
      text2.SetFocus
   End If

End If
End Sub

Private Sub Text2_GotFocus()
    'Me.text2.SelStart = Len(Me.text2.Text)
End Sub

Private Sub text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

      text3.SetFocus


End If

End Sub

Private Sub Text3_GotFocus()
    'Me.text3.SelStart = Len(Me.text3.Text)
End Sub

Private Sub text3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   If text1.Text = "" Then
      text1.SetFocus
      Exit Sub
     Else
      Call Command1_Click
   End If

End If
End Sub
