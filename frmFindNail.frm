VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Nail"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblEg 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "3's"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "2's"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "1's"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  SendKeys "{Tab}"
End If

End Sub

Private Sub Text1_GotFocus()
    Me.Text1.SelStart = Len(Me.Text1.Text)
End Sub

Private Sub Text2_GotFocus()
    Me.Text2.SelStart = Len(Me.Text2.Text)
End Sub

Private Sub Text3_GotFocus()
    Me.Text3.SelStart = Len(Me.Text3.Text)
End Sub
