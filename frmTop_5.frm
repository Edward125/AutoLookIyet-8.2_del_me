VERSION 5.00
Begin VB.Form frmTop_5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Top 5"
   ClientHeight    =   1875
   ClientLeft      =   3030
   ClientTop       =   435
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleLeft       =   1000
   ScaleMode       =   0  'User
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   3015
      Left            =   120
      TabIndex        =   76
      Top             =   1800
      Width           =   7935
      Begin VB.Label lTimes 
         Alignment       =   2  'Center
         Caption         =   "30"
         Height          =   195
         Index           =   0
         Left            =   700
         TabIndex        =   78
         Top             =   300
         Width           =   195
      End
      Begin VB.Label Label6 
         Caption         =   "99"
         Height          =   200
         Index           =   0
         Left            =   1320
         TabIndex        =   77
         Top             =   840
         Width           =   200
      End
      Begin VB.Line Line6 
         Index           =   0
         X1              =   600
         X2              =   700
         Y1              =   400
         Y2              =   400
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808000&
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   1335
         Index           =   4
         Left            =   5640
         Top             =   1080
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000080FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   1335
         Index           =   3
         Left            =   4560
         Top             =   1080
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF00FF&
         FillColor       =   &H00C000C0&
         FillStyle       =   0  'Solid
         Height          =   1335
         Index           =   2
         Left            =   3480
         Top             =   1080
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   1335
         Index           =   1
         Left            =   2400
         Top             =   1080
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   1335
         Index           =   0
         Left            =   1320
         Top             =   1080
         Width           =   375
      End
      Begin VB.Line Line5 
         X1              =   600
         X2              =   6840
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line4 
         X1              =   600
         X2              =   600
         Y1              =   390
         Y2              =   2400
      End
   End
   Begin VB.CommandButton cmdCall 
      Caption         =   "debug"
      Height          =   735
      Left            =   8400
      TabIndex        =   75
      Top             =   720
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Height          =   1815
      Left            =   8040
      TabIndex        =   74
      Top             =   0
      Width           =   7335
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   3960
         Top             =   720
      End
      Begin VB.Timer tmrMove 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2400
         Top             =   720
      End
      Begin VB.Timer tmrCheck 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   3000
         Top             =   720
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5415
      Left            =   360
      TabIndex        =   25
      Top             =   5160
      Width           =   14415
      Begin VB.TextBox txtPartName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtPartName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtPartName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtPartName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   4680
         TabIndex        =   32
         Text            =   "Text2"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtPartName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   6000
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtPartName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   7320
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtPartName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   8640
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtPartName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   9960
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtPartName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   11280
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtPartName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   12600
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF00FF&
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   1215
         Index           =   9
         Left            =   13080
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H80000008&
         FillStyle       =   0  'Solid
         Height          =   1215
         Index           =   8
         Left            =   11760
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   1215
         Index           =   7
         Left            =   10440
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H008080FF&
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   1215
         Index           =   6
         Left            =   9120
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   1215
         Index           =   5
         Left            =   7800
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H00C0C000&
         FillStyle       =   0  'Solid
         Height          =   1215
         Index           =   4
         Left            =   6480
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000040C0&
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   1215
         Index           =   3
         Left            =   5160
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C000C0&
         FillColor       =   &H00C000C0&
         FillStyle       =   0  'Solid
         Height          =   1215
         Index           =   2
         Left            =   3840
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   1905
         Index           =   1
         Left            =   2520
         Top             =   2880
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   0
         Left            =   1200
         Top             =   3720
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   360
         X2              =   360
         Y1              =   200
         Y2              =   4800
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   360
         X2              =   14040
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   495
         X2              =   360
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label lTimesDisplay 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   200
         Index           =   0
         Left            =   1200
         TabIndex        =   36
         Top             =   3480
         Width           =   200
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   120
      TabIndex        =   37
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton cmdDelAll 
         Caption         =   "Del All"
         Height          =   315
         Left            =   6600
         TabIndex        =   79
         Top             =   150
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   3600
         Left            =   720
         Top             =   2040
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   24
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   23
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   22
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   21
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   20
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   19
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   18
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   17
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   16
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   15
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   14
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   13
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   12
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   11
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   10
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   9
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   8
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   7
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   6
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   5
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   4
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   3
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   2
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdDebug 
         Appearance      =   0  'Flat
         Caption         =   "Team 1 del"
         Height          =   255
         Index           =   0
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdDebug 
         Appearance      =   0  'Flat
         Caption         =   "Team 2 del"
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   42
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdDebug 
         Appearance      =   0  'Flat
         Caption         =   "Team 3 del"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdDebug 
         Appearance      =   0  'Flat
         Caption         =   "Team 4 del"
         Height          =   255
         Index           =   3
         Left            =   6600
         TabIndex        =   40
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdDebug 
         Appearance      =   0  'Flat
         Caption         =   "Team 5 del"
         Height          =   255
         Index           =   4
         Left            =   6600
         TabIndex        =   39
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   975
         Left            =   4920
         TabIndex        =   38
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Times"
         Height          =   255
         Left            =   6000
         TabIndex        =   73
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "TestType"
         Height          =   255
         Left            =   4920
         TabIndex        =   72
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Node/Ibus"
         Height          =   255
         Left            =   3600
         TabIndex        =   71
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Node/Sbus"
         Height          =   255
         Left            =   2040
         TabIndex        =   70
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Device/Node"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   14655
      Begin VB.OptionButton Option1 
         Caption         =   "07:30-08:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   23
         Left            =   13320
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "05:30-07:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   22
         Left            =   12120
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "20:30-21:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "21:30-22:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   20
         Left            =   1320
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "22:30-23:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   19
         Left            =   2520
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "23:30-00:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   18
         Left            =   3720
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "00:30-01:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   17
         Left            =   4920
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "01:30-02:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   16
         Left            =   6120
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "02:30-03:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   15
         Left            =   7320
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "03:30-04:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   14
         Left            =   8520
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "04:30-05:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   13
         Left            =   9720
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "05:30-06:30"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   12
         Left            =   10920
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "18:30-19:30"
         Height          =   255
         Index           =   10
         Left            =   12120
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "17:30-18:30"
         Height          =   255
         Index           =   9
         Left            =   10920
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "16:30-17:30"
         Height          =   255
         Index           =   8
         Left            =   9720
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "15:30-16:30"
         Height          =   255
         Index           =   7
         Left            =   8520
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "14:30-15:30"
         Height          =   255
         Index           =   6
         Left            =   7320
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "13:30-14:30"
         Height          =   255
         Index           =   5
         Left            =   6120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "12:30-13:30"
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "11:30-12:30"
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "10:30-11:30"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "09:30-10:30"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "08:30-09:30"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "19:30-20:30"
         Height          =   255
         Index           =   11
         Left            =   13320
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmTop_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'//Download by http://www.codefans.net
'//本功能代码作者   gvu
'//子类化代码作者   pctgl
'//版权归作者所有
'//转载请保留作者信息

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_EXITSIZEMOVE = &H232
Private Const WM_MOVING = &H216
Private Type RECT
        left As Long
        top As Long
        right As Long
        bottom As Long
End Type
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private WithEvents c_Subclass   As iSubClass_2
Attribute c_Subclass.VB_VarHelpID = -1

Private Const SIZE_SHOW         As Long = 60    '隐藏后留出来的宽度或高度,单位缇
Private Const SHOWHIDE_SPEED    As Long = 30    '(自动显示隐藏速度，单位缇)
'显示标识
'0  自动隐藏
'1  自动显示
Private m_ShowFlag              As Long
'显示方向
'0  向左
'1  向右
'2  向上
Private m_ShowOrient            As Long
'显示速度
Private m_ShowSpeed             As Long
'是否已经启动自动隐藏(为了防止WM_MOVING调整窗口位置)
Private m_MoveEnabled           As Boolean

'//下面是把窗口移动Top=0且Left=0或Right=Screen.Width的时候让窗口高度=屏幕高度
'是否自动调整了大小
Private m_AutoSize              As Boolean
Private m_OldHeight             As Long

Dim strDeviceTestTimes(5)
Dim intTime As Integer
Dim strTimesPath As String


Private Sub cmdDebug_Click(Index As Integer)
Dim tmpInt(5) As Integer
 On Error Resume Next
Erase strDeviceTestTimes
Dim tmpFileName As String
Erase tmpInt
 If Index = 0 Then
    If Trim(Text1(0).Text) = "" Then
      Exit Sub
      Else
        If Text1(3).Text = "[Testjet]" Then
             tmpFileName = Text1(0).Text & "," & Text1(1).Text & "," & Text1(2).Text & "," & Text1(3).Text
        End If
        If Text1(3).Text = "[Open]" Then
             tmpFileName = Text1(0).Text & "," & Text1(1).Text & "," & Text1(3).Text
        End If
        If Text1(3).Text = "[Analog]" Then
            tmpFileName = Text1(0).Text & "," & Text1(3).Text
        End If
       'Kill strTimesPath & "\" & Text1(0).Text & "." & Text1(3).Text & ".zuoai"
          Kill strTimesPath & "\" & tmpFileName & ".zuoai"
          tmpFileName = ""
        For i = 0 To 4
          Text1(i).Text = ""
        Next
       
    End If
 End If
 
  If Index = 1 Then
    If Trim(Text1(5).Text) = "" Then
      Exit Sub
      Else

        If Text1(8).Text = "[Testjet]" Then
             tmpFileName = Text1(5).Text & "," & Text1(6).Text & "," & Text1(7).Text & "," & Text1(8).Text
        End If
        If Text1(8).Text = "[Open]" Then
             tmpFileName = Text1(5).Text & "," & Text1(6).Text & "," & Text1(8).Text
        End If
        If Text1(8).Text = "[Analog]" Then
            tmpFileName = Text1(5).Text & "," & Text1(8).Text
        End If

          Kill strTimesPath & "\" & tmpFileName & ".zuoai"
          tmpFileName = ""

        For i = 5 To 9
          Text1(i).Text = ""
        Next

    End If
 End If
'

  If Index = 2 Then
    If Trim(Text1(10).Text) = "" Then
      Exit Sub
      Else
        If Text1(13).Text = "[Testjet]" Then
             tmpFileName = Text1(10).Text & "," & Text1(11).Text & "," & Text1(12).Text & "," & Text1(13).Text
        End If
        If Text1(13).Text = "[Open]" Then
             tmpFileName = Text1(10).Text & "," & Text1(11).Text & "," & Text1(13).Text
        End If
        If Text1(13).Text = "[Analog]" Then
            tmpFileName = Text1(10).Text & "," & Text1(13).Text
        End If

          Kill strTimesPath & "\" & tmpFileName & ".zuoai"
          tmpFileName = ""




       Kill strTimesPath & "\" & Text1(10).Text & "." & Text1(13).Text & ".zuoai"

        For i = 10 To 14
          Text1(i).Text = ""
        Next

    End If
 End If
 
  If Index = 3 Then
    If Trim(Text1(15).Text) = "" Then
      Exit Sub
      Else
        If Text1(18).Text = "[Testjet]" Then
             tmpFileName = Text1(15).Text & "," & Text1(16).Text & "," & Text1(17).Text & "," & Text1(18).Text
        End If
        If Text1(18).Text = "[Open]" Then
             tmpFileName = Text1(15).Text & "," & Text1(16).Text & "," & Text1(18).Text
        End If
        If Text1(18).Text = "[Analog]" Then
            tmpFileName = Text1(15).Text & "," & Text1(18).Text
        End If
          Kill strTimesPath & "\" & tmpFileName & ".zuoai"
          tmpFileName = ""
        For i = 15 To 19
          Text1(i).Text = ""
        Next

    End If
 End If
 
  If Index = 4 Then
    If Trim(Text1(20).Text) = "" Then
      Exit Sub
      Else

        If Text1(23).Text = "[Testjet]" Then
             tmpFileName = Text1(20).Text & "," & Text1(21).Text & "," & Text1(22).Text & "," & Text1(23).Text
        End If
        If Text1(23).Text = "[Open]" Then
             tmpFileName = Text1(20).Text & "," & Text1(21).Text & "," & Text1(23).Text
        End If
        If Text1(23).Text = "[Analog]" Then
            tmpFileName = Text1(20).Text & "," & Text1(23).Text
        End If
          Kill strTimesPath & "\" & tmpFileName & ".zuoai"
          tmpFileName = ""
        For i = 20 To 24
          Text1(i).Text = ""
        Next

    End If
 End If
  
 
 
 
 
 
 
    T = 0
   j = 0
   
    For i = 1 To 5
        T = T + 4
       tmpInt(i) = Val(Text1(T).Text)
       T = T + 1
       strDeviceTestTimes(i) = Text1(j).Text & "," & Text1(j + 1).Text & "," & Text1(j + 2).Text & "," & Text1(j + 3).Text & "," & Text1(j + 4).Text
       
       j = j + 5
       
    Next
    
'    tmpInt(0) = 20
'    tmpInt(1) = 18
'    tmpInt(2) = 70
'    tmpInt(3) = 68
'    tmpInt(4) = 1
'    tmpInt(5) = 522
'
    Call PaiXu(tmpInt())
T = 0
W = 4

    For i = 5 To 1 Step -1
       
      ' tmpInt(i) = Val(Text1(T).Text)
        
       strFenPei = Split(strDeviceTestTimes(i), ",")
       S = 0
       For Y = T To W
           Text1(Y).Text = strFenPei(S)
            S = S + 1
       Next
       T = T + 5
       W = W + 5
    Next
    
    
 Open strTimesPath & "\TestTime.log" For Output As #7
   For G = 5 To 1 Step -1
      Print #7, strDeviceTestTimes(G)
   Next
 
 
End Sub

Private Sub cmdDelAll_Click()
On Error Resume Next
strTimesPath = "C:\WINDOWS\system\Top10"
        For i = 0 To 24
           Text1(i).Text = ""
        Next
Kill strTimesPath & "\*.*"

End Sub

'Dim strDevicePartNumber(9) As String


Private Sub Command1_Click()
Dim inti(9) As Integer
 
For i = 0 To 9
   inti(i) = i + 7
   Debug.Print inti(i)
Next
Call Map_Updata(inti())
End Sub

Private Sub cmdCall_Click()
         ' Call Top_10(strDeviceName, strDeviceType)
'
''Dim b(5) As Integer
''Dim c(5) As String
''Dim e(5) As String
''For i = 0 To 24
''Text1(i) = i
''Next
'' Call PaiXu(b(), 1)
'
'Dim a As String
'Dim b As String
'Dim C As String
'a = "fuc0k"
'b = "DASDF"
'C = "ASDFASDF"
'Call Top_10(a, a)
'Call Top_10(a, C)
'Call Top_10(b, a)
'Call Top_10(C, b)
'Call Top_10(C, a)
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim strFenPei()  As String

strTimesPath = "C:\WINDOWS\system\Top10"
  If Dir(strTimesPath & "\TestTime.log") <> "" Then
      Open strTimesPath & "\TestTime.log" For Input As #10
       jj = 0
       rr = 4
       ww = 0
        Do Until EOF(10)
            Line Input #10, tmpStr
            tmpStr = Trim(tmpStr)
            strFenPei = Split(tmpStr, ",")
            
            For i = jj To rr
               Text1(i).Text = Trim(strFenPei(ww))
              ww = ww + 1
            Next
            jj = jj + 5
            rr = rr + 5
            ww = 0
            
        Loop
      Close #10
  End If
intTime = 0
   Set c_Subclass = New iSubClass_2
   c_Subclass.SetMsgHook Me.hWnd
   'Me.Top = 100
   tmrCheck.Enabled = True
 '  Timer2.Enabled = True
  ' Call Load_Set
  Call Load_Set_2
End Sub

Private Sub c_Subclass_GetWindowMessage(Result As Long, ByVal cHwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case Message
        Case WM_NCLBUTTONDOWN
            Const HTCAPTION = 2
            If wParam = HTCAPTION Then
                '点击标题栏让所有Timer停止工作
                m_MoveEnabled = True
                tmrCheck.Enabled = False
                tmrMove.Enabled = False
            End If
            
        Case WM_MOVING
            If m_MoveEnabled = False Then Exit Sub
            '这里仅仅是为了不让窗口移出屏幕，可以忽略
            Dim rcMov   As RECT
            Dim rcWnd   As RECT
            Dim lScrW   As Long
            '获取窗口矩形
            Call GetWindowRect(cHwnd, rcWnd)
            '//屏幕宽度
            lScrW = Screen.Width / Screen.TwipsPerPixelX
            '获取移动目标位置矩形
            Call CopyMemory(rcMov, ByVal lParam, Len(rcMov))
            With rcMov
                If .left < 0 Then
                    .left = 0
                    .right = rcWnd.right - rcWnd.left
                End If
                If .top < 0 Then
                    .top = 0
                    .bottom = rcWnd.bottom - rcWnd.top
                End If
                If .right > lScrW Then
                    .left = lScrW - (rcWnd.right - rcWnd.left)
                    .right = .left + (rcWnd.right - rcWnd.left)
                End If
            End With
            '//如果窗口的靠在右上角或左上角，则把高度设置为屏幕高度
            If rcMov.top = 0 And (rcMov.left = 0 Or rcMov.right = Screen.Width / Screen.TwipsPerPixelX) Then
                If m_AutoSize = False Then
                    m_AutoSize = True
                    '保存旧的高度
                    m_OldHeight = rcMov.bottom - rcMov.top
                '    rcMov.Bottom = Screen.Height / Screen.TwipsPerPixelY
                End If
            Else
                If m_AutoSize Then
                    m_AutoSize = False
                    '设置旧的高度
                 '   rcMov.Bottom = rcMov.Top + m_OldHeight
                End If
            End If
            Call CopyMemory(ByVal lParam, rcMov, Len(rcMov))
            
        Case WM_EXITSIZEMOVE
            m_MoveEnabled = False
            Call GetWindowRect(cHwnd, rcWnd)
            If rcWnd.left <= 0 Or rcWnd.top <= 0 Or _
                rcWnd.right >= Screen.Width / Screen.TwipsPerPixelX Then
                '如果窗口停靠在屏幕边缘
                '让检查鼠标位置的Timer工作
                
                '设置显示方向
                If rcWnd.left = 0 Then
                    m_ShowOrient = 0
                ElseIf rcWnd.right >= Screen.Width / Screen.TwipsPerPixelX Then
                    m_ShowOrient = 1
                ElseIf rcWnd.top = 0 Then
                    m_ShowOrient = 2
                End If
                tmrCheck.Enabled = True
            End If
    End Select
    Result = c_Subclass.CallDefaultWindowProc(cHwnd, Message, wParam, lParam)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If unloadTop = False Then
    Cancel = 1
    Me.Hide
    frmAuto1.cmdTop5.Caption = "ShowTop5"
End If
End Sub

Private Sub Picture1_Click()

End Sub

 

Private Sub Text1_Click(Index As Integer)

Dim i As Integer
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)



i = 5
For j = 0 To 4
  Text1(j).BackColor = &HFFC0C0
Next
For j = 10 To 14
  Text1(j).BackColor = &HFFC0C0
Next
For j = 20 To 24
  Text1(j).BackColor = &HFFC0C0
Next
For j = 5 To 9
 Text1(j).BackColor = &HC0FFFF
Next
For j = 15 To 19
 Text1(j).BackColor = &HC0FFFF
Next
Text1(Index).BackColor = &HFF80FF
End Sub

Private Sub Timer2_Timer()
'intTime = intTime + 1
On Error Resume Next
Dim aa
aa = Format("hhmm", Now)
'If intTime = 3600 Then
    
   For i = 0 To 24
     Text1(i).Text = ""
   Next
   Kill strTimesPath & "\*.*"
  ' intTime = 0
'End If
End Sub

Private Sub tmrCheck_Timer()
    Dim pt As POINTAPI
    Dim rc As RECT
    Call GetCursorPos(pt)
    Call GetWindowRect(Me.hWnd, rc)
    If PtInRect(rc, pt.X, pt.Y) Then
        '鼠标停留在窗口上
        If m_ShowFlag = 1 Then Exit Sub
        m_ShowSpeed = SHOWHIDE_SPEED
        m_ShowFlag = 1
        tmrMove.Enabled = True
    Else
        '鼠标不再窗口上
        If m_ShowFlag = 0 Then Exit Sub
        m_ShowSpeed = SHOWHIDE_SPEED
        m_ShowFlag = 0
        tmrMove.Enabled = True
    End If
End Sub

Private Sub tmrMove_Timer()
    Dim nTop    As Long
    Dim nLeft   As Long
    m_ShowSpeed = m_ShowSpeed + SHOWHIDE_SPEED
    '如果大于300T则加快速度
    If m_ShowSpeed > 300 Then m_ShowSpeed = m_ShowSpeed + m_ShowSpeed * 0.2
    Select Case m_ShowOrient
        Case 0  '0  向左
            If m_ShowFlag = 0 Then
                nLeft = Me.left - m_ShowSpeed
                If nLeft < -Me.Width + SIZE_SHOW Then nLeft = -Me.Width + SIZE_SHOW: tmrMove.Enabled = False
            Else
                nLeft = Me.left + m_ShowSpeed
                If nLeft > -SIZE_SHOW Then nLeft = -SIZE_SHOW: tmrMove.Enabled = False
            End If
            Me.left = nLeft
            
        Case 1  '1  向右
            If m_ShowFlag = 0 Then
                nLeft = Me.left + m_ShowSpeed
                If nLeft > Screen.Width - SIZE_SHOW Then nLeft = Screen.Width - SIZE_SHOW: tmrMove.Enabled = False
            Else
                nLeft = Me.left - m_ShowSpeed
                If nLeft < Screen.Width - Me.Width + SIZE_SHOW Then nLeft = Screen.Width - Me.Width + SIZE_SHOW: tmrMove.Enabled = False
            End If
            Me.left = nLeft
            
        Case 2  '2  向上
            If m_ShowFlag = 0 Then
                nTop = Me.top - m_ShowSpeed
                If nTop < -Me.Height + SIZE_SHOW Then nTop = -Me.Height + SIZE_SHOW: tmrMove.Enabled = False
            Else
                nTop = Me.top + m_ShowSpeed
                If nTop > -SIZE_SHOW Then nTop = -SIZE_SHOW: tmrMove.Enabled = False
            End If
            Me.top = nTop
            
    End Select
End Sub

Private Sub Load_Set()
On Error Resume Next
 'h=4500
 'top=300
 Dim intTimes As Integer
 intTimes = 50
 
  
 For i = 1 To 25
     Load lTimes(i)
     lTimes(i).Alignment = 2
     lTimes(i).top = 200 + i * (90 * 2)
     lTimes(i).Caption = intTimes - 2
     intTimes = intTimes - 2
     lTimes(i).Visible = True
     Load Line3(i)
   Line3(i).X1 = 480
   Line3(i).X2 = 345
   Line3(i).Y1 = 300 + i * (90 * 2)
   Line3(i).Y2 = 300 + i * (90 * 2)
   Line3(i).Visible = True

 Next
danGeHeight = 0 * (4500 / 50)
For i = 0 To 9

 Shape1(i).top = 4800
 Shape1(i).top = Shape1(i).top - danGeHeight
Shape1(i).Height = danGeHeight
    If i > 0 Then
           Load lTimesDisplay(i)
        lTimesDisplay(i).left = Shape1(i).left
        lTimesDisplay(i).top = Shape1(i).top - 200
        lTimesDisplay(i).Visible = True
        Else
        lTimesDisplay(i).left = Shape1(i).left
        lTimesDisplay(i).top = Shape1(i).top - 200
    End If


Next

End Sub

Private Sub Map_Updata(intDeviceTestTimes() As Integer)  '(strDevicePartName As String)
'strDevicePartNumber
Dim danGeHeight As Integer
'danGeHeight = 50 * (4500 / 50)
For i = 0 To 9
    danGeHeight = intDeviceTestTimes(i) * (4500 / 50)
 Shape1(i).top = 4800
 Shape1(i).top = Shape1(i).top - danGeHeight
Shape1(i).Height = danGeHeight
 
        lTimesDisplay(i).left = Shape1(i).left
        lTimesDisplay(i).top = Shape1(i).top - 200
        
       lTimesDisplay(i).Caption = intDeviceTestTimes(i)


Next


End Sub


Private Sub Top_10(strDeviceName As String, strDeviceType As String)
   On Error Resume Next
'   Dim CurrentTime
'   strCurrentTime = Format(Now, "hhmm")
'   CurrentTime = Val(Right(strCurrentTime, 2))
'   MsgBox CurrentTime
Dim strTmpPath As String
Dim tmpTimes
Dim tmpInt(5) As Integer
Dim TestTimes As Integer
Dim strFenPei() As String
strTmpPath = strTimesPath
'strTmpPath = "C:\WINDOWS\system\To10"
   MkDir strTmpPath
   strTmpPath = strTimesPath & "\"
   Erase strDeviceTestTimes
   Erase tmpInt
   Erase strFenPei
'strTmpPath = "C:\WINDOWS\system\Top10\"
   If Dir(strTmpPath & strDeviceName & "." & strDeviceType & ".zuoai") <> "" Then
       Open strTmpPath & strDeviceName & "." & strDeviceType & ".zuoai" For Input As #7
          Line Input #7, tmpTimes
       Close #7
        TestTimes = Val(Trim(tmpTimes))
        TestTimes = TestTimes + 1
        Open strTmpPath & strDeviceName & "." & strDeviceType & ".zuoai" For Output As #7
           Print #7, TestTimes
        Close #7
        'TestTimes = 0
     Else
        Open strTmpPath & strDeviceName & "." & strDeviceType & ".zuoai" For Output As #7
           Print #7, "1"
        Close #7
        TestTimes = 1
   End If
   strDeviceTestTimes(0) = strDeviceName & "," & strSbusNode & "," & strIbusNode & "," & strDeviceType & "," & TestTimes
   tmpInt(0) = TestTimes
   T = 0
   j = 0
   
    For i = 1 To 5
        T = T + 4
       tmpInt(i) = Val(Text1(T).Text)
       T = T + 1
       strDeviceTestTimes(i) = Text1(j).Text & "," & Text1(j + 1).Text & "," & Text1(j + 2).Text & "," & Text1(j + 3).Text & "," & Text1(j + 4).Text
       j = j + 5
       
    Next
    
'    tmpInt(0) = 20
'    tmpInt(1) = 18
'    tmpInt(2) = 70
'    tmpInt(3) = 68
'    tmpInt(4) = 1
'    tmpInt(5) = 522
'
    Call PaiXu(tmpInt())
T = 0
W = 4

    For i = 5 To 1 Step -1
       
      ' tmpInt(i) = Val(Text1(T).Text)
        
       strFenPei = Split(strDeviceTestTimes(i), ",")
       S = 0
       For Y = T To W
           Text1(Y).Text = strFenPei(S)
            S = S + 1
       Next
       T = T + 5
       W = W + 5
    Next
    
    
 Open strTmpPath & "TestTime.log" For Output As #7
   For G = 5 To 1 Step -1
      Print #7, strDeviceTestTimes(G)
   Next
 Close #7
    
    
'   If Text1(0).Text = "" Then
'     Text1(0).Text = strDeviceName
'     Text1(3).Text = strDeviceType
'     Text1(4).Text = TestTimes
'   End If
'
'  If Text1(5).Text = "" And strDeviceName <> Text1(0).Text Then
'     Text1(5).Text = strDeviceName
'     Text1(8).Text = strDeviceType
'     Text1(9).Text = TestTimes
'   End If
   
   TestTimes = 0
End Sub
Private Sub PaiXu(intI_1() As Integer) ', intCurrentDevice As Integer)

For i = 0 To 5
      For j = 0 To 4

       If intI_1(j) > intI_1(j + 1) Then
          tmpStr = intI_1(j)
          intI_1(j) = intI_1(j + 1)
          intI_1(j + 1) = tmpStr
          tmpStr2 = strDeviceTestTimes(j)
          strDeviceTestTimes(j) = strDeviceTestTimes(j + 1)
          strDeviceTestTimes(j + 1) = tmpStr2

       End If
      Next j
   Next i


 End Sub

Private Sub Load_Set_2()
On Error Resume Next
Dim bLeft As Boolean
 
 Dim intTimes As Integer
 intTimes = 30
 
  
 For i = 1 To 15
     Load lTimes(i)
     lTimes(i).Alignment = 2
     lTimes(i).top = 300 + i * (67 * 2)
     lTimes(i).Caption = intTimes - 2
     intTimes = intTimes - 2
      If bLeft = False Then
          lTimes(i).left = 360
          bLeft = True

         Else
           lTimes(i).left = 700
           bLeft = False

      End If
     lTimes(i).Visible = True
     Load Line6(i)
   Line6(i).X1 = 600
   Line6(i).X2 = 700
   Line6(i).Y1 = 400 + i * (67 * 2)
   Line6(i).Y2 = 400 + i * (67 * 2)
   Line6(i).Visible = True

 Next
danGeHeight = 20 * (2010 / 30)
For i = 0 To 4

 Shape2(i).top = 2400
 Shape2(i).top = Shape2(i).top - danGeHeight
Shape2(i).Height = danGeHeight
    If i > 0 Then
           Load Label6(i)
        Label6(i).left = Shape2(i).left
        Label6(i).top = Shape2(i).top - 300
        Label6(i).Caption = 20
        Label6(i).Visible = True
        Else
        Label6(i).left = Shape2(i).left
        Label6(i).top = Shape2(i).top - 300
        Label6(i).Caption = 20
    End If


Next

End Sub

