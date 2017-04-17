VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBoardView 
   AutoRedraw      =   -1  'True
   Caption         =   "No Image"
   ClientHeight    =   5040
   ClientLeft      =   480
   ClientTop       =   5685
   ClientWidth     =   6180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6180
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   0
      Width           =   1095
   End
   Begin VB.CheckBox chkEnableRightMouse 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkShowPartName 
      Caption         =   "&Show part Name"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picNavigate 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      FillColor       =   &H0000FFFF&
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   240
      ScaleHeight     =   1275
      ScaleWidth      =   1755
      TabIndex        =   4
      ToolTipText     =   "Press ""F3"" to show or hide it "
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar stbBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4785
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10398
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5355
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   240
      Width           =   7935
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4200
         Top             =   1560
      End
   End
   Begin VB.Label Label1 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lblTB 
      Caption         =   "No image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilebar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu tt 
      Caption         =   "Option..."
      Enabled         =   0   'False
      Begin VB.Menu mnuFindComponet 
         Caption         =   "&Componet(Part)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFindNail 
         Caption         =   "N&ail"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFindPin 
         Caption         =   "&Pin"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFindNet 
         Caption         =   "&Net"
         Enabled         =   0   'False
      End
      Begin VB.Menu chkEnableRightMouse1 
         Caption         =   "&Enable Right Mouse Function"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Zoom__ 
      Caption         =   "Zoom"
      Enabled         =   0   'False
      Begin VB.Menu mnuZoomin 
         Caption         =   "Zoom&In"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuZoomout 
         Caption         =   "Zoom&Out"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuZoomPerfect 
         Caption         =   "&ZoomPerfect"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTurn 
         Caption         =   "&Turn"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRotate 
         Caption         =   "&Rotate"
         Enabled         =   0   'False
      End
      Begin VB.Menu MrrorTurn180 
         Caption         =   "&MrrorTurn180"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About PCBView..."
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmBoardView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Type pPoint
  x As Double
  y As Double
  R As Double
  Group As Integer
End Type

Private Type pPin
  x As Double
  y As Double
  Number As Integer
  Name As String
  Part As String
  Layer As Integer
  NetName As String
  Color As Integer
End Type

Private Type pPart
  X1 As Double
  Y1 As Double
  X2 As Double
  Y2 As Double
  Name As String
  Color As Integer
End Type

Private Type pNail
  x As Double
  y As Double
  Nail As String
  Type As String
  Grid As String
  NetName As String
  Color As Integer
End Type


Private sngXshift As Single
Private sngYshift As Single

'动态存放各零件的信息特别是XY，将不停地随图像而变动
Private OutLineP() As pPoint
Private PinsB() As pPin
Private PinsT() As pPin
Private PartB() As pPart
Private PartT() As pPart
Private NailB() As pNail
Private NailT() As pNail

'Private fNailDetail As frmNailDetail  '用来显示点Nail时的详细信息的一个窗体。

Private imgNavigate As StdPicture   '用来存放导航图片
Private blnTopimg As Boolean   '是否导航图是存的top面的.
Private blnRedrawNvigate As Boolean '用来标志是否需要重画导航图

Private pCenter As pPoint     '记录板子的中心坐标(变动中)
Private dblWidth_Original As Double    '记录板子的原始宽
Private dblHeight_Original As Double
Private dblWidth_Board As Double    '记录板子的适应窗口后的宽
Private dblHeight_Board As Double
Private blnHorizontal As Boolean   '记录板子是否为横放,也就是一开始加载时的方式 没有旋转过.

Private MouseX As Single       '存放随时鼠标坐标
Private MouseY As Single
Private MouseDownX As Single   '存放MouseDown 时鼠标的坐标
Private MouseDownY As Single

Private blnTop As Boolean '用来存当前是不是top，True = top ;false=Bottom

Private lngRate As Long  '放大比率   初始化时给定
Private intCircleRate As Integer  '设定画圆时半径与放大比之间的比率   初始化时给定
Private sngSpace As Single '设定点到Part边框的空白处大小的比例   初始化时给定
Private lngFindRate  As Long   '设定当找一个东西时找到后所要放大的倍数   初始化时给定

'Nail 的颜色和Pin 的颜色一样用。
Private intPinColorNormal As Integer  '设定pin的颜色  初始化时给定
Private intPinColorSel   As Integer  '设定pin被选中时的颜色  初始化时给定
Private intPinColorFind(3) As Integer  '设定找到时的3种颜色 初始化时给定

Private intPartColorNormal As Integer
Private intPartColorSel As Integer

Private imgPic As New StdPicture '框选时使用，框选时先把图片保存起来，设为背景，再对picture1操作，OK。
Private intMouseKey As Integer  '存放MouseDown时的是左键还是右键。
Private blnBoxSelect As Boolean   '标志当前是不是处在框选状态下。
Private DBclikc As Boolean      'show whether mouseup is double click

Private str_ColorPinT As String  '存放有色的Pin的号码，格式 "1 3 4 " 表示1，3，4 是有色的Pin
Private str_ColorPinB As String
Private str_ColorPartT As String '存放有色的Part的号码，格式同上
Private str_ColorPartB As String
Private str_ColorNailT As String '存放有色的Nail的号码，格式同上
Private str_ColorNailB As String

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  
Dim strCommand As String
Dim DownFkey As Boolean
Dim DownDkey As Boolean


Private Sub EnableCtls()
On Error Resume Next
  Dim ctl As Control
  For Each ctl In Me.Controls
   ctl.Enabled = True
  Next ctl
End Sub

Private Sub chkEnableRightMouse_Click()
 Me.Picture1.SetFocus
End Sub

Private Sub chkShowPartName_Click()
 Call drawPic
 Me.Picture1.SetFocus
 
End Sub

 
 
Private Sub Command1_Click()
Timer1.Enabled = Not Timer1.Enabled
If Timer1.Enabled = True Then
 
    Command1.Caption = "Stop"
    tt.Enabled = False
    BoardViewTrue = True
    Zoom__.Enabled = False

  Else

    Command1.Caption = "Start"
    tt.Enabled = True
    Zoom__.Enabled = True
    
    BoardViewDevice = ""
    BoardViewTrue = False
    
End If

End Sub

Private Sub Form_Activate()
If Dir(strCommand) <> "" And strCommand <> "" Then
  Call LoadData(strCommand)
 
  strCommand = ""
tt.Enabled = False
Zoom__.Enabled = False
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Here is short_Cut key

If Me.mnuTurn.Enabled = False Then Exit Sub
'If Command1.Caption = "Stop" Then Exit Sub
Select Case KeyCode
  Case Is = 39 ' right key
     Call MovePic(-1 * Me.Picture1.Width / 10, 0)
  Case Is = 37 'left
     Call MovePic(Me.Picture1.Width / 10, 0)
  Case Is = 40  'down
     Call MovePic(0, -1 * Me.Picture1.Height / 10)
  Case Is = 38   'up
     Call MovePic(0, Me.Picture1.Height / 10)
  Case Is = vbKeyF 'f great add
        Call MrrorTurn180_Click
  Case Is = vbKeyD
      DownDkey = True
      Call mnuFindPin_Click 'great
      DownDkey = False
  Case Is = 187  '"+"
     'Call Picture1_DblClick  'Great guo
     'Call ZoomIn
      Call Zoom(2, MouseX, MouseY)
      DBclikc = True
      Call drawPic
  Case Is = 189 ' "-"
       Call Zoom(0.5, MouseX, MouseY)
       DBclikc = True
       Call drawPic
   '  Call ZoomOut
  Case Is = 107  '"+"'great
      Call Zoom(2, MouseX, MouseY)
      DBclikc = True
      Call drawPic

  Case Is = 109 ' "-"'great
       Call Zoom(0.5, MouseX, MouseY)
       DBclikc = True
       Call drawPic
       
       
  Case Is = vbKeyS 'great
     Call mnuFindNail_Click
  Case Is = vbKeySpace  'turn
     Call mnuTurn_Click
  Case Is = 80 ' Asc("P"), Asc("p")
      Call mnuFindPin_Click
  Case Is = 67 ' Asc("C"), Asc("c")
        
        Call mnuFindComponet_Click 'great
  Case Is = 78 ' Asc("N"), Asc("n")
     Call mnuFindNet_Click
  Case Is = 65 ' Asc("A"), Asc("a")
     Call mnuFindNail_Click
  Case Is = 90 ' Asc("Z"), Asc("z")
     Call mnuZoomPerfect_Click
  Case Is = 82 'Asc("R"), Asc("r")
    Call mnuRotate_Click
  Case Is = 114   'F3
     Me.picNavigate.Visible = Not Me.picNavigate.Visible
  Case Is = 27   'ESC
     Call ClearPinNailColor
     Call ClearPartColor
     Call drawPic
 End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
BoardViewTrue = False
Me.CommonDialog1.FileName = ""
BoardViewDevice = ""
frmAuto1.BoardView_.Caption = "BoardViewFalse"
  frmAuto1.Show
End Sub

Private Sub mnuFindComponet_Click()
  Call FindPart_Name
  
End Sub

Private Sub mnuFindNail_Click()
  Call FindNail_Name
End Sub

Private Sub mnuFindNet_Click()
  Call FindNet_Name
  
End Sub

Private Sub mnuFindPin_Click()
  Call FindPin_Name
End Sub




Private Sub mnuFileExit_Click()
  'unload the form
  Unload Me
End Sub
 
Private Sub mnuFileOpen_Click()
  Dim strFile As String
  On Error GoTo errH
  Me.CommonDialog1.Filter = "*.bv2|*.bv2|*.*|*.*"
  Me.CommonDialog1.CancelError = True
  Me.CommonDialog1.ShowOpen
  If Dir(Me.CommonDialog1.FileName) = "" Then
     MsgBox "Can't find the File", vbCritical
  Else
     
    Call LoadData(Me.CommonDialog1.FileName)
     Timer1.Enabled = True
  End If
  
errH:
   Screen.MousePointer = vbNormal
End Sub

Private Sub LoadData(Optional strFilename As String = "")
' Dim cn As New ADODB.Connection
' Dim rs As New ADODB.Recordset
' Dim X1 As Double        '存放最左上角的那个点的坐标，为第一次平移时用，load时就把图片放好。
' Dim Y1 As Double
Dim xMin  As Double, xMax As Double 'X方向的最大最小
Dim yMin As Double, yMax As Double 'Y方向的最大最小
Dim iOutline As Integer, iTopPins As Integer, iBotPins As Integer '总个数
Dim iTopNails As Integer, iBotNails As Integer

Dim strLine As String '每行的内容
Dim ar() As String '每行的分解内容
Dim strReading As String '当前读文件读到哪里了，是#Layout#,#Nail#,#Pin#
Dim j As Integer

Dim blnLoadMax As Boolean

Dim iFreeFile As Integer

 Screen.MousePointer = vbHourglass
  If Dir(strFilename) = "" Or strFilename = "" Then MsgBox "Can't find the file " & strFilename, vbCritical
  
 Dim i As Integer
' lngRate = 250
 
 ReDim OutLineP(0)
 ReDim PinsT(0)
 ReDim PinsB(0)
 ReDim NailB(0)
 ReDim NailT(0)
 ReDim PartT(0)
 ReDim PartB(0)
 
 blnTop = True
 Me.lblTB.Caption = "TOP"
 Me.Caption = strFilename

 blnLoadMax = False
 
 'pre read find out the total information
 iFreeFile = FreeFile
 Open strFilename For Input As #iFreeFile
    Do Until EOF(iFreeFile)
        Line Input #iFreeFile, strLine
        
        If strLine = "#Layout#" Then
            strReading = "#Layout#"
            Line Input #iFreeFile, strLine
            Line Input #iFreeFile, strLine
            
            ElseIf strLine = "#Nail#" Then
                 strReading = "#Nail#"
                 Line Input #iFreeFile, strLine
                 Line Input #iFreeFile, strLine
                 
                 ElseIf strLine = "#Pin#" Then
                    strReading = "#Pin#"
                    Line Input #iFreeFile, strLine
                    Line Input #iFreeFile, strLine
        End If
        
        If strReading = "#Layout#" And strLine <> "" Then
            iOutline = iOutline + 1
            'add to OutLineP
            ar = Split(strLine, ",")
            
            ReDim Preserve OutLineP(UBound(OutLineP) + 1)
            
            OutLineP(UBound(OutLineP) - 1).x = ar(0)
            OutLineP(UBound(OutLineP) - 1).y = -ar(1)
            OutLineP(UBound(OutLineP) - 1).R = ar(2)
            OutLineP(UBound(OutLineP) - 1).Group = ar(3)
            
            If blnLoadMax = False Then
                xMax = ar(0)
                xMin = ar(0)
                yMax = -ar(1)
                yMin = -ar(1)
                blnLoadMax = True
            Else
                If ar(0) > xMax Then xMax = ar(0)
                If ar(0) < xMin Then xMin = ar(0)
                If -ar(1) > yMax Then yMax = -ar(1)
                If -ar(1) < yMin Then yMin = -ar(1)
            End If
            
            ElseIf strReading = "#Nail#" Then
                ar = Split(strLine, ",")
                If ar(5) = "(B)" Then
                    ReDim Preserve NailB(UBound(NailB) + 1)
                    NailB(UBound(NailB) - 1).Nail = ar(0)
                    NailB(UBound(NailB) - 1).x = ar(1)
                    NailB(UBound(NailB) - 1).y = -ar(2)
                    NailB(UBound(NailB) - 1).Type = ar(3)
                    NailB(UBound(NailB) - 1).Grid = ar(4)
'                    NailB(UBound(NailB)).Net = ar(6)
                    NailB(UBound(NailB) - 1).NetName = ar(7)
                    NailB(UBound(NailB) - 1).Color = intPinColorNormal
                     
                Else
                    ReDim Preserve NailT(UBound(NailT) + 1)
                    NailT(UBound(NailT) - 1).Nail = ar(0)
                    NailT(UBound(NailT) - 1).x = ar(1)
                    NailT(UBound(NailT) - 1).y = -ar(2)
                    NailT(UBound(NailT) - 1).Type = ar(3)
                    NailT(UBound(NailT) - 1).Grid = ar(4)
'                    NailT(UBound(NailT)).Net = ar(6)
                    NailT(UBound(NailT) - 1).NetName = ar(7)
                    NailT(UBound(NailT) - 1).Color = intPinColorNormal
                End If
                ElseIf strReading = "#Pin#" Then
                    ar = Split(strLine, ",")
                    If ar(1) = "(B)" Then
                        ReDim Preserve PinsB(UBound(PinsB) + 1)
                        PinsB(UBound(PinsB) - 1).Part = ar(0)
                        PinsB(UBound(PinsB) - 1).Number = ar(2)
                        PinsB(UBound(PinsB) - 1).Name = ar(3)
                        PinsB(UBound(PinsB) - 1).x = ar(4)
                        PinsB(UBound(PinsB) - 1).y = -ar(5)
                        PinsB(UBound(PinsB) - 1).Layer = ar(6)
                        PinsB(UBound(PinsB) - 1).NetName = ar(7)
                        PinsB(UBound(PinsB) - 1).Color = intPinColorNormal
                    Else
                        ReDim Preserve PinsT(UBound(PinsT) + 1)
                        PinsT(UBound(PinsT) - 1).Part = ar(0)
                        PinsT(UBound(PinsT) - 1).Number = ar(2)
                        PinsT(UBound(PinsT) - 1).Name = ar(3)
                        PinsT(UBound(PinsT) - 1).x = ar(4)
                        PinsT(UBound(PinsT) - 1).y = -ar(5)
                        PinsT(UBound(PinsT) - 1).Layer = ar(6)
                        PinsT(UBound(PinsT) - 1).NetName = ar(7)
                        PinsT(UBound(PinsT) - 1).Color = intPinColorNormal
                    
                    End If
                    If blnLoadMax = False Then
                        xMax = ar(4)
                        xMin = ar(4)
                        yMax = -ar(5)
                        yMin = -ar(5)
                        blnLoadMax = True
                    Else
'                        If ar(4) > xMax Then xMax = ar(4)
'                        If ar(4) < xMin Then xMin = ar(4)
'                        If ar(5) > yMax Then yMax = ar(5)
'                        If ar(5) < yMin Then yMin = ar(5)
                    End If
                    
        End If
        DoEvents
        
    Loop
  
 Close #iFreeFile
  
'load top  part
 For i = 0 To UBound(PinsT)
    
    For j = 0 To UBound(PartT)
        If PartT(j).Name = PinsT(i).Part Then
            Exit For
        End If
    Next
'    Debug.Assert PinsT(i).Part <> "U7"
    
    If j > UBound(PartT) Then
        ReDim Preserve PartT(UBound(PartT) + 1)
        j = UBound(PartT) - 1
        PartT(j).Name = PinsT(i).Part
        PartT(j).X1 = PinsT(i).x
        PartT(j).Y1 = PinsT(i).y
        PartT(j).X2 = PinsT(i).x
        PartT(j).Y2 = PinsT(i).y
        PartT(j).Color = intPartColorNormal

    End If
     
'   PartT(i).X1 = rs.Fields("X1") * lngRate - lngRate * sngSpace / intCircleRate
    If PartT(j).X1 >= PinsT(i).x Then
        PartT(j).X1 = PinsT(i).x
    End If
    If PartT(j).Y1 >= PinsT(i).y Then
        PartT(j).Y1 = PinsT(i).y
    End If
    If PartT(j).X2 <= PinsT(i).x Then
        PartT(j).X2 = PinsT(i).x
    End If
    If PartT(j).Y2 <= PinsT(i).y Then
        PartT(j).Y2 = PinsT(i).y
    End If
 Next
 'set part outline ,top
 For i = 0 To UBound(PartT)
    PartT(i).X1 = PartT(i).X1 - sngSpace / intCircleRate
    PartT(i).Y1 = PartT(i).Y1 - sngSpace / intCircleRate
    PartT(i).X2 = PartT(i).X2 + sngSpace / intCircleRate
    PartT(i).Y2 = PartT(i).Y2 + sngSpace / intCircleRate
 Next
 
 'load bottom part
 For i = 0 To UBound(PinsB)
    
    For j = 0 To UBound(PartB)
        If PartB(j).Name = PinsB(i).Part Then
            Exit For
        End If
    Next
    If j > UBound(PartB) Then
        ReDim Preserve PartB(UBound(PartB) + 1)
        j = UBound(PartB) - 1
        PartB(j).Name = PinsB(i).Part
        PartB(j).X1 = PinsB(i).x
        PartB(j).Y1 = PinsB(i).y
        PartB(j).X2 = PinsB(i).x
        PartB(j).Y2 = PinsB(i).y
        PartB(j).Color = intPartColorNormal

    End If
     
'   PartT(i).X1 = rs.Fields("X1") * lngRate - lngRate * sngSpace / intCircleRate
    If PartB(j).X1 >= PinsB(i).x Then
        PartB(j).X1 = PinsB(i).x
    End If
    If PartB(j).Y1 >= PinsB(i).y Then
        PartB(j).Y1 = PinsB(i).y
    End If
    
    If PartB(j).X2 <= PinsB(i).x Then
        PartB(j).X2 = PinsB(i).x
    End If
    If PartB(j).Y2 <= PinsB(i).y Then
        PartB(j).Y2 = PinsB(i).y
    End If
     
 Next
 'set part outline ,bot
 For i = 0 To UBound(PartB)
    PartB(i).X1 = PartB(i).X1 - sngSpace / intCircleRate
    PartB(i).Y1 = PartB(i).Y1 - sngSpace / intCircleRate
    PartB(i).X2 = PartB(i).X2 + sngSpace / intCircleRate
    PartB(i).Y2 = PartB(i).Y2 + sngSpace / intCircleRate
 Next
 
 'load OutLine

 
 'Load Pins
 'load Top pins
   
  'load Bottom Pins
'
'    For i = 0 To rs.RecordCount - 1
'      PinsB(i).x = rs.Fields("X") * lngRate
'      PinsB(i).y = -1 * rs.Fields("Y") * lngRate
'      PinsB(i).Number = rs.Fields("Pin")
'      PinsB(i).Layer = rs.Fields("Layer")
'      PinsB(i).Name = rs.Fields("Name")
'      PinsB(i).NetName = rs.Fields("Net")
'      PinsB(i).Part = rs.Fields("Part")
'      PinsB(i).Color = intPinColorNormal
'      rs.MoveNext
'    Next i
'   End If
   
'load Part
  'load top parts
'  If rs.State <> adStateClosed Then rs.Close
'   rs.open " SELECT part as name,  min(x) AS X1, max(x) AS X2, min(-y) AS Y1, max(-y) AS Y2 From pin where tb='(T)' GROUP BY part ", cn
'   If rs.EOF = True Then
'       MsgBox "No fond the Top Part ", vbCritical
'    Else
'        For i = 0 To rs.RecordCount - 1
'
'          PartT(i).X1 = rs.Fields("X1") * lngRate - lngRate * sngSpace / intCircleRate
'          PartT(i).X2 = rs.Fields("X2") * lngRate + lngRate * sngSpace / intCircleRate
'          PartT(i).Y1 = rs.Fields("Y1") * lngRate - lngRate * sngSpace / intCircleRate
'          PartT(i).Y2 = rs.Fields("Y2") * lngRate + lngRate * sngSpace / intCircleRate
'          PartT(i).Name = rs.Fields("Name")
'          PartT(i).Color = intPartColorNormal
'
'          rs.MoveNext
'        Next i
'   End If
   
    
  'load Nails
     'load bottom Nails
'  If rs.State <> adStateClosed Then rs.Close
'   rs.open "select * from Nail where tb='(B)'", cn
'   If rs.EOF = True Then
'        MsgBox "No fond the Bottom nail", vbCritical
'    Else
'
'        For i = 0 To rs.RecordCount - 1
'          NailB(i).x = rs.Fields("X") * lngRate
'          NailB(i).y = -1 * rs.Fields("Y") * lngRate
'          NailB(i).Grid = rs.Fields("Grid")
'          NailB(i).Nail = rs.Fields("Nail")
'          NailB(i).Net = rs.Fields("NetName")
'          NailB(i).Type = rs.Fields("Type")
'          NailB(i).Color = intPinColorNormal
'          rs.MoveNext
'        Next i
'   End If
     

' 'Load width ,height to calculate lngRate
   
   
    dblWidth_Original = xMax - xMin ' rs.Fields("Xmax") - rs.Fields("xmin")
    dblHeight_Original = yMax - yMin ' rs.Fields("Ymax") - rs.Fields("ymin")
    pCenter.x = (xMin + xMax) / 2
    pCenter.y = (yMin + yMax) / 2
    
    If Picture1.Width / dblWidth_Original < Picture1.Height / dblHeight_Original Then
       lngRate = Picture1.Width / dblWidth_Original
    Else
       lngRate = Picture1.Height / dblHeight_Original
    End If
    
   '使他们的比例一致，要不然会使导航图出错。
'        dblWidth_Original = Picture1.Width / lngRate
'        dblHeight_Original = Picture1.Height / lngRate
        
        dblWidth_Board = Picture1.Width / lngRate
        dblHeight_Board = Picture1.Height / lngRate
        
    lngRate = lngRate * 0.9   '不能把全屏都盖住，留点边 *
' 'load X1 Y1 为平移
'   If rs.State <> adStateClosed Then rs.Close
'   rs.Open "select min(x) as x1,min(y) as y1 from layout", cn
'   X1 = rs.Fields("x1") * 0.95 * lngRate    '左边也要留出一点框
'   Y1 = rs.Fields("y1") * 0.95 * lngRate    '上边也要留出一点框
 
 pCenter.x = pCenter.x * lngRate
 pCenter.y = pCenter.y * lngRate
 
 
 For i = 0 To UBound(OutLineP)
    OutLineP(i).x = OutLineP(i).x * lngRate
    OutLineP(i).y = OutLineP(i).y * lngRate
 Next
 
 For i = 0 To UBound(PinsT)
    PinsT(i).x = PinsT(i).x * lngRate
    PinsT(i).y = PinsT(i).y * lngRate
 Next
 
 For i = 0 To UBound(PinsB)
    PinsB(i).x = PinsB(i).x * lngRate
    PinsB(i).y = PinsB(i).y * lngRate
 Next
 
 For i = 0 To UBound(NailB)
    NailB(i).x = NailB(i).x * lngRate
    NailB(i).y = NailB(i).y * lngRate
 Next
 
 For i = 0 To UBound(NailT)
    NailT(i).x = NailT(i).x * lngRate
    NailT(i).y = NailT(i).y * lngRate
 Next
 
 
 For i = 0 To UBound(PartT)
    PartT(i).X1 = PartT(i).X1 * lngRate
    PartT(i).X2 = PartT(i).X2 * lngRate
    PartT(i).Y1 = PartT(i).Y1 * lngRate
    PartT(i).Y2 = PartT(i).Y2 * lngRate
 Next
 For i = 0 To UBound(PartB)
    PartB(i).X1 = PartB(i).X1 * lngRate
    PartB(i).X2 = PartB(i).X2 * lngRate
    PartB(i).Y1 = PartB(i).Y1 * lngRate
    PartB(i).Y2 = PartB(i).Y2 * lngRate
 Next
 On Error Resume Next
 ReDim Preserve OutLineP(UBound(OutLineP) - 1)
 ReDim Preserve PinsT(UBound(PinsT) - 1)
 ReDim Preserve PinsB(UBound(PinsB) - 1)
 ReDim Preserve NailB(UBound(NailB) - 1)
 ReDim Preserve NailT(UBound(NailT) - 1)
' ReDim Preserve PartT(UBound(PartT) - 1)
' ReDim Preserve PartB(UBound(PartB) - 1)
  'end load data
  
 
 Call MovePic(Me.Picture1.Width / 2 - pCenter.x, Me.Picture1.Height / 2 - pCenter.y)
 
  'Draw Navigate picture and save it
 Me.picNavigate.Visible = True
 Dim lngTmprate As Long
 lngTmprate = Me.Picture1.Width / Me.picNavigate.Width
 Set Me.picNavigate.Picture = Nothing
 Me.picNavigate.Cls
 For i = 0 To UBound(OutLineP) - 1
   If OutLineP(i).Group = OutLineP(i + 1).Group Then
    Me.picNavigate.Line (OutLineP(i).x / lngTmprate, OutLineP(i).y / lngTmprate)-(OutLineP(i + 1).x / lngTmprate, OutLineP(i + 1).y / lngTmprate), RGB(255, 0, 0)
   End If
'   Me.Picture1.Width / 2 - pCenter.X

 Next i
 Set imgNavigate = Me.picNavigate.Image
 'End draw Navigate picture
 
 Call EnableCtls
 Call InitialParameter
 
 Picture1.Enabled = True
 Picture1.SetFocus
 Screen.MousePointer = vbNormal
End Sub

 

Private Sub ZoomOut()
 
Call Zoom(0.5, Picture1.Width / 2, Picture1.Height / 2)
Call drawPic

End Sub

Private Sub ZoomIn()
' Call Picture1_MouseUp(vbRightButton, 0, Me.Picture1.Width / 2, Me.Picture1.Height / 2)
Call Zoom(2, Me.Picture1.Width / 2, Me.Picture1.Height / 2)
Call drawPic

End Sub
  
Private Sub Form_Load()
lngRate = 250  '设定放大倍数
intCircleRate = 130  '设定画圆时的比率
lngFindRate = 4000  '设定找到某东西时放大的倍数

'intPinColorNormal = 7
'intPinColorSel = 10
'
'intPartColorNormal = 8
'intPartColorSel = 10

intPinColorNormal = 13
intPinColorSel = 10

intPartColorNormal = 11
intPartColorSel = 10


sngSpace = 1.1 '设定part的边到pin中心的距离是pin半径的倍数
intPinColorFind(0) = 12
intPinColorFind(1) = 9
intPinColorFind(2) = 10
intPinColorFind(3) = 6


strCommand = Interaction.Command


 
End Sub

Private Sub Form_Resize()
Dim n As Integer

On Error Resume Next
With Me.Picture1
  .left = 0
  .Width = Me.Width
  .Height = Me.Height - .top - Me.stbBar.Height - 650
'  Me.picNavigate.Left = .Left + .Width - picNavigate.Width - 100
'  Me.picNavigate.Top = 0
  If Me.picNavigate.left <> 10 Then
    Me.picNavigate.left = 10
    Me.picNavigate.top = .top
    n = .Height / Me.picNavigate.Height
    Me.picNavigate.Height = .Height / n
    Me.picNavigate.Width = .Width / n
    Call LoadLogo
  End If
End With
 
End Sub

Private Sub mnuPrint_Click()
    Dim i As Single
    i = InputBox("Please input size ", "Input", 2)
    
    Call DrawOutLine2Printer(i)
    Call DrawPart2Printer(i)
    Printer.EndDoc
    
End Sub

Private Sub mnuRotate_Click()
   Call Rotate
   Call mnuZoomPerfect_Click
End Sub

Private Sub mnuTurn_Click()
     Call TurnBack
    Call drawPic
End Sub

Private Sub mnuZoomin_Click()
  Call ZoomIn
End Sub

Private Sub mnuZoomout_Click()
  Call ZoomOut
End Sub

Private Sub mnuZoomPerfect_Click()
  Call ZoomPerfect
End Sub

Private Sub MrrorTurn180_Click()
       DownFkey = True
       Call TurnBack
       Call drawPic
       DownFkey = False
End Sub

Private Sub Picture1_DblClick()
If intMouseKey = vbLeftButton Then

  Call Zoom(2, MouseX, MouseY)
Else
  Call Zoom(0.5, MouseX, MouseY)
End If
DBclikc = True

 Call drawPic
 
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim strTmpPinSel As String
  
  MouseDownX = x
  MouseDownY = y
  intMouseKey = Button
  
blnBoxSelect = False

Set imgPic = Picture1.Image
If Button = vbLeftButton Then
'
'
'       Call FindPinNail_XY   'if find it will clear  old color
'      If Trim(str_ColorPinT & str_ColorPinB & str_ColorNailT & str_ColorNailB) = "" Then
'    '     如果没有找到Pin 和 Nail 那就找Part
'         Call FindPart_XY
'      End If
ElseIf Button = vbRightButton Then
    Set Picture1.Picture = imgPic
 
'
'  Call FindPinNail_XY
'  Call drawPic
'  If Trim(str_ColorNailT & str_ColorNailB) = "" Then
'     '没有找到Nail
'     Exit Sub
'  Else '找到了Nail
'     Call ShowNailDetail  '在这里str_colorNail里一定要有东西
'  End If
  
End If
  
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

 Dim sngX As Single
 Dim sngY As Single
 Dim i As Integer
 
 If Button = vbLeftButton Then  '拖动
   sngX = x - MouseX
   sngY = y - MouseY
   Call MovePic(sngX, sngY)
   
'   For i = 0 To UBound(OutLineP)
'    OutLineP(i).X = OutLineP(i).X + sngX
'    OutLineP(i).Y = OutLineP(i).Y + sngY
'
'   Next i
'
'
'   For i = 0 To UBound(PinsT)
'    PinsT(i).X = PinsT(i).X + sngX
'    PinsT(i).Y = PinsT(i).Y + sngY
'
'   Next i
'
'   For i = 0 To UBound(PinsB)
'    PinsB(i).X = PinsB(i).X + sngX
'    PinsB(i).Y = PinsB(i).Y + sngY
'
'   Next i
'
'   'deal Part
'   For i = 0 To UBound(PartT)
'    PartT(i).X1 = PartT(i).X1 + sngX
'    PartT(i).Y1 = PartT(i).Y1 + sngY
'    PartT(i).X2 = PartT(i).X2 + sngX
'    PartT(i).Y2 = PartT(i).Y2 + sngY
'   Next i
'
'   For i = 0 To UBound(PartB)
'    PartB(i).X1 = PartB(i).X1 + sngX
'    PartB(i).Y1 = PartB(i).Y1 + sngY
'    PartB(i).X2 = PartB(i).X2 + sngX
'    PartB(i).Y2 = PartB(i).Y2 + sngY
'   Next i
'
'
'   Call drawPic
  ElseIf Button = vbRightButton Then   '框选。
'    If Picture1.Picture = 0 Then
'    Set Picture1.Picture = imgPic
'    End If
    Picture1.AutoRedraw = False
    Picture1.Cls
    Picture1.Line (MouseDownX, MouseDownY)-(x, y), , B
    blnBoxSelect = True
    Picture1.AutoRedraw = True
  End If
 
 MouseX = x
 MouseY = y
 
 
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim i As Integer
Dim strP_N As String

If Button = vbLeftButton Then
  strP_N = str_ColorPinT & str_ColorPinB & str_ColorNailT & str_ColorNailB
  If x = MouseDownX And y = MouseDownY And DBclikc = False Then
       Call FindPinNail_XY   'if find it will clear  old color
      If str_ColorPinT & str_ColorPinB & str_ColorNailT & str_ColorNailB = strP_N Then
    '     如果没有找到Pin 和 Nail 那就找Part
         Call FindPart_XY
      End If
  End If
  
'Call DrawPinNailColor
Else  '看是不是框选完毕？'right button
  If blnBoxSelect = True Then
    Dim X1 As Single
    Dim Y1 As Single
    Dim X2 As Single
    Dim Y2 As Single
    
    Dim tmpRate As Single
    
    X1 = IIf(MouseDownX < x, MouseDownX, x)
    X2 = IIf(MouseDownX < x, x, MouseDownX)
    Y1 = IIf(MouseDownY < y, MouseDownY, y)
    Y2 = IIf(MouseDownY < y, y, MouseDownY)
    
    If X2 = X1 Or Y1 = Y2 Then Exit Sub
     
    '选一个小的值做为放大倍数
    tmpRate = Picture1.Width / (X2 - X1)
    If tmpRate > Picture1.Height / (Y2 - Y1) Then
      tmpRate = Picture1.Height / (Y2 - Y1)
    End If
    Call MovePic(X1 * -1, Y1 * -1)
    Call Zoom(tmpRate, 0, 0)
  ElseIf Me.chkEnableRightMouse.Value = 1 Then
    Call ClearNailColor
    Call FindPinNail_XY
'    Call drawPic
    If Trim(str_ColorNailT & str_ColorNailB) = "" Then
       '没有找到Nail
       Exit Sub
    Else '找到了Nail
       'Call ShowNailDetail  '在这里str_colorNail里一定要有东西
    End If

  End If
  
End If

DBclikc = False
 
 Call drawPic

End Sub
Private Sub DrawOutLine()
Dim i As Integer
On Error Resume Next

    For i = 0 To UBound(OutLineP) - 1
'      If OutLineP(i).X > 0 Or OutLineP(i + 1).X > 0 Then
'         If OutLineP(i).X < Me.Picture1.Width Or OutLineP(i + 1).X < Me.Picture1.Width Then
'            If OutLineP(i).Y > 0 Or OutLineP(i).Y > 0 Then
'               If OutLineP(i).Y < Me.Picture1.Height Or OutLineP(i).Y < Me.Picture1.Height Then
'                    Picture1.Line (OutLineP(i).X, OutLineP(i).Y)-(OutLineP(i + 1).X, OutLineP(i + 1).Y), RGB(255, 0, 0)
'                End If
'             End If
'          End If
'       End If
       If OutLineP(i).Group = OutLineP(i + 1).Group Then
        Picture1.Line (OutLineP(i).x, OutLineP(i).y)-(OutLineP(i + 1).x, OutLineP(i + 1).y), RGB(255, 0, 0)
       End If
    Next i
          
End Sub
Private Sub DrawOutLine2Printer(iSize As Single)
Dim i As Integer
           
    For i = 0 To UBound(OutLineP) - 1
       If OutLineP(i).Group = OutLineP(i + 1).Group Then
           Printer.Line (OutLineP(i).x * iSize, OutLineP(i).y * iSize)-(OutLineP(i + 1).x * iSize, OutLineP(i + 1).y * iSize), RGB(255, 0, 0)
       End If


    Next i
          
End Sub

Private Sub DrawPin()
 On Error Resume Next '防止图片中只有Top 或是只有botoom

Dim i As Long


 If blnTop = True Then   'Draw top Pin
    For i = 0 To UBound(PinsT) - 1
      If PinsT(i).x > 0 Then
         If PinsT(i).x < Me.Picture1.Width Then
            If PinsT(i).y > 0 Then
               If PinsT(i).y < Me.Picture1.Height Then
                    Picture1.Circle (PinsT(i).x, PinsT(i).y), lngRate / intCircleRate, QBColor(PinsT(i).Color)
                    If PinsT(i).Color <> intPinColorNormal Then Picture1.Print PinsT(i).Part & "." & PinsT(i).Name
                End If
             End If
          End If
       End If
    Next i
  Else                  'Draw Bottom Pin
    For i = 0 To UBound(PinsB) - 1
      If PinsB(i).x > 0 Then
         If PinsB(i).x < Me.Picture1.Width Then
            If PinsB(i).y > 0 Then
               If PinsB(i).y < Me.Picture1.Height Then
                    Picture1.Circle (PinsB(i).x, PinsB(i).y), lngRate / intCircleRate, QBColor(PinsB(i).Color)
                    If PinsB(i).Color <> intPinColorNormal Then Picture1.Print PinsB(i).Part & "." & PinsB(i).Name
                End If
             End If
          End If
       End If
           
    Next i
  End If

End Sub
Private Sub DrawPin2Printer()
'On Error Resume Next '防止图片中只有Top 或是只有botoom

Dim i As Long
 If blnTop = True Then   'Draw top Pin
    For i = 0 To UBound(PinsT) - 1
        Printer.Circle (PinsT(i).x, PinsT(i).y), lngRate / intCircleRate, QBColor(PinsT(i).Color)
    Next i
  Else                  'Draw Bottom Pin
    For i = 0 To UBound(PinsB) - 1
        Printer.Circle (PinsB(i).x, PinsB(i).y), lngRate / intCircleRate, QBColor(PinsB(i).Color)
    Next i
  End If

End Sub


Private Sub DrawNail()
Dim i As Long
Dim dblRadius As Single

 On Error Resume Next '防止图片中只有Top 或是只有botoom

dblRadius = lngRate / intCircleRate
 If blnTop = True Then   'Draw top Parts
    For i = 0 To UBound(NailT)
'      Picture1.Circle (NailT(i).X, NailT(i).Y), lngRate / intCircleRate, QBColor(NailT(i).Color)
       Picture1.Line (NailT(i).x, NailT(i).y - dblRadius)-(NailT(i).x - dblRadius, NailT(i).y), QBColor(NailT(i).Color)
       Picture1.Line (NailT(i).x - dblRadius, NailT(i).y)-(NailT(i).x, NailT(i).y + dblRadius), QBColor(NailT(i).Color)
       Picture1.Line (NailT(i).x, NailT(i).y + dblRadius)-(NailT(i).x + dblRadius, NailT(i).y), QBColor(NailT(i).Color)
       Picture1.Line (NailT(i).x + dblRadius, NailT(i).y)-(NailT(i).x, NailT(i).y - dblRadius), QBColor(NailT(i).Color)
      
      
      If NailT(i).Color <> intPinColorNormal Then
        Picture1.Print NailT(i).Nail
      End If
    Next i
  
  Else                  'Draw Bottom Parts
    For i = 0 To UBound(NailB)
'      Picture1.Circle (NailB(i).X, NailB(i).Y), lngRate / intCircleRate, QBColor(NailB(i).Color)
       Picture1.Line (NailB(i).x, NailB(i).y - dblRadius)-(NailB(i).x - dblRadius, NailB(i).y), QBColor(NailB(i).Color)
       Picture1.Line (NailB(i).x - dblRadius, NailB(i).y)-(NailB(i).x, NailB(i).y + dblRadius), QBColor(NailB(i).Color)
       Picture1.Line (NailB(i).x, NailB(i).y + dblRadius)-(NailB(i).x + dblRadius, NailB(i).y), QBColor(NailB(i).Color)
       Picture1.Line (NailB(i).x + dblRadius, NailB(i).y)-(NailB(i).x, NailB(i).y - dblRadius), QBColor(NailB(i).Color)

      If NailB(i).Color <> intPinColorNormal Then
        Picture1.Print NailB(i).Nail
      End If
    Next i
  End If
End Sub
Private Sub DrawPart()
 On Error Resume Next '防止图片中只有Top 或是只有botoom

Dim i As Long
 If blnTop = True Then   'Draw top Parts
    For i = 0 To UBound(PartT)
'      If partt(i).X1 > 0 Then
'         If partt(i).X1 < Me.Picture1.Width Then
'            If partt(i).Y1 > 0 Then
'               If partt(i).Y1 < Me.Picture1.Height Then
'                    Picture1.Line (partt(i).X1, partt(i).Y1)-(partt(i).X2, partt(i).Y2), QBColor(PinsT(i).Color), B
'                End If
'             End If
'          End If
'       End If
      Picture1.Line (PartT(i).X2, PartT(i).Y2)-(PartT(i).X1, PartT(i).Y1), QBColor(PartT(i).Color), B
      If PartT(i).Color <> intPartColorNormal Or Me.chkShowPartName.Value = 1 Then
        Picture1.Print PartT(i).Name
      End If
      
    Next i
  
  Else                  'Draw Bottom Parts
    For i = 0 To UBound(PartB)
     Picture1.Line (PartB(i).X2, PartB(i).Y2)-(PartB(i).X1, PartB(i).Y1), QBColor(PartB(i).Color), B
     If PartB(i).Color <> intPartColorNormal Or Me.chkShowPartName.Value = 1 Then
        Picture1.Print PartB(i).Name
     End If
     
    Next i
  End If

End Sub

Private Sub DrawPart2Printer(iSize As Single)
'On Error Resume Next '防止图片中只有Top 或是只有botoom

Dim i As Long
 If blnTop = True Then   'Draw top Parts
    For i = 0 To UBound(PartT)
'      If partt(i).X1 > 0 Then
'         If partt(i).X1 < Me.Picture1.Width Then
'            If partt(i).Y1 > 0 Then
'               If partt(i).Y1 < Me.Picture1.Height Then
'                    Picture1.Line (partt(i).X1, partt(i).Y1)-(partt(i).X2, partt(i).Y2), QBColor(PinsT(i).Color), B
'                End If
'             End If
'          End If
'       End If
      Printer.Line (PartT(i).X2 * iSize, PartT(i).Y2 * iSize)-(PartT(i).X1 * iSize, PartT(i).Y1 * iSize), QBColor(PartT(i).Color), B
       
    Next i
  
  Else                  'Draw Bottom Parts
    For i = 0 To UBound(PartB)
     Printer.Line (PartB(i).X2 * iSize, PartB(i).Y2 * iSize)-(PartB(i).X1 * iSize, PartB(i).Y1 * iSize), QBColor(PartB(i).Color), B
      
    Next i
  End If

End Sub

Private Sub TurnBack()
Dim i As Integer

'On Error Resume Next '防止图片中只有Top 或是只有botoom

For i = 0 To UBound(OutLineP)
 OutLineP(i).x = OutLineP(i).x * -1 + Me.Picture1.Width
Next i
For i = 0 To UBound(PinsT)
 PinsT(i).x = PinsT(i).x * -1 + Me.Picture1.Width
Next i
For i = 0 To UBound(PinsB)
 PinsB(i).x = PinsB(i).x * -1 + Me.Picture1.Width
Next i

For i = 0 To UBound(PartT)
 PartT(i).X1 = PartT(i).X1 * -1 + Me.Picture1.Width
 PartT(i).X2 = PartT(i).X2 * -1 + Me.Picture1.Width
Next i
For i = 0 To UBound(PartB)
 PartB(i).X1 = PartB(i).X1 * -1 + Me.Picture1.Width
 PartB(i).X2 = PartB(i).X2 * -1 + Me.Picture1.Width
Next i
'do Nail top
For i = 0 To UBound(NailT)
 NailT(i).x = NailT(i).x * -1 + Me.Picture1.Width
Next i
'do Nail Bottom
For i = 0 To UBound(NailB)
 NailB(i).x = NailB(i).x * -1 + Me.Picture1.Width
Next i

'Do center
pCenter.x = pCenter.x * -1 + Me.Picture1.Width
    blnTop = Not blnTop
    
'blnRedrawNvigate = True '重画导航图

End Sub

Private Sub drawPic()
Set Picture1.Picture = Nothing

Me.Picture1.Cls
 Call DrawOutLine
 If DownFkey = True Then 'great add
     blnTop = Not blnTop
 End If
 Call DrawPin
 Call DrawPart
 Call DrawNail
 Call DrawNavigate
If blnTop = True Then
     
    Me.lblTB.Caption = "TOP"
    '   Set Me.picNavigate.Picture = imgNavigate
    Else
       Me.lblTB.Caption = "Bottom"
    '   Set Me.picNavigate.Picture = Nothing
    '   Me.picNavigate.PaintPicture imgNavigate, 0, 0, Me.picNavigate.Width, Me.picNavigate.Height, Me.picNavigate.Width, 0, Me.picNavigate.Width * -1, Me.picNavigate.Height
End If
End Sub

Private Sub FindPinNail_XY()
 Dim i As Long
 
' On Error Resume Next '防止图片中只有Top 或是只有botoom
 
' Call ClearPinNailColor       '只有当找到时才清除上次的选中
' Call ClearPartColor          '现在变了，找不找的到都要先清除
                               
 
 If blnTop = True Then 'now it's top
    'Find inPins
    For i = 0 To UBound(PinsT)
      If MouseDownX > PinsT(i).x - lngRate / intCircleRate Then
          If MouseDownX < PinsT(i).x + lngRate / intCircleRate Then
             If MouseDownY > PinsT(i).y - lngRate / intCircleRate Then
               If MouseDownY < PinsT(i).y + lngRate / intCircleRate Then
'                    Call ClearPinNailColor       '只有当找到时才清除上次的选中
                    If PinsT(i).Color = intPinColorSel Then
                         PinsT(i).Color = intPinColorNormal
                         str_ColorPinT = Replace(str_ColorPinT, " " & i, "")
                    Else
                         PinsT(i).Color = intPinColorSel
                         str_ColorPinT = str_ColorPinT & " " & i
                    End If
                    Me.stbBar.Panels(1).Text = "Part:" & PinsT(i).Part & "  Pin:" & PinsT(i).Name & "  Net:" & PinsT(i).NetName
                   Exit For
               End If
             End If
          End If
      End If
    Next i
    
    'Find in Nail  top
    For i = 0 To UBound(NailT)
      If MouseDownX > NailT(i).x - lngRate / intCircleRate Then
          If MouseDownX < NailT(i).x + lngRate / intCircleRate Then
             If MouseDownY > NailT(i).y - lngRate / intCircleRate Then
               If MouseDownY < NailT(i).y + lngRate / intCircleRate Then
'                    Call ClearPinNailColor       '只有当找到时才清除上次的选中
                    If NailT(i).Color = intPinColorSel Then
                        NailT(i).Color = intPinColorNormal
                        str_ColorNailT = Replace(str_ColorNailT, " " & i, "")
                    Else
                        NailT(i).Color = intPinColorSel
                        str_ColorNailT = str_ColorNailT & " " & i
                    End If
                    
                    
                    Me.stbBar.Panels(1).Text = "Nail:" & NailT(i).Nail & "  Type:" & NailT(i).Type & "  Grid:" & NailT(i).Grid & "  Net:" & NailT(i).NetName
                   Exit For
               End If
             End If
          End If
      End If
    Next i
 Else     'now it's bottom
    'Find in Pins
    For i = 0 To UBound(PinsB)
      If MouseDownX > PinsB(i).x - lngRate / intCircleRate Then
          If MouseDownX < PinsB(i).x + lngRate / intCircleRate Then
             If MouseDownY > PinsB(i).y - lngRate / intCircleRate Then
               If MouseDownY < PinsB(i).y + lngRate / intCircleRate Then
'                    Call ClearPinNailColor       '只有当找到时才清除上次的选中
                   If PinsB(i).Color = intPinColorSel Then
                    PinsB(i).Color = intPinColorNormal
                    str_ColorPinB = Replace(str_ColorPinB, " " & i, "")
                   Else
                    PinsB(i).Color = intPinColorSel
                    str_ColorPinB = str_ColorPinB & " " & i
                   End If
                    
                   Me.stbBar.Panels(1).Text = "Part:" & PinsB(i).Part & "  Pin:" & PinsB(i).Name & "  Net:" & PinsB(i).NetName
                  Exit For
               End If
             End If
          End If
      End If
    Next i
    'Find in Nails Bottom
    For i = 0 To UBound(NailB)
      If MouseDownX > NailB(i).x - lngRate / intCircleRate Then
          If MouseDownX < NailB(i).x + lngRate / intCircleRate Then
             If MouseDownY > NailB(i).y - lngRate / intCircleRate Then
               If MouseDownY < NailB(i).y + lngRate / intCircleRate Then
'                    Call ClearPinNailColor       '只有当找到时才清除上次的选中
                    If NailB(i).Color = intPinColorSel Then
                        NailB(i).Color = intPinColorNormal
                        str_ColorNailB = Replace(str_ColorNailB, " " & i, "")
                    Else
                        NailB(i).Color = intPinColorSel
                        str_ColorNailB = str_ColorNailB & " " & i
                    End If
                    
                    Me.stbBar.Panels(1).Text = "Nail:" & NailB(i).Nail & "  Type:" & NailB(i).Type & "  Grid:" & NailB(i).Grid & "  Net:" & NailB(i).NetName
                   Exit For
               End If
             End If
          End If
      End If
    Next i
 End If
 
End Sub
Private Sub FindPart_XY()
 Dim i As Long
 
' On Error Resume Next '防止图片中只有Top 或是只有botoom
 
 If blnTop = True Then 'now it's top
    'Find in Part
    For i = 0 To UBound(PartT)
      If (MouseDownX - PartT(i).X1) * (MouseDownX - PartT(i).X2) < 0 Then
         If (MouseDownY - PartT(i).Y1) * (MouseDownY - PartT(i).Y2) < 0 Then
                Call ClearPartColor       '只有当找到时才清除上次的选中
                PartT(i).Color = intPartColorSel
                str_ColorPartT = str_ColorPartT & " " & i
                Me.stbBar.Panels(1).Text = "Part:" & PartT(i).Name ' & "  Name:" & PinsT(i).Name & "  Net:" & PinsT(i).Net
                Exit For
          End If
      End If
    Next i
    
 Else     'now it's bottom
    'Find in Part
    For i = 0 To UBound(PartB)
      If (MouseDownX - PartB(i).X2) * (MouseDownX - PartB(i).X1) < 0 Then
             If (MouseDownY - PartB(i).Y1) * (MouseDownY - PartB(i).Y2) < 0 Then
                    Call ClearPartColor       '只有当找到时才清除上次的选中
                    PartB(i).Color = intPartColorSel
                    str_ColorPartB = str_ColorPartB & " " & i
                   Me.stbBar.Panels(1).Text = "Part:" & PartB(i).Name  '& "  Name:" & PinsB(i).Name & "  Net:" & PinsB(i).Net
                  Exit For
             End If
      End If
    Next i
 
 End If
 
End Sub

'Private Sub DrawPinNailColor()
''只画出有颜色的Pin和Nail
'  Dim i As Long
'  Dim ar() As String
'  ar = Split(Trim(str_ColorPinT), " ")
''现在处理Pin
'  For i = 0 To UBound(ar)
''   PinsT(Val(ar(i))).Color = intPinColorSel
'   If blnTop = True Then   '如果现在是top
'        Picture1.Circle (PinsT(Val(Val(ar(i)))).X, PinsT(Val(ar(i))).Y), lngRate / intCircleRate, QBColor(intPinColorSel)
'   End If
'  Next i
'
'  ar = Split(Trim(str_ColorPinB), " ")
'  For i = 0 To UBound(ar)
''   PinsB(Val(ar(i))).Color = intPinColorSel
'   If blnTop = False Then '如果现在是bottom
'        Picture1.Circle (PinsB(Val(Val(ar(i)))).X, PinsB(Val(ar(i))).Y), lngRate / intCircleRate, QBColor(intPinColorSel)
'   End If
'  Next i
'  'Pin处理完
'
'  'Now Nail
'     'Now Top
'  ar = Split(Trim(str_ColorNailT), " ")
'  For i = 0 To UBound(ar)
''   PinsT(Val(ar(i))).Color = intPinColorSel
'   If blnTop = True Then   '如果现在是top
'        Picture1.Circle (NailT(Val(Val(ar(i)))).X, NailT(Val(ar(i))).Y), lngRate / intCircleRate, QBColor(intPinColorSel)
'   End If
'  Next i
'
'    'Now Bottom
'  ar = Split(Trim(str_ColorNailB), " ")
'  For i = 0 To UBound(ar)
''   PinsB(Val(ar(i))).Color = intPinColorSel
'   If blnTop = False Then '如果现在是bottom
'        Picture1.Circle (NailB(Val(Val(ar(i)))).X, NailB(Val(ar(i))).Y), lngRate / intCircleRate, QBColor(intPinColorSel)
'   End If
'  Next i
'
'  'Nail End
'
'
'End Sub
Private Sub ClearPinNailColor()
'clear pin and Nail color and redraw it as Normal color

  Call ClearPinColor
  Call ClearNailColor
 
End Sub
Private Sub ClearPinColor()
On Error Resume Next

  Dim i As Integer
  Dim ar() As String
'Now do Pin
  ar = Split(Trim(str_ColorPinT), " ")
  For i = 0 To UBound(ar)
   PinsT(Val(ar(i))).Color = intPinColorNormal
'   If blnTop = True Then   '如果现在是top
'        Picture1.Circle (PinsT(Val(Val(ar(i)))).X, PinsT(Val(ar(i))).Y), lngRate / intCircleRate, QBColor(intPinColorNormal)
'   End If
  Next i
  str_ColorPinT = ""
  
  ar = Split(Trim(str_ColorPinB), " ")
  For i = 0 To UBound(ar)
   PinsB(Val(ar(i))).Color = intPinColorNormal
'   If blnTop = False Then '如果现在是bottom
'        Picture1.Circle (PinsB(Val(Val(ar(i)))).X, PinsB(Val(ar(i))).Y), lngRate / intCircleRate, QBColor(intPinColorNormal)
'   End If
  Next i
  str_ColorPinB = ""
'Pin End
End Sub
Private Sub ClearNailColor()
On Error Resume Next
  Dim i As Integer
  Dim ar() As String
'---------------------------------------------
'Now Nail
  ar = Split(Trim(str_ColorNailT), " ")
  For i = 0 To UBound(ar)
   NailT(Val(ar(i))).Color = intPinColorNormal
'   If blnTop = True Then   '如果现在是top
'        Picture1.Circle (NailT(Val(Val(ar(i)))).X, NailT(Val(ar(i))).Y), lngRate / intCircleRate, QBColor(intPinColorNormal)
'   End If
  Next i
  str_ColorNailT = ""
  
  ar = Split(Trim(str_ColorNailB), " ")
  For i = 0 To UBound(ar)
   NailB(Val(ar(i))).Color = intPinColorNormal
'   If blnTop = False Then '如果现在是bottom
'        Picture1.Circle (NailB(Val(Val(ar(i)))).X, NailB(Val(ar(i))).Y), lngRate / intCircleRate, QBColor(intPinColorNormal)
'   End If
  Next i
  str_ColorNailB = ""
 'Nail End
End Sub
Private Sub ClearPartColor()
On Error Resume Next
'clear pin and Nail color and redraw it as Normal color
  Dim i As Integer
  Dim ar() As String

'Now do Part
  ar = Split(Trim(str_ColorPartT), " ")
  For i = 0 To UBound(ar)
   PartT(Val(ar(i))).Color = intPartColorNormal
  Next i
  str_ColorPartT = ""
  
  ar = Split(Trim(str_ColorPartB), " ")
  For i = 0 To UBound(ar)
   PartB(Val(ar(i))).Color = intPartColorNormal
  Next i
  str_ColorPartB = ""
'Part End
'---------------------------------------------
  
End Sub
Private Sub Zoom(N_large As Single, X_center As Single, Y_center As Single)
'N_large 放大到现在的倍数 可以为小数（即缩小）
'X，Y以这点为中心放大缩小
'On Error Resume Next '防止图片中只有Top 或是只有botoom

Dim i As Integer
 
 If lngRate * N_large < 256000 And lngRate * N_large > 30 Then '控制不能放的太大,或是太小。
     lngRate = lngRate * N_large
 Else
    Exit Sub
 End If
  
  'do OutLine
For i = 0 To UBound(OutLineP)
 OutLineP(i).x = (OutLineP(i).x - X_center) * N_large + X_center
 OutLineP(i).y = (OutLineP(i).y - Y_center) * N_large + Y_center
Next i

'do Pins top
For i = 0 To UBound(PinsT)
 PinsT(i).x = (PinsT(i).x - X_center) * N_large + X_center
 PinsT(i).y = (PinsT(i).y - Y_center) * N_large + Y_center
Next i
'do pins bottom
For i = 0 To UBound(PinsB)
 PinsB(i).x = (PinsB(i).x - X_center) * N_large + X_center
 PinsB(i).y = (PinsB(i).y - Y_center) * N_large + Y_center
Next i
'do parts top
For i = 0 To UBound(PartT)
 PartT(i).X1 = (PartT(i).X1 - X_center) * N_large + X_center
 PartT(i).Y1 = (PartT(i).Y1 - Y_center) * N_large + Y_center
 PartT(i).X2 = (PartT(i).X2 - X_center) * N_large + X_center
 PartT(i).Y2 = (PartT(i).Y2 - Y_center) * N_large + Y_center
Next i
'do parts Bottom
For i = 0 To UBound(PartB)
 PartB(i).X1 = (PartB(i).X1 - X_center) * N_large + X_center
 PartB(i).Y1 = (PartB(i).Y1 - Y_center) * N_large + Y_center
 PartB(i).X2 = (PartB(i).X2 - X_center) * N_large + X_center
 PartB(i).Y2 = (PartB(i).Y2 - Y_center) * N_large + Y_center
Next i

'do Pins top
For i = 0 To UBound(NailT)
 NailT(i).x = (NailT(i).x - X_center) * N_large + X_center
 NailT(i).y = (NailT(i).y - Y_center) * N_large + Y_center
Next i
'do Pins Bottom
For i = 0 To UBound(NailB)
 NailB(i).x = (NailB(i).x - X_center) * N_large + X_center
 NailB(i).y = (NailB(i).y - Y_center) * N_large + Y_center
Next i

 pCenter.x = (pCenter.x - X_center) * N_large + X_center
 pCenter.y = (pCenter.y - Y_center) * N_large + Y_center

End Sub
 
Private Sub MovePic(Xshift As Single, Yshift As Single)
' On Error Resume Next '防止图片中只有Top 或是只有botoom
 
  Dim i As Long
  Dim sngX As Single
  Dim sngY As Single
  sngX = Xshift
  sngY = Yshift
  
   For i = 0 To UBound(OutLineP)
    OutLineP(i).x = OutLineP(i).x + sngX
    OutLineP(i).y = OutLineP(i).y + sngY
   Next i
   
   'do Pin Top
   For i = 0 To UBound(PinsT)
    PinsT(i).x = PinsT(i).x + sngX
    PinsT(i).y = PinsT(i).y + sngY
   Next i
   'Do Pin Bottom
   For i = 0 To UBound(PinsB)
    PinsB(i).x = PinsB(i).x + sngX
    PinsB(i).y = PinsB(i).y + sngY
   Next i
   
   'deal Part
   For i = 0 To UBound(PartT)
    PartT(i).X1 = PartT(i).X1 + sngX
    PartT(i).Y1 = PartT(i).Y1 + sngY
    PartT(i).X2 = PartT(i).X2 + sngX
    PartT(i).Y2 = PartT(i).Y2 + sngY
   Next i

   For i = 0 To UBound(PartB)
    PartB(i).X1 = PartB(i).X1 + sngX
    PartB(i).Y1 = PartB(i).Y1 + sngY
    PartB(i).X2 = PartB(i).X2 + sngX
    PartB(i).Y2 = PartB(i).Y2 + sngY
   Next i
   
   'do Nail Top
   For i = 0 To UBound(NailT)
    NailT(i).x = NailT(i).x + sngX
    NailT(i).y = NailT(i).y + sngY
   Next i
   'do Nail Bottom
   For i = 0 To UBound(NailB)
    NailB(i).x = NailB(i).x + sngX
    NailB(i).y = NailB(i).y + sngY
   Next i
   
   'do pCenter point
   
   pCenter.x = pCenter.x + sngX
   pCenter.y = pCenter.y + sngY
   
   Call drawPic

End Sub
Private Sub FindPart_Name()
  Dim strTmp As String
  Dim i As Long
  Dim blnFind As Boolean  '是否找到
  Dim blnAtTop As Boolean  '如果找到那么是在哪面
  Dim strTmp2 As String
  Dim strTmp3 As String
  Dim Find1True As Boolean
  Dim Find2True As Boolean
  Dim Find3True As Boolean
  
  ' Dim f As New frmFindDvs
   
  
  Find1True = True
  Find2True = True
  Find3True = True
  
   On Error Resume Next '防止图片中只有Top 或是只有botoom
  

 '  f.Show vbModal
'   strTmp = UCase(Trim(f.text1.Text))
'   If strTmp = "" Then Find1True = False
'   strTmp2 = UCase(Trim(f.text2.Text))
'   If strTmp2 = "" Then Find2True = False
'   strTmp3 = UCase(Trim(f.text3.Text))
'   If strTmp3 = "" Then Find3True = False
'   Unload f
'  If Find1True = False And Find2True = False And Find3True = False Then Exit Sub
  
  If BoardViewTrue = True Then
    strTmp = UCase(BoardViewDevice)
  End If
  
  If BoardViewDevice = "" Then
       strTmp = UCase(Trim(InputBox("Please Input the Componet(Part) Name e.g. ：" & PartT(0).Name & vbCrLf & "Or you can input part.pinNumber e.g.:" & PinsT(0).Part & "." & PinsT(0).Number, "Find Componet(Part)")))
  End If
   
   If strTmp = "" Then Exit Sub
   
  Call ClearPinNailColor
  Call ClearPartColor
  
  If InStr(strTmp, ".") > 1 Then   '不光是要找part，还要找part的一个或是多个脚
     '如果还要找脚，那就多call一个程序。
     Dim strPins As String
     
     Dim strAr() As String
     strAr = Split(strTmp, ".")
     strTmp = strAr(0)
     For i = 1 To UBound(strAr)
       strPins = strPins & " " & strAr(0) & "." & strAr(i)
     Next i
          
     FindPin_Name_Parameter strPins
  End If
  
  
  blnFind = False
  'Find 1------------------------------
  If Find1True = True Then
      For i = 0 To UBound(PartT)
        If PartT(i).Name = strTmp Then
    '      Call ClearPartColor  'Find Clear history
    
          PartT(i).Color = intPartColorSel
          str_ColorPartT = i
          blnFind = True
          blnAtTop = True
          Exit For
        End If
      Next i
   End If
   If Find1True = True Then
      If blnFind = False Then '在top 面没有找到 现在找Tottom
            For i = 0 To UBound(PartB)
              If PartB(i).Name = strTmp Then
    '            Call ClearPartColor  'Find Clear history
                
                PartB(i).Color = intPartColorSel
                str_ColorPartB = i
                blnFind = True
                blnAtTop = False
                Exit For
              End If
            Next i
      End If
   End If
  'Find 1 end --------------------
'If Find1True = True Then
'      If blnFind = False Then Call drawPic: Exit Sub    '没有找到
'
'
'    '找到了。处理....
'    If blnTop <> blnAtTop Then '不在同一面 ,需要先翻转
'        Call TurnBack
'    End If
'     '平移 + Draw
'
'     Call Zoom(lngFindRate / lngRate, 0, 0)  '放大。不重画
'
'     If blnTop = True Then '在top面
'
'       Call MovePic(Me.Picture1.Width / 2 - (PartT(i).X1 + PartT(i).X2) / 2, Me.Picture1.Height / 2 - (PartT(i).Y1 + PartT(i).Y2) / 2) '重画
'
'     Else
'        Call MovePic(Me.Picture1.Width / 2 - (PartB(i).X1 + PartB(i).X2) / 2, Me.Picture1.Height / 2 - (PartB(i).Y1 + PartB(i).Y2) / 2)
'
'     End If
' End If  'Find 2------------------------------
'  If Find2True = True Then
'      For i = 0 To UBound(PartT)
'        If PartT(i).Name = strTmp2 Then
'    '      Call ClearPartColor  'Find Clear history
'
'          PartT(i).Color = intPartColorSel
'          str_ColorPartT = i
'          blnFind = True
'          blnAtTop = True
'          Exit For
'        End If
'      Next i
'   End If
'   If Find2True = True Then
'      If blnFind = False Then '在top 面没有找到 现在找Tottom
'            For i = 0 To UBound(PartB)
'              If PartB(i).Name = strTmp2 Then
'    '            Call ClearPartColor  'Find Clear history
'
'                PartB(i).Color = intPartColorSel
'                str_ColorPartB = i
'                blnFind = True
'                blnAtTop = False
'                Exit For
'              End If
'            Next i
'      End If
'   End If
'  'Find 2 end --------------------
'If Find2True = True Then
'      If blnFind = False Then Call drawPic: Exit Sub    '没有找到
'
'
'    '找到了。处理....
'    If blnTop <> blnAtTop Then '不在同一面 ,需要先翻转
'        Call TurnBack
'    End If
'     '平移 + Draw
'
'     Call Zoom(lngFindRate / lngRate, 0, 0)  '放大。不重画
'
'     If blnTop = True Then '在top面
'
'       Call MovePic(Me.Picture1.Width / 2 - (PartT(i).X1 + PartT(i).X2) / 2, Me.Picture1.Height / 2 - (PartT(i).Y1 + PartT(i).Y2) / 2) '重画
'
'     Else
'        Call MovePic(Me.Picture1.Width / 2 - (PartB(i).X1 + PartB(i).X2) / 2, Me.Picture1.Height / 2 - (PartB(i).Y1 + PartB(i).Y2) / 2)
'
'     End If
' End If
'
'  'Find 3------------------------------
'  If Find3True = True Then
'      For i = 0 To UBound(PartT)
'        If PartT(i).Name = strTmp3 Then
'    '      Call ClearPartColor  'Find Clear history
'
'          PartT(i).Color = intPartColorSel
'          str_ColorPartT = i
'          blnFind = True
'          blnAtTop = True
'          Exit For
'        End If
'      Next i
'   End If
'   If Find3True = True Then
'      If blnFind = False Then '在top 面没有找到 现在找Tottom
'            For i = 0 To UBound(PartB)
'              If PartB(i).Name = strTmp3 Then
'    '            Call ClearPartColor  'Find Clear history
'
'                PartB(i).Color = intPartColorSel
'                str_ColorPartB = i
'                blnFind = True
'                blnAtTop = False
'                Exit For
'              End If
'            Next i
'      End If
'   End If
'  'Find 3 end --------------------
   
If Find1True = True Then
      If blnFind = False Then Call drawPic: Exit Sub    '没有找到
          
      
    '找到了。处理....
    If blnTop <> blnAtTop Then '不在同一面 ,需要先翻转
        Call TurnBack
    End If
     '平移 + Draw
       
     Call Zoom(lngFindRate / lngRate, 0, 0)  '放大。不重画
     
     If blnTop = True Then '在top面
       
       Call MovePic(Me.Picture1.Width / 2 - (PartT(i).X1 + PartT(i).X2) / 2, Me.Picture1.Height / 2 - (PartT(i).Y1 + PartT(i).Y2) / 2) '重画
       
     Else
        Call MovePic(Me.Picture1.Width / 2 - (PartB(i).X1 + PartB(i).X2) / 2, Me.Picture1.Height / 2 - (PartB(i).Y1 + PartB(i).Y2) / 2)
       
     End If
 End If
 
End Sub
Private Sub FindNail_Name()
'通过nail的名字来找nail的信息。
'要求：可以一次找1-3个nail，并放大，平移到中心。

'On Error Resume Next '防止图片中只有Top 或是只有botoom

Dim f As New frmFind
Dim i As Integer
Dim j As Long
Dim intFindTotal As Integer '共找到了几个
Dim strName1 As String  '要找的3 个目标名
Dim strName2 As String
Dim strName3 As String
Dim lngFind As Long        '找到后的位置
Dim blnFindTop As Boolean  '找到后是在哪一面

If BoardViewDevice = "" Then
    f.lblEg.Caption = "Please Input Nail Name e.g. ：" & NailT(0).Nail
    f.Caption = "Find Nail"
    
    'f.text1.Text = "$"
    'f.text2.Text = "$"
    'f.text3.Text = "$"
    
    f.Show vbModal
    
    
     
        strName1 = "$" & UCase(Trim(f.Text1.Text))
    
        
        
    strName2 = "$" & UCase(Trim(f.Text2.Text))
    strName3 = "$" & UCase(Trim(f.Text3.Text))
    Unload f
  Else
     strName1 = "$" & BoardViewDevice
End If
If strName1 = "" Or strName1 = "$" Then intFindTotal = intFindTotal + 1
If strName2 = "" Or strName2 = "$" Then intFindTotal = intFindTotal + 1
If strName3 = "" Or strName3 = "$" Then intFindTotal = intFindTotal + 1

lngFind = -1 '没找之前
Call ClearPinNailColor
Call ClearPartColor

For i = 1 To 3
  For j = 0 To UBound(NailT)
    If intFindTotal = 3 Then Exit For
    If NailT(j).Nail = strName1 Or NailT(j).Nail = strName2 Or NailT(j).Nail = strName3 Then
      intFindTotal = intFindTotal + 1
      NailT(j).Color = intPinColorFind(intFindTotal Mod 4)
      
      str_ColorNailT = str_ColorNailT & " " & j

      If lngFind < 0 Then '是第一个找到的点，保存
        lngFind = j
        blnFindTop = True
      End If
    End If
  Next j
  
  For j = 0 To UBound(NailB)
    If intFindTotal = 3 Then Exit For
    If NailB(j).Nail = strName1 Or NailB(j).Nail = strName2 Or NailB(j).Nail = strName3 Then
      intFindTotal = intFindTotal + 1
      NailB(j).Color = intPinColorFind(intFindTotal Mod 4)
      
      str_ColorNailB = str_ColorNailB & " " & j
      If lngFind < 0 Then '是第一个找到的点，保存
        lngFind = j
        blnFindTop = False
      End If
    End If
  Next j
Next i

  'Now show Status Bar
 If lngFind < 0 Then
     Call drawPic
     Exit Sub
 Else
 
   If blnFindTop = True Then
       Me.stbBar.Panels(1).Text = "Nail:" & NailT(lngFind).Nail & " Net:" & NailT(lngFind).NetName & " Type:" & NailT(lngFind).Type & "  Grid:" & NailT(lngFind).Grid
   Else
       Me.stbBar.Panels(1).Text = "Nail:" & NailB(lngFind).Nail & " Net:" & NailB(lngFind).NetName & " Type:" & NailB(lngFind).Type & "  Grid:" & NailB(lngFind).Grid
   End If
 End If
  'Now turn
  If blnFindTop <> blnTop Then
    Call TurnBack
  End If
  
 Call Zoom(lngFindRate / lngRate, 0, 0)  '放大。不重画
  
  If blnTop = True Then
'      Call MovePic(NailT(lngFind).X, NailT(lngFind).Y)
      Call MovePic(Me.Picture1.Width / 2 - NailT(lngFind).x, Me.Picture1.Height / 2 - NailT(lngFind).y) '重画
  Else
      Call MovePic(Me.Picture1.Width / 2 - NailB(lngFind).x, Me.Picture1.Height / 2 - NailB(lngFind).y) '重画
  End If
End Sub
Private Sub FindPin_Name()
'通过Part的名字和Pin的号码来找Pin的信息。
'要求：可以一次找1-3个Pin，并放大，平移到中心。

'On Error Resume Next '防止图片中只有Top 或是只有botoom

Dim f As New frmFind
Dim f1 As New frmFindDvs
Dim i As Integer
Dim j As Long
Dim intFindTotal As Integer '共找到了几个
Dim strName1 As String  '要找的3 个目标名
Dim strName2 As String
Dim strName3 As String
Dim lngFind As Long        '找到后的位置
Dim blnFindTop As Boolean  '找到后是在哪一面
If DownDkey = False Then
    f.lblEg.Caption = "Please Input PartName and Pin Number e.g. ：" & PinsT(0).Part & "." & PinsT(0).Number
    f.Caption = "Find Pin"
    
    f.Show vbModal
     strName1 = UCase(Trim(f.Text1.Text))
     strName2 = UCase(Trim(f.Text2.Text))
     strName3 = UCase(Trim(f.Text3.Text))
    Unload f
  Else
    f1.Show vbModal
     strName1 = UCase(Trim(f1.Text1.Text)) & ".1"
     strName2 = UCase(Trim(f1.Text2.Text)) & ".1"
     strName3 = UCase(Trim(f1.Text3.Text)) & ".1"
    Unload f1
End If
 
Dim strTmp As String
Call ClearPinNailColor
Call ClearPartColor

 strTmp = FindPin_Name_Parameter(strName1 & " " & strName2 & " " & strName3)
 
'
'If strName1 = "" Then intFindTotal = intFindTotal + 1
'If strName2 = "" Then intFindTotal = intFindTotal + 1
'If strName3 = "" Then intFindTotal = intFindTotal + 1
'
'lngFind = -1 '没找之前

'
'For i = 1 To 3
'  For j = 0 To UBound(PinsT)
'    If intFindTotal = 3 Then Exit For
'    If PinsT(j).Part & "." & PinsT(j).Number = strName1 Or PinsT(j).Part & "." & PinsT(j).Number = strName2 Or PinsT(j).Part & "." & PinsT(j).Number = strName3 Then
'      intFindTotal = intFindTotal + 1
'      PinsT(j).Color = intPinColorFind(intFindTotal - 1)
'
'      str_ColorPinT = str_ColorPinT & " " & j
'
'      If lngFind < 0 Then '是第一个找到的点，保存
'        lngFind = j
'        blnFindTop = True
'      End If
'    End If
'  Next j
'
'  For j = 0 To UBound(PinsB)
'    If intFindTotal = 3 Then Exit For
'    If PinsB(j).Part & "." & PinsB(j).Number = strName1 Or PinsB(j).Part & "." & PinsB(j).Number = strName2 Or PinsB(j).Part & "." & PinsB(j).Number = strName3 Then
'      intFindTotal = intFindTotal + 1
'      PinsB(j).Color = intPinColorFind(intFindTotal - 1)
'
'      str_ColorPinB = str_ColorPinB & " " & j
'      If lngFind < 0 Then '是第一个找到的点，保存
'        lngFind = j
'        blnFindTop = False
'      End If
'    End If
'  Next j
'Next i

  'Now show Status Bar
' If lngFind < 0 Then
'     Exit Sub
' Else
If strTmp = "" Then
   Call drawPic
   Exit Sub
Else
'   If blnFindTop = True Then
'       Me.stbBar.Panels(1).Text = "Part:" & PinsT(lngFind).Part & "  Number:" & PinsT(lngFind).Number & " Name:" & PinsT(lngFind).Name & " Net:" & PinsT(lngFind).Net
'   Else
'       Me.stbBar.Panels(1).Text = "Part:" & PinsB(lngFind).Part & " Number:" & PinsB(lngFind).Number & " Name:" & PinsB(lngFind).Name & "  Net:" & PinsB(lngFind).Net
'   End If
   Dim strTAr() As String
   strTAr = Split(strTmp, " ")
   lngFind = Val(strTAr(1))
   If UCase(left(strTmp, 1)) = "T" Then 'top
       blnFindTop = True
       Me.stbBar.Panels(1).Text = "Part:" & PinsT(lngFind).Part & "  Number:" & PinsT(lngFind).Number & " Name:" & PinsT(lngFind).Name & " Net:" & PinsT(lngFind).NetName
   Else
       blnFindTop = False
       Me.stbBar.Panels(1).Text = "Part:" & PinsB(lngFind).Part & " Number:" & PinsB(lngFind).Number & " Name:" & PinsB(lngFind).Name & "  Net:" & PinsB(lngFind).NetName
   End If
End If
  'Now turn
  If blnFindTop <> blnTop Then
    Call TurnBack
  End If
  
 Call Zoom(lngFindRate / lngRate, 0, 0)  '放大。不重画
  
  If blnTop = True Then
'      Call MovePic(NailT(lngFind).X, NailT(lngFind).Y)
      Call MovePic(Me.Picture1.Width / 2 - PinsT(lngFind).x, Me.Picture1.Height / 2 - PinsT(lngFind).y) '重画
  Else
      Call MovePic(Me.Picture1.Width / 2 - PinsB(lngFind).x, Me.Picture1.Height / 2 - PinsB(lngFind).y) '重画
  End If
   
   
End Sub
Private Function FindPin_Name_Parameter(strAllPin_PartDotNumber As String) As String
'通过Part的名字和Pin的号码来找Pin的信息。 不清除以前的颜色。
'返回第一个脚所在的面和  i  如"T 12" top 面第12个脚
'要求：可以一次找1-3个Pin 参数请用空格分开
Dim i As Integer
Dim j As Long
Dim intFindTotal As Integer '共找到了几个
Dim strName() As String  '要找的3 个目标名
Dim lngFind As Long        '找到后的位置
Dim blnFindTop As Boolean  '找到后是在哪一面

'On Error Resume Next '防止图片中只有Top 或是只有botoom

strAllPin_PartDotNumber = Trim(strAllPin_PartDotNumber)
If strAllPin_PartDotNumber = "" Then Exit Function

strName = Split(strAllPin_PartDotNumber, " ")

For i = 0 To UBound(strName)
   If strName(i) = "" Then intFindTotal = intFindTotal + 1
Next i

lngFind = -1 '没找之前

'For i = 1 To UBound(strName)
  For j = 0 To UBound(PinsT)
    If intFindTotal = UBound(strName) + 1 Then Exit For
    
    For i = 0 To UBound(strName)
     If PinsT(j).Part & "." & PinsT(j).Number = strName(i) Then Exit For
    Next i
    
    If i < UBound(strName) + 1 Then
      intFindTotal = intFindTotal + 1
      PinsT(j).Color = intPinColorFind(2)
      
      str_ColorPinT = str_ColorPinT & " " & j

      If lngFind < 0 Then '是第一个找到的点，保存
        lngFind = j
        blnFindTop = True
      End If
    End If
  Next j
  
  For j = 0 To UBound(PinsB)
    If intFindTotal = UBound(strName) + 1 Then Exit For
    
    For i = 0 To UBound(strName)
     If PinsB(j).Part & "." & PinsB(j).Number = strName(i) Then Exit For
    Next i
    
    If i < UBound(strName) + 1 Then
      intFindTotal = intFindTotal + 1
      PinsB(j).Color = intPinColorFind(intFindTotal Mod 4)
      str_ColorPinB = str_ColorPinB & " " & j
      If lngFind < 0 Then '是第一个找到的点，保存
        lngFind = j
        blnFindTop = False
      End If
    End If
  Next j
'Next i

  'Now show Status Bar
 If lngFind < 0 Then
     FindPin_Name_Parameter = ""
     Exit Function
 End If
  
 If blnFindTop = True Then
    FindPin_Name_Parameter = "T " & lngFind
 Else
    FindPin_Name_Parameter = "B " & lngFind
 End If
End Function

Private Sub FindNet_Name()
  Dim strTmp As String
  Dim i As Long
'  Dim blnFind As Boolean  '是否找到
'  Dim blnAtTop As Boolean  '如果找到那么是在哪面
  
 On Error Resume Next '防止图片中只有Top 或是只有botoom
  
  If BoardViewTrue = True Then
    strTmp = UCase(BoardViewDevice)
  End If

If BoardViewDevice = "" Then
  strTmp = UCase(Trim(InputBox("Please Input the Net Name e.g. ：" & PinsT(0).NetName, "Find Net")))
End If
  
  
  If strTmp = "" Then Exit Sub
  
  Call ClearPinNailColor  '  Clear all history color
  Call ClearPartColor
'Now find the net in pin top
  For i = 0 To UBound(PinsT)
    If PinsT(i).NetName = strTmp Then
      PinsT(i).Color = intPinColorSel
      str_ColorPinT = str_ColorPinT & " " & i
    End If
  Next i
'    fin in pin bottom
  For i = 0 To UBound(PinsB)
    If PinsB(i).NetName = strTmp Then
      PinsB(i).Color = intPinColorSel
      str_ColorPinB = str_ColorPinB & " " & i
    End If
  Next i
'End pin

'Now Nail
   'Top
  For i = 0 To UBound(NailT)
    If NailT(i).NetName = strTmp Then
      NailT(i).Color = intPinColorSel
      str_ColorNailT = str_ColorNailT & " " & i
    End If
  Next i
  'Nail Bottom
  For i = 0 To UBound(NailB)
    If NailB(i).NetName = strTmp Then
      NailB(i).Color = intPinColorSel
      str_ColorNailB = str_ColorNailB & " " & i
    End If
  Next i
  
  Me.stbBar.Panels(1).Text = "Net:" & strTmp
  
  Call ZoomPerfect
  
'End nail
'
'  If blnFind = False Then Exit Sub     '没有找到
'
'
''找到了。处理....
'If blnTop <> blnAtTop Then '不在同一面 ,需要先翻转
'    Call TurnBack
'End If
' '平移 + Draw
'
' Call Zoom(lngFindRate / lngRate, 0, 0)  '放大。不重画
'
' If blnTop = True Then '在top面
'
'   Call MovePic(Me.Picture1.Width / 2 - PartT(i).X1, Me.Picture1.Height / 2 - PartT(i).Y1) '重画
'
' Else
'   Call MovePic(Me.Picture1.Width / 2 - PartB(i).X1, Me.Picture1.Height / 2 - PartB(i).Y1)
' End If


End Sub
Private Sub ZoomPerfect()
 Dim sngRate As Single
 On Error Resume Next
 
  If Me.Picture1.Width / dblWidth_Board < Me.Picture1.Height / dblHeight_Board Then
        sngRate = Me.Picture1.Width * 0.9 / (dblWidth_Board * lngRate)
  Else
        sngRate = Me.Picture1.Height * 0.9 / (dblHeight_Board * lngRate)
  End If
 Call Zoom(sngRate, 0, 0)
 Call MovePic(Me.Picture1.Width / 2 - pCenter.x, Me.Picture1.Height / 2 - pCenter.y)
' Call drawPic
 
End Sub
'Private Sub ShowNailDetail()
'   Dim i As Long
'   Dim j As Long
'   Dim strTmp As String
'   '确认不会是没有选到Nail
'
''   On Error Resume Next '防止图片中只有Top 或是只有botoom
'
'   If blnTop = True Then
'        If Trim(str_ColorNailT) = "" Then
'            Exit Sub
'        Else
'            j = Val(str_ColorNailT)
'            strTmp = NailT(j).NetName
'        End If
'   Else
'        If Trim(str_ColorNailB) = "" Then
'            Exit Sub
'        Else
'            j = Val(str_ColorNailB)
'            strTmp = NailB(j).NetName
'        End If
'   End If
'
'  'If fNailDetail Is Nothing Then Set fNailDetail = New frmNailDetail
'
'
'
'  '初始化frmNaildetail
'  If blnTop Then
'        fNailDetail.labNail.Caption = "Nail:" & NailT(j).Nail
'        fNailDetail.labNet.Caption = "Net:" & NailT(j).NetName
'  Else
'        fNailDetail.labNail.Caption = "Nail:" & NailB(j).Nail
'        fNailDetail.labNet.Caption = "Net:" & NailB(j).NetName
'  End If
'  fNailDetail.lstDetail.Clear
'
' 'Now find the net in pin top
'  For i = 0 To UBound(PinsT)
'    If PinsT(i).NetName = strTmp Then
''      PinsT(i).Color = intPinColorSel
''      str_ColorPinT = str_ColorPinT & " " & i
'       fNailDetail.lstDetail.AddItem PinsT(i).Part & "." & PinsT(i).Name   '& "  Name:" & PinsT(i).Name
'    End If
'  Next i
''    fin in pin bottom
'  For i = 0 To UBound(PinsB)
'    If PinsB(i).NetName = strTmp Then
''      PinsB(i).Color = intPinColorSel
''      str_ColorPinB = str_ColorPinB & " " & i
'       fNailDetail.lstDetail.AddItem PinsB(i).Part & "." & PinsB(i).Name   ' & "  Name:" & PinsB(i).Name
'    End If
'  Next i
''End pin
'
''Now Nail
'   'Top
'  For i = 0 To UBound(NailT)
'    If NailT(i).NetName = strTmp Then
''      NailT(i).Color = intPinColorSel
''      str_ColorNailT = str_ColorNailT & " " & i
'       fNailDetail.lstDetail.AddItem NailT(i).Nail
'    End If
'  Next i
'  'Nail Bottom
'  For i = 0 To UBound(NailB)
'    If NailB(i).NetName = strTmp Then
''      NailB(i).Color = intPinColorSel
''      str_ColorNailB = str_ColorNailB & " " & i
'       fNailDetail.lstDetail.AddItem NailB(i).Nail
'    End If
'  Next i
'
'fNailDetail.Show vbModeless, Me
'
'End Sub
Private Sub DrawNavigate()
  Dim X1 As Double
  Dim Y1 As Double
  Dim X2 As Double
  Dim Y2 As Double
 
 Dim i As Long
 On Error Resume Next
 Me.picNavigate.Cls
   
 If imgNavigate Is Nothing Or blnRedrawNvigate = True Then
 'Draw Navigate picture and save it
     Me.picNavigate.Visible = True
     Dim lngTmprate As Double
     lngTmprate = Me.Picture1.Width / (Me.picNavigate.Width)
     Set Me.picNavigate.Picture = Nothing
     Me.picNavigate.Cls
     For i = 0 To UBound(OutLineP) - 1
       If OutLineP(i).Group = OutLineP(i + 1).Group Then
        Me.picNavigate.Line (OutLineP(i).x / lngTmprate, OutLineP(i).y / lngTmprate)-(OutLineP(i + 1).x / lngTmprate, OutLineP(i + 1).y / lngTmprate), RGB(255, 0, 0)
       End If
    '   Me.Picture1.Width / 2 - pCenter.X
    
     Next i
     Set imgNavigate = Me.picNavigate.Image
     blnRedrawNvigate = False
     If blnTop = True Then
        blnTopimg = True
     Else
        blnTopimg = False
     End If
     
 End If

'Me.picNavigate.Cls
'Set Me.picNavigate.Picture = Nothing
'Exit Sub

If blnTop = blnTopimg Then
   Set Me.picNavigate.Picture = imgNavigate

Else
   Me.picNavigate.PaintPicture imgNavigate, 0, 0, Me.picNavigate.Width, Me.picNavigate.Height, Me.picNavigate.Width, 0, Me.picNavigate.Width * -1, Me.picNavigate.Height
End If
 
  X1 = (0.5 - 0.9 * pCenter.x / (dblWidth_Board * lngRate)) * Me.picNavigate.Width
  X2 = (0.5 - 0.9 * (pCenter.x - Me.Picture1.Width) / (dblWidth_Board * lngRate)) * Me.picNavigate.Width
  Y1 = (0.5 - 0.9 * pCenter.y / (dblHeight_Board * lngRate)) * Me.picNavigate.Height
  Y2 = (0.5 - 0.9 * (pCenter.y - Me.Picture1.Height) / (dblHeight_Board * lngRate)) * Me.picNavigate.Height
  
  
  
  Me.picNavigate.Line (X1, Y1)-(X2, Y2), , B

End Sub
 
Private Sub LoadLogo()
 Dim intScaleMode As Integer
 intScaleMode = Me.Picture1.ScaleMode
 Me.Picture1.ScaleMode = vbMillimeters
  
 ' Picture1.PaintPicture Me.Image1.Picture, (Me.Picture1.ScaleWidth * 100 - Me.Image1.Picture.Width) / 200, (Me.Picture1.ScaleHeight * 100 - Me.Image1.Picture.Height) / 200     ',   Me.Picture1.Width, Me.Picture1.Height
Me.Picture1.ScaleMode = intScaleMode

End Sub
Private Sub Rotate()
Dim i As Integer
Dim dblV As Double

'On Error Resume Next '

For i = 0 To UBound(OutLineP)
 dblV = OutLineP(i).x
 OutLineP(i).x = OutLineP(i).y
 OutLineP(i).y = dblV * -1
Next i
For i = 0 To UBound(PinsT)
 dblV = PinsT(i).x
 PinsT(i).x = PinsT(i).y
 PinsT(i).y = dblV * -1
 
Next i
For i = 0 To UBound(PinsB)
 dblV = PinsB(i).x
 PinsB(i).x = PinsB(i).y
 PinsB(i).y = dblV * -1
Next i

For i = 0 To UBound(PartT)
 
 dblV = PartT(i).X1
 PartT(i).X1 = PartT(i).Y1
 PartT(i).Y1 = dblV * -1
 dblV = PartT(i).X2
 PartT(i).X2 = PartT(i).Y2
 PartT(i).Y2 = dblV * -1
  
Next i
For i = 0 To UBound(PartB)
 dblV = PartB(i).X1
 PartB(i).X1 = PartB(i).Y1
 PartB(i).Y1 = dblV * -1
 dblV = PartB(i).X2
 PartB(i).X2 = PartB(i).Y2
 PartB(i).Y2 = dblV * -1

Next i
'do Nail top
For i = 0 To UBound(NailT)
 dblV = NailT(i).x
 NailT(i).x = NailT(i).y
 NailT(i).y = dblV * -1
Next i
'do Nail Bottom
For i = 0 To UBound(NailB)
 dblV = NailB(i).x
 NailB(i).x = NailB(i).y
 NailB(i).y = dblV * -1
Next i

'Do center
dblV = pCenter.x
pCenter.x = pCenter.y
pCenter.y = dblV * -1

dblV = lngRate

If blnHorizontal = True Then  '本来是横放,现在要竖放了
dblWidth_Board = dblHeight_Original
dblHeight_Board = dblWidth_Original
   blnHorizontal = False
Else
  dblWidth_Board = dblWidth_Original
  dblHeight_Board = dblHeight_Original
  blnHorizontal = True
End If

    If Picture1.Width / dblWidth_Board < Picture1.Height / dblHeight_Board Then
       lngRate = Picture1.Width / dblWidth_Board
    Else
       lngRate = Picture1.Height / dblHeight_Board
    End If
    
   dblWidth_Board = Picture1.Width / lngRate
   dblHeight_Board = Picture1.Height / lngRate
   lngRate = lngRate * 0.9   '不能把全屏都盖住，留点边 *

lngRate = lngRate * (dblV / lngRate) 'lngRate 已重新计算过了,但所有点的坐标并没有改变,所以lngRate不能正确反映情况,
                                     '名

blnRedrawNvigate = True '重画导航图

End Sub
Private Sub InitialParameter()
str_ColorPinT = ""
str_ColorPinB = ""
str_ColorPartB = ""
str_ColorPartT = ""
str_ColorNailT = ""
str_ColorNailB = ""
blnHorizontal = True

End Sub

Private Sub Timer1_Timer()
  BoardViewDevice = UCase(Trim(BoardViewDevice))
If BoardViewDevice <> "" Then
  If AnalogView = True Then
     OpenView = False
     Call FindPart_Name
     BoardViewDevice = ""
    'AnalogView = False
    
  End If
  
  If OpenView = True Then
     
     AnalogView = False
    ' Call FindNet_Name
      Call FindNail_Name
      BoardViewDevice = ""
      OpenView = False
  End If
  
  
  
 Else
'     Call ZoomPerfect
'     Call ClearPinNailColor
'     Call ClearPartColor
'     Call drawPic
End If

End Sub
