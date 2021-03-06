VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Board_xy To Board View"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Caption         =   "Boards Same"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Panel Boards"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Node range"
      Height          =   180
      Left            =   6840
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Device range"
      Height          =   180
      Left            =   6840
      TabIndex        =   8
      Top             =   360
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Add BRC Probe"
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "open board file"
      ToolTipText     =   "Double click to choose file"
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6840
      Picture         =   "frmMain.frx":1CF2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "open board_xy file"
      ToolTipText     =   "Double click to choose file"
      Top             =   600
      Width           =   6615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Output file"
      ToolTipText     =   "Double click to choose file"
      Top             =   1560
      Width           =   6615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "open fixture file"
      ToolTipText     =   "Double click to choose file"
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Label L 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ToolPath As String
Dim DeviceNode As String
Dim StopRun As Boolean
Dim BoardFileOK As Boolean
Dim FixtureFileOK As Boolean
Dim Boards As Boolean
Dim BoardsSame As Boolean
Dim JJ As Integer
Dim sBoardsName(1 To 30) As String
Dim iBoardsNumber As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Text1.Enabled = True
   Else
   Text1.Enabled = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
  Boards = True
 ' Check3.Visible = True
 Else
 Check3.Visible = False
' Check3.Value = 0
 'BoardsSame = False
  Boards = False
End If
End Sub

'Private Sub Check3_Click()
' If Check2.Value = 1 Then
'    BoardsSame = True
'    Else
'    BoardsSame = False
'    Check3.Value = 0
' End If
'End Sub

Private Sub cmdGo_Click()
 StopRun = False
If Text4.Text = "" Then Text4.Text = "Please open board file": cmdGo.Enabled = False: Exit Sub
If Dir(Text4.Text) = "" Then Text4.Text = "Please open board file": cmdGo.Enabled = False: Exit Sub
If Text2.Text = "" Then Text2.Text = "Please open board_xy file": cmdGo.Enabled = False: Exit Sub
If Dir(Text2.Text) = "" Then Text2.Text = "Please open board_xy file": cmdGo.Enabled = False: Exit Sub
If Check1.Value = 1 Then
   If Text1.Text = "" Then Text1.Text = "Please open fixture file": cmdGo.Enabled = False: Exit Sub
   If Dir(Text1.Text) = "" Then Text1.Text = "Please open fixture file": cmdGo.Enabled = False: Exit Sub
End If
cmdGo.Enabled = False
Check1.Enabled = False
Check2.Enabled = False

If Boards = True Then
   Check3.Enabled = False
End If
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
If Option1.Value = True Then
  Call Read_Board
   If BoardFileOK = False Then
      e = MsgBox("Board file is Node range!" & Chr(13) & Chr(10) & "Advise use board consultant set (Board File List Format: Device) range!", 16, "Warning")
      StopRun = True
     GoTo EX
   End If
  Else
   If Option2 = True Then
      Call Read_Board_Node
   If BoardFileOK = False Then
      e = MsgBox("Board file is Device range!" & Chr(13) & Chr(10) & "Advise use board consultant set (Board File List Format: Node) range!", 16, "Warning")
      StopRun = True
     GoTo EX
   End If
      
   End If
End If
If StopRun = True Then GoTo EX
   If Boards = False Then
      Call Read_Board_xy
    Else
      Call Read_Boards_xy
   End If
 If StopRun = True Then GoTo EX

If Check1.Value = 1 Then '
   If Text1.Text = "" Then Text1.Text = "Please open fixture file": cmdGo.Enabled = False: Exit Sub
   If Dir(Text1.Text) = "" Then Text1.Text = "Please open fixture file": cmdGo.Enabled = False: Exit Sub
   Call Read_fixture
    If StopRun = True Then GoTo EX

    Call Add_BRC
End If
EX:
Check1.Enabled = True
Check2.Enabled = True
If Boards = True Then
   Check3.Enabled = True
End If


cmdGo.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
On Error Resume Next
Kill ToolPath & "FixToBoardView\ReadBoard\*.tmp"
Kill ToolPath & "FixToBoardView\ReadFixture\*.BRC"
End Sub



Private Sub Command1_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
On Error Resume Next

ToolPath = App.Path
If Right(ToolPath, 1) <> "\" Then ToolPath = ToolPath & "\"
MkDir ToolPath & "FixToBoardView"
Open ToolPath & "FixToBoardView\NotDelete.sys" For Output As #77
Text3.Text = ToolPath & "FixToBoardView\NoBrcBoardView.bv2"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill ToolPath & "FixToBoardView\ReadBoard\*.tmp"
Kill ToolPath & "FixToBoardView\ReadFixture\*.BRC"
Close #77
Kill ToolPath & "FixToBoardView\NotDelete.sys"
End Sub

Private Sub Option2_Click()
A = MsgBox("Use Node range speed drop!" & Chr(13) & Chr(10) & "Advise use board consultant set (Board File List Format: Device) range!", vbQuestion, "Warning")
End Sub

Private Sub Text1_DblClick()
On Error GoTo errH
With Me.CommonDialog1
    .FileName = "fixture"
    .CancelError = True
    .Filter = "*.*|*.*"
    .ShowOpen
    Me.Text1.Text = .FileName
End With
 If Text1.Text = "" Then Text1.Text = "please open fixture file": cmdGo.Enabled = False: Exit Sub
 If Dir(Text1.Text) = "" Then
    cmdGo.Enabled = False
    Text1.Text = "Please open fixture file"
    MsgBox "No file!", vbCritical
    cmdGo.Enabled = False
    Exit Sub
   Else
        Open Text1.Text For Input As #1
          Line Input #1, tmpStr
        Close #1
         If Left(tmpStr, 9) <> "!!!!   13" Then
            MsgBox "File format error,no fixture type!", vbCritical
            Text1.Text = "Please open fixture file"
            cmdGo.Enabled = False
            Exit Sub
           Else
             If Text2.Text <> "Please open board_xy file" And Text2.Text <> "open board_xy file" And Text4.Text <> "Please open board file" And Text4.Text <> "open board file" Then
               cmdGo.Enabled = True
             End If
         End If

 End If
errH:
End Sub

Private Sub Text2_DblClick()
On Error GoTo errH
With Me.CommonDialog1
    .FileName = "board_xy"
    .CancelError = True
    .Filter = "*.*|*.*"
    .ShowOpen
    Me.Text2.Text = .FileName
End With
 If Text2.Text = "" Then Text2.Text = "Please open board_xy file": cmdGo.Enabled = False: Exit Sub
 If Dir(Text2.Text) = "" Then
    cmdGo.Enabled = False
    Text2.Text = "Please open board_xy file"
    MsgBox "No file!", vbCritical
    cmdGo.Enabled = False
    Exit Sub
   Else
        Open Text2.Text For Input As #1
          Line Input #1, tmpStr
        Close #1
         If Left(tmpStr, 9) <> "!!!!   15" Then
            MsgBox "File format error,no board_xy type!", vbCritical
            Text2.Text = "Please open board_xy file"
            cmdGo.Enabled = False
            Exit Sub
           Else
           If Text4.Text <> "Please open board file" And Text4.Text <> "open board file" Then
             cmdGo.Enabled = True
           End If
         End If
 End If

errH:
End Sub

Private Sub Text3_DblClick()
On Error GoTo errH
With Me.CommonDialog1
 
    .FileName = Trim(Text3.Text) & ".bv2"
    .CancelError = True
    .Filter = "*.bv2|*.bv2"
    .ShowSave
    Me.Text3.Text = .FileName
End With
 
errH:
End Sub
Private Sub Read_Board_xy()
 On Error Resume Next

 Dim MyStr As String
 Dim NoProbeTrue As Boolean
 Dim FOutLine As Integer
 Dim FOutLineStr As String
 Dim FindOutLine As Boolean
 Dim OutLine() As String
 Dim tmpNode1() As String
 Dim NodeXY() As String
 Dim DvsName As String
 Dim DvsPin As String
 Dim FindDvsTrue As Boolean
 Dim NodeName As String
 Dim FindNodeXYTrue As Boolean
 Dim NodeX As String
 Dim NodeY As String
 Dim tmpNode2()  As String
 Dim bTB As String
 Dim NodeNumber As String
 Dim i As Integer
 Dim NextNode As Boolean
 Dim tmpNode3 As String
 Dim tmpDvs1() As String
 Dim tmpDvs2() As String
 Dim tmpDvs3() As String
 Dim DvsX As String
 Dim DvsY As String
 Dim SameDvs As String
 Dim tt As Integer
 Dim ee As Integer
 Dim yy As Integer
 Dim MyDevice As String
 Dim tmpStr1() As String
 Dim BoardsGe As Integer
 FOutLine = 0
 BoardsGe = 0
 yy = 0
 tt = 1
 i = 0
Open Text2.Text For Input As #1
  Open Text3.Text For Output As #2
    Print #2, "#Layout#"
    Print #2, "X,Y,R,Group"
    
  Do Until EOF(1) '#1
      If StopRun = True Then Exit Do
      
      Line Input #1, MyStr
      MyStr = Trim(MyStr)
       JJ = JJ + 1
       NoProbeTrue = False
      If MyStr <> "" Then
       
            If MyStr = "OUTLINE" Then
               FindOutLine = True
                
            End If
            If MyStr <> "OUTLINE" And FindOutLine = True Then
               OutLine = Split(MyStr, ",")

               OutLine(0) = Trim(OutLine(0))
                OutLine(0) = (Val(OutLine(0))) / 12000
               
               
               'Line x
               ' OutLine(0) = Left(Trim(OutLine(0)), Len(OutLine(0)) - 3) & "." & Right(Trim(OutLine(0)), 3)
               
               
               
               If Left(OutLine(0), 1) = "." Then
                  OutLine(0) = "0" & OutLine(0)
               End If
               OutLine(1) = Trim(OutLine(1))
               
               
               
               If Right(OutLine(1), 1) = ";" Then
                 FindOutLine = False
                  OutLine(1) = Trim(Left(OutLine(1), Len(OutLine(1)) - 1))
               End If
              
               OutLine(1) = (Val(OutLine(1))) / 12000
            
            'line Y
                'OutLine(1) = Left(Trim(OutLine(1)), Len(OutLine(1)) - 3) & "." & Right(Trim(OutLine(1)), 3)
               If Left(OutLine(1), 1) = "." Then
                  OutLine(1) = "0" & OutLine(1)
               End If
               
               
               If FOutLine = 0 Then
                 FOutLineStr = OutLine(0) & "," & OutLine(1) & ",0.000,2"
               End If
               Print #2, OutLine(0) & "," & OutLine(1) & ",0.000,2"
               FOutLine = FOutLine + 1
               If FindOutLine = False Then
                  Print #2, FOutLineStr
                  FOutLineStr = ""
                  OutLine(1) = ""
                  OutLine(0) = ""
                  Print #2, "#Nail#"
                  Print #2, "Nail,X,Y,Type,Grid,TB,Net,NetName"
                 
               End If
               
            End If
            If UCase(Left(MyStr, 5)) = "NODE " Then
                NodeName = Trim(Right(MyStr, Len(MyStr) - 4))
                
                NodeNumber = "#" & i
                'no probe
                If InStr(UCase(NodeName), "NO_ACCESS;") <> 0 Then
                   tmpNode1 = Split(NodeName, " ")
                   NodeName = Trim(tmpNode1(0))
                   NoProbeTrue = True
                   
''great 20080521
                   
                   If NoProbeTrue = True Then

                      Print #2, "$0,0,0,NO_PROBE,E1,(B),#0," & NodeName

                   End If
                   
                  Else
                   FindNodeXYTrue = True
                   
                End If
                
            End If
            If FindNodeXYTrue = True And UCase(Left(MyStr, 5)) <> "NODE " Then
               If Right(UCase(MyStr), 10) = "MANDATORY;" <> 0 Or InStr(UCase(MyStr), "TOP;") <> 0 Or Right(UCase(MyStr), 10) = "PREFERRED;" <> 0 Then
                  NodeXY = Split(MyStr, ",")
                  NodeX = Trim(NodeXY(0))
                   NodeX = (Val(NodeX)) / 12000
                  'NodeX = Left(Trim(NodeX), Len(NodeX) - 3) & "." & Right(Trim(NodeX), 3)
               
               
               If Left(NodeX, 1) = "." Then
                  NodeX = "0" & NodeX
               End If
                  
                  '(T) and (B)--
                  If InStr(UCase(NodeXY(1)), "TOP") <> 0 Then
                     bTB = "(T)"
                    Else
                     bTB = "(B)"
                  End If
                  '-------
                  If InStr(UCase(NodeXY(1)), " ") <> 0 Then
                     tmpNode2 = Split(Trim(NodeXY(1)), " ")
                     NodeY = Trim(tmpNode2(0))
                     NodeY = (Val(NodeY)) / 12000
                    ' NodeY = Left(Trim(tmpNode2(0)), Len(tmpNode2(0)) - 3) & "." & Right(Trim(tmpNode2(0)), 3)
               
               
                         If Left(NodeY, 1) = "." Then
                             NodeY = "0" & NodeY
                         End If
                    
                    Else
                     NodeY = Replace(Trim(NodeXY(1)), ";", "")
                      NodeY = (Val(NodeY)) / 12000
                  '  NodeY = Left(Trim(NodeY), Len(NodeY) - 3) & "." & Right(Trim(NodeY), 3)
                         If Left(NodeY, 1) = "." Then
                             NodeY = "0" & NodeY
                         End If
                    
                  End If
                  If NodeX <> "" And NodeY <> "" And bTB <> "" Then
                    Print #2, "$0," & NodeX & "," & NodeY & ",5,E1," & bTB & "," & NodeNumber & "," & NodeName
                    'Print #2, "$0," & NodeX & ","; NodeY & ",5,E1," & bTB & "," & NodeNumber & "," & NodeName
                     'tmpNode3 = NodeName
                     'NextNode = True
                    'FindNodeXYTrue = False
                    

                        

                    
                    
                  End If
               End If
            End If
            If UCase(MyStr) = "OTHER" Then
              FindNodeXYTrue = False
              Print #2, "#Pin#"
              Print #2, "Part,TB,Pin,Name,X,Y,Layer,Netname"
               FindDvsTrue = True
            End If
            If FindDvsTrue = True And UCase(MyStr) <> "OTHER" And UCase(MyStr) <> "ALTERNATES" And UCase(MyStr) <> "DEVICES" Then
                tmpDvs1 = Split(MyStr, ",")
               ' DvsX = Left(Trim(tmpDvs1(0)), Len(tmpDvs1(0)) - 3) & "." & Right(Trim(tmpDvs1(0)), 3)
                         If Left(DvsX, 1) = "." Then
                             DvsX = "0" & DvsX
                         End If
                
                 DvsX = (Val(Trim(tmpDvs1(0)))) / 12000
                If InStr(UCase(tmpDvs1(1)), "TOP") <> 0 Then
                   bTB = "(T)"
                  Else
                   bTB = "(B)"
                End If
                tmpDvs2 = Split(Trim(tmpDvs1(1)), " ")
                
                DvsY = (Val(Trim(tmpDvs2(0)))) / 12000
                
               'DvsY = Left(Trim(tmpDvs2(0)), Len(tmpDvs2(0)) - 3) & "." & Right(Trim(tmpDvs2(0)), 3)
                         If Left(DvsY, 1) = "." Then
                             DvsY = "0" & DvsY
                         End If
                
                
                tmpDvs3 = Split(Trim(tmpDvs2(1)), ".")
                DvsName = Trim(tmpDvs3(0))
                
                If yy = 0 Then
                   SameDvs = DvsName
                  yy = 1
                  ee = 0
                End If
                If SameDvs <> DvsName Then
                   yy = 0
                End If
                DvsPin = Trim(tmpDvs3(1))
                tt = Val(DvsPin)
                If tt = 0 Then
                   ee = ee + 1
                   Else
                   ee = Val(DvsPin)
                End If
                
'readboard floder   file
               If Option1 = True Then
                    If Dir(ToolPath & "FixToBoardView\ReadBoard\" & DvsName & ".tmp") <> "" Then
                       Open ToolPath & "FixToBoardView\ReadBoard\" & DvsName & ".tmp" For Input As #6
                          Do Until EOF(6)
                             Line Input #6, MyDevice
                             MyDevice = Trim(MyDevice)
                             If MyDevice <> "" Then
                                If InStr(MyDevice, ".") <> 0 Then
                                   tmpStr1 = Split(MyDevice, ".")
                                   If UCase(tmpStr1(0)) = UCase(DvsPin) Then
                                      DeviceNode = Trim(tmpStr1(1))
                                      Exit Do
                                   End If
                                   
                                End If
                             End If
                          Loop
                       Close #6
                     End If
                   Else
                    If Option2 = True Then
                        If Dir(ToolPath & "FixToBoardView\ReadBoard\" & DvsName & "." & DvsPin & ".tmp") <> "" Then
                           Open ToolPath & "FixToBoardView\ReadBoard\" & DvsName & "." & DvsPin & ".tmp" For Input As #6
                                 Line Input #6, MyDevice
                                 MyDevice = Trim(MyDevice)
                                 If MyDevice <> "" Then
                                     DeviceNode = Trim(MyDevice)
    
                                 End If
                           Close #6
                        End If
                   End If
                End If
                
                
                
                 If DeviceNode = "" Then DeviceNode = "NotFindNode"
                Print #2, DvsName & "," & bTB & "," & ee & "," & DvsPin; "," & DvsX & "," & DvsY & ","; "2" & "," & DeviceNode
                DvsName = ""
                bTB = ""
                DvsX = ""
                DvsY = ""
                DeviceName = ""
                
            End If
            If MyStr = "DEVICES" Then
               FindDvsTrue = False
               FindNodeXYTrue = False
            End If
          Else
           FindNodeXYTrue = False
            NodeX = ""
            NodeY = ""
            NodeNumber = ""
            NodeName = ""
            
            bTB = ""
      End If
      DoEvents
      L.Caption = "Read board_xy file line: " & JJ
  Loop '#1
  Close #2
Close #1
Kill ToolPath & "FixToBoardView\ReadBoard\*.tmp"
L.Caption = "Save BoardView.bv2 file ok"
MsgBox "NoBrcBoardView.bv2 creat OK !", vbQuestion
 Exit Sub
EX:
 MsgBox Err.Description, vbCritical
End Sub
Private Sub Read_fixture()
 ' On Error GoTo ex
Dim FindStart As Boolean
Dim PinsStart As Boolean
Dim Node As String
Dim Brc As String
Dim BrcFindOk As Boolean
Dim MilFindOk As Boolean
Dim MyStr As String
Dim Probes As String
Dim FindProbe As Boolean
Dim StrMil As String

'Dim FindPin As Boolean
Dim i As Integer
Dim T As Integer
T = 1
i = 0
 On Error Resume Next
 MkDir ToolPath & "FixToBoardView\ReadFixture"
 Kill ToolPath & "FixToBoardView\ReadFixture\*.BRC"
    Open Text1.Text For Input As #1
      ' Open ToolPath & "FixToBoardView\Net.tmp" For Output As #7
       '  Print #7, "!# Read file is in fxture file"
       '  Print #7, "$0" & "," & "#" & T & "," & "GND"
       Do Until EOF(1)
         If StopRun = True Then: Exit Do
         Line Input #1, MyStr
         i = i + 1
         MyStr = Trim(MyStr)
         If MyStr <> "" Then
            If Left(UCase(MyStr), 5) = "NODE " Then
               If Left(UCase(MyStr), 5) = "NODE " And Trim(Right(UCase(MyStr), Len(MyStr) - 5)) <> "GND GROUND" Then
                   Node = Trim(Right(UCase(MyStr), Len(MyStr) - 5))
                   Node = Replace(Node, """", "")
                   If Boards = True Then
                       Node = Replace(Node, Left(Node, InStr(Node, "%")), "")
                      Node = Right(Node, Left(Node, InStr(Node, "%")))
                   End If
                   T = T + 1
                    FindStart = True
               End If
            End If
            If FindStart = True And UCase(MyStr) = "PINS" Then
               PinsStart = True
            End If
            If PinsStart = True And Right(MyStr, 1) = ";" And Left(UCase(MyStr), 5) <> "NODE " And UCase(MyStr) <> "PINS" Then
               Brc = Left(MyStr, Len(MyStr) - 1)
               If Len(Brc) > 6 Then
                 Brc = Trim(Left(Brc, 6))
                 BrcFindOk = True
                 PinsStart = False
               End If
                 BrcFindOk = True
                 PinsStart = False
'               If Brc = "" Then
'                  Print #7, "$0" & "," & "#" & T & "," & Node
'                    Open ToolPath & "FixToBoardView\ReadFixture\" & Node & ".BRC" For Output As #8
'                      Print #8, "$0" & "," & "#" & T & "," & Node
'                    Close #8
'                 Else
'                  Print #7, "$" & Brc & "," & "#" & T & "," & Node
'                    Open ToolPath & "FixToBoardView\ReadFixture\" & Node & ".BRC" For Output As #8
'                      Print #8, "$" & Brc & "," & "#" & T & "," & Node
'                    Close #8
'               End If
'               PinsStart = False
'               FindStart = False
'               Brc = ""
'               Node = ""
            End If
            If FindStart = True And Trim(UCase(MyStr)) = "PROBES" Then
              FindProbe = True
            End If
            If FindStart = True And FindProbe = True And Left(Trim(UCase(MyStr)), 1) = "P" And Left(UCase(MyStr), 5) <> "NODE " And Right(Trim(MyStr), 1) = ";" And Trim(UCase(MyStr)) <> "PROBES" Then
               If InStr(UCase(MyStr), "50MIL") <> 0 Then
                  StrMil = "50MIL"
                 Else
                   If InStr(UCase(MyStr), "75MIL") <> 0 Then
                      StrMil = "75MIL"
                   End If
                   If InStr(UCase(MyStr), "75MIL") = 0 And InStr(UCase(MyStr), "50MIL") = 0 Then
                      StrMil = "100MIL"
                   End If
               End If
               If StrMil = "" Then
                 StrMil = "Unknown"
               End If
               FindProbe = False
               MilFindOk = True
            End If
            
            If BrcFindOk = True And MilFindOk = True Then
               If Brc = "" Then
                 ' Print #7, "$0" & "," & "#" & T & "," & Node
                    Open ToolPath & "FixToBoardView\ReadFixture\" & Node & ".BRC" For Output As #8
                      Print #8, "$0" & "," & "#" & T & "," & Node & "," & StrMil
                    Close #8
                 Else
                 ' Print #7, "$" & Brc & "," & "#" & T & "," & Node
                    Open ToolPath & "FixToBoardView\ReadFixture\" & Node & ".BRC" For Output As #8
                      Print #8, "$" & Brc & "," & "#" & T & "," & Node & "," & StrMil
                    Close #8
               End If
           
               PinsStart = False
               FindStart = False
               FindProbe = False
               Brc = ""
               Node = ""
               StrMil = ""
               BrcFindOk = False
               MilFindOk = False
           End If
           

               DoEvents
              
              L.Caption = "Read fixture file line: " & i
         End If
      Loop
   '  Close #7
    Close #1

    
    L.Caption = "Fixture file read ok!:"
    
    
   ' MsgBox "Save file OK", vbInformation
 Exit Sub
EX:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub Text4_DblClick()
On Error GoTo errH
With Me.CommonDialog1
    .FileName = "board"
    .CancelError = True
    .Filter = "*.*|*.*"
    .ShowOpen
    Me.Text4.Text = .FileName
End With
 If Text4.Text = "" Then Text4.Text = "Please open board file": cmdGo.Enabled = False: Exit Sub
 If Dir(Text4.Text) = "" Then
    cmdGo.Enabled = False
    Text4.Text = "Please open board file"
    MsgBox "No file!", vbCritical
    cmdGo.Enabled = False
    Exit Sub
   Else
        Open Text4.Text For Input As #1
          Line Input #1, tmpStr
        Close #1
         If Left(tmpStr, 9) <> "!!!!   12" Then
            MsgBox "File format error,no board file type!", vbCritical
            Text4.Text = "Please open board file"
            cmdGo.Enabled = False
            Exit Sub
           Else
           If Text2.Text <> "Please open board_xy file" And Text2.Text <> "open board_xy file" Then
             cmdGo.Enabled = True
           End If
         End If
 End If

errH:

End Sub
Private Sub Read_Board()
Dim BoardsNameOk As Boolean
Dim FindBoards As Boolean
Dim MyStr As String
Dim tmpStr() As String
Dim FindPin As Boolean
Dim TmpDvs As String
Dim FindStart As Boolean
Dim haha
Dim BoardGe As Integer
Dim q As Integer
 On Error Resume Next
 iBoardsNumber = 1
 q = 1
 MkDir ToolPath & "FixToBoardView\ReadBoard"
 Kill ToolPath & "FixToBoardView\ReadBoard\*.tmp"
 BoardGe = 0
  Open Text4.Text For Input As #4
  Do Until EOF(4)
    If StopRun = True Then: BoardFileOK = False: Exit Do
    Line Input #4, MyStr
    MyStr = Trim(MyStr)
    If MyStr <> "" Then
    'boards
'        If Boards = True Then
'          If UCase(MyStr) = "BOARDS" Then
'             BoardsNameOk = True
'          End If
'          haha = iBoardsNumber
'          If BoardsNameOk = True And UCase(MyStr) <> "boards" Then
'             If Left(MyStr, Len(haha)) = iBoardsNumber And Right(MyStr, 1) = ";" Then
'                 sBoardsName(iBoardsNumber) = Trim(Mid(MyStr, Len(haha) + 1, Len(MyStr) - (Len(haha) + 1)))
'                 iBoardsNumber = iBoardsNumber + 1
'             End If
'
'             If Left(UCase(MyStr), 6) = "BOARD " Then
'                tmpp = Trim(Replace(MyStr, "BOARD", ""))
'                If tmpp = sBoardsName(q) Then
'                   FindBoards = True
'                End If
'             End If
'          End If
'        End If
      
          

      
          If UCase(MyStr) = "DEVICES" Then
             FindPin = True
             BoardGe = BoardGe + 1
          End If
          
          If Boards = True And Replace(UCase(MyStr), " ", "") = "ENDBOARD" Then
             FindPin = False
          End If
          
          If FindPin = True And UCase(MyStr) <> "DEVICES" And InStr(MyStr, ".") = 0 And UCase(MyStr) <> "END" Then
              TmpDvs = MyStr
              FindStart = True
              BoardFileOK = True
                If Boards = True And FindBoards = True And BoardsNameOk = True Then
                   If Boards = True Then
                     Open ToolPath & "FixToBoardView\ReadBoard\" & BoardGe & "%" & TmpDvs & "," & sBoardsName(q) & ".tmp" For Append As #5
                     Else
                     Open ToolPath & "FixToBoardView\ReadBoard\" & TmpDvs & "," & sBoardsName(q) & ".tmp" For Append As #5
                   End If
                   Else
                   
                   If Boards = True Then
                     Open ToolPath & "FixToBoardView\ReadBoard\" & BoardGe & "%" & TmpDvs & ".tmp" For Append As #5
                     Else
                       Open ToolPath & "FixToBoardView\ReadBoard\" & TmpDvs & ".tmp" For Append As #5

                   
                   End If
               End If
          End If
          
          If FindPin = True And FindStart = True And InStr(MyStr, ".") <> 0 Then
             If Right(MyStr, 1) = ";" Then
                MyStr = Trim(Left(MyStr, Len(MyStr) - 1))
                FindStart = False

                     Print #5, MyStr

                TmpDvs = ""
                Close #5
             End If
             Print #5, MyStr

          End If
         Else
'         'boards
'          If Boards = True And FindBoards = True And BoardsNameOk = True Then
'                If UCase(MyStr) = "END BOARD" Then
'                    q = q + 1
'                   FindBoards = False
'                   FindPin = False
'                   If q > iBoardsNumber Then
'                     BoardsNameOk = False
'                     Exit Do
'                   End If
'                End If
'              Else
'                If UCase(MyStr) = "END" Then
'                   FindPin = False
'                   Exit Do
'                End If
'         End If
   End If
'          If Boards = True And FindBoards = True And BoardsNameOk = True Then
'                If UCase(MyStr) = "END BOARD" Then
'                    q = q + 1
'                   FindBoards = False
'                   FindPin = False
'                   If q > iBoardsNumber Then
'                     BoardsNameOk = False
'                     Exit Do
'                   End If
'                End If
'          End If
        DoEvents
        i = i + 1
        L.Caption = "Read board file line: " & i & " (BoardType: Device range)"
  Loop
  Close #4
 ' MsgBox "OK"
Exit Sub

errH:
MsgBox Err.Description, vbCritical
End Sub
Private Sub Add_BRC()
Dim MyStr As String
Dim tmpStr() As String
Dim tmpStr1() As String
Dim Brc As String
Dim i As String
Dim NodeNumber As String
 On Error Resume Next
 Dim TmpNodeStr As String
 
 Open Text3.Text For Input As #9
 Open ToolPath & "FixToBoardView\BrcBoardView.bv2" For Output As #10
   Do Until EOF(9)
     If StopRun = True Then Exit Do
     Line Input #9, MyStr
        MyStr = Trim(MyStr)
           If MyStr <> "" Then
               If Left(MyStr, 1) = "$" Then
                   tmpStr = Split(MyStr, ",")
                   If Dir(ToolPath & "FixToBoardView\ReadFixture\" & tmpStr(7) & ".BRC") <> "" Then
                         Open ToolPath & "FixToBoardView\ReadFixture\" & tmpStr(7) & ".BRC" For Input As #11
                             Line Input #11, TmpNodeStr
                         Close #11
                         tmpStr1 = Split(TmpNodeStr, ",")
                         Brc = Trim(tmpStr1(0))
                         Print #10, Brc & "," & tmpStr(1) & "," & tmpStr(2) & "," & tmpStr1(3) & "," & tmpStr(4) & "," & tmpStr(5) & "," & tmpStr1(1) & "," & tmpStr(7)
                      Else
                        Print #10, MyStr
                       
                   End If
                  Else
                   Print #10, MyStr
                    
               End If
           End If
        DoEvents
        
        L.Caption = "Add BRC ,Please wait..."
   Loop
 Close #9
 Close #10
 Kill ToolPath & "FixToBoardView\ReadFixture\*.BRC"
  L.Caption = "BrcBoardView.bv2 file create ok!"
  MsgBox L.Caption, vbInformation
End Sub




Private Sub Read_Board_Node()

Dim MyStr As String
Dim tmpStr() As String
Dim FindPin As Boolean
Dim TmpDvs As String
Dim FindStart As Boolean
Dim BoardGe  As Integer
 On Error Resume Next
 MkDir ToolPath & "FixToBoardView\ReadBoard"
 Kill ToolPath & "FixToBoardView\ReadBoard\*.tmp"
  Open Text4.Text For Input As #4
  Do Until EOF(4)
    If StopRun = True Then: BoardFileOK = False: Exit Do
    Line Input #4, MyStr
    MyStr = Trim(MyStr)
    If MyStr <> "" Then
          If UCase(MyStr) = "CONNECTIONS" Then
             FindPin = True
             BoardGe = BoardGe + 1
          End If

          If FindPin = True And UCase(MyStr) <> "CONNECTIONS" And InStr(MyStr, ".") = 0 And UCase(MyStr) <> "END" And UCase(MyStr) <> "DEVICES" Then
              TmpDvs = Replace(MyStr, """", "")
              FindStart = True
              ' Open ToolPath & "FixToBoardView\ReadBoard\" & TmpDvs & ".tmp" For Append As #5
          End If
          If FindPin = True And FindStart = True And InStr(MyStr, ".") <> 0 Then
             If Right(MyStr, 1) = ";" Then
                MyStr = Trim(Left(MyStr, Len(MyStr) - 1))
                FindStart = False
                BoardFileOK = True
              If Boards = True Then
                  Open ToolPath & "FixToBoardView\ReadBoard\" & BoardGe & "%" & MyStr & ".tmp" For Append As #5
                Else
                
                Open ToolPath & "FixToBoardView\ReadBoard\" & MyStr & ".tmp" For Append As #5
              End If    '

                      Print #5, TmpDvs
               
                Close #5
                TmpDvs = ""
            End If
              If Boards = True Then
                  Open ToolPath & "FixToBoardView\ReadBoard\" & BoardGe & "%" & MyStr & ".tmp" For Append As #5
                Else
                Open ToolPath & "FixToBoardView\ReadBoard\" & MyStr & ".tmp" For Append As #5
              End If

                      Print #5, TmpDvs

                Close #5
          End If
        Else
         If FindPin = True And MyStr = "" Or UCase(MyStr) = "DEVICES" Or UCase(MyStr) = "END" Then
            FindPin = False
            Exit Do
         End If
    End If
    DoEvents
    i = i + 1
    L.Caption = "Read board file line: " & i & " (BoardType: Node range)"
  Loop
  Close #4
 ' MsgBox "OK"
Exit Sub

errH:
MsgBox Err.Description, vbCritical

End Sub




Private Sub Read_Boards_xy()
 On Error Resume Next

 Dim MyStr As String
 Dim NoProbeTrue As Boolean
 Dim FOutLine As Integer
 Dim FOutLineStr As String
 Dim FindOutLine As Boolean
 Dim OutLine() As String
 Dim tmpNode1() As String
 Dim NodeXY() As String
 Dim DvsName As String
 Dim DvsPin As String
 Dim FindDvsTrue As Boolean
 Dim NodeName As String
 Dim FindNodeXYTrue As Boolean
 Dim NodeX As String
 Dim NodeY As String
 Dim tmpNode2()  As String
 Dim bTB As String
 Dim NodeNumber As String
 Dim i As Integer
 Dim NextNode As Boolean
 Dim tmpNode3 As String
 Dim tmpDvs1() As String
 Dim tmpDvs2() As String
 Dim tmpDvs3() As String
 Dim DvsX As String
 Dim DvsY As String
 Dim SameDvs As String
 Dim tt As Integer
 Dim ee As Integer
 Dim yy As Integer
 Dim MyDevice As String
 Dim tmpStr1() As String
 Dim BoardsGe As Integer
 Dim BoardsX
 Dim FindBoardOk As Boolean
 Dim BoardsY
  Dim strBoardsX(500)
  Dim strBoardsY(500)
  Dim strBoardsName(500) As String
  Dim strBoardsNumber(500)
  Dim strTmpBoardsXY(500) As String
  Dim strTmpBoartStr As String
  Dim intNumbet As Integer
  Dim bStartBoards As Boolean
  Dim strTmpBoardName As String
  Dim strMyTmp As String
  Dim intCurrNumber As Integer
  intNumbet = 0
 FOutLine = 0
 intCurrNumb = 0
 BoardsGe = 0
 strTmpBoartStr = ""
 yy = 0
 tt = 1
 i = 0
 Open ToolPath & "FixToBoardView\ReadBoard\Layout.txt" For Output As 11
 Close #11
  Open ToolPath & "FixToBoardView\ReadBoard\Nail.txt" For Output As 11
 Close #11
   Open ToolPath & "FixToBoardView\ReadBoard\Pin.txt" For Output As 11
 Close #11

  Open ToolPath & "FixToBoardView\ReadBoard\Layout.txt" For Append As 11
    Open ToolPath & "FixToBoardView\ReadBoard\Nail.txt" For Append As 22
  Open ToolPath & "FixToBoardView\ReadBoard\Pin.txt" For Append As 33

 
Open Text2.Text For Input As #1
  Open Text3.Text For Output As #2
     Print #2, "#Layout#"
     Print #2, "X,Y,R,Group"
    
  Do Until EOF(1) '#1
      If StopRun = True Then Exit Do
      MyStr = ""
      Line Input #1, MyStr
      MyStr = Trim(MyStr)
       JJ = JJ + 1
       NoProbeTrue = False
      If MyStr <> "" Then
            

            If UCase(MyStr) = "OUTLINE" And BoardsGe = 0 Then
               FindOutLine = True
                BoardsGe = BoardsGe + 1
            End If
             If BoardsGe > 1 Then
               FindOutLine = False
             End If
            
            If UCase(MyStr) <> "OUTLINE" And FindOutLine = True Then
               OutLine = Split(MyStr, ",")

               OutLine(0) = Trim(OutLine(0))
                OutLine(0) = (Val(OutLine(0))) / 12000
               
               
               'Line x
               ' OutLine(0) = Left(Trim(OutLine(0)), Len(OutLine(0)) - 3) & "." & Right(Trim(OutLine(0)), 3)
               
               
               
               If Left(OutLine(0), 1) = "." Then
                  OutLine(0) = "0" & OutLine(0)
               End If
               OutLine(1) = Trim(OutLine(1))
               
               
               
               If Right(OutLine(1), 1) = ";" Then
                 FindOutLine = False
                  OutLine(1) = Trim(Left(OutLine(1), Len(OutLine(1)) - 1))
               End If
              
               OutLine(1) = (Val(OutLine(1))) / 12000
            
            'line Y
                'OutLine(1) = Left(Trim(OutLine(1)), Len(OutLine(1)) - 3) & "." & Right(Trim(OutLine(1)), 3)
               If Left(OutLine(1), 1) = "." Then
                  OutLine(1) = "0" & OutLine(1)
               End If
               
               
               If FOutLine = 0 Then
                 FOutLineStr = OutLine(0) & "," & OutLine(1) & ",0.000,2"
               End If
               Print #11, OutLine(0) & "," & OutLine(1) & ",0.000,2"
               FOutLine = FOutLine + 1
               If FindOutLine = False Then
                  Print #11, FOutLineStr
                  FOutLineStr = ""
                  OutLine(1) = ""
                  OutLine(0) = ""
                 ' Print #22, "#Nail#"
                '  Print #22, "Nail,X,Y,Type,Grid,TB,Net,NetName"
                 
               End If
               
            End If
            'boards
             
            If UCase(MyStr) = "BOARDS" Then
              FindBoardOk = True
            End If
            If FindBoardOk = True And UCase(MyStr) <> "BOARDS" Then
                strTmpBoartStr = MyStr
                  strTmpBoartStr = Replace(strTmpBoartStr, """", "")
                  strBoardsNumber(intNumbet) = Left(strTmpBoartStr, InStr(strTmpBoartStr, " "))
                  strTmpBoartStr = Trim(Right(strTmpBoartStr, Len(strTmpBoartStr) - Len(strBoardsNumber(intNumbet))))
                  strBoardsNumber(intNumbet) = Trim(strBoardsNumber(intNumbet))
                  'strBoardsNumber(intNumbet) = Trim(Replace(strBoardsNumber(intNumbet), """", ""))
                  strBoardsName(intNumbet) = Left(strTmpBoartStr, InStr(strTmpBoartStr, " "))
                  strTmpBoartStr = Trim(Right(strTmpBoartStr, Len(strTmpBoartStr) - Len(strBoardsName(intNumbet))))
                 ' strBoardsName(intNumbet) = Trim(Replace(strBoardsName(intNumbet), """", ""))
                  strBoardsName(intNumbet) = Trim(strBoardsName(intNumbet))
                  strBoardsX(intNumbet) = Left(strTmpBoartStr, InStr(strTmpBoartStr, " "))
                  strTmpBoartStr = Trim(Right(strTmpBoartStr, Len(strTmpBoartStr) - Len(strBoardsX(intNumbet))))
                  strBoardsX(intNumbet) = Trim(strBoardsX(intNumbet))
                  strBoardsY(intNumbet) = Left(strTmpBoartStr, InStr(strTmpBoartStr, " "))
                  strTmpBoartStr = Trim(Right(strTmpBoartStr, Len(strTmpBoartStr) - Len(strBoardsY(intNumbet))))
                  strBoardsY(intNumbet) = Trim(strBoardsX(intNumbet))
                  
                 If intNumbet <> 0 And strBoardsName(intNumbet) <> "" Then
                   BoardsX = InputBox("Please input board ( " & strBoardsName(intNumbet) & " )X,Y offset ! ( 0-5000000)", "Board X,Y offset + 0-5000000", 0)
                   BoardsX = Val(BoardsX)
                   strBoardsX(intNumbet) = strBoardsX(intNumbet) + BoardsX
                   strBoardsY(intNumbet) = strBoardsY(intNumbet) + BoardsX
                 End If
                  If strTmpBoartStr <> "" Then
                    intNumbet = intNumbet + 1
                  End If
                  If Left(UCase(MyStr), 6) = "BOARD " Then
                     FindBoardOk = False
                     bStartBoards = True
                  End If
                  
'               strTmpBoardsXY = Split(MyStr, " ")
'               If strTmpBoardsXY(o) <> "" Then
'                   strBoardsNumber (0)
'               End If
'               For o = 1 To UBound(strTmpBoardsXY)
'                    If strTmpBoardsXY(o) <> "" Then
'                       strBoardsNumber (o)
'                    End If
'
'               Next
            End If
            
             If Left(UCase(MyStr), 6) = "BOARD " And bStartBoards = True Then
                strTmpBoardName = Replace(UCase(MyStr), "BOARD ", "")
                strTmpBoardName = Replace(strTmpBoardName, """", "")
                strTmpBoardName = Trim(strTmpBoardName)
                For yyYy = 0 To intNumbet
                  If UCase(strTmpBoardName) = UCase(strBoardsName(yyYy)) Then
                      intCurrNumber = yyYy
                      Exit For
                  End If
                Next
             End If
            
            
            If UCase(Left(MyStr, 5)) = "NODE " Then
                NodeName = Trim(Right(MyStr, Len(MyStr) - 4))
                
                NodeNumber = "#" & i
                'no probe
                If InStr(UCase(NodeName), "NO_ACCESS;") <> 0 Then
                   tmpNode1 = Split(NodeName, " ")
                   NodeName = Trim(tmpNode1(0))
                   NoProbeTrue = True
                   
''great 20080521

                   
                   If NoProbeTrue = True Then

                      Print #22, "$0,0,0,NO_PROBE,E1,(B),#0," & NodeName

                   End If
                   
                  Else
                   FindNodeXYTrue = True
                   
                End If
                
            End If
            If FindNodeXYTrue = True And UCase(Left(MyStr, 5)) <> "NODE " Then
               If Right(UCase(MyStr), 10) = "MANDATORY;" <> 0 Or InStr(UCase(MyStr), "TOP;") <> 0 Or Right(UCase(MyStr), 10) = "PREFERRED;" <> 0 Then
                  NodeXY = Split(MyStr, ",")
                  strBoardsX(intCurrNumber) = Val(strBoardsX(intCurrNumber))
                  NodeX = Trim(NodeXY(0))
                  NodeX = (Val(NodeX))
                  
                  
                   If intCurrNumber <> 0 Then
                       NodeX = NodeX + strBoardsX(intCurrNumber) ' + 100000
                       Else
                        NodeX = NodeX + strBoardsX(intCurrNumber)
                   End If
                   NodeX = NodeX / 12000
                  'NodeX = Left(Trim(NodeX), Len(NodeX) - 3) & "." & Right(Trim(NodeX), 3)
               
               
               If Left(NodeX, 1) = "." Then
                  NodeX = "0" & NodeX
               End If
                  
                  '(T) and (B)--
                  If InStr(UCase(NodeXY(1)), "TOP") <> 0 Then
                     bTB = "(T)"
                    Else
                     bTB = "(B)"
                  End If
                  '-------
                  If InStr(UCase(NodeXY(1)), " ") <> 0 Then
                     tmpNode2 = Split(Trim(NodeXY(1)), " ")
                     NodeY = Trim(tmpNode2(0))
                     strBoardsY(intCurrNumber) = Val(strBoardsY(intCurrNumber))
                      NodeY = (Val(NodeY))
                      
                      If intCurrNumber <> 0 Then
                           NodeY = NodeY + strBoardsY(intCurrNumber) '+ 100000
                           Else
                            NodeY = NodeY + strBoardsY(intCurrNumber)
                      End If
                     NodeY = NodeY / 12000
                    ' NodeY = Left(Trim(tmpNode2(0)), Len(tmpNode2(0)) - 3) & "." & Right(Trim(tmpNode2(0)), 3)
               
               
                         If Left(NodeY, 1) = "." Then
                             NodeY = "0" & NodeY
                         End If
                    
                    Else
                     NodeY = Replace(Trim(NodeXY(1)), ";", "")
                     
                     strBoardsY(intCurrNumber) = Val(strBoardsY(intCurrNumber))
                      NodeY = (Val(NodeY))
                      
                     If intCurrNumber <> 0 Then
                         NodeY = NodeY + strBoardsY(intCurrNumber) '+ 100000
                         Else
                         NodeY = NodeY + strBoardsY(intCurrNumber)
                     End If
                      NodeY = NodeY / 12000
                  '  NodeY = Left(Trim(NodeY), Len(NodeY) - 3) & "." & Right(Trim(NodeY), 3)
                         If Left(NodeY, 1) = "." Then
                             NodeY = "0" & NodeY
                         End If
                    
                  End If
                  If NodeX <> "" And NodeY <> "" And bTB <> "" Then
                    Print #22, "$0," & NodeX & "," & NodeY & ",5,E1," & bTB & "," & NodeNumber & "," & NodeName
                    'Print #2, "$0," & NodeX & ","; NodeY & ",5,E1," & bTB & "," & NodeNumber & "," & NodeName
                     'tmpNode3 = NodeName
                     'NextNode = True
                    'FindNodeXYTrue = False
                    

                        

                    
                    
                  End If
               End If
            End If
            If UCase(MyStr) = "OTHER" Then
              FindNodeXYTrue = False
             ' Print #2, "#Pin#"
             ' Print #2, "Part,TB,Pin,Name,X,Y,Layer,Netname"
               FindDvsTrue = True
            End If
            If FindDvsTrue = True And UCase(MyStr) <> "OTHER" And UCase(MyStr) <> "ALTERNATES" And UCase(MyStr) <> "DEVICES" Then
                tmpDvs1 = Split(MyStr, ",")
               ' DvsX = Left(Trim(tmpDvs1(0)), Len(tmpDvs1(0)) - 3) & "." & Right(Trim(tmpDvs1(0)), 3)
                         If Left(DvsX, 1) = "." Then
                             DvsX = "0" & DvsX
                         End If
                
                 strBoardsX(intCurrNumber) = Val(strBoardsX(intCurrNumber))
                
                
                 'DvsX = (Val(Trim(tmpDvs1(0)))) '/ 12000
                 DvsX = Val(Trim(tmpDvs1(0)))
                 
                 
                If intCurrNumber <> 0 Then
                    DvsX = DvsX + strBoardsX(intCurrNumber) '+ 100000
                    Else
                     DvsX = DvsX + strBoardsX(intCurrNumber)
                End If
                 DvsX = DvsX / 12000
                If InStr(UCase(tmpDvs1(1)), "TOP") <> 0 Then
                   bTB = "(T)"
                  Else
                   bTB = "(B)"
                End If
                tmpDvs2 = Split(Trim(tmpDvs1(1)), " ")
                
                'DvsY = (Val(Trim(tmpDvs2(0)))) / 12000
                strBoardsY(intCurrNumber) = Val(strBoardsY(intCurrNumber))
                DvsY = Val(Trim(tmpDvs2(0)))
                
                
                If intCurrNumber <> 0 Then
                   DvsY = DvsY + strBoardsY(intCurrNumber) '+ 100000
                   Else
                   DvsY = DvsY + strBoardsY(intCurrNumber)
                End If
                
                
                DvsY = DvsY / 12000
                
               'DvsY = Left(Trim(tmpDvs2(0)), Len(tmpDvs2(0)) - 3) & "." & Right(Trim(tmpDvs2(0)), 3)
                         If Left(DvsY, 1) = "." Then
                             DvsY = "0" & DvsY
                         End If
                
                
                tmpDvs3 = Split(Trim(tmpDvs2(1)), ".")
                DvsName = Trim(tmpDvs3(0))
                
                If yy = 0 Then
                   SameDvs = DvsName
                  yy = 1
                  ee = 0
                End If
                If SameDvs <> DvsName Then
                   yy = 0
                End If
                DvsPin = Trim(tmpDvs3(1))
                tt = Val(DvsPin)
                If tt = 0 Then
                   ee = ee + 1
                   Else
                   ee = Val(DvsPin)
                End If
                
'readboard floder   file
               If Option1 = True Then
                    If Dir(ToolPath & "FixToBoardView\ReadBoard\" & intCurrNumber + 1 & "%" & DvsName & ".tmp") <> "" Then
                       Open ToolPath & "FixToBoardView\ReadBoard\" & intCurrNumber + 1 & "%" & DvsName & ".tmp" For Input As #6
                          Do Until EOF(6)
                             Line Input #6, MyDevice
                             MyDevice = Trim(MyDevice)
                             If MyDevice <> "" Then
                                If InStr(MyDevice, ".") <> 0 Then
                                   tmpStr1 = Split(MyDevice, ".")
                                   If UCase(tmpStr1(0)) = UCase(DvsPin) Then
                                      DeviceNode = Trim(tmpStr1(1))
                                      Exit Do
                                   End If
                                   
                                End If
                             End If
                          Loop
                       Close #6
                     End If
                   Else
                    If Option2 = True Then
                        If Dir(ToolPath & "FixToBoardView\ReadBoard\" & DvsName & "." & DvsPin & ".tmp") <> "" Then
                           Open ToolPath & "FixToBoardView\ReadBoard\" & DvsName & "." & DvsPin & ".tmp" For Input As #6
                                 Line Input #6, MyDevice
                                 MyDevice = Trim(MyDevice)
                                 If MyDevice <> "" Then
                                     DeviceNode = Trim(MyDevice)
    
                                 End If
                           Close #6
                        End If
                   End If
                End If
                
                
                
                 If DeviceNode = "" Then DeviceNode = "NotFindNode"
                Print #33, intCurrNumber + 1 & "%" & DvsName & "," & bTB & "," & ee & "," & DvsPin; "," & DvsX & "," & DvsY & ","; "2" & "," & DeviceNode
                DvsName = ""
                bTB = ""
                DvsX = ""
                DvsY = ""
                DeviceName = ""
                
            End If
            If UCase(MyStr) = "DEVICES" Then
               FindDvsTrue = False
               FindNodeXYTrue = False
            End If
            
            If Replace(UCase(MyStr), " ", "") = "ENDBOARD" Then
               FindDvsTrue = False
               FindNodeXYTrue = False
            End If
            
            
            If Trim(UCase(MyStr)) = "END" Then
                    FindDvsTrue = False
                    FindNodeXYTrue = False
                     NodeX = ""
                     NodeY = ""
                     NodeNumber = ""
                     NodeName = ""
                     
                     bTB = ""
            
            
               Exit Do
            End If
          Else
           FindNodeXYTrue = False
            NodeX = ""
            NodeY = ""
            NodeNumber = ""
            NodeName = ""
            
            bTB = ""
      End If
      DoEvents
      L.Caption = "Read board_xy file line: " & JJ
  Loop '#1
  
Close #1
Close #11
Close #22
Close #33

'all add
  Open ToolPath & "FixToBoardView\ReadBoard\Layout.txt" For Input As 11
   Do Until EOF(11) '#11
     MyStr = ""
     Line Input #11, MyStr
     MyStr = Trim(MyStr)
     If MyStr <> "" Then
        Print #2, MyStr
     End If
     DoEvents
   Loop
  Close #11
                   Print #2, "#Nail#"
                  Print #2, "Nail,X,Y,Type,Grid,TB,Net,NetName"
  
    Open ToolPath & "FixToBoardView\ReadBoard\Nail.txt" For Input As 22
    
    Do Until EOF(22) '#22
     MyStr = ""
     Line Input #22, MyStr
     MyStr = Trim(MyStr)
     If MyStr <> "" Then
        Print #2, MyStr
     End If
     DoEvents
   Loop
  Close #22
    
                Print #2, "#Pin#"
              Print #2, "Part,TB,Pin,Name,X,Y,Layer,Netname"
    
    
  Open ToolPath & "FixToBoardView\ReadBoard\Pin.txt" For Input As 33

    Do Until EOF(33) '#33
     MyStr = ""
     Line Input #33, MyStr
     MyStr = Trim(MyStr)
     If MyStr <> "" Then
        Print #2, MyStr
     End If
     DoEvents
   Loop
  Close #33




Close #2
 Kill ToolPath & "FixToBoardView\ReadBoard\Layout.txt"
 Kill ToolPath & "FixToBoardView\ReadBoard\Nail.txt"
 Kill ToolPath & "FixToBoardView\ReadBoard\Pin.txt"
Kill ToolPath & "FixToBoardView\ReadBoard\*.tmp"
L.Caption = "Save BoardView.bv2 file ok"
MsgBox "NoBrcBoardView.bv2 creat OK !", vbQuestion
 Exit Sub
EX:
 MsgBox Err.Description, vbCritical
End Sub





