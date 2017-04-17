VERSION 5.00
Begin VB.Form frmView 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   8490
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   840
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BeginPath Lib "gdi32" _
                (ByVal hdc As Long) _
                As Long

Private Declare Function SetBkMode Lib "gdi32" _
                (ByVal hdc As Long, _
                ByVal nBkMode As Long) _
                As Long

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" _
                (ByVal hdc As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal lpString As String, _
                ByVal nCount As Long) _
                As Long

Private Declare Function EndPath Lib "gdi32" _
                (ByVal hdc As Long) _
                As Long

Private Declare Function PathToRegion Lib "gdi32" _
                (ByVal hdc As Long) _
                As Long

Private Declare Function SetWindowRgn Lib "user32" _
                (ByVal hwnd As Long, _
                ByVal hRgn As Long, _
                ByVal bRedraw As Boolean) _
                As Long

Private Declare Function SelectObject Lib "gdi32" _
                (ByVal hdc As Long, _
                ByVal hObject As Long) _
                As Long
                
Private Declare Function CreateFont Lib "gdi32" _
                Alias "CreateFontA" _
                (ByVal H As Long, _
                ByVal W As Long, _
                ByVal e As Long, _
                ByVal O As Long, _
                ByVal W As Long, _
                ByVal i As Long, _
                ByVal u As Long, _
                ByVal S As Long, _
                ByVal C As Long, _
                ByVal OP As Long, _
                ByVal CP As Long, _
                ByVal Q As Long, _
                ByVal PAF As Long, _
                ByVal f As String) _
                As Long
                
Private Const OPAQUE = 2
Private Const TRANSPARENT = 1

Private Const ANSI_CHARSET = 0
Private Const FW_HEAVY = 900
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.
Dim II As Integer

' 窗口置前=========
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'-------------------

Private Sub Form_Load()
Dim myval
      myval = SetWindowPos(frmView.hwnd, -1, 0, 0, 0, 0, 3)

   Timer1.Enabled = True
   II = 0
    Dim dc As Long
    Dim m_wndRgn As Long
    Dim m_Font As Long
    Dim m_OldFont As Long
    If ViewText = "" Then Me.Hide
    LenViewText = Len(ViewText)
    dc = Me.hdc
    m_Font = CreateFont(120, 50, 0, 0, FW_HEAVY, 1, 0, _
                       0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, _
                       CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
                       DEFAULT_PITCH Or FF_SWISS, "宋体")
    BeginPath dc
    '开始记录窗体轮廓路径
    SetBkMode dc, TRANSPARENT
    '设置背景为透明模式,这是必须有的
    m_OldFont = SelectObject(dc, m_Font)
    TextOut dc, 0, 0, ViewText, LenViewText
    SelectObject dc, m_OldFont
    EndPath dc
    '结束记录窗体轮廓路径
    m_wndRgn = PathToRegion(dc)
    '把所记录的路径转化为窗体轮廓句柄
    SetWindowRgn Me.hwnd, m_wndRgn, True
    '赋予窗体指定的轮廓形状
    
End Sub

Private Sub Timer1_Timer()
II = II + 1
If II >= 2 Then
  ViewText = ""
  Timer1.Enabled = False
  Unload Me
End If
End Sub
