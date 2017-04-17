Attribute VB_Name = "Top5"

Public unloadTop As Boolean
Dim strDeviceTestTimes(5)
Dim intTime As Integer
Dim strTimesPath As String

Public Sub Top_10(strDeviceName As String, strDeviceType As String)
   On Error Resume Next
'   Dim CurrentTime
'   strCurrentTime = Format(Now, "hhmm")
'   CurrentTime = Val(Right(strCurrentTime, 2))
'   MsgBox CurrentTime
Dim strTmpPath As String
Dim tmpTimes
Dim strSbusNode As String
Dim strIbusNode As String
Dim tmpInt(5) As Integer
Dim TestTimes As Integer
Dim TestTimes_All As Integer
Dim strFenPei() As String
Dim strFenPei_2() As String
Dim strTimesPath As String
Dim bSame As Boolean
strTimesPath = "C:\WINDOWS\system\Top10"
strTmpPath = strTimesPath
'strTmpPath = "C:\WINDOWS\system\To10"
   MkDir strTmpPath
   strTmpPath = strTimesPath & "\"
   Erase strDeviceTestTimes
   Erase tmpInt
   Erase strFenPei
'strTmpPath = "C:\WINDOWS\system\Top10\"
   If Dir(strTmpPath & strDeviceName & "," & strDeviceType & ".zuoai") <> "" Then
       Open strTmpPath & strDeviceName & "," & strDeviceType & ".zuoai" For Input As #7
          Line Input #7, tmpTimes
       Close #7
        TestTimes = Val(Trim(tmpTimes))
        TestTimes = TestTimes + 1
        tmpTimes = 0
        Open strTmpPath & strDeviceName & "," & strDeviceType & ".zuoai" For Output As #7
           Print #7, TestTimes
        Close #7
        'TestTimes = 0
     Else
        Open strTmpPath & strDeviceName & "," & strDeviceType & ".zuoai" For Output As #7
           Print #7, "1"
        Close #7
        TestTimes = 1
   End If
'analog s i bus
strSbusNode = ""
strIbusNode = ""
   If strDeviceType = "[Open]" Then
      strFenPei_2 = Split(strDeviceName, ",")
 
      strDeviceName = strFenPei_2(0)
      strSbusNode = strFenPei_2(1)
      
   End If
   
   If strDeviceType = "[Testjet]" Then
      strFenPei_2 = Split(strDeviceName, ",")
 
      strDeviceName = strFenPei_2(0)
      strSbusNode = strFenPei_2(1)
      strIbusNode = strFenPei_2(2)
   End If
   
   If strDeviceType = "[Analog]" And Dir(strBoardPath & "analog\" & strDeviceName) <> "" Then
 
     Open strBoardPath & "analog\" & strDeviceName For Input As #7
       Do Until EOF(7)
          Line Input #7, tmpAnalogStr
          tmpAnalogStr = Trim(Replace(tmpAnalogStr, " ", ""))
          If tmpAnalogStr <> "" And left(Trim(tmpAnalogStr), 1) <> "!" Then
              If left(tmpAnalogStr, 17) = "testpoweredanalog" Then
                 Exit Do
              End If
              If left(tmpAnalogStr, 10) = "connectsto" Then
                   strFenPei_2 = Split(tmpAnalogStr, """")
                   strSbusNode = strFenPei_2(1)
                   If left(strSbusNode, 2) = "#%" Then
                       strSbusNode = right(strSbusNode, Len(strSbusNode) - 2)
                   End If
              End If
              If left(tmpAnalogStr, 10) = "connectito" Then
                   strFenPei_2 = Split(tmpAnalogStr, """")
                   strIbusNode = strFenPei_2(1)
                   If left(strIbusNode, 2) = "#%" Then
                       strIbusNode = right(strIbusNode, Len(strIbusNode) - 2)
                   End If
              End If
          End If
          Erase strFenPei_2
          DoEvents
       Loop
     Close #7
   End If
   '------------------------------------------------------------------------------------------
'   If Dir(strTmpPath & strDeviceName & "." & strDeviceType & ".zuoaiBa") <> "" Then
'       Open strTmpPath & strDeviceName & "." & strDeviceType & ".zuoaiBa" For Input As #7
'          Line Input #7, tmpTimes
'       Close #7
'        TestTimes_All = Val(Trim(tmpTimes))
'        TestTimes_All = TestTimes + 1
'        Open strTmpPath & strDeviceName & "." & strDeviceType & ".zuoaiBa" For Output As #7
'           Print #7, TestTimes_All
'        Close #7
'        'TestTimes = 0
'     Else
'        Open strTmpPath & strDeviceName & "." & strDeviceType & ".zuoaiBa" For Output As #7
'           Print #7, "1"
'        Close #7
'        TestTimes_All = 1
'   End If
   
  '-------------------------------------------------------------------------------
   
   
   fu = 0
   fk = 0
   For uu = 0 To 4
       If strDeviceName & strSbusNode & strIbusNode & strDeviceType = frmTop_5.Text1(fu).Text & frmTop_5.Text1(fu + 1).Text & frmTop_5.Text1(fu + 2).Text & frmTop_5.Text1(fu + 3).Text Then
           frmTop_5.Text1(fu + 4).Text = TestTimes
           bSame = True
       End If
       fu = fu + 5
   Next
   
    If bSame = False Then
       strDeviceTestTimes(0) = strDeviceName & "," & strSbusNode & "," & strIbusNode & "," & strDeviceType & "," & TestTimes
       tmpInt(0) = TestTimes
       Else
        strDeviceTestTimes(0) = ",,,,"
        tmpInt(0) = 0
        bSame = False
   End If
   
   
   T = 0
   j = 0
   
    For i = 1 To 5
        T = T + 4
       tmpInt(i) = Val(frmTop_5.Text1(T).Text)
       T = T + 1
       strDeviceTestTimes(i) = frmTop_5.Text1(j).Text & "," & frmTop_5.Text1(j + 1).Text & "," & frmTop_5.Text1(j + 2).Text & "," & frmTop_5.Text1(j + 3).Text & "," & frmTop_5.Text1(j + 4).Text
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
       
      ' tmpInt(i) = Val(frmTop_5.Text1(T).Text)
        
       strFenPei = Split(strDeviceTestTimes(i), ",")
       S = 0
       For y = T To W
           frmTop_5.Text1(y).Text = strFenPei(S)
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
    
    
'   If frmTop_5.Text1(0).Text = "" Then
'     frmTop_5.Text1(0).Text = strDeviceName
'     frmTop_5.Text1(3).Text = strDeviceType
'     frmTop_5.Text1(4).Text = TestTimes
'   End If
'
'  If frmTop_5.Text1(5).Text = "" And strDeviceName <> frmTop_5.Text1(0).Text Then
'     frmTop_5.Text1(5).Text = strDeviceName
'     frmTop_5.Text1(8).Text = strDeviceType
'     frmTop_5.Text1(9).Text = TestTimes
'   End If
   
   TestTimes = 0
End Sub
Public Sub PaiXu(intI_1() As Integer) ', intCurrentDevice As Integer)

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

