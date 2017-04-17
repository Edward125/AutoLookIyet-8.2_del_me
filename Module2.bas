Attribute VB_Name = "Module2"
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long ' Note that if you declare the lpData parameter as String, you must pass it By Value.
'download by http://www.codefans.net
Declare Function EnumWindows Lib "user32" (ByVal wndenmprc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10

Private Target As String

Public Const REG_SZ = 1
Public Const HKEY_CURRENT_USER = &H80000001
' Check a returned task to see if we should
' kill it.
Public Function EnumCallback(ByVal app_hWnd As Long, ByVal param As Long) As Long
Dim buf As String * 256
Dim title As String
Dim length As Long

    ' Get the window's title.
    length = GetWindowText(app_hWnd, buf, Len(buf))
    title = left$(buf, length)

    ' See if this is the target window.
    If InStr(title, Target) <> 0 Then
        ' Kill the window.
        SendMessage app_hWnd, WM_CLOSE, 0, 0
    End If
    
    ' Continue searching.
    EnumCallback = 1
End Function

' Ask Windows for the list of tasks.
Public Sub TerminateTask(app_name As String)
    Target = app_name
    EnumWindows AddressOf EnumCallback, 0
End Sub

