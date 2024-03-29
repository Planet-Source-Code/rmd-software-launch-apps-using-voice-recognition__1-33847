Attribute VB_Name = "Module1"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
    End Type
    'constants required by Shell_NotifyIcon
    '     API call:
    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    Public Const WM_MOUSEMOVE = &H200
    Public Const WM_LBUTTONDOWN = &H201 'Button down
    Public Const WM_LBUTTONUP = &H202 'Button up
    Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
    Public Const WM_RBUTTONDOWN = &H204 'Button down
    Public Const WM_RBUTTONUP = &H205 'Button up
    Public Const WM_RBUTTONDBLCLK = &H206 'Double-click

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Type POINTAPI
    x As Long
    y As Long
End Type
Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    Public nid As NOTIFYICONDATA

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

