Attribute VB_Name = "TryIcon"
Public Const WM_MOUSEISMOVING = &H200 ' Mouse is moving
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_SETHOTKEY = &H32

' The API Call
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As enm_NIM_Shell, pnid As NOTIFYICONDATA) As Boolean

' User defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

' This is an Enum that tells the API what to do...
' Constants required by Shell_NotifyIcon API call:
Public Enum enm_NIM_Shell
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
    NIF_MESSAGE = &H1
    NIF_ICON = &H2
    NIF_TIP = &H4
    WM_MOUSEMOVE = &H200
End Enum

Public nidProgramData As NOTIFYICONDATA

