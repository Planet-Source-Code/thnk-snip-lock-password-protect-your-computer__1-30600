Attribute VB_Name = "Module1"
'Snip Lock v. 1.0.1

'If you have any question email me at thnk@aol.com


Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SPI_SCREENSAVERRUNNING = 97

Public Const WM_COMMAND = &H111

Public Const MIN_ALL = 419
Public Const MIN_ALL_UNDO = 416

Public Const LB_SETHORIZONTALEXTENT = &H194

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5


Public Sub WriteToINI(iSection As String, iKey As String, iKeyValue As String, iDirectory As String)
    Call WritePrivateProfileString(iSection$, UCase$(iKey$), iKeyValue$, iDirectory$)
End Sub

Public Function GetFromINI(iSection As String, iKey As String, iDirectory As String) As String
Dim strBuffer As String
    
    strBuffer = String(750, Chr(0))
    iKey$ = LCase$(iKey$)
    GetFromINI$ = Left(strBuffer, GetPrivateProfileString(iSection$, ByVal iKey$, "", strBuffer, Len(strBuffer), iDirectory$))
End Function

Public Sub Pause(TheTime As Long)
Dim Now As Long
    Now = Timer
    Do Until Timer - Now >= TheTime
        DoEvents
    Loop
End Sub

Public Function Desktop_IconsHide()
Dim Progman As Long, Shelldlldefview As Long, Syslistview As Long
    
    Progman = FindWindow("progman", vbNullString)
    Shelldlldefview = FindWindowEx(Progman, 0&, "shelldll_defview", vbNullString)
    Syslistview = FindWindowEx(Shelldlldefview, 0&, "syslistview32", vbNullString)

Call ShowWindow(Syslistview, SW_HIDE)
End Function

Public Function Desktop_IconsShow()
Dim Progman As Long, Shelldlldefview As Long, Syslistview As Long
    
    Progman = FindWindow("progman", vbNullString)
    Shelldlldefview = FindWindowEx(Progman, 0&, "shelldll_defview", vbNullString)
    Syslistview = FindWindowEx(Shelldlldefview, 0&, "syslistview32", vbNullString)

Call ShowWindow(Syslistview, SW_SHOW)
End Function

Public Function Disable_CtrlAltDel()
Dim ret As Integer

    ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, vbNullString, 0)
End Function

Public Function Enable_CtrlAltDel()
Dim ret As Integer

    ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, vbNullString, 0)
End Function

Public Function Taskbar_Show()
Dim Shelltraywnd&
    
    Shelltraywnd& = FindWindow("shell_traywnd", vbNullString)
Call ShowWindow(Shelltraywnd&, SW_SHOW)
End Function


Public Function Taskbar_Hide()
Dim Shelltraywnd&
    
    Shelltraywnd& = FindWindow("shell_traywnd", vbNullString)
Call ShowWindow(Shelltraywnd&, SW_HIDE)
End Function


Sub Minimize_AllWin(Optional Restore As Boolean)
    Dim hwnd As Long
    hwnd = FindWindow("Shell_TrayWnd", vbNullString)
    If Restore Then
        SendMessage hwnd, WM_COMMAND, MIN_ALL_UNDO, ByVal 0&
    Else
        SendMessage hwnd, WM_COMMAND, MIN_ALL, ByVal 0&
    End If
End Sub
