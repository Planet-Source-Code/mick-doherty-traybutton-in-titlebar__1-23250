Attribute VB_Name = "TrayNotify"
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uid As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        sztip As String * 64
End Type


Const NIM_ADD = &H0
Const NIM_DELETE = &H2
Const NIM_MODIFY = &H1
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4

Const MF_GRAYED = &H1&
Const MF_STRING = &H0&
Const MF_SEPARATOR = &H800&
Const TPM_NONOTIFY = &H80&
Const TPM_RETURNCMD = &H100&

Public bTraySet As Boolean
Public rctMenu As RECT, hMenu As Long, tMenu As Long

Public Sub TraySet(frm As Form, sztip As String, hIcon As Long)
    
    Dim NID As NOTIFYICONDATA
    
    With NID
        .cbSize = Len(NID)
        .hIcon = hIcon
        .hwnd = frm.hwnd
        .sztip = sztip & vbNullChar
        .uCallbackMessage = WM_LBUTTONUP
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uid = 1&
    End With
    
    Shell_NotifyIcon NIM_ADD, NID
    
    frm.Hide
    bTraySet = True
    
End Sub

Public Sub TrayRestore(frm As Form)
    
    Dim NID As NOTIFYICONDATA
    
    With NID
        .cbSize = Len(NID)
        .hwnd = frm.hwnd
        .uid = 1&
    End With
    
    Shell_NotifyIcon NIM_DELETE, NID
    frm.Show
    bTraySet = False
    
End Sub

Public Sub TrayMenu(frm As Form)
    
    GetCursorPos MP
    hMenu = CreatePopupMenu()
    If bTraySet Then
        AppendMenu hMenu, MF_STRING, 1000, "Restore"
    Else
        AppendMenu hMenu, MF_STRING Or MF_GRAYED, 1000, "Restore"
    End If
    AppendMenu hMenu, MF_SEPARATOR, 0&, 0&
    AppendMenu hMenu, MF_STRING, 1010, "Exit"
    
    tMenu = TrackPopupMenu(hMenu, TPM_NONOTIFY Or TPM_RETURNCMD, MP.x, MP.y, 0&, frm.hwnd, 0&)
    
    Select Case tMenu
        Case 1000
            TrayRestore frm
        Case 1010
            TrayRestore frm
            UnHook
            Unload frm
        Case Else
            'do nothing
    End Select
    
    DestroyMenu hMenu
    
End Sub
