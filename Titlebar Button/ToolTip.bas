Attribute VB_Name = "ToolTip"
Const WS_EX_TOPMOST = &H8&
Const TTS_ALWAYSTIP = &H1
Const HWND_TOPMOST = -1

Const SWP_NOACTIVATE = &H10
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Const WM_USER = &H400
Const TTM_ADDTOOLA = (WM_USER + 4)
Const TTM_SETDELAYTIME = (WM_USER + 3)
Const TTF_SUBCLASS = &H10
Const TTDT_INITIAL = 3

Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uid As Long
    RECT As RECT
    hinst As Long
    lpszText As String
    lParam As Long
End Type

Public hWndTT As Long
Public TTReset As Boolean

Public Sub CreateTip(hwndForm As Long, szText As String, rct As RECT)
    
    hWndTT = CreateWindowEx(WS_EX_TOPMOST, "tooltips_class32", "", TTS_ALWAYSTIP, _
                            0, 0, 0, 0, hwndForm, 0&, App.hInstance, 0&)

    SetWindowPos hWndTT, HWND_TOPMOST, 0, 0, 0, 0, _
                        SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE

    Dim TI As TOOLINFO
    
    With TI
        .cbSize = Len(TI)
        .uFlags = TTF_SUBCLASS
        .hwnd = hwndForm
        .hinst = App.hInstance
        .uid = 1&
        .lpszText = szText & vbNullChar
        .RECT = rct
    End With
    
    SendMessage hWndTT, TTM_ADDTOOLA, 0, TI
    SendMessage hWndTT, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal 10 'milliseconds
    
End Sub

Public Sub KillTip()
    DestroyWindow hWndTT
End Sub
