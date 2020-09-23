Attribute VB_Name = "DrawButton"
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Declare Function GetTitleBarInfo Lib "user32" (ByVal hwnd As Long, pti As TitleBarInfo) As Boolean
Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type TitleBarInfo
    cbSize As Long
    rcTitleBar As RECT  'A RECT structure that receives the coordinates of the title bar
    rgState(5) As Long  'An array that receives a DWORD value for each element of the title bar
End Type
            'rgState array Values
            '0  The titlebar Itself
            '1  Reserved
            '2  Min button
            '3  Max button
            '4  Help button
            '5  Close button
            '
            'rgstate return constatnts
            'STATE_SYSTEM_FOCUSABLE = &H00100000
            'STATE_SYSTEM_INVISIBLE = &H00008000
            'STATE_SYSTEM_OFFSCREEN = &H00010000
            'STATE_SYSTEM_PRESSED = &H00000008
            'STATE_SYSTEM_UNAVAILABLE = &H00000001
            
Const DFC_BUTTON = 4
Const DFCS_BUTTONPUSH = &H10
Const DFCS_PUSHED = &H200

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
   x As Long
   y As Long
End Type

Const SM_CXFRAME = 32
 
Const COLOR_BTNTEXT = 18

Dim lDC As Long
Public R As RECT

Public Sub ButtonDraw(frm As Form, bState As Boolean)

    Dim TBButtons As Integer
    Dim TBarHeight As Integer
    Dim TBButtonHeight As Integer
    Dim TBButtonWidth As Integer
    Dim DrawWidth As Integer
    Dim TBI As TitleBarInfo
    Dim TBIRect As RECT
    Dim bRslt As Boolean
    Dim WinBorder As Integer
    
    With frm
        If .BorderStyle = 0 Then Exit Sub ' Don't draw a button if there is no titlebar
        
        '----How Many Buttons in TitleBar------------------------------------------
        If Not .ControlBox Then TBButtons = 0
        If .ControlBox Then TBButtons = 1
        If .ControlBox And .WhatsThisButton Then
            If .BorderStyle < 4 Then
                TBButtons = 2
            Else
                tButtons = 1
            End If
        End If
        If .ControlBox And .MinButton And .BorderStyle = 2 Then TBButtons = 3
        If .ControlBox And .MinButton And .BorderStyle = 5 Then TBButtons = 1
        If .ControlBox And .MaxButton And .BorderStyle = 2 Then TBButtons = 3
        If .ControlBox And .MaxButton And .BorderStyle = 5 Then TBButtons = 1
        '------------------------------------------------------------------------
        
        '----Get height of Titlebar----------------------------------------------
        'Using this method gets the height of the titlebar regardless of the window
        'style. It does, however, restrict its use to Win98/2000. So if you want to
        'use this code in Win95, then call GetSystemMetrics to find the windowstyle
        'and titlebar size.
        TBI.cbSize = Len(TBI)
        bRslt = GetTitleBarInfo(.hwnd, TBI)
        TBIRect = TBI.rcTitleBar
        TBarHeight = TBIRect.Bottom - TBIRect.Top - 1
        '-----------------------------------------------------------------------
        
        '----Get WindowBorder Size----------------------------------------------
        If .BorderStyle = 2 Or .BorderStyle = 5 Then
            R.Top = GetSystemMetrics(32) + 2
            WinBorder = R.Top - 6
        Else
            R.Top = 5
            WinBorder = -1
        End If
    End With
    '---------------------------------------------------------------------------
    
    '----Use Titlebar Height to determin button size----------------------------
    TBButtonHeight = TBarHeight - 4
    TBButtonWidth = TBButtonHeight + 2
    'and the size and space of the dot on the button
    DrawWidth = TBarHeight / 8
    '---------------------------------------------------------------------------
    
    '----Determin the position of our button------------------------------------
    R.Bottom = R.Top + TBButtonHeight
    
    Select Case TBButtons
        Case 1
            R.Right = frm.ScaleWidth - (TBButtonWidth) + WinBorder
        Case 2
            R.Right = frm.ScaleWidth - ((TBButtonWidth * 2) + 2) + WinBorder
        Case 3
            R.Right = frm.ScaleWidth - ((TBButtonWidth * 3) + 2) + WinBorder
        Case Else
            R.Right = frm.ScaleWidth
    End Select
    
    R.Left = R.Right - TBButtonWidth
    '--------------------------------------------------------------------------
    
    '----Get the Widow DC so that we may draw in the title bar-----------------
    lDC = GetWindowDC(frm.hwnd)
    '--------------------------------------------------------------------------
    
    '----Determin the position of the dot--------------------------------------
    Dim StartXY As Integer, EndXY As Integer

    Select Case TBarHeight
        Case Is < 20
            StartXY = DrawWidth + 1
            EndXY = DrawWidth - 1
        Case Else
            StartXY = (DrawWidth * 2)
            EndXY = DrawWidth
    End Select
    '--------------------------------------------------------------------------
    
    '----We have all the information we need So Draw the button----------------
    Dim rDot As RECT
    If bState Then
        DrawFrameControl lDC, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED
        rDot.Left = R.Right - (1 + StartXY): rDot.Top = R.Bottom - (1 + StartXY)
        rDot.Right = R.Right - (1 + EndXY): rDot.Bottom = R.Bottom - (1 + EndXY)
    Else
        DrawFrameControl lDC, R, DFC_BUTTON, DFCS_BUTTONPUSH
        rDot.Left = R.Right - (2 + StartXY): rDot.Top = R.Bottom - (2 + StartXY)
        rDot.Right = R.Right - (2 + EndXY): rDot.Bottom = R.Bottom - (2 + EndXY)
    End If
    
    FillRect lDC, rDot, GetSysColorBrush(COLOR_BTNTEXT)
    '---------------------------------------------------------------------------
    
    '----Set Tooltip------------------------------------------------------------
    If TTReset Then
        Dim TTRect As RECT
        
        TTRect.Bottom = R.Bottom + (TBarHeight - ((TBarHeight * 2) + WinBorder + 5))
        TTRect.Left = R.Left - (4 - WinBorder)
        TTRect.Right = R.Right - (4 - WinBorder)
        TTRect.Top = R.Top + (TBarHeight - ((TBarHeight * 2) + WinBorder + 5))
        
        If hWndTT <> 0 Then KillTip 'ToolTip KillTip()
        CreateTip appForm.hwnd, "System Tray", TTRect 'ToolTip CreateTip()
        TTReset = False
    End If
    
End Sub
