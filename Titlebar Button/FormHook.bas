Attribute VB_Name = "FormHook"
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long, _
        ByVal hwnd As Long, _
        ByVal Msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Const GWL_WNDPROC = -4

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200

Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5

Public Const WM_ACTIVATE = &H6
Public Const WM_NCPAINT = &H85
Public Const WM_PAINT = &HF
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_COMMAND = &H111
Public Const WM_NCACTIVATE = &H86

Public Const WM_DESTROY = &H2
Public Const WM_SIZE = &H5

Global lpPrevWndProc As Long
Global gHW As Long
Global appForm As Form
Public MP As POINTAPI

Private Function MakePoints(lParam As Long) As POINTAPI
    Dim hexstr As String
    hexstr = Right("00000000" & Hex(lParam), 8)

    MakePoints.x = CLng("&H" & Right(hexstr, 4)) - (appForm.Left / Screen.TwipsPerPixelX)
    MakePoints.y = CLng("&H" & Left(hexstr, 4)) - (appForm.Top / Screen.TwipsPerPixelY)
End Function

Public Sub Hook(frm As Form)
    gHW = frm.hwnd
    Set appForm = frm
    frm.ScaleMode = vbPixels 'API works in pixels.
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
    TrayRestore frm
End Sub

Public Sub UnHook()
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hwnd As Long, _
            ByVal uMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long

    '------------------------------------------------------------------------------
    'Messing around in here can cause allsorts of problems.
    'So, if you must, make sure you save everytihing you want to keep
    'before you run the program.
    'Don't run anything outside of a message selection as it will be
    'executed so many times per second that it will slow down system response.
    Dim lRslt As Long
    Dim retProc As Boolean

    Static STButtonState As Boolean
    Static Toggle As Boolean
    
    On Error Resume Next
    
    Select Case uMsg
            
        Case WM_DESTROY
            TrayRestore appForm
            KillTip 'ToolTip KillTip()
            UnHook
            retProc = True
            
        Case WM_NCMOUSEMOVE
            'Only draw the button when necessary
            If GetAsyncKeyState(vbLeftButton) < 0 Then
                If OverButton(lParam) Then
                    If Toggle = False Then
                        Toggle = True
                        ButtonDraw appForm, Toggle 'DrawButton ButtonDraw()
                    End If
                Else
                    If Toggle = True Then
                        Toggle = False
                        ButtonDraw appForm, Toggle 'DrawButton ButtonDraw()
                    End If
                End If
            Else
                STButtonState = False
                retProc = True
            End If
        
        Case WM_NCLBUTTONDOWN
            If OverButton(lParam) Then
                STButtonState = True
                ButtonDraw appForm, True 'DrawButton ButtonDraw()
            Else
                STButtonState = False
                retProc = True
            End If
              
        Case WM_NCLBUTTONUP
            STButtonState = False
            If OverButton(lParam) Then
                TraySet appForm, appForm.Caption, appForm.Icon 'TrayNotify TraySet()
                ButtonDraw appForm, False 'DrawButton ButtonDraw()
                retProc = False
            Else
                retProc = True
            End If
           
        Case WM_LBUTTONUP
            STButtonState = False
            ButtonDraw appForm, False 'DrawButton ButtonDraw()
            If GetAsyncKeyState(vbLeftButton) < 0 And bTraySet Then
                TrayMenu appForm 'TrayNotify TrayMenu()
            End If
            retProc = True

        Case WM_NCLBUTTONDBLCLK, WM_NCRBUTTONDOWN
            If Not OverButton(lParam) Then
                retProc = True
            End If
            
        Case WM_NCPAINT, WM_PAINT, WM_COMMAND
            ButtonDraw appForm, False 'DrawButton ButtonDraw()
            retProc = True
            
        Case WM_SIZE, WM_ACTIVATEAPP, WM_NCACTIVATE, WM_ACTIVATE, WM_MOUSEACTIVATE
            TTReset = True
            ButtonDraw appForm, False 'DrawButton ButtonDraw()
            retProc = True
       
        Case Else
            retProc = True
            
    End Select
    
    
    If retProc Then
        WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
    Else
        WindowProc = 0
    End If
    
End Function

Private Function OverButton(lParam As Long) As Boolean
    
    MP = MakePoints(lParam)
        
    If PtInRect(R, MP.x, MP.y) Then OverButton = True

End Function
