VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'This class demonstrates Api Windows via the cWindow class
'
'I specifically wanted to show that...
'  cWindow, and all the WinSubHook2 Thunks can callback to a class
'  The WinSubHook2 thunks visibly survive ide stop. Follow the instructions in the MsgBox.
'
'For a *real* world scenario see the Shadow control.
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'==================================================================================================
Option Explicit

Private hWnd1      As Long
Private hWnd2      As Long
Private Window     As cWindow    'Window class instance

Implements WinSubHook2.iWindow  'We're implementing the interface

Private Declare Function TextOut Lib "gdi32" Alias "TextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long

Public Sub Display()
  Dim sMsg As String

  Set Window = New cWindow      'Create the window class instance
  With Window
    'Register the window class
    If Not .WindowClassRegister("ApiWindowClass") Then
      'This will happen *if* somehow the window class didn't get unregistered at the end of the last run.
      'Not to worry too much though, we can reuse the window class.
      Call MsgBox("Window class registration failed.")
    End If
    
    'Set the owner
    Set .Owner = Me
    
    'Add the messages that we're interested in
    Call .AddMsg(WM_CLOSE)
    Call .AddMsg(WM_NCHITTEST)
    Call .AddMsg(WM_PAINT)
    
    'Create and show a window
    hWnd1 = .WindowCreate(WS_EX_TOPMOST, WS_POPUPWINDOW Or WS_CAPTION Or WS_SYSMENU, , 100, 50, 150, 150, "API Window 1")
    Call WinSubHook2.ShowWindow(hWnd1, WinSubHook2.SW_SHOWNOACTIVATE)
    
    'Create and show another window
    hWnd2 = .WindowCreate(WS_EX_TOPMOST Or WS_EX_CLIENTEDGE, WS_OVERLAPPEDWINDOW, , 100, 220, 300, 100, "API Window 2")
    Call WinSubHook2.ShowWindow(hWnd2, WinSubHook2.SW_SHOWNOACTIVATE)
  End With 'WINDOW
  
  If Window.InIDE Then
    'There's no point in suggesting the following if we're compiled.
    sMsg = "Just for fun, try clicking the Yes button to execute an End statement.." & vbNewLine & vbNewLine & "If you then move the mouse pointer over an api window, you should see it get destroyed." & vbNewLine & _
     "If the window *appears* to remain visible then it's probably the case that you" & vbNewLine & "don't have any code windows open. The VB IDE MDI background doesn't repaint itself." & vbNewLine & _
     "Even though the app/IDE has stopped, the cWindow asm thunk is still alive, it" & vbNewLine & "receives the WM_MOUSEMOVE message, detects that the ide has stopped and destroys" & vbNewLine & _
     "the window/unregisters the window class.. all by itself." & vbNewLine & vbNewLine & "The same technique is used in all the WinSubHook2 thunks for ide protection." & vbNewLine & _
     "Here, by clicking Yes, the process becomes observable." & vbNewLine & vbNewLine & "Do you want to execute an End statement to stop the app/IDE?"

    If MsgBox(sMsg, vbQuestion + vbYesNo) = vbYes Then
      End
    End If
  End If
  
  'Loop until both windows are closed
  Do While hWnd1 Or hWnd2
    DoEvents
  Loop
  
  Set Window = Nothing
End Sub

Private Sub iWindow_Proc(ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef hWnd As Long, ByRef uMsg As WinSubHook2.eMsg, ByRef wParam As Long, ByRef lParam As Long)
Const HTCAPTION As Long = 2
Const HTCLIENT  As Long = 1
  Dim hDC       As Long
  Dim sMsg      As String
  
  Select Case uMsg
  Case WM_CLOSE
    'Destroy the window
    If hWnd = hWnd1 Then
      hWnd1 = 0
    ElseIf hWnd = hWnd2 Then
      hWnd2 = 0
    End If
    Call Window.WindowDestroy(hWnd)
    
  Case WM_NCHITTEST
    If hWnd = hWnd2 Then
      'Call the default window proc first, if it says that the mouse is over the client area...
      If Window.CallDefWndProc(hWnd, uMsg, wParam, lParam) = HTCLIENT Then
        lReturn = HTCAPTION 'Lie that the mouse is over the caption. You can now click and drag on the client area
        bHandled = True     'Don't call DefWndProc hereafter
      End If
    End If
  
  Case WM_PAINT
    If hWnd = hWnd2 Then
      sMsg = "Try clicking and dragging the client area"
      hDC = WinSubHook2.GetDC(hWnd)
      Call TextOut(hDC, 0, 0, StrPtr(sMsg), Len(sMsg))
      lReturn = 0
    End If
  End Select
End Sub
