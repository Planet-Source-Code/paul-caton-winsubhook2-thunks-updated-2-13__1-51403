VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HookMouse"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3840
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================================
'HookMouse - a low-level system-wide (global) mouse hook demonstration.
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'==================================================================================================
Option Explicit

Private Const SWP_NOMOVE      As Long = &H2
Private Const SWP_NOSIZE      As Long = &H1
Private Const HWND_TOPMOST    As Long = -1
Private Const HWND_NOTOPMOST  As Long = -2

Private bPSAPI  As Boolean
Private hk      As cHook

Implements iHook

Private Declare Function GetWindowTextA Lib "user32" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
  Dim hModule As Long
  
  'Create the hook instance
  Set hk = New cHook
  
  'Hook the mouse system wide
  Call hk.Hook(Me, WH_MOUSE_LL, False)
  
  With Me
    'Check that the PSAPI DLL exists... if so we can (easily) determine the exe from the hWnd
    hModule = LoadLibraryA("PSAPI.DLL")
    If hModule Then
      Call FreeLibrary(hModule)
      bPSAPI = True
      .Height = .Height + .TextHeight("My")
    End If
    
    Call .Show
    DoEvents
    
    'Make this window stay on top
    Call SetWindowPos(.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Destroy the hook instance
  Set hk = Nothing
End Sub

Private Sub iHook_Proc(ByVal bBefore As Boolean, _
                       ByRef bHandled As Boolean, _
                       ByRef lReturn As Long, _
                       ByRef nCode As WinSubHook2.eHookCode, _
                       ByRef wParam As Long, _
                       ByRef lParam As Long)
  Dim hWndHk As Long              'Handle of the window under the mouse pointer
  Dim nLen   As Long              'Length returned by GetWindowText
  Dim sText  As String            'Store the window text (caption) here
  Dim dat    As tMSLLHOOKSTRUCT   'Low-level mouse data
  
  If nCode = HC_ACTION Then
    If Not bBefore Then
      'After
      
      'lParam points to the low-level mouse data, copy it to dat
      dat = hk.xMSLLHOOKSTRUCT(lParam)
        
      With dat.pt
        Me.Cls
        Me.Print "X" & vbTab & " " & Format$(.x, "#,###")
        Me.Print "Y" & vbTab & " " & Format$(.y, "#,###")
        
        'Get the window under the mouse pointer
        hWndHk = WindowFromPoint(.x, .y)
        
        Me.Print "hWnd" & vbTab & " " & hWndHk
        
        If hWndHk <> 0 Then
          'We have a valid window handle
    
          'Initialize a string buffer
          sText = Space$(255)
          
          'Get the window text (caption)
          nLen = GetWindowTextA(hWndHk, sText, 255)
          
          Me.Print "Caption:";
          
          'If the window has a caption
          If nLen > 0 Then
            Me.Print " " & Left$(sText, nLen); ""
          Else
            Me.Print vbNullString
          End If
          
          If bPSAPI Then
            Me.Print "EXE" & vbTab & " " & mhWndToExe.ExeName(hWndHk)
          End If
        End If
      End With
    End If
  End If
End Sub
