VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   Caption         =   "HooKey"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5625
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
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin prjHooKey.ucShadow Shadow 
      Left            =   4845
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   2625
      Left            =   30
      TabIndex        =   0
      Top             =   615
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   4630
      _Version        =   393217
      BackColor       =   14737632
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4785
      Top             =   1245
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItm 
         Caption         =   "&Save..."
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileItm 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileItm 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditItm 
         Caption         =   "&Copy"
         Index           =   0
      End
      Begin VB.Menu mnuEditItm 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEditItm 
         Caption         =   "Select &All"
         Index           =   2
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptItm 
         Caption         =   "Active"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOptItm 
         Caption         =   "Show right-hand modifiers"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuOptItm 
         Caption         =   "Topmost"
         Checked         =   -1  'True
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================================
'HooKey - a low-level system-wide (global) keyboard hook sample.
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'==================================================================================================

Option Explicit

'Api constants
Private Const SWP_NOMOVE              As Long = &H2
Private Const SWP_NOSIZE              As Long = &H1
Private Const HWND_TOPMOST            As Long = -1
Private Const HWND_NOTOPMOST          As Long = -2
Private Const HSHELL_WINDOWACTIVATED  As Long = 4

'Display type
Private Enum eDispType
  Keyboard = 0
  Injected
  Activation
End Enum

'Private module variables
Private bWelcome              As Boolean                    'Whether the welcome text is on-screen
Private nLastLen              As Long                       'Last update text length
Private uMsgNotify            As Long                       'Shell hook notification message
Private WithEvents hk         As cHooKey                    'cHooKey class
Attribute hk.VB_VarHelpID = -1
Private sc                    As cSubclass                  'Subclasser for shell hook notifications

Implements WinSubHook2.iSubclass                            'We implement the iSubclass interface

'Api declares
Private Declare Function DeregisterShellHookWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowTextA Lib "user32" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function InitCommonControls Lib "comctl32" () As Long
Private Declare Function RegisterShellHookWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'For XP manifests
Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
  Dim os As WinSubHook2.tOSVERSIONINFO
  
  'Check the OS is Win2k, XP or better.
  os.dwOSVersionInfoSize = Len(os)
  Call WinSubHook2.GetVersionEx(os)
  
  If os.dwMajorVersion < 5 Then
  
    Call MsgBox("Sorry! You'll need Window 2000 or better to run " & App.Title, vbCritical)
    Call Unload(Me)
    Exit Sub
  End If
  
  'Check that we're the only instance running
  If App.PrevInstance Then
  
    Call MsgBox("You shouldn't be running more than one instance of " & App.Title & " at a time.", vbCritical)
    Call Unload(Me)
    Exit Sub
  End If
  
  With Me
    'Default to running topmost - see the Options menu.
    Call SetWindowPos(.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    .mnuOptItm(2).Checked = True
    
    'Load the welcome text
    On Error Resume Next
      With .rtf
        Call .LoadFile(App.Path & "\Welcome.rtf", rtfRTF)
        bWelcome = (Err.Number = 0)
        .SelStart = Len(.Text)
      End With
    On Error GoTo 0
    
    'Size the window based on a known client area width/height
    Call .Move(.Left, .Top, 7230# + (.Width - .ScaleWidth), 8205# + (.Height - .ScaleHeight))
    Call .Show
    DoEvents
    
    'Create a unique message number for the shell to notify us with
    uMsgNotify = RegisterWindowMessage(ByVal "SHELLHOOK")
    
    'We need a subclasser to catch the shell hook notification messages
    Set sc = New cSubclass
    With sc
      Call .AddMsg(uMsgNotify, MSG_AFTER)
      Call .Subclass(Me.hWnd, Me)
    End With
    
    'Register our window so as to receive shell notifications
    Call RegisterShellHookWindow(Me.hWnd)
    
    'Create the HooKey instance
    Set hk = New cHooKey
    hk.Active = True
    .mnuOptItm(0).Checked = hk.Active
    
    hk.ShowLR = True
    .mnuOptItm(1).Checked = hk.ShowLR
    
    .SetFocus
  End With
End Sub

Private Sub Form_Resize()
  With Me
    If .WindowState <> vbMinimized Then
      Call .rtf.Move(0, 0, .ScaleWidth, .ScaleHeight)
    End If
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call Shadow.FadeOut
  Call DeregisterShellHookWindow(Me.hWnd)                   'De-register our window from shell notification
  
  Set hk = Nothing                                          'Destroy the cHooKey instance
  Set sc = Nothing                                          'Destroy the cSubclass instance
End Sub

Private Sub mnuEdit_Click()
  With Me
    .mnuEditItm(0).Enabled = (.rtf.SelLength > 0)
  End With
End Sub

Private Sub mnuEditItm_Click(Index As Integer)
  With Me.rtf
    Select Case Index
    Case 0:
      Call Clipboard.SetText(.SelText)
    Case 2
      .SelStart = 0
      .SelLength = Len(.Text)
    End Select
  End With
End Sub

Private Sub mnuFileItm_Click(Index As Integer)
  Select Case Index
    Case 0: Call SaveAs
    Case 2: Call Unload(Me)
  End Select
End Sub

Private Sub mnuOptions_Click()
  With Me
    .mnuOptItm(0).Checked = hk.Active
    .mnuOptItm(1).Checked = hk.ShowLR
  End With
End Sub

Private Sub mnuOptItm_Click(Index As Integer)
  Dim zPos As Long
  
  With Me
    Select Case Index
    Case 0
      hk.Active = Not hk.Active
      .mnuOptItm(0).Checked = hk.Active
      
    Case 1
      hk.ShowLR = Not hk.ShowLR
      .mnuOptItm(1).Checked = hk.ShowLR
      
    Case 2:
      .mnuOptItm(2).Checked = Not .mnuOptItm(2).Checked
      If .mnuOptItm(2).Checked Then
        zPos = HWND_TOPMOST
      Else
        zPos = HWND_NOTOPMOST
      End If
      
      Call SetWindowPos(.hWnd, zPos, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End Select
  End With
End Sub

'Key-strokes are reported here...
Private Sub hk_KeyPress(ByVal sKey As String, ByVal bInjected As Boolean)
  If Injected Then
    Call Update(sKey, Injected)
  Else
    Call Update(sKey, Keyboard)
  End If
End Sub

'Notification messages from the shell hook land here
Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
  Dim nLen     As Long
  Dim sCaption As String
  Dim sModule  As String
  Dim sLine    As String
  
  'Debug.Print hWnd, uMsg, wParam, lParam
  
  If hk.Active Then
    If (wParam And &HFF&) = HSHELL_WINDOWACTIVATED Then
      sCaption = Space$(256)
      nLen = GetWindowTextA(lParam, sCaption, 256)
      
      If nLen Then
        sCaption = Left$(sCaption, nLen)
      Else
        sCaption = vbNullString
      End If
      
      sModule = mhWndToExe.ExeFileName(lParam)
      
      sLine = "Active:" & vbTab & sModule & vbNewLine & _
              "Caption:" & vbTab & Chr$(34) & sCaption & Chr$(34)
              
      If nLen > 0 Then
        Call Update(sLine, Activation)
      End If
    End If
  End If
End Sub

Private Sub SaveAs()
  On Error GoTo Catch
  
  With Me.dlg
    .CancelError = True
    .DialogTitle = "Save as"
    .Filter = "Rich text files (*.rtf)|*.rtf|Text files (*.txt)|*.txt"
    .flags = cdlOFNOverwritePrompt
    Call .ShowSave
    
    If StrComp(Right$(.FileName, 4), ".rtf") = 0 Then
      Call Me.rtf.SaveFile(.FileName, rtfRTF)
    ElseIf StrComp(Right$(.FileName, 4), ".txt") = 0 Then
      Call Me.rtf.SaveFile(.FileName, rtfText)
    Else
      If .FilterIndex = 1 Then
        Call Me.rtf.SaveFile(.FileName & ".rtf", rtfRTF)
      Else
        Call Me.rtf.SaveFile(.FileName & ".txt", rtfText)
      End If
    End If
  End With
Catch:
  On Error GoTo 0
End Sub

Private Sub Update(ByVal sData As String, DispType As eDispType)
  Dim nLen As Long
  
  nLen = Len(sData)
  
  With Me.rtf
    If bWelcome Then
      bWelcome = False
      .TextRTF = vbNullString
      .SelBold = .Font.Bold
      .SelFontName = .Font.Name
      .SelFontSize = .Font.Size
      .SelItalic = .Font.Italic
      .SelUnderline = .Font.Underline
    End If
  
    Select Case DispType
      Case eDispType.Keyboard:    .SelColor = RGB(160, 0, 0)
      Case eDispType.Injected:    .SelColor = RGB(0, 0, 160)
      Case eDispType.Activation:  .SelColor = RGB(0, 96, 0)
    End Select
    
    .SelStart = Len(.Text)
    
    If Len(.Text) = 0 Then
      .SelText = sData
    Else
      If nLen > 1 Then
        .SelText = vbNewLine & sData
      Else
        If nLastLen > 1 Then
          .SelText = vbNewLine & sData
        Else
          .SelText = sData
        End If
      End If
    End If
    
    nLastLen = nLen
  End With
End Sub
