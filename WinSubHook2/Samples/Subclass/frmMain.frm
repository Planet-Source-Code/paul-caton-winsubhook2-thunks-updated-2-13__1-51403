VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Subclass..."
   ClientHeight    =   7710
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   775
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picHeader 
      Height          =   300
      Left            =   -15
      ScaleHeight     =   240
      ScaleWidth      =   8505
      TabIndex        =   11
      Top             =   0
      Width           =   8565
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "######## When.. lReturn. hWnd.... uMsg.... wParam.. lParam.. Message name....... "
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -15
         TabIndex        =   12
         Top             =   -30
         UseMnemonic     =   0   'False
         Width           =   8580
      End
   End
   Begin VB.PictureBox pic 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7710
      Left            =   8550
      ScaleHeight     =   7650
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   0
      Width           =   3075
      Begin VB.CheckBox chkAfter 
         Caption         =   "After original WndProc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   3990
         Width           =   1950
      End
      Begin VB.PictureBox picOptAfter 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   150
         ScaleHeight     =   465
         ScaleWidth      =   2190
         TabIndex        =   6
         Top             =   4320
         Width           =   2190
         Begin VB.OptionButton optAfter 
            Caption         =   "All messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optAfter 
            Caption         =   "Selected messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   270
            Width           =   2175
         End
      End
      Begin VB.CheckBox chkBefore 
         Caption         =   "Before original WndProc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   5
         Top             =   150
         Width           =   2040
      End
      Begin VB.PictureBox picOptBefore 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   150
         ScaleHeight     =   465
         ScaleWidth      =   2190
         TabIndex        =   2
         Top             =   480
         Width           =   2190
         Begin VB.OptionButton optBefore 
            Caption         =   "Selected messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   4
            Top             =   270
            Width           =   2175
         End
         Begin VB.OptionButton optBefore 
            Caption         =   "All messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1455
         End
      End
      Begin Subclass.ucShadow Shadow 
         Left            =   2400
         Top             =   510
         _ExtentX        =   847
         _ExtentY        =   847
         Transparency    =   150
      End
      Begin MSComctlLib.ListView lvBefore 
         Height          =   2640
         Left            =   150
         TabIndex        =   1
         Top             =   1050
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   4657
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4260
         EndProperty
      End
      Begin MSComctlLib.ListView lvAfter 
         Height          =   2640
         Left            =   150
         TabIndex        =   10
         Top             =   4875
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   4657
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4260
         EndProperty
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuItm 
         Caption         =   "Do nothing"
         Index           =   0
      End
      Begin VB.Menu mnuItm 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuItm 
         Caption         =   "E&xit"
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
'Subclass - This form demonstrates the cSubclass class
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'==================================================================================================
Option Explicit

Private Const SW_INVALIDATE As Long = &H2
Private Const LV_KEY        As String = "k"

Private nTxtHeight          As Long                    'Height of a text line
Private nLastLine           As Long                    'Lowest vertical position where a line of text is completely visible
Private nMsgNo              As Long                    'Just a message counter
Private rc                  As WinSubHook2.tRECT       'Scrolling rectangle
Private sc                  As cSubclass               'Declare the subclasser

Implements WinSubHook2.iSubclass                       'Tell VB that we guarantee to implement the interface the iSubclass intercae

'Api declares - well, those we don't already have in WinSubHook2.tlb
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function ScrollWindowEx Lib "user32" (ByVal hWnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As WinSubHook2.tRECT, lprcClip As WinSubHook2.tRECT, ByVal hrgnUpdate As Long, lprcUpdate As Any, ByVal fuScroll As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
  Dim i As WinSubHook2.eMsg
  Dim s As String
  
  With Me
    'Calculate the height of a line of text
    nTxtHeight = .TextHeight("My")
    
    'Adjust the height of the window... like the IntegralHeight property in a listbox
    Height = Height - (((.ScaleHeight Mod nTxtHeight) - 2) * Screen.TwipsPerPixelY)
    DoEvents

    Call Form_Resize
    
    'Start printing at the top
    CurrentY = rc.Top
  End With 'ME

  'Populate the listview controls
  For i = 0 To &H400
    s = GetMsgName(i)
    If Asc(Left$(s, 1)) <> 48 Then
      Call lvBefore.ListItems.Add(, "k" & i, s)
      Call lvAfter.ListItems.Add(, "k" & i, s)
    End If
  Next i
  
  lvBefore.Sorted = True
  lvAfter.Sorted = True
  
  lvBefore.ColumnHeaders(1).Width = 2430#
  lvAfter.ColumnHeaders(1).Width = 2430#
  
  'Initialize the subclasser
  Set sc = New cSubclass
  Call sc.Subclass(Me.hWnd, Me)
End Sub

Private Sub Form_Resize()
  With Me
    If .WindowState <> vbMinimized Then
    
      'Set the extent of the scrolling rectangle
      rc.Left = 0
      rc.Top = nTxtHeight + 3
      rc.Right = .ScaleWidth - .pic.Width
      rc.Bottom = .ScaleHeight
      
      'If the current print position goes below nLastLine then we need to scroll the dc up a line
      nLastLine = .ScaleHeight - nTxtHeight
      
      .picHeader.Width = .ScaleWidth - .pic.Width + 1#
      .lblHeader.Width = .picHeader.ScaleWidth + 30#
    End If
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Kill the subclasser
  Set sc = Nothing
  
  'Fade away...
  Call Shadow.FadeOut(Shadow.FadeTime)
End Sub

'Enable/disable *after* subclassing
Private Sub chkAfter_Click()
  If chkAfter = 0 Then
    optAfter(0).Enabled = False
    optAfter(0).Value = False
    optAfter(1).Enabled = False
    optAfter(1).Value = False
    lvAfter.Enabled = False
    lvAfter.TextBackground = lvwTransparent
    Call Deselect(lvAfter)
    Call sc.DelMsg(ALL_MESSAGES, MSG_AFTER)
  Else
    optAfter(0).Enabled = True
    optAfter(0).Value = False
    optAfter(1).Enabled = True
    optAfter(1).Value = False
  End If
End Sub

'Enable/disable *before* subclassing
Private Sub chkBefore_Click()
  Dim i As Long
  
  If chkBefore.Value = 0 Then
    optBefore(0).Enabled = False
    optBefore(0).Value = False
    optBefore(1).Enabled = False
    optBefore(1).Value = False
    lvBefore.Enabled = False
    lvBefore.TextBackground = lvwTransparent
    Call Deselect(lvBefore)
    Call sc.DelMsg(ALL_MESSAGES, MSG_BEFORE)
  Else
    optBefore(0).Enabled = True
    optBefore(0).Value = False
    optBefore(1).Enabled = True
    optBefore(1).Value = False
  End If
End Sub

'After list check box set/unset
Private Sub lvAfter_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  Dim nMsg As WinSubHook2.eMsg
  
  nMsg = Val(Mid$(Item.Key, 2))
  
  If Item.Checked Then
    Call sc.AddMsg(nMsg, MSG_AFTER)
  Else
    Call sc.DelMsg(Val(Mid$(Item.Key, 2)), MSG_AFTER)
  End If
  
  If nMsg = WM_MOUSEWHEEL Then
    'The mousewheel events will be captured/stolen  by the listview, so set the focus elsewhere
    Call chkAfter.SetFocus
  End If
End Sub

'Before list check box set/unset
Private Sub lvBefore_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  Dim nMsg As WinSubHook2.eMsg
  
  nMsg = Val(Mid$(Item.Key, 2))
  
  If Item.Checked Then
    Call sc.AddMsg(nMsg, MSG_BEFORE)
  Else
    Call sc.DelMsg(nMsg, MSG_BEFORE)
  End If
  
  If nMsg = WM_MOUSEWHEEL Then
    'The mousewheel events will be captured/stolen  by the listview, so set the focus elsewhere
    Call chkBefore.SetFocus
  End If
End Sub

Private Sub mnuItm_Click(Index As Integer)
  If Index = 2 Then Call Unload(Me)
End Sub

'After all or selected
Private Sub optAfter_Click(Index As Integer)
  Dim i As Long
  
  If Index = 0 Then
    lvAfter.Enabled = False
    lvAfter.TextBackground = lvwTransparent
    Call Deselect(lvAfter)
    Call sc.AddMsg(ALL_MESSAGES, MSG_AFTER)
  Else
    lvAfter.TextBackground = lvwOpaque
    lvAfter.Enabled = True
    Call sc.DelMsg(ALL_MESSAGES, MSG_AFTER)
  End If
End Sub

'Before all or selected
Private Sub optBefore_Click(Index As Integer)
  Dim i As Long
  
  If Index = 0 Then
    lvBefore.Enabled = False
    lvBefore.TextBackground = lvwTransparent
    Call Deselect(lvBefore)
    Call sc.AddMsg(ALL_MESSAGES, MSG_BEFORE)
  Else
    lvBefore.TextBackground = lvwOpaque
    lvBefore.Enabled = True
    Call sc.DelMsg(ALL_MESSAGES, MSG_BEFORE)
  End If
End Sub

'Subclasser implemented interface
Private Sub iSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef hWnd As Long, ByRef uMsg As WinSubHook2.eMsg, ByRef wParam As Long, ByRef lParam As Long)
  Dim sWhen As String
  
  If uMsg = WM_PAINT Then
    'If we try to display the paint message we'll just cause another paint message... vicious circle.
    Exit Sub
  End If
  
  If bBefore Then
    sWhen = "Before "
  Else
    sWhen = "After  "
  End If
  
  Call Display(sWhen, lReturn, hWnd, uMsg, wParam, lParam)
End Sub

'Uncheck all of the messages in the passed listview
Private Sub Deselect(ByVal lv As ListView)
  Dim itm As MSComctlLib.ListItem
  
  For Each itm In lv.ListItems
    itm.Checked = False
  Next itm
End Sub

'Display a line
Private Sub Display(ByVal sWhen As String, ByVal lReturn As Long, ByVal lngHWnd As Long, ByVal uMsg As WinSubHook2.eMsg, ByVal wParam As Long, ByVal lParam As Long)
  Dim sMsg    As String

  nMsgNo = nMsgNo + 1         'Just a counter
  sMsg = GetMsgName(uMsg)     'Get the message name

  If CurrentY > nLastLine Then
    
    'If we were to print now, we'd be printing, at least partialy, below the extent of the visible client area, so...
    'Scroll the displayed output up one text line
    Call ScrollWindowEx(Me.hWnd, 0, -nTxtHeight, rc, rc, 0, ByVal 0&, SW_INVALIDATE)
    Call WinSubHook2.UpdateWindow(Me.hWnd)
    CurrentY = nLastLine
  End If

  'Print the line
  Print FmtHex(nMsgNo) & _
        sWhen & _
        FmtHex(lReturn) & _
        FmtHex(lngHWnd) & _
        FmtHex(uMsg) & _
        FmtHex(wParam) & _
        FmtHex(lParam) & _
        sMsg
End Sub

'Return the passed Long value as a hex string with leading zeros, if required, to a width of eight characters, plus a trailing space
Private Function FmtHex(ByVal nValue As Long) As String
  Dim s As String
  
  s = Hex$(nValue)
  FmtHex = String$(8 - Len(s), "0") & s & " "
End Function

'Return the name of the passed messag number
Private Function GetMsgName(ByVal uMsg As WinSubHook2.eMsg) As String
  Select Case uMsg
   Case WinSubHook2.WM_ACTIVATE
    GetMsgName = "WM_ACTIVATE"
   Case WinSubHook2.WM_ACTIVATEAPP
    GetMsgName = "WM_ACTIVATEAPP"
   Case WinSubHook2.WM_ASKCBFORMATNAME
    GetMsgName = "WM_ASKCBFORMATNAME"
   Case WinSubHook2.WM_CANCELJOURNAL
    GetMsgName = "WM_CANCELJOURNAL"
   Case WinSubHook2.WM_CANCELMODE
    GetMsgName = "WM_CANCELMODE"
   Case WinSubHook2.WM_CAPTURECHANGED
    GetMsgName = "WM_CAPTURECHANGED"
   Case WinSubHook2.WM_CHANGECBCHAIN
    GetMsgName = "WM_CHANGECBCHAIN"
   Case WinSubHook2.WM_CHAR
    GetMsgName = "WM_CHAR"
   Case WinSubHook2.WM_CHARTOITEM
    GetMsgName = "WM_CHARTOITEM"
   Case WinSubHook2.WM_CHILDACTIVATE
    GetMsgName = "WM_CHILDACTIVATE"
   Case WinSubHook2.WM_CLEAR
    GetMsgName = "WM_CLEAR"
   Case WinSubHook2.WM_CLOSE
    GetMsgName = "WM_CLOSE"
   Case WinSubHook2.WM_COMMAND
    GetMsgName = "WM_COMMAND"
   Case WinSubHook2.WM_COMPACTING
    GetMsgName = "WM_COMPACTING"
   Case WinSubHook2.WM_COMPAREITEM
    GetMsgName = "WM_COMPAREITEM"
   Case WinSubHook2.WM_COPY
    GetMsgName = "WM_COPY"
   Case WinSubHook2.WM_COPYDATA
    GetMsgName = "WM_COPYDATA"
   Case WinSubHook2.WM_CREATE
    GetMsgName = "WM_CREATE"
   Case WinSubHook2.WM_CTLCOLORBTN
    GetMsgName = "WM_CTLCOLORBTN"
   Case WinSubHook2.WM_CTLCOLORDLG
    GetMsgName = "WM_CTLCOLORDLG"
   Case WinSubHook2.WM_CTLCOLOREDIT
    GetMsgName = "WM_CTLCOLOREDIT"
   Case WinSubHook2.WM_CTLCOLORLISTBOX
    GetMsgName = "WM_CTLCOLORLISTBOX"
   Case WinSubHook2.WM_CTLCOLORMSGBOX
    GetMsgName = "WM_CTLCOLORMSGBOX"
   Case WinSubHook2.WM_CTLCOLORSCROLLBAR
    GetMsgName = "WM_CTLCOLORSCROLLBAR"
   Case WinSubHook2.WM_CTLCOLORSTATIC
    GetMsgName = "WM_CTLCOLORSTATIC"
   Case WinSubHook2.WM_CUT
    GetMsgName = "WM_CUT"
   Case WinSubHook2.WM_DEADCHAR
    GetMsgName = "WM_DEADCHAR"
   Case WinSubHook2.WM_DELETEITEM
    GetMsgName = "WM_DELETEITEM"
   Case WinSubHook2.WM_DESTROY
    GetMsgName = "WM_DESTROY"
   Case WinSubHook2.WM_DESTROYCLIPBOARD
    GetMsgName = "WM_DESTROYCLIPBOARD"
   Case WinSubHook2.WM_DRAWCLIPBOARD
    GetMsgName = "WM_DRAWCLIPBOARD"
   Case WinSubHook2.WM_DRAWITEM
    GetMsgName = "WM_DRAWITEM"
   Case WinSubHook2.WM_DROPFILES
    GetMsgName = "WM_DROPFILES"
   Case WinSubHook2.WM_ENABLE
    GetMsgName = "WM_ENABLE"
   Case WinSubHook2.WM_ENDSESSION
    GetMsgName = "WM_ENDSESSION"
   Case WinSubHook2.WM_ENTERIDLE
    GetMsgName = "WM_ENTERIDLE"
   Case WinSubHook2.WM_ENTERMENULOOP
    GetMsgName = "WM_ENTERMENULOOP"
   Case WinSubHook2.WM_ENTERSIZEMOVE
    GetMsgName = "WM_ENTERSIZEMOVE"
   Case WinSubHook2.WM_ERASEBKGND
    GetMsgName = "WM_ERASEBKGND"
   Case WinSubHook2.WM_EXITMENULOOP
    GetMsgName = "WM_EXITMENULOOP"
   Case WinSubHook2.WM_EXITSIZEMOVE
    GetMsgName = "WM_EXITSIZEMOVE"
   Case WinSubHook2.WM_FONTCHANGE
    GetMsgName = "WM_FONTCHANGE"
   Case WinSubHook2.WM_GETDLGCODE
    GetMsgName = "WM_GETDLGCODE"
   Case WinSubHook2.WM_GETFONT
    GetMsgName = "WM_GETFONT"
   Case WinSubHook2.WM_GETHOTKEY
    GetMsgName = "WM_GETHOTKEY"
   Case WinSubHook2.WM_GETMINMAXINFO
    GetMsgName = "WM_GETMINMAXINFO"
   Case WinSubHook2.WM_GETTEXT
    GetMsgName = "WM_GETTEXT"
   Case WinSubHook2.WM_GETTEXTLENGTH
    GetMsgName = "WM_GETTEXTLENGTH"
   Case WinSubHook2.WM_HOTKEY
    GetMsgName = "WM_HOTKEY"
   Case WinSubHook2.WM_HSCROLL
    GetMsgName = "WM_HSCROLL"
   Case WinSubHook2.WM_HSCROLLCLIPBOARD
    GetMsgName = "WM_HSCROLLCLIPBOARD"
   Case WinSubHook2.WM_ICONERASEBKGND
    GetMsgName = "WM_ICONERASEBKGND"
   Case WinSubHook2.WM_IME_CHAR
    GetMsgName = "WM_IME_CHAR"
   Case WinSubHook2.WM_IME_COMPOSITION
    GetMsgName = "WM_IME_COMPOSITION"
   Case WinSubHook2.WM_IME_COMPOSITIONFULL
    GetMsgName = "WM_IME_COMPOSITIONFULL"
   Case WinSubHook2.WM_IME_CONTROL
    GetMsgName = "WM_IME_CONTROL"
   Case WinSubHook2.WM_IME_ENDCOMPOSITION
    GetMsgName = "WM_IME_ENDCOMPOSITION"
   Case WinSubHook2.WM_IME_KEYDOWN
    GetMsgName = "WM_IME_KEYDOWN"
   Case WinSubHook2.WM_IME_KEYLAST
    GetMsgName = "WM_IME_KEYLAST"
   Case WinSubHook2.WM_IME_KEYUP
    GetMsgName = "WM_IME_KEYUP"
   Case WinSubHook2.WM_IME_NOTIFY
    GetMsgName = "WM_IME_NOTIFY"
   Case WinSubHook2.WM_IME_SELECT
    GetMsgName = "WM_IME_SELECT"
   Case WinSubHook2.WM_IME_SETCONTEXT
    GetMsgName = "WM_IME_SETCONTEXT"
   Case WinSubHook2.WM_IME_STARTCOMPOSITION
    GetMsgName = "WM_IME_STARTCOMPOSITION"
   Case WinSubHook2.WM_INITDIALOG
    GetMsgName = "WM_INITDIALOG"
   Case WinSubHook2.WM_INITMENU
    GetMsgName = "WM_INITMENU"
   Case WinSubHook2.WM_INITMENUPOPUP
    GetMsgName = "WM_INITMENUPOPUP"
   Case WinSubHook2.WM_KEYDOWN
    GetMsgName = "WM_KEYDOWN"
   Case WinSubHook2.WM_KEYFIRST
    GetMsgName = "WM_KEYFIRST"
   Case WinSubHook2.WM_KEYLAST
    GetMsgName = "WM_KEYLAST"
   Case WinSubHook2.WM_KEYUP
    GetMsgName = "WM_KEYUP"
   Case WinSubHook2.WM_KILLFOCUS
    GetMsgName = "WM_KILLFOCUS"
   Case WinSubHook2.WM_LBUTTONDBLCLK
    GetMsgName = "WM_LBUTTONDBLCLK"
   Case WinSubHook2.WM_LBUTTONDOWN
    GetMsgName = "WM_LBUTTONDOWN"
   Case WinSubHook2.WM_LBUTTONUP
    GetMsgName = "WM_LBUTTONUP"
   Case WinSubHook2.WM_MBUTTONDBLCLK
    GetMsgName = "WM_MBUTTONDBLCLK"
   Case WinSubHook2.WM_MBUTTONDOWN
    GetMsgName = "WM_MBUTTONDOWN"
   Case WinSubHook2.WM_MBUTTONUP
    GetMsgName = "WM_MBUTTONUP"
   Case WinSubHook2.WM_MDIACTIVATE
    GetMsgName = "WM_MDIACTIVATE"
   Case WinSubHook2.WM_MDICASCADE
    GetMsgName = "WM_MDICASCADE"
   Case WinSubHook2.WM_MDICREATE
    GetMsgName = "WM_MDICREATE"
   Case WinSubHook2.WM_MDIDESTROY
    GetMsgName = "WM_MDIDESTROY"
   Case WinSubHook2.WM_MDIGETACTIVE
    GetMsgName = "WM_MDIGETACTIVE"
   Case WinSubHook2.WM_MDIICONARRANGE
    GetMsgName = "WM_MDIICONARRANGE"
   Case WinSubHook2.WM_MDIMAXIMIZE
    GetMsgName = "WM_MDIMAXIMIZE"
   Case WinSubHook2.WM_MDINEXT
    GetMsgName = "WM_MDINEXT"
   Case WinSubHook2.WM_MDIREFRESHMENU
    GetMsgName = "WM_MDIREFRESHMENU"
   Case WinSubHook2.WM_MDIRESTORE
    GetMsgName = "WM_MDIRESTORE"
   Case WinSubHook2.WM_MDISETMENU
    GetMsgName = "WM_MDISETMENU"
   Case WinSubHook2.WM_MDITILE
    GetMsgName = "WM_MDITILE"
   Case WinSubHook2.WM_MEASUREITEM
    GetMsgName = "WM_MEASUREITEM"
   Case WinSubHook2.WM_MENUCHAR
    GetMsgName = "WM_MENUCHAR"
   Case WinSubHook2.WM_MENUSELECT
    GetMsgName = "WM_MENUSELECT"
   Case WinSubHook2.WM_MOUSEACTIVATE
    GetMsgName = "WM_MOUSEACTIVATE"
   Case WinSubHook2.WM_MOUSEMOVE
    GetMsgName = "WM_MOUSEMOVE"
   Case WinSubHook2.WM_MOUSEWHEEL
    GetMsgName = "WM_MOUSEWHEEL"
   Case WinSubHook2.WM_MOVE
    GetMsgName = "WM_MOVE"
   Case WinSubHook2.WM_MOVING
    GetMsgName = "WM_MOVING"
   Case WinSubHook2.WM_NCACTIVATE
    GetMsgName = "WM_NCACTIVATE"
   Case WinSubHook2.WM_NCCALCSIZE
    GetMsgName = "WM_NCCALCSIZE"
   Case WinSubHook2.WM_NCCREATE
    GetMsgName = "WM_NCCREATE"
   Case WinSubHook2.WM_NCDESTROY
    GetMsgName = "WM_NCDESTROY"
   Case WinSubHook2.WM_NCHITTEST
    GetMsgName = "WM_NCHITTEST"
   Case WinSubHook2.WM_NCLBUTTONDBLCLK
    GetMsgName = "WM_NCLBUTTONDBLCLK"
   Case WinSubHook2.WM_NCLBUTTONDOWN
    GetMsgName = "WM_NCLBUTTONDOWN"
   Case WinSubHook2.WM_NCLBUTTONUP
    GetMsgName = "WM_NCLBUTTONUP"
   Case WinSubHook2.WM_NCMBUTTONDBLCLK
    GetMsgName = "WM_NCMBUTTONDBLCLK"
   Case WinSubHook2.WM_NCMBUTTONDOWN
    GetMsgName = "WM_NCMBUTTONDOWN"
   Case WinSubHook2.WM_NCMBUTTONUP
    GetMsgName = "WM_NCMBUTTONUP"
   Case WinSubHook2.WM_NCMOUSEMOVE
    GetMsgName = "WM_NCMOUSEMOVE"
   Case WinSubHook2.WM_NCPAINT
    GetMsgName = "WM_NCPAINT"
   Case WinSubHook2.WM_NCRBUTTONDBLCLK
    GetMsgName = "WM_NCRBUTTONDBLCLK"
   Case WinSubHook2.WM_NCRBUTTONDOWN
    GetMsgName = "WM_NCRBUTTONDOWN"
   Case WinSubHook2.WM_NCRBUTTONUP
    GetMsgName = "WM_NCRBUTTONUP"
   Case WinSubHook2.WM_NEXTDLGCTL
    GetMsgName = "WM_NEXTDLGCTL"
   Case WinSubHook2.WM_NULL
    GetMsgName = "WM_NULL"
   Case WinSubHook2.WM_PAINT
    GetMsgName = "WM_PAINT"
   Case WinSubHook2.WM_PAINTCLIPBOARD
    GetMsgName = "WM_PAINTCLIPBOARD"
   Case WinSubHook2.WM_PAINTICON
    GetMsgName = "WM_PAINTICON"
   Case WinSubHook2.WM_PALETTECHANGED
    GetMsgName = "WM_PALETTECHANGED"
   Case WinSubHook2.WM_PALETTEISCHANGING
    GetMsgName = "WM_PALETTEISCHANGING"
   Case WinSubHook2.WM_PARENTNOTIFY
    GetMsgName = "WM_PARENTNOTIFY"
   Case WinSubHook2.WM_PASTE
    GetMsgName = "WM_PASTE"
   Case WinSubHook2.WM_PENWINFIRST
    GetMsgName = "WM_PENWINFIRST"
   Case WinSubHook2.WM_PENWINLAST
    GetMsgName = "WM_PENWINLAST"
   Case WinSubHook2.WM_POWER
    GetMsgName = "WM_POWER"
   Case WinSubHook2.WM_QUERYDRAGICON
    GetMsgName = "WM_QUERYDRAGICON"
   Case WinSubHook2.WM_QUERYENDSESSION
    GetMsgName = "WM_QUERYENDSESSION"
   Case WinSubHook2.WM_QUERYNEWPALETTE
    GetMsgName = "WM_QUERYNEWPALETTE"
   Case WinSubHook2.WM_QUERYOPEN
    GetMsgName = "WM_QUERYOPEN"
   Case WinSubHook2.WM_QUEUESYNC
    GetMsgName = "WM_QUEUESYNC"
   Case WinSubHook2.WM_QUIT
    GetMsgName = "WM_QUIT"
   Case WinSubHook2.WM_RBUTTONDBLCLK
    GetMsgName = "WM_RBUTTONDBLCLK"
   Case WinSubHook2.WM_RBUTTONDOWN
    GetMsgName = "WM_RBUTTONDOWN"
   Case WinSubHook2.WM_RBUTTONUP
    GetMsgName = "WM_RBUTTONUP"
   Case WinSubHook2.WM_RENDERALLFORMATS
    GetMsgName = "WM_RENDERALLFORMATS"
   Case WinSubHook2.WM_RENDERFORMAT
    GetMsgName = "WM_RENDERFORMAT"
   Case WinSubHook2.WM_SETCURSOR
    GetMsgName = "WM_SETCURSOR"
   Case WinSubHook2.WM_SETFOCUS
    GetMsgName = "WM_SETFOCUS"
   Case WinSubHook2.WM_SETFONT
    GetMsgName = "WM_SETFONT"
   Case WinSubHook2.WM_SETHOTKEY
    GetMsgName = "WM_SETHOTKEY"
   Case WinSubHook2.WM_SETREDRAW
    GetMsgName = "WM_SETREDRAW"
   Case WinSubHook2.WM_SETTEXT
    GetMsgName = "WM_SETTEXT"
   Case WinSubHook2.WM_SHOWWINDOW
    GetMsgName = "WM_SHOWWINDOW"
   Case WinSubHook2.WM_SIZE
    GetMsgName = "WM_SIZE"
   Case WinSubHook2.WM_SIZING
    GetMsgName = "WM_SIZING"
   Case WinSubHook2.WM_SIZECLIPBOARD
    GetMsgName = "WM_SIZECLIPBOARD"
   Case WinSubHook2.WM_SPOOLERSTATUS
    GetMsgName = "WM_SPOOLERSTATUS"
   Case WinSubHook2.WM_SYSCHAR
    GetMsgName = "WM_SYSCHAR"
   Case WinSubHook2.WM_SYSCOLORCHANGE
    GetMsgName = "WM_SYSCOLORCHANGE"
   Case WinSubHook2.WM_SYSCOMMAND
    GetMsgName = "WM_SYSCOMMAND"
   Case WinSubHook2.WM_SYSDEADCHAR
    GetMsgName = "WM_SYSDEADCHAR"
   Case WinSubHook2.WM_SYSKEYDOWN
    GetMsgName = "WM_SYSKEYDOWN"
   Case WinSubHook2.WM_SYSKEYUP
    GetMsgName = "WM_SYSKEYUP"
   Case WinSubHook2.WM_TIMECHANGE
    GetMsgName = "WM_TIMECHANGE"
   Case WinSubHook2.WM_TIMER
    GetMsgName = "WM_TIMER"
   Case WinSubHook2.WM_UNDO
    GetMsgName = "WM_UNDO"
   Case WinSubHook2.WM_USER
    GetMsgName = "WM_USER"
   Case WinSubHook2.WM_VKEYTOITEM
    GetMsgName = "WM_VKEYTOITEM"
   Case WinSubHook2.WM_VSCROLL
    GetMsgName = "WM_VSCROLL"
   Case WinSubHook2.WM_VSCROLLCLIPBOARD
    GetMsgName = "WM_VSCROLL"
   Case WinSubHook2.WM_WINDOWPOSCHANGED
    GetMsgName = "WM_WINDOWPOSCHANGED"
   Case WinSubHook2.WM_WINDOWPOSCHANGING
    GetMsgName = "WM_WINDOWPOSCHANGING"
   Case WinSubHook2.WM_WININICHANGE
    GetMsgName = "WM_WININICHANGE"
   Case Else
    GetMsgName = FmtHex(uMsg)
  End Select
End Function



