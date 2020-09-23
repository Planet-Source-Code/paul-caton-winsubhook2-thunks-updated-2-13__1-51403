VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   6495
   ClientLeft      =   2190
   ClientTop       =   2235
   ClientWidth     =   9975
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
   ScaleHeight     =   6495
   ScaleWidth      =   9975
   Begin RichTextLib.RichTextBox rt 
      Height          =   5610
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   9895
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      RightMargin     =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItm 
         Caption         =   "&Open"
         Index           =   0
         Shortcut        =   ^O
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
         Shortcut        =   ^C
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
   Begin VB.Menu mnuAbout 
      Caption         =   "About!"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================================
'McXtract
'
' Extract a hex string representation from the code of an executable or bin file.
' Look for patch-marker sequences "ABCABC??", substitute with "xxxxx??x". Create
' Const's to indicate the patch-marker buffer offsets.
'
Option Explicit

Private Const PATH              As String = "Path"
Private Const OFN_HIDEREADONLY  As Long = &H4

Private Type OPENFILENAME
  lStructSize       As Long
  hWndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String
End Type

Private ofd As OPENFILENAME

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
  Dim nWidth  As Single
  Dim nHeight As Single
  
  Call Clipboard.Clear

  With ofd
    .lStructSize = Len(ofd)
    .hInstance = App.hInstance
    .hWndOwner = Me.hWnd
    .lpstrFilter = "Binary files (*.bin)" & vbNullChar & "*.bin" & vbNullChar & "Executables (*.exe *.dll)" & vbNullChar & "*.exe;*.dll" & vbNullChar & "All files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    .nMaxFile = 255
    .nMaxFileTitle = 255
    .lpstrInitialDir = GetSetting(App.Title, PATH, PATH, CurDir)
    .lpstrTitle = "Select a binary/executable..."
    .Flags = OFN_HIDEREADONLY
  End With
  
  Me.Caption = App.Title
  nWidth = (Me.Width - Me.ScaleWidth) + rt.Width + 480
  nHeight = (Me.Height - Me.ScaleHeight) + rt.Top + rt.Height + 240

  Call Move((Screen.Width - nWidth) / 2, (Screen.Height - nHeight) / 4, nWidth, nHeight)
  Call Show
  DoEvents

  Call Form_Resize
  Call OpenFile
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  rt.Text = vbNullString
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  Call rt.Move(0, 0, Me.ScaleWidth, Me.ScaleHeight)
  On Error GoTo 0
End Sub

Private Sub mnuAbout_Click()
  Call MsgBox(App.Title & " - " & App.Comments & vbNewLine & vbNewLine & _
             "Version: " & App.Major & "." & Format$(App.Minor, "#0") & "." & Format$(App.Revision, "0####"), _
              vbInformation, _
             "About " & App.Title)
End Sub

Private Sub mnuEdit_Click()
  Dim nLen As Long
  Dim nSel As Long
  
  With rt
    nLen = Len(.Text)
    nSel = .SelLength
  End With 'RT
  
  mnuEditItm(0).Enabled = (nSel > 0)
  mnuEditItm(2).Enabled = (nLen > 0) And (nLen <> nSel)
End Sub

Private Sub mnuEditItm_Click(Index As Integer)
  Select Case Index
  Case 0
    Call Clipboard.SetText(rt.SelText)
  Case 2
    With rt
      .SelStart = 0
      .SelLength = Len(.Text)
    End With 'RT
  End Select
End Sub

Private Sub mnuFileItm_Click(Index As Integer)
  Select Case Index
  Case 0
    Call OpenFile
  Case 2
    Call Unload(Me)
  End Select
End Sub

Private Sub ColorCode(ByVal sSearch As String, ByVal nColor As Long, Optional ByVal nExtra As Long = 0)
  Dim nPos As Long
  
  With rt
Again:
    nPos = .Find(sSearch, nPos)
    If nPos <> -1 Then
      .SelLength = .SelLength + nExtra
      .SelColor = nColor
      nPos = nPos + 1
      GoTo Again
    End If
  End With 'RT
End Sub

Private Function FindByVirtual(ByVal ValX As Long) As Long
  Dim i As Long
  
  For FindByVirtual = 0 To UBound(HeaderSections)
    i = HeaderSections(FindByVirtual).VirtualAddress
    If ValX >= i Then
      If ValX <= i + HeaderSections(FindByVirtual).VirtualSize Then
        Exit Function
      End If
    End If
  Next FindByVirtual

  FindByVirtual = -1
End Function

Private Function FindExecutive() As Long
  FindExecutive = FindByVirtual(HeaderNT.OptionalHeader.AddressOfEntryPoint)

  If FindExecutive = -1 Then
    FindExecutive = FindByVirtual(HeaderNT.OptionalHeader.BaseOfCode)
  End If
End Function

Private Function GetPath(ByVal sFile As String) As String
  Dim i As Long

  i = InStrRev(sFile, "\")
  If i Then
    GetPath = Left$(sFile, i - 1)
  Else 'I = FALSE/0
    GetPath = vbNullString
  End If
End Function

Private Sub OpenFile()
  Dim bBin                As Boolean
  Dim Data()              As Byte
  Dim i                   As Long
  Dim j                   As Long
  Dim k                   As Long
  Dim nPos                As Long
  Dim nSecn               As Long
  Dim s                   As String
  Dim sTemp               As String
  Dim sOff                As String
  Dim sPatch              As String
  Dim a()                 As String
  
  With ofd
    .lpstrFile = Space$(254)
    .lpstrFileTitle = Space$(254)
    .lpstrInitialDir = GetSetting(App.Title, PATH, PATH, App.PATH)
    If GetOpenFileName(ofd) = 0 Then
      Exit Sub
    End If

    Call SaveSetting(App.Title, PATH, PATH, GetPath(.lpstrFile))
    s = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
  End With 'OFD
  
  Me.Caption = App.Title & " " & s

  Call ChDir(App.PATH)
  rt.Text = vbNullString

  Open s For Binary As #1

  ReDim Data(LOF(1) - 1)
  Get #1, , Data
  Close #1

  bBin = (StrComp(Right$(s, 4), ".bin", vbTextCompare) = 0)

  If Not bBin Then
    If mPE.ReadPE(Data) = 0 Then
      Call MsgBox("File isn't a Win32 Executeable!", vbCritical)
      Exit Sub
    End If

    nSecn = FindExecutive
    If nSecn = -1 Then
      Call MsgBox("Cannot trace the executive Code!", vbCritical)
      Exit Sub
    End If
  End If

  Screen.MousePointer = vbHourglass
  rt.Text = vbNullString
  DoEvents

  If bBin Then
    i = 0
    j = UBound(Data()) + 1
    k = j
  Else 'BBIN = FALSE/0
    i = mPE.HeaderSections(nSecn).PointerToRawData
    j = mPE.HeaderSections(nSecn).VirtualSize
    k = i + j
  End If

  ReDim a(j)

  j = 0
  Do While i < k
    s = Hex$(Data(i))
    If Len(s) = 2 Then
      a(j) = s
    Else 'NOT LEN(S)...
      a(j) = "0" & s
    End If

    i = i + 1
    j = j + 1
  Loop

  s = Join(a, vbNullString)
  Erase a

  If Len(s) = 0 Then
    GoTo Bail
  End If

  sTemp = s

  Do
    nPos = InStr(1, sTemp, "ABCABC")
    If nPos = 0 Then
      Exit Do
    End If

    sPatch = Hex$(Val("&H" & Mid$(sTemp, nPos + 6, 2)))

    If Len(sPatch) = 1 Then
      sPatch = "0" & sPatch
    End If

    sTemp = Left$(sTemp, nPos - 1) & "xxxxx" & sPatch & "x" & Mid$(sTemp, nPos + 8)
    sOff = sOff & "Const PATCH_" & sPatch & " As Long = " & CInt(nPos \ 2) & vbNewLine
  Loop

  With rt
    Call LockWindowUpdate(.hWnd)

    .Text = sOff & "Const CODE_STR As String = """ & sTemp & """"
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelColor = 0

    Call ColorCode("Const", RGB(0, 0, 192))
    Call ColorCode("As Long", RGB(0, 0, 192))
    Call ColorCode("As String", RGB(0, 0, 192))
    Call ColorCode("xxxxx", RGB(192, 0, 0), 3)

    .SelStart = 0
    .SelLength = 0

    Call LockWindowUpdate(0)
  End With 'rt

Bail:
  Screen.MousePointer = vbDefault
End Sub
