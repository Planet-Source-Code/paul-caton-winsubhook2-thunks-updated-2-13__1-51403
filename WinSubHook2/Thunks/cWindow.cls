VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cWindow - module-less, IDE safe, machine code api window thunk
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'
'v1.00 20030107 First cut..........................................................................
'v1.01 20031118 Allow control over callback gating
'               Use global memory for the machine code buffer
'               Reform the assembler...............................................................
'v1.02 20040118 Use EbMode for breakpoint/stop detection rather than callback gating
'               Further reform the assembler for greater speed and smaller size
'               As well DestroyWindow also UnregisterClass on IDE stop
'               Made InIDE public..................................................................
'
'==================================================================================================
Option Explicit

Private Const PATCH_02  As Long = 32                        'Table entry count
Private Const PATCH_0A  As Long = 148                       'Location of class name string

Private nMsgCnt         As Long                             'Msg table entry count
Private aMsgTbl()       As WinSubHook2.eMsg                 'Msg table array
Private nAddrWnd        As Long                             'The address of the WndProc
Private m_sClass        As String                           'Class name
Private m_Owner         As iWindow                          'The owning class/form/usercontrol
Private col_hWnds       As Collection                       'Collection of window handles

'============================================
'Class creation/destruction
'============================================

'Build the window thunk into allocated memory
Private Sub Class_Initialize()
Const PATCH_01 As Long = 18                                 'Code buffer offset to the location of the relative address to EbMode
Const PATCH_05 As Long = 108                                'Relative address of DefWindowProc
Const PATCH_06 As Long = 127                                'Relative address of DestroyWindow
Const PATCH_07 As Long = 132                                'Application hInstance for UnregisterClass
Const PATCH_08 As Long = 137                                'Address of the class name for UnregisterClass
Const PATCH_09 As Long = 142                                'Relative address of UnregisterClass
Const FUNC_EBM As String = "EbMode"                         'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_DWP As String = "DefWindowProcA"                 'DefWindowProc is used to call the default window proc
Const FUNC_DW  As String = "DestroyWindow"                  'DestroyWindow is used if the thunk detects that that the IDE has stopped
Const FUNC_URC As String = "UnregisterClassA"               'UnregisterClass is used if the thunk detects that that the IDE has stopped
Const MOD_VBA5 As String = "vba5"                           'Location of the EbMode function if running VB5
Const MOD_VBA6 As String = "vba6"                           'Location of the EbMode function if running VB6
Const MOD_USER As String = "user32"                         'Location of the SetWindowLong & CallWindowProc functions
  Dim i        As Long                                      'Loop index
  Dim nLen     As Long                                      'String lengths
  Dim sHex     As String                                    'Hex code string
  Dim sCode    As String                                    'Binary code string

  'Store the hex pair machine code representation in sHex
  sHex = "5589E583C4F85731C08945FC8945F8EB0EE8xxxxx01x83F802744485C0745CB9xxxxx02xE33909C9780CBFxxxxx03x8B450CF2AF75298D4514508D4510508D450C508D4508508D45FC508D45F850B8xxxxx04x508B00FF501C837DF8007514FF7514FF7510FF750CFF7508E8xxxxx05x8945FC5F8B45FCC9C21000FF7508E8xxxxx06x68xxxxx07x68xxxxx08xE8xxxxx09xEBDFxxxxx0Ax"
  nLen = Len(sHex)                                          'Length of hex pair string
  
  'Convert the string from hex pairs to bytes and store in the ASCII string opcode buffer
  For i = 1 To nLen Step 2                                  'For each pair of hex characters
    sCode = sCode & ChrB$(Val("&H" & Mid$(sHex, i, 2)))     'Convert a pair of hex characters to a byte and append to the ASCII string
  Next i                                                    'Next pair
  
  nLen = LenB(sCode)                                        'Get the machine code length
  nAddrWnd = GlobalAlloc(0, nLen + 255)                     'Allocate fixed memory for the machine code buffer and the class name string

  'Copy the code to allocated memory
  Call WinSubHook2.CopyMemory(ByVal nAddrWnd, ByVal StrPtr(sCode), nLen)
      
  If InIDE Then
    'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
    Call WinSubHook2.CopyMemory(ByVal nAddrWnd + 15, &H9090, 2)
    
    i = AddrFunc(MOD_VBA6, FUNC_EBM)                        'Get the address of EbMode in vba6.dll
    If i = 0 Then                                           'Found?
      i = AddrFunc(MOD_VBA5, FUNC_EBM)                      'VB5 perhaps, try vba5.dll
    End If

    Debug.Assert i                                          'Ensure the EbMode function was found
    Call PatchRel(PATCH_01, i)                              'Patch the relative address to the EbMode api function
  End If
  
  'Patch the WndProc thunk with runtime values
  Call PatchVal(PATCH_02, 0)                                'Initialize the MsgCount
  Call PatchRel(PATCH_05, AddrFunc(MOD_USER, FUNC_DWP))     'Address of the DefWindowProc api function
  Call PatchRel(PATCH_06, AddrFunc(MOD_USER, FUNC_DW))      'Address of the DestroyWindow api function
  Call PatchVal(PATCH_07, App.hInstance)                    'Application hInstance for UnregisterClass
  Call PatchVal(PATCH_08, nAddrWnd + PATCH_0A)              'Address of the class name for UnregisterClass
  Call PatchRel(PATCH_09, AddrFunc(MOD_USER, FUNC_URC))     'Address of the UnregisterClass function)
  Set col_hWnds = New Collection                            'Create instance of the window handles collection
End Sub

'Destroy window(s) and release the allocated memory
Private Sub Class_Terminate()
  Dim i As Long
  
  Call PatchVal(PATCH_02, 0)                                'Patch the code to ensure no further iWindow_Proc callbacks
  
  For i = col_hWnds.Count To 1 Step -1                      'For each window created (and not yet destroyed)
    Call WinSubHook2.DestroyWindow(col_hWnds.Item(i))       'Destroy the window
    Call col_hWnds.Remove(i)                                'Remove from the collection
  Next i                                                    'Next window
  
  Set col_hWnds = Nothing                                   'Destroy the window handles collection
  
  If Len(m_sClass) > 0 Then                                 'If a class was registered
    Call UnregisterClass(m_sClass, App.hInstance)           'Unregister the window class
  End If
  
  Call GlobalFree(nAddrWnd)                                 'Release the allocated memory
End Sub

'============================================
'Public interface
'============================================

'Add a message to the msg callback table. NB This method can be called at any time
Public Sub AddMsg(ByVal uMsg As WinSubHook2.eMsg)
Const PATCH_03 As Long = 43                                 'Table address
  Dim nEntry   As Long
  
  If uMsg = ALL_MESSAGES Then                               'If ALL_MESSAGES
    nMsgCnt = -1                                            'Indicates that all messages are to callback
  Else                                                      'Else a specific message number
    For nEntry = 1 To nMsgCnt                               'For each existing entry. NB will skip if 0 or -1 (ALL_MESSAGES)
      Select Case aMsgTbl(nEntry)                           'Select on the message number stored in this table entry
      Case -1                                               'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                              'Re-use this entry
        Exit Sub                                            'Bail
      Case uMsg                                             'The msg is already in the table!
        Exit Sub                                            'Bail
      End Select
    Next nEntry                                             'Next entry
    
'Make space for the new entry
    ReDim Preserve aMsgTbl(1 To nEntry)                     'Increase the size of the table. NB nEntry = nMsgCnt + 1
    nMsgCnt = nEntry                                        'Bump the entry count
    aMsgTbl(nEntry) = uMsg                                  'Store the message in the table
  End If
  
  Call PatchVal(PATCH_02, nMsgCnt)                          'Patch the Before table entry count
  Call PatchVal(PATCH_03, AddrMsgTbl())                     'Patch the address of the Before message table. We need do this because there's no guarantee that the table existed at WindowCreate time, the table only gets created if a message number is added.
End Sub

'Arbitarily call the DefWindowProc - Normally, if bHandled isn't set in the WndProc callback, the DefWindowProc
'is called in the assembler thunk after the callback. Use this method to call the DefWindowProc first.
Public Function CallDefWndProc(ByVal hWnd As Long, ByVal uMsg As WinSubHook2.eMsg, ByVal wParam As Long, ByVal lParam) As Long
  CallDefWndProc = WinSubHook2.DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function

'Return the window class name
Public Property Get Class() As String
  Class = m_sClass
End Property

'Call this method to delete a message from the msg table. NB This method can be called at any time
Public Sub DelMsg(ByVal uMsg As WinSubHook2.eMsg)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                               'If deleting all messages (specific or ALL_MESSAGES)
    nMsgCnt = 0                                             'Message count is now zero
    Call PatchVal(PATCH_02, 0)                              'Patch the before table message count
  Else                                                      'Else deleteting a specific message
    For nEntry = 1 To nMsgCnt                               'For each table entry
      If aMsgTbl(nEntry) = uMsg Then                        'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = -1                                'Mark the table slot as available
        Exit For                                            'Bail
      End If
    Next nEntry                                             'Next entry
  End If
End Sub

'Return whether we're running in the IDE. Public for general utility purposes
Public Function InIDE() As Boolean
  Debug.Assert SetTrue(InIDE)
End Function

'Set the window class owner, Form/Class/UserControl
Public Property Set Owner(NewOwner As WinSubHook2.iWindow)
Const PATCH_04 As Long = 79
  
  Set m_Owner = NewOwner
  Call PatchVal(PATCH_04, ObjPtr(m_Owner))                  'Owner object address for iWindow_WndProc
End Property

'Register the window class, call this before creating windows--unless one of the predefined window classes is required.
Public Function WindowClassRegister(ByVal sClass As String, _
                                    Optional ByVal colBackground As Long = &HFFFFFF, _
                                    Optional ByVal Style As WinSubHook2.eClassStyle = 0, _
                                    Optional ByVal hCursor As Long = 0, _
                                    Optional ByVal hIcon As Long = 0, _
                                    Optional ByVal hIconSm As Long = 0, _
                                    Optional ByVal cbClassExtra As Long = 0, _
                                    Optional ByVal cbWndExtra As Long = 0) As Boolean
  Dim wc      As tWNDCLASSEX
  Dim sAscii  As String
  
  Debug.Assert (Len(m_sClass) = 0)                          'This method should only be called once... or never for a predefined class
  m_sClass = sClass                                         'Store the class name
  sAscii = StrConv(m_sClass & vbNullChar, vbFromUnicode)    'Get the ascii (rather than unicode) version of the class name
  
  With wc
    .cbSize = Len(wc)                                       'Size of the window class type
    .cbClsExtra = cbClassExtra                              'Number of class extra bytes
    .cbWndExtra = cbWndExtra                                'Number of window extra bytes
    .hbrBackground = CreateSolidBrush(colBackground)        'Class background
    .hCursor = hCursor                                      'Class cursor
    .hIcon = hIcon                                          'Class icon
    .hIconSm = hIconSm                                      'Class small icon
    .hInstance = App.hInstance                              'Application instance handle
    .lpfnWndProc = nAddrWnd                                 'Class WndProc address
    .Style = Style                                          'Class style
    .lpszClassName = StrPtr(sAscii)                         'Class name
  End With

  'Copy the class name to the thunk
  Call WinSubHook2.CopyMemory(ByVal (nAddrWnd + PATCH_0A), ByVal StrPtr(sAscii), LenB(sAscii))
  
  WindowClassRegister = (RegisterClassEx(wc) <> 0)          'Register the window class
End Function

'Create a window, return the window handle
Public Function WindowCreate(Optional ByVal dwExStyle As WinSubHook2.eWindowStyleEx, _
                             Optional ByVal dwStyle As WinSubHook2.eWindowStyle, _
                             Optional ByVal Class As WinSubHook2.eWindowClass = AS_WINDOWCLASS, _
                             Optional ByVal x As Long = 0, _
                             Optional ByVal y As Long = 0, _
                             Optional ByVal nWidth As Long = 0, _
                             Optional ByVal nHeight As Long = 0, _
                             Optional ByVal sCaption As String = "", _
                             Optional ByVal hWndParent As Long = 0, _
                             Optional ByVal hMenu As Long = 0, _
                             Optional ByVal lParam As Long = 0) As Long
  Dim hWnd    As Long
  Dim sClass  As String
  
  Debug.Assert (Not (m_Owner Is Nothing))                   'LOGIC ERROR! the Owner must be set before calling this method
  
  Select Case Class
    Case WinSubHook2.eWindowClass.AS_WINDOWCLASS:             sClass = m_sClass  'User defined window class
    Case WinSubHook2.eWindowClass.PREDEFINED_BUTTON:          sClass = "BUTTON"  'Predefined window classes
    Case WinSubHook2.eWindowClass.PREDEFINED_COMBOBOX:        sClass = "COMBOBOX"
    Case WinSubHook2.eWindowClass.PREDEFINED_EDIT:            sClass = "EDIT"
    Case WinSubHook2.eWindowClass.PREDEFINED_LISTBOX:         sClass = "LISTBOX"
    Case WinSubHook2.eWindowClass.PREDEFINED_MDICLIENT:       sClass = "MDICLIENT"
    Case WinSubHook2.eWindowClass.PREDEFINED_RICHEDIT:        sClass = "RichEdit"
    Case WinSubHook2.eWindowClass.PREDEFINED_RICHEDIT_CLASS:  sClass = "RICHEDIT_CLASS"
    Case WinSubHook2.eWindowClass.PREDEFINED_SCROLLBAR:       sClass = "SCROLLBAR"
    Case WinSubHook2.eWindowClass.PREDEFINED_STATIC:          sClass = "STATIC"
  End Select
  Debug.Assert (Len(sClass) > 0)                              'LOGIC ERROR! Class name not defined
  
'Create the window
  hWnd = WinSubHook2.CreateWindowEx(dwExStyle, sClass, sCaption, dwStyle, x, y, nWidth, nHeight, hWndParent, hMenu, App.hInstance, lParam)
  Debug.Assert hWnd                                         'CreateWindow failed
  Call col_hWnds.Add(hWnd, "h" & hWnd)                      'Add the window handle to the collection
  WindowCreate = hWnd
End Function

'Destroy window
Public Function WindowDestroy(ByVal hWnd As Long) As Boolean
  Dim sKey As String
  
  On Error GoTo Catch
    sKey = "h" & hWnd
    hWnd = col_hWnds.Item(sKey)                             'Ensure the handle is in the collection
    Call WinSubHook2.DestroyWindow(hWnd)                    'Destroy the window
    Call col_hWnds.Remove(sKey)                             'Remove the handle from the collection
    WindowDestroy = True
Catch:
  On Error GoTo 0
End Function

'============================================
'Private interface
'============================================

'Return the address of the passed function in the passed dll
Private Function AddrFunc(ByVal sDLL As String, _
                          ByVal sProc As String) As Long
  AddrFunc = WinSubHook2.GetProcAddress(WinSubHook2.GetModuleHandle(sDLL), sProc)
  
  'You may want to comment out the following line if you're using vb5 else the EbMode
  'GetProcAddress will stop here everytime because we look in vba6.dll first
  Debug.Assert AddrFunc
End Function

'Return the address of the low bound of the passed table array
Private Function AddrMsgTbl() As Long
  On Error Resume Next                                      'The table may not be dimensioned yet so we need protection
    AddrMsgTbl = VarPtr(aMsgTbl(1))                         'Get the address of the first element of the passed message table
  On Error GoTo 0                                           'Switch off error protection
End Function

'Patch the machine code buffer offset with the relative address to the target address
Private Sub PatchRel(ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call WinSubHook2.CopyMemory(ByVal (nAddrWnd + nOffset), nTargetAddr - nAddrWnd - nOffset - 4, 4)
End Sub

'Patch the machine code buffer offset with the passed value
Private Sub PatchVal(ByVal nOffset As Long, ByVal nValue As Long)
  Call WinSubHook2.CopyMemory(ByVal (nAddrWnd + nOffset), nValue, 4)
End Sub

'Worker function for InIDE - will only be called whilst running in the IDE
Private Function SetTrue(ByRef bValue As Boolean) As Boolean
  SetTrue = True
  bValue = True
End Function
