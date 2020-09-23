Attribute VB_Name = "mhWndToExe"
'==================================================================================================
'Return the full exe path\filename of the process that created the passed window handle
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'==================================================================================================

Option Explicit

'Api constants
Private Const PROCESS_QUERY_INFORMATION As Long = &H400&
Private Const PROCESS_VM_READ           As Long = &H10&

'Api declarations
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi" (ByVal hProcess As Long, lphModule As Any, cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, nSize As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Function ExeFileName(ByVal hWnd As Long) As String
Const opFlags       As Long = PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ
Const nMaxMods      As Long = 256
Const nBaseModule   As Long = 1
Const nBytesPerLong As Long = 4
Const MAX_PATH      As Long = 260
  Dim hModules()    As Long
  Dim hProcess      As Long
  Dim nProcessID    As Long
  Dim nBufferSize   As Long
  Dim nBytesNeeded  As Long
  Dim nRet          As Long
  Dim sBuffer       As String
  
  'Get the process ID from the window handle
  Call GetWindowThreadProcessId(hWnd, nProcessID)

  'Open the process so we can read some module info.
  hProcess = OpenProcess(opFlags, False, nProcessID)
  
  If hProcess Then
    'Get list of process modules.
    ReDim hModules(1 To nMaxMods) As Long
    nBufferSize = UBound(hModules) * nBytesPerLong
    nRet = EnumProcessModules(hProcess, hModules(nBaseModule), nBufferSize, nBytesNeeded)
    
    If nRet = False Then
      'Check to see if we need to allocate more space for results.
      If nBytesNeeded > nBufferSize Then
        
        ReDim m_Mods(nBaseModule To nBytesNeeded \ nBytesPerLong) As Long
        nBufferSize = nBytesNeeded
        nRet = EnumProcessModules(hProcess, hModules(nBaseModule), nBufferSize, nBytesNeeded)
      End If
    End If

    'Get the module name.
    sBuffer = Space$(MAX_PATH)
    nRet = GetModuleFileNameEx(hProcess, hModules(nBaseModule), sBuffer, MAX_PATH)
    
    If nRet Then
      ExeFileName = Left$(sBuffer, nRet)
    End If
    
    'Clean up
    Call CloseHandle(hProcess)
  End If
End Function
