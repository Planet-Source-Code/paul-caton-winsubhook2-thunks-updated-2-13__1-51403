VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cShift - Fast assembler bitshift class.
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'==================================================================================================
Option Explicit

Private pMe     As Long
Private nEntry  As Long
Private sCode() As String

Private Sub Class_Initialize()
  Call WinSubHook2.CopyMemory(pMe, ByVal ObjPtr(Me), 4)
  Call Inject("8B4424088B4C240C8B542410D3E0890231C0C21000") 'ShiftLeft machine-code
  Call Inject("8B4424088B4C240C8B542410D3E8890231C0C21000") 'ShiftRight machine-code
End Sub

Public Function ShiftLeft(ByVal lngValue As Long, ByVal lngShift As Long) As Long
End Function

Public Function ShiftRight(ByVal lngValue As Long, ByVal lngShift As Long) As Long
End Function

Private Sub Inject(ByVal sHexCode As String)
  Dim i As Long
  Dim s As String
  
  ReDim Preserve sCode(0 To nEntry)
  
  For i = 1 To Len(sHexCode) Step 2
    s = s & ChrB$(Val("&H" & Mid$(sHexCode, i, 2)))
  Next i
  
  sCode(nEntry) = s
  Call WinSubHook2.CopyMemory(ByVal pMe + &H1C + (nEntry * 4), StrPtr(sCode(nEntry)), 4)
  nEntry = nEntry + 1
End Sub

