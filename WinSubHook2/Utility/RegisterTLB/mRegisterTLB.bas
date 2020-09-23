Attribute VB_Name = "mRegisterTLB"
'==================================================================================================
'Handy little Type Library registration tool
'
Option Explicit

Private Type OPENFILENAME
  lStructSize       As Long
  hwndOwner         As Long
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

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Sub Main()
  Dim sTLB      As String
  Dim o_TypeLib As TypeLibInfo
  Dim of        As OPENFILENAME

  With of
    .Flags = &H281004
    .lpstrDefExt = "tlb"
    .lpstrFile = String$(260, 0)
    .lpstrFilter = "Type Libraries" & vbNullChar & "*.tlb" & vbNullChar & vbNullChar
    .lpstrTitle = "Select a Type Library"
    .nFilterIndex = 1
    .nMaxFile = 260
    .nMaxFileTitle = 260
    .lStructSize = LenB(of)

    If GetOpenFileName(of) = 0 Then Exit Sub
  
    sTLB = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
  End With
  
  If Len(sTLB) > 0 Then
    Set o_TypeLib = TLIApplication.TypeLibInfoFromFile(sTLB)

    If MsgBox("Do you want to register the type library?" & vbNewLine & vbNewLine & "File..." & vbNewLine & "  " & sTLB & vbNewLine & vbNewLine & "HelpString..." & vbNewLine & " '" & o_TypeLib.HelpString & "'" & vbNewLine, vbQuestion Or vbYesNo, "Register Type Library?") = vbYes Then
      o_TypeLib.Register
      MsgBox "Type Library Registered", vbInformation
    End If
  End If

  Set o_TypeLib = Nothing
End Sub
