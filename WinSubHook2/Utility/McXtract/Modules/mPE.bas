Attribute VB_Name = "mPE"
'==================================================================================================
'mPE - Read the PE (Portable Executable) Header
'
Option Explicit

Public Type IMAGEDOSHEADER
  e_magic                        As String * 2
  e_cblp                         As Integer
  e_cp                           As Integer
  e_crlc                         As Integer
  e_cparhdr                      As Integer
  e_minalloc                     As Integer
  e_maxalloc                     As Integer
  e_ss                           As Integer
  e_sp                           As Integer
  e_csum                         As Integer
  e_ip                           As Integer
  e_cs                           As Integer
  e_lfarlc                       As Integer
  e_ovno                         As Integer
  e_res(1 To 4)                  As Integer
  e_oemid                        As Integer
  e_oeminfo                      As Integer
  e_res2(1 To 10)                As Integer
  e_lfanew                       As Long
End Type

Public Type IMAGE_SECTION_HEADER
  NameSec                        As String * 6
  PhysicalAddr                   As Integer
  VirtualSize                    As Long
  VirtualAddress                 As Long
  SizeOfRawData                  As Long
  PointerToRawData               As Long
  PointerToRelocations           As Long
  PointerToLinenumbers           As Long
  NumberOfRelocations            As Integer
  NumberOfLinenumbers            As Integer
  Characteristics                As Long
End Type

Public Type IMAGE_FILE_HEADER
  Machine                        As Integer
  NumberOfSections               As Integer
  TimeDateStamp                  As Long
  PointerToSymbolTable           As Long
  NumberOfSymbols                As Long
  SizeOfOptionalHeader           As Integer
  Characteristics                As Integer
End Type

Public Type IMAGE_DATA_DIRECTORY
  VirtualAddress                 As Long
  Size                           As Long
End Type

Public Type IMAGE_OPTIONAL_HEADER
  Magic                          As Integer
  MajorLinkerVersion             As Byte
  MinorLinkerVersion             As Byte
  SizeOfCode                     As Long
  SizeOfInitializedData          As Long
  SizeOfUninitializedData        As Long
  AddressOfEntryPoint            As Long
  BaseOfCode                     As Long
  BaseOfData                     As Long
  ImageBase                      As Long
  SectionAlignment               As Long
  FileAlignment                  As Long
  MajorOperatingSystemVersion    As Integer
  MinorOperatingSystemVersion    As Integer
  MajorImageVersion              As Integer
  MinorImageVersion              As Integer
  MajorSubsystemVersion          As Integer
  MinorSubsystemVersion          As Integer
  Win32VersionValue              As Long
  SizeOfImage                    As Long
  SizeOfHeaders                  As Long
  CheckSum                       As Long
  Subsystem                      As Integer
  DllCharacteristics             As Integer
  SizeOfStackReserve             As Long
  SizeOfStackCommit              As Long
  SizeOfHeapReserve              As Long
  SizeOfHeapCommit               As Long
  LoaderFlags                    As Long
  NumberOfRvaAndSizes            As Long
  DataDirectory(0 To 15)         As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_NT_HEADERS
  Signature                      As String * 4
  FileHeader                     As IMAGE_FILE_HEADER
  OptionalHeader                 As IMAGE_OPTIONAL_HEADER
End Type

Public HeaderNT                  As IMAGE_NT_HEADERS
Public HeaderSections()          As IMAGE_SECTION_HEADER

Private HeaderDos                As IMAGEDOSHEADER

Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Function ReadPE(Data() As Byte) As Boolean
  Dim i     As Long
  Dim Count As Long

  On Error GoTo Catch
  Call RtlMoveMemory(HeaderDos, Data(0), Len(HeaderDos))

  If HeaderDos.e_magic <> "MZ" Then
    Exit Function
  End If

  Call RtlMoveMemory(HeaderNT, Data(HeaderDos.e_lfanew), Len(HeaderNT))

  Count = Count + HeaderDos.e_lfanew + Len(HeaderNT)

  If HeaderNT.Signature <> "PE" & vbNullChar & vbNullChar Then
    Exit Function
  End If

  ReDim HeaderSections(HeaderNT.FileHeader.NumberOfSections - 1)

  For i = 0 To UBound(HeaderSections)
    Call RtlMoveMemory(HeaderSections(i), Data(Count), Len(HeaderSections(0)))
    Count = Count + Len(HeaderSections(0))
  Next i

  ReadPE = True
Catch:
  On Error GoTo 0
End Function
