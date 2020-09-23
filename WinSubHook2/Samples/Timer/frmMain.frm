VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "cTimer Test"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fm 
      Caption         =   "cTimer 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   2
      Left            =   135
      TabIndex        =   6
      Top             =   1735
      Width           =   2445
      Begin VB.Label lblTmr 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   105
         TabIndex        =   7
         Top             =   315
         Width           =   2220
      End
   End
   Begin VB.CommandButton cmdMsgBox 
      Caption         =   "MsgBox"
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
      Height          =   390
      Left            =   1290
      TabIndex        =   5
      Top             =   165
      Width           =   990
   End
   Begin VB.CommandButton cmdTmr 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   165
      Width           =   990
   End
   Begin VB.Timer tmrVB 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fm 
      Caption         =   "VB Timer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   2685
      Width           =   2445
      Begin VB.Label lblTmr 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   2220
      End
   End
   Begin VB.Frame fm 
      Caption         =   "cTimer 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   785
      Width           =   2445
      Begin VB.Label lblTmr 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   2220
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================================
'frmMain - This form demonstrates the cTimer class
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'==================================================================================================
Option Explicit

Private Const SWP_NOMOVE    As Long = &H2
Private Const SWP_NOSIZE    As Long = &H1
Private Const HWND_TOPMOST  As Long = -1

Private bStop               As Boolean
Private nCountVB            As Long
Private nCountCT(1 To 2)    As Long
Private tmrCT(1 To 2)       As cTimer

Implements WinSubHook2.iTimer

Private Declare Function InitCommonControls Lib "comctl32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  Set tmrCT(1) = New cTimer
  Set tmrCT(2) = New cTimer
  
  'Make this window stay on top
  Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tmrCT(1) = Nothing
  Set tmrCT(2) = Nothing
End Sub

Private Sub cmdMsgBox_Click()
  MsgBox "VB timers only work with Forms and UserControls." & vbNewLine & _
         "cTimer will also work with classes... ideal for form-less ActiveX DLL's." & vbNewLine & _
         "VB timers pause when a MsgBox is active whilst running in the IDE." & vbNewLine & _
         "cTimer will continue when a MsgBox is active whilst running in the IDE." & vbNewLine & _
         "cTimer, like VB timers, will pause on an IDE breakpoint." & vbNewLine & _
         "cTimer automatically provides the total elapsed time since timer start." & vbNewLine & _
         "cTimer is faster and lighter than a VB timer.", vbInformation, "Note how the cTimer continues..."
End Sub

Private Sub cmdTmr_Click()
Const INTERVAL_MS As Long = 100
  
  If Not bStop Then
    bStop = True
    nCountVB = 0
    nCountCT(1) = 0
    nCountCT(2) = 0
    cmdTmr.Caption = "Stop"
    cmdMsgBox.Enabled = True
    tmrVB.Interval = INTERVAL_MS
    tmrVB.Enabled = True
    Call tmrCT(1).TmrStart(Me, INTERVAL_MS / 10, 1)
    Call tmrCT(2).TmrStart(Me, INTERVAL_MS, 2)
  Else
    bStop = False
    cmdTmr.Caption = "Start"
    cmdMsgBox.Enabled = False
    tmrVB.Enabled = False
    Call tmrCT(1).TmrStop
    Call tmrCT(2).TmrStop
  End If
End Sub

Private Sub iTimer_Proc(ByVal lElapsedMS As Long, ByVal lTimerID As Long)
  nCountCT(lTimerID) = nCountCT(lTimerID) + 1
  lblTmr(lTimerID).Caption = Format$(nCountCT(lTimerID), "#,###") & " - Elapsed: " & Format$(lElapsedMS / 1000#, "0.00") & "s"
End Sub

Private Sub tmrVB_Timer()
  nCountVB = nCountVB + 1
  lblTmr(3).Caption = Format$(nCountVB, "#,###")
End Sub
