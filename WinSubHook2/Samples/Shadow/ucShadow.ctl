VERSION 5.00
Begin VB.UserControl ucShadow 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   87
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   82
End
Attribute VB_Name = "ucShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================================================================================
'ucShadow - Fades and shadows the parent form.
'
'This is the compiled in variant of the ucShadow control. If you prefer to use compiled ocx's then
'make the project in the "Shadow ocx" directory.
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'
'v1.0.0000 20030112 First cut..........................................................................
'v1.0.0001 20030203 Improve comments...................................................................
'v1.0.0002 20030205 Change from SetLayeredWindowAttributes to UpdateLayeredWindow, better performance..
'v1.0.0003 20030213 Moved UpdateLayeredWindow out of the type library
'                   New corner algorithm from vbAccelerator............................................
'v1.0.0004 20030217 Detect XP's Luna theme, don't just assume Luna if XP
'                   Added IsLuna property
'                   Detect display changes
'                   Detect theme changes
'                   Fixed issue with depths < 5 on XP
'                   Fixed a 'No Aliasing' compiler optimization problem in the FadeOut sub.............
'v1.0.0005 20031118 Converted to WinSubHook2...................................................
'======================================================================================================
Option Explicit

'Public events
Public Event FadedIn()                                                'FadeIn completed event
Attribute FadedIn.VB_Description = "Event raised on completion of fade in."
Public Event FadedOut()                                               'FadeOut completed event
Attribute FadedOut.VB_Description = "Event raised on completion of fade out."

'Defines a dimension within a VB array header/descriptor block
Private Type tSAFEARRAYBOUND
  cElements       As Long                                             'Number of elements within an array dimension
  lLbound         As Long                                             'The dimensions lowest bound
End Type

'VB array header/descriptor block for a 2 dimension array
Private Type tSAFEARRAY2D
  cDims           As Integer                                          'Number of dimensions (Bounds array)
  fFeatures       As Integer                                          'Array features
  cbElements      As Long                                             'Bytes per element (4 = Long, 2 = Integer ...)
  cLocks          As Long                                             'Array locks
  pvData          As Long                                             'Memory address of the array data
  Bounds(0 To 1)  As tSAFEARRAYBOUND                                  'Two dimensions
End Type

'Api function to return the memory address of a VB safearray header block (tSAFEARRAY)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" ( _
  Ptr() As Any) As Long

'This function can't be referenced via the type library because it won't exist on pre Win2k/XP systems.
Private Declare Function UpdateLayeredWindow Lib "user32.dll" ( _
  ByVal hWnd As Long, _
  ByVal hDcDest As Long, _
  ptDst As Any, _
  pSize As Any, _
  ByVal hDcSrc As Long, _
  ptSrc As Any, _
  ByVal crKey As Long, _
  pBlend As Any, _
  ByVal dwFlags As Long) As Long

'These functions can't be referenced via the type library because they won't exist on pre Windows XP systems.
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" ( _
  ByVal pszThemeFileName As Long, _
  ByVal cchMaxNameChars As Long, _
  ByVal pszColorBuff As Long, _
  ByVal cchMaxColorChars As Long, _
  ByVal pszSizeBuff As Long, _
  ByVal cchMaxSizeChars As Long) As Long

Private Declare Function GetThemeDocumentationProperty Lib "uxtheme.dll" ( _
  ByVal pszThemeName As Long, _
  ByVal pszPropertyName As Long, _
  ByVal pszValueBuff As Long, _
  ByVal cchMaxValChars As Long) As Long

'Property constants
Private Const PRP_COLOR As String = "Color"                           'Color property name
Private Const DEF_COLOR As Long = 0                                   'Default Color, black
Private Const PRP_DEPTH As String = "Depth"                           'Depth property name
Private Const DEF_DEPTH As Long = 10                                  'Default Depth
Private Const PRP_FADES As String = "FadeIn"                          'FadeIn property name
Private Const DEF_FADES As Boolean = True                             'Default FadeIn
Private Const PRP_FADET As String = "FadeTime"                        'FadeTime property name
Private Const DEF_FADET As Long = 500                                 'Default FadeTime
Private Const PRP_HIDEM As String = "HideMove"                        'HideMove property name
Private Const DEF_HIDEM As Boolean = False                            'Default HideMove
Private Const PRP_HIDES As String = "HideSize"                        'HideSize property name
Private Const DEF_HIDES As Boolean = False                            'Default HideSize
Private Const PRP_SOFTS As String = "SoftShadow"                      'SoftShadow property name
Private Const DEF_SOFTS As Boolean = True                             'Default SoftShadow
Private Const PRP_TRANS As String = "Transparency"                    'Transparency property name
Private Const DEF_TRANS As Long = 120                                 'Default Transparency
Private Const PRP_SHOWN As String = "Visible"                         'Visible property name
Private Const DEF_SHOWN As Boolean = True                             'Default Visible

Private Const FADE_STEP As Long = 20                                  'Fade timer interval step in milliseconds

'Member variables
Private m_Color         As OLE_COLOR                                  'Color property, shadow color
Private m_Depth         As Long                                       'Depth property, depth of the shadow
Private m_FadeS         As Boolean                                    'FadeIn property
Private m_FadeT         As Long                                       'FadeTime property
Private m_HideM         As Boolean                                    'HideMove property, whether the shadow is shown when the the parent is moved
Private m_HideS         As Boolean                                    'HideSize property, whether the shadow is shown when the the parent is sized
Private m_SoftS         As Boolean                                    'Soft shadow property, wether the shadow has a soft edge
Private m_Trans         As Long                                       'Transparency property, transparency of the shadow
Private m_Shown         As Boolean                                    'Visible property, whether the shadows are shown

'Member read-only values
Private m_IsLayered     As Boolean                                    'Layered windows supported
Private m_IsLuna        As Boolean                                    'Luna theme?
Private m_IsXP          As Boolean                                    'Windows XP?

'Control variables
Private bBlock          As Boolean                                    'Wether to block during FadeOut
Private bFadeIn         As Boolean                                    'Wether we're fading In or Out
Private nColor          As Long                                       'Translated m_Color
Private nFaderStart     As Long                                       'Fade start time
Private nFadeTime       As Long                                       'Fade duration
Private hWndBt          As Long                                       'Bottom shadow window handle
Private hWndRt          As Long                                       'Right shadow window handle
Private hWndFader       As Long                                       'Fader window handle
Private hWndParent      As Long                                       'Parent window handle
Private wp              As tWINDOWPOS                                 'Parent window position
Private bf              As tBLENDFUNCTION                             'Fader blend function
Private scParent        As cSubclass                                  'Parent subclasser
Private wnFader         As cWindow                                    'Fader window class
Private wnShadow        As cWindow                                    'Shadow window class

Implements WinSubHook2.iSubclass                                      'Guarantee that we'll implement the iSubclass interface
Implements WinSubHook2.iWindow                                        'Guarantee that we'll implement the iWindow interface

'---------------------------------------------------------------------
'PROPERTIES

'Return the shadow color
Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Return/set the shadow color."
  Color = m_Color
End Property

'Assign the shadow color
Public Property Let Color(NewValue As OLE_COLOR)
  If NewValue <> m_Color Then                                         'If the new value differs from the current value
    m_Color = NewValue                                                'Store the new value
    If Ambient.UserMode Then                                          'If we're running, not designing
      nColor = TranslateColor(m_Color)                                'Translate system color indices
      Call ShadowCreate                                               'Re-create the Shadow
    Else                                                              'If we're in design mode, not running
      PropertyChanged PRP_COLOR                                       'Tell the container
    End If
  End If
End Property

'Return the shadow depth
Public Property Get Depth() As Long
Attribute Depth.VB_Description = "Return/set the shadow depth"
  Depth = m_Depth
End Property

'Assign the shadow depth
Public Property Let Depth(NewValue As Long)
  If NewValue < 0 Then
    NewValue = 0
  ElseIf NewValue > 99 Then
    NewValue = 99                                                     'Ensure the value is sensible-ish
  End If
  If NewValue <> m_Depth Then                                         'If the new value differs from the current value
    m_Depth = NewValue                                                'Store the new value
    If Ambient.UserMode Then                                          'If we're running
      Call ShadowCreate                                               'Resize the shadow
    Else                                                              'If we're not running (design mode)
      PropertyChanged PRP_DEPTH                                       'Tell the container
    End If
  End If
End Property

'Return the FadeIn setting
Public Property Get FadeIn() As Boolean
Attribute FadeIn.VB_Description = "Return/set whether the parent form will be faded in on show."
  FadeIn = m_FadeS
End Property

'Assign the FadeIn setting
Public Property Let FadeIn(NewValue As Boolean)
  If NewValue <> m_FadeS Then                                         'If the new value differs from the current value
    m_FadeS = NewValue                                                'Store the new value
    If Not Ambient.UserMode Then                                      'If we're not running
      PropertyChanged PRP_FADES                                       'Tell the container
    End If
  End If
End Property

'Return the FadeTime
Public Property Get FadeTime() As Long
Attribute FadeTime.VB_Description = "Return/set the duration in milliseconds of a fade in."
  FadeTime = m_FadeT
End Property

'Assign the FadeTime
Public Property Let FadeTime(NewValue As Long)
  If NewValue <> m_Depth Then                                         'If the new value differs from the current value
    m_FadeT = NewValue                                                'Store the new value
    If Not Ambient.UserMode Then                                      'If we're not running (design mode)
      PropertyChanged PRP_FADET                                       'Tell the container
    End If
  End If
End Property

'Return the HideMove setting
Public Property Get HideMove() As Boolean
Attribute HideMove.VB_Description = "Return/set whether the shadows should disappear whilst the parent form is being moved."
  HideMove = m_HideM
End Property

'Assign the HideMove setting
Public Property Let HideMove(NewValue As Boolean)
  If NewValue <> m_HideM Then                                         'If the new value differs from the current value
    m_HideM = NewValue                                                'Store the new value
    If Not Ambient.UserMode Then                                      'If we're not running (design mode)
      PropertyChanged PRP_HIDEM                                       'Tell the container
    End If
  End If
End Property

'Return the HideSize setting
Public Property Get HideSize() As Boolean
Attribute HideSize.VB_Description = "Return/set whether the shadows should disappear whilst the parent form is being sized."
  HideSize = m_HideS
End Property

'Assign the HideSize setting
Public Property Let HideSize(NewValue As Boolean)
  If NewValue <> m_HideS Then                                         'If the new value differs from the current value
    m_HideS = NewValue                                                'Store the new value
    If Not Ambient.UserMode Then                                      'If we're not running (design mode)
      PropertyChanged PRP_HIDES                                       'Tell the container
    End If
  End If
End Property

'Return whether the OS/screen settings support layered windows
Public Property Get IsLayered() As Boolean
Attribute IsLayered.VB_Description = "Return whether layered windows are supported. OS / color depth."
  IsLayered = m_IsLayered
End Property

'Return whether the Windows XP Luna theme is running
Public Property Get IsLuna() As Boolean
Attribute IsLuna.VB_Description = "If running on Windows XP returns whether the Luna theme is active."
  IsLuna = m_IsLuna
End Property

'Return the system setting indicating whether fade animations should be used.
'It's up to the programmer to honor this setting or not.
Public Property Get IsSysFadeEnabled() As Boolean
Attribute IsSysFadeEnabled.VB_Description = "Return whether fading is enabled at the OS level. It is up to the programmer to decide whether this setting is honored."
Dim nFadeEnabled As Long

  Call SystemParametersInfo(SPI_GETMENUANIMATION, 0, nFadeEnabled, 0)
  IsSysFadeEnabled = (nFadeEnabled <> 0)
End Property

'Return the system setting indicating whether drop shadows should be shown, only truly valid on XP.
'It's up to the programmer to honor this setting or not.
Public Property Get IsSysShadowEnabled() As Boolean
Attribute IsSysShadowEnabled.VB_Description = "Return whether menu shadows are enabled at the OS level (only valid on Windows XP). It is up to the programmer to decide whether this setting is honored."
Dim nDropShadow As Long
  
  If m_IsXP Then
    Call SystemParametersInfo(SPI_GETDROPSHADOW, 0, nDropShadow, 0)
    IsSysShadowEnabled = (nDropShadow <> 0)
  Else
    IsSysShadowEnabled = True
  End If
End Property

'Return whether we're running on XP
Public Property Get IsXP() As Boolean
Attribute IsXP.VB_Description = "Return whether Windows XP is the OS."
  IsXP = m_IsXP
End Property

'Return the soft shadow setting
Public Property Get SoftShadow() As Boolean
Attribute SoftShadow.VB_Description = "Return/set whether to display the shadow with soft edges."
  SoftShadow = m_SoftS
End Property

'Assign the soft shadow setting
Public Property Let SoftShadow(NewValue As Boolean)
  If NewValue <> m_SoftS Then                                         'If the new value differs from the current value
    m_SoftS = NewValue                                                'Store the new value
    If Ambient.UserMode Then                                          'If we're running
      Call ShadowCreate                                               'Create the shadows
    Else                                                              'If we're not running (design mode)
      PropertyChanged PRP_SOFTS                                       'Tell the container
    End If
  End If
End Property

'Return the transparency
Public Property Get Transparency() As Long
Attribute Transparency.VB_Description = "Return/set the shadow transparency."
  Transparency = m_Trans
End Property

'Assign the transparency
Public Property Let Transparency(NewValue As Long)
  NewValue = NewValue Mod 256                                         'Ensure the new value doesn't overflow                                        '
  If NewValue <> m_Trans Then                                         'If the new value differs from the current value
    m_Trans = NewValue                                                'Store the new value
    If Ambient.UserMode Then                                          'If we're running
      Call ShadowCreate                                               'Re-create the shadow
    Else                                                              'If we're not running (design mode)
      PropertyChanged PRP_TRANS                                       'Tell the container
    End If
  End If
End Property

'Return the shadow visibility
Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Return/set whether the shadow is visiblle."
  Visible = m_Shown
End Property

'Assign the shadow visibility
Public Property Let Visible(NewValue As Boolean)
  If NewValue <> m_Shown Then                                         'If the new value differs from the current value
    m_Shown = NewValue                                                'Store the new value
    If Ambient.UserMode Then                                          'If we're running
      If IsWindowVisible(hWndParent) = 1 Then                         'If the parent is visible
        Call ShadowShow(m_Shown, True)                                'Set the visibility
      End If
    Else
      PropertyChanged PRP_SHOWN                                       'Tell the container
    End If
  End If
End Property

'---------------------------------------------------------------------
'METHODS

'Why a method for FadeOut rather than a property like FadeIn? - 'Cos on the unload of a form will
'take the fader with it just after fade start, or done another way prevent the form unload, or...
'
'Use the block parameter set to True if the form is unloading, if merely hiding then you takes your choice... don't forget
'that an event will be raised on completion if you wish to know when the fade was completed but didn't block.
Public Sub FadeOut(Optional nTimeMS As Long = 500, Optional Block As Boolean = True)
Attribute FadeOut.VB_Description = "Method called to fade the form out."
  If m_IsLayered Then                                                 'If we have the transparency support
    bBlock = Block                                                    'Store wether to block
    nFadeTime = nTimeMS                                               'Store the fade time
    bFadeIn = False                                                   'Store that we're fading out
    Call FaderCreate                                                  'Create the fader window
    Call ShowWindow(hWndParent, SW_HIDE)                              'Hide the parent
    
    Do While Blocking                                                 'Loop until fade completed
      DoEvents
    Loop
  End If
End Sub

'Do While bBlock causes a problem if you compile with 'No aliasing'. Shipping
'the read of the variable into a separate function prevents aliasing.
Private Function Blocking() As Boolean
  Blocking = bBlock
End Function

'---------------------------------------------------------------------
'USERCONTROL EVENTS

'This event is raised when a shadow control is first placed on a form
Private Sub UserControl_InitProperties()
  m_Color = DEF_COLOR                                                 'Set the default Color
  m_Depth = DEF_DEPTH                                                 'Set the default Depth
  m_FadeS = DEF_FADES                                                 'Set the default FadeIn
  m_FadeT = DEF_FADET                                                 'Set the default FadeTime
  m_HideM = DEF_HIDEM                                                 'Set the default HideMove
  m_HideS = DEF_HIDES                                                 'Set the default HideSize
  m_SoftS = DEF_SOFTS                                                 'Set the default SoftShadow
  m_Trans = DEF_TRANS                                                 'Set the default Transparency
  m_Shown = DEF_SHOWN                                                 'Set the default Visibile
  Debug.Assert (TypeName(Parent) = "Form")                            'The ucShadow control must be placed on a form NOT a UserControl
End Sub

'This UserControl is invisible at runtime therefore the paint event only fires at design time. We could of course, do
'the usual thing, display a bitmap but done this way we avoid the bitmap overhead and also use the appropriate system
'colors in our little representation of a form with a shadow.
Private Sub UserControl_Paint()
Const COL_BTNFACE       As Long = 15, _
      COL_BTNSHADOW     As Long = 16, _
      COL_HOTLIGHT      As Long = 26
  
'Paint the UserControl (at design time) to look like a form with a shadow.

  With UserControl
'Caption area
    Call Rectangle(.hDc, 1, 1, 27, 6, TranslateColor(COL_SYS_MASK Or COL_HOTLIGHT))
'Client area
    Call Rectangle(.hDc, 1, 8, 27, 20, TranslateColor(COL_SYS_MASK Or COL_BTNFACE))
'Bottom background
    Call Rectangle(.hDc, 0, 29, 3, 3, TranslateColor(COL_SYS_MASK Or COL_BTNFACE))
'Bottom shadow
    Call Rectangle(.hDc, 3, 29, 29, 3, TranslateColor(COL_SYS_MASK Or COL_BTNSHADOW))
'Right background
    Call Rectangle(.hDc, 29, 0, 3, 3, TranslateColor(COL_SYS_MASK Or COL_BTNFACE))
'Right shadow
    Call Rectangle(.hDc, 29, 3, 3, 27, TranslateColor(COL_SYS_MASK Or COL_BTNSHADOW))
  End With
End Sub

'Read the properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    m_Color = .ReadProperty(PRP_COLOR, DEF_COLOR)                     'Shadow color
    m_Depth = .ReadProperty(PRP_DEPTH, DEF_DEPTH)                     'Shadow depth
    m_FadeS = .ReadProperty(PRP_FADES, DEF_FADES)                     'Form FadeIn
    m_FadeT = .ReadProperty(PRP_FADET, DEF_FADET)                     'FadeIn time
    m_HideM = .ReadProperty(PRP_HIDEM, DEF_HIDEM)                     'Hide shadow on move
    m_HideS = .ReadProperty(PRP_HIDES, DEF_HIDES)                     'Hide shadow on size
    m_SoftS = .ReadProperty(PRP_SOFTS, DEF_SOFTS)                     'Soft shadow
    m_Trans = .ReadProperty(PRP_TRANS, DEF_TRANS)                     'Shadow transparency
    m_Shown = .ReadProperty(PRP_SHOWN, DEF_SHOWN)                     'Shadow visible
  End With
  
  If Not Ambient.UserMode Then                                        'If we're in design mode
    Exit Sub                                                          'Nothing happens in design mode, outta here
  End If
  
  Call CheckEnvironment                                               'Check OS, color depth etc.
  If m_IsLayered = False Then                                         'If we don't have the support
    Exit Sub                                                          'Outta here
  End If
  
'Ok, we're running not designing, let's set everything up
  nColor = TranslateColor(m_Color)                                    'Translate system color indices
  hWndParent = UserControl.Parent.hWnd                                'Get the parent form window handle
  
  Set scParent = New cSubclass                                        'Create the parent form subclasser
  With scParent
    Call .Subclass(hWndParent, Me)                                    'Subclass the parent form
    Call .AddMsg(WM_WINDOWPOSCHANGED, MSG_BEFORE)                     'Callback on parent position changing
    Call .AddMsg(WM_SIZE, MSG_BEFORE)                                 'Callback on parent restore/min/max
    Call .AddMsg(WM_SIZING, MSG_BEFORE)                               'Callback on parent being re-sized
    Call .AddMsg(WM_MOVING, MSG_BEFORE)                               'Callback on parent being moved
    Call .AddMsg(WM_EXITSIZEMOVE, MSG_BEFORE)                         'Callback on exit from re-sizing or moving
    Call .AddMsg(WM_SHOWWINDOW, MSG_BEFORE)                           'Callback on show for fader creation
    Call .AddMsg(WM_DISPLAYCHANGE, MSG_BEFORE)                        'Callback on display changes
    If m_IsXP Then
      Call .AddMsg(WM_THEMECHANGED, MSG_BEFORE)                       'Callback on theme change
    End If
  End With
  
  Call ShadowCreateWindows
End Sub

'The control has been resized, resize back to it's fixed size
Private Sub UserControl_Resize()
  Call UserControl.Size(480, 480)
End Sub

'The control is terminating
Private Sub UserControl_Terminate()
  Set wnShadow = Nothing
  Set scParent = Nothing
End Sub

'Write the properties to the property bag
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty(PRP_COLOR, m_Color, DEF_COLOR)                'Shadow color
    Call .WriteProperty(PRP_DEPTH, m_Depth, DEF_DEPTH)                'Shadow depth
    Call .WriteProperty(PRP_FADES, m_FadeS, DEF_FADES)                'Form FadeIn
    Call .WriteProperty(PRP_FADET, m_FadeT, DEF_FADET)                'FadeIn time
    Call .WriteProperty(PRP_HIDEM, m_HideM, DEF_HIDEM)                'Shadow hide on move
    Call .WriteProperty(PRP_HIDES, m_HideS, DEF_HIDES)                'Shadow hide on size
    Call .WriteProperty(PRP_SOFTS, m_SoftS, DEF_SOFTS)                'Soft shadow
    Call .WriteProperty(PRP_TRANS, m_Trans, DEF_TRANS)                'Shadow transparency
    Call .WriteProperty(PRP_SHOWN, m_Shown, DEF_SHOWN)                'Shadow visible
  End With
End Sub

'---------------------------------------------------------------------
'PARENT WINDOW SUBCLASSER IMPLEMENTED INTERFACE CALLBACK
Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
Static bHidden  As Boolean                                            'Temporarily hidden (resize/move)
Dim tLayered    As Boolean
  
  Select Case uMsg                                                    'Select the message number
  Case WM_SHOWWINDOW                                                  'Show or Hide
    If lParam = 0 Then                                                'Regular show or hide, not an uncover
      If wParam <> 0 Then                                             'Show
        If m_FadeS Then                                               'Are we set for fade in?
          bFadeIn = True                                              'Yep
          nFadeTime = m_FadeT                                         'Store the fade time to use
          Call FaderCreate                                            'Create the fader window
        End If
      End If
    End If
  
  Case WM_WINDOWPOSCHANGED                                            'Parent form position/size has changed
    Call CopyMemory(wp, ByVal lParam, Len(wp))                        'Copy the WINDOWPOS data
    If Not bHidden Then                                               'If not hidden
      Call ShadowSizePos                                              'Position shadows
    End If
  
  Case WM_SIZE                                                        'Parent form has been minimized/restored/maximized
    If wParam = SIZE_RESTORED Then                                    'If the parent has been restored
      If IsWindowVisible(hWnd) = 1 Then                               'If the parent is visible
        If Not bHidden Then                                           'If we're not re-sizing
          Call ShadowShow(True)                                       'Show shadows
        End If
      End If
    End If
  
  Case WM_MOVING                                                      'The parent form is being moved
    If m_HideM Then                                                   'If we're supposed to hide on move
      If Not bHidden Then                                             'If we're not already hidden
        bHidden = True                                                'Set the hidden flag
        Call ShadowShow(False)                                        'Hide shadows
      End If
    End If
  
  Case WM_SIZING                                                      'The parent form is being re-sized
    If m_HideS Then                                                   'If we're supposed to hide on re-size
      If Not bHidden Then                                             'If we're not already hidden
        bHidden = True                                                'Set the hidden flag
        Call ShadowShow(False)                                        'Hide shadows
      End If
    End If
  
  Case WM_EXITSIZEMOVE                                                'If we've exited from re-size/move
    If bHidden Then                                                   'If we're hidden
      bHidden = False                                                 'Unset the hidden flag
      Call ShadowSizePos                                              'Position shadows
      Call ShadowShow(True)                                           'Show shadows
    End If
    
  Case WM_DISPLAYCHANGE                                               'The display settings have been changed
    tLayered = m_IsLayered                                            'Save the the existing layered status
    Call CheckEnvironment                                             'Check OS, color depth etc.
    If tLayered Then                                                  'If we were layer enabled
      If Not m_IsLayered Then                                         'But not anymore
        Set wnShadow = Nothing                                        'Get rid of the shadow
      End If
    Else
      If m_IsLayered Then                                             'If we're now layered but previously weren't
        Call ShadowCreateWindows                                      'Create the shadow windows
        Call ShadowCreate                                             'Create the shadows
        If m_Shown Then                                               'If shadows are visible
          If IsWindowVisible(hWndParent) Then                         'If the parent is visible
            Call ShadowShow(True, True)                               'Display the shadows
          End If
        End If
      End If
    End If
    
  Case WM_THEMECHANGED                                                'Windows XP theme has changed
    Call CheckEnvironment                                             'Check OS, color depth etc.
    Call ShadowCreate                                                 'Create the shadows
    
  End Select
End Sub

'---------------------------------------------------------------------
'WINDOW INTERFACE CALLBACK
'Fader window callback. NB the shadow windows don't need to callback, all 'painting' is taken care of by UpdateLayeredWindow
Private Sub iWindow_Proc(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
  Select Case uMsg
  Case WM_CREATE
    Call FaderStart(hWnd)                                             'Create the fader window, start the fade process
  Case WM_TIMER
    Call FaderStep(hWnd)                                              'Timer event, step the fade
  End Select
End Sub

'---------------------------------------------------------------------
' WORKER SUBROUTINES

'Create the fader window
Private Sub FaderCreate()
  Set wnFader = New cWindow                                           'Create the shadow window class
  With wnFader
    Set .Owner = Me                                                   'Set the owner of the implemented callback interface (iWindow_WndProc)
'NB - It's important to create a unique window class name hence the '& hWndParent'
    Call .WindowClassRegister("FadeCls" & hWndParent)                 'Define the window class
    Call .AddMsg(WM_CREATE)                                           'Fader window calls back on window create
    Call .AddMsg(WM_TIMER)                                            'Fader window calls back on timer
'Create the fader window, because we're hooked into the WM_CREATE message, execution will proceed
'to iWindow_WndProc (WM_CREATE) which will call FaderStart, all before the WindowCreate function returns
    hWndFader = .WindowCreate(WS_EX_TOPMOST Or WS_EX_TOOLWINDOW Or WS_EX_LAYERED, WS_POPUP)
  End With
End Sub

'Start the fade process
Private Sub FaderStart(hWnd As Long)
Dim rc    As tRECT, _
    ptDst As tPOINT, _
    ptSrc As tPOINT, _
    sz    As tSIZE

  Call GetWindowRect(hWndParent, rc)                                  'Get the window rect of the parent window
    
  With rc
    ptDst.x = .Left                                                   'x location of the fader window
    ptDst.y = .Top                                                    'y location of the fader window
    ptSrc.x = .Left                                                   'x location in the source dc (the screen)
    ptSrc.y = .Top                                                    'y location in the source dc (the screen)
    sz.cx = .Right - .Left                                            'Width of the fader window
    sz.cy = .bottom - .Top                                            'Height of the fader window
    If m_Shown Then                                                   'If the shadows are visible then we should fade them in/out as well
      sz.cx = sz.cx + m_Depth                                         'Bump the width to include the right shadow
      sz.cy = sz.cy + m_Depth                                         'Bump the height to include the bottom shadow
    End If
  End With
  
'Blendfunction for window fade
  With bf
    .AlphaFormat = 0
    .BlendFlags = 0
    .BlendOp = AC_SRC_OVER                                            'Alpha overlay
    .SourceConstantAlpha = 255                                        'Window transparency, fully opaque initialy
  End With
  
'You may be wondering how one routine serves for both fade in and fade out without any
'conditional (if fade in else...) code.
'Fade in...
'   before the parent appears we create the fader window above where the parent will appear
'   the fader window contents are what lay behind it. Then the parent is displayed, though
'   we can't see it because the fader window is top most and initialy fully opaque. As the
'   timer fires we make the fader window more and more transparent allowing the parent to show
'   through. Instead of fading the parent in we're fading out the image of the background over
'   the top of the parent window. Cute, but best of all, this works for fade out as well.
'Fade out...
'   the parent is fully visible, we create the fader window over the top of the parent using
'   the screen image at that location (which is that of the parent) then the parent is hidden
'   just leaving its image in the fader window, which is faded away as the timer fires.
'In summary, the technique works both ways (fade in, fade out) without change.
  Call UpdateLayeredWindow(hWnd, GetDC(hWnd), ptDst, sz, GetDC(0), ptSrc, 0, bf, ULW_ALPHA)
  Call ShowWindow(hWnd, SW_SHOW)
  
  nFaderStart = GetTickCount                                          'Remember when we started the fade
  Call SetTimer(hWnd, hWnd, FADE_STEP, 0)                             'Create the timer
End Sub

'Timer has fired, step the fade
Private Sub FaderStep(hWnd As Long)
Dim nStep     As Long, _
    nAlpha    As Long, _
    nDuration As Long
  
  nDuration = GetTickCount - nFaderStart                              'Calculate the duration
  
  If nDuration < nFadeTime Then                                       'Ensure we don't take ANY longer than requested
    nAlpha = bf.SourceConstantAlpha
    nStep = nAlpha / ((nFadeTime - nDuration) / FADE_STEP)            'For smoothness and time accuracy, continuously recalculate the step each tick
    
    If nStep < 1 Then
      nStep = 1
    End If
    
    If nAlpha > nStep Then
      bf.SourceConstantAlpha = nAlpha - nStep
      
'Update the transparency of the fader window
      Call UpdateLayeredWindow(hWnd, 0, ByVal 0, ByVal 0, 0, ByVal 0, 0, bf, ULW_ALPHA)
'Exit here while we haven't finished fading
      Exit Sub
    End If
  End If
  
'If we're here then the fade window and timer should be killed
  Call KillTimer(hWnd, hWnd)                                          'Destroy the timer
  Set wnFader = Nothing                                               'Destroy the fader window
  
  If bFadeIn Then
    RaiseEvent FadedIn                                                'Raise the Faded in event
  Else
    RaiseEvent FadedOut                                               'Raise the faded out event
    bBlock = False                                                    'If we're blocking on fade out, un-block
  End If
End Sub

Private Sub ShadowCreateWindows()
Const WS_EX As Long = WS_EX_LAYERED Or WS_EX_TRANSPARENT Or WS_EX_NOPARENTNOTIFY

  Set wnShadow = New cWindow                                          'Create the shadow window class
  With wnShadow
    Set .Owner = Me                                                   'Set the owner of the implemented callback interface (iWindow_WndProc)
'NB - It's important to create a unique window class name hence the '& hWndParent'
    Call .WindowClassRegister("ShadowCls" & hWndParent)               'Define the show window class, we need a unique window class else two shadowed forms within the same application would callback into the first created shadow control instance
    hWndRt = .WindowCreate(WS_EX, WS_POPUP, , , , , , , hWndParent)   'Create the right shadow window
    hWndBt = .WindowCreate(WS_EX, WS_POPUP, , , , , , , hWndParent)   'Create the bottom shadow window
  End With
End Sub

'Create the right and bottom shadows
Private Sub ShadowCreate()
  If Not m_IsLayered Then                                             'If the OS doesn't support transparency
    Exit Sub                                                          'Bail
  End If
  
  If IsWindowVisible(hWndParent) = 0 Then                             'If the parent window isn't visible
    Exit Sub                                                          'Bail
  End If
  
  With wp
'Right shadow
    Call ShadowCreateSub(.x + .cx, .y + m_Depth, m_Depth, .cy, True)
'Bottom shadow
    Call ShadowCreateSub(.x + m_Depth, .y + .cy, .cx - m_Depth, m_Depth, False)
  End With
End Sub

'Size/position the shadows
Private Sub ShadowSizePos()
Static cx As Long, _
       cy As Long
       
  With wp
    If .flags And SWP_HIDEWINDOW Then                                 'If the parent form is being hidden
      Call ShadowShow(False)
    Else
      If .cx <> cx Then                                               'If the parent's width has changed
        cx = .cx                                                      'Store the new width
'Parent width change means we need to modify the bottom window
        Call ShadowCreateSub(.x + m_Depth, .y + .cy, .cx - m_Depth, m_Depth, False)
      End If
      
      If .cy <> cy Then                                               'If the parent's height has changed
        cy = .cy                                                      'Store the new height
'Parent height change means we need to modify the right window
        Call ShadowCreateSub(.x + .cx, .y + m_Depth, m_Depth, .cy, True)
      End If
                                                                  
'Position the shadow windows
      Call MoveWindow(hWndRt, .x + .cx, .y + m_Depth, m_Depth, .cy, False)
      Call MoveWindow(hWndBt, .x + m_Depth, .y + .cy, .cx - m_Depth, m_Depth, False)
      
      If (.flags And SWP_SHOWWINDOW) Then
        Call ShadowShow(True)
      End If
    End If
  End With
End Sub

'Show/hide the shadow windows
Private Sub ShadowShow(bShow As Boolean, Optional bForce As Boolean = False)
Static bLastShow As Boolean
  
  If Not bForce Then
'If we're not forcing the issue
    If bLastShow = bShow Then
'No change in visiblility status so bail
      Exit Sub
    End If
  End If
  
  bLastShow = bShow
  
  If bShow Then
    If m_Shown Then
      Call ShowWindow(hWndRt, SW_SHOWNOACTIVATE)
      Call ShowWindow(hWndBt, SW_SHOWNOACTIVATE)
    End If
  Else
    Call ShowWindow(hWndRt, SW_HIDE)
    Call ShowWindow(hWndBt, SW_HIDE)
  End If
End Sub

'Create the content of the indicated shadow window
Private Sub ShadowCreateSub(x As Long, y As Long, cx As Long, cy As Long, Right As Boolean)
Dim dc        As Long, _
    iX        As Long, _
    iY        As Long, _
    hDib      As Long, _
    hWin      As Long, _
    nPixel    As Long, _
    nAlpha    As Long, _
    aPixels() As Long, _
    pBmpBits  As Long, _
    pt0       As tPOINT, _
    pt        As tPOINT, _
    sz        As tSIZE, _
    bs        As tBLENDFUNCTION, _
    bmpHeader As tBITMAPINFOHEADER, _
    SafeArray As tSAFEARRAY2D
  
'Create a screen compatible memory dc
  dc = CreateCompatibleDC(0)

'Initialize a bitmap header
  With bmpHeader
    .biSize = Len(bmpHeader)                                          'Bitmap header size
    .biWidth = cx                                                     'Bitmap/window pixel width
    .biHeight = cy                                                    'Bitmap/window pixel height
    .biPlanes = 1                                                     'Graphics planes
    .biBitCount = 32                                                  '32bits per pixel BGRA (Blue, Green, Red, Alpha)
    .biSizeImage = cx * cy * 4                                        'Memory size, width * height * 32bit
  End With

'Create a device independant bitmap as per the header, compatible with the dc (compatible with the screen)
  hDib = CreateDIBSection(dc, bmpHeader, 0, pBmpBits, 0, 0)
  
'Construct a VB safearray header that matches the specs of the bitmap
  With SafeArray
    .cbElements = 4                                                   '4 bytes per element - 32bits per pixel
    .cDims = 2                                                        'We'll treat the pixels as a two dimensional array
    .pvData = pBmpBits                                                'The data pointer points to the bitmap data (pixels)
'Describes the x dimension
    .Bounds(0).lLbound = 0                                            'Lowest bound will be 0
    .Bounds(0).cElements = cy                                         'The number of elements
'Describes the y dimension
    .Bounds(1).lLbound = 0                                            'Lowest bound will be 0
    .Bounds(1).cElements = cx                                         'The number of elements
  End With
  
'Copy the address of our safearray over the address of aPixels() safearray
  Call CopyMemory(ByVal VarPtrArray(aPixels), VarPtr(SafeArray), 4)
'Now when we access the array aPixels() we're accessing the bitmap pixels directly in memory - COOL!

  If Right Then
    hWin = hWndRt
  Else
    hWin = hWndBt
  End If
  
  If m_SoftS Then
'Soft shadow
    If Right Then
'Right shadow
      For iY = 0 To cy - 1
        If (iY < cx) Then
'Near the bottom right corner
          nAlpha = (255 * iY) \ cx
        ElseIf iY >= (cy - cx) Then
'Near the top right corner
          nAlpha = ((cy - iY) * 255) \ cx
        Else
           nAlpha = 255
        End If
        
        For iX = 0 To cx - 1
          aPixels(iX, iY) = FixColAlpha((nAlpha * (cx - iX)) \ cx)
        Next iX
      Next iY
    Else
'Bottom shadow
      For iX = 0 To cx - 1
        If (iX < cy) Then
'Near the bottom left corner
          nAlpha = (255 * iX) \ cy
        Else
          nAlpha = 255
        End If
        
        For iY = 0 To cy - 1
          aPixels(iX, iY) = FixColAlpha((nAlpha * iY) \ cy)
        Next iY
      Next iX
    End If
  Else
'Hard shadow
    nPixel = FixColAlpha(255)
    For iX = 0 To cx - 1
      For iY = 0 To cy - 1
        aPixels(iX, iY) = nPixel
      Next iY
    Next iX
  End If
  
  If Right Then
    If m_IsLuna Then
'Luna interface, draw the top right curved corner.
      
'Protect against shadow depths less than 5
      On Error Resume Next
        aPixels(cx - 1, cy - 1) = 0
        aPixels(cx - 2, cy - 1) = 0
        aPixels(cx - 3, cy - 1) = 0
        aPixels(cx - 4, cy - 1) = 0
        aPixels(cx - 5, cy - 1) = 0
  
        aPixels(cx - 1, cy - 2) = 0
        aPixels(cx - 2, cy - 2) = 0
        aPixels(cx - 3, cy - 2) = 0
  
        aPixels(cx - 1, cy - 3) = 0
        aPixels(cx - 2, cy - 3) = 0
  
        aPixels(cx - 1, cy - 4) = 0
        aPixels(cx - 1, cy - 5) = 0
      On Error GoTo 0
    End If
  End If

'Clean up the array header else there will be problems
  Call CopyMemory(ByVal VarPtrArray(aPixels), 0&, 4)

'Setup the blend function
  With bs
    .AlphaFormat = WinSubHook2.AC_SRC_ALPHA                            'Use the alpha channel for individual pixel transparency
    .BlendFlags = 0
    .BlendOp = AC_SRC_OVER                                            'Alpha overlay
    .SourceConstantAlpha = m_Trans                                    'Alpha transparency for overall transparency
  End With

'Setup the window position and size data
  pt.x = x
  pt.y = y
  sz.cx = cx
  sz.cy = cy

'Select the bitmap into the memory display context
  hDib = SelectObject(dc, hDib)

'Here we go...
'  hWin - Display this window
'  dc   - Matching this display context
'  pt   - At this position
'  sz   - Of this size
'  dc   - Using this display context for the window's visual content
'  pt0  - Use this as the starting point in the dc for the visual image (0, 0)
'  0    - Using this value as the color key (not used)
'  bs   - Using this blendfunction data to describe how to blend/layer the window with the background
'  flag - Alpha, not opaque nor color keyed
  Call UpdateLayeredWindow(hWin, dc, pt, sz, dc, pt0, 0, bs, ULW_ALPHA)

'Trash the bitmap
  Call SelectObject(dc, hDib)

'Delete the memory display context
  Call DeleteDC(dc)
End Sub

'Draw a filled rectangle, used to draw the control at design time, see UserControl_Paint
Private Sub Rectangle(hDc As Long, x As Long, y As Long, Width As Long, Height As Long, Color As Long)
Dim rc     As tRECT, _
    hBrush As Long
    
  With rc
    .Left = x
    .Top = y
    .Right = x + Width
    .bottom = y + Height
  End With
  
  hBrush = CreateSolidBrush(Color)
  Call FillRect(hDc, rc, hBrush)
  Call DeleteObject(hBrush)
End Sub

'Premultiply the shadow color with the passed alpha value. This is needed to get nice looking colors according to MSDN.
'NB Alpha should range from 0 to 255
Private Function FixColAlpha(Alpha As Long) As Long
Dim fFactor As Double, _
    BGRA    As tBGRA
    
  fFactor = CDbl(Alpha) / 255#                                        'Calculate the factor
  
'Note that nColor is in RGB format, part of this process is to convert to BGRA format
  With BGRA                                                           'Blue, Green, Red, Alpha
    .b = ((nColor And &HFF0000) \ &H10000) * fFactor                  'Factor the blue component
    .g = ((nColor And &HFF00&) \ &H100&) * fFactor                    'Factor the green component
    .r = (nColor And &HFF) * fFactor                                  'Factor the red component
    .a = Alpha                                                        'Store the alpha value
  End With
  
'Copy the BGRA type to long
  Call CopyMemory(FixColAlpha, BGRA, 4)
End Function

'If the passed color is a system color then translate it to it's actual color
Private Function TranslateColor(Color As OLE_COLOR) As OLE_COLOR
  If (Color And COL_SYS_MASK) Then                                    'If the system color bit is set
    TranslateColor = GetSysColor(Color Xor COL_SYS_MASK)              'Get the translated system color
  Else
    TranslateColor = Color                                            'Not a system color
  End If
End Function

'Return whether the OS supports layered windows
Private Sub CheckEnvironment()
  Dim OSV As tOSVERSIONINFO
  
  m_IsLayered = False
  
  With OSV
    .dwOSVersionInfoSize = Len(OSV)                                   'Set the length element
    Call GetVersionEx(OSV)                                            'Fill the type with OS version info
   
    If .dwPlatformId = 2 Then                                         'If it's an NT based OS
      If .dwMajorVersion = 5 Then                                     'If the major version is 5 OS supports transparency
        If .dwMinorVersion > 0 Then
          m_IsXP = True                                               '
          m_IsLuna = bIsLuna                                          'Determine if the Luna theme is active
        End If
        
        m_IsLayered = (GetDeviceCaps(GetDC(0), BITSPIXEL) >= 16)      'Ensure we have enough screen colors
      End If
    End If
  End With
  
'DEVELOPER!! Alpha transparency isn't supported on your platform
  Debug.Assert m_IsLayered
End Sub

'Return whether the Luna theme is loaded
Private Function bIsLuna() As Boolean
  Dim hLib    As Long
  Dim nPos    As Long
  Dim sTheme  As String
  Dim sName   As String
  
'Be sure that the theme dll is present
  hLib = LoadLibrary("uxtheme.dll")
  If hLib <> 0 Then
'Get the theme file name
    sTheme = String$(255, 0)
    Call GetCurrentThemeName(StrPtr(sTheme), Len(sTheme), 0, 0, 0, 0)
    nPos = InStr(1, sTheme, Chr$(0))
    If nPos > 0 Then
'Get the canonical theme name
      sTheme = Left$(sTheme, nPos - 1)
      sName = String$(255, 0)
      Call GetThemeDocumentationProperty(StrPtr(sTheme), StrPtr("ThemeName"), StrPtr(sName), Len(sName))
      nPos = InStr(1, sName, Chr$(0))
      If nPos > 0 Then
'Is it Luna?
        sName = Left$(sName, nPos - 1)
        bIsLuna = (StrComp(sName, "Luna", vbTextCompare) = 0)
      End If
    End If
    
    Call FreeLibrary(hLib)
  End If
End Function
