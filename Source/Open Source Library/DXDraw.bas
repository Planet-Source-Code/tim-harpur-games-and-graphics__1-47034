Attribute VB_Name = "DXDraw"
'***************************************************************************************************************
'
' DirectX VisualBASIC Interface for DirectDraw/Direct3D Support
'                                                     - written by Tim Harpur for Logicon Enterprises
'
' Don't forget to add the appropriate Project->Reference to the DirectX7 library
'
' Version 2.6
'
' ----------- User Licensing Notice -----------
'
' This file and all source code herein is property of Logicon Enterprises. Licensed users of this file
' and its associated library files are authorized to include this file in their VisualBASIC projects, and
' may redistribute the code herein free of any additional licensing fee, so long as no part of this file,
' whether in its original or modified form, is redistributed in uncompiled format.
'
' Whether in its original or modified form, Logicon Enterprises retains ownership of this file.
'
'***************************************************************************************************************

Option Explicit
Option Compare Text
Option Base 0

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private dx_DirectX As New DirectX7

'DXDraw Control variables
Public Const PiByTwo As Single = 1.5707963267949
Public Const Pi As Single = 3.14159265358979
Public Const ThreePiByTwo As Single = 4.71238898038469
Public Const TwoPi As Single = 6.2831853071795

Public Enum BlitterFX
  BFXStretch = 1
  BFXMirrorLeftRight = 2
  BFXMirrorTopBottom = 4
  BFXTargetColour = 8
  BFXTransparent = 16
End Enum

Public Enum FillStyles
  FSSolid = 0
  FSNoFill
  FSHorizontalLine
  FSVerticalLine
  FSUpwardDiagonal
  FSDownwardDiagonal
  FSCross
  FSDiagonalCross
End Enum

Public Enum LineStyles
  LSSolid = 0
  LSDash
  LSDot
  LSDashDot
  LSDashDotDot
  LSNoLine
  LSInsideSolid
End Enum

Private dx_DirectDraw As DirectDraw7

Private dx_DirectDrawEnabled As Boolean, dx_FullScreenMode As Boolean, dx_SystemMemoryOnly As Boolean
Private dx_Direct3DEnabled As Boolean, dx_Width As Long, dx_Height As Long, dx_BitDepth As Long
Private m_ClippingWindow As Object

Public m_ClippingRectangleX As Long, m_ClippingRectangleY As Long
Public m_ClippingRectangleWidth As Long, m_ClippingRectangleHeight As Long

Private dx_DirectDrawPrimarySurface As DirectDrawSurface7
Private dx_DirectDrawPrimaryPalette As DirectDrawPalette
Private dx_DirectDrawPrimaryColourControl As DirectDrawColorControl
Private dx_DirectDrawPrimaryGammaControl As DirectDrawGammaControl
Private dx_DirectDrawBackSurface As DirectDrawSurface7

Private dx_DirectDrawStaticSurface() As DirectDrawSurface7
Private dx_DirectDrawStaticSurfaceDesc() As DDSURFACEDESC2
Private dx_StaticSurfaceWidth() As Long, dx_StaticSurfaceHeight() As Long
Private m_TotalStaticSurfaces As Long, m_StaticSurfaceFileName() As String
Private m_StaticSurfaceValid() As Boolean, m_StaticSurfaceTrans() As Long
Private m_StaticSurfaceIsTexture() As Boolean, m_StaticSurfacePriority() As Long
Private m_StaticSurfacePath As String

Private dx_Direct3D As Direct3D7
Private dx_Direct3DDevice As Direct3DDevice7
Private m_D3DViewPort(0 To 0) As D3DRECT
Private dx_D3DZSurface As DirectDrawSurface7, dx_D3DZDepth As Long
Private m_TexturePath As String

Public D3DTextureEnable As Boolean
Public D3DWorldAmbient_Red As Single, D3DWorldAmbient_Green As Single, D3DWorldAmbient_Blue As Single
Public D3DWorldBrightnessAdjust As Single, D3DWorldIntensity As Single

Public ISOmetricViewScale As Long

Public Type DisplayInfo
  currentMode As Long
  currentWidth As Long
  currentHeight As Long
  currentDepth As Long
  currentZDepth As Long
  
  availableModes As Long
  availableWidth() As Long
  availableHeight() As Long
  availableDepth() As Long
  
  totalDisplayMemory As Long
  availableDisplayMemory As Long
  
  directDrawDriver As String
End Type

'Timing routine in milliseconds
Public Function DelayTillTime(returnTime As Long, Optional maxCarryOver As Long = 0, Optional ByVal useRelativeTime As Boolean = False, Optional ByVal callDoEvents As Boolean = False)
  Dim CarryOver As Long
  
  If useRelativeTime Then returnTime = timeGetTime() + returnTime
  
  Do
    If callDoEvents Then DoEvents 'if enabled let the system process other tasks while idle
    
    'this would be a good place to put a jump to a quick routine that could be called to take
    'advantage of any slack time - perhaps perform some extra computer AI logic
    ' ...
    ' ...
    ' ...
    
    DelayTillTime = timeGetTime()
  Loop While DelayTillTime < returnTime
  
  CarryOver = DelayTillTime - returnTime
  
  If CarryOver > maxCarryOver Then
    DelayTillTime = DelayTillTime - maxCarryOver
  ElseIf CarryOver > 0 Then
    DelayTillTime = DelayTillTime - CarryOver
  End If
End Function

'fills in and returns a type DisplayInfo
Public Function Get_DisplaySettings() As DisplayInfo
  Dim DisplayModesEnum As DirectDrawEnumModes, ddsd2 As DDSURFACEDESC2
  Dim loop1 As Long, m_caps As DDSCAPS2, ddraw As DirectDraw7, ddId As DirectDrawIdentifier
  
  On Error GoTo badMode
  
  With Get_DisplaySettings
    Set ddraw = dx_DirectX.DirectDrawCreate("")
    
    If .availableModes = 0 Then
      Set DisplayModesEnum = ddraw.GetDisplayModesEnum(DDEDM_DEFAULT, ddsd2)
        
      .availableModes = DisplayModesEnum.GetCount()
      
      ReDim .availableWidth(1 To .availableModes)
      ReDim .availableHeight(1 To .availableModes)
      ReDim .availableDepth(1 To .availableModes)
        
      For loop1 = 1 To .availableModes
        DisplayModesEnum.GetItem loop1, ddsd2
        
        .availableWidth(loop1) = ddsd2.lWidth
        .availableHeight(loop1) = ddsd2.lHeight
        .availableDepth(loop1) = ddsd2.ddpfPixelFormat.lRGBBitCount
      Next loop1
      
      Set ddId = ddraw.GetDeviceIdentifier(DDGDI_DEFAULT)
      
      .directDrawDriver = ddId.GetDescription
    End If
    
    If dx_DirectDrawEnabled Then
      .currentWidth = dx_Width
      .currentHeight = dx_Height
      .currentZDepth = dx_D3DZDepth
      
      If dx_FullScreenMode Then
        .currentMode = 2
        .currentDepth = dx_BitDepth
      Else
        .currentMode = 1
        
        ddraw.GetDisplayMode ddsd2
        
        If ddsd2.ddpfPixelFormat.lFlags And DDPF_RGB Then
          .currentDepth = ddsd2.ddpfPixelFormat.lRGBBitCount
        ElseIf ddsd2.ddpfPixelFormat.lFlags And DDPF_YUV Then
          .currentDepth = ddsd2.ddpfPixelFormat.lYUVBitCount
        ElseIf ddsd2.ddpfPixelFormat.lFlags And DDPF_PALETTEINDEXED1 Then
          .currentDepth = 1
        ElseIf ddsd2.ddpfPixelFormat.lFlags And DDPF_PALETTEINDEXED2 Then
          .currentDepth = 2
        ElseIf ddsd2.ddpfPixelFormat.lFlags And DDPF_PALETTEINDEXED4 Then
          .currentDepth = 4
        ElseIf ddsd2.ddpfPixelFormat.lFlags And DDPF_PALETTEINDEXED8 Then
          .currentDepth = 8
        Else
          .currentDepth = 0
        End If
      End If
    Else
      .currentMode = 0
      .currentWidth = 0
      .currentHeight = 0
      .currentDepth = 0
      .currentZDepth = 0
    End If
    
    On Error Resume Next
    
    m_caps.lCaps = DDSCAPS_VIDEOMEMORY
    .totalDisplayMemory = ddraw.GetAvailableTotalMem(m_caps)
    .availableDisplayMemory = ddraw.GetFreeMem(m_caps)
  End With
  
  Exit Function
  
badMode:
  With Get_DisplaySettings
    .currentMode = 0
    .currentWidth = 0
    .currentHeight = 0
    .currentDepth = 0
    .currentZDepth = 0
    
    .availableModes = 0
    
    .totalDisplayMemory = 0
    .availableDisplayMemory = 0
  End With
End Function

'Initialize the direct draw routines for use as a graphics window - make sure that the ClippingWindow object is set for scale mode pixel
Public Sub Init_DXDrawWindow(parentForm As Object, Optional ClippingWindow As Object = Nothing, Optional ByVal NumberOfStaticSurfaces As Long = 0, _
      Optional ByVal useSystemMemory As Boolean = False, Optional ByVal requestedZBufferDepth As Long = -1)
      
  Dim loop1 As Long, loop2 As Long, clearMode As Boolean
  Dim dx_DirectDrawPrimarySurfaceDesc As DDSURFACEDESC2
  Dim dx_DirectDrawBackSurfaceDesc As DDSURFACEDESC2
  Dim dx_DirectDrawPrimaryClipper As DirectDrawClipper
  
  On Error GoTo badInit
  
  CleanUp_DXDraw
  
  dx_SystemMemoryOnly = useSystemMemory
  
  Set m_ClippingWindow = ClippingWindow
  
  Set dx_DirectDraw = dx_DirectX.DirectDrawCreate("")
  
  dx_DirectDraw.SetCooperativeLevel parentForm.hWnd, DDSCL_NORMAL
  
  'initailize the primary surface
  With dx_DirectDrawPrimarySurfaceDesc
    .lFlags = DDSD_CAPS
    
    If dx_SystemMemoryOnly Then
      .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_SYSTEMMEMORY
    Else
      .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End If
  End With
  
  Set dx_DirectDrawPrimarySurface = dx_DirectDraw.CreateSurface(dx_DirectDrawPrimarySurfaceDesc)
  
  If ClippingWindow Is Nothing Then
    dx_Width = Screen.Width / Screen.TwipsPerPixelX
    dx_Height = Screen.Height / Screen.TwipsPerPixelY
    
    clearMode = False
  Else
    'create a full window clipping rectangle for the primary surface
    Set dx_DirectDrawPrimaryClipper = dx_DirectDraw.CreateClipper(0)
    dx_DirectDrawPrimaryClipper.SetHWnd m_ClippingWindow.hWnd
    dx_DirectDrawPrimarySurface.SetClipper dx_DirectDrawPrimaryClipper
    
    Dim t_Rect As RECT
    
    dx_DirectX.GetWindowRect m_ClippingWindow.hWnd, t_Rect
    
    'find actual width/height of render window - note that if window is resized DirectDraw must be refreshed
    With t_Rect
      dx_Width = .Right - .Left
      dx_Height = .Bottom - .Top
    End With
    
    clearMode = True
  End If
  
  'initailize the back surface
  With dx_DirectDrawBackSurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
    
    If dx_SystemMemoryOnly Then .ddsCaps.lCaps = .ddsCaps.lCaps Or DDSCAPS_SYSTEMMEMORY
    
    .lWidth = dx_Width
    .lHeight = dx_Height
  End With
  
  Set dx_DirectDrawBackSurface = dx_DirectDraw.CreateSurface(dx_DirectDrawBackSurfaceDesc)
  
  CommonInit NumberOfStaticSurfaces, requestedZBufferDepth
  
  If clearMode Then Clear_Display 'pre-clear the display window's surface to black (unless the desktop is used)
  
  On Error Resume Next
  
  Dim caps1 As DDCAPS, caps2 As DDCAPS
  dx_DirectDraw.GetCaps caps1, caps2
  
  If caps1.lCaps2 And DDCAPS2_PRIMARYGAMMA Then Set dx_DirectDrawPrimaryGammaControl = dx_DirectDrawPrimarySurface.GetDirectDrawGammaControl
  If caps1.lCaps2 And DDCAPS2_COLORCONTROLPRIMARY Then Set dx_DirectDrawPrimaryColourControl = dx_DirectDrawPrimarySurface.GetDirectDrawColorControl
  
  dx_DirectDrawEnabled = True
  
  Exit Sub
  
badInit:
  dx_DirectDrawEnabled = False
End Sub

'Initialize the direct draw routines for use as a graphics screen
Public Sub Init_DXDrawScreen(parentForm As Object, Optional ByVal PixelWidth As Long = 800, Optional ByVal PixelHeight As Long = 600, _
      Optional ByVal PixelDepth As Long = 16, Optional ByVal NumberOfStaticSurfaces As Long = 0, _
      Optional ByVal useSystemMemory As Boolean = False, Optional ByVal requestedZBufferDepth As Long = -1)
      
  Dim loop1 As Long, loop2 As Long
  Dim dx_DirectDrawPrimarySurfaceDesc As DDSURFACEDESC2
  
  On Error GoTo badInit
  
  CleanUp_DXDraw
  
  dx_SystemMemoryOnly = useSystemMemory
 
  Set dx_DirectDraw = dx_DirectX.DirectDrawCreate("")
  
  dx_DirectDraw.SetCooperativeLevel parentForm.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
  dx_DirectDraw.SetDisplayMode PixelWidth, PixelHeight, PixelDepth, 0, DDSDM_DEFAULT
  
  'initailize the primary surface
  With dx_DirectDrawPrimarySurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_3DDEVICE
    
    If dx_SystemMemoryOnly Then .ddsCaps.lCaps = .ddsCaps.lCaps Or DDSCAPS_SYSTEMMEMORY
    
    .lBackBufferCount = 1
  End With
  
  Set dx_DirectDrawPrimarySurface = dx_DirectDraw.CreateSurface(dx_DirectDrawPrimarySurfaceDesc)
  
  If PixelDepth = 8 Then
    Dim ptemp(0 To 255) As PALETTEENTRY
    
    Set dx_DirectDrawPrimaryPalette = dx_DirectDraw.CreatePalette(DDPCAPS_8BIT Or DDPCAPS_ALLOW256, ptemp)
    
    Reset_DefaultPalette
    
    dx_DirectDrawPrimarySurface.SetPalette dx_DirectDrawPrimaryPalette
  End If
  
  'initailize the back surface
  Dim Caps As DDSCAPS2
  Caps.lCaps = DDSCAPS_BACKBUFFER
  
  Set dx_DirectDrawBackSurface = dx_DirectDrawPrimarySurface.GetAttachedSurface(Caps)
  
  dx_Width = PixelWidth
  dx_Height = PixelHeight
  
  CommonInit NumberOfStaticSurfaces, requestedZBufferDepth
  
  Clear_Display 'pre-clear the display surfaces to black
  SynchronizeBuffers
  
  On Error Resume Next
  
  Dim caps1 As DDCAPS, caps2 As DDCAPS
  dx_DirectDraw.GetCaps caps1, caps2
  
  If caps1.lCaps2 And DDCAPS2_PRIMARYGAMMA Then Set dx_DirectDrawPrimaryGammaControl = dx_DirectDrawPrimarySurface.GetDirectDrawGammaControl
  If caps1.lCaps2 And DDCAPS2_COLORCONTROLPRIMARY Then Set dx_DirectDrawPrimaryColourControl = dx_DirectDrawPrimarySurface.GetDirectDrawColorControl
  
  dx_DirectDrawEnabled = True
  dx_FullScreenMode = True
  
  Exit Sub
  
badInit:
  dx_DirectDrawEnabled = False
  dx_Width = 0
  dx_Height = 0
  dx_BitDepth = 0
End Sub

'more DXDraw initialization - common to either windowed or full screen modes
Private Sub CommonInit(ByVal NumberOfStaticSurfaces As Long, ByVal requestedZBufferDepth As Long)
  Dim loop1 As Long, loop2 As Long, t_Rect As RECT, colourkey As DDCOLORKEY
  Dim pixelCaps As DDPIXELFORMAT
  Dim dx_DirectDrawSurfaceDesc As DDSURFACEDESC2
  
  On Error Resume Next
  
  dx_DirectDrawPrimarySurface.GetPixelFormat pixelCaps
  
  If pixelCaps.lFlags & DDPF_RGB Then
    dx_BitDepth = pixelCaps.lRGBBitCount
    
    If dx_BitDepth = 16 And pixelCaps.lGBitMask = &H3E0 Then dx_BitDepth = 15
  ElseIf pixelCaps.lFlags & DDPF_PALETTEINDEXED8 Then
    dx_BitDepth = 8
  Else
    dx_BitDepth = 0
  End If
  
  're-dim the static surfaces
  ReDim dx_DirectDrawStaticSurface(0 To NumberOfStaticSurfaces)
  ReDim dx_DirectDrawStaticSurfaceDesc(0 To NumberOfStaticSurfaces)
  ReDim dx_StaticSurfaceWidth(0 To NumberOfStaticSurfaces)
  ReDim dx_StaticSurfaceHeight(0 To NumberOfStaticSurfaces)
  ReDim m_StaticSurfaceValid(0 To NumberOfStaticSurfaces)
  ReDim m_StaticSurfaceFileName(0 To NumberOfStaticSurfaces)
  ReDim m_StaticSurfaceTrans(0 To NumberOfStaticSurfaces)
  ReDim m_StaticSurfaceIsTexture(0 To NumberOfStaticSurfaces)
  ReDim m_StaticSurfacePriority(0 To NumberOfStaticSurfaces)
  
  Set dx_DirectDrawStaticSurface(0) = dx_DirectDrawBackSurface
  
  dx_StaticSurfaceWidth(0) = dx_Width
  dx_StaticSurfaceHeight(0) = dx_Height
  m_StaticSurfaceValid(0) = True
  
  m_TotalStaticSurfaces = NumberOfStaticSurfaces
  
  ISOmetricViewScale = 1
  
  Set_SSurfacePath "\"
  Set_TexturePath "\"
  
  Initialize_DX3D requestedZBufferDepth
  
  Reset_ClippingRectangle True
End Sub

'returns TRUE if gamma control is available
Public Function Get_GammaControlAvailable() As Boolean
  If dx_DirectDrawPrimaryGammaControl Is Nothing Then
    Get_GammaControlAvailable = False
  Else
    Get_GammaControlAvailable = True
  End If
End Function

'only if gamma control available
Public Sub Get_GammaRamp(ByRef gammaRamp As DDGAMMARAMP)
  On Error Resume Next
  
  dx_DirectDrawPrimaryGammaControl.GetGammaRamp DDSGR_DEFAULT, gammaRamp
End Sub

'only if gamma control available
Public Sub Set_GammaRamp(ByRef gammaRamp As DDGAMMARAMP)
  On Error Resume Next
  
  dx_DirectDrawPrimaryGammaControl.SetGammaRamp DDSGR_DEFAULT, gammaRamp
End Sub

'only if gamma control available
Public Sub Reset_GammaRamp()
  Dim loop1 As Long, temp1 As Long, gammaRamp As DDGAMMARAMP
  
  On Error Resume Next
  
  For loop1 = 0 To 255
    temp1 = loop1 * 256 + loop1
    
    If temp1 > 32767 Then temp1 = temp1 - 65536
    
    gammaRamp.Red(loop1) = temp1
    gammaRamp.Green(loop1) = temp1
    gammaRamp.Blue(loop1) = temp1
  Next loop1
        
  Set_GammaRamp gammaRamp
End Sub

Public Function Get_ColourControlAvailable() As Boolean
  If dx_DirectDrawPrimaryColourControl Is Nothing Then
    Get_ColourControlAvailable = False
  Else
    Get_ColourControlAvailable = True
  End If
End Function

'only if colour control available
Public Function Get_Brightness() As Long
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_BRIGHTNESS Then
    Get_Brightness = ddcolour.lBrightness
  Else
    Get_Brightness = -1
  End If
End Function

'only if colour control available
Public Function Get_Contrast() As Long
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_CONTRAST Then
    Get_Contrast = ddcolour.lContrast
  Else
    Get_Contrast = -1
  End If
End Function

'only if colour control available
Public Function Get_Gamma() As Long
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_GAMMA Then
    Get_Gamma = ddcolour.lGamma
  Else
    Get_Gamma = -1
  End If
End Function

'only if colour control available
Public Function Get_ColorEnable() As Long
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_COLORENABLE Then
    Get_ColorEnable = ddcolour.lColorEnable
  Else
    Get_ColorEnable = -1
  End If
End Function

'only if colour control available
Public Function Get_Saturation() As Long
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_SATURATION Then
    Get_Saturation = ddcolour.lSaturation
  Else
    Get_Saturation = -1
  End If
End Function

'only if colour control available
Public Function Get_Sharpness() As Long
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_SHARPNESS Then
    Get_Sharpness = ddcolour.lSharpness
  Else
    Get_Sharpness = -1
  End If
End Function

'only if colour control available
Public Function Get_Hue() As Long
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_HUE Then
    Get_Hue = ddcolour.lHue
  Else
    Get_Hue = -181
  End If
End Function

'only if colour control available
Public Sub Set_Gamma(Optional ByVal Gamma As Long = 1)
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_GAMMA
  ddcolour.lGamma = Gamma
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Sub

'only if colour control available
Public Sub Set_ColourEnable(Optional ByVal ColourEnable As Long = 1)
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_COLORENABLE
  ddcolour.lColorEnable = ColourEnable
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Sub

'only if colour control available
Public Sub Set_Brightness(Optional ByVal Brightness As Long = 750)
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_BRIGHTNESS
  ddcolour.lBrightness = Brightness
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Sub

'only if colour control available
Public Sub Set_Contrast(Optional ByVal Contrast As Long = 10000)
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_CONTRAST
  ddcolour.lContrast = Contrast
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Sub

'only if colour control available
Public Sub Set_Saturation(Optional ByVal Saturation As Long = 10000)
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_SATURATION
  ddcolour.lSaturation = Saturation
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Sub

'only if colour control available
Public Sub Set_Sharpness(Optional ByVal Sharpness As Long = 5)
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_SHARPNESS
  ddcolour.lSharpness = Sharpness
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Sub

'only if colour control available
Public Sub Set_Hue(Optional ByVal Hue As Long = 0)
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_HUE
  ddcolour.lHue = Hue
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Sub

'returns the status of the display window - a false requires the display to be initialized or re-initialized
Public Function TestDisplayValid() As Boolean
  On Error GoTo notValid
  
  If dx_DirectDrawEnabled Then
    If dx_DirectDraw.TestCooperativeLevel() = DD_OK Then TestDisplayValid = True
  End If
  
notValid:
End Function

Public Function TestDisplay3DValid() As Boolean
  If TestDisplayValid() Then TestDisplay3DValid = dx_Direct3DEnabled
End Function

'copies the back buffer to the visible display window on the main screen if in windowed mode
'flips render/display buffers if in full screen mode
Public Sub RefreshDisplay(Optional ByVal waitForVB As Boolean = False)
  Dim t_Rect As RECT, s_Rect As RECT, s_Desc As DDSURFACEDESC2

  If dx_FullScreenMode Then 'flip back buffer and display surface
    On Error Resume Next
    
    'this is done to prevent flicker that can occur on the last blit operation when the surfaces are flipped
    'I tried to use GetBlitStatus but for some reason it does not work properly - claims the blitter is
    'done yet the flicker effect still occurs
    dx_DirectDrawBackSurface.Lock t_Rect, s_Desc, DDLOCK_WAIT, 0
    dx_DirectDrawBackSurface.Unlock t_Rect
    
    'wait for vertical blank
    If waitForVB Then
      On Error GoTo noVBfullscreen
      
      Do While dx_DirectDraw.GetVerticalBlankStatus() <> 0
        'DoEvents
      Loop
    End If
    
skipVBfullscreen:
    On Error Resume Next
    
    'flip front and back buffers
    dx_DirectDrawPrimarySurface.Flip Nothing, DDFLIP_WAIT
  Else 'copy back buffer to display surface
    If waitForVB Then 'wait for vertical blank
      On Error GoTo noVBwindow
      
      Do While dx_DirectDraw.GetVerticalBlankStatus() <> 0
        'DoEvents
      Loop
    End If
    
skipVBwindow:
    On Error Resume Next
    
    'copy back buffer to window
    If Not (m_ClippingWindow Is Nothing) Then dx_DirectX.GetWindowRect m_ClippingWindow.hWnd, t_Rect
    
    dx_DirectDrawPrimarySurface.Blt t_Rect, dx_DirectDrawBackSurface, s_Rect, DDBLT_WAIT
  End If
  
  Exit Sub
  
noVBwindow:
  Resume skipVBwindow

noVBfullscreen:
  Resume skipVBfullscreen
End Sub

'sysncronize contents of the rendering buffer with the display buffer if in full screen mode
'this is different from flipping the buffers as this makes both buffers' contents identical by copying the
'contents of the back render buffer into the display buffer - used to prevent graphical glitches during
'things such as "pause modes" where the display is not longer being updated but may still be being
'flipped
Public Sub SynchronizeBuffers()
  Dim t_Rect As RECT, s_Rect As RECT
  
  On Error Resume Next
  
  If Not dx_FullScreenMode Then Exit Sub
  
  dx_DirectDrawPrimarySurface.Blt t_Rect, dx_DirectDrawBackSurface, s_Rect, DDBLT_WAIT
End Sub

'get colour converted from RGB to target bit depth or screen's current bit depth
Public Function Get_ConvertedColour(ByVal rgbColour As Long, Optional ByVal targetBitDepth As Variant)
  Dim compRed As Long, compGreen As Long, compBlue As Long
  Dim palEntry As Long, palMatch As Long, palDelta As Long
  Dim DeltaR As Long, DeltaG As Long, DeltaB As Long, DeltaT As Long
  Dim paletteEntries(0 To 255) As PALETTEENTRY
  
  compRed = (rgbColour And &HFF0000) \ 65536
  compGreen = (rgbColour And 65281) \ 256 'I don't use &HFF00 here as VB will erroneously convert this at run time to &HFFFFFF00
  compBlue = rgbColour And &HFF
  
  If IsMissing(targetBitDepth) Then targetBitDepth = dx_BitDepth
  
  Select Case targetBitDepth
    Case 8
      Get_Palette paletteEntries
      
      palMatch = 0
      palDelta = 999
      
      For palEntry = 0 To 255 'scan palette entries for closest match
        DeltaR = paletteEntries(palEntry).Red - compRed
        DeltaT = DeltaR
        If DeltaR < 0 Then DeltaR = -DeltaR
        
        DeltaG = paletteEntries(palEntry).Green - compGreen
        DeltaT = DeltaT + DeltaG
        If DeltaG < 0 Then DeltaG = -DeltaG
        
        DeltaB = paletteEntries(palEntry).Blue - compBlue
        DeltaT = DeltaT + DeltaB
        If DeltaB < 0 Then DeltaB = -DeltaB
        
        DeltaT = DeltaT + DeltaR + DeltaG + DeltaB
        
        If palDelta > DeltaT Then
          palDelta = DeltaT
          palMatch = palEntry
        End If
      Next palEntry
      
      Get_ConvertedColour = palMatch
    Case 15
      compRed = compRed \ 8
      compGreen = compGreen \ 8
      compBlue = compBlue \ 8
      
      Get_ConvertedColour = ((compRed * 32) Or compGreen) * 32 Or compBlue
    Case 16
      compRed = compRed \ 8
      compGreen = compGreen \ 4
      compBlue = compBlue \ 8
      
      Get_ConvertedColour = ((compRed * 64) Or compGreen) * 32 Or compBlue
    Case 24, 32
      Get_ConvertedColour = rgbColour
  End Select
End Function

'Clear the display to the indicated colour (subject to clipping rectangle)
Public Sub Clear_Display(Optional ByVal clearColour As Long = 0, Optional ByVal clearZbufferOnly As Boolean = False)
  Dim t_Rect As RECT
  
  On Error Resume Next
  
  If clearZbufferOnly = True Then 'only the 3D Zbuffer is to be cleared
    dx_Direct3DDevice.Clear 1, m_D3DViewPort(), D3DCLEAR_ZBUFFER, 0, 1#, 0
  ElseIf dx_D3DZSurface Is Nothing Then 'no 3D Zbuffer being used - just clear using 2D blitter
    With t_Rect
      .Left = m_ClippingRectangleX
      .Top = m_ClippingRectangleY
      .Right = m_ClippingRectangleWidth + .Left
      .Bottom = m_ClippingRectangleHeight + .Top
    End With
    
    dx_DirectDrawBackSurface.BltColorFill t_Rect, clearColour
  Else 'clear both the 3D Zbuffer and the display
    dx_Direct3DDevice.Clear 1, m_D3DViewPort(), D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, clearColour, 1#, 0
  End If
End Sub

'Wash the target area with the RGB colour
Public Sub Wash(ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single, ByVal washLevel As Single)
  Dim tVertex(0 To 3) As D3DTLVERTEX, rgbColour As Long
  Dim redL As Long, greenL As Long, blueL As Long, tOpacity As Long
  
  redL = Red * 255
  If redL > 255 Then redL = 255
  
  greenL = Green * 255
  If greenL > 255 Then greenL = 255
  
  blueL = Blue * 255
  If blueL > 255 Then blueL = 255
  
  tOpacity = washLevel * 255
  If tOpacity > 255 Then tOpacity = 255
  
  If tOpacity > 127 Then
    tOpacity = (&H1000000 * (255 - tOpacity)) Xor &HFF000000
  Else
    tOpacity = &H1000000 * tOpacity
  End If
  
  rgbColour = tOpacity Or (redL * 65536) Or (greenL * 256) Or blueL
  
  tVertex(0).Color = rgbColour
  tVertex(1).Color = rgbColour
  tVertex(2).Color = rgbColour
  tVertex(3).Color = rgbColour
  
  tVertex(0).rhw = 1
  tVertex(1).rhw = 1
  tVertex(2).rhw = 1
  tVertex(3).rhw = 1
  
  tVertex(0).sx = m_ClippingRectangleX
  tVertex(1).sx = m_ClippingRectangleX
  tVertex(2).sx = m_ClippingRectangleX + m_ClippingRectangleWidth
  tVertex(3).sx = m_ClippingRectangleX + m_ClippingRectangleWidth
  
  tVertex(0).sy = m_ClippingRectangleY + m_ClippingRectangleHeight
  tVertex(1).sy = m_ClippingRectangleY
  tVertex(2).sy = m_ClippingRectangleY + m_ClippingRectangleHeight
  tVertex(3).sy = m_ClippingRectangleY
  
  dx_Direct3DDevice.BeginScene
  dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, tVertex(0), 4, D3DDP_DEFAULT
  dx_Direct3DDevice.EndScene
End Sub

'BlitClear the area of display to the background colour (subject to clipping rectangle)
Public Sub BlitClear(ByVal xPos As Long, ByVal yPos As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal clearColour As Long = 0)
  Dim t_Rect As RECT
  
  On Error Resume Next
  
  With t_Rect
    .Left = m_ClippingRectangleX + xPos
    .Top = m_ClippingRectangleY + yPos
    .Right = Width + .Left
    .Bottom = Height + .Top
    
    If .Left < m_ClippingRectangleX Then .Left = m_ClippingRectangleX
    If .Left >= m_ClippingRectangleX + m_ClippingRectangleWidth Then Exit Sub
    If .Right <= m_ClippingRectangleX Then Exit Sub
    If .Right > m_ClippingRectangleX + m_ClippingRectangleWidth Then .Right = m_ClippingRectangleX + m_ClippingRectangleWidth
    
    If .Top < m_ClippingRectangleY Then .Top = m_ClippingRectangleY
    If .Top >= m_ClippingRectangleY + m_ClippingRectangleHeight Then Exit Sub
    If .Bottom <= m_ClippingRectangleY Then Exit Sub
    If .Bottom > m_ClippingRectangleY + m_ClippingRectangleHeight Then .Bottom = m_ClippingRectangleY + m_ClippingRectangleHeight
  End With
  
  dx_DirectDrawBackSurface.BltColorFill t_Rect, clearColour
End Sub

'Blit the source surface to the co-ordinates on the back buffer using transparency copy
Public Sub BlitTransparent(ByVal staticSurfaceIndex As Long, ByVal targetOffsetX As Long, _
        ByVal targetOffsetY As Long, ByVal sourceOffsetX As Long, ByVal sourceOffsetY As Long, ByVal sourceWidth As Long, _
        ByVal sourceHeight As Long)
        
  Dim s_Rect As RECT, t_Rect As RECT
  
  On Error Resume Next
  
  If targetOffsetX >= m_ClippingRectangleWidth Or targetOffsetY >= m_ClippingRectangleHeight Then Exit Sub
  
  With s_Rect
    If targetOffsetX < 0 Then
      .Left = sourceOffsetX - targetOffsetX
      sourceWidth = sourceWidth + targetOffsetX
      targetOffsetX = 0
    Else
      .Left = sourceOffsetX
    End If
    
    If targetOffsetY < 0 Then
      .Top = sourceOffsetY - targetOffsetY
      sourceHeight = sourceHeight + targetOffsetY
      targetOffsetY = 0
    Else
      .Top = sourceOffsetY
    End If
    
    If targetOffsetX + sourceWidth > m_ClippingRectangleWidth Then sourceWidth = m_ClippingRectangleWidth - targetOffsetX
    .Right = .Left + sourceWidth
    
    If targetOffsetY + sourceHeight > m_ClippingRectangleHeight Then sourceHeight = m_ClippingRectangleHeight - targetOffsetY
    .Bottom = .Top + sourceHeight
  End With
  
  If sourceWidth <= 0 Or sourceHeight <= 0 Then Exit Sub
  
  With t_Rect
    .Left = targetOffsetX + m_ClippingRectangleX
    .Top = targetOffsetY + m_ClippingRectangleY
    .Right = .Left + sourceWidth
    .Bottom = .Top + sourceHeight
  End With
  
  dx_DirectDrawBackSurface.Blt t_Rect, dx_DirectDrawStaticSurface(staticSurfaceIndex), s_Rect, DDBLT_WAIT Or DDBLT_KEYSRC
End Sub

'Blit the source static surface to the co-ordinates on the target static surface using transparent copy
Public Sub BlitTransparentS2S(ByVal sourceStaticSurfaceIndex As Long, ByVal targetStaticSurfaceIndex As Long, _
      ByVal sourceOffsetX As Long, ByVal sourceOffsetY As Long, ByVal sourceWidth As Long, _
      ByVal sourceHeight As Long, ByVal targetOffsetX As Long, ByVal targetOffsetY As Long)

  Dim s_Rect As RECT, t_Rect As RECT, m_width As Long, m_height As Long
  
  On Error Resume Next
  
  m_width = dx_StaticSurfaceWidth(targetStaticSurfaceIndex)
  m_height = dx_StaticSurfaceHeight(targetStaticSurfaceIndex)
  
  If targetOffsetX >= m_width Or targetOffsetY >= m_height Then Exit Sub
  
  With s_Rect
    If targetOffsetX < 0 Then
      .Left = sourceOffsetX - targetOffsetX
      sourceWidth = sourceWidth + targetOffsetX
      targetOffsetX = 0
    Else
      .Left = sourceOffsetX
    End If
   
    If targetOffsetY < 0 Then
      .Top = sourceOffsetY - targetOffsetY
      sourceHeight = sourceHeight + targetOffsetY
      targetOffsetY = 0
    Else
      .Top = sourceOffsetY
    End If
    
    If targetOffsetX + sourceWidth > m_width Then sourceWidth = m_width - targetOffsetX
    .Right = .Left + sourceWidth
    
    If targetOffsetY + sourceHeight > m_height Then sourceHeight = m_height - targetOffsetY
    .Bottom = .Top + sourceHeight
  End With
  
  If sourceWidth <= 0 Or sourceHeight <= 0 Then Exit Sub
  
  With t_Rect
    .Left = targetOffsetX
    .Top = targetOffsetY
    .Right = .Left + sourceWidth
    .Bottom = .Top + sourceHeight
  End With
  
  dx_DirectDrawStaticSurface(targetStaticSurfaceIndex).Blt t_Rect, dx_DirectDrawStaticSurface(sourceStaticSurfaceIndex), s_Rect, DDBLT_WAIT Or DDBLT_KEYSRC
End Sub

'Blit the source surface to the co-ordinates on the back buffer using solid copy
Public Sub BlitSolid(ByVal staticSurfaceIndex As Long, ByVal targetOffsetX As Long, ByVal targetOffsetY As Long, _
        ByVal sourceOffsetX As Long, ByVal sourceOffsetY As Long, ByVal sourceWidth As Long, ByVal sourceHeight As Long)
        
  Dim s_Rect As RECT, t_Rect As RECT
  
  On Error Resume Next
  
  If targetOffsetX >= m_ClippingRectangleWidth Or targetOffsetY >= m_ClippingRectangleHeight Then Exit Sub
  
  With s_Rect
    If targetOffsetX < 0 Then
      .Left = sourceOffsetX - targetOffsetX
      sourceWidth = sourceWidth + targetOffsetX
      targetOffsetX = 0
    Else
      .Left = sourceOffsetX
    End If
    
    If targetOffsetY < 0 Then
      .Top = sourceOffsetY - targetOffsetY
      sourceHeight = sourceHeight + targetOffsetY
      targetOffsetY = 0
    Else
      .Top = sourceOffsetY
    End If
    
    If targetOffsetX + sourceWidth > m_ClippingRectangleWidth Then sourceWidth = m_ClippingRectangleWidth - targetOffsetX
    .Right = .Left + sourceWidth
    
    If targetOffsetY + sourceHeight > m_ClippingRectangleHeight Then sourceHeight = m_ClippingRectangleHeight - targetOffsetY
    .Bottom = .Top + sourceHeight
  End With
  
  If sourceWidth <= 0 Or sourceHeight <= 0 Then Exit Sub
  
  With t_Rect
    .Left = targetOffsetX + m_ClippingRectangleX
    .Top = targetOffsetY + m_ClippingRectangleY
    .Right = .Left + sourceWidth
    .Bottom = .Top + sourceHeight
  End With
  
  dx_DirectDrawBackSurface.Blt t_Rect, dx_DirectDrawStaticSurface(staticSurfaceIndex), s_Rect, DDBLT_WAIT
End Sub

'Blit the source static surface to the co-ordinates on the target static surface using solid copy
Public Sub BlitSolidS2S(ByVal sourceStaticSurfaceIndex As Long, ByVal targetStaticSurfaceIndex As Long, _
      ByVal sourceOffsetX As Long, ByVal sourceOffsetY As Long, ByVal sourceWidth As Long, _
      ByVal sourceHeight As Long, ByVal targetOffsetX As Long, ByVal targetOffsetY As Long)

  Dim s_Rect As RECT, t_Rect As RECT, m_width As Long, m_height As Long
  
  On Error Resume Next
  
  m_width = dx_StaticSurfaceWidth(targetStaticSurfaceIndex)
  m_height = dx_StaticSurfaceHeight(targetStaticSurfaceIndex)
  
  If targetOffsetX >= m_width Or targetOffsetY >= m_height Then Exit Sub
  
  With s_Rect
    If targetOffsetX < 0 Then
      .Left = sourceOffsetX - targetOffsetX
      sourceWidth = sourceWidth + targetOffsetX
      targetOffsetX = 0
    Else
      .Left = sourceOffsetX
    End If
   
    If targetOffsetY < 0 Then
      .Top = sourceOffsetY - targetOffsetY
      sourceHeight = sourceHeight + targetOffsetY
      targetOffsetY = 0
    Else
      .Top = sourceOffsetY
    End If
    
    If targetOffsetX + sourceWidth > m_width Then sourceWidth = m_width - targetOffsetX
    .Right = .Left + sourceWidth
    
    If targetOffsetY + sourceHeight > m_height Then sourceHeight = m_height - targetOffsetY
    .Bottom = .Top + sourceHeight
    
    If sourceWidth <= 0 Or sourceHeight <= 0 Then Exit Sub
  End With
  
  With t_Rect
    .Left = targetOffsetX
    .Top = targetOffsetY
    .Right = .Left + sourceWidth
    .Bottom = .Top + sourceHeight
  End With
  
  dx_DirectDrawStaticSurface(targetStaticSurfaceIndex).Blt t_Rect, dx_DirectDrawStaticSurface(sourceStaticSurfaceIndex), s_Rect, DDBLT_WAIT
End Sub

'Blit the source static surface to the co-ordinates on the back buffer using transparency and SpecialFX
Public Sub BlitFX(ByVal staticSurfaceIndex As Long, ByVal targetOffsetX As Long, ByVal targetOffsetY As Long, _
        ByVal sourceOffsetX As Long, ByVal sourceOffsetY As Long, ByVal sourceWidth As Long, ByVal sourceHeight As Long, _
        ByVal specialFX As BlitterFX, Optional ByVal scaleWidth As Double = 1, _
        Optional ByVal scaleHeight As Double = 1, Optional ByVal targetColour As Long = 0)
        
  Dim s_Rect As RECT, t_Rect As RECT
  Dim t_flags As Long, t_FX As DDBLTFX
  
  On Error Resume Next
  
  If targetOffsetX >= m_ClippingRectangleWidth Or targetOffsetY >= m_ClippingRectangleHeight Then Exit Sub
  
  With s_Rect
    t_flags = DDBLT_WAIT
    
    If specialFX And BFXStretch Then
      If targetOffsetX < 0 Then
        If specialFX And BFXMirrorLeftRight Then
          .Left = sourceOffsetX
        Else
          .Left = sourceOffsetX - targetOffsetX / scaleWidth
        End If
        
        sourceWidth = sourceWidth + targetOffsetX / scaleWidth
        targetOffsetX = 0
      Else
        .Left = sourceOffsetX
      End If
      
      If targetOffsetY < 0 Then
        If specialFX And BFXMirrorTopBottom Then
          .Top = sourceOffsetY
        Else
          .Top = sourceOffsetY - targetOffsetY / scaleHeight
        End If
        
        sourceHeight = sourceHeight + targetOffsetY / scaleHeight
        targetOffsetY = 0
      Else
        .Top = sourceOffsetY
      End If
      
      If targetOffsetX + sourceWidth * scaleWidth > m_ClippingRectangleWidth Then
        If specialFX And BFXMirrorLeftRight Then
          .Left = .Left + sourceWidth
          sourceWidth = (dx_Width - targetOffsetX) / scaleWidth
          .Left = .Left - sourceWidth
        Else
          sourceWidth = (m_ClippingRectangleWidth - targetOffsetX) / scaleWidth
        End If
      End If
      
      .Right = .Left + sourceWidth
      
      If targetOffsetY + sourceHeight * scaleHeight > m_ClippingRectangleHeight Then
        If specialFX And BFXMirrorTopBottom Then
          .Top = .Top + sourceHeight
          sourceHeight = (dx_Height - targetOffsetY) / scaleHeight
          .Top = .Top - sourceHeight
        Else
          sourceHeight = (m_ClippingRectangleHeight - targetOffsetY) / scaleHeight
        End If
      End If
      
      .Bottom = .Top + sourceHeight
      
      With t_Rect
        .Top = targetOffsetY + m_ClippingRectangleY
        .Left = targetOffsetX + m_ClippingRectangleX
        .Bottom = .Top + sourceHeight * scaleHeight
        .Right = .Left + sourceWidth * scaleWidth
        
        If .Right > m_ClippingRectangleWidth Then .Right = m_ClippingRectangleWidth
        If .Bottom > m_ClippingRectangleHeight Then .Bottom = m_ClippingRectangleHeight
      End With
    Else
      If targetOffsetX < 0 Then
        If specialFX And BFXMirrorLeftRight Then
          .Left = sourceOffsetX
        Else
          .Left = sourceOffsetX - targetOffsetX
        End If
        
        sourceWidth = sourceWidth + targetOffsetX
        targetOffsetX = 0
      Else
        .Left = sourceOffsetX
      End If
      
      If targetOffsetY < 0 Then
        If specialFX And BFXMirrorTopBottom Then
          .Top = sourceOffsetY
        Else
          .Top = sourceOffsetY - targetOffsetY
        End If
        
        sourceHeight = sourceHeight + targetOffsetY
        targetOffsetY = 0
      Else
        .Top = sourceOffsetY
      End If
      
      If targetOffsetX + sourceWidth > m_ClippingRectangleWidth Then
        If specialFX And BFXMirrorLeftRight Then
          .Left = .Left + sourceWidth
          sourceWidth = m_ClippingRectangleWidth - targetOffsetX
          .Left = .Left - sourceWidth
        Else
          sourceWidth = m_ClippingRectangleWidth - targetOffsetX
        End If
      End If
      
      .Right = .Left + sourceWidth
      
      If targetOffsetY + sourceHeight > m_ClippingRectangleHeight Then
        If specialFX And BFXMirrorTopBottom Then
          .Top = .Top + sourceHeight
          sourceHeight = m_ClippingRectangleHeight - targetOffsetY
          .Top = .Top - sourceHeight
        Else
          sourceHeight = m_ClippingRectangleHeight - targetOffsetY
        End If
      End If
      
      .Bottom = .Top + sourceHeight
      
      With t_Rect
        .Top = targetOffsetY + m_ClippingRectangleY
        .Left = targetOffsetX + m_ClippingRectangleX
        .Bottom = .Top + sourceHeight
        .Right = .Left + sourceWidth
      End With
    End If
  End With
  
  If sourceWidth <= 0 Or sourceHeight <= 0 Then Exit Sub
    
  With t_FX
    If specialFX And BFXTransparent Then t_flags = t_flags Or DDBLT_KEYSRC
      
    If specialFX And BFXTargetColour Then
      t_flags = t_flags Or DDBLT_KEYDESTOVERRIDE
      
      .ddckDestColorKey_low = targetColour
      .ddckDestColorKey_high = targetColour
    End If
    
    If specialFX And BFXMirrorLeftRight Then
      .lDDFX = .lDDFX Or DDBLTFX_MIRRORLEFTRIGHT
      t_flags = t_flags Or DDBLT_DDFX
    End If
    
    If specialFX And BFXMirrorTopBottom Then
      .lDDFX = .lDDFX Or DDBLTFX_MIRRORUPDOWN
      t_flags = t_flags Or DDBLT_DDFX
    End If
  End With
  
  dx_DirectDrawBackSurface.BltFx t_Rect, dx_DirectDrawStaticSurface(staticSurfaceIndex), s_Rect, t_flags, t_FX
End Sub

'Blit the source static (texture) surface to the co-ordinates on the back buffer using transparency, translucency and scaling
Public Sub BlitTexture(ByVal staticSurfaceIndex As Long, _
        ByVal targetOffsetX As Long, ByVal targetOffsetY As Long, ByVal targetWidth As Long, ByVal targetHeight As Long, _
        ByVal sourceULeft As Single, ByVal sourceURight As Single, ByVal sourceVTop As Single, ByVal sourceVBottom As Single, _
        ByVal baseColour As Long)
        
  Dim bVertex(0 To 3) As D3DTLVERTEX
  
  On Error Resume Next
  
  If staticSurfaceIndex > 0 Then
    dx_Direct3DDevice.SetTexture 0, DXDraw.GetDirectDrawSurface(staticSurfaceIndex)
  Else
    dx_Direct3DDevice.SetTexture 0, Nothing
  End If
  
  bVertex(1).Color = baseColour
  bVertex(1).rhw = 1
  bVertex(1).sx = m_ClippingRectangleX + targetOffsetX
  bVertex(1).sy = m_ClippingRectangleY + targetOffsetY
  bVertex(1).tU = sourceULeft
  bVertex(1).tV = sourceVTop
  
  bVertex(0).Color = baseColour
  bVertex(0).rhw = 1
  bVertex(0).sx = bVertex(1).sx
  bVertex(0).sy = bVertex(1).sy + targetHeight
  bVertex(0).tU = sourceULeft
  bVertex(0).tV = sourceVBottom
  
  bVertex(2).Color = baseColour
  bVertex(2).rhw = 1
  bVertex(2).sx = bVertex(1).sx + targetWidth
  bVertex(2).sy = bVertex(0).sy
  bVertex(2).tU = sourceURight
  bVertex(2).tV = sourceVBottom
  
  bVertex(3).Color = baseColour
  bVertex(3).rhw = 1
  bVertex(3).sx = bVertex(2).sx
  bVertex(3).sy = bVertex(1).sy
  bVertex(3).tU = sourceURight
  bVertex(3).tV = sourceVTop
  
  dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, bVertex(0), 4, D3DDP_DEFAULT
End Sub

'Blit the source static (texture) surface to the co-ordinates on the back buffer using transparency, translucency, scaling, and rotation
Public Sub BlitTextureRotated(ByVal staticSurfaceIndex As Long, ByVal imageRotation As Single, _
        ByVal targetOffsetX As Long, ByVal targetOffsetY As Long, ByVal targetWidth As Long, ByVal targetHeight As Long, _
        ByVal sourceULeft As Single, ByVal sourceURight As Single, ByVal sourceVTop As Single, ByVal sourceVBottom As Single, _
        ByVal baseColour As Long, Optional ByVal useTargetCentre As Boolean = False)
        
  Dim bVertex(0 To 3) As D3DTLVERTEX, baseX As Long, baseY As Long, cAngle As Single, sAngle As Single
  Dim cWidth As Single, cHeight As Single, sWidth As Single, sHeight As Single
  
  On Error Resume Next
  
  If staticSurfaceIndex > 0 Then
    dx_Direct3DDevice.SetTexture 0, DXDraw.GetDirectDrawSurface(staticSurfaceIndex)
  Else
    dx_Direct3DDevice.SetTexture 0, Nothing
  End If
  
  baseX = m_ClippingRectangleX + targetOffsetX
  baseY = m_ClippingRectangleY + targetOffsetY
  cAngle = Cos(imageRotation)
  sAngle = Sin(imageRotation)
  
  bVertex(1).Color = baseColour
  bVertex(1).rhw = 1
  bVertex(1).tU = sourceULeft
  bVertex(1).tV = sourceVTop
  
  bVertex(0).Color = baseColour
  bVertex(0).rhw = 1
  bVertex(0).tU = sourceULeft
  bVertex(0).tV = sourceVBottom
  
  bVertex(2).Color = baseColour
  bVertex(2).rhw = 1
  bVertex(2).tU = sourceURight
  bVertex(2).tV = sourceVBottom
  
  bVertex(3).Color = baseColour
  bVertex(3).rhw = 1
  bVertex(3).tU = sourceURight
  bVertex(3).tV = sourceVTop
  
  If useTargetCentre Then
    cWidth = cAngle * (targetWidth / 2)
    cHeight = cAngle * (targetHeight / 2)
    sWidth = sAngle * (targetWidth / 2)
    sHeight = sAngle * (targetHeight / 2)
    
    bVertex(1).sx = baseX - cWidth - sHeight
    bVertex(1).sy = baseY - cHeight + sWidth
    
    bVertex(0).sx = baseX - cWidth + sHeight
    bVertex(0).sy = baseY + cHeight + sWidth
    
    bVertex(2).sx = baseX + cWidth + sHeight
    bVertex(2).sy = baseY + cHeight - sWidth
    
    bVertex(3).sx = baseX + cWidth - sHeight
    bVertex(3).sy = baseY - cHeight - sWidth
  Else
    cWidth = cAngle * targetWidth
    cHeight = cAngle * targetHeight
    sWidth = sAngle * targetWidth
    sHeight = sAngle * targetHeight
    
    bVertex(1).sx = baseX
    bVertex(1).sy = baseY
    
    bVertex(0).sx = baseX + sHeight
    bVertex(0).sy = baseY + cHeight
    
    bVertex(2).sx = baseX + cWidth + sHeight
    bVertex(2).sy = baseY + cHeight - sWidth
    
    bVertex(3).sx = baseX + cWidth
    bVertex(3).sy = baseY - sWidth
  End If
  
  dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, bVertex(0), 4, D3DDP_DEFAULT
End Sub

Public Function Get_TotalStaticSurfaces() As Long
  Get_TotalStaticSurfaces = m_TotalStaticSurfaces
End Function

'initialize a static surface surface from a bitmap file (if no file is specified then a blank surface will be created)
'note: set transparentColour = -1 for first pixel transparency, or RGB
Public Function Init_StaticSurface(ByVal staticSurfaceIndex As Long, ByVal surfaceFileName As String, ByVal surfaceWidth As Long, ByVal surfaceHeight As Long, Optional ByVal transparentColour As Long = -1, Optional ByVal forcedBitDepth As Long = 0) As Boolean
  Dim colourkey As DDCOLORKEY, t_Rect As RECT
  
  On Error GoTo badLoad
  
  If staticSurfaceIndex < 1 Or staticSurfaceIndex > m_TotalStaticSurfaces Then Exit Function
  
  Release_StaticSurface staticSurfaceIndex
  
  With dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex)
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
    If dx_SystemMemoryOnly Then .ddsCaps.lCaps = .ddsCaps.lCaps Or DDSCAPS_SYSTEMMEMORY
    
    .lWidth = surfaceWidth
    .lHeight = surfaceHeight
    
    If forcedBitDepth = 16 Then
      .lFlags = .lFlags Or DDSD_PIXELFORMAT
      
      With .ddpfPixelFormat
        .lRGBAlphaBitMask = &HFFFFFF
        .lRBitMask = &H7C00
        .lGBitMask = &H3E0&
        .lBBitMask = &H1F&
        
        .lRGBBitCount = 16
        .lFlags = DDPF_RGB
      End With
    ElseIf forcedBitDepth = 32 Then
      .lFlags = .lFlags Or DDSD_PIXELFORMAT
      
      With .ddpfPixelFormat
        .lRGBAlphaBitMask = &HFFFFFF
        .lRBitMask = &HFF0000
        .lGBitMask = &HFF00&
        .lBBitMask = &HFF&
        
        .lRGBBitCount = 32
        .lFlags = DDPF_RGB
      End With
    End If
  End With
  
  m_StaticSurfaceIsTexture(staticSurfaceIndex) = False
  
  dx_StaticSurfaceWidth(staticSurfaceIndex) = surfaceWidth
  dx_StaticSurfaceHeight(staticSurfaceIndex) = surfaceHeight
  
  m_StaticSurfaceTrans(staticSurfaceIndex) = transparentColour
  
  If surfaceFileName = "" Then
    Set dx_DirectDrawStaticSurface(staticSurfaceIndex) = dx_DirectDraw.CreateSurface(dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex))
    
    dx_DirectDrawStaticSurface(staticSurfaceIndex).BltColorFill t_Rect, 0
    
    If transparentColour = -1 Then transparentColour = 0
  Else
    If m_StaticSurfacePath = "Resource" Then
      Set dx_DirectDrawStaticSurface(staticSurfaceIndex) = dx_DirectDraw.CreateSurfaceFromResource("", surfaceFileName, dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex))
    ElseIf Right$(surfaceFileName, 4) <> ".BMP" Then 'direct draw routines only certain for .BMP format
      Dim srcDC As Long, trgDC As Long, srcPicture As StdPicture
      
      Set dx_DirectDrawStaticSurface(staticSurfaceIndex) = dx_DirectDraw.CreateSurface(dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex))
      
      Set srcPicture = LoadPicture(m_StaticSurfacePath & surfaceFileName)
      
      srcDC = CreateCompatibleDC(ByVal 0&)
      SelectObject srcDC, srcPicture.Handle
      trgDC = dx_DirectDrawStaticSurface(staticSurfaceIndex).GetDC
      
      BitBlt trgDC, 0, 0, surfaceWidth, surfaceHeight, srcDC, 0, 0, vbSrcCopy
      
      dx_DirectDrawStaticSurface(staticSurfaceIndex).ReleaseDC trgDC
      
      DeleteDC srcDC
      
      Set srcPicture = Nothing
    Else
      Set dx_DirectDrawStaticSurface(staticSurfaceIndex) = dx_DirectDraw.CreateSurfaceFromFile(m_StaticSurfacePath & surfaceFileName, dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex))
    End If
    
    If transparentColour = -1 Then
      dx_DirectDrawStaticSurface(staticSurfaceIndex).Lock t_Rect, dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex), 0, 0
      
      transparentColour = dx_DirectDrawStaticSurface(staticSurfaceIndex).GetLockedPixel(0, 0)
      
      dx_DirectDrawStaticSurface(staticSurfaceIndex).Unlock t_Rect
    End If
  End If
  
  colourkey.low = transparentColour
  colourkey.high = transparentColour
  
  dx_DirectDrawStaticSurface(staticSurfaceIndex).SetColorKey DDCKEY_SRCBLT, colourkey
  
  m_StaticSurfaceFileName(staticSurfaceIndex) = surfaceFileName
  m_StaticSurfaceValid(staticSurfaceIndex) = True
  
  Init_StaticSurface = True
  
  Exit Function
  
badLoad:
  Release_StaticSurface staticSurfaceIndex
End Function

'initialize a static surface surface from the current Windows Desktop
Public Function Init_StaticSurfaceFromScreen(ByVal staticSurfaceIndex As Long, Optional ByVal transparentColour As Long = -1, Optional ByVal forcedBitDepth As Long = 0) As Boolean
  Dim colourkey As DDCOLORKEY, surfaceWidth As Long, surfaceHeight As Long
  Dim t_Rect As RECT, s_Rect As RECT
  
  On Error GoTo badLoad
  
  If staticSurfaceIndex < 1 Or staticSurfaceIndex > m_TotalStaticSurfaces Then Exit Function
  
  Release_StaticSurface staticSurfaceIndex
  
  If dx_FullScreenMode Then
    surfaceWidth = dx_Width
    surfaceHeight = dx_Height
  Else
    surfaceWidth = Screen.Width / Screen.TwipsPerPixelX
    surfaceHeight = Screen.Height / Screen.TwipsPerPixelY
  End If
  
  With dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex)
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
    If dx_SystemMemoryOnly Then .ddsCaps.lCaps = .ddsCaps.lCaps Or DDSCAPS_SYSTEMMEMORY
    
    .lWidth = surfaceWidth
    .lHeight = surfaceHeight
    
    If forcedBitDepth = 16 Then
      .lFlags = .lFlags Or DDSD_PIXELFORMAT
      
      With .ddpfPixelFormat
        .lRGBAlphaBitMask = &HFFFFFF
        .lRBitMask = &H7C00
        .lGBitMask = &H3E0&
        .lBBitMask = &H1F&
        
        .lRGBBitCount = 16
        .lFlags = DDPF_RGB
      End With
    ElseIf forcedBitDepth = 32 Then
      .lFlags = .lFlags Or DDSD_PIXELFORMAT
      
      With .ddpfPixelFormat
        .lRGBAlphaBitMask = &HFFFFFF
        .lRBitMask = &HFF0000
        .lGBitMask = &HFF00&
        .lBBitMask = &HFF&
        
        .lRGBBitCount = 32
        .lFlags = DDPF_RGB
      End With
    End If
  End With
  
  m_StaticSurfaceIsTexture(staticSurfaceIndex) = False
  
  dx_StaticSurfaceWidth(staticSurfaceIndex) = surfaceWidth
  dx_StaticSurfaceHeight(staticSurfaceIndex) = surfaceHeight
  
  m_StaticSurfaceTrans(staticSurfaceIndex) = transparentColour
  
  Set dx_DirectDrawStaticSurface(staticSurfaceIndex) = dx_DirectDraw.CreateSurface(dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex))
  
  dx_DirectDrawStaticSurface(staticSurfaceIndex).Blt t_Rect, dx_DirectDrawPrimarySurface, s_Rect, DDBLT_WAIT
  
  If transparentColour = -1 Then
    dx_DirectDrawStaticSurface(staticSurfaceIndex).Lock t_Rect, dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex), 0, 0
    
    transparentColour = dx_DirectDrawStaticSurface(staticSurfaceIndex).GetLockedPixel(0, 0)
    
    dx_DirectDrawStaticSurface(staticSurfaceIndex).Unlock t_Rect
  End If
  
  colourkey.low = transparentColour
  colourkey.high = transparentColour
  
  dx_DirectDrawStaticSurface(staticSurfaceIndex).SetColorKey DDCKEY_SRCBLT, colourkey
  
  m_StaticSurfaceFileName(staticSurfaceIndex) = "Screen Shot"
  m_StaticSurfaceValid(staticSurfaceIndex) = True
  
  Init_StaticSurfaceFromScreen = True
  
  Exit Function
  
badLoad:
  Release_StaticSurface staticSurfaceIndex
End Function

'initialize a static surface surface from a bitmap file for use as a texture (if no file is specified then a blank surface will be created) - (part of the 3D library)
'note: set transparentColour = -2 for no transparency, -1 for first pixel transparency, or RGB
Public Function Init_TextureSurface(ByVal staticSurfaceIndex As Long, ByVal surfaceFileName As String, _
          ByVal surfaceWidth As Long, ByVal surfaceHeight As Long, Optional ByVal transparentColour As Long = -2, _
          Optional ByVal texturePriority As Long = 0, Optional ByVal forcedBitDepth As Long = 0) As Boolean
          
  Dim colourkey As DDCOLORKEY, t_Rect As RECT
  
  On Error GoTo badLoad
  
  If staticSurfaceIndex < 1 Or staticSurfaceIndex > m_TotalStaticSurfaces Then Exit Function
  
  Release_StaticSurface staticSurfaceIndex
  
  ' Enumerate the available texture formats and find a device-supported format
  With dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex)
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_TEXTURESTAGE
    
    .ddsCaps.lCaps = DDSCAPS_TEXTURE
    
    'use the texture manager for hardware accelerated 3D
    If dx_SystemMemoryOnly Then
      .ddsCaps.lCaps = .ddsCaps.lCaps Or DDSCAPS_SYSTEMMEMORY
    Else
      .ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
    End If
    
    .lTextureStage = 0
    
    .lWidth = surfaceWidth
    .lHeight = surfaceHeight
    
    If forcedBitDepth = 16 Then
      .lFlags = .lFlags Or DDSD_PIXELFORMAT
      
      With .ddpfPixelFormat
        .lRGBAlphaBitMask = &HFFFFFF
        .lRBitMask = &H7C00
        .lGBitMask = &H3E0&
        .lBBitMask = &H1F&
        
        .lRGBBitCount = 16
        .lFlags = DDPF_RGB
      End With
    ElseIf forcedBitDepth = 32 Then
      .lFlags = .lFlags Or DDSD_PIXELFORMAT
      
      With .ddpfPixelFormat
        .lRGBAlphaBitMask = &HFFFFFF
        .lRBitMask = &HFF0000
        .lGBitMask = &HFF00&
        .lBBitMask = &HFF&
        
        .lRGBBitCount = 32
        .lFlags = DDPF_RGB
      End With
    End If
  End With
  
  m_StaticSurfaceIsTexture(staticSurfaceIndex) = True
  
  dx_StaticSurfaceWidth(staticSurfaceIndex) = surfaceWidth
  dx_StaticSurfaceHeight(staticSurfaceIndex) = surfaceHeight
  
  m_StaticSurfaceTrans(staticSurfaceIndex) = transparentColour
  m_StaticSurfacePriority(staticSurfaceIndex) = texturePriority
  
  If surfaceFileName = "" Then
    Set dx_DirectDrawStaticSurface(staticSurfaceIndex) = dx_DirectDraw.CreateSurface(dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex))
    
    dx_DirectDrawStaticSurface(staticSurfaceIndex).BltColorFill t_Rect, 0
    
    If transparentColour = -1 Then transparentColour = 0
  ElseIf m_TexturePath = "Resource" Then
    Set dx_DirectDrawStaticSurface(staticSurfaceIndex) = dx_DirectDraw.CreateSurfaceFromResource("", surfaceFileName, dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex))
  ElseIf Right$(surfaceFileName, 4) <> ".BMP" Then  'direct draw routines only certain for .BMP format
    Dim srcDC As Long, trgDC As Long, srcPicture As StdPicture
    
    Set dx_DirectDrawStaticSurface(staticSurfaceIndex) = dx_DirectDraw.CreateSurface(dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex))
    
    Set srcPicture = LoadPicture(m_TexturePath & surfaceFileName)
    
    srcDC = CreateCompatibleDC(ByVal 0&)
    SelectObject srcDC, srcPicture.Handle
    trgDC = dx_DirectDrawStaticSurface(staticSurfaceIndex).GetDC
    
    BitBlt trgDC, 0, 0, surfaceWidth, surfaceHeight, srcDC, 0, 0, vbSrcCopy
    
    dx_DirectDrawStaticSurface(staticSurfaceIndex).ReleaseDC trgDC
    
    DeleteDC srcDC
    
    Set srcPicture = Nothing
  Else
    Set dx_DirectDrawStaticSurface(staticSurfaceIndex) = dx_DirectDraw.CreateSurfaceFromFile(m_TexturePath & surfaceFileName, dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex))
  End If
  
  If transparentColour <> -2 Then
    If transparentColour = -1 Then
      dx_DirectDrawStaticSurface(staticSurfaceIndex).Lock t_Rect, dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex), 0, 0
      
      transparentColour = dx_DirectDrawStaticSurface(staticSurfaceIndex).GetLockedPixel(0, 0)
      
      dx_DirectDrawStaticSurface(staticSurfaceIndex).Unlock t_Rect
    End If
    
    colourkey.low = transparentColour
    colourkey.high = transparentColour
    
    dx_DirectDrawStaticSurface(staticSurfaceIndex).SetColorKey DDCKEY_SRCBLT, colourkey
  End If
  
  m_StaticSurfaceFileName(staticSurfaceIndex) = surfaceFileName
  m_StaticSurfaceValid(staticSurfaceIndex) = True
  
  If texturePriority > 0 Then dx_DirectDrawStaticSurface(staticSurfaceIndex).SetPriority texturePriority
  
  Init_TextureSurface = True
  
  Exit Function
  
badLoad:
  Release_StaticSurface staticSurfaceIndex
End Function

'switch the static surfaces - this is useful for animating textured surfaces
Public Sub Swap_StaticSurfaces(ByVal firstSurface As Long, ByVal secondSurface As Long)
  Dim t_DirectDrawStaticSurface As DirectDrawSurface7
  Dim t_DirectDrawStaticSurfaceDesc As DDSURFACEDESC2
  Dim t_StaticSurfaceWidth As Long, t_StaticSurfaceHeight As Long
  Dim t_StaticSurfaceFileName As String
  Dim t_StaticSurfaceValid As Boolean, t_StaticSurfaceTrans As Long
  Dim t_StaticSurfaceIsTexture As Boolean, t_StaticSurfacePriority As Long
  
  If firstSurface < 1 Or firstSurface > m_TotalStaticSurfaces Then Exit Sub
  If secondSurface < 1 Or secondSurface > m_TotalStaticSurfaces Then Exit Sub
  
  Set t_DirectDrawStaticSurface = dx_DirectDrawStaticSurface(firstSurface)
  t_DirectDrawStaticSurfaceDesc = dx_DirectDrawStaticSurfaceDesc(firstSurface)
  t_StaticSurfaceWidth = dx_StaticSurfaceWidth(firstSurface)
  t_StaticSurfaceHeight = dx_StaticSurfaceHeight(firstSurface)
  t_StaticSurfaceFileName = m_StaticSurfaceFileName(firstSurface)
  t_StaticSurfaceValid = m_StaticSurfaceValid(firstSurface)
  t_StaticSurfaceTrans = m_StaticSurfaceTrans(firstSurface)
  t_StaticSurfaceIsTexture = m_StaticSurfaceIsTexture(firstSurface)
  t_StaticSurfacePriority = m_StaticSurfacePriority(firstSurface)
  
  Set dx_DirectDrawStaticSurface(firstSurface) = dx_DirectDrawStaticSurface(secondSurface)
  dx_DirectDrawStaticSurfaceDesc(firstSurface) = dx_DirectDrawStaticSurfaceDesc(secondSurface)
  dx_StaticSurfaceWidth(firstSurface) = dx_StaticSurfaceWidth(secondSurface)
  dx_StaticSurfaceHeight(firstSurface) = dx_StaticSurfaceHeight(secondSurface)
  m_StaticSurfaceFileName(firstSurface) = m_StaticSurfaceFileName(secondSurface)
  m_StaticSurfaceValid(firstSurface) = m_StaticSurfaceValid(secondSurface)
  m_StaticSurfaceTrans(firstSurface) = m_StaticSurfaceTrans(secondSurface)
  m_StaticSurfaceIsTexture(firstSurface) = m_StaticSurfaceIsTexture(secondSurface)
  m_StaticSurfacePriority(firstSurface) = m_StaticSurfacePriority(secondSurface)
  
  Set dx_DirectDrawStaticSurface(secondSurface) = t_DirectDrawStaticSurface
  dx_DirectDrawStaticSurfaceDesc(secondSurface) = t_DirectDrawStaticSurfaceDesc
  dx_StaticSurfaceWidth(secondSurface) = t_StaticSurfaceWidth
  dx_StaticSurfaceHeight(secondSurface) = t_StaticSurfaceHeight
  m_StaticSurfaceFileName(secondSurface) = t_StaticSurfaceFileName
  m_StaticSurfaceValid(secondSurface) = t_StaticSurfaceValid
  m_StaticSurfaceTrans(secondSurface) = t_StaticSurfaceTrans
  m_StaticSurfaceIsTexture(secondSurface) = t_StaticSurfaceIsTexture
  m_StaticSurfacePriority(secondSurface) = t_StaticSurfacePriority
End Sub

'Clears the entire static surface
Public Sub Clear_StaticSurface(ByVal staticSurfaceIndex As Long, Optional ByVal clearColour As Long = 0)
  Dim t_Rect As RECT
  
  On Error Resume Next
  
  dx_DirectDrawStaticSurface(staticSurfaceIndex).BltColorFill t_Rect, clearColour
End Sub

Public Sub Set_SSurfacePath(ByVal surfaceFilePath As String)
  If surfaceFilePath = "Resource" Then
    m_StaticSurfacePath = "Resource"
  Else
    If Left$(surfaceFilePath, 1) = "\" Then surfaceFilePath = App.Path & surfaceFilePath
    
    If Right$(surfaceFilePath, 1) <> "\" Then
      m_StaticSurfacePath = surfaceFilePath & "\"
    Else
      m_StaticSurfacePath = surfaceFilePath
    End If
  End If
End Sub

Public Function Get_SSurfacePath() As String
  Get_SSurfacePath = m_StaticSurfacePath
End Function

Public Sub Set_TexturePath(ByVal textureFilePath As String)
  If textureFilePath = "Resource" Then
    m_TexturePath = "Resource"
  Else
    If Left$(textureFilePath, 1) = "\" Then textureFilePath = App.Path & textureFilePath
    
    If Right$(textureFilePath, 1) <> "\" Then
      m_TexturePath = textureFilePath & "\"
    Else
      m_TexturePath = textureFilePath
    End If
  End If
End Sub

Public Function Get_TexturePath() As String
  Get_TexturePath = m_TexturePath
End Function

Public Function Get_SSurfaceFileName(ByVal staticSurfaceIndex As Long) As String
  If staticSurfaceIndex < 1 Or staticSurfaceIndex > m_TotalStaticSurfaces Then
    Get_SSurfaceFileName = ""
    
    Exit Function
  End If
  
  If m_StaticSurfaceValid(staticSurfaceIndex) Then
    Get_SSurfaceFileName = m_StaticSurfaceFileName(staticSurfaceIndex)
  Else
    Get_SSurfaceFileName = ""
  End If
End Function

Public Sub Get_SSurfaceSettings(ByVal staticSurfaceIndex As Long, ByRef surfaceWidth As Long, ByRef surfaceHeight As Long, ByRef surfaceTransparency As Long, ByRef surfaceIsTexture As Boolean)
  If staticSurfaceIndex < 1 Or staticSurfaceIndex > m_TotalStaticSurfaces Then
    surfaceWidth = 0
    surfaceHeight = 0
    surfaceTransparency = 0
    surfaceIsTexture = False
    
    Exit Sub
  End If
  
  If m_StaticSurfaceValid(staticSurfaceIndex) Then
    surfaceWidth = dx_StaticSurfaceWidth(staticSurfaceIndex)
    surfaceHeight = dx_StaticSurfaceHeight(staticSurfaceIndex)
    surfaceTransparency = m_StaticSurfaceTrans(staticSurfaceIndex)
    surfaceIsTexture = m_StaticSurfaceIsTexture(staticSurfaceIndex)
  Else
    surfaceWidth = 0
    surfaceHeight = 0
    surfaceTransparency = 0
    surfaceIsTexture = False
  End If
End Sub

Public Function Get_SSurfaceIndex(ByVal surfaceShortFileName As String) As Long
  Dim loop1 As Long
  
  For loop1 = 1 To m_TotalStaticSurfaces
    If Get_SSurfaceValid(loop1) Then
      If Get_SSurfaceFileName(loop1) = surfaceShortFileName Then
        Get_SSurfaceIndex = loop1
        
        Exit Function
      End If
    End If
  Next loop1
  
  Get_SSurfaceIndex = 0
End Function

Public Function Get_SSurfaceFreeIndex() As Long
  Dim loop1 As Long
  
  For loop1 = 1 To m_TotalStaticSurfaces
    If Get_SSurfaceValid(loop1) = False Then
      Get_SSurfaceFreeIndex = loop1
      
      Exit Function
    End If
  Next loop1
  
  Get_SSurfaceFreeIndex = 0
End Function

Public Function Get_SSurfaceValid(ByVal staticSurfaceIndex As Long) As Boolean
  If staticSurfaceIndex < 1 Or staticSurfaceIndex > m_TotalStaticSurfaces Then
    Get_SSurfaceValid = False
    
    Exit Function
  End If
  
  Get_SSurfaceValid = m_StaticSurfaceValid(staticSurfaceIndex)
End Function

'this is usually done when the surface is initialized, but if you need to change the transparent colour you can
Public Sub Set_SSurfaceTransparency(ByVal staticSurfaceIndex As Long, Optional ByVal transparentColour As Long = -1)
  Dim colourkey As DDCOLORKEY
  
  If staticSurfaceIndex < 1 Or staticSurfaceIndex > m_TotalStaticSurfaces Then Exit Sub
  
  m_StaticSurfaceTrans(staticSurfaceIndex) = transparentColour
  
  If transparentColour = -1 Then
    Dim t_Rect As RECT
    
    dx_DirectDrawStaticSurface(staticSurfaceIndex).Lock t_Rect, dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex), 0, 0
    
    transparentColour = dx_DirectDrawStaticSurface(staticSurfaceIndex).GetLockedPixel(0, 0)
    
    dx_DirectDrawStaticSurface(staticSurfaceIndex).Unlock t_Rect
  End If
  
  colourkey.low = transparentColour
  colourkey.high = transparentColour
   
  dx_DirectDrawStaticSurface(staticSurfaceIndex).SetColorKey DDCKEY_SRCBLT, colourkey
End Sub

Public Sub Release_StaticSurface(ByVal staticSurfaceIndex As Long)
  Dim t_DirectDrawStaticSurfaceDesc As DDSURFACEDESC2
  
  If staticSurfaceIndex < 1 Or staticSurfaceIndex > m_TotalStaticSurfaces Then Exit Sub
  
  Set dx_DirectDrawStaticSurface(staticSurfaceIndex) = Nothing
  dx_DirectDrawStaticSurfaceDesc(staticSurfaceIndex) = t_DirectDrawStaticSurfaceDesc
  dx_StaticSurfaceWidth(staticSurfaceIndex) = 0
  dx_StaticSurfaceHeight(staticSurfaceIndex) = 0
  m_StaticSurfaceFileName(staticSurfaceIndex) = ""
  m_StaticSurfaceValid(staticSurfaceIndex) = False
  m_StaticSurfaceTrans(staticSurfaceIndex) = 0
  m_StaticSurfaceIsTexture(staticSurfaceIndex) = False
  m_StaticSurfacePriority(staticSurfaceIndex) = 0
End Sub

Public Sub Release_AllStaticSurfaces(Optional ByVal onlyTextures As Boolean = False)
  Dim loop1 As Long
  
  For loop1 = 1 To m_TotalStaticSurfaces
    If m_StaticSurfaceIsTexture(loop1) = True Or onlyTextures = False Then
      Release_StaticSurface loop1
    End If
  Next loop1
End Sub

'displays the selected picture - tiled over the clipping window - good for scrolling backdrops
Public Sub Display_TiledImage(ByVal staticSurfaceIndex As Long, _
        Optional ByVal pictureWidth As Long = 0, Optional ByVal pictureHeight As Long = 0, _
        Optional ByVal sourceOffsetX As Long = 0, Optional ByVal sourceOffsetY As Long = 0, _
        Optional ByVal tileShiftX As Long = 0, Optional ByVal tileShiftY As Long = 0, _
        Optional ByVal displayTransparent As Boolean = False)
        
  Dim RowX As Long, RowY As Long, startCol As Long
  
  If pictureWidth <= 0 Then pictureWidth = dx_StaticSurfaceWidth(staticSurfaceIndex)
  If pictureHeight <= 0 Then pictureHeight = dx_StaticSurfaceHeight(staticSurfaceIndex)
  
  If tileShiftY < 0 Then
    RowY = (-tileShiftY) Mod pictureHeight - pictureHeight
  ElseIf tileShiftY > 0 Then
    RowY = -(tileShiftY Mod pictureHeight)
  End If
  
  If tileShiftX < 0 Then
    startCol = (-tileShiftX) Mod pictureWidth - pictureWidth
  ElseIf tileShiftX > 0 Then
    startCol = -(tileShiftX Mod pictureWidth)
  End If
  
  If displayTransparent Then
    Do While RowY < m_ClippingRectangleHeight
      RowX = startCol
      
      Do While RowX < m_ClippingRectangleWidth
        BlitTransparent staticSurfaceIndex, RowX, RowY, sourceOffsetX, sourceOffsetY, pictureWidth, pictureHeight
        
        RowX = RowX + pictureWidth
      Loop
      
      RowY = RowY + pictureHeight
    Loop
  Else
    Do While RowY < m_ClippingRectangleHeight
      RowX = startCol
      
      Do While RowX < m_ClippingRectangleWidth
        BlitSolid staticSurfaceIndex, RowX, RowY, sourceOffsetX, sourceOffsetY, pictureWidth, pictureHeight
        
        RowX = RowX + pictureWidth
      Loop
      
      RowY = RowY + pictureHeight
    Loop
  End If
End Sub

Public Sub Set_Font(mFont As IFont, Optional ByVal staticSurfaceIndex As Long = 0)
  Dim dds As DirectDrawSurface7
  
  On Error Resume Next
  
  If staticSurfaceIndex = 0 Then
    Set dds = dx_DirectDrawBackSurface
  Else
    Set dds = dx_DirectDrawStaticSurface(staticSurfaceIndex)
  End If
  
  dds.SetFont mFont
End Sub

Public Sub Draw_Text(textString As String, ByVal xPos As Long, ByVal yPos As Long, ByVal foreColour As Long, Optional ByVal backColour = 0, Optional ByVal Transparency As Boolean = True, Optional ByVal staticSurfaceIndex As Long = 0)
  On Error Resume Next
  
  If staticSurfaceIndex = 0 Then
    With dx_DirectDrawBackSurface
      .SetForeColor foreColour
      .SetFontBackColor backColour
      .SetFontTransparency Transparency
      
      .DrawText xPos + m_ClippingRectangleX, yPos + m_ClippingRectangleY, textString, False
    End With
  Else
    With dx_DirectDrawStaticSurface(staticSurfaceIndex)
      .SetForeColor foreColour
      .SetFontBackColor backColour
      .SetFontTransparency Transparency
      
      .DrawText xPos, yPos, textString, False
    End With
  End If
End Sub

Public Sub Draw_Box(ByVal xPos As Long, ByVal yPos As Long, ByVal boxWidth As Long, ByVal boxHeight As Long, ByVal lineColour As Long, Optional ByVal lineWidth As Long = 1, Optional ByVal fillColour As Long = 0, Optional ByVal lineStyle As LineStyles = LSSolid, Optional ByVal fillStyle As FillStyles = FSNoFill, Optional ByVal staticSurfaceIndex As Long = 0)
  On Error Resume Next
  
  If staticSurfaceIndex = 0 Then
    With dx_DirectDrawBackSurface
      .SetForeColor lineColour
      .setDrawStyle lineStyle
      .setDrawWidth lineWidth
      .SetFillColor fillColour
      .SetFillStyle fillStyle
      
      .DrawBox xPos + m_ClippingRectangleX, yPos + m_ClippingRectangleY, xPos + m_ClippingRectangleX + boxWidth, yPos + m_ClippingRectangleY + boxHeight
    End With
  Else
    With dx_DirectDrawStaticSurface(staticSurfaceIndex)
      .SetForeColor lineColour
      .setDrawStyle lineStyle
      .setDrawWidth lineWidth
      .SetFillColor fillColour
      .SetFillStyle fillStyle
      
      .DrawBox xPos, yPos, xPos + boxWidth, yPos + boxHeight
    End With
  End If
End Sub

Public Sub Draw_Ellipse(ByVal xPos As Long, ByVal yPos As Long, ByVal boundingWidth As Long, ByVal boundingHeight As Long, ByVal lineColour As Long, Optional ByVal lineWidth As Long = 1, Optional ByVal fillColour As Long = 0, Optional ByVal lineStyle As LineStyles = LSSolid, Optional ByVal fillStyle As FillStyles = FSNoFill, Optional ByVal staticSurfaceIndex As Long = 0)
  On Error Resume Next
  
  If staticSurfaceIndex = 0 Then
    With dx_DirectDrawBackSurface
      .SetForeColor lineColour
      .setDrawStyle lineStyle
      .setDrawWidth lineWidth
      .SetFillColor fillColour
      .SetFillStyle fillStyle
      
      .DrawEllipse xPos + m_ClippingRectangleX, yPos + m_ClippingRectangleY, xPos + m_ClippingRectangleX + boundingWidth, yPos + m_ClippingRectangleY + boundingHeight
    End With
  Else
    With dx_DirectDrawStaticSurface(staticSurfaceIndex)
      .SetForeColor lineColour
      .setDrawStyle lineStyle
      .setDrawWidth lineWidth
      .SetFillColor fillColour
      .SetFillStyle fillStyle
      
      .DrawEllipse xPos, yPos, xPos + boundingWidth, yPos + boundingHeight
    End With
  End If
End Sub

Public Sub Draw_Circle(ByVal xPos As Long, ByVal yPos As Long, ByVal circleRadius As Long, ByVal lineColour As Long, Optional ByVal lineWidth As Long = 1, Optional ByVal fillColour As Long = 0, Optional ByVal lineStyle As LineStyles = LSSolid, Optional ByVal fillStyle As FillStyles = FSNoFill, Optional ByVal staticSurfaceIndex As Long = 0)
  On Error Resume Next
  
  If staticSurfaceIndex = 0 Then
    With dx_DirectDrawBackSurface
      .SetForeColor lineColour
      .setDrawStyle lineStyle
      .setDrawWidth lineWidth
      .SetFillColor fillColour
      .SetFillStyle fillStyle
      
      .DrawCircle xPos + m_ClippingRectangleX, yPos + m_ClippingRectangleY, circleRadius
    End With
  Else
    With dx_DirectDrawStaticSurface(staticSurfaceIndex)
      .SetForeColor lineColour
      .setDrawStyle lineStyle
      .setDrawWidth lineWidth
      .SetFillColor fillColour
      .SetFillStyle fillStyle
      
      .DrawCircle xPos, yPos, circleRadius
    End With
  End If
End Sub

Public Sub Draw_Line(ByVal xPosStart As Long, ByVal yPosStart As Long, ByVal xPosStop As Long, ByVal yPosStop As Long, ByVal lineColour As Long, Optional ByVal lineWidth As Long = 1, Optional ByVal lineStyle As LineStyles = LSSolid, Optional ByVal staticSurfaceIndex As Long = 0)
  On Error Resume Next
  
  If staticSurfaceIndex = 0 Then
    With dx_DirectDrawBackSurface
      .SetForeColor lineColour
      .setDrawStyle lineStyle
      .setDrawWidth lineWidth
      
      .DrawLine xPosStart + m_ClippingRectangleX, yPosStart + m_ClippingRectangleY, xPosStop + m_ClippingRectangleX, yPosStop + m_ClippingRectangleY
    End With
  Else
    With dx_DirectDrawStaticSurface(staticSurfaceIndex)
      .SetForeColor lineColour
      .setDrawStyle lineStyle
      .setDrawWidth lineWidth
      
      .DrawLine xPosStart, yPosStart, xPosStop, yPosStop
    End With
  End If
End Sub

Public Sub Set_ClippingRectangle(ByVal xPos As Long, ByVal yPos As Long, ByVal Width As Long, ByVal Height As Long, Optional setD3DViewPort As Boolean = False)
  Dim m_ViewPort As D3DVIEWPORT7
  
  On Error Resume Next
  
  m_ClippingRectangleX = xPos
  m_ClippingRectangleY = yPos
  m_ClippingRectangleWidth = Width
  m_ClippingRectangleHeight = Height
  
  If setD3DViewPort Then
    With m_D3DViewPort(0)
      .X1 = DXDraw.m_ClippingRectangleX
      .Y1 = DXDraw.m_ClippingRectangleY
      .X2 = DXDraw.m_ClippingRectangleX + DXDraw.m_ClippingRectangleWidth
      .Y2 = DXDraw.m_ClippingRectangleY + DXDraw.m_ClippingRectangleHeight
    End With
    
    With m_ViewPort
      .lX = DXDraw.m_ClippingRectangleX
      .lY = DXDraw.m_ClippingRectangleY
      .lWidth = DXDraw.m_ClippingRectangleWidth
      .lHeight = DXDraw.m_ClippingRectangleHeight
      .minZ = 0
      .maxZ = 1
    End With
    
    dx_Direct3DDevice.SetViewport m_ViewPort
  End If
End Sub

Public Sub Get_ClippingRectangle(ByRef xPos As Long, ByRef yPos As Long, ByRef Width As Long, ByRef Height As Long)
  xPos = m_ClippingRectangleX
  yPos = m_ClippingRectangleY
  Width = m_ClippingRectangleWidth
  Height = m_ClippingRectangleHeight
End Sub

Public Sub Reset_ClippingRectangle(Optional ByVal resetD3DViewPort As Boolean = False)
  Dim m_ViewPort As D3DVIEWPORT7
  
  On Error Resume Next
  
  m_ClippingRectangleX = 0
  m_ClippingRectangleY = 0
  m_ClippingRectangleWidth = dx_Width
  m_ClippingRectangleHeight = dx_Height
  
  If resetD3DViewPort Then
    With m_D3DViewPort(0)
      .X1 = DXDraw.m_ClippingRectangleX
      .Y1 = DXDraw.m_ClippingRectangleY
      .X2 = DXDraw.m_ClippingRectangleX + DXDraw.m_ClippingRectangleWidth
      .Y2 = DXDraw.m_ClippingRectangleY + DXDraw.m_ClippingRectangleHeight
    End With
    
    With m_ViewPort
      .lX = DXDraw.m_ClippingRectangleX
      .lY = DXDraw.m_ClippingRectangleY
      .lWidth = DXDraw.m_ClippingRectangleWidth
      .lHeight = DXDraw.m_ClippingRectangleHeight
      .minZ = 0
      .maxZ = 1
    End With
    
    dx_Direct3DDevice.SetViewport m_ViewPort
  End If
End Sub

Public Sub CleanUp_DXDraw(Optional ByVal resetStaticSurfaces As Boolean = True)
  Dim loop1 As Long
  
  On Error Resume Next
  
  If Not dx_DirectDrawEnabled Then Exit Sub
  
  CleanUp_DX3D
  
  If dx_FullScreenMode Then
    dx_DirectDraw.RestoreDisplayMode
    dx_DirectDraw.SetCooperativeLevel 0, DDSCL_NORMAL
  End If
  
  If Not (dx_DirectDrawPrimarySurface Is Nothing) Then
    dx_DirectDrawPrimarySurface.SetClipper Nothing
    dx_DirectDrawPrimarySurface.DeleteAttachedSurface Nothing
  End If
  
  If Not (dx_DirectDrawBackSurface Is Nothing) Then dx_DirectDrawBackSurface.DeleteAttachedSurface Nothing
  
  dx_DirectDrawEnabled = False
  dx_FullScreenMode = False
  dx_Width = 0
  dx_Height = 0
  dx_BitDepth = 0
  
  Set dx_DirectDrawPrimarySurface = Nothing
  Set dx_DirectDrawPrimaryPalette = Nothing
  Set dx_DirectDrawPrimaryColourControl = Nothing
  Set dx_DirectDrawPrimaryGammaControl = Nothing
  Set dx_DirectDrawBackSurface = Nothing
  
  Erase dx_DirectDrawStaticSurface()
  
  Set m_ClippingWindow = Nothing
  Set dx_DirectDraw = Nothing
  
  If resetStaticSurfaces Then
    For loop1 = 1 To m_TotalStaticSurfaces
      m_StaticSurfaceFileName(loop1) = ""
      m_StaticSurfaceValid(loop1) = False
    Next loop1
    
    m_TotalStaticSurfaces = 0
  End If
End Sub




'**********************************************************************************************************

'**************************** the following routines are for use with 3D only ***************************

'**********************************************************************************************************

'initializes the Direct3D engine
Private Sub Initialize_DX3D(ByVal requestedZBufferDepth As Long)
  Dim loop1 As Long, foundEnum As Long, pixelCaps As DDPIXELFORMAT
  Dim dx_DirectDrawSurfaceDesc As DDSURFACEDESC2, matTemp As D3DMATRIX
  Dim ddpfPixelFormat As DDPIXELFORMAT, d3dEnumPFs As Direct3DEnumPixelFormats
  Dim TextureEnum As Direct3DEnumPixelFormats
  
  On Error GoTo gotError
  
  Set dx_Direct3D = dx_DirectDraw.GetDirect3D()
  
  dx_D3DZDepth = 0
  
  If requestedZBufferDepth >= 0 Then ' find and attach an appropriate Z Buffer
    If dx_SystemMemoryOnly Then
      Set d3dEnumPFs = dx_Direct3D.GetEnumZBufferFormats("IID_IDirect3DRGBDevice")
    Else
      Set d3dEnumPFs = dx_Direct3D.GetEnumZBufferFormats("IID_IDirect3DHALDevice")
    End If
    
    loop1 = d3dEnumPFs.GetCount()
    
    foundEnum = 0
     
    Do While loop1 > 0
      d3dEnumPFs.GetItem loop1, ddpfPixelFormat
      
      With ddpfPixelFormat
        If .lFlags = DDPF_ZBUFFER Then
          If requestedZBufferDepth = 0 Then 'if no Z buffer depth specified then pick one that is closest to screen depth
            If .lZBufferBitDepth = dx_BitDepth Then
              dx_D3DZDepth = .lZBufferBitDepth
              foundEnum = loop1
              
              loop1 = 0
            ElseIf .lZBufferBitDepth > dx_D3DZDepth Then
              dx_D3DZDepth = .lZBufferBitDepth
              foundEnum = loop1
            End If
          ElseIf .lZBufferBitDepth = requestedZBufferDepth Then
            dx_D3DZDepth = .lZBufferBitDepth
            foundEnum = loop1
            
            loop1 = 0
          ElseIf .lZBufferBitDepth > dx_D3DZDepth Then
            dx_D3DZDepth = .lZBufferBitDepth
            foundEnum = loop1
          End If
        End If
      End With
      
      loop1 = loop1 - 1
    Loop
    
    If foundEnum > 0 Then
      d3dEnumPFs.GetItem foundEnum, ddpfPixelFormat
      
      With dx_DirectDrawSurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_PIXELFORMAT
        
        If dx_SystemMemoryOnly Then
          .ddsCaps.lCaps = DDSCAPS_ZBUFFER Or DDSCAPS_SYSTEMMEMORY
        Else
          .ddsCaps.lCaps = DDSCAPS_ZBUFFER Or DDSCAPS_VIDEOMEMORY
        End If
        
        .lWidth = dx_Width
        .lHeight = dx_Height
        .ddpfPixelFormat = ddpfPixelFormat
      End With
      
      Set dx_D3DZSurface = dx_DirectDraw.CreateSurface(dx_DirectDrawSurfaceDesc)
      
      dx_DirectDrawBackSurface.AddAttachedSurface dx_D3DZSurface
    End If
  End If
  
  'initialize the D3D device
  If dx_SystemMemoryOnly Then
    Set dx_Direct3DDevice = dx_Direct3D.CreateDevice("IID_IDirect3DRGBDevice", dx_DirectDrawBackSurface)
  Else
    Set dx_Direct3DDevice = dx_Direct3D.CreateDevice("IID_IDirect3DHALDevice", dx_DirectDrawBackSurface)
  End If
  
  'currently the library is set to handle its own lighting calculations
  dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_LIGHTING, False
  
  D3DTextureEnable = True
  
  Set_D3DProjection
  
  If requestedZBufferDepth >= 0 Then
    dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_ZENABLE, D3DZB_TRUE
  Else
    dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_ZENABLE, D3DZB_FALSE
  End If
  
  Set_D3DShadeMode True
  Set_D3DFilterMode
  Set_D3DWorldAmbient
  
  dx_Direct3DDevice.SetTextureStageState 0, D3DTSS_ADDRESS, D3DTADDRESS_WRAP
  
  'always use the diffuse surface colour for alpha (not the texture's alpha) unless software rendering
  If dx_SystemMemoryOnly = False Then
    dx_Direct3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
    dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_SRCBLEND, D3DBLEND_SRCALPHA
    dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_DESTBLEND, D3DBLEND_INVSRCALPHA
  End If
  
  dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, True
  dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, True
  
  dx_Direct3DEnabled = True
    
gotError:
  
End Sub

Public Sub Set_D3DProjection(Optional ByVal nearPlane As Single = 1, Optional ByVal farPlane As Single = 1000, Optional ByVal fieldOfView As Single = Pi / 2)
  Dim matTemp As D3DMATRIX
  
  On Error Resume Next
  
  dx_DirectX.IdentityMatrix matTemp
  dx_DirectX.ProjectionMatrix matTemp, nearPlane, farPlane, fieldOfView
  
  With matTemp 'normalize the projection matrix
    .rc11 = .rc11 / .rc34
    .rc22 = .rc22 / .rc34
    .rc33 = .rc33 / .rc34
    .rc43 = .rc43 / .rc34
    .rc34 = 1
  End With
    
  dx_Direct3DDevice.SetTransform D3DTRANSFORMSTATE_PROJECTION, matTemp
End Sub

Public Sub Set_D3DCamera(ByVal xPos As Single, ByVal yPos As Single, ByVal zPos As Single, _
      ByVal Rotation As Single, ByVal Tilt As Single, Optional ByVal xUp As Single = 0, _
      Optional ByVal yUp As Single = 1, Optional ByVal zUp As Single = 0)
      
  Dim matView  As D3DMATRIX, vecFrom As D3DVECTOR, vecTo As D3DVECTOR, vecUp As D3DVECTOR
  
  On Error Resume Next
  
  With vecFrom
    .X = xPos
    .Y = yPos
    .z = zPos
    
    vecTo.X = .X + Cos(Tilt) * Sin(Rotation)
    vecTo.Y = .Y + Sin(Tilt)
    vecTo.z = .z + Cos(Tilt) * Cos(Rotation)
  End With
  
  With vecUp
    .X = xUp
    .Y = yUp
    .z = zUp
  End With
  
  dx_DirectX.IdentityMatrix matView
  
  dx_DirectX.ViewMatrix matView, vecFrom, vecTo, vecUp, 0
    
  dx_Direct3DDevice.SetTransform D3DTRANSFORMSTATE_VIEW, matView
End Sub

'changes to this will only take effect on objects that have their vertices re-calculated afterwards
Public Sub Set_D3DWorldAmbient(Optional ByVal Red As Single = 0, Optional ByVal Green As Single = 0, Optional ByVal Blue As Single = 0, Optional ByVal brightnessAdjust As Single = 0, Optional ByVal Intensity As Single = 1)
  D3DWorldAmbient_Red = Red
  D3DWorldAmbient_Green = Green
  D3DWorldAmbient_Blue = Blue
  
  D3DWorldBrightnessAdjust = brightnessAdjust
  D3DWorldIntensity = Intensity
End Sub

Public Sub Set_D3DZBufferActive(Optional ByVal ZBufferActive As Boolean = True)
  On Error Resume Next
  
  If dx_D3DZSurface Is Nothing Then Exit Sub
  
  If ZBufferActive Then
    dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_ZENABLE, D3DZB_TRUE
  Else
    dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_ZENABLE, D3DZB_FALSE
  End If
End Sub

Public Sub Set_D3DShadeMode(Optional ByVal shadeModeGouraud As Boolean = False)
  On Error Resume Next
  
  If shadeModeGouraud Then
    dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_SHADEMODE, D3DSHADE_GOURAUD
  Else
    dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_SHADEMODE, D3DSHADE_FLAT
  End If
End Sub

Public Sub Set_D3DFilterMode(Optional ByVal filterModeBiLinear As Boolean = False)
  On Error Resume Next
  
  If filterModeBiLinear Then
    dx_Direct3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTFP_LINEAR
    dx_Direct3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTFG_LINEAR
    dx_Direct3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTFN_LINEAR
  Else
    dx_Direct3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTFP_POINT
    dx_Direct3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTFG_POINT
    dx_Direct3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTFN_POINT
  End If
End Sub

Public Sub Set_D3DWorldFog(Optional ByVal fogEnable As Boolean = False, _
      Optional fogStart As Single = 0, Optional fogEnd As Single = 10000, _
      Optional fogColour As Long = &HFFFFFF)
  
  On Error Resume Next
  
  dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_FOGENABLE, fogEnable
  dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_FOGCOLOR, fogColour
  dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_FOGVERTEXMODE, D3DFOG_LINEAR
  dx_Direct3DDevice.SetRenderStateSingle D3DRENDERSTATE_FOGSTART, fogStart
  dx_Direct3DDevice.SetRenderStateSingle D3DRENDERSTATE_FOGEND, fogEnd
  dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_RANGEFOGENABLE, True
End Sub

'starts the 3D scene renderer
Public Sub Begin_D3DScene()
  On Error Resume Next
  
  dx_Direct3DDevice.BeginScene
End Sub

'stops the 3D scene renderer and cause the back buffer to be updated
Public Sub End_D3DScene()
  On Error Resume Next
  
  dx_Direct3DDevice.EndScene
End Sub

Public Sub Clear_ManagedTextureMemory()
  On Error Resume Next
  
  dx_Direct3D.EvictManagedTextures
End Sub

'clean up Direct3D interface
Private Sub CleanUp_DX3D()
  On Error Resume Next
  
  dx_D3DZDepth = 0
  
  If Not (dx_Direct3DDevice Is Nothing) Then dx_Direct3DDevice.SetTexture 0, Nothing
  
  If Not (dx_Direct3D Is Nothing) Then dx_Direct3D.EvictManagedTextures
  
  Set dx_D3DZSurface = Nothing
  Set dx_Direct3DDevice = Nothing
  Set dx_Direct3D = Nothing
  
  dx_Direct3DEnabled = False
End Sub

'**********************************************************************************************************

'************************* palette control routines only used on 8 bit displays *************************

'**********************************************************************************************************

'Only used on 8 bit displays
Public Sub Reset_DefaultPalette()
  Dim m_palette(0 To 255) As PALETTEENTRY, loop1 As Long, temp1 As Long
  
  On Error Resume Next
  
  For loop1 = 0 To 31 'grey range
    With m_palette(loop1)
      temp1 = loop1 * 8 + loop1 \ 4
      .Red = temp1
      .Green = temp1
      .Blue = temp1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'red primary range
    With m_palette(loop1 + 32)
      .Red = loop1 * 16 + loop1
      .Green = 0
      .Blue = 0
    End With
  Next loop1
  
  For loop1 = 0 To 15 'red upper range
    With m_palette(loop1 + 48)
      .Red = 255
      .Green = loop1 * 16 + loop1
      .Blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'green primary range
    With m_palette(loop1 + 64)
      .Red = 0
      .Green = loop1 * 16 + loop1
      .Blue = 0
    End With
  Next loop1
  
  For loop1 = 0 To 15 'green upper range
    With m_palette(loop1 + 80)
      .Red = loop1 * 16 + loop1
      .Green = 255
      .Blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'blue primary range
    With m_palette(loop1 + 96)
      .Red = 0
      .Green = 0
      .Blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'blue upper range
    With m_palette(loop1 + 112)
      .Red = loop1 * 16 + loop1
      .Green = loop1 * 16 + loop1
      .Blue = 255
    End With
  Next loop1
  
  For loop1 = 0 To 15 'purple primary range
    With m_palette(loop1 + 128)
      .Red = loop1 * 16 + loop1
      .Green = 0
      .Blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'purple upper range
    With m_palette(loop1 + 144)
      .Red = 255
      .Green = loop1 * 16 + loop1
      .Blue = 255
    End With
  Next loop1
  
  For loop1 = 0 To 15 'yellow primary range
    With m_palette(loop1 + 160)
      .Red = loop1 * 16 + loop1
      .Green = loop1 * 16 + loop1
      .Blue = 0
    End With
  Next loop1
  
  For loop1 = 0 To 15 'yellow upper range
    With m_palette(loop1 + 176)
      .Red = 255
      .Green = 255
      .Blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'cyan primary range
    With m_palette(loop1 + 192)
      .Red = 0
      .Green = loop1 * 16 + loop1
      .Blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'cyan upper range
    With m_palette(loop1 + 208)
      .Red = loop1 * 16 + loop1
      .Green = 255
      .Blue = 255
    End With
  Next loop1
  
  For loop1 = 0 To 15 'brown primary range
    With m_palette(loop1 + 224)
      .Red = loop1 * 16
      .Green = loop1 * 12
      .Blue = loop1 * 5
    End With
  Next loop1
  
  For loop1 = 0 To 7 'brown upper range
    With m_palette(loop1 + 240)
      .Red = 255
      .Green = 191 + loop1 * 8
      .Blue = 96 + loop1 * 20
    End With
  Next loop1
  
  For loop1 = 0 To 7 'silver range
    With m_palette(loop1 + 248)
      .Red = 84 + loop1 * 16
      .Green = 96 + loop1 * 16
      .Blue = 128 + loop1 * 16
    End With
  Next loop1
  
  dx_DirectDrawPrimaryPalette.SetEntries 0, 256, m_palette
End Sub

'Only used on 8 bit displays
Public Sub Load_PaletteFromBMP(bitmapFilePathName As String)
  Dim m_palette(0 To 255) As PALETTEENTRY
  Dim tPalette As DirectDrawPalette
  
  On Error Resume Next
  
  Set tPalette = dx_DirectDraw.LoadPaletteFromBitmap(bitmapFilePathName)
  
  tPalette.GetEntries 0, 256, m_palette
  dx_DirectDrawPrimaryPalette.SetEntries 0, 256, m_palette
End Sub

'Only used on 8 bit displays
Public Sub Set_Palette(paletteEntries() As PALETTEENTRY)
  On Error Resume Next
   
  dx_DirectDrawPrimaryPalette.SetEntries 0, 256, paletteEntries
End Sub

'Only used on 8 bit displays
Public Sub Get_Palette(ByRef paletteEntries() As PALETTEENTRY)
  On Error Resume Next
  
  dx_DirectDrawPrimaryPalette.GetEntries 0, 256, paletteEntries
End Sub

'Only used on 8 bit displays
Public Sub Set_PaletteEntry(ByVal entryNumber As Long, ByVal paletteRed As Long, ByVal paletteGreen As Long, ByVal paletteBlue As Long)
  Dim m_palette(0 To 0) As PALETTEENTRY
  
  On Error Resume Next
  
  With m_palette(0)
    .Blue = paletteBlue
    .Green = paletteGreen
    .Red = paletteRed
  End With
  
  dx_DirectDrawPrimaryPalette.SetEntries entryNumber, 1, m_palette
End Sub

'Only used on 8 bit displays
Public Sub Get_PaletteEntry(ByVal entryNumber As Long, ByRef paletteRed As Long, ByRef paletteGreen As Long, ByRef paletteBlue As Long)
  Dim m_palette(0 To 0) As PALETTEENTRY
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryPalette.GetEntries entryNumber, 1, m_palette
  
  With m_palette(0)
    paletteBlue = .Blue
    paletteGreen = .Green
    paletteRed = .Red
  End With
End Sub



'**********************************************************************************************************

'************* in case other modules or objects need to tap into any of these objects *****************

'**********************************************************************************************************

Public Function GetDirectX() As DirectX7
  Set GetDirectX = dx_DirectX
End Function

Public Function GetDirectDraw() As DirectDraw7
  Set GetDirectDraw = dx_DirectDraw
End Function

Public Function GetDirectDrawSurface(staticSurfaceIndex As Long) As DirectDrawSurface7
  If staticSurfaceIndex > 0 And staticSurfaceIndex <= m_TotalStaticSurfaces Then
    Set GetDirectDrawSurface = dx_DirectDrawStaticSurface(staticSurfaceIndex)
  End If
End Function

Public Function GetDirectDrawBackSurface() As DirectDrawSurface7
  Set GetDirectDrawBackSurface = dx_DirectDrawBackSurface
End Function

Public Function GetDirect3D() As Direct3D7
  Set GetDirect3D = dx_Direct3D
End Function

Public Function GetDirect3DDevice() As Direct3DDevice7
  Set GetDirect3DDevice = dx_Direct3DDevice
End Function
