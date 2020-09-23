VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Games & Graphics - 3D Demo - Press Esc to Quit"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10140
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   591
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   676
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Me.Left = 0
  Me.Top = 0
  
  'set the flag indicating that the form has loaded and the program may proceed with initialization
  masterMode = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'if the user tries to unload the program - signal a shut down, but abort until properly processed
  'unless the main program has already shut down (masterMode = -2) then it is ok to unload
  'if it isn't done this way, the form may be unloaded with the main program loop still executing
  'causing either an error or more likely a free running program fragment that can only be shut
  'down using Ctrl-Alt-Del (Task Manager)
  If masterMode <> -2 Then
    masterMode = -1
    
    Cancel = 1
  End If
End Sub

Public Sub Initialize_Display()
  'initialize the display
  If windowMode Then 'initialize the display for windowed 3D mode
    Select Case displayMode
      Case 1
        Me.Width = 640 * Screen.TwipsPerPixelX
        Me.Height = 480 * Screen.TwipsPerPixelY
      Case 2
        Me.Width = 800 * Screen.TwipsPerPixelX
        Me.Height = 600 * Screen.TwipsPerPixelY
      Case Else
        Me.Width = 1024 * Screen.TwipsPerPixelX
        Me.Height = 768 * Screen.TwipsPerPixelY
    End Select
    
    DXDraw.Init_DXDrawWindow Me, Me, 7, useSoftware, 0
  Else  'initialize the display for full screen 3D mode
    Select Case displayMode
      Case 1
        DXDraw.Init_DXDrawScreen Me, 640, 480, colourMode * 16, 7, useSoftware, 0
      Case 2
        DXDraw.Init_DXDrawScreen Me, 800, 600, colourMode * 16, 7, useSoftware, 0
      Case Else
        DXDraw.Init_DXDrawScreen Me, 1024, 768, colourMode * 16, 7, useSoftware, 0
    End Select
    
    'if full screen 3D failed, then set to 640x480 no 3D
    If DXDraw.TestDisplay3DValid() = False Then DXDraw.Init_DXDrawScreen Me, 640, 480, 16, 7
  End If
  
  'associate this form's font as the font to use with DXDraw.Draw_Text routine
  DXDraw.Set_Font Me.Font
  
  'initialize the keyboard input as normal, and mouse input - normal for window and exclusive for full screen
  DXInput.Init_DXInput Me
  DXInput.Acquire_Keyboard
  
  If windowMode Then
    DXInput.Acquire_Mouse
  Else
    DXInput.Acquire_Mouse IM_ForegroundExclusive
  End If
  
  'load the surface textures
  DXDraw.Set_TexturePath "\textures"
  
  DXDraw.Init_TextureSurface 1, "MetalSurface1.gif", 128, 128, , , forceTexBDepth
  DXDraw.Init_TextureSurface 2, "MetalSurface2.gif", 128, 128, , , forceTexBDepth
  DXDraw.Init_TextureSurface 3, "MetalSurface3.gif", 128, 128, , , forceTexBDepth
  DXDraw.Init_TextureSurface 4, "MetalBarrel.gif", 32, 128, , , forceTexBDepth
  DXDraw.Init_TextureSurface 5, "Flare1.gif", 64, 64, , , forceTexBDepth
  DXDraw.Init_TextureSurface 6, "Rock4.gif", 128, 128, , , forceTexBDepth
  DXDraw.Init_TextureSurface 7, "Thrust2.gif", 64, 64, -1, forceTexBDepth
  
  'indicate that display and input initialization is complete
  masterMode = 2
End Sub
