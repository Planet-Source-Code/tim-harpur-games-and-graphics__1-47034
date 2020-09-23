VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Games & Graphics - ISOmetric Demo - Press Esc to Quit"
   ClientHeight    =   8820.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11700
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   588
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Target Visibility Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   16
      Top             =   8160
      Width           =   6315
      Begin VB.OptionButton optGhostMode 
         Caption         =   "Target Object Forced Visibility"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   3030
         TabIndex        =   18
         Top             =   270
         Value           =   -1  'True
         Width           =   3195
      End
      Begin VB.OptionButton optGhostMode 
         Caption         =   "Front Object Ghosting"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   270
         Width           =   2355
      End
   End
   Begin VB.HScrollBar scrZoom 
      Height          =   165
      Left            =   8070
      Max             =   10
      Min             =   1
      TabIndex        =   15
      Top             =   8640.001
      Value           =   4
      Width           =   3525
   End
   Begin VB.CheckBox chkShadow 
      Caption         =   "Shado&w Cursor"
      Height          =   345
      Left            =   3180
      TabIndex        =   14
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton cmdFlatten 
      Caption         =   "Fla&tten Map Terrain"
      Height          =   315
      Left            =   8070
      TabIndex        =   13
      Top             =   7800
      Width           =   3525
   End
   Begin VB.CommandButton cmdResetFog 
      Caption         =   "Reset FogOf&War"
      Height          =   315
      Left            =   5730
      TabIndex        =   12
      Top             =   7800
      Width           =   2265
   End
   Begin VB.CheckBox chkFog 
      Caption         =   "Show &FogOfWar"
      Height          =   315
      Left            =   5730
      TabIndex        =   11
      Top             =   7470
      Width           =   2265
   End
   Begin VB.CheckBox chkCursor 
      Caption         =   "3D C&ursor"
      Height          =   345
      Left            =   4200
      TabIndex        =   10
      Top             =   7470
      Width           =   1485
   End
   Begin VB.CheckBox chkNight 
      Caption         =   "&Night"
      Height          =   345
      Left            =   3180
      TabIndex        =   7
      Top             =   7470
      Width           =   975
   End
   Begin VB.CheckBox chkObjects 
      Caption         =   "&Show Objects"
      Height          =   315
      Left            =   1050
      TabIndex        =   6
      Top             =   7500
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.CommandButton cmdRandomize 
      Caption         =   "&Randomize Terrain Elevations"
      Height          =   315
      Left            =   8070
      TabIndex        =   5
      Top             =   7440
      Width           =   3525
   End
   Begin VB.CheckBox chkCBoxes 
      Caption         =   "Show All &Collision Boxes"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   7830
      Width           =   3135
   End
   Begin VB.HScrollBar scrAxis 
      Height          =   165
      Index           =   1
      LargeChange     =   1000
      Left            =   8070
      Max             =   30000
      Min             =   -30000
      SmallChange     =   67
      TabIndex        =   3
      Top             =   8430.001
      Width           =   3525
   End
   Begin VB.HScrollBar scrAxis 
      Height          =   165
      Index           =   0
      LargeChange     =   1000
      Left            =   8070
      Max             =   30000
      Min             =   -30000
      SmallChange     =   67
      TabIndex        =   2
      Top             =   8220.001
      Width           =   3525
   End
   Begin VB.CheckBox chkMapGrid 
      Caption         =   "&Grid"
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   7470
      Width           =   975
   End
   Begin VB.PictureBox picDisplay 
      ClipControls    =   0   'False
      Height          =   7245
      Left            =   60
      ScaleHeight     =   479
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   766
      TabIndex        =   0
      Top             =   120
      Width           =   11550
   End
   Begin VB.Label lblYpos 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6450
      TabIndex        =   9
      Top             =   8490.001
      Width           =   1545
   End
   Begin VB.Label lblXpos 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6450
      TabIndex        =   8
      Top             =   8160
      Width           =   1545
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCBoxes_Click()
  Dim objectList As clsISO_Object
  
  Set objectList = objectListStart0
  
  If chkCBoxes.value = 0 Then
    Do While Not (objectList Is Nothing)
      With objectList
        .CollisionBoxVisible = False
        
        Set objectList = .chainNext
      End With
    Loop
  Else
    Do While Not (objectList Is Nothing)
      With objectList
        .CollisionBoxVisible = True
        
        Set objectList = .chainNext
      End With
    Loop
  End If
End Sub

Private Sub cmdFlatten_Click()
  Dim loop1 As Long, loop2 As Long
  
  With mapBase
    For loop1 = 0 To 14
      For loop2 = 0 To 14
        .Set_MapCell_Altitude loop1, loop2, 0
      Next loop2
    Next loop1
  End With
    
  mapBase.Calculate_Map
  
  If chkNight.value <> 0 Then
    With mapBase
      .Apply_Light 5500 * .worldBase, 5500 * .worldBase, 4000 * .worldBase, &H505000
      .Apply_Light 10500 * .worldBase, 7500 * .worldBase, 4000 * .worldBase, &H505000
      .Apply_Light 2000 * .worldBase, 11000 * .worldBase, 4000 * .worldBase, &H505000
    End With
  End If
  
  Create_TerrainObjects
End Sub

Private Sub cmdRandomize_Click()
  Dim loop1 As Long, loop2 As Long
  Dim temp1 As Long, temp2 As Long, temp3 As Long, temp4 As Long, temp5 As Long
  
  cmdRandomize.Enabled = False
  
  With mapBase
    For loop1 = 0 To 14
      For loop2 = 0 To 14
        .Set_MapCell_Altitude loop1, loop2, Rnd() * 30
      Next loop2
    Next loop1
  End With
    
  mapBase.Validate_Map 14, 0
  mapBase.Calculate_Map
  
  If chkNight.value <> 0 Then
    With mapBase
      .Apply_Light 5500 * .worldBase, 5500 * .worldBase, 4000 * .worldBase, &H505000
      .Apply_Light 10500 * .worldBase, 7500 * .worldBase, 4000 * .worldBase, &H505000
      .Apply_Light 2000 * .worldBase, 11000 * .worldBase, 4000 * .worldBase, &H505000
    End With
  End If
  
  Create_TerrainObjects
  
  cmdRandomize.Enabled = True
End Sub

Private Sub chkNight_Click()
  With mapBase
    If chkNight.value = 0 Then
      .minBrightness = 0.4
      .maxBrightness = 1#
    Else
      .minBrightness = 0.1
      .maxBrightness = 0.4
    End If
    
    .Calculate_Map
    
    If chkNight.value <> 0 Then
      .Apply_Light 5500 * .worldBase, 5500 * .worldBase, 4000 * .worldBase, &H505000
      .Apply_Light 10500 * .worldBase, 7500 * .worldBase, 4000 * .worldBase, &H505000
      .Apply_Light 2000 * .worldBase, 11000 * .worldBase, 4000 * .worldBase, &H505000
    End If
  End With
  
  DXDraw.Swap_StaticSurfaces 3, 4
End Sub

Private Sub cmdResetFog_Click()
  mapBase.ApplyAll_FogThick
  
  oldCursorX = -100000
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape And masterMode <> -2 Then masterMode = -1
End Sub

Private Sub Form_Load() 'wait till for has loaded before proceeding to initialization
  picDisplay.Width = 770
  picDisplay.Height = 483
  
  masterMode = 1
End Sub

Public Sub Initialize_Display()
  DXDraw.Init_DXDrawWindow Me, picDisplay, 10 ', True
  DXDraw.Set_D3DFilterMode True
  
  masterMode = 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  cursorX = -1000000
  cursorY = -1000000
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'prevent unloading this form until cleanup has occurred
  If masterMode <> -2 Then
    'signal that the program should clean up and shut down
    masterMode = -1
    Cancel = 1
  End If
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If masterMode >= 3 Then
    If chkCBoxes.value = 0 Then
      If Not (targetObject Is Nothing) Then targetObject.CollisionBoxVisible = False
      
      Set targetObject = mapBase.Get_DisplayedISOmetricObject(X - DXDraw.m_ClippingRectangleX, Y - DXDraw.m_ClippingRectangleY, 2)
      
      If Not (targetObject Is Nothing) Then targetObject.CollisionBoxVisible = True
    Else
      Set targetObject = Nothing
    End If
    
    mapBase.Get_WorldXYZ_from_DisplayXY X - DXDraw.m_ClippingRectangleX, Y - DXDraw.m_ClippingRectangleY, 100, cursorX, cursorY, cursorZ
    
    If cursorX = -1000000000 Then
      lblXpos.Caption = "x: outbounds"
      lblYpos.Caption = "y: outbounds"
    Else
      lblXpos.Caption = "x: " & cursorX
      lblYpos.Caption = "y: " & cursorY
    End If
  End If
End Sub

Private Sub scrAxis_Change(Index As Integer)
  With mapBase 'map has been scrolled - adjust map and update object position
    .DisplayColumn_1000ths = scrAxis(0).value
    .DisplayRow_1000ths = scrAxis(1).value
  End With
End Sub

Private Sub scrZoom_Change()
  DXDraw.ISOmetricViewScale = scrZoom.value
  
  mapBase.Calculate_Map
  
  If chkNight.value <> 0 Then
    With mapBase
      .Apply_Light 5500 * .worldBase, 5500 * .worldBase, 4000 * .worldBase, &H505000
      .Apply_Light 10500 * .worldBase, 7500 * .worldBase, 4000 * .worldBase, &H505000
      .Apply_Light 2000 * .worldBase, 11000 * .worldBase, 4000 * .worldBase, &H505000
    End With
  End If
End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim tempX As Long, tempY As Long, tempZ As Long
  
  mapBase.Get_WorldXYZ_from_DisplayXY X - DXDraw.m_ClippingRectangleX, Y - DXDraw.m_ClippingRectangleY, 100, tempX, tempY, tempZ
  
  Create_EnergyRing tempX, tempY, tempZ
End Sub
