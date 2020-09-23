Attribute VB_Name = "modMain"
Option Explicit

Public timeCycle As Long, masterMode As Long
Public backdropOffset As Long

Public objectListStart0 As clsISO_Object 'ground level objects (map level 1)
Public objectListStart1 As clsISO_Object 'floating objects (map level 2)
Public mapBase As clsISO_MapObject
Public cloud As clsISO_Object, targetObject As clsISO_Object, cursorX As Long, cursorY As Long, cursorZ As Long, oldCursorX As Long, oldCursorY As Long
Public EnergyRings As clsISO_Object, Explosions As clsISO_Object

Public Sub Main()
   'initialize variables
  masterMode = 0
  
  Randomize
  
  'initialize the timer - this will return instantly (0 and relative) with the current time + 16 milliseconds
  timeCycle = DXDraw.DelayTillTime(0, 0, True) + 16
  
  Do While masterMode > -1 'this is the main game control loop
    'timing is set for 32 milliseconds with a maximum carryover of 16 milliseconds
    'since this demo is using windowed mode, with windows controls
    'this routine must be called with the callDoEvents set to TRUE
    timeCycle = DXDraw.DelayTillTime(timeCycle, 16, , True) + 32
    
    Select Case masterMode 'primary game dispatch
      Case 0
        frmMain.Show
      Case 1
        frmMain.Initialize_Display
      Case 2
        InitializeScene
      Case 3
        UpdateScene
    End Select
  Loop
  
  'clean up the objects
  mapBase.Destroy_Object 'this also detaches any ISOmetric world objects from the map
  objectListStart0.Destroy_Chain
  objectListStart1.Destroy_Chain
  
  'release input and display
  DXDraw.CleanUp_DXDraw
  
  'set the flag that it's OK for the main form to unload now, and then tell it to unload
  masterMode = -2
  Unload frmMain
End Sub

Private Sub InitializeScene()
  Dim loop1 As Long, loop2 As Long
  Dim temp1 As Long, temp2 As Long, temp3 As Long, temp4 As Long, temp5 As Long
  Dim zX As Long, zY As Long, zXY As Long, tArray(0 To 14) As Variant, xArray As Variant
  
  'leave a small blank gap around the display just to demonstrate clipping
  DXDraw.Set_ClippingRectangle 14, 14, 738, 450, True
  
  'load bitmaps used for graphics (if system is not rendering correctly try setting a forced bit depth)
  DXDraw.Init_TextureSurface 1, "ISO Terrain.gif", 256, 256
  DXDraw.Init_TextureSurface 2, "ISO EnergyRing.gif", 256, 256, -1
  DXDraw.Init_TextureSurface 3, "ISO ObjectsDark.gif", 256, 256, -1
  DXDraw.Init_TextureSurface 4, "ISO Objects.gif", 256, 256, -1
  DXDraw.Init_TextureSurface 5, "ISO Explosion.gif", 256, 256, -1
  
  DXDraw.ISOmetricViewScale = 4
  
  'initialize the base map
  Set mapBase = New clsISO_MapObject

  With mapBase
    'the width (in texture ratios) of the images used for texturing the terrain, the number of images per row
    'and column, and also the map's displayed cell dimension - worldbase (in pixels)
    .SurfacePixelRatioU = 1 / 256
    .SurfacePixelRatioV = 1 / 256
    .ImageWidthU = 32 / 256
    .ImageHeightV = 32 / 256
    .ImagesPerRow = 8
    .worldBase = 8
    
    'the brightness values are used to calculate front side/back side shading effects for the terrain
    .minBrightness = 0.4
    .maxBrightness = 1#
    .gridColour = &H80000000
    
    'any portion of the display that is off the map will be cleared to this colour
    .OutBoundsColour = &HFF000000
    
    'initialize the map for a 15x15 array of image cells
    .Init_MapArray 15, 15
    
    'this is bizzare method to create a  map -  usually you would design a map editor to
    'do this when working on large projects  - I am doing it this way for the example
    'so you may see how a map is put together, and then you can design an editor
    tArray(0) = Array(4, 16, 0, 4, 8, 0, 4, 8, 0, 4, 0, 0, 25, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 28, 0, 1, 26, 0, 1, 4, 0, 0, 4, 0, 4)
    tArray(1) = Array(4, 12, 0, 4, 8, 0, 4, 8, 0, 4, 0, 0, 14, 0, 1, 6, 0, 1, 16, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 13, 0, 1, 4, 0, 0, 4, 0, 0)
    tArray(2) = Array(4, 8, 0, 4, 8, 0, 4, 4, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 25, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 28, 0, 1, 19, 0, 1, 18, 0, 1)
    tArray(3) = Array(4, 8, 0, 4, 4, 0, 4, 4, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 24, 0, 1, 3, 0, 1, 3, 0, 1, 10, 0, 1, 16, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1)
    tArray(4) = Array(4, 8, 0, 4, 4, 0, 4, 0, 0, 4, 4, 0, 4, 0, 0, 4, 0, 0, 14, 0, 1, 7, 0, 1, 6, 0, 1, 20, 0, 1, 25, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1, 3, 0, 1)
    tArray(5) = Array(4, 4, 0, 4, 4, 0, 4, 4, 0, 4, 4, 0, 4, 0, 0, 4, 0, 1, 4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 14, 0, 1, 6, 0, 1, 16, 0, 1, 3, 0, 1, 10, 0, 1)
    tArray(6) = Array(4, 4, 0, 4, 4, 0, 4, 4, 0, 4, 4, 0, 4, 0, 0, 4, 0, 0, 4, 4, 0, 4, 4, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 14, 0, 1, 16, 0, 1, 13, 0, 1)
    tArray(7) = Array(4, 0, 0, 4, 0, 0, 4, 4, 0, 4, 4, 0, 4, 4, 0, 4, 4, 0, 4, 4, 0, 4, 4, 0, 4, 0, 0, 4, 0, 0, 4, 0, 1, 4, 0, 0, 4, 0, 0, 24, 0, 1, 12, 0, 1)
    tArray(8) = Array(4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 25, 0, 1, 15, 0, 1)
    tArray(9) = Array(4, 0, 1, 4, 0, 1, 4, 0, 1, 4, 0, 1, 4, 0, 1, 4, 0, 1, 4, 0, 1, 4, 0, 1, 4, 0, 1, 4, 0, 0, 8, 0, 1, 19, 0, 1, 18, 0, 1, 22, 0, 1, 13, 0, 1)
    tArray(10) = Array(4, 0, 1, 4, 0, 1, 4, 4, 1, 4, 4, 0, 4, 4, 0, 4, 4, 0, 4, 4, 0, 4, 4, 0, 4, 0, 0, 4, 0, 0, 24, 0, 1, 10, 0, 1, 7, 0, 1, 9, 0, 1, 20, 0, 1)
    tArray(11) = Array(4, 0, 1, 4, 4, 1, 4, 4, 1, 4, 4, 0, 4, 8, 0, 4, 4, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 8, 0, 1, 22, 0, 1, 13, 0, 1, 4, 0, 0, 4, 0, 0, 4, 4, 1)
    tArray(12) = Array(4, 0, 1, 4, 4, 0, 4, 4, 0, 4, 8, 0, 4, 4, 0, 4, 4, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 25, 0, 1, 10, 0, 1, 20, 0, 1, 4, 4, 0, 4, 4, 0, 4, 4, 1)
    tArray(13) = Array(4, 0, 1, 4, 4, 0, 4, 8, 0, 4, 8, 0, 4, 8, 0, 4, 4, 0, 4, 0, 0, 4, 0, 0, 4, 0, 0, 24, 0, 1, 13, 0, 1, 4, 0, 0, 4, 4, 0, 4, 8, 0, 4, 6, 0)
    tArray(14) = Array(4, 0, 1, 4, 4, 0, 4, 8, 0, 4, 12, 0, 4, 8, 0, 4, 4, 0, 4, 4, 0, 4, 0, 0, 8, 0, 1, 22, 0, 1, 12, 0, 1, 4, 0, 0, 4, 4, 0, 4, 8, 0, 4, 10, 0)
    
    For loop1 = 0 To 14
      For loop2 = 0 To 14
        xArray = tArray(loop1)
        
        .Set_MapCell loop1, loop2, 1, xArray(loop2 * 3), 0, xArray(loop2 * 3 + 1), xArray(loop2 * 3 + 2)
      Next loop2
    Next loop1
    
    'calculate the map's vertices - must be done before the map can be displayed
    .Validate_Map 7, 7
    .Calculate_Map
  End With
  
  Create_TerrainObjects
  createCloud
  
  masterMode = 3
End Sub

'this is the main render loop
Private Sub UpdateScene()
  Dim tObject1 As clsISO_Object, tObject2 As clsISO_Object, tVal As Long
  
  'redraw the ISOmetric image map
  If frmMain.chkMapGrid.value <> 0 Then
    mapBase.Render_Map_with_Grid
  Else
    mapBase.Render_Map
  End If
  
  Set tObject1 = EnergyRings 'process/animate any energy ring objects
  
  Do While Not (tObject1 Is Nothing)
    With tObject1
      Set tObject2 = .chainNext
      
      .FrameTickCounter = .FrameTickCounter - 1
      If .FrameTickCounter <= 0 Then
        .FrameTickCounter = .TicksPerFrame
        .MainActionTickCounter = .MainActionTickCounter - 1
        
        If .MainActionTickCounter <= 0 Then
          mapBase.Remove_ISOmetricObject tObject1 'remove the object from both the map
          
          .RemoveFrom_Chain EnergyRings 'and its own ISOmetric object chain
        Else
          tVal = (.MainActionTickCounter * 255) \ .TicksPerMainAction
          
          With .SurfaceList
            .DisplayHeight = .DisplayHeight + 4
            .DisplayWidth = .DisplayHeight * 2
            
            .DisplayOffsetX = -.DisplayWidth \ 2
            .DisplayOffsetY = -.DisplayHeight \ 2
            
            .ActionFrame = .ActionFrame + 1
            If .ActionFrame > .ActionLast Then .ActionFrame = 0
            
            If tVal > 127 Then
              tVal = tVal - 128
              
              .SurfaceColour = (tVal * &H1000000) Or &H80FFFFFF
            Else
              .SurfaceColour = (tVal * &H1000000) Or &HFFFFFF
            End If
          End With
        End If
      End If
    End With
    
    Set tObject1 = tObject2
  Loop
  
  Set tObject1 = Explosions 'process/animate any explosion objects
  
  Do While Not (tObject1 Is Nothing)
    With tObject1
      Set tObject2 = .chainNext
      
      .FrameTickCounter = .FrameTickCounter - 1
      If .FrameTickCounter <= 0 Then
        .FrameTickCounter = .TicksPerFrame
        .MainActionTickCounter = .MainActionTickCounter - 1
        
        If .MainActionTickCounter <= 0 Then
          mapBase.Remove_ISOmetricObject tObject1 'remove the object from both the map
          
          .RemoveFrom_Chain Explosions 'and its own ISOmetric object chain
        Else
          With .SurfaceList
            .ActionFrame = .ActionFrame + 1
            If .ActionFrame > .ActionLast Then .ActionFrame = .ActionLast
          End With
        End If
      End If
    End With
    
    Set tObject1 = tObject2
  Loop
  
  With cloud 'move cloud
    .PosY_1000ths = .PosY_1000ths + 500
    .PosX_1000ths = .PosX_1000ths + 500
    
    If .PosY_1000ths > 800000 Then
      .PosY_1000ths = -100000
      .PosX_1000ths = -100000
    End If
    
    ' this isn't really necessary since it's the only object in the map's level 1 chain
    mapBase.Move_ISOmetricObject cloud
  End With
  
  If (targetObject Is Nothing) Or (frmMain.chkObjects.value = 0) Then
    If frmMain.chkCursor.value = 0 Then ' convert world to map coordinates by dividing by worldbase * 1000
      If frmMain.chkShadow.value = 0 Then
        mapBase.Draw_TerrainCursor cursorY \ 8000, cursorX \ 8000, &HFFFFFFFF
      Else
        mapBase.Draw_TerrainCursor cursorY \ 8000, cursorX \ 8000, &HFFFFFFFF, , True
      End If
    Else
      If frmMain.chkShadow.value = 0 Then
        mapBase.Draw_TerrainCursor cursorY \ 8000, cursorX \ 8000, &HFFFFFFFF, True
      Else
        mapBase.Draw_TerrainCursor cursorY \ 8000, cursorX \ 8000, &HFFFFFFFF, True, True
      End If
    End If
  End If
  
  'redraw the ISOmetric world objects
  If frmMain.chkObjects.value <> 0 Then
    With mapBase
      If frmMain.optGhostMode(0).value Then
        If Not (targetObject Is Nothing) Then
          Dim cRow As Long, cCol As Long
          
          With targetObject
            cRow = .mapChainRow
            cCol = .mapChainColumn
          End With
        
          .Ghost_MapCell cRow, cCol + 1
          .Ghost_MapCell cRow + 1, cCol
          .Ghost_MapCell cRow + 1, cCol + 1
        End If
        
        If frmMain.chkFog.value <> 0 Then
          .Render_WorldObjects True
        Else
          .Render_WorldObjects
        End If
      Else
        If frmMain.chkFog.value <> 0 Then
          .Render_WorldObjects True, targetObject
        Else
          .Render_WorldObjects , targetObject
        End If
      End If
    End With
  End If
  
  If frmMain.chkFog.value <> 0 Then 'display FogOfWar
    mapBase.Restore_Fog oldCursorX, oldCursorY, 30000
    
    oldCursorX = cursorX
    oldCursorY = cursorY
    
    mapBase.Clear_Fog cursorX, cursorY, 30000
    mapBase.Render_FogOfWar
  End If
    
  DXDraw.RefreshDisplay
End Sub

'routines to create various objects used in demo
'only the values that are relevant to the usage of the object need to be set
Public Sub createTree1(ByVal xPos_1000ths As Long, ByVal yPos_1000ths As Long, ByVal zPos_1000ths As Long)
  Dim newObject As clsISO_Object, newSurface As clsISO_ObjectSurface
  
  Set newObject = New clsISO_Object
    
  With newObject
    .TypeID = 1
    
    .Collision_Mask = 2
    .Collision_WidthX_1000ths = 3000
    .Collision_WidthY_1000ths = 3000
    .Collision_WidthZ_1000ths = 12500
    .CollisionBoxColour = &HFFFFFF00
    
    .PosX_1000ths = xPos_1000ths
    .PosY_1000ths = yPos_1000ths
    .PosZ_1000ths = zPos_1000ths
  
    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .SurfaceOffsetU = 2 / 256
      .SurfaceOffsetV = 88 / 256
      .SurfaceIndex = 4
      
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      .ImageWidthU = 56 / 256
      .ImageHeightV = 62 / 256
      
      .DisplayOffsetX = -7
      .DisplayOffsetY = -14
      .DisplayWidth = 14
      .DisplayHeight = 16
      
      .ActionFrame = -1 'this object is not animated - it has only one frame
    End With
    
    .Add_Surface newSurface
  End With
  
  newObject.AddTo_Chain objectListStart0
End Sub

Public Sub createTree2(ByVal xPos_1000ths As Long, ByVal yPos_1000ths As Long, ByVal zPos_1000ths As Long)
  Dim newObject As clsISO_Object, newSurface As clsISO_ObjectSurface
  
  Set newObject = New clsISO_Object
    
  With newObject
    .TypeID = 1
    
    .Collision_Mask = 2
    .Collision_WidthX_1000ths = 3000
    .Collision_WidthY_1000ths = 3000
    .Collision_WidthZ_1000ths = 12000
    .CollisionBoxColour = &HFFFFFF00
    
    .PosX_1000ths = xPos_1000ths
    .PosY_1000ths = yPos_1000ths
    .PosZ_1000ths = zPos_1000ths
 
    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .SurfaceOffsetU = 62 / 256
      .SurfaceOffsetV = 92 / 256
      .SurfaceIndex = 4
      
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      .ImageWidthU = 46 / 256
      .ImageHeightV = 58 / 256
      
      .DisplayOffsetX = -6
      .DisplayOffsetY = -14
      .DisplayWidth = 12
      .DisplayHeight = 15
      
      .ActionFrame = -1 'this object is not animated - it has only one frame
    End With
    
    .Add_Surface newSurface
  End With
  
  newObject.AddTo_Chain objectListStart0
End Sub

Public Sub createTree3(ByVal xPos_1000ths As Long, ByVal yPos_1000ths As Long, ByVal zPos_1000ths As Long)
  Dim newObject As clsISO_Object, newSurface As clsISO_ObjectSurface
  
  Set newObject = New clsISO_Object
    
  With newObject
    .TypeID = 1
    
    .Collision_Mask = 2
    .Collision_WidthX_1000ths = 1500
    .Collision_WidthY_1000ths = 1500
    .Collision_WidthZ_1000ths = 11500
    .CollisionBoxColour = &HFFFFFF00
    
    .PosX_1000ths = xPos_1000ths
    .PosY_1000ths = yPos_1000ths
    .PosZ_1000ths = zPos_1000ths
 
    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .SurfaceOffsetU = 113 / 256
      .SurfaceOffsetV = 93 / 256
      .SurfaceIndex = 4
      
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      .ImageWidthU = 32 / 256
      .ImageHeightV = 58 / 256
      
      .DisplayOffsetX = -4
      .DisplayOffsetY = -14
      .DisplayWidth = 8
      .DisplayHeight = 15
      
      .ActionFrame = -1 'this object is not animated - it has only one frame
    End With
    
    .Add_Surface newSurface
  End With
  
  newObject.AddTo_Chain objectListStart0
End Sub

Public Sub createTree4(ByVal xPos_1000ths As Long, ByVal yPos_1000ths As Long, ByVal zPos_1000ths As Long)
  Dim newObject As clsISO_Object, newSurface As clsISO_ObjectSurface
  
  Set newObject = New clsISO_Object
    
  With newObject
    .TypeID = 1
    
    .Collision_Mask = 2
    .Collision_WidthX_1000ths = 1000
    .Collision_WidthY_1000ths = 1000
    .Collision_WidthZ_1000ths = 11000
    .CollisionBoxColour = &HFFFFFF00
    
    .PosX_1000ths = xPos_1000ths
    .PosY_1000ths = yPos_1000ths
    .PosZ_1000ths = zPos_1000ths

    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .SurfaceOffsetU = 149 / 256
      .SurfaceOffsetV = 95 / 256
      .SurfaceIndex = 4
      
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      .ImageWidthU = 26 / 256
      .ImageHeightV = 54 / 256
      
      .DisplayOffsetX = -3
      .DisplayOffsetY = -13
      .DisplayWidth = 7
      .DisplayHeight = 14
      
      .ActionFrame = -1 'this object is not animated - it has only one frame
    End With
    
    .Add_Surface newSurface
  End With
  
  newObject.AddTo_Chain objectListStart0
End Sub

Public Sub createHouse1(ByVal xPos_1000ths As Long, ByVal yPos_1000ths As Long, ByVal zPos_1000ths As Long)
  Dim newObject As clsISO_Object, newSurface As clsISO_ObjectSurface
  
  Set newObject = New clsISO_Object
    
  With newObject
    .TypeID = 2
    
    .Collision_Mask = 4
    .Collision_WidthX_1000ths = 4000
    .Collision_WidthY_1000ths = 5500
    .Collision_WidthZ_1000ths = 10000
    .CollisionBoxColour = &HFFFF00FF
    
    .PosX_1000ths = xPos_1000ths
    .PosY_1000ths = yPos_1000ths
    .PosZ_1000ths = zPos_1000ths

    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .SurfaceOffsetU = 181 / 256
      .SurfaceOffsetV = 4 / 256
      .SurfaceIndex = 4
      
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      .ImageWidthU = 68 / 256
      .ImageHeightV = 72 / 256
      
      .DisplayOffsetX = -9
      .DisplayOffsetY = -14
      .DisplayWidth = 17
      .DisplayHeight = 18
      
      .ActionFrame = -1 'this object is not animated - it has only one frame
    End With
    
    .Add_Surface newSurface
  End With
  
  newObject.AddTo_Chain objectListStart0
End Sub

Public Sub createHouse2(ByVal xPos_1000ths As Long, ByVal yPos_1000ths As Long, ByVal zPos_1000ths As Long)
  Dim newObject As clsISO_Object, newSurface As clsISO_ObjectSurface
  
  Set newObject = New clsISO_Object
    
  With newObject
    .TypeID = 2
    
    .Collision_Mask = 4
    .Collision_WidthX_1000ths = 4500
    .Collision_WidthY_1000ths = 4500
    .Collision_WidthZ_1000ths = 10000
    .CollisionBoxColour = &HFFFF00FF
    
    .PosX_1000ths = xPos_1000ths
    .PosY_1000ths = yPos_1000ths
    .PosZ_1000ths = zPos_1000ths
    
    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .SurfaceOffsetU = 177 / 256
      .SurfaceOffsetV = 84 / 256
      .SurfaceIndex = 4
      
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      .ImageWidthU = 76 / 256
      .ImageHeightV = 72 / 256
      
      .DisplayOffsetX = -10
      .DisplayOffsetY = -14
      .DisplayWidth = 19
      .DisplayHeight = 18
      
      .ActionFrame = -1 'this object is not animated - it has only one frame
    End With
    
    .Add_Surface newSurface
  End With
  
  newObject.AddTo_Chain objectListStart0
End Sub

Public Sub createTower(ByVal xPos_1000ths As Long, ByVal yPos_1000ths As Long, ByVal zPos_1000ths As Long)
  Dim newObject As clsISO_Object, newSurface As clsISO_ObjectSurface
  
  Set newObject = New clsISO_Object
    
  With newObject
    .TypeID = 3
    
    .Collision_Mask = 4
    .Collision_WidthX_1000ths = 4000
    .Collision_WidthY_1000ths = 3500
    .Collision_WidthZ_1000ths = 14000
    .CollisionBoxColour = &HFFFF00FF
    
    .PosX_1000ths = xPos_1000ths
    .PosY_1000ths = yPos_1000ths
    .PosZ_1000ths = zPos_1000ths
    
    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .SurfaceOffsetU = 1 / 256
      .SurfaceOffsetV = 0
      .SurfaceIndex = 4
      
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      .ImageWidthU = 54 / 256
      .ImageHeightV = 86 / 256
      
      .DisplayOffsetX = -7
      .DisplayOffsetY = -18
      .DisplayWidth = 14
      .DisplayHeight = 22
      
      .ActionFrame = -1 'this object is not animated - it has only one frame
    End With
    
    .Add_Surface newSurface
  End With
  
  newObject.AddTo_Chain objectListStart0
End Sub

Public Sub createTowerWallX(ByVal xPos_1000ths As Long, ByVal yPos_1000ths As Long, ByVal zPos_1000ths As Long)
  Dim newObject As clsISO_Object, newSurface As clsISO_ObjectSurface
  
  Set newObject = New clsISO_Object
    
  With newObject
    .TypeID = 4
    
    .Collision_Mask = 4
    .Collision_WidthX_1000ths = 4000
    .Collision_WidthY_1000ths = 3500
    .Collision_WidthZ_1000ths = 9000
    .CollisionBoxColour = &HFFFF00FF
    
    .PosX_1000ths = xPos_1000ths
    .PosY_1000ths = yPos_1000ths
    .PosZ_1000ths = zPos_1000ths

    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .SurfaceOffsetU = 56 / 256
      .SurfaceOffsetV = 1 / 256
      .SurfaceIndex = 4
      
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      .ImageWidthU = 56 / 256
      .ImageHeightV = 64 / 256
      
      .DisplayOffsetX = -7
      .DisplayOffsetY = -13
      .DisplayWidth = 14
      .DisplayHeight = 16
      
      .ActionFrame = -1 'this object is not animated - it has only one frame
    End With
    
    .Add_Surface newSurface
  End With
  
  newObject.AddTo_Chain objectListStart0
End Sub

Public Sub createTowerWallY(ByVal xPos_1000ths As Long, ByVal yPos_1000ths As Long, ByVal zPos_1000ths As Long)
  Dim newObject As clsISO_Object, newSurface As clsISO_ObjectSurface
  
  Set newObject = New clsISO_Object
    
  With newObject
    .TypeID = 4
    
    .Collision_Mask = 4
    .Collision_WidthX_1000ths = 3500
    .Collision_WidthY_1000ths = 4000
    .Collision_WidthZ_1000ths = 9000
    .CollisionBoxColour = &HFFFF00FF
    
    .PosX_1000ths = xPos_1000ths
    .PosY_1000ths = yPos_1000ths
    .PosZ_1000ths = zPos_1000ths
    
    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .SurfaceOffsetU = 116 / 256
      .SurfaceOffsetV = 1 / 256
      .SurfaceIndex = 4
      
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      .ImageWidthU = 56 / 256
      .ImageHeightV = 64 / 256
      
      .DisplayOffsetX = -7
      .DisplayOffsetY = -13
      .DisplayWidth = 14
      .DisplayHeight = 16
      
      .ActionFrame = -1 'this object is not animated - it has only one frame
    End With
    
    .Add_Surface newSurface
  End With
  
  newObject.AddTo_Chain objectListStart0
End Sub

Public Sub createCloud()
  Dim newSurface As clsISO_ObjectSurface
  
  If Not (objectListStart1 Is Nothing) Then
    mapBase.Remove_ISOmetricObjectChain objectListStart1
    objectListStart1.Destroy_Chain
  End If
  
  Set objectListStart1 = Nothing
  
  Set cloud = New clsISO_Object
    
  With cloud
    .TypeID = 5
    
    .Collision_Mask = 0
    
    .PosX_1000ths = -50000
    .PosY_1000ths = -50000
    .PosZ_1000ths = 50000
    
    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .SurfaceOffsetU = 0 / 256
      .SurfaceOffsetV = 156 / 256
      .SurfaceIndex = 4
      
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      .ImageWidthU = 256 / 256
      .ImageHeightV = 100 / 256
      
      .DisplayOffsetX = -32
      .DisplayOffsetY = -14
      .DisplayWidth = 64
      .DisplayHeight = 25
      
      .SurfaceColour = &HE0FFFFFF
      
      .ActionFrame = -1 'this object is not animated - it has only one frame
    End With
    
    .Add_Surface newSurface
  End With
  
  cloud.AddTo_Chain objectListStart1
  mapBase.Add_ISOmetricObjectChain objectListStart1, 2
End Sub

Public Sub Create_TerrainObjects()
  Dim isoMod As Long, loop1 As Long, temp1 As Long, temp2 As Long
  
  If Not (objectListStart0 Is Nothing) Then
    mapBase.Remove_ISOmetricObjectChain objectListStart0
    objectListStart0.Destroy_Chain
  End If
  
  Set objectListStart0 = Nothing
  
  'and again - normally this would be done with a map editor and then the file defining map and object
  'properties could just be loaded
  With mapBase
    isoMod = .worldBase * 1000
    
    For loop1 = 1 To 25
      Do While True
        temp1 = Rnd() * 14500
        temp2 = Rnd() * 14500
        
        If mapBase.Get_MapCell_FlatLock(temp1 \ 1000, temp2 \ 1000) = False Then Exit Do
      Loop
      
      createTree1 (temp2 * .worldBase), (temp1 * .worldBase), .Get_Altitude_from_WorldXY(temp2 * .worldBase, temp1 * .worldBase)
    Next loop1
    
    For loop1 = 1 To 25
      Do While True
        temp1 = Rnd() * 14500
        temp2 = Rnd() * 14500
        
        If mapBase.Get_MapCell_FlatLock(temp1 \ 1000, temp2 \ 1000) = False Then Exit Do
      Loop
      
      createTree2 (temp2 * .worldBase), (temp1 * .worldBase), .Get_Altitude_from_WorldXY(temp2 * .worldBase, temp1 * .worldBase)
    Next loop1
    
    For loop1 = 1 To 25
      Do While True
        temp1 = Rnd() * 14500
        temp2 = Rnd() * 14500
        
        If mapBase.Get_MapCell_FlatLock(temp1 \ 1000, temp2 \ 1000) = False Then Exit Do
      Loop
      
      createTree3 (temp2 * .worldBase), (temp1 * .worldBase), .Get_Altitude_from_WorldXY(temp2 * .worldBase, temp1 * .worldBase)
    Next loop1
    
    For loop1 = 1 To 25
      Do While True
        temp1 = Rnd() * 14500
        temp2 = Rnd() * 14500
        
        If mapBase.Get_MapCell_FlatLock(temp1 \ 1000, temp2 \ 1000) = False Then Exit Do
      Loop
      
      createTree4 (temp2 * .worldBase), (temp1 * .worldBase), .Get_Altitude_from_WorldXY(temp2 * .worldBase, temp1 * .worldBase)
    Next loop1
   
    
    createHouse1 (5.5 * isoMod), (5.5 * isoMod), .Get_Altitude_from_MapRC(5, 5)
    createHouse1 (10.5 * isoMod), (7.5 * isoMod), .Get_Altitude_from_MapRC(7, 10)
    
    createHouse2 (2 * isoMod), (11 * isoMod), .Get_Altitude_from_MapRC(11, 2)
    
    createTower (0.5 * isoMod), (9.5 * isoMod), .Get_Altitude_from_MapRC(9, 0)
    createTowerWallX (1 * isoMod) + 3000, (9 * isoMod) + 4500, .Get_Altitude_from_MapRC(9, 1)
    createTowerWallX (2 * isoMod) + 2000, (9 * isoMod) + 4500, .Get_Altitude_from_MapRC(9, 2)
    createTowerWallX (3 * isoMod) + 1000, (9 * isoMod) + 4500, .Get_Altitude_from_MapRC(9, 3)
    createTowerWallX (4 * isoMod) + 0, (9 * isoMod) + 4500, .Get_Altitude_from_MapRC(9, 4)
    createTowerWallX (5 * isoMod) - 1000, (9 * isoMod) + 4500, .Get_Altitude_from_MapRC(9, 5)
    createTowerWallX (6 * isoMod) - 2000, (9 * isoMod) + 4500, .Get_Altitude_from_MapRC(9, 6)
    createTowerWallX (7 * isoMod) - 3000, (9 * isoMod) + 4500, .Get_Altitude_from_MapRC(9, 7)
    createTower (8 * isoMod) - 4000, (9 * isoMod) + 4750, .Get_Altitude_from_MapRC(9, 8)
    createTowerWallY (0 * isoMod) + 4000, (10 * isoMod) + 3500, .Get_Altitude_from_MapRC(10, 0)
    createTowerWallY (0 * isoMod) + 4000, (11 * isoMod) + 2500, .Get_Altitude_from_MapRC(11, 0)
    createTowerWallY (0 * isoMod) + 4000, (12 * isoMod) + 1500, .Get_Altitude_from_MapRC(12, 0)
    createTowerWallY (0 * isoMod) + 4000, (13 * isoMod) + 500, .Get_Altitude_from_MapRC(13, 0)
    createTowerWallY (0) + 4000, (14 * isoMod) - 500, .Get_Altitude_from_MapRC(13, 0)
  End With
  
  Dim objectList As clsISO_Object
  
  Set objectList = objectListStart0
  
  If frmMain.chkCBoxes.value = 0 Then
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
  
  mapBase.Add_ISOmetricObjectChain objectListStart0, 1
End Sub

'the following routine creates a two piece  energy ring/explosion and adds it to both
'an ISOmetric object chain and the ISOmetric map
Public Sub Create_EnergyRing(ByVal xPos_1000ths As Long, yPos_1000ths As Long, zPos_1000ths As Long)
  Dim newObject As clsISO_Object, newSurface As clsISO_ObjectSurface
  
  Set newObject = New clsISO_Object
  
  With newObject
    .PosX_1000ths = xPos_1000ths
    .PosY_1000ths = yPos_1000ths
    .PosZ_1000ths = zPos_1000ths
    
    .MainActionTickCounter = 15
    .TicksPerMainAction = 15
    
    .TicksPerFrame = 2
    
    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .TypeID = 1
      
      .ZPriority = -1
      
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      
      .SurfaceOffsetU = 0 / 256
      .SurfaceOffsetV = 0 / 256
      
      .ImageWidthU = 0.5
      .ImageHeightV = 0.25
      
      .DisplayHeight = 2
      .DisplayWidth = 4
      
      .DisplayOffsetX = -.DisplayWidth \ 2
      .DisplayOffsetY = -.DisplayHeight \ 2
      
      .ImagesPerRow = 2
      
      .SurfaceIndex = 2
      
      .Set_AnimationSequence Array(0, 1, 2, 3, 4, 5, 6, 7)
    End With
    
    .Add_Surface newSurface
  End With
  
  newObject.AddTo_Chain EnergyRings
  
  mapBase.Add_ISOmetricObject newObject, 0
  
  Set newObject = New clsISO_Object
  
  With newObject
    .PosX_1000ths = xPos_1000ths
    .PosY_1000ths = yPos_1000ths
    .PosZ_1000ths = zPos_1000ths
    
    .MainActionTickCounter = 15
    .TicksPerMainAction = 15
    
    .TicksPerFrame = 2
    .FrameTickCounter = 2
    
    Set newSurface = New clsISO_ObjectSurface
    
    With newSurface
      .SurfacePixelRatioU = 1 / 256
      .SurfacePixelRatioV = 1 / 256
      
      .SurfaceOffsetU = 0 / 256
      .SurfaceOffsetV = 0 / 256
      
      .ImageWidthU = 50 / 256
      .ImageHeightV = 75 / 256
      
      .DisplayHeight = 30
      .DisplayWidth = 20
      
      .DisplayOffsetX = -10
      .DisplayOffsetY = -24
      
      .ImagesPerRow = 5
      
      .SurfaceIndex = 5
      
      .SurfaceColour = &HA0FFFFFF
      
      .Set_AnimationSequence Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14)
    End With
    
    .Add_Surface newSurface
  End With
  
  newObject.AddTo_Chain Explosions
  
  mapBase.Add_ISOmetricObject newObject, 1
End Sub

