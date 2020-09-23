Attribute VB_Name = "modMain"
Option Explicit

Public timeCycle As Long, masterMode As Long
Public backdropOffset As Long

Public objectListStart As cls2D_Object, mapBase As cls2D_MapObject

Public Sub Main()
   'initialize variables
  masterMode = 0
  
  'initialize the timer - this will return instantly (0 and relative) with the current time + 16 milliseconds
  timeCycle = DXDraw.DelayTillTime(0, 0, True, True) + 16
  
  Do While masterMode > -1 'this is the main game control loop
    'timing is set for 16 milliseconds with a maximum carryover of 16 miiliseconds
    'since this demo is using windowed mode, with windows controls
    'this routine must be called with the callDoEvents set to TRUE
    timeCycle = DXDraw.DelayTillTime(timeCycle, 16, , True) + 16
    
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
  mapBase.Destroy_Object
  objectListStart.Destroy_Chain
  
  'release input and display
  DXDraw.CleanUp_DXDraw
  
  'set the flag that it's OK for the main form to unload now, and then tell it to unload
  masterMode = -2
  Unload frmMain
End Sub

Private Sub InitializeScene()
  Dim loop1 As Long, loop2 As Long, tArray(0 To 8) As Variant, xArray As Variant
  
  'leave a small blank gap around the display just to demonstrate clipping
  DXDraw.Set_ClippingRectangle 14, 14, 450, 450
  
  'load bitmaps used for graphics
  DXDraw.Init_StaticSurface 1, "2D Backdrop.gif", 640, 480
  DXDraw.Init_StaticSurface 2, "2D Terrain.gif", 300, 150
  DXDraw.Init_StaticSurface 3, "2D Animation.gif", 314, 130, -1
  DXDraw.Init_StaticSurface 4, "2D Explosions.gif", 384, 320, -1
  
  'initialize the base map
  Set mapBase = New cls2D_MapObject

  With mapBase
    .OutBoundsColour = 0
    .OutBoundsClear = True
    
    .gridColour = 0
    
    .worldBase = 50
    .ImagesPerRow = 6
    
    .Init_MapArray 9, 9
    
    'this is an OK method to create a small map, but usually you would design a map editor to
    'do this when working on large projects as this method would become way to slow and confusing
    tArray(0) = Array(2, 2, 2, 2, 2, 13, 7, 4, 8)
    tArray(1) = Array(16, 12, 12, 15, 2, 2, 13, 10, 14)
    tArray(2) = Array(11, 4, 4, 5, 12, 15, 2, 2, 2)
    tArray(3) = Array(13, 7, 4, 4, 4, 9, 2, 2, 16)
    tArray(4) = Array(2, 13, 7, 4, 8, 14, 2, 2, 11)
    tArray(5) = Array(2, 2, 13, 10, 14, 2, 16, 12, 6)
    tArray(6) = Array(2, 2, 2, 2, 2, 2, 11, 4, 4)
    tArray(7) = Array(2, 2, 2, 2, 2, 16, 6, 4, 4)
    tArray(8) = Array(2, 2, 2, 2, 2, 11, 4, 4, 4)
    
    For loop1 = 0 To 8
      For loop2 = 0 To 8
        xArray = tArray(loop1)
        
        .Set_MapCell loop1, loop2, 2, xArray(loop2), 0
      Next loop2
    Next loop1
  End With
  
  createBomber
  createFighter
  
  masterMode = 3
End Sub

Private Sub UpdateScene()
  Dim loop1 As Long
  Dim oList As cls2D_Object, oxList As cls2D_Object, sList As cls2D_ObjectSurface, sxList As cls2D_ObjectSurface
  
  If frmMain.optBackdrop(0).value Then
    DXDraw.Clear_Display
  ElseIf frmMain.optBackdrop(1).value Then
    DXDraw.Display_TiledImage 1, , , , , backdropOffset / 4, backdropOffset
    
    backdropOffset = backdropOffset - 1
  Else
    If frmMain.chkMapGrid.value <> 0 Then
      mapBase.Render_Map_with_Grid
    Else
      mapBase.Render_Map
    End If
  End If
  
  'process all the 2D objects in the User list
  Set oList = objectListStart
  
  Do While Not (oList Is Nothing)
    With oList
      Set oxList = oList
      Set oList = .chainNext
      
      Select Case .TypeID
        Case 1 'bomber
          If .MainAction = 1 Then 'flying normal
            If .direction = 1 Then
              .PosX_1000ths = .PosX_1000ths + 1000
              
              If .PosX_1000ths > 350000 Then .direction = -1
            Else
              .PosX_1000ths = .PosX_1000ths - 1000
              
              If .PosX_1000ths < 100000 Then .direction = 1
            End If
            
            .Calculate_CollsionBox
            
            .FrameTickCounter = .FrameTickCounter - 1 'check if surface frames need changing
            If .FrameTickCounter = 0 Then
              .FrameTickCounter = .TicksPerFrame
              
              Set sList = .SurfaceList
              
              Do While Not (sList Is Nothing)
                With sList
                  .ActionFrame = .ActionFrame + 1
                  If .ActionFrame > .ActionLast Then .ActionFrame = 0
                  
                  Set sList = .NextSurface
                End With
              Loop
            End If
          Else 'it's blowing up
            .FrameTickCounter = .FrameTickCounter - 1 'check if surface frames need changing
            If .FrameTickCounter = 0 Then
              .FrameTickCounter = .TicksPerFrame
              
              Set sList = .SurfaceList
              
              Do While Not (sList Is Nothing)
                With sList
                  Set sxList = sList
                  Set sList = .NextSurface
                  
                  If .TypeID >= 10 Then 'only process explosions
                    If .MainAction = 1 Then
                      .ActionFrame = .ActionFrame + 1
                      
                      If .ActionFrame > .ActionLast Then
                        If .TypeID = 16 Then 'this is the last explosion - reset the bomber
                          With oxList
                            .Collision_Mask = 1
                            If frmMain.chkCBoxes.value <> 0 Then .CollisionBoxVisible = True
                            
                            .MainAction = 1
                            
                            For loop1 = 2 To 3 'turn on thrust
                              .Get_SurfaceID(loop1).Visible = True
                            Next loop1
                          End With
                        End If
                          
                        'remove the explosion (2D surface objects don't need to be explicitly destroyed)
                        oxList.Remove_Surface sxList
                      End If
                    Else
                      .MainActionTickCounter = .MainActionTickCounter - 1
                      If .MainActionTickCounter = 0 Then
                        .MainAction = 1
                        
                        .Visible = True
                      End If
                    End If
                  End If
                End With
              Loop
            End If
          End If
        Case 2 'fighter
          If .direction = 1 Then
            .PosX_1000ths = .PosX_1000ths + 1500
            
            If .PosX_1000ths > 350000 Then .direction = -1
          Else
            .PosX_1000ths = .PosX_1000ths - 1500
            
            If .PosX_1000ths < 100000 Then .direction = 1
          End If
          
          .Calculate_CollsionBox
          
          .FrameTickCounter = .FrameTickCounter - 1 'check if surface frames need changing
          If .FrameTickCounter = 0 Then
            .FrameTickCounter = .TicksPerFrame
            
            Set sList = .SurfaceList
            
            Do While Not (sList Is Nothing)
              With sList
                .ActionFrame = .ActionFrame + 1
                If .ActionFrame > .ActionLast Then .ActionFrame = 0
                
                Set sList = .NextSurface
              End With
            Loop
          End If
        Case 3 'missile
          If .MainAction = 1 Then
            .PosY_1000ths = .PosY_1000ths - 3000
            
            If .PosY_1000ths < 10000 Then
              .PosY_1000ths = 10000
              
              blowUpMissile oxList 'missile hit edge of screen - so blow it up
            Else
              .Calculate_CollsionBox
              
              'it would actually be more efficient to hold the bomber in an object for quick reference
              'since there is only one of them in this demo
              With objectListStart.Get_ChainObjectID(1)
                If .Collision_Mask <> 0 Then 'is it already blowing up?
                  If .Test_Collision(oxList, 2) Then 'did we hit it? - if so then start blowing things up
                    blowUpMissile oxList
                    
                    blowUpBomber
                  End If
                End If
              End With
              
              .FrameTickCounter = .FrameTickCounter - 1 'check if surface frames need changing
              If .FrameTickCounter = 0 Then
                .FrameTickCounter = .TicksPerFrame
                
                With .SurfaceList
                  .ActionFrame = .ActionFrame + 1
                  If .ActionFrame > .ActionLast Then .ActionFrame = 0
                End With
              End If
            End If
          Else 'its blowing up
            .FrameTickCounter = .FrameTickCounter - 1 'once this sequence runs out - get rid of missile
              If .FrameTickCounter = 0 Then
                .FrameTickCounter = .TicksPerFrame
                
                With .SurfaceList
                  .ActionFrame = .ActionFrame + 1
                  
                  If .ActionFrame > .ActionLast Then
                    oxList.RemoveFrom_Chain objectListStart
                  
                    oxList.Destroy_Object
                  End If
                End With
              End If
          End If
      End Select
    End With
  Loop
  
  'redraw the 2D animation chain
  objectListStart.ReDraw_Chain 0, 0
  
  DXDraw.RefreshDisplay
End Sub

Public Sub createBomber()
  Dim newObject As cls2D_Object, newSurface As cls2D_ObjectSurface
  
  Set newObject = New cls2D_Object
  
  With newObject
    .TypeID = 1
    
    .PosX_1000ths = 225000
    .PosY_1000ths = 80000
    .direction = 1 'start sliding to right
    
    .MainAction = 1 'all normal - just flying
    
    .Collision_Mask = 1
    .CollisionBoxColour = &HFF&
    .CollisionBoxVisible = False
    
    .Collision_WidthX_1000ths = 60000
    .Collision_WidthY_1000ths = 20000
    
    .TicksPerFrame = 3 'master frame rate
    .FrameTickCounter = 3
    
    'create the main body surface
    Set newSurface = New cls2D_ObjectSurface
    
    With newSurface
      .TypeID = 1 'give it unique id
      
      .ZPriority = 1 'must appear over thrust surfaces
      
      .ImageWidth = 157 'initialize image size and source
      .ImageHeight = 61
      .ImagesPerRow = 2
      .SurfaceIndex = 3
      .SurfaceOffsetX = 0
      .SurfaceOffsetY = 0
      
      .DisplayOffsetX = -78
      .DisplayOffsetY = -30
      
      'set up the animation sequence
      .Set_AnimationSequence Array(1, 1, 1, 1, 1, 0)
    End With
    
    .Add_Surface newSurface
    
    'add thrust
    Set newSurface = New cls2D_ObjectSurface
    
    With newSurface
      .TypeID = 2 'give it unique id
      
      .ZPriority = 0 'must appear under other surfaces
      
      .ImageWidth = 16 'initialize image size and source
      .ImageHeight = 10
      .ImagesPerRow = 6
      .SurfaceIndex = 3
      .SurfaceOffsetX = 0
      .SurfaceOffsetY = 61
      
      .DisplayOffsetX = -36
      .DisplayOffsetY = -23
      
      'set up the animation sequence
      .Set_AnimationSequence Array(0, 1, 2, 3, 4, 5)
    End With
    
    .Add_Surface newSurface
    
    'add thrust
    Set newSurface = New cls2D_ObjectSurface
    
    With newSurface
      .TypeID = 3 'give it unique id
      
      .ZPriority = 0 'must appear under other surfaces
      
      .ImageWidth = 16 'initialize image size and source
      .ImageHeight = 10
      .ImagesPerRow = 6
      .SurfaceIndex = 3
      .SurfaceOffsetX = 0
      .SurfaceOffsetY = 61
      
      .DisplayOffsetX = 20
      .DisplayOffsetY = -23
      
      'set up the animation sequence
      .Set_AnimationSequence Array(0, 1, 2, 3, 4, 5)
    End With
    
    .Add_Surface newSurface
    .AddTo_Chain objectListStart
  End With
End Sub

Public Sub createFighter()
  Dim newObject As cls2D_Object, newSurface As cls2D_ObjectSurface
  
  Set newObject = New cls2D_Object
  
  With newObject
    .TypeID = 2
    
    .PosX_1000ths = 225000
    .PosY_1000ths = 380000
    .direction = 1 'start sliding to right
    
    .Collision_Mask = 2
    .CollisionBoxColour = &HFF00&
    .CollisionBoxVisible = False
    
    .Collision_WidthX_1000ths = 10000
    .Collision_WidthY_1000ths = 18000
    
    .TicksPerFrame = 3 'master frame rate
    .FrameTickCounter = 3
    
    'create the main body surface
    Set newSurface = New cls2D_ObjectSurface
    
    With newSurface
      .TypeID = 1 'give it unique id
      
      .ZPriority = 1 'must appear over thrust surfaces
      
      .ImageWidth = 35 'initialize image size and source
      .ImageHeight = 54
      .ImagesPerRow = 2
      .SurfaceIndex = 3
      .SurfaceOffsetX = 0
      .SurfaceOffsetY = 75
      
      .DisplayOffsetX = -17
      .DisplayOffsetY = -27
      
      'set up the animation sequence
      .Set_AnimationSequence Array(1, 1, 1, 1, 1, 0)
    End With
    
    .Add_Surface newSurface
    
    'add thrust
    Set newSurface = New cls2D_ObjectSurface
    
    With newSurface
      .TypeID = 2 'give it unique id
      
      .ZPriority = 0 'must appear under other surfaces
      
      .ImageWidth = 16 'initialize image size and source
      .ImageHeight = 10
      .ImagesPerRow = 5
      .SurfaceIndex = 3
      .SurfaceOffsetX = 97
      .SurfaceOffsetY = 61
      
      .DisplayOffsetX = -8
      .DisplayOffsetY = 18
      
      'set up the animation sequence
      .Set_AnimationSequence Array(0, 1, 2, 3, 4)
    End With
    
    .Add_Surface newSurface
    .AddTo_Chain objectListStart
  End With
End Sub

Public Sub createMissiles(ByVal xPos As Long, yPos As Long)
  Dim newObject As cls2D_Object, newSurface As cls2D_ObjectSurface
  
  Set newObject = New cls2D_Object
  
  With newObject
    .TypeID = 3
    .MainAction = 1 'flying normal on target
    
    .renderPriority = -1 'when using 2D (not 2D ISOmetric) this value can be set to User priority
    
    .PosX_1000ths = xPos
    .PosY_1000ths = yPos
    .direction = 1 'we don't use this in the demo, but normally the missile's direction would be set
    
    .Collision_Mask = 2
    .CollisionBoxColour = &HFF00&
    
    If frmMain.chkCBoxes.value = 0 Then
      .CollisionBoxVisible = False
    Else
      .CollisionBoxVisible = True
    End If
    
    .Collision_WidthX_1000ths = 2000
    .Collision_WidthY_1000ths = 4000
    
    .TicksPerFrame = 3 'master frame rate
    .FrameTickCounter = 3
    
    'create the main body surface
    Set newSurface = New cls2D_ObjectSurface
    
    With newSurface
      .TypeID = 1 'give it unique id
      
      .ImageWidth = 10 'initialize image size and source
      .ImageHeight = 30
      .ImagesPerRow = 4
      .SurfaceIndex = 3
      .SurfaceOffsetX = 83
      .SurfaceOffsetY = 93
      
      .DisplayOffsetX = -5
      .DisplayOffsetY = -15
      
      'set up the animation sequence
      .Set_AnimationSequence Array(0, 1, 2, 3)
    End With
    
    .Add_Surface newSurface
    .AddTo_Chain objectListStart
  End With
End Sub

Public Sub blowUpMissile(ByRef missileObject As cls2D_Object)
  With missileObject
    .MainAction = 2 'blow it up
    
    .Collision_Mask = 0 'don't process further collisions
    .CollisionBoxVisible = False
    
    With .SurfaceList
      .ImageWidth = 32
      .ImageHeight = 32
      .ImagesPerRow = 4
      .SurfaceOffsetX = 0
      .SurfaceOffsetY = 160
      .SurfaceIndex = 4
             
      .DisplayOffsetX = -16
      .DisplayOffsetY = -16
      
      .Set_AnimationSequence Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19)
    End With
  End With
End Sub

Public Sub blowUpBomber()
  Dim loop1 As Long, newSurface As cls2D_ObjectSurface
  
  'it would actually be more efficient to hold the bomber in an object for quick reference
  'since there is only one of them in this demo
  With objectListStart.Get_ChainObjectID(1)
    .MainAction = 2
    .Collision_Mask = 0 'don't process further collisions
    .CollisionBoxVisible = False
    
    For loop1 = 2 To 3 'turn off thrust
      .Get_SurfaceID(loop1).Visible = False
    Next loop1
    
    'create a number of surface explosions
    For loop1 = 1 To 6
      Set newSurface = New cls2D_ObjectSurface
      
      With newSurface
        .TypeID = loop1 + 10
        
        .ZPriority = 10 'make sure the explosions are on top
        
        If loop1 > 1 Then
          .Visible = False
          .MainAction = 2 'holding pattern
          .MainActionTickCounter = loop1 * 6 'each explosion occurs 6 frames apart
        Else
          .MainAction = 1
        End If
        
        If loop1 And 1 Then 'big and small explosions
          .ImageWidth = 64
          .ImageHeight = 64
          .ImagesPerRow = 4
          .SurfaceOffsetX = 126
          .SurfaceOffsetY = 0
          .SurfaceIndex = 4
          
          .DisplayOffsetX = 38 - Rnd() * 140
          .DisplayOffsetY = -19 - Rnd() * 26
        Else
          .ImageWidth = 32
          .ImageHeight = 32
          .ImagesPerRow = 4
          .SurfaceOffsetX = 0
          .SurfaceOffsetY = 160
          .SurfaceIndex = 4
                 
          .DisplayOffsetX = 54 - Rnd() * 140
          .DisplayOffsetY = -3 - Rnd() * 26
        End If
        
        .Set_AnimationSequence Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19)
      End With
      
      .Add_Surface newSurface
    Next loop1
  End With
End Sub
  
