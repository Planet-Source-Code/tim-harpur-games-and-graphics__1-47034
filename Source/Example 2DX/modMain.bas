Attribute VB_Name = "modMain"
Option Explicit

Public timeCycle As Long, masterMode As Long
Public backdropOffsetX As Single, backdropOffsetY As Single, backdropOffsetXL As Long, backdropOffsetYL As Long
Public rotorSpeed As Single, bodySpeed As Single, rSound As Single, missileSide As Long

Public objectListStart As cls2DX_Object

Public Sub Main()
  MsgBox "Use cursor keys to adjust rotor speeds. Space bar launches missile. Esc quits demo", vbOKOnly, "2DX Demo"
  
  'initialize variables
  masterMode = 0
  
  'initialize the timer - this will return instantly (0 and relative) with the current time + 16 milliseconds
  timeCycle = DXDraw.DelayTillTime(0, 0, True, True) + 16
  
  Do While masterMode > -1 'this is the main game control loop
    'timing is set for 16 milliseconds with a maximum carryover of 16 miiliseconds
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
  
  'clean up the 2DX chain
  objectListStart.Destroy_Chain
  
  'release sound and display
  DXSound.CleanUp_DXSound
  DXDraw.CleanUp_DXDraw
  
  'set the flag that it's OK for the main form to unload now, and then tell it to unload
  masterMode = -2
  Unload frmMain
End Sub

Private Sub InitializeScene()
  'load bitmaps used for graphics
  DXDraw.Init_StaticSurface 1, "2DX Backdrop.gif", 640, 480
  DXDraw.Init_TextureSurface 2, "2DX Heli.gif", 512, 256, -1
  
  DXSound.Init_DXSound frmMain, 3
  
  DXSound.Load_SoundBuffer 1, "2DX Heli.wav"
  DXSound.Load_SoundBuffer 2, "2DX Missile.wav"
  
  DXSound.Change_SoundSettings 2, 16000
  
  DXSound.Duplicate_SoundBuffer 3, 2
  
  createHeli
  
  backdropOffsetX = 0
  backdropOffsetY = 0
  
  missileSide = 28000
  
  masterMode = 3
End Sub

Private Sub UpdateScene()
  Dim sList As cls2DX_ObjectSurface, oList As cls2DX_Object
  
  'redraw the backdrop
  DXDraw.Display_TiledImage 1, , , , , backdropOffsetX, backdropOffsetY
  
  'process all the 2D objects in the 2DX object chain
  Set oList = objectListStart
  
  Do While Not (oList Is Nothing)
    With oList
      Set oList = .chainNext
      
      Select Case .TypeID
        Case 1 'process the helicopter
          Set sList = .SurfaceList
          
          If rotorSpeed > 0.15 Then .DisplayRotation = .DisplayRotation + bodySpeed
          
          If .DisplayRotation > 6.25 Then
            .DisplayRotation = .DisplayRotation - 6.25
          ElseIf .DisplayRotation < -6.25 Then
            .DisplayRotation = .DisplayRotation + 6.25
          End If
          
          'helicopter is being kept centre display always so move the backdrop offset and the recalculate
          'the helicopter's world position by placing it centre display
          'I haven't  bothered to place any boundary logic in the demo, so if it runs long enough
          'in the same direction it will get an overflow error (it would have to run a very very long time)
          If rotorSpeed > 0.15 Then
            backdropOffsetX = backdropOffsetX - Sin(.DisplayRotation) * (rotorSpeed - 0.14) * 20
            backdropOffsetY = backdropOffsetY - Cos(.DisplayRotation) * (rotorSpeed - 0.14) * 20
          End If
          
          'convert backdrop offset into object world coordinates (1000ths of a pixel)
          backdropOffsetXL = backdropOffsetX * 1000
          backdropOffsetYL = backdropOffsetY * 1000
          
          .PosX_1000ths = backdropOffsetXL + 400000
          .PosY_1000ths = backdropOffsetYL + 300000
          
          Do While Not (sList Is Nothing)
            With sList
              If .TypeID = 2 Then
                .DisplayRotation = .DisplayRotation + rotorSpeed
                
                If rSound < .DisplayRotation Then
                  rSound = .DisplayRotation + 1.2
                  
                  DXSound.Play_Sound 1
                End If
                
                If .DisplayRotation > 6.25 Then
                  .DisplayRotation = .DisplayRotation - 6.25
                  
                  rSound = rSound - 6.25
                End If
              End If
              
              Set sList = .NextSurface
            End With
          Loop
        Case 2 'process any missiles
          .PosX_1000ths = .PosX_1000ths - .speed_1000ths * Sin(.DisplayRotation)
          .PosY_1000ths = .PosY_1000ths - .speed_1000ths * Cos(.DisplayRotation)
          
          If .PosX_1000ths < backdropOffsetXL - 100000 Or .PosY_1000ths < backdropOffsetYL - 100000 Or .PosX_1000ths > backdropOffsetXL + 900000 Or .PosY_1000ths > backdropOffsetYL + 700000 Then
            .RemoveFrom_Chain objectListStart 'the missile has left the display area so remove it
            
            .Destroy_Object 'then destroy it so that all surface objects and references will be released
          Else
            With .Get_SurfaceID(2) 'process the smoke trail animation frame
              .ActionFrame = .ActionFrame + 1
              
              If .ActionFrame > .ActionLast Then .ActionFrame = 0
            End With
          End If
      End Select
    End With
  Loop
  
  'redraw the 2DX animation chain
  objectListStart.ReDraw_Chain backdropOffsetXL, backdropOffsetYL
  
  DXDraw.RefreshDisplay 'using full screen mode - this will flip the back render surface with display surface
End Sub

'create a helicopter object (this routine is only called once in this demo)
Public Sub createHeli()
  Dim newObject As cls2DX_Object, newSurface As cls2DX_ObjectSurface
  
  Set newObject = New cls2DX_Object
  
  With newObject
    .TypeID = 1
    
    .DisplayRotationEnable = True 'this object can be rotated
    
    'create the main body surface
    Set newSurface = New cls2DX_ObjectSurface
    
    With newSurface
      .TypeID = 1 'give it unique id
      
      .DisplayWidth = 77 'initialize image size and source
      .DisplayHeight = 219
      .ImagesPerRow = 1
      .SurfaceIndex = 2
      .SurfaceOffsetU = 257 / 512
      .SurfaceOffsetV = 0
      .ImageWidthU = 77 / 512
      .ImageHeightV = 219 / 256
      
      .SurfacePixelRatioU = 1 / 512
      .SurfacePixelRatioV = 1 / 256
      
      .DisplayOffsetX = -39
      .DisplayOffsetY = -63
    End With
    
    .Add_Surface newSurface
    
    'add top rotor
    Set newSurface = New cls2DX_ObjectSurface
    
    With newSurface
      .TypeID = 2 'give it unique id
      
      .ZPriority = 1 'must appear over other surfaces
      
      .DisplayWidth = 257 'initialize image size and source
      .DisplayHeight = 245
      
      .ImagesPerRow = 1
      .SurfaceIndex = 2
      .SurfaceOffsetU = 0
      .SurfaceOffsetV = 0
      .ImageWidthU = 257 / 512
      .ImageHeightV = 245 / 256
      
      .SurfacePixelRatioU = 1 / 512
      .SurfacePixelRatioV = 1 / 256
      
      .DisplayOffsetX = -129
      .DisplayOffsetY = -123
    End With
    
    .Add_Surface newSurface
    
    .AddTo_Chain objectListStart
  End With
End Sub

'create missile (this routine is called anytime the space is pressed)
Public Sub createMissile()
  Dim newObject As cls2DX_Object, newSurface As cls2DX_ObjectSurface
  Dim heliObject As cls2DX_Object
  
  'get the helicopter object - we need to access some of its values
  Set heliObject = objectListStart.Get_ChainObjectID(1)
  
  Set newObject = New cls2DX_Object
  
  With newObject
    .TypeID = 2
    
    .renderPriority = -1 'the missile should appear below the helicopter
    
    .DisplayRotationEnable = True 'this object can be rotated
    
    'position the missile under the helicopter - alternating left or right side
    'and play the launch sound
    missileSide = -missileSide
    
    If missileSide < 0 Then
      DXSound.Play_Sound 2
    Else
      DXSound.Play_Sound 3
    End If
    
    .PosX_1000ths = heliObject.PosX_1000ths - Cos(heliObject.DisplayRotation) * missileSide
    .PosY_1000ths = heliObject.PosY_1000ths + Sin(heliObject.DisplayRotation) * missileSide
    
    .speed_1000ths = 7000 'missile speed
    
    'set the missile direction to match the helicopter's current direction
    .DisplayRotation = heliObject.DisplayRotation
    
    'create the main body surface
    Set newSurface = New cls2DX_ObjectSurface
    
    With newSurface
      .TypeID = 1 'give it unique id
      
      .DisplayWidth = 17 'initialize image size and source
      .DisplayHeight = 57
      .ImagesPerRow = 1
      .SurfaceIndex = 2
      .SurfaceOffsetU = 336 / 512
      .SurfaceOffsetV = 0
      .ImageWidthU = 17 / 512
      .ImageHeightV = 57 / 256
      
      .SurfacePixelRatioU = 1 / 512
      .SurfacePixelRatioV = 1 / 256
      
      .DisplayOffsetX = -8
      .DisplayOffsetY = -28
    End With
    
    .Add_Surface newSurface
    
    'add smoke trail
    Set newSurface = New cls2DX_ObjectSurface
    
    With newSurface
      .TypeID = 2 'give it unique id
      
      .ZPriority = 1 'must appear over other surfaces of this object
      
      .DisplayWidth = 15 'initialize image size and source
      .DisplayHeight = 81
      
      .ImagesPerRow = 4
      .SurfaceIndex = 2
      .SurfaceOffsetU = 354 / 512
      .SurfaceOffsetV = 0
      .ImageWidthU = 15 / 512
      .ImageHeightV = 81 / 256
      
      .SurfacePixelRatioU = 1 / 512
      .SurfacePixelRatioV = 1 / 256
      
      .DisplayOffsetX = -6
      .DisplayOffsetY = 29
      
      .Set_AnimationSequence Array(0, 1, 2, 3) 'smoke trail has 4 frames of animation
      
      .SurfaceColour = &H40FFFFFF 'smoke trail is 75% translucent
    End With
    
    .Add_Surface newSurface
    
    .AddTo_Chain objectListStart
  End With
End Sub

