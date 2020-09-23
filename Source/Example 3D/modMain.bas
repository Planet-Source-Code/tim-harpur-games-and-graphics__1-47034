Attribute VB_Name = "modMain"
Option Explicit
Option Base 0

'game mode flags and timer
Public masterMode As Long, timeCycle As Long, lastKey As Long

'demo mode flags
Public textureMode As Boolean, filterMode As Boolean, shadeMode As Boolean
Public transMode As Single, transparencyMode As Boolean, colourMode As Long
Public displayMode As Long, useSoftware As Boolean, windowMode As Boolean
Public animationMode As Boolean, forceTexBDepth As Long, colBoxes As Boolean

'3D render list
Public objectListStart As cls3D_Object

'camera position
Public cameraRotation As Single, cameraTilt As Single, zoomFactor As Single

'planet's rotation
Public planetRotation As Single

'the main game routine - this is the start point and also central displatch
Public Sub Main()
  'initialize variables
  zoomFactor = 1000

  textureMode = True 'enable textures
  shadeMode = True 'enable gourad shading
  transMode = 1#  'all objects intially opaque
  transparencyMode = True 'turn on transparency
  displayMode = 2 '800x600
  colourMode = 1 '16bit colour
  windowMode = True 'start in window
  
  'initialize the timer - this will return instantly (0 and relative) with the current time + 16 milliseconds
  timeCycle = DXDraw.DelayTillTime(0, 0, True) + 16
  
  Do While masterMode > -1 'this is the main game control loop
    'timing is set for 16 milliseconds with a maximum carryover of 16 miiliseconds
    timeCycle = DXDraw.DelayTillTime(timeCycle, 16) + 16
    
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
  
  'clean up the 3D render list
  objectListStart.Destroy_Chain
  
  'release input and display
  DXInput.CleanUp_DXInput
  DXDraw.CleanUp_DXDraw
  
  'set the flag that it's OK for the main form to unload now, and then tell it to unload
  masterMode = -2
  Unload frmMain
End Sub

'this routine initializes the 3D scene
Public Sub InitializeScene()
  Dim tobj As cls3D_Object, loop1 As Long
  
  'set up some of the 3D world parameters - others are using defaults until changed
  DXDraw.Set_D3DProjection 1, 40000
  DXDraw.Set_D3DWorldAmbient 0.35, 0.35, 0.35
  
  'clear the old render list
  If Not (objectListStart Is Nothing) Then objectListStart.Destroy_Chain
  Set objectListStart = Nothing
  
  'add one planet, a starfield and a terlaxian cruiser
  Planet
  StarField
  TerlaxianCruiser
  
  'now add a few extra Terlaxian Cruisers by duplication, to create a fleet
  For loop1 = 2 To 7
    Set tobj = objectListStart.Get_ChainObjectName("TerlaxianCruiser", 1).Duplicate_Object("TerlaxianCruiser", loop1)
  
    With tobj
      .PrimaryAxisRotation = (Pi / 4) * loop1
      
      .BasePosY = -400 + loop1 * 100
      
      If loop1 And 1 Then 'have every other cruiser going in opposite direction
        .BasePosX = 700 * Cos(.PrimaryAxisRotation)
        .BasePosZ = 700 * Sin(.PrimaryAxisRotation) + 50 * loop1 - 50
      Else
        .BasePosX = -700 * Cos(.PrimaryAxisRotation)
        .BasePosZ = -700 * Sin(.PrimaryAxisRotation) + 50 * loop1 - 50
      End If
      
      'the object is done being created so calculate its vertices
      .Calculate_Vertices
      
      'although this object may move in the future, it starts off stopped so update its starting position
      .Update_ObjectPosition
      .AddTo_Chain objectListStart, False
    End With
  Next loop1

  masterMode = 3
End Sub

Public Sub UpdateScene()
  Dim loop1 As Long
  
  'see if any keys have been pressed, and respond
  checkKeyPress
  
  'if the 3D could not be initialized then don't try to render the scene
  If DXDraw.TestDisplay3DValid() Then
    'check for mouse input and adjust camera accordingly
    DXInput.Poll_Mouse
    
    cameraRotation = cameraRotation + DXInput.dx_MouseState.X / 100
    
    Do While cameraRotation > 2 * Pi
      cameraRotation = cameraRotation - 2 * Pi
    Loop
    
    Do While cameraRotation < 0
      cameraRotation = cameraRotation + 2 * Pi
    Loop
    
    cameraTilt = cameraTilt + DXInput.dx_MouseState.Y / 100
    
    If cameraTilt >= Pi / 2 Then cameraTilt = Pi / 2 - 0.0001
    If cameraTilt <= -Pi / 2 Then cameraTilt = -Pi / 2 + 0.0001
    
    If DXInput.dx_MouseState.buttons(0) Then zoomFactor = zoomFactor - 5
    If DXInput.dx_MouseState.buttons(1) Then zoomFactor = zoomFactor + 5
    
    If zoomFactor < -1000 Then zoomFactor = -1000
    If zoomFactor > 2000 Then zoomFactor = 2000
    
    'update the camera's position
    DXDraw.Set_D3DCamera -Sin(cameraRotation) * Cos(cameraTilt) * zoomFactor, -Sin(cameraTilt) * zoomFactor, Cos(cameraRotation) * Cos(cameraTilt) * -zoomFactor, cameraRotation, cameraTilt
    
    'if the demo animation has been activated - move the cruisers
    If animationMode Then
      planetRotation = planetRotation + 0.01
      If planetRotation > 1 Then planetRotation = planetRotation - 1
      
      'there are two ways to rotate the planet - one way is to scroll the surface texture - it is not
      'quite as visually appealling but it is faster than the alternative, which is to rotate
      'the whole planet and recalulate the lighting
      'of course - if the lighting was applied evenly to the whole planet then it could be rotated
      'without needing to have its vertices re-calculates - this would fast but not as visually
      'appealing as the planet's lighting would be 'flat'
      With objectListStart.Get_ChainObjectName("Planet")
        .Set_TextureAxisY 80, Pi / 4, , planetRotation
        .Calculate_VerticesUV
        
        '.PrimaryAxisRotation = planetRotation
        '.Apply_VertexAmbient -100, -100, -100, 300, 0.9, 1, 0.8, 1, True, True
        '.Calculate_VerticesColour
        '.Update_ObjectPosition
      End With
      
      For loop1 = 1 To 7
        With objectListStart.Get_ChainObjectName("TerlaxianCruiser", loop1)
          'for the purposes of this demo I just set up a simple circular orbit for each cruiser
          If loop1 And 1 Then 'have every other cruiser going in opposite direction
            .PrimaryAxisRotation = .PrimaryAxisRotation + 0.0005 * loop1
            
            If .PrimaryAxisRotation > 2 * Pi Then .PrimaryAxisRotation = .PrimaryAxisRotation - 2 * Pi
          
            .BasePosX = 700 * Cos(.PrimaryAxisRotation)
            .BasePosZ = 700 * Sin(.PrimaryAxisRotation) + 50 * loop1 - 50
          Else
            .PrimaryAxisRotation = .PrimaryAxisRotation - 0.0005 * loop1
            
            If .PrimaryAxisRotation < 0 Then .PrimaryAxisRotation = .PrimaryAxisRotation + 2 * Pi
            
            .BasePosX = -700 * Cos(.PrimaryAxisRotation)
            .BasePosZ = -700 * Sin(.PrimaryAxisRotation) + 50 * loop1 - 50
          End If
          
          'when an object is moved it needs to have its position updated
          .Update_ObjectPosition
        End With
      Next loop1
    End If
    
    'render 3D scene
    DXDraw.Clear_Display
    DXDraw.Begin_D3DScene
    
    If transMode = 1# Then
      objectListStart.ReDraw_Chain transparencyMode, False
    Else
      objectListStart.ReDraw_Chain transparencyMode
    End If
    
    DXDraw.End_D3DScene
  Else
    DXDraw.Clear_Display
    
    DXDraw.Draw_Text "Requested 3D Display Mode is Not Available", 300, 220, 65535
  End If
  
  'regular 2D graphics can be added to the display after the 3D scene has been rendered
  DXDraw.Draw_Text "Press Esc to end this demo", 0, 100, &HFFFFFF
  DXDraw.Draw_Text "Press A to toggle animation " & animationMode, 0, 125, &HFFFFFF
  DXDraw.Draw_Text "Press E to toggle engines", 0, 150, &HFFFFFF
  DXDraw.Draw_Text "Press G/H to adjust opacity down/up (ghosting) " & Format(transMode, "0.0"), 0, 175, &HFFFFFF
  DXDraw.Draw_Text "Press S to toggle shade mode " & shadeMode, 0, 200, &HFFFFFF
  DXDraw.Draw_Text "Press T to toggle textures " & textureMode, 0, 225, &HFFFFFF
  DXDraw.Draw_Text "Press B to toggle starfield backdrop", 0, 250, &HFFFFFF
  DXDraw.Draw_Text "Press P to toggle transparency " & transparencyMode, 0, 275, &HFFFFFF
  DXDraw.Draw_Text "Press F to toggle filter mode " & filterMode, 0, 300, &HFFFFFF
  DXDraw.Draw_Text "Press D to toggle forced texture bit depth (0/16/32) " & forceTexBDepth, 0, 325, &HFFFFFF
  DXDraw.Draw_Text "Press F1/F2 to toggle window/full screen " & Not windowMode, 0, 350, &HFFFFFF
  DXDraw.Draw_Text "Press F3/F4 to toggle software/hardware mode " & Not useSoftware, 0, 375, &HFFFFFF
  DXDraw.Draw_Text "Press F5-F7 to set display resolution 640/800/1024 (" & displayMode & ")", 0, 400, &HFFFFFF
  DXDraw.Draw_Text "Press F8/F9 to toggle 16/32 bit colour " & colourMode * 16, 0, 425, &HFFFFFF
  DXDraw.Draw_Text "Use the mouse and left/right buttons to navigate", 0, 450, &HFFFFFF
  
  DXDraw.RefreshDisplay 'the display surface is updated
End Sub

'This routine creates and adds a planet to the 3D render chain
Public Sub Planet()
  Dim pObject As cls3D_Object
  
  'start with a new 3D object
  Set pObject = New cls3D_Object
  
  With pObject
    .objectName = "Planet"
    .TypeID = 1000
    .primaryAxis = 1
    
    .Create_Sphere 16, 6, 50, 50, 50, 1, 1, 1, , 6
    
    'the sphere primitive is made up of surface bands - since they all use the same texture and
    'don't need to be accessed independently they may as well be joined - also this improves
    'the lighting effects as edges between bands is softened
    .Join_AllSurfaces
    
    'recalculate the normals for the sphere using adjacent normal blending - needs to be done on
    'any surface that has been changed and is going to apply lighting
    .Calculate_Normals True
    
    'do to the design nature of the sphere primitive we want the "peaks" to be on the N/S poles
    '-turn of textures and you'll see what I mean
    'just a reminder, the 'Flip...' routines flip the normals also, so they don't need recalculating
    .Flip_xAxis90
    
    'increase planet's scale by 4
    .Scale_Object 4, 4, 4
    
    'apply texture wrapping around Y axis
    .Set_TextureAxisY 80, Pi / 4, , planetRotation
    
    'apply lighting effects to the planet
    .Apply_VertexAmbient -100, -100, -100, 300, 0.9, 1, 0.8, 1, True, True
    
    'since the object was just created it needs to have it surface vertices calculated
    .Calculate_Vertices
    
    'this object doesn't move so it only needs to have its position updated once - when it is created
    .Update_ObjectPosition
    
    'add the planet to the 3D render chain - this is the first object in the chain
    .AddTo_Chain objectListStart
  End With
End Sub

'This routine adds a 3D object used for a starfield backdrop to the 3D Render Chain
Public Sub StarField()
  Dim bdropObject As cls3D_Object
  
  'create a new 3D object to work with
  Set bdropObject = New cls3D_Object
  
  With bdropObject
    'this is a very big object so that the stars appear to be at great distance
    'notice how small the particle density is - be careful when creating large particle clouds
    'to keep the density down or you may wind up with a very large number of particles
    .Create_Particles 20000, 20000, 20000, 0.0000000001, 1, 1, 1, 0.15, 0.15, 0.5, 0.5, 0.5, 4
    
    .objectName = "StarField"
    .TypeID = 1001
    
    'initially the starfield is turned off
    .Visible = False
    
    'since the object was just created it needs to have it surface vertices calculated
    .Calculate_Vertices
    'this object doesn't move so it only needs to have its position updated once - when it is created
    .Update_ObjectPosition
    
    'add this object to the 3D rendering chain
    .AddTo_Chain objectListStart
  End With
End Sub

'This routine adds a 3D object called a TerlaxianCruiser to the 3D Render Chain
Public Sub TerlaxianCruiser()
  Dim mObject As cls3D_Object, sObject As cls3D_ObjectSurface
  Dim aObject As cls3D_Object, bObject As cls3D_Object
  
  Set mObject = New cls3D_Object
  
  'There are 2 ways to initialize a 3D object
  'The first way is to create it from scratch (shown)
  'The second way is to load its definition from file (shown but commented out)
  'In order to use this second method the object must have first been created and saved to file
  'but once created it makes for FAR smaller code
  
  '-----------------------------------------------------------------------------------
  '  With mObject
  '    .Load_FromFile App.Path & "\TerlaxianCruiser.3DObj"
  
  '    .Calculate_Vertices
  '    .Update_3DObjectPosition
  '  End With
  
  '  .AddTo_3DChain objectListStart
  
  '  Exit Sub
  '-----------------------------------------------------------------------------------
  
  With mObject 'main body/hull
    .objectName = "TerlaxianCruiser"
    .TypeID = 1
    
    .Create_Box 50, 20, 80, 0.5, 0.5, , -5, 0.7, 0.85, 0.9, , 2
    .primaryAxis = 1 'y axis is primary axis
    
    .BasePosX = 700 'start cruiser at 700 world units along x axis
    .BasePosY = -300
    
    'apply texture scaling to different faces of hull
    .Get_SurfaceID(1).Set_TexturePlaneXZ 10, 10, , , , False
    .Get_SurfaceID(2).Set_TexturePlaneXZ -10, 10, , , , False
    
    .Get_SurfaceID(3).Set_TexturePlaneYZ 10, -10, , , True, False
    .Get_SurfaceID(4).Set_TexturePlaneYZ -10, -10, , , True, False
    
    .Remove_Surface .Get_SurfaceID(5)
    .Get_SurfaceID(6).Set_TexturePlaneXY 10, 10, , , , False
    
    'improve lighting quality by doubling the number of vetices on the hull
    .Double_ObjectComplexity
    
    Set aObject = New cls3D_Object
    
    With aObject 'add a front nose piece to reduce the blocky appearance
      .Create_Wedge 10, 2, 25, , , , , , 0.7, 0.85, 0.9, , 2
      
      .Remove_Surface .Get_SurfaceID(1)
      
      .Get_SurfaceID(2).Set_TextureAxisZ 10, Pi, , 0.2, True
      .Get_SurfaceID(3).Set_TextureAxisZ 10, Pi, , 0.2, True
      .Get_SurfaceID(4).Set_TexturePlaneXY 10, 10, , , True, False
      .Get_SurfaceID(5).Set_TexturePlaneXY 10, 10, , , True, False
      
      'these values affect how/where this object's surfaces are attached to the hull (below)
      .BasePosY = -5
      .BasePosZ = 41
      
      .PrimaryAxisTilt = Pi / 2
      .PrimaryAxisPrecession = -Pi / 2
    End With
    
    'use the front nose object as a template and add its the surfaces to the hull object
    .Add_SurfacesFromTemplate aObject
    
    'join all hull object's surfaces together since they all use the same texture and surface properties
    .Join_AllSurfaces
    
    Set aObject = New cls3D_Object
    
    With aObject 'create some fins to add onto the hull
      .Create_TriPane 30, 10, -20, 0.8, 0.85, 0.9, 2, 2, True, 0.1
      .primaryAxis = 0
      
      .Get_SurfaceID(1).Set_TexturePlaneXY 10, -10, , , True, False
      .Get_SurfaceID(2).Set_TexturePlaneXY -10, -10, , , True, False
      
      'these values affect how/where this object's surfaces are attached to the hull (below)
      .BasePosX = -20
      .BasePosY = 7
      .BasePosZ = -20
      
      .PrimaryAxisRotation = -Pi / 5.3
      .PrimaryAxisTilt = Pi / 2.2
      .PrimaryAxisPrecession = -Pi / 24
    End With

    'use the fin object as a template and add its the surfaces to the hull object
    .Add_SurfacesFromTemplate aObject
    
    'flip the fin object - this will be a top fin
    aObject.Mirror_yzPlane True
    
    'use the fin object as a template and add its the surfaces to the hull object
    .Add_SurfacesFromTemplate aObject
    
    'flip the fin object and make some adjustments to its position/angle
    'this will be a bottom fin
    With aObject
      .Mirror_xzPlane True
      
      .BasePosY = -5
      .PrimaryAxisRotation = Pi / 25
    End With
    
    'use the fin object as a template and add its the surfaces to the hull object
    .Add_SurfacesFromTemplate aObject
    
    'flip the fin object - this will be a bottom fin
    aObject.Mirror_yzPlane True
    
    'use the fin object as a template and add its the surfaces to the hull object
    .Add_SurfacesFromTemplate aObject
    
    Set aObject = New cls3D_Object
    
    With aObject 'create the main bridge
      .Create_Box 20, 4, 15, 0.5, , , , 0.8, 0.8, 0.8, , 2
      
      'rember that the primaryAxis is the x axis by default so this will tilt the bridge up
      .PrimaryAxisRotation = 0.1
      
      .BasePosY = 7
      .BasePosZ = -20
      
      .Remove_Surface .Get_SurfaceID(1)
      .Remove_Surface .Get_SurfaceID(6)
      
      With .Get_SurfaceID(5)
        .TextureSurfaceIndex = 3
        .Set_TexturePlaneXY 1, 1, -0.2, , True
        .SurfaceEmissiveRed = 0.5
        .SurfaceEmissiveGreen = 0.5
        .SurfaceEmissiveBlue = 0.5
      End With
      
      With .Get_SurfaceID(3)
        .TextureSurfaceIndex = 3
        .Set_TexturePlaneYZ 1, 1, -0.2
        .SurfaceEmissiveRed = 0.5
        .SurfaceEmissiveGreen = 0.5
        .SurfaceEmissiveBlue = 0.5
      End With
      
      With .Get_SurfaceID(4)
        .TextureSurfaceIndex = 3
        .Set_TexturePlaneYZ 1, 1, -0.2
        .SurfaceEmissiveRed = 0.5
        .SurfaceEmissiveGreen = 0.5
        .SurfaceEmissiveBlue = 0.5
      End With
    End With
    
    .Add_SurfacesFromTemplate aObject
    
    With aObject 'create 2 secondary bridges using the main bridge template as a starting point
      .Scale_Object 0.5, 0.5, 0.5
      
      .BasePosX = -12
      .BasePosY = 4.5
      .BasePosZ = 0
      
      .PrimaryAxisRotation = 0.2
    End With
    
    .Add_SurfacesFromTemplate aObject
    
    aObject.BasePosX = 12
    .Add_SurfacesFromTemplate aObject
    
    With aObject
      .BasePosX = 0
      .BasePosY = 9
      .BasePosZ = -20
      
      With .Get_SurfaceID(2)
        .SurfaceRed = 0.6
        .SurfaceGreen = 0.6
        .SurfaceBlue = 0.6
      End With
    End With
    
    .Add_SurfacesFromTemplate aObject
    
    Set aObject = New cls3D_Object
    
    With aObject 'create 3 engines - create the template first
      .Create_Rod 6, 5, 5, 4, 0.75, 0.75, , , -2, , 0.2, 0.2, 0.1
      
      .BasePosZ = -42
      
      .Remove_Surface .Get_SurfaceID(6)
      
      With .Get_SurfaceID(4)
        .TextureSurfaceIndex = 5
        .Set_TexturePlaneXY
        
        .SurfaceVertexArrayRed = Array(0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8)
        .SurfaceVertexArrayGreen = Array(0.6, 0, 0, 0.6, 0, 0, 0.6, 0, 0, 0.6, 0, 0, 0.6, 0, 0, 0.6, 0, 0, 0.6, 0, 0)
      End With
    End With
    
    .Add_SurfacesFromTemplate aObject
    
    aObject.BasePosX = -16
    .Add_SurfacesFromTemplate aObject
    
    aObject.BasePosX = 16
    .Add_SurfacesFromTemplate aObject
    
    Set aObject = New cls3D_Object
    
    With aObject 'side cannon
      .Create_Rod 10, 2, 3, 60, 1, 1, , , , 3, , , , , 4
      
      .BasePosX = -20
      .BasePosY = -5
      
      With .Get_SurfaceID(6) 'front glowing piece
        .SurfaceEmissiveRed = 0.1
        .SurfaceEmissiveGreen = 0.2
        .SurfaceEmissiveBlue = 0.3
        .TextureSurfaceIndex = 5
        .Set_TexturePlaneXY
      End With
      
      .Get_SurfaceID(2).Set_TextureAxisZ 2, Pi * 1.5, , 0.42, True
      
      .Remove_Surface .Get_SurfaceID(4)
    End With
    
    .Add_SurfacesFromTemplate aObject
    
    aObject.Mirror_yzPlane True 'create a second cannon just like the first - just mirrored to other side
    .Add_SurfacesFromTemplate aObject
    
    With aObject 'use the side cannon as a starting point for 4 small top cannons
      .Scale_Object 0.2, 0.2, 0.1
      
      .BasePosX = -10
      .BasePosY = 3
      .BasePosZ = 25
      
      .PrimaryAxisRotation = 0.4
    End With
    
    .Add_SurfacesFromTemplate aObject
    
    aObject.BasePosX = -8
    .Add_SurfacesFromTemplate aObject
    
    aObject.Mirror_yzPlane True
    .Add_SurfacesFromTemplate aObject
    
    aObject.BasePosX = 10
    .Add_SurfacesFromTemplate aObject
    
    Set aObject = New cls3D_Object
    
    With aObject 'engine flange
      .Create_Box 50, 20, 4, 0.9, 0.9, , , 0.7, 0.6, 0.5, 2, 2, , True
      
      'this flange is going to have lighting effects applied when engines turned off/on so we want
      'a reasonable quality to the object for this to look good
      .Double_ObjectComplexity
      .Double_ObjectComplexity
      
      .Apply_VertexAmbient 0, 0, -32, 35, 0, 0.4, 0.3, 0.2, False, False
      
      'object needed minor adjusting
      .Scale_Object 0.9, 0.8, 1
      
      'place flange on tail of cruiser hull
      .BasePosZ = -42.2
      
      'apply texturing to each surface
      .Get_SurfaceID(1).Set_TexturePlaneXZ 20, 20, , , , False
      .Get_SurfaceID(2).Set_TexturePlaneXZ -20, 20, , , , False
      .Get_SurfaceID(3).Set_TexturePlaneYZ 20, -20, , , True, False
      .Get_SurfaceID(4).Set_TexturePlaneYZ -20, -20, , , True, False
      
      .Get_SurfaceID(7).Set_TexturePlaneXZ 20, 20, , , , False
      .Get_SurfaceID(8).Set_TexturePlaneXZ -20, 20, , , , False
      .Get_SurfaceID(9).Set_TexturePlaneYZ 20, -20, , , True, False
      .Get_SurfaceID(10).Set_TexturePlaneYZ -20, -20, , , True, False
      
      .Get_SurfaceID(11).Set_TexturePlaneXY -20, -20, , , , False
      
      'remove those surfaces from the flange that are not desired
      .Remove_Surface .Get_SurfaceID(5)
      .Remove_Surface .Get_SurfaceID(6)
      .Remove_Surface .Get_SurfaceID(12)
      
      'turn the flange into one big surface - ID 0
      .Join_AllSurfaces 0
    End With
    
    'give the flange a unique surface ID so it can be manipulated later - when thrust is activated this
    'flange has lighting effects applied to it - its ID is 0 + 5000 = 5000
    .Add_SurfacesFromTemplate aObject, 5000
    
    Set aObject = New cls3D_Object
    
    With aObject 'engine thrust
      .Create_Panel 15, 8, 1, 1, 1, 7, 7, True, 0.6
      .Get_SurfaceID(1).Set_TexturePlaneXY -1
      
      Set bObject = .Duplicate_Object("", 0) 'create a second 'engine thrust' at 90 degrees to first
      bObject.PrimaryAxisRotation = Pi / 2
      .Add_SurfacesFromTemplate bObject
      
      Set bObject = .Duplicate_Object("", 0)
      'create a third flat 'engine thrust' at 90 degrees to first two, that only need to be one sided
      With bObject
        .Create_Panel 7, 7, 1, 1, 1, 7, , , 0.6
        
        .PrimaryAxisRotation = Pi
        .PrimaryAxisTilt = Pi / 2
        
        .BasePosX = 5
      End With
      
      .Add_SurfacesFromTemplate bObject
      
      'the whole object needs to be rotated 90 degrees around Y axis and moved into position
      'before being joined with the hull object
      .primaryAxis = 1
      .PrimaryAxisRotation = Pi / 2
      
      .BasePosZ = -49
    End With
    
    Set bObject = New cls3D_Object
    
    'use the engine thrust object just created and create a thrust object that has 3 exhausts
    With bObject
      .objectName = "Thrust"
      
      .Visible = False
      
      .Add_SurfacesFromTemplate aObject
      
      aObject.BasePosX = -16
      .Add_SurfacesFromTemplate aObject
      
      aObject.BasePosX = 16
      .Add_SurfacesFromTemplate aObject
      
      .Join_AllSurfaces
    End With
    
    'add the 'Thrust' to the hull object as an attached subordinate 3D object - this way it moves
    'with the hull, but can still be manipulated independently
    .Add_Attached3DObject bObject
    
    'the object is done being created so calculate its vertices
    .Set_Opacity transMode
    .Calculate_Vertices
    
    'although this object may move in the future, it starts off stopped so update its starting position
    .Update_ObjectPosition
    
    'The completed object description can be saved for future use without needing to reuse all
    'this code
    '.Save_ToFile App.Path & "\TerlaxianCruiser.3DObj"
    
    .AddTo_Chain objectListStart, False
  End With
End Sub

Private Sub checkKeyPress()
  Dim loop1 As Long
  
  'check the keyboard for user input
  DXInput.Poll_Keyboard
  
  'if the user pressed Esc then signal that the program is to shut down by setting the game's
  'master mode flag to -1 and abort the sub
  If DXInput.dx_KeyboardState.Key(DIK_ESCAPE) <> 0 Then
    masterMode = -1
    
    Exit Sub
  ElseIf DXInput.dx_KeyboardState.Key(DIK_T) <> 0 Then 'texture mode toggle
    If lastKey <> DIK_T Then
      textureMode = Not textureMode
      
      DXDraw.D3DTextureEnable = textureMode
      
      lastKey = DIK_T
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_F) <> 0 Then 'bilinear filter mode toggle
    If lastKey <> DIK_F Then
      filterMode = Not filterMode
      
      DXDraw.Set_D3DFilterMode filterMode
      
      lastKey = DIK_F
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_S) <> 0 Then 'shade mode toggle
    If lastKey <> DIK_S Then
      shadeMode = Not shadeMode
      
      DXDraw.Set_D3DShadeMode shadeMode
      
      lastKey = DIK_S
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_G) <> 0 Then 'decrease opacity
    If lastKey <> DIK_G Then
      transMode = transMode - 0.1
     If transMode < 0 Then transMode = 0
      
      For loop1 = 1 To 7
        With objectListStart.Get_ChainObjectName("TerlaxianCruiser", loop1)
          .Set_Opacity transMode
          .Calculate_VerticesColour
        End With
      Next loop1
      
      lastKey = DIK_G
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_H) <> 0 Then 'increase opacity
    If lastKey <> DIK_H Then
      transMode = transMode + 0.1
      If transMode > 1 Then transMode = 1
      
      For loop1 = 1 To 7
        With objectListStart.Get_ChainObjectName("TerlaxianCruiser", loop1)
          .Set_Opacity transMode
          .Calculate_VerticesColour
        End With
      Next loop1
      
      lastKey = DIK_H
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_P) <> 0 Then 'transparency toggle
    If lastKey <> DIK_P Then
      transparencyMode = Not transparencyMode
      
      lastKey = DIK_P
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_E) <> 0 Then 'engine toggle
    If lastKey <> DIK_E Then
      Dim tobj As cls3D_Object
      
      For loop1 = 1 To 7
        Set tobj = objectListStart.Get_ChainObjectName("TerlaxianCruiser", loop1)
        
        With tobj.Get_Attached3DObjectName("Thrust")
          If .Visible Then
            .Visible = False
            
            With tobj.Get_SurfaceID(5000)
              .SurfaceEmissiveRed = 0
              .SurfaceEmissiveGreen = 0
              
              .Calculate_VerticesColour
            End With
          Else
            .Visible = True
                 
            With tobj.Get_SurfaceID(5000)
              .SurfaceEmissiveRed = 0.15
              .SurfaceEmissiveGreen = 0.08
            
              .Calculate_VerticesColour
            End With
          End If
        End With
      Next loop1
      
      lastKey = DIK_E
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_B) <> 0 Then 'backdrop toggle
    If lastKey <> DIK_B Then
      With objectListStart.Get_ChainObjectName("StarField")
        .Visible = Not .Visible
      End With
      
      lastKey = DIK_B
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_A) <> 0 Then 'animation mode toggle
    If lastKey <> DIK_A Then
      animationMode = Not animationMode
      
      lastKey = DIK_A
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_D) <> 0 Then ' toggle forced texture bit depth
    If lastKey <> DIK_D Then
      Select Case forceTexBDepth
        Case 0
          forceTexBDepth = 16
        Case 16
          forceTexBDepth = 32
        Case Else
          forceTexBDepth = 0
      End Select
      
      masterMode = 1
      
      lastKey = DIK_D
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_F1) <> 0 Then ' toggle window mode
    If lastKey <> DIK_F1 Then
      windowMode = True
      
      masterMode = 1
      
      lastKey = DIK_F1
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_F2) <> 0 Then ' toggle full screen mode
    If lastKey <> DIK_F2 Then
      windowMode = False
      
      masterMode = 1
      
      lastKey = DIK_F2
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_F3) <> 0 Then ' toggle software mode
    If lastKey <> DIK_F3 Then
      useSoftware = True
      
      masterMode = 1
      
      lastKey = DIK_F3
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_F4) <> 0 Then ' toggle hardware mode
    If lastKey <> DIK_F4 Then
      useSoftware = False
      
      masterMode = 1
      
      lastKey = DIK_F4
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_F5) <> 0 Then ' toggle dispaly mode 640x480
    If lastKey <> DIK_F5 Then
      displayMode = 1
      
      masterMode = 1
      
      lastKey = DIK_F5
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_F6) <> 0 Then ' toggle dispaly mode 800x600
    If lastKey <> DIK_F6 Then
      displayMode = 2
      
      masterMode = 1
      
      lastKey = DIK_F6
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_F7) <> 0 Then ' toggle dispaly mode 1024x768
    If lastKey <> DIK_F7 Then
      displayMode = 3
      
      masterMode = 1
      
      lastKey = DIK_F7
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_F8) <> 0 Then ' toggle colour mode 16bit
    If lastKey <> DIK_F8 Then
      colourMode = 1
      
      masterMode = 1
      
      lastKey = DIK_F8
    End If
  ElseIf DXInput.dx_KeyboardState.Key(DIK_F9) <> 0 Then ' toggle colour mode 32bit
    If lastKey <> DIK_F9 Then
      colourMode = 2
      
      masterMode = 1
      
      lastKey = DIK_F9
    End If
  Else
    lastKey = 0
  End If
End Sub
