VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls2D_Object"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************************
'
' cls2D_Object class definition
'                                                     - written by Tim Harpur for Logicon Enterprises
'
' Don't forget to add the appropriate Project->Reference to the DirectX7 library
' User defined values can be used as needed - they have no effect on any DXDraw routines
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
Option Base 0

'***************************************************************************************************************
'
' The following section contains the core member variables for this class - do not remove.
'
'***************************************************************************************************************

Public objectName As String                                 ' object's name

Public Visible As Boolean                                         ' object's master visibility flag (default TRUE)

Public chainPrevious As cls2D_Object       ' previous object in chain
Public chainNext As cls2D_Object              ' next object in chain

Public renderPriority As Long                                ' object's priority - user defined

Public SurfaceList As cls2D_ObjectSurface                    ' list of object's image surfaces
  
Public TypeID As Long                                             ' object ID value
  
Public displayPosX As Long           ' object's X & Y display position - these values are automatically
Public displayPosY As Long           ' loaded by a call to Translate_Object or ReDraw_Chain
  
Public PosX_1000ths As Long                                       ' object's X world position
Public PosY_1000ths As Long                                       ' object's Y world position

Public Collision_Mask As Long                           ' object's collision bit mask
Public Collision_WidthX_1000ths As Long                   ' X width of collision box
Public Collision_WidthY_1000ths As Long                   ' Y width of collision box

Public Collision_XLower As Long, Collision_XUpper As Long         'calculated collision bounds
Public Collision_YLower As Long, Collision_YUpper As Long         'calculated collision bounds

Public CollisionBoxColour As Long                  ' line colour for drawing collision box
Public CollisionBoxVisible As Boolean               'indicates if this object's collision box is visible

'***************************************************************************************************************
'
' The following section can be used to add custom member variables to this class.
' It is stongly advised that a COPY of this file be placed in the project directory, and that only the copy
' be customized and used. In this way the original class definition is preserved, while ensuring that each
' project has its own customized class definition that is not at risk of being modified by another project.
'
' Be certain to make changes to the routines Load_FromOpenFile, Save_ToOpenFile, and
' Duplicate_Object to reflect any changes made to the member variables here.
'
'***************************************************************************************************************

Public Health As Long                                              ' user defined value
Public Damage As Long                                          ' user defined value
Public Level As Long                                                ' user defined value
Public Experience As Long                                      ' user defined value
Public LifeSpan As Long                                          ' user defined value
Public Age As Long                                                  ' user defined value
Public value As Long                                               ' user defined value

Public MainAction As Long                                       ' user defined value (can be used as master)
Public TicksPerMainAction As Long                         ' user defined value (can be used as master)
Public MainActionTickCounter As Long                   ' user defined value (can be used as master)
Public SubAction As Long                                        ' user defined value (can be used as master)
Public TicksPerSubAction As Long                         ' user defined value (can be used as master)
Public SubActionTickCounter As Long                    ' user defined value (can be used as master)
  
Public TicksPerFrame As Long                               ' user defined value (can be used as master)
Public FrameTickCounter As Long                          ' user defined value (can be used as master)
  
Public direction As Long                                             ' user defined value
Public Range_1000ths As Long                                  ' user defined value
Public speed_1000ths As Long                                  ' user defined value

Public VelocityX_1000ths As Long                                  ' user defined value
Public VelocityY_1000ths As Long                                  ' user defined value
Public AccelerationX_1000ths As Long                          ' user defined value
Public AccelerationY_1000ths As Long                          ' user defined value
  
Public TargetMode As Long                                           ' user defined value
Public TargetX_1000ths As Long                                  ' user defined value
Public TargetY_1000ths As Long                                  ' user defined value

Public User1 As Long                                                  ' user defined value
Public User2 As Long                                                  ' user defined value
Public User3 As Long                                                  ' user defined value

'***************************************************************************************************************



Public Sub Add_Surface(ByVal addSurface As cls2D_ObjectSurface)
  Dim tempObject As cls2D_ObjectSurface
  
  Set tempObject = SurfaceList
  
  With addSurface
    Set .ParentObject = Me
    
    If tempObject Is Nothing Then
      Set SurfaceList = addSurface
      
      Set .PrevSurface = Nothing
      Set .NextSurface = Nothing
    Else
      Do While True
        If .ZPriority <= tempObject.ZPriority Then
          Set .PrevSurface = tempObject.PrevSurface
          Set tempObject.PrevSurface = addSurface
          
          If .PrevSurface Is Nothing Then
            Set SurfaceList = addSurface
          Else
            Set .PrevSurface.NextSurface = addSurface
          End If
          
          Set .NextSurface = tempObject
          
          Exit Do
        ElseIf tempObject.NextSurface Is Nothing Then
          Set .PrevSurface = tempObject
          Set tempObject.NextSurface = addSurface
          
          Set .NextSurface = Nothing
          
          Exit Do
        Else
          Set tempObject = tempObject.NextSurface
        End If
      Loop
    End If
  End With
End Sub

'Removes the animation layer from the layer list and also
'releases references that the object may hold so it can be released by the system if not needed
Public Sub Remove_Surface(ByVal removeSurface As cls2D_ObjectSurface)
  With removeSurface
    If .PrevSurface Is Nothing Then
      Set SurfaceList = .NextSurface
    Else
      Set .PrevSurface.NextSurface = .NextSurface
    End If
    
    If Not (.NextSurface Is Nothing) Then
      Set .NextSurface.PrevSurface = .PrevSurface
      Set .NextSurface = Nothing
    End If
    
    Set .PrevSurface = Nothing
    Set .ParentObject = Nothing
  End With
End Sub

'Removes all animation layers from the layer list and also
'releases references that the objects may hold so they can be released by the system if not needed
Public Sub RemoveAll_Surfaces()
  Dim AList As cls2D_ObjectSurface
  
  Set AList = SurfaceList
  Set SurfaceList = Nothing
  
  Do While Not (AList Is Nothing)
    With AList
      Set AList = .NextSurface
      
      Set .NextSurface = Nothing
      Set .PrevSurface = Nothing
      Set .ParentObject = Nothing
    End With
  Loop
End Sub

Public Function Get_Surface(ByVal SurfaceIndex As Long) As cls2D_ObjectSurface
  Dim ALayer As cls2D_ObjectSurface
  
  Set ALayer = SurfaceList
  
  Do While Not (ALayer Is Nothing)
    SurfaceIndex = SurfaceIndex - 1
    
    If SurfaceIndex = 0 Then
      Set Get_Surface = ALayer
      
      Exit Function
    End If
    
    Set ALayer = ALayer.NextSurface
  Loop
End Function

Public Function Get_SurfaceID(ByVal LayerTypeID As Long) As cls2D_ObjectSurface
  Dim ALayer As cls2D_ObjectSurface
  
  Set ALayer = SurfaceList
  
  Do While Not (ALayer Is Nothing)
    If LayerTypeID = ALayer.TypeID Then
      Set Get_SurfaceID = ALayer
      
      Exit Function
    End If
    
    Set ALayer = ALayer.NextSurface
  Loop
End Function

Public Function Get_SurfaceCount() As Long
  Dim sList As cls2D_ObjectSurface
  
  Get_SurfaceCount = 0
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    Get_SurfaceCount = Get_SurfaceCount + 1
    
    Set sList = sList.NextSurface
  Loop
End Function

Public Function Get_ChainObjectCount() As Long
  Dim AList As cls2D_Object
  
  Set AList = Me
  
  Do While Not (AList Is Nothing)
    Get_ChainObjectCount = Get_ChainObjectCount + 1
    
    Set AList = AList.chainNext
  Loop
End Function

Public Function Duplicate_Object(ByVal newObjectName As String, ByVal newObjectID As Long) As cls2D_Object
  Dim loop1 As Long, tCount As Long, oldSurface As cls2D_ObjectSurface
    
  Set Duplicate_Object = New cls2D_Object
  
  With Duplicate_Object
    .objectName = newObjectName
    .Visible = Visible
    
    .renderPriority = renderPriority
    
    .TypeID = newObjectID
    
    .displayPosX = displayPosX
    .displayPosY = displayPosY
    
    .PosX_1000ths = PosX_1000ths
    .PosY_1000ths = PosY_1000ths
    
    .Collision_Mask = Collision_Mask
    .Collision_WidthX_1000ths = Collision_WidthX_1000ths
    .Collision_WidthY_1000ths = Collision_WidthY_1000ths
    
    .Collision_XLower = Collision_XLower
    .Collision_XUpper = Collision_XUpper
    .Collision_YLower = Collision_YLower
    .Collision_YUpper = Collision_YUpper
    
    .CollisionBoxVisible = CollisionBoxVisible
    .CollisionBoxColour = CollisionBoxColour
    
    tCount = Get_SurfaceCount() ' add attached surfaces
    
    For loop1 = 1 To tCount
      Set oldSurface = Get_Surface(loop1)
      
      .Add_Surface oldSurface.Duplicate_Object(oldSurface.TypeID)
    Next loop1
    
    '**********************************************************************************************
    'use this section for custom member variables
    '**********************************************************************************************
    
    .Health = Health
    .Damage = Damage
    .Level = Level
    .Experience = Experience
    .LifeSpan = LifeSpan
    .Age = Age
    .value = value
    
    .MainAction = MainAction
    .TicksPerMainAction = TicksPerMainAction
    .MainActionTickCounter = MainActionTickCounter
    .SubAction = SubAction
    .TicksPerSubAction = TicksPerSubAction
    .SubActionTickCounter = SubActionTickCounter
    
    .TicksPerFrame = TicksPerFrame
    .FrameTickCounter = FrameTickCounter
    
    .direction = direction
    .Range_1000ths = Range_1000ths
    .speed_1000ths = speed_1000ths
    
    .VelocityX_1000ths = VelocityX_1000ths
    .VelocityY_1000ths = VelocityY_1000ths
    .AccelerationX_1000ths = AccelerationX_1000ths
    .AccelerationY_1000ths = AccelerationY_1000ths
    
    .TargetMode = TargetMode
    .TargetX_1000ths = TargetX_1000ths
    .TargetY_1000ths = TargetY_1000ths
    
    .User1 = User1
    .User2 = User2
    .User3 = User3
    
    '**********************************************************************************************
  End With
End Function

Public Sub Render_Object()
  Dim objectSurface As cls2D_ObjectSurface, ImageXOffset As Long, ImageYOffset As Long
  
  Set objectSurface = SurfaceList
  
  Do While Not (objectSurface Is Nothing)
    With objectSurface
      If .Visible Then
        If .ActionFrame = -1 Then
          DXDraw.BlitTransparent .SurfaceIndex, displayPosX + .DisplayOffsetX, displayPosY + .DisplayOffsetY, .SurfaceOffsetX, .SurfaceOffsetY, .ImageWidth, .ImageHeight
        Else
          ImageXOffset = .ActionSequence(.ActionFrame)
          ImageYOffset = (ImageXOffset \ .ImagesPerRow) * .ImageHeight + .SurfaceOffsetY
          ImageXOffset = (ImageXOffset Mod .ImagesPerRow) * .ImageWidth + .SurfaceOffsetX
          
          DXDraw.BlitTransparent .SurfaceIndex, displayPosX + .DisplayOffsetX, displayPosY + .DisplayOffsetY, ImageXOffset, ImageYOffset, .ImageWidth, .ImageHeight
        End If
      End If
      
      Set objectSurface = .NextSurface
    End With
  Loop
End Sub

Public Sub Draw_CollisionBox()
  DXDraw.Draw_Box displayPosX - Collision_WidthX_1000ths / 1000, displayPosY - Collision_WidthY_1000ths / 1000, Collision_WidthX_1000ths / 500, Collision_WidthY_1000ths / 500, CollisionBoxColour, 1
End Sub

'translates the world XY coordinates to the display XY coordinates
Public Sub Translate_Object(ByVal BasePosX_1000ths As Long, ByVal BasePosY_1000ths As Long)
  displayPosX = (PosX_1000ths - BasePosX_1000ths) \ 1000
  displayPosY = (PosY_1000ths - BasePosY_1000ths) \ 1000
End Sub

'translates the world XY coordinates to the display XY coordinates - wrapping X and Y
Public Sub Translate_Object_Wrap(ByVal BasePosX_1000ths As Long, ByVal BasePosY_1000ths As Long, ByVal wrapX_1000ths As Long, ByVal wrapY_1000ths As Long)
  Dim worldCentreX_1000ths As Long, worldCentreY_1000ths As Long
  
  worldCentreX_1000ths = BasePosX_1000ths + DXDraw.m_ClippingRectangleWidth * 500
  worldCentreY_1000ths = BasePosY_1000ths + DXDraw.m_ClippingRectangleHeight * 500
  
  If PosX_1000ths < worldCentreX_1000ths Then
    If worldCentreX_1000ths - PosX_1000ths <= wrapX_1000ths \ 2 Then
      displayPosX = (PosX_1000ths - BasePosX_1000ths) \ 1000
    Else
      displayPosX = (PosX_1000ths - BasePosX_1000ths + wrapX_1000ths) \ 1000
    End If
  Else
    If PosX_1000ths - worldCentreX_1000ths <= wrapX_1000ths \ 2 Then
      displayPosX = (PosX_1000ths - BasePosX_1000ths) \ 1000
    Else
      displayPosX = (PosX_1000ths - BasePosX_1000ths - wrapX_1000ths) \ 1000
    End If
  End If
  
  If PosY_1000ths < worldCentreY_1000ths Then
    If worldCentreY_1000ths - PosY_1000ths <= wrapY_1000ths \ 2 Then
      displayPosY = (PosY_1000ths - BasePosY_1000ths) \ 1000
    Else
      displayPosY = (PosY_1000ths - BasePosY_1000ths + wrapY_1000ths) \ 1000
    End If
  Else
    If PosY_1000ths - worldCentreY_1000ths <= wrapY_1000ths \ 2 Then
      displayPosY = (PosY_1000ths - BasePosY_1000ths) \ 1000
    Else
      displayPosY = (PosY_1000ths - BasePosY_1000ths - wrapY_1000ths) \ 1000
    End If
  End If
End Sub

Public Sub Calculate_CollsionBox()
  Collision_XLower = PosX_1000ths - Collision_WidthX_1000ths
  Collision_XUpper = PosX_1000ths + Collision_WidthX_1000ths
  Collision_YLower = PosY_1000ths - Collision_WidthY_1000ths
  Collision_YUpper = PosY_1000ths + Collision_WidthY_1000ths
End Sub

Public Function Test_Collision(ByVal targetObject As cls2D_Object, Optional ByVal targetMask As Long = &HFFFFFFFF) As Boolean
  With targetObject
    If targetMask And .Collision_Mask Then
      If Collision_XLower <= .Collision_XUpper Then
        If Collision_XUpper >= .Collision_XLower Then
          If Collision_YLower <= .Collision_YUpper Then
            If Collision_YUpper >= .Collision_YLower Then Test_Collision = True
          End If
        End If
      End If
    End If
  End With
End Function

Public Function Test_Collision_Wrap(ByVal targetObject As cls2D_Object, _
        wrapX_1000ths As Long, wrapY_1000ths As Long, Optional ByVal targetMask As Long = &HFFFFFFFF) As Boolean
        
  With targetObject
    If targetMask And .Collision_Mask Then
      If Collision_XLower <= .Collision_XUpper Then
        If Collision_XUpper >= .Collision_XLower Then
          If Collision_YLower <= .Collision_YUpper Then
            If Collision_YUpper >= .Collision_YLower Then
              Test_Collision_Wrap = True
            ElseIf Collision_YLower + wrapY_1000ths <= .Collision_YUpper Then
              If Collision_YUpper + wrapY_1000ths >= .Collision_YLower Then Test_Collision_Wrap = True
            End If
          ElseIf Collision_YLower - wrapY_1000ths <= .Collision_YUpper Then
            If Collision_YUpper - wrapY_1000ths >= .Collision_YLower Then Test_Collision_Wrap = True
          End If
        ElseIf Collision_XLower + wrapX_1000ths <= .Collision_XUpper Then
          If Collision_XUpper + wrapX_1000ths >= .Collision_XLower Then
            If Collision_YLower <= .Collision_YUpper Then
              If Collision_YUpper >= .Collision_YLower Then
                Test_Collision_Wrap = True
              ElseIf Collision_YLower + wrapY_1000ths <= .Collision_YUpper Then
                If Collision_YUpper + wrapY_1000ths >= .Collision_YLower Then Test_Collision_Wrap = True
              End If
            ElseIf Collision_YLower - wrapY_1000ths <= .Collision_YUpper Then
              If Collision_YUpper - wrapY_1000ths >= .Collision_YLower Then Test_Collision_Wrap = True
            End If
          End If
        End If
      ElseIf Collision_XLower - wrapX_1000ths <= .Collision_XUpper Then
        If Collision_XUpper - wrapX_1000ths >= .Collision_XLower Then
          If Collision_YLower <= .Collision_YUpper Then
            If Collision_YUpper >= .Collision_YLower Then
              Test_Collision_Wrap = True
            ElseIf Collision_YLower + wrapY_1000ths <= .Collision_YUpper Then
              If Collision_YUpper + wrapY_1000ths >= .Collision_YLower Then Test_Collision_Wrap = True
            End If
          ElseIf Collision_YLower - wrapY_1000ths <= .Collision_YUpper Then
            If Collision_YUpper - wrapY_1000ths >= .Collision_YLower Then Test_Collision_Wrap = True
          End If
        End If
      End If
    End If
  End With
End Function

'get the distance between this object an a point in space
Public Function Get_Distance(ByVal targetXPos_1000ths As Long, ByVal targetYPos_1000ths As Long) As Long
  Get_Distance = Sqr((targetXPos_1000ths - PosX_1000ths) ^ 2 + (targetYPos_1000ths - PosY_1000ths) ^ 2)
End Function

'get the distance between this object an a point in space (world wraps)
Public Function Get_DistanceWrap(ByVal targetXPos_1000ths As Long, ByVal targetYPos_1000ths As Long, ByVal worldSizeX_1000ths As Long, ByVal worldSizeY_1000ths As Long) As Long
  targetXPos_1000ths = targetXPos_1000ths - PosX_1000ths
  targetYPos_1000ths = targetYPos_1000ths - PosY_1000ths
  
  If targetXPos_1000ths > worldSizeX_1000ths \ 2 Then
    targetXPos_1000ths = targetXPos_1000ths - worldSizeX_1000ths
  ElseIf targetXPos_1000ths < worldSizeX_1000ths \ -2 Then
    targetXPos_1000ths = targetXPos_1000ths + worldSizeX_1000ths
  End If
  
  If targetYPos_1000ths > worldSizeY_1000ths \ 2 Then
    targetYPos_1000ths = targetYPos_1000ths - worldSizeY_1000ths
  ElseIf targetYPos_1000ths < worldSizeY_1000ths \ -2 Then
    targetYPos_1000ths = targetYPos_1000ths + worldSizeY_1000ths
  End If
  
  Get_DistanceWrap = Sqr(targetXPos_1000ths ^ 2 + targetYPos_1000ths ^ 2)
End Function

'get the direction between this object an a point in space CCW from +Y
Public Function Get_Direction(ByVal targetXPos_1000ths As Long, ByVal targetYPos_1000ths As Long) As Single
  targetXPos_1000ths = targetXPos_1000ths - PosX_1000ths
  targetYPos_1000ths = targetYPos_1000ths - PosY_1000ths
  
  If targetYPos_1000ths < 0 Then
    Get_Direction = Pi + Atn(targetXPos_1000ths / targetYPos_1000ths)
  ElseIf targetYPos_1000ths > 0 Then
    Get_Direction = Atn(targetXPos_1000ths / targetYPos_1000ths)
    
    If Get_Direction < 0 Then Get_Direction = Get_Direction + TwoPi
  ElseIf targetXPos_1000ths <= 0 Then
    Get_Direction = ThreePiByTwo
  Else
    Get_Direction = PiByTwo
  End If
End Function

'get the direction between this object an a point in space CCW from +Y (world wraps)
Public Function Get_DirectionWrap(ByVal targetXPos_1000ths As Long, ByVal targetYPos_1000ths As Long, ByVal worldSizeX_1000ths As Long, ByVal worldSizeY_1000ths As Long) As Single
  targetXPos_1000ths = targetXPos_1000ths - PosX_1000ths
  targetYPos_1000ths = targetYPos_1000ths - PosY_1000ths
  
  If targetXPos_1000ths > worldSizeX_1000ths \ 2 Then
    targetXPos_1000ths = targetXPos_1000ths - worldSizeX_1000ths
  ElseIf targetXPos_1000ths < worldSizeX_1000ths \ -2 Then
    targetXPos_1000ths = targetXPos_1000ths + worldSizeX_1000ths
  End If
  
  If targetYPos_1000ths > worldSizeY_1000ths \ 2 Then
    targetYPos_1000ths = targetYPos_1000ths - worldSizeY_1000ths
  ElseIf targetYPos_1000ths < worldSizeY_1000ths \ -2 Then
    targetYPos_1000ths = targetYPos_1000ths + worldSizeY_1000ths
  End If
  
  If targetYPos_1000ths < 0 Then
    Get_DirectionWrap = Pi + Atn(targetXPos_1000ths / targetYPos_1000ths)
  ElseIf targetYPos_1000ths > 0 Then
    Get_DirectionWrap = Atn(targetXPos_1000ths / targetYPos_1000ths)
    
    If Get_DirectionWrap < 0 Then Get_DirectionWrap = Get_DirectionWrap + TwoPi
  ElseIf targetXPos_1000ths <= 0 Then
    Get_DirectionWrap = ThreePiByTwo
  Else
    Get_DirectionWrap = PiByTwo
  End If
End Function

'get speed and direction (in radians CCW from +Y axis) based on component x and y velocities
Public Sub Get_Speed(ByVal xVel_1000ths As Long, ByVal yVel_1000ths As Long, ByRef speed_1000ths As Long, ByRef direction As Single)
  speed_1000ths = Sqr(xVel_1000ths ^ 2 + yVel_1000ths ^ 2)
  
  If yVel_1000ths < 0 Then
    direction = Pi + Atn(xVel_1000ths / yVel_1000ths)
  ElseIf yVel_1000ths > 0 Then
    direction = Atn(xVel_1000ths / yVel_1000ths)
    
    If direction < 0 Then direction = direction + TwoPi
  ElseIf xVel_1000ths <= 0 Then
    direction = ThreePiByTwo
  Else
    direction = PiByTwo
  End If
End Sub

'get component x and y velocities based on speed_1000ths and direction
Public Sub Get_Velocity(ByVal speed_1000ths As Long, ByVal direction As Single, ByRef xVel_1000ths As Long, ByRef yVel_1000ths As Long)
  xVel_1000ths = speed_1000ths * Sin(direction)
  yVel_1000ths = speed_1000ths * Cos(direction)
End Sub

'get the difference between the source angle and the target angle (in radians)
Public Function Get_DifferenceAngle(ByVal sourceAngle As Single, ByVal targetAngle As Single) As Single
  Get_DifferenceAngle = targetAngle - sourceAngle
  
  Do While Get_DifferenceAngle > Pi
    Get_DifferenceAngle = Get_DifferenceAngle - TwoPi
  Loop
  
  Do While Get_DifferenceAngle < -Pi
    Get_DifferenceAngle = Get_DifferenceAngle + TwoPi
  Loop
End Function

Private Sub Class_Initialize()
  Visible = True
End Sub

'Releases references that the object may hold so it can be released by the system
'The object should not exist in any chains or that chain will become corrupt - first remove from chain
Public Sub Destroy_Object()
  Dim AList As cls2D_ObjectSurface
  
  On Error Resume Next
  
  Set chainNext = Nothing
  Set chainPrevious = Nothing
  
  Set AList = SurfaceList
  Set SurfaceList = Nothing
    
  Do While Not (AList Is Nothing)
    With AList
      Set AList = .NextSurface
      
      Set .ParentObject = Nothing
      Set .PrevSurface = Nothing
      Set .NextSurface = Nothing
    End With
  Loop
End Sub

' Saves 2D object definition to file
Public Sub Save_ToFile(filePathName As String)
  Dim fileHandle As Long
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Output As #fileHandle
    Write #fileHandle, "Logicon Enterprises cls2D_Object Description Ver2.6"
    
    Save_ToOpenFile fileHandle
  Close fileHandle
  
  Exit Sub
  
badFile:
  MsgBox "Error writing 2D object to -> " & filePathName, vbCritical, "ERROR"
End Sub

' Saves 2D chain definition to file
Public Sub Save_ChainToFile(filePathName As String)
  Dim fileHandle As Long, objectList As cls2D_Object, tCount As Long
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Output As #fileHandle
    Write #fileHandle, "Logicon Enterprises cls2D_Chain Description Ver2.6"
    
    tCount = Get_ChainObjectCount()
    Write #fileHandle, tCount
    
    Set objectList = Me
    
    Do While Not (objectList Is Nothing)
      With objectList
        .Save_ToOpenFile fileHandle
        
        Set objectList = .chainNext
      End With
    Loop
  Close fileHandle
  
  Exit Sub
  
badFile:
  MsgBox "Error writing 2D chain to -> " & filePathName, vbCritical, "ERROR"
End Sub

' This should not be called directly... instead call Save_ToFile()
Public Sub Save_ToOpenFile(ByVal fileHandle)
  Dim sList As cls2D_ObjectSurface, tCount As Long
  Dim bitmapSurfaceName As String, bitmapSurfaceWidth As Long, bitmapIsTexture As Boolean
  Dim bitmapSurfaceHeight As Long, bitmapSurfaceTransparency As Long, numberOfBitmapSurfaces As Long
  Dim loop1 As Long, loop2 As Long, staticSurfaceInUse() As Boolean
  
  Write #fileHandle, objectName
  Write #fileHandle, Visible

  Write #fileHandle, renderPriority
  
  Write #fileHandle, TypeID
  
  Write #fileHandle, displayPosX
  Write #fileHandle, displayPosY
  
  Write #fileHandle, PosX_1000ths
  Write #fileHandle, PosY_1000ths
  
  Write #fileHandle, Collision_Mask
  Write #fileHandle, Collision_WidthX_1000ths
  Write #fileHandle, Collision_WidthY_1000ths

  Write #fileHandle, Collision_XLower, Collision_XUpper
  Write #fileHandle, Collision_YLower, Collision_YUpper
  
  Write #fileHandle, CollisionBoxVisible
  Write #fileHandle, CollisionBoxColour
  
  numberOfBitmapSurfaces = DXDraw.Get_TotalStaticSurfaces()
  
  'I am using 0 even though it shouldn't occur just to prevent possible error
  ReDim staticSurfaceInUse(0 To numberOfBitmapSurfaces)
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      staticSurfaceInUse(.SurfaceIndex) = True
      
      Set sList = .NextSurface
    End With
  Loop
  
  loop2 = 0
  
  For loop1 = 1 To numberOfBitmapSurfaces
    If staticSurfaceInUse(loop1) = True Then loop2 = loop2 + 1
  Next loop1
  
  Write #fileHandle, numberOfBitmapSurfaces
  Write #fileHandle, loop2
  
  For loop1 = 1 To numberOfBitmapSurfaces 'save current static surface info for reload
    If staticSurfaceInUse(loop1) = True Then
      Write #fileHandle, loop1
      Write #fileHandle, DXDraw.Get_SSurfaceFileName(loop1)
      
      DXDraw.Get_SSurfaceSettings loop1, bitmapSurfaceWidth, bitmapSurfaceHeight, bitmapSurfaceTransparency, bitmapIsTexture
      
      Write #fileHandle, bitmapSurfaceWidth
      Write #fileHandle, bitmapSurfaceHeight
      Write #fileHandle, bitmapSurfaceTransparency
    End If
  Next loop1
  
  tCount = Get_SurfaceCount() ' save attached surfaces
  Write #fileHandle, tCount
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Save_ToOpenFile fileHandle
      
      Set sList = .NextSurface
    End With
  Loop
  
  '**********************************************************************************************
  'use this section for custom member variables
  '**********************************************************************************************
  
  Write #fileHandle, Health
  Write #fileHandle, Damage
  Write #fileHandle, Level
  Write #fileHandle, Experience
  Write #fileHandle, LifeSpan
  Write #fileHandle, Age
  Write #fileHandle, value

  Write #fileHandle, MainAction
  Write #fileHandle, TicksPerMainAction
  Write #fileHandle, MainActionTickCounter
  Write #fileHandle, SubAction
  Write #fileHandle, TicksPerSubAction
  Write #fileHandle, SubActionTickCounter
  
  Write #fileHandle, TicksPerFrame
  Write #fileHandle, FrameTickCounter
  
  Write #fileHandle, direction
  Write #fileHandle, Range_1000ths
  Write #fileHandle, speed_1000ths
  
  Write #fileHandle, VelocityX_1000ths
  Write #fileHandle, VelocityY_1000ths
  Write #fileHandle, AccelerationX_1000ths
  Write #fileHandle, AccelerationY_1000ths
  
  Write #fileHandle, TargetMode
  Write #fileHandle, TargetX_1000ths
  Write #fileHandle, TargetY_1000ths
  
  Write #fileHandle, User1
  Write #fileHandle, User2
  Write #fileHandle, User3
  
  '**********************************************************************************************
End Sub

' Loads 2D object definition from file
Public Sub Load_FromFile(filePathName As String)
  Dim fileHandle As Long, versionString As String
  
  'On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Input As #fileHandle
    Input #fileHandle, versionString
    
    Load_FromOpenFile fileHandle
  Close fileHandle
  
  Exit Sub
  
badFile:
  MsgBox "Error loading 2D object -> " & filePathName, vbCritical, "ERROR"
End Sub

' Loads 2D chain definition to file
Public Sub Load_ChainFromFile(filePathName As String)
  Dim fileHandle As Long, newObject As cls2D_Object, lastObject As cls2D_Object
  Dim tCount As Long, loop1 As Long, versionString As String
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Input As #fileHandle
    Input #fileHandle, versionString
    Input #fileHandle, tCount
    
    Destroy_Chain
    Load_FromOpenFile fileHandle
    
    Set lastObject = Me
    
    For loop1 = 2 To tCount
      Set newObject = New cls2D_Object
      
      With newObject
        .Load_FromOpenFile fileHandle
        
        Set .chainPrevious = lastObject 'don't use AddTo_Chain as we wish to preserve the order
        Set lastObject.chainNext = newObject
        
        Set lastObject = newObject
      End With
    Next loop1
  Close fileHandle
  
  Exit Sub
  
badFile:
  MsgBox "Error loading 2D chain from -> " & filePathName, vbCritical, "ERROR"
End Sub

' This should not be called directly... instead call Load_FromFile()
Public Sub Load_FromOpenFile(ByVal fileHandle As Long)
  Dim sList As cls2D_ObjectSurface, tCount As Long
  Dim bitmapSurfaceIndex() As Long, bitmapSurfaceName As String, bitmapSurfaceWidth As Long
  Dim bitmapSurfaceHeight As Long, bitmapSurfaceTransparency As Long, numberOfBitmapSurfaces As Long
  Dim loop1 As Long, numberOfSurfacesInUse As Long, surfaceNumber As Long
  
  Input #fileHandle, objectName
  Input #fileHandle, Visible

  Input #fileHandle, renderPriority

  Input #fileHandle, TypeID
    
  Input #fileHandle, displayPosX
  Input #fileHandle, displayPosY
  
  Input #fileHandle, PosX_1000ths
  Input #fileHandle, PosY_1000ths
  
  Input #fileHandle, Collision_Mask
  Input #fileHandle, Collision_WidthX_1000ths
  Input #fileHandle, Collision_WidthY_1000ths

  Input #fileHandle, Collision_XLower, Collision_XUpper
  Input #fileHandle, Collision_YLower, Collision_YUpper
  
  Input #fileHandle, CollisionBoxVisible
  Input #fileHandle, CollisionBoxColour
  
  Input #fileHandle, numberOfBitmapSurfaces
  Input #fileHandle, numberOfSurfacesInUse
  
  ReDim bitmapSurfaceIndex(0 To numberOfBitmapSurfaces)
  
  For loop1 = 1 To numberOfSurfacesInUse 'create a static surface remapper
    Input #fileHandle, surfaceNumber
    Input #fileHandle, bitmapSurfaceName
    Input #fileHandle, bitmapSurfaceWidth
    Input #fileHandle, bitmapSurfaceHeight
    Input #fileHandle, bitmapSurfaceTransparency
    
    bitmapSurfaceIndex(surfaceNumber) = DXDraw.Get_SSurfaceIndex(bitmapSurfaceName)
      
    If bitmapSurfaceIndex(surfaceNumber) = 0 Then
      bitmapSurfaceIndex(surfaceNumber) = DXDraw.Get_SSurfaceFreeIndex()
        
      If bitmapSurfaceIndex(surfaceNumber) <> 0 Then
        DXDraw.Init_StaticSurface bitmapSurfaceIndex(surfaceNumber), bitmapSurfaceName, bitmapSurfaceWidth, bitmapSurfaceHeight, bitmapSurfaceTransparency
      End If
    End If
  Next loop1
  
  RemoveAll_Surfaces 'load attached surfaces
  Input #fileHandle, tCount
  
  For loop1 = 1 To tCount
    Set sList = New cls2D_ObjectSurface
    
    With sList
      .Load_FromOpenFile fileHandle
      
      .SurfaceIndex = bitmapSurfaceIndex(.SurfaceIndex)
    End With
    
    Add_Surface sList
  Next loop1
  
  '**********************************************************************************************
  'use this section for custom member variables
  '**********************************************************************************************
  
  Input #fileHandle, Health
  Input #fileHandle, Damage
  Input #fileHandle, Level
  Input #fileHandle, Experience
  Input #fileHandle, LifeSpan
  Input #fileHandle, Age
  Input #fileHandle, value

  Input #fileHandle, MainAction
  Input #fileHandle, TicksPerMainAction
  Input #fileHandle, MainActionTickCounter
  Input #fileHandle, SubAction
  Input #fileHandle, TicksPerSubAction
  Input #fileHandle, SubActionTickCounter
  
  Input #fileHandle, TicksPerFrame
  Input #fileHandle, FrameTickCounter
  
  Input #fileHandle, direction
  Input #fileHandle, Range_1000ths
  Input #fileHandle, speed_1000ths
  
  Input #fileHandle, VelocityX_1000ths
  Input #fileHandle, VelocityY_1000ths
  Input #fileHandle, AccelerationX_1000ths
  Input #fileHandle, AccelerationY_1000ths
  
  Input #fileHandle, TargetMode
  Input #fileHandle, TargetX_1000ths
  Input #fileHandle, TargetY_1000ths
  
  Input #fileHandle, User1
  Input #fileHandle, User2
  Input #fileHandle, User3
  
  '**********************************************************************************************
End Sub

'**********************************************************************************************************

'******************************************** chains *********************************************

'**********************************************************************************************************

'Redraws the 2D object chain
Public Sub ReDraw_Chain(ByVal BasePosX_1000ths As Long, ByVal BasePosY_1000ths As Long, Optional ByVal inclusionZone As Long = 100000)
  Dim chainList As cls2D_Object
  Dim clipLowX As Long, clipHighX As Long, clipLowY As Long, clipHighY As Long
  
  clipLowX = -inclusionZone
  clipLowY = -inclusionZone
  clipHighX = DXDraw.m_ClippingRectangleWidth + inclusionZone
  clipHighY = DXDraw.m_ClippingRectangleHeight + inclusionZone
  
  Set chainList = Me
  
  Do While Not (chainList Is Nothing)
    With chainList
      If .Visible Then
        .displayPosX = (.PosX_1000ths - BasePosX_1000ths) \ 1000
        .displayPosY = (.PosY_1000ths - BasePosY_1000ths) \ 1000
        
        If .displayPosX >= clipLowX And .displayPosX <= clipHighX Then
          If .displayPosY >= clipLowY And .displayPosY <= clipHighY Then
            .Render_Object
            
            If .CollisionBoxVisible Then DXDraw.Draw_Box .displayPosX - .Collision_WidthX_1000ths / 1000, .displayPosY - .Collision_WidthY_1000ths / 1000, .Collision_WidthX_1000ths / 500, .Collision_WidthY_1000ths / 500, .CollisionBoxColour, 1
          End If
        End If
      End If
      
      Set chainList = .chainNext
    End With
  Loop
End Sub

'Redraws the 2D object chain - wrapping the X & Y
Public Sub ReDraw_Chain_Wrap(ByVal BasePosX_1000ths As Long, ByVal BasePosY_1000ths As Long, _
          ByVal wrapX_1000ths As Long, ByVal wrapY_1000ths As Long, Optional ByVal inclusionZone As Long = 100000)
          
  Dim chainList As cls2D_Object
  Dim clipLowX As Long, clipHighX As Long, clipLowY As Long, clipHighY As Long
  Dim worldCentreX_1000ths As Long, worldCentreY_1000ths As Long
  Dim wrapXHalf_1000ths As Long, wrapYHalf_1000ths As Long
  
  wrapXHalf_1000ths = wrapX_1000ths \ 2
  wrapYHalf_1000ths = wrapY_1000ths \ 2
  worldCentreX_1000ths = BasePosX_1000ths + DXDraw.m_ClippingRectangleWidth * 500
  worldCentreY_1000ths = BasePosY_1000ths + DXDraw.m_ClippingRectangleHeight * 500
  
  clipLowX = -inclusionZone
  clipLowY = -inclusionZone
  clipHighX = DXDraw.m_ClippingRectangleWidth + inclusionZone
  clipHighY = DXDraw.m_ClippingRectangleHeight + inclusionZone
  
  Set chainList = Me
  
  Do While Not (chainList Is Nothing)
    With chainList
      If .Visible Then
        If .PosX_1000ths < worldCentreX_1000ths Then
          If worldCentreX_1000ths - .PosX_1000ths <= wrapXHalf_1000ths Then
            .displayPosX = (.PosX_1000ths - BasePosX_1000ths) \ 1000
          Else
            .displayPosX = (.PosX_1000ths - BasePosX_1000ths + wrapX_1000ths) \ 1000
          End If
        Else
          If .PosX_1000ths - worldCentreX_1000ths <= wrapXHalf_1000ths Then
            .displayPosX = (.PosX_1000ths - BasePosX_1000ths) \ 1000
          Else
            .displayPosX = (.PosX_1000ths - BasePosX_1000ths - wrapX_1000ths) \ 1000
          End If
        End If
        
        If .PosY_1000ths < worldCentreY_1000ths Then
          If worldCentreY_1000ths - .PosY_1000ths <= wrapYHalf_1000ths Then
            .displayPosY = (.PosY_1000ths - BasePosY_1000ths) \ 1000
          Else
            .displayPosY = (.PosY_1000ths - BasePosY_1000ths + wrapY_1000ths) \ 1000
          End If
        Else
          If .PosY_1000ths - worldCentreY_1000ths <= wrapYHalf_1000ths Then
            .displayPosY = (.PosY_1000ths - BasePosY_1000ths) \ 1000
          Else
            .displayPosY = (.PosY_1000ths - BasePosY_1000ths - wrapY_1000ths) \ 1000
          End If
        End If
        
        If .displayPosX >= clipLowX And .displayPosX <= clipHighX Then
          If .displayPosY >= clipLowY And .displayPosY <= clipHighY Then
            .Render_Object
            
            If .CollisionBoxVisible Then DXDraw.Draw_Box .displayPosX - .Collision_WidthX_1000ths / 1000, .displayPosY - .Collision_WidthY_1000ths / 1000, .Collision_WidthX_1000ths / 500, .Collision_WidthY_1000ths / 500, .CollisionBoxColour, 1
          End If
        End If
      End If
      
      Set chainList = .chainNext
    End With
  Loop
End Sub

Public Sub Calculate_Chain_CollisionBox()
  Dim chainList As cls2D_Object
  
  Set chainList = Me
   
  Do While Not (chainList Is Nothing)
    With chainList
      .Collision_XLower = .PosX_1000ths - .Collision_WidthX_1000ths
      .Collision_XUpper = .PosX_1000ths + .Collision_WidthX_1000ths
      .Collision_YLower = .PosY_1000ths - .Collision_WidthY_1000ths
      .Collision_YUpper = .PosY_1000ths + .Collision_WidthY_1000ths
      
      Set chainList = .chainNext
    End With
  Loop
End Sub

'removes the 2D object from the chain and returns the first object in the chain
'removing an object from the chain that was not in the chain will cause corruption in
'the chain - object may then be destroyed
Public Sub RemoveFrom_Chain(ByRef chainStart As cls2D_Object)
  If Not (chainNext Is Nothing) Then Set chainNext.chainPrevious = chainPrevious
    
  If chainPrevious Is Nothing Then
    Set chainStart = chainNext
  Else
    Set chainPrevious.chainNext = chainNext
    Set chainPrevious = Nothing
  End If
  
  Set chainNext = Nothing
End Sub

'adds the 2D object to the chain at its appropriate location for display priority and
'adding an object already in the chain will cause corruption in the chain
Public Sub AddTo_Chain(ByRef chainStart As cls2D_Object)
  Dim chainList As cls2D_Object
   
  If chainStart Is Nothing Then
    Set chainNext = Nothing
    Set chainPrevious = Nothing
    
    Set chainStart = Me
    
    Exit Sub
  End If
  
  Set chainList = chainStart
  
  Do While Not (chainList Is Nothing) 'do forward scan
    With chainList
      If renderPriority <= .renderPriority Then
        Set chainNext = chainList
        Set chainPrevious = .chainPrevious
        
        If chainPrevious Is Nothing Then
          Set chainStart = Me
        Else
          Set .chainPrevious.chainNext = Me
        End If
        
        Set .chainPrevious = Me
        
        Set chainList = Nothing
      ElseIf .chainNext Is Nothing Then
        Set .chainNext = Me
        
        Set chainPrevious = chainList
        Set chainNext = Nothing
        
        Set chainList = Nothing
      Else
        Set chainList = .chainNext
      End If
    End With
  Loop
End Sub

'recalculates the 2D object's position in the chain and returns the first object in the chain
'the selected object must be a valid member of the a chain or corruption may occur
Public Sub MoveIn_Chain(ByRef chainStart As cls2D_Object)
  RemoveFrom_Chain chainStart
  AddTo_Chain chainStart
End Sub

'Releases references that the 2D objects in the chain may hold so they can be released by the system
Public Sub Destroy_Chain()
  Dim chainList As cls2D_Object
  
  Set chainList = Me
  
  Do While Not (chainList Is Nothing)
    With chainList
      Set chainList = .chainNext
      
      .Destroy_Object
    End With
  Loop
End Sub

Public Function Get_ChainObjectID(ByVal objectID As Long) As cls2D_Object
  Dim chainList As cls2D_Object
  
  Set chainList = Me
  
  Do While Not (chainList Is Nothing)
    With chainList
      If .TypeID = objectID Then
        Set Get_ChainObjectID = chainList
        
        Exit Function
      End If
      
      Set chainList = .chainNext
    End With
  Loop
End Function

Public Function Get_ChainObjectName(ByVal objectName As String, Optional ByVal objectID As Long = 0) As cls2D_Object
  Dim chainList As cls2D_Object
  
  Set chainList = Me
  
  If objectID = 0 Then
    Do While Not (chainList Is Nothing)
      With chainList
        If .objectName = objectName Then
          Set Get_ChainObjectName = chainList
          
          Exit Function
        End If
        
        Set chainList = .chainNext
      End With
    Loop
  Else
    Do While Not (chainList Is Nothing)
      With chainList
        If .objectName = objectName Then
          If .TypeID = objectID Then
            Set Get_ChainObjectName = chainList
            
            Exit Function
          End If
        End If
        
        Set chainList = .chainNext
      End With
    Loop
  End If
End Function

