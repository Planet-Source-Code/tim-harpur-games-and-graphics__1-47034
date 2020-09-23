VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsISO_Object"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************************
'
' clsISO_Object class definition
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

Private dx_DirectX As DirectX7
Private dx_Direct3D As Direct3D7
Private dx_Direct3DDevice As Direct3DDevice7

Public objectName As String                                 ' object's name

Public Visible As Boolean                                         ' object's master visibility flag (default TRUE)
Public visibleInFOW As Boolean           ' object's visibility override flag when in FOW (default TRUE)
Public wasDisplayed As Long               ' if hidden by FOW this will be 0 (system use only)

Public chainPrevious As clsISO_Object       ' previous object in chain
Public chainNext As clsISO_Object              ' next object in chain

Public mapChainPrevious As clsISO_Object       ' previous object in map cell chain
Public mapChainNext As clsISO_Object              ' next object in map cell chain
Public mapChainLevel As Long
Public mapChainRow As Long
Public mapChainColumn As Long

Public renderPriority As Long                                ' object's render priority - system provided

Public SurfaceList As clsISO_ObjectSurface                    ' list of object's image surfaces

Public TypeID As Long                                             ' object ID value
  
Public displayPosX As Long           ' object's X & Y display position - these values are automatically
Public displayPosY As Long           ' loaded by a call to Translate_Chain_to_???View()
  
Public PosX_1000ths As Long                                       ' object's X world position
Public PosY_1000ths As Long                                       ' object's Y world position
Public PosZ_1000ths As Long                                       ' object's Z world position

Public Collision_Mask As Long                                     ' object's collision bit mask
Public Collision_WidthX_1000ths As Long                   ' X width of collision box
Public Collision_WidthY_1000ths As Long                   ' Y width of collision box
Public Collision_WidthZ_1000ths As Long                    ' Z width (height) of collision box (IOSmetric only)

Public Collision_XLower As Long, Collision_XUpper As Long         'calculated collision bounds
Public Collision_YLower As Long, Collision_YUpper As Long         'calculated collision bounds
Public Collision_ZLower As Long, Collision_ZUpper As Long         'calculated collision bounds

Public CollisionBoxColour As Long                  ' ARGB line colour for drawing collision box
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
Public VelocityZ_1000ths As Long                                  ' user defined value
Public AccelerationX_1000ths As Long                          ' user defined value
Public AccelerationY_1000ths As Long                          ' user defined value
Public AccelerationZ_1000ths As Long                          ' user defined value
  
Public TargetMode As Long                                           ' user defined value
Public TargetX_1000ths As Long                                  ' user defined value
Public TargetY_1000ths As Long                                  ' user defined value
Public TargetZ_1000ths As Long                                  ' user defined value

Public User1 As Long                                                  ' user defined value
Public User2 As Long                                                  ' user defined value
Public User3 As Long                                                  ' user defined value

'***************************************************************************************************************



Public Sub Add_Surface(ByVal addSurface As clsISO_ObjectSurface)
  Dim tempObject As clsISO_ObjectSurface
  
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
Public Sub Remove_Surface(ByVal removeSurface As clsISO_ObjectSurface)
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
  Dim AList As clsISO_ObjectSurface
  
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

Public Function Get_Surface(ByVal SurfaceIndex As Long) As clsISO_ObjectSurface
  Dim ALayer As clsISO_ObjectSurface
  
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

Public Function Get_SurfaceID(ByVal LayerTypeID As Long) As clsISO_ObjectSurface
  Dim ALayer As clsISO_ObjectSurface
  
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
  Dim sList As clsISO_ObjectSurface
  
  Get_SurfaceCount = 0
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    Get_SurfaceCount = Get_SurfaceCount + 1
    
    Set sList = sList.NextSurface
  Loop
End Function

Public Function Get_ChainObjectCount() As Long
  Dim AList As clsISO_Object
  
  Set AList = Me
  
  Do While Not (AList Is Nothing)
    Get_ChainObjectCount = Get_ChainObjectCount + 1
    
    Set AList = AList.chainNext
  Loop
End Function

Public Sub Set_ObjectColour(ByVal objectColour As Long)
  Dim ALayer As clsISO_ObjectSurface
  
  Set ALayer = SurfaceList
  
  Do While Not (ALayer Is Nothing)
    With ALayer
      .SurfaceColour = objectColour
      
      Set ALayer = .NextSurface
    End With
  Loop
End Sub

Public Function Duplicate_Object(ByVal newObjectName As String, ByVal newObjectID As Long) As clsISO_Object
  Dim loop1 As Long, oldSurface As clsISO_ObjectSurface
    
  Set Duplicate_Object = New clsISO_Object
  
  With Duplicate_Object
    .objectName = newObjectName
    .Visible = Visible
    
    .renderPriority = renderPriority
    
    .TypeID = newObjectID
    
    .displayPosX = displayPosX
    .displayPosY = displayPosY
    
    .PosX_1000ths = PosX_1000ths
    .PosY_1000ths = PosY_1000ths
    .PosZ_1000ths = PosZ_1000ths
    
    .Collision_Mask = Collision_Mask
    .Collision_WidthX_1000ths = Collision_WidthX_1000ths
    .Collision_WidthY_1000ths = Collision_WidthY_1000ths
    .Collision_WidthZ_1000ths = Collision_WidthZ_1000ths
    
    .Collision_XLower = Collision_XLower
    .Collision_XUpper = Collision_XUpper
    .Collision_YLower = Collision_YLower
    .Collision_YUpper = Collision_YUpper
    .Collision_ZLower = Collision_ZLower
    .Collision_ZUpper = Collision_ZUpper
    
    .CollisionBoxVisible = CollisionBoxVisible
    .CollisionBoxColour = CollisionBoxColour
    
    Set oldSurface = SurfaceList
    
    Do While Not (oldSurface Is Nothing)
      .Add_Surface oldSurface.Duplicate_Object(oldSurface.TypeID)
      
      Set oldSurface = oldSurface.NextSurface
    Loop
    
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
    .VelocityZ_1000ths = VelocityZ_1000ths
    .AccelerationX_1000ths = AccelerationX_1000ths
    .AccelerationY_1000ths = AccelerationY_1000ths
    .AccelerationZ_1000ths = AccelerationZ_1000ths
    
    .TargetMode = TargetMode
    .TargetX_1000ths = TargetX_1000ths
    .TargetY_1000ths = TargetY_1000ths
    .TargetZ_1000ths = TargetZ_1000ths
    
    .User1 = User1
    .User2 = User2
    .User3 = User3
    
    '**********************************************************************************************
  End With
End Function

Public Sub Render_Object()
  Dim objectSurface As clsISO_ObjectSurface, imageBase As Long, ImageUOffset As Single, ImageVOffset As Single
  
  Set objectSurface = SurfaceList
  
  Do While Not (objectSurface Is Nothing)
    With objectSurface
      If .Visible Then
        If .ActionFrame = -1 Then
          DXDraw.BlitTexture .SurfaceIndex, displayPosX + .DisplayOffsetX * DXDraw.ISOmetricViewScale, displayPosY + .DisplayOffsetY * DXDraw.ISOmetricViewScale, .DisplayWidth * DXDraw.ISOmetricViewScale, .DisplayHeight * DXDraw.ISOmetricViewScale, .SurfaceOffsetU, .SurfaceOffsetU + .ImageWidthU, .SurfaceOffsetV, .SurfaceOffsetV + .ImageHeightV, .SurfaceColour
        Else
          imageBase = .ActionSequence(.ActionFrame)
          ImageVOffset = (imageBase \ .ImagesPerRow) * .ImageHeightV + .SurfaceOffsetV
          ImageUOffset = (imageBase Mod .ImagesPerRow) * .ImageWidthU + .SurfaceOffsetU
          
          DXDraw.BlitTexture .SurfaceIndex, displayPosX + .DisplayOffsetX * DXDraw.ISOmetricViewScale, displayPosY + .DisplayOffsetY * DXDraw.ISOmetricViewScale, .DisplayWidth * DXDraw.ISOmetricViewScale, .DisplayHeight * DXDraw.ISOmetricViewScale, ImageUOffset, ImageUOffset + .ImageWidthU, ImageVOffset, ImageVOffset + .ImageHeightV, .SurfaceColour
        End If
      End If
      
      Set objectSurface = .NextSurface
    End With
  Loop
End Sub

Public Sub Render_ObjectGhosted()
  Dim objectSurface As clsISO_ObjectSurface, imageBase As Long, ImageUOffset As Single, ImageVOffset As Single
  
  Set objectSurface = SurfaceList
  
  Do While Not (objectSurface Is Nothing)
    With objectSurface
      If .Visible Then
        If .ActionFrame = -1 Then
          DXDraw.BlitTexture .SurfaceIndex, displayPosX + .DisplayOffsetX * DXDraw.ISOmetricViewScale, displayPosY + .DisplayOffsetY * DXDraw.ISOmetricViewScale, .DisplayWidth * DXDraw.ISOmetricViewScale, .DisplayHeight * DXDraw.ISOmetricViewScale, .SurfaceOffsetU, .SurfaceOffsetU + .ImageWidthU, .SurfaceOffsetV, .SurfaceOffsetV + .ImageHeightV, .SurfaceColour And &H80FFFFFF
        Else
          imageBase = .ActionSequence(.ActionFrame)
          ImageVOffset = (imageBase \ .ImagesPerRow) * .ImageHeightV + .SurfaceOffsetV
          ImageUOffset = (imageBase Mod .ImagesPerRow) * .ImageWidthU + .SurfaceOffsetU
          
          DXDraw.BlitTexture .SurfaceIndex, displayPosX + .DisplayOffsetX * DXDraw.ISOmetricViewScale, displayPosY + .DisplayOffsetY * DXDraw.ISOmetricViewScale, .DisplayWidth * DXDraw.ISOmetricViewScale, .DisplayHeight * DXDraw.ISOmetricViewScale, ImageUOffset, ImageUOffset + .ImageWidthU, ImageVOffset, ImageVOffset + .ImageHeightV, .SurfaceColour And &H80FFFFFF
        End If
      End If
      
      Set objectSurface = .NextSurface
    End With
  Loop
End Sub

Public Sub Render_ObjectShadowed()
  Dim objectSurface As clsISO_ObjectSurface, imageBase As Long, ImageUOffset As Single, ImageVOffset As Single
  
  Set objectSurface = SurfaceList
  
  Do While Not (objectSurface Is Nothing)
    With objectSurface
      If .Visible Then
        If .ActionFrame = -1 Then
          DXDraw.BlitTexture .SurfaceIndex, displayPosX + .DisplayOffsetX * DXDraw.ISOmetricViewScale, displayPosY + .DisplayOffsetY * DXDraw.ISOmetricViewScale, .DisplayWidth * DXDraw.ISOmetricViewScale, .DisplayHeight * DXDraw.ISOmetricViewScale, .SurfaceOffsetU, .SurfaceOffsetU + .ImageWidthU, .SurfaceOffsetV, .SurfaceOffsetV + .ImageHeightV, &H40000000
        Else
          imageBase = .ActionSequence(.ActionFrame)
          ImageVOffset = (imageBase \ .ImagesPerRow) * .ImageHeightV + .SurfaceOffsetV
          ImageUOffset = (imageBase Mod .ImagesPerRow) * .ImageWidthU + .SurfaceOffsetU
          
          DXDraw.BlitTexture .SurfaceIndex, displayPosX + .DisplayOffsetX * DXDraw.ISOmetricViewScale, displayPosY + .DisplayOffsetY * DXDraw.ISOmetricViewScale, .DisplayWidth * DXDraw.ISOmetricViewScale, .DisplayHeight * DXDraw.ISOmetricViewScale, ImageUOffset, ImageUOffset + .ImageWidthU, ImageVOffset, ImageVOffset + .ImageHeightV, &H40000000
        End If
      End If
      
      Set objectSurface = .NextSurface
    End With
  Loop
End Sub

'translates the world XYZ coordinates to the display ISOmetric coordinates
Public Sub Translate_Object(ByVal BasePosX_1000ths As Long, ByVal BasePosY_1000ths As Long, ByVal BasePosZ_1000ths As Long)
  Dim effX As Long, effY As Long
  
  effX = PosX_1000ths - BasePosX_1000ths
  effY = PosY_1000ths - BasePosY_1000ths
  
  displayPosX = (DXDraw.m_ClippingRectangleWidth \ 2) + ((effX - effY) * DXDraw.ISOmetricViewScale) \ 1000
  displayPosY = (((effX + effY) \ 2 - (PosZ_1000ths - BasePosZ_1000ths)) * DXDraw.ISOmetricViewScale) \ 1000
End Sub

Public Sub Scale_Object(ByVal xScale As Single, ByVal yScale As Single)
  Dim sList As clsISO_ObjectSurface
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .DisplayWidth = .DisplayWidth * xScale
      .DisplayOffsetX = .DisplayOffsetX * xScale
      
      .DisplayHeight = .DisplayHeight * yScale
      .DisplayOffsetY = .DisplayOffsetY * yScale
      
      Set sList = .NextSurface
    End With
  Loop
End Sub

Public Sub Calculate_CollsionBox()
  Collision_XLower = PosX_1000ths - Collision_WidthX_1000ths
  Collision_XUpper = PosX_1000ths + Collision_WidthX_1000ths
  Collision_YLower = PosY_1000ths - Collision_WidthY_1000ths
  Collision_YUpper = PosY_1000ths + Collision_WidthY_1000ths
  Collision_ZLower = PosZ_1000ths
  Collision_ZUpper = PosZ_1000ths + Collision_WidthZ_1000ths
End Sub

Public Function Test_Collision(ByVal targetObject As clsISO_Object, Optional ByVal targetMask As Long = &HFFFFFFFF) As Boolean
  With targetObject
    If targetMask And .Collision_Mask Then
      If Collision_XLower <= .Collision_XUpper Then
        If Collision_XUpper >= .Collision_XLower Then
          If Collision_YLower <= .Collision_YUpper Then
            If Collision_YUpper >= .Collision_YLower Then
              If Collision_ZLower <= .Collision_ZUpper Then
                If Collision_ZUpper >= .Collision_ZLower Then Test_Collision = True
              End If
            End If
          End If
        End If
      End If
    End If
  End With
End Function

'get the distance between this object an a point in space
Public Function Get_Distance(ByVal targetXPos_1000ths As Long, ByVal targetYPos_1000ths As Long, Optional ByVal targetZPos_1000ths As Long = 0) As Long
  Get_Distance = Sqr((targetXPos_1000ths - PosX_1000ths) ^ 2 + (targetYPos_1000ths - PosY_1000ths) ^ 2 + (targetZPos_1000ths - PosZ_1000ths) ^ 2)
End Function

'get the XY direction between this object an a point in space (CCW from +Y)
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

'get the angle of elevation (Y axis) between this object an a point in space (-PiByTwo to +PiByTwo)
Public Function Get_Elevation(ByVal targetXPos_1000ths As Long, ByVal targetYPos_1000ths As Long, ByVal targetZPos_1000ths As Long) As Single
  targetXPos_1000ths = Sqr((targetXPos_1000ths - PosX_1000ths) ^ 2 + (targetYPos_1000ths - PosY_1000ths) ^ 2)
  targetZPos_1000ths = targetZPos_1000ths - PosZ_1000ths
  
  If targetXPos_1000ths > 0 Then
    Get_Elevation = Atn(targetZPos_1000ths / targetXPos_1000ths)
  ElseIf targetZPos_1000ths <= 0 Then
    Get_Elevation = -PiByTwo
  Else
    Get_Elevation = PiByTwo
  End If
End Function

'get speed, direction (CCW from +Y axis), and elevation (Z axis) based on component x, y, and z velocities
Public Sub Get_Speed(ByVal xVel_1000ths As Long, ByVal yVel_1000ths As Long, ByVal zVel_1000ths As Long, ByRef speed_1000ths As Long, ByRef direction As Single, elevation As Single)
  Dim speedXY As Long
  
  speedXY = Sqr(xVel_1000ths ^ 2 + yVel_1000ths ^ 2)
  speed_1000ths = Sqr(speedXY + zVel_1000ths ^ 2)
  
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
  
  If speedXY > 0 Then
    elevation = Atn(zVel_1000ths / speedXY)
  ElseIf zVel_1000ths <= 0 Then
    elevation = -PiByTwo
  Else
    elevation = PiByTwo
  End If
End Sub

'get component x, y, and z velocities based on speed, direction (CCW from +Y), and elevation (Z axis)
Public Sub Get_Velocity(ByVal speed_1000ths As Single, ByVal direction As Single, ByVal elevation As Single, ByRef xVel_1000ths As Long, ByRef yVel_1000ths As Long, ByRef zVel_1000ths As Long)
  xVel_1000ths = speed_1000ths * Sin(direction) * Cos(elevation)
  yVel_1000ths = speed_1000ths * Cos(direction) * Cos(elevation)
  zVel_1000ths = speed_1000ths * Sin(elevation)
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

Public Sub Draw_CollisionBox()
  Dim crX As Long, crY As Long
  Dim displayX(0 To 7) As Long, displayY(0 To 7) As Long
  Dim effX As Long, effY As Long, effZ As Long, effXpY As Long, effXmY As Long
  Dim gridVertices(0 To 9) As D3DTLVERTEX
  
  On Error Resume Next
  
  effX = (Collision_WidthX_1000ths * DXDraw.ISOmetricViewScale) \ 1000
  effY = (Collision_WidthY_1000ths * DXDraw.ISOmetricViewScale) \ 1000
  effZ = (Collision_WidthZ_1000ths * DXDraw.ISOmetricViewScale) \ 1000
    
  crX = DXDraw.m_ClippingRectangleX + displayPosX
  crY = DXDraw.m_ClippingRectangleY + displayPosY
  
  effXpY = effX + effY
  effXmY = effX - effY
  
  displayX(0) = crX - effXmY
  displayY(0) = crY - effXpY \ 2
  displayY(1) = displayY(0) - effZ
  
  displayX(2) = crX + effXpY
  displayY(2) = crY + effXmY \ 2
  displayY(3) = displayY(2) - effZ
  
  displayX(4) = crX + effXmY
  displayY(4) = crY + effXpY \ 2
  displayY(5) = displayY(4) - effZ
  
  displayX(6) = crX - effXpY
  displayY(6) = crY - effXmY \ 2
  displayY(7) = displayY(6) - effZ
  
  With gridVertices(0)
    .sx = displayX(0)
    .sy = displayY(1)
    .Color = CollisionBoxColour
    .rhw = 1
  End With
  
  With gridVertices(1)
    .sx = displayX(2)
    .sy = displayY(3)
    .Color = CollisionBoxColour
    .rhw = 1
  End With
  
  With gridVertices(2)
    .sx = displayX(4)
    .sy = displayY(5)
    .Color = CollisionBoxColour
    .rhw = 1
  End With
  
  With gridVertices(3)
    .sx = displayX(6)
    .sy = displayY(7)
    .Color = CollisionBoxColour
    .rhw = 1
  End With
  
  With gridVertices(4)
    .sx = displayX(0)
    .sy = displayY(1)
    .Color = CollisionBoxColour
    .rhw = 1
  End With
  
  With gridVertices(5)
    .sx = displayX(0)
    .sy = displayY(0)
    .rhw = 1
  End With
  
  With gridVertices(6)
    .sx = displayX(2)
    .sy = displayY(2)
    .Color = CollisionBoxColour
    .rhw = 1
  End With
  
  With gridVertices(7)
    .sx = displayX(4)
    .sy = displayY(4)
    .Color = CollisionBoxColour
    .rhw = 1
  End With
  
  With gridVertices(8)
    .sx = displayX(6)
    .sy = displayY(6)
    .Color = CollisionBoxColour
    .rhw = 1
  End With
  
  With gridVertices(9)
    .sx = displayX(0)
    .sy = displayY(0)
    .rhw = 1
  End With
  
  dx_Direct3DDevice.SetTexture 0, Nothing
  dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, gridVertices(0), 10, D3DDP_DEFAULT
  
  With gridVertices(0)
    .sx = displayX(6)
    .sy = displayY(6)
  End With
  
  With gridVertices(1)
    .sx = displayX(6)
    .sy = displayY(7)
  End With
  
  With gridVertices(2)
    .sx = displayX(4)
    .sy = displayY(4)
  End With
  
  With gridVertices(3)
    .sx = displayX(4)
    .sy = displayY(5)
  End With
  
  With gridVertices(4)
    .sx = displayX(2)
    .sy = displayY(2)
  End With
  
  With gridVertices(5)
    .sx = displayX(2)
    .sy = displayY(3)
    .Color = CollisionBoxColour
  End With
  
  dx_Direct3DDevice.DrawPrimitive D3DPT_LINELIST, D3DFVF_TLVERTEX, gridVertices(0), 6, D3DDP_DEFAULT
End Sub

Public Sub ReInitialize_D3D()
  Set dx_DirectX = DXDraw.GetDirectX()
  Set dx_Direct3D = DXDraw.GetDirect3D()
  Set dx_Direct3DDevice = DXDraw.GetDirect3DDevice()
End Sub

Private Sub Class_Initialize()
  ReInitialize_D3D
  
  Visible = True
  visibleInFOW = True
  
  mapChainLevel = -1 ' initially not connected to any ISOmetric maps
End Sub

'Releases references that the object may hold so it can be released by the system
'The object should not exist in any chains or that chain will become corrupt - first remove from chain
Public Sub Destroy_Object()
  Dim AList As clsISO_ObjectSurface
  
  On Error Resume Next
  
  Set dx_DirectX = Nothing
  Set dx_Direct3D = Nothing
  Set dx_Direct3DDevice = Nothing
  
  Set chainNext = Nothing
  Set chainPrevious = Nothing
  Set mapChainNext = Nothing
  Set mapChainPrevious = Nothing
  
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

' Saves ISO object definition to file
Public Sub Save_ToFile(filePathName As String)
  Dim fileHandle As Long
  
  'On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Output As #fileHandle
    Write #fileHandle, "Logicon Enterprises clsISO_Object Description Ver2.6"
    
    Save_ToOpenFile fileHandle
  Close fileHandle
  
  Exit Sub
  
badFile:
  MsgBox "Error writing ISO object to -> " & filePathName, vbCritical, "ERROR"
End Sub

Public Sub Save_ChainToFile(filePathName As String)
  Dim fileHandle As Long, objectList As clsISO_Object, tCount As Long
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Output As #fileHandle
    Write #fileHandle, "Logicon Enterprises clsISO_Chain Description Ver2.6"
    
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
  MsgBox "Error writing ISOmetric object chain to -> " & filePathName, vbCritical, "ERROR"
End Sub

' This should not be called directly... instead call Save_ToFile()
Public Sub Save_ToOpenFile(ByVal fileHandle)
  Dim sList As clsISO_ObjectSurface, tCount As Long
  Dim bitmapSurfaceName As String, bitmapSurfaceWidth As Long, bitmapIsTexture As Boolean
  Dim bitmapSurfaceHeight As Long, bitmapSurfaceTransparency As Long, numberOfBitmapSurfaces As Long
  Dim loop1 As Long, loop2 As Long, staticSurfaceInUse() As Boolean
  
  Write #fileHandle, objectName
  Write #fileHandle, Visible
  Write #fileHandle, visibleInFOW
  
  Write #fileHandle, renderPriority
  
  Write #fileHandle, TypeID
  
  Write #fileHandle, displayPosX
  Write #fileHandle, displayPosY
  
  Write #fileHandle, PosX_1000ths
  Write #fileHandle, PosY_1000ths
  Write #fileHandle, PosZ_1000ths
  
  Write #fileHandle, Collision_Mask
  Write #fileHandle, Collision_WidthX_1000ths
  Write #fileHandle, Collision_WidthY_1000ths
  Write #fileHandle, Collision_WidthZ_1000ths

  Write #fileHandle, Collision_XLower, Collision_XUpper
  Write #fileHandle, Collision_YLower, Collision_YUpper
  Write #fileHandle, Collision_ZLower, Collision_ZUpper
  
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
  Write #fileHandle, VelocityZ_1000ths
  Write #fileHandle, AccelerationX_1000ths
  Write #fileHandle, AccelerationY_1000ths
  Write #fileHandle, AccelerationZ_1000ths
  
  Write #fileHandle, TargetMode
  Write #fileHandle, TargetX_1000ths
  Write #fileHandle, TargetY_1000ths
  Write #fileHandle, TargetZ_1000ths
  
  Write #fileHandle, User1
  Write #fileHandle, User2
  Write #fileHandle, User3
  
  '**********************************************************************************************
End Sub

' Loads ISO object definition from file
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
  MsgBox "Error loading ISO object -> " & filePathName, vbCritical, "ERROR"
End Sub

' Loads ISOmetric user chain definition from file
Public Sub Load_ChainFromFile(filePathName As String)
  Dim fileHandle As Long, newObject As clsISO_Object, lastObject As clsISO_Object
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
      Set newObject = New clsISO_Object
      
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
  MsgBox "Error loading ISOmetric object chain from -> " & filePathName, vbCritical, "ERROR"
End Sub

' This should not be called directly... instead call Load_FromFile or Load_ChainFromFile
Public Sub Load_FromOpenFile(ByVal fileHandle As Long)
  Dim sList As clsISO_ObjectSurface, tCount As Long
  Dim bitmapSurfaceIndex() As Long, bitmapSurfaceName As String, bitmapSurfaceWidth As Long
  Dim bitmapSurfaceHeight As Long, bitmapSurfaceTransparency As Long, numberOfBitmapSurfaces As Long
  Dim loop1 As Long, numberOfSurfacesInUse As Long, surfaceNumber As Long
  
  Input #fileHandle, objectName
  Input #fileHandle, Visible
  Input #fileHandle, visibleInFOW
  
  Input #fileHandle, renderPriority

  Input #fileHandle, TypeID
  
  Input #fileHandle, displayPosX
  Input #fileHandle, displayPosY
  
  Input #fileHandle, PosX_1000ths
  Input #fileHandle, PosY_1000ths
  Input #fileHandle, PosZ_1000ths
  
  Input #fileHandle, Collision_Mask
  Input #fileHandle, Collision_WidthX_1000ths
  Input #fileHandle, Collision_WidthY_1000ths
  Input #fileHandle, Collision_WidthZ_1000ths

  Input #fileHandle, Collision_XLower, Collision_XUpper
  Input #fileHandle, Collision_YLower, Collision_YUpper
  Input #fileHandle, Collision_ZLower, Collision_ZUpper
  
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
    Set sList = New clsISO_ObjectSurface
    
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
  Input #fileHandle, VelocityZ_1000ths
  Input #fileHandle, AccelerationX_1000ths
  Input #fileHandle, AccelerationY_1000ths
  Input #fileHandle, AccelerationZ_1000ths
  
  Input #fileHandle, TargetMode
  Input #fileHandle, TargetX_1000ths
  Input #fileHandle, TargetY_1000ths
  Input #fileHandle, TargetZ_1000ths
  
  Input #fileHandle, User1
  Input #fileHandle, User2
  Input #fileHandle, User3
  
  '**********************************************************************************************
End Sub

'**********************************************************************************************************

'******************************************** user chains *********************************************

'**********************************************************************************************************

Public Sub Calculate_Chain_CollisionBox()
  Dim chainList As clsISO_Object
  
  Set chainList = Me
  
  Do While Not (chainList Is Nothing)
    With chainList
      .Collision_XLower = .PosX_1000ths - .Collision_WidthX_1000ths
      .Collision_XUpper = .PosX_1000ths + .Collision_WidthX_1000ths
      .Collision_YLower = .PosY_1000ths - .Collision_WidthY_1000ths
      .Collision_YUpper = .PosY_1000ths + .Collision_WidthY_1000ths
      .Collision_ZLower = .PosZ_1000ths
      .Collision_ZUpper = .PosZ_1000ths + .Collision_WidthZ_1000ths
      
      Set chainList = .chainNext
    End With
  Loop
End Sub

'removes the ISO object from the chain
Public Sub RemoveFrom_Chain(ByRef chainStart As clsISO_Object)
  If Not (chainNext Is Nothing) Then Set chainNext.chainPrevious = chainPrevious
  
  If chainPrevious Is Nothing Then
    Set chainStart = chainNext
  Else
    Set chainPrevious.chainNext = chainNext
    Set chainPrevious = Nothing
  End If
  
  Set chainNext = Nothing
End Sub

'adds the ISO object to the chain
Public Sub AddTo_Chain(ByRef chainStart As clsISO_Object)
  If chainStart Is Nothing Then
    Set chainNext = Nothing
  Else
    Set chainStart.chainPrevious = Me
    Set chainNext = chainStart
  End If
  
  Set chainPrevious = Nothing
  Set chainStart = Me
End Sub

'Releases references that the ISO objects in the chain may hold so they can be released by the system
Public Sub Destroy_Chain()
  Dim chainList As clsISO_Object
  
  Set chainList = Me
  
  Do While Not (chainList Is Nothing)
    With chainList
      Set chainList = .chainNext
      
      .Destroy_Object
    End With
  Loop
End Sub

Public Function Get_ChainObjectID(ByVal objectID As Long) As clsISO_Object
  Dim chainList As clsISO_Object
  
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

Public Function Get_ChainObjectName(ByVal chainStart As clsISO_Object, ByVal objectName As String, Optional ByVal objectID As Long = 0) As clsISO_Object
  Dim chainList As clsISO_Object
  
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

