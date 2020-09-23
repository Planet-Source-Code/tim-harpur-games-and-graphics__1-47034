VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls3D_Object"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************************
'
' cls3D_Object class definition
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

Public objectName As String                  ' object's name

Public Visible As Boolean                                         ' object's visibility flag (default TRUE)

Public ParentObject As cls3D_Object                 ' parent object or Nothing
Public attachedList As cls3D_Object                   ' list of attached objects
Public chainPrevious As cls3D_Object             ' previous object in render list or attached list
Public chainNext As cls3D_Object                    ' next object in render list or attached list
Public SurfaceList As cls3D_ObjectSurface               ' list of object's surfaces

Private worldMatrix As D3DMATRIX    ' world translation and rotation matrix to be applied to all surfaces of this object

Public TypeID As Long                                             ' object type identifier

Public primaryAxis As Long                         'object's primary axis (x=0, y=1, z=2, billboard=3)
Public PrimaryAxisRotation As Single                  ' object's rotation about primary axis
Public PrimaryAxisTilt As Single                           ' object's primary axis tilt
Public PrimaryAxisPrecession As Single                  ' object's primary axis precession

Public BasePosX As Single                                       ' object's X position (either world - if this is the parent
Public BasePosY As Single                                       ' object's Y position   object or relative - if this is a
Public BasePosZ As Single                                       ' object's Z position    attached object)
  
Public Collision_Mask As Long                                ' object's collision bit mask
Public Collision_WidthX As Single                                  ' X width of collision box
Public Collision_WidthY As Single                                  ' Y width of collision box
Public Collision_WidthZ As Single                                  ' Z width of collision box

Public Collision_XLower As Single, Collision_XUpper As Single         'calculated collision bounds
Public Collision_YLower As Single, Collision_YUpper As Single         'calculated collision bounds
Public Collision_ZLower As Single, Collision_ZUpper As Single         'calculated collision bounds

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
  
Public MainAction As Long                                       ' user defined value
Public TicksPerMainAction As Long                         ' user defined value
Public MainActionTickCounter As Long                   ' user defined value
Public SubAction As Long                                        ' user defined value
Public TicksPerSubAction As Long                         ' user defined value
Public SubActionTickCounter As Long                    ' user defined value
  
Public direction As Long                                             ' user defined value
Public Range As Single                                              ' user defined value
Public Speed As Single                                               ' user defined value

Public VelocityX As Single                                  ' user defined value
Public VelocityY As Single                                  ' user defined value
Public VelocityZ As Single                                  ' user defined value
Public AccelerationX As Single                          ' user defined value
Public AccelerationY As Single                          ' user defined value
Public AccelerationZ As Single                          ' user defined value
  
Public TargetMode As Long                                           ' user defined value
Public TargetX As Single                                  ' user defined value
Public TargetY As Single                                  ' user defined value
Public TargetZ As Single                                  ' user defined value

Public User1 As Long                                                  ' user defined value
Public User2 As Long                                                  ' user defined value
Public User3 As Long                                                  ' user defined value

'***************************************************************************************************************



'recalculates the WorldMatrix based on object's position and rotation if nescessary
'and then calls Update_SurfacePosition() for all attached surfaces
Public Sub Update_ObjectPosition()
  Dim AList As cls3D_Object, matSpinA As D3DMATRIX, matSpinB As D3DMATRIX
  Dim deltaX As Single, deltaY As Single, deltaZ As Single
  
  With dx_DirectX
    .IdentityMatrix worldMatrix
    
    With worldMatrix
      .rc41 = BasePosX
      .rc42 = BasePosY
      .rc43 = BasePosZ
    End With
    
    Select Case primaryAxis
      Case 0
        .RotateXMatrix matSpinA, PrimaryAxisRotation
        .RotateYMatrix matSpinB, PrimaryAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateXMatrix matSpinB, PrimaryAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
      Case 1
        .RotateYMatrix matSpinA, PrimaryAxisRotation
        .RotateZMatrix matSpinB, PrimaryAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateYMatrix matSpinB, PrimaryAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
      Case 2
        .RotateZMatrix matSpinA, PrimaryAxisRotation
        .RotateXMatrix matSpinB, PrimaryAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateZMatrix matSpinB, PrimaryAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
      Case Else 'object is billboard
        dx_Direct3DDevice.GetTransform D3DTRANSFORMSTATE_VIEW, matSpinA
        
        With matSpinA
          deltaX = BasePosX + .rc41
          deltaY = BasePosY + .rc42
          deltaZ = BasePosZ + .rc43
        End With
        
        If deltaZ = 0 Then
          If deltaY > 0 Then
            PrimaryAxisTilt = -Pi / 2
          Else
            PrimaryAxisTilt = Pi / 2
          End If
        Else
          PrimaryAxisTilt = Atn(deltaY / Abs(deltaZ))
        End If
        
        If deltaZ = 0 Then
          If deltaX > 0 Then
            PrimaryAxisPrecession = 1.5 * Pi
          Else
            PrimaryAxisPrecession = 0.5 * Pi
          End If
        ElseIf deltaX > 0 Then
          If deltaZ > 0 Then
            PrimaryAxisPrecession = 2 * Pi - Atn(deltaX / deltaZ)
          Else
            PrimaryAxisPrecession = Pi + Atn(deltaX / -deltaZ)
          End If
        Else
          If deltaZ > 0 Then
            PrimaryAxisPrecession = Atn(deltaX / -deltaZ)
          Else
            PrimaryAxisPrecession = Pi - Atn(deltaX / deltaZ)
          End If
        End If
        
        .RotateXMatrix matSpinA, PrimaryAxisTilt
        .RotateYMatrix matSpinB, PrimaryAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
    End Select
    
    .MatrixMultiply worldMatrix, matSpinA, worldMatrix
  End With
  
  Set AList = attachedList
  
  Do While Not (AList Is Nothing)
    With AList
      .Update_Attached3DObjectPosition worldMatrix
      
      Set AList = .chainNext
    End With
  Loop
End Sub

' this should only be called by the Update_ObjectPosition() routine of the parent object
Public Sub Update_Attached3DObjectPosition(ByRef parentWorldMatrix As D3DMATRIX)
  Dim AList As cls3D_Object, matSpinA As D3DMATRIX, matSpinB As D3DMATRIX
  Dim deltaX As Single, deltaY As Single, deltaZ As Single
  
  With dx_DirectX
    .IdentityMatrix worldMatrix
    
    With worldMatrix
      .rc41 = BasePosX
      .rc42 = BasePosY
      .rc43 = BasePosZ
    End With
    
    Select Case primaryAxis
      Case 0
        .RotateXMatrix matSpinA, PrimaryAxisRotation
        .RotateYMatrix matSpinB, PrimaryAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateXMatrix matSpinB, PrimaryAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
      Case 1
        .RotateYMatrix matSpinA, PrimaryAxisRotation
        .RotateZMatrix matSpinB, PrimaryAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateYMatrix matSpinB, PrimaryAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
      Case 2
        .RotateZMatrix matSpinA, PrimaryAxisRotation
        .RotateXMatrix matSpinB, PrimaryAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateZMatrix matSpinB, PrimaryAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
      Case Else 'object is billboard
        dx_Direct3DDevice.GetTransform D3DTRANSFORMSTATE_VIEW, matSpinA
        
        With matSpinA
          deltaX = (parentWorldMatrix.rc41 + BasePosX) + .rc41
          deltaY = (parentWorldMatrix.rc42 + BasePosY) + .rc42
          deltaZ = (parentWorldMatrix.rc43 + BasePosZ) + .rc43
        End With
        
        If deltaZ = 0 Then
          If deltaY > 0 Then
            PrimaryAxisTilt = -Pi / 2
          Else
            PrimaryAxisTilt = Pi / 2
          End If
        Else
          PrimaryAxisTilt = Atn(deltaY / Abs(deltaZ))
        End If
        
        If deltaZ = 0 Then
          If deltaX > 0 Then
            PrimaryAxisPrecession = 1.5 * Pi
          Else
            PrimaryAxisPrecession = 0.5 * Pi
          End If
        ElseIf deltaX > 0 Then
          If deltaZ > 0 Then
            PrimaryAxisPrecession = 2 * Pi - Atn(deltaX / deltaZ)
          Else
            PrimaryAxisPrecession = Pi + Atn(deltaX / -deltaZ)
          End If
        Else
          If deltaZ > 0 Then
            PrimaryAxisPrecession = Atn(deltaX / -deltaZ)
          Else
            PrimaryAxisPrecession = Pi - Atn(deltaX / deltaZ)
          End If
        End If
        
        .RotateXMatrix matSpinA, PrimaryAxisTilt
        .RotateYMatrix matSpinB, PrimaryAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        
        With parentWorldMatrix
          .rc11 = 1
          .rc12 = 0
          .rc13 = 0
          .rc21 = 0
          .rc22 = 1
          .rc23 = 0
          .rc31 = 0
          .rc32 = 0
          .rc33 = 1
        End With
    End Select
    
    .MatrixMultiply worldMatrix, matSpinA, worldMatrix
    .MatrixMultiply worldMatrix, worldMatrix, parentWorldMatrix
  End With
   
  Set AList = attachedList
  
  Do While Not (AList Is Nothing)
    With AList
      .Update_Attached3DObjectPosition worldMatrix
      
      Set AList = .chainNext
    End With
  Loop
End Sub

Public Sub Calculate_CollsionBox()
  Dim matSpinA As D3DMATRIX, matSpinB As D3DMATRIX, matSpinR As D3DMATRIX
  Dim xMin As Single, xMax As Single, yMin As Single, yMax As Single, zMin As Single, zMax As Single
  
  With dx_DirectX
    Select Case primaryAxis
      Case 0
        .RotateXMatrix matSpinA, PrimaryAxisRotation
        .RotateYMatrix matSpinB, PrimaryAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateXMatrix matSpinB, PrimaryAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
      Case 1
        .RotateYMatrix matSpinA, PrimaryAxisRotation
        .RotateZMatrix matSpinB, PrimaryAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateYMatrix matSpinB, PrimaryAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
      Case 2
        .RotateZMatrix matSpinA, PrimaryAxisRotation
        .RotateXMatrix matSpinB, PrimaryAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateZMatrix matSpinB, PrimaryAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
    End Select
    
    .IdentityMatrix matSpinB
    
    With matSpinB '-,-,-
      .rc41 = -Collision_WidthX
      .rc42 = -Collision_WidthY
      .rc43 = -Collision_WidthZ
    End With
    
    .MatrixMultiply matSpinR, matSpinB, matSpinA
    
    With matSpinR
      xMin = .rc41
      xMax = .rc41
      yMin = .rc42
      yMax = .rc42
      zMin = .rc43
      zMax = .rc43
    End With
    
    With matSpinB '+,-,-
      .rc41 = Collision_WidthX
      .rc42 = -Collision_WidthY
      .rc43 = -Collision_WidthZ
    End With
    
    .MatrixMultiply matSpinR, matSpinB, matSpinA
    
    With matSpinR
      If xMin > .rc41 Then xMin = .rc41
      If xMax < .rc41 Then xMax = .rc41
      If yMin > .rc42 Then yMin = .rc42
      If yMax < .rc42 Then yMax = .rc42
      If zMin > .rc43 Then zMin = .rc43
      If zMax < .rc43 Then zMax = .rc43
    End With
    
    With matSpinB '-,+,-
      .rc41 = -Collision_WidthX
      .rc42 = Collision_WidthY
      .rc43 = -Collision_WidthZ
    End With
    
    .MatrixMultiply matSpinR, matSpinB, matSpinA
    
    With matSpinR
      If xMin > .rc41 Then xMin = .rc41
      If xMax < .rc41 Then xMax = .rc41
      If yMin > .rc42 Then yMin = .rc42
      If yMax < .rc42 Then yMax = .rc42
      If zMin > .rc43 Then zMin = .rc43
      If zMax < .rc43 Then zMax = .rc43
    End With
    
    With matSpinB '+,+,-
      .rc41 = Collision_WidthX
      .rc42 = Collision_WidthY
      .rc43 = -Collision_WidthZ
    End With
    
    .MatrixMultiply matSpinR, matSpinB, matSpinA
    
    With matSpinR
      If xMin > .rc41 Then xMin = .rc41
      If xMax < .rc41 Then xMax = .rc41
      If yMin > .rc42 Then yMin = .rc42
      If yMax < .rc42 Then yMax = .rc42
      If zMin > .rc43 Then zMin = .rc43
      If zMax < .rc43 Then zMax = .rc43
    End With
    
    With matSpinB '-,-,+
      .rc41 = -Collision_WidthX
      .rc42 = -Collision_WidthY
      .rc43 = Collision_WidthZ
    End With
    
    .MatrixMultiply matSpinR, matSpinB, matSpinA
    
    With matSpinR
      If xMin > .rc41 Then xMin = .rc41
      If xMax < .rc41 Then xMax = .rc41
      If yMin > .rc42 Then yMin = .rc42
      If yMax < .rc42 Then yMax = .rc42
      If zMin > .rc43 Then zMin = .rc43
      If zMax < .rc43 Then zMax = .rc43
    End With
    
    With matSpinB '+,-,+
      .rc41 = Collision_WidthX
      .rc42 = -Collision_WidthY
      .rc43 = Collision_WidthZ
    End With
    
    .MatrixMultiply matSpinR, matSpinB, matSpinA
    
    With matSpinR
      If xMin > .rc41 Then xMin = .rc41
      If xMax < .rc41 Then xMax = .rc41
      If yMin > .rc42 Then yMin = .rc42
      If yMax < .rc42 Then yMax = .rc42
      If zMin > .rc43 Then zMin = .rc43
      If zMax < .rc43 Then zMax = .rc43
    End With
    
    With matSpinB '-,+,+
      .rc41 = -Collision_WidthX
      .rc42 = Collision_WidthY
      .rc43 = Collision_WidthZ
    End With
    
    .MatrixMultiply matSpinR, matSpinB, matSpinA
    
    With matSpinR
      If xMin > .rc41 Then xMin = .rc41
      If xMax < .rc41 Then xMax = .rc41
      If yMin > .rc42 Then yMin = .rc42
      If yMax < .rc42 Then yMax = .rc42
      If zMin > .rc43 Then zMin = .rc43
      If zMax < .rc43 Then zMax = .rc43
    End With
    
    With matSpinB '+,+,+
      .rc41 = Collision_WidthX
      .rc42 = Collision_WidthY
      .rc43 = Collision_WidthZ
    End With
    
    .MatrixMultiply matSpinR, matSpinB, matSpinA
    
    With matSpinR
      If xMin > .rc41 Then xMin = .rc41
      If xMax < .rc41 Then xMax = .rc41
      If yMin > .rc42 Then yMin = .rc42
      If yMax < .rc42 Then yMax = .rc42
      If zMin > .rc43 Then zMin = .rc43
      If zMax < .rc43 Then zMax = .rc43
    End With
  End With
  
  Collision_XLower = BasePosX + xMin
  Collision_XUpper = BasePosX + xMax
  Collision_YLower = BasePosY + yMin
  Collision_YUpper = BasePosY + yMax
  Collision_ZLower = BasePosZ + zMin
  Collision_ZUpper = BasePosZ + zMax
End Sub

Public Sub Draw_CollisionBox(ByVal boxColour As Long)
  Dim gridVertices(0 To 7) As D3DLVERTEX, NullMatrix As D3DMATRIX
  
  On Error Resume Next
  
  dx_DirectX.IdentityMatrix NullMatrix
  dx_Direct3DDevice.SetTransform D3DTRANSFORMSTATE_WORLD, NullMatrix
  
  With gridVertices(0)
    .X = Collision_XLower
    .Y = Collision_YLower
    .z = Collision_ZLower
    
    .Color = boxColour
  End With
  
  With gridVertices(1)
    .X = Collision_XUpper
    .Y = Collision_YLower
    .z = Collision_ZLower
    
    .Color = boxColour
  End With
  
  With gridVertices(2)
    .X = Collision_XUpper
    .Y = Collision_YLower
    .z = Collision_ZUpper
    
    .Color = boxColour
  End With
  
  With gridVertices(3)
    .X = Collision_XLower
    .Y = Collision_YLower
    .z = Collision_ZUpper
    
    .Color = boxColour
  End With
  
  With gridVertices(4)
    .X = Collision_XLower
    .Y = Collision_YLower
    .z = Collision_ZLower
    
    .Color = boxColour
  End With
  
  dx_Direct3DDevice.SetTexture 0, Nothing
  dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_LVERTEX, gridVertices(0), 5, D3DDP_DEFAULT
  
  gridVertices(0).Y = Collision_YUpper
  gridVertices(1).Y = Collision_YUpper
  gridVertices(2).Y = Collision_YUpper
  gridVertices(3).Y = Collision_YUpper
  gridVertices(4).Y = Collision_YUpper
  
  dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_LVERTEX, gridVertices(0), 5, D3DDP_DEFAULT
  
  With gridVertices(0)
    .X = Collision_XLower
    .Y = Collision_YLower
    .z = Collision_ZLower
  End With
  
  With gridVertices(1)
    .X = Collision_XLower
    .Y = Collision_YUpper
    .z = Collision_ZLower
  End With
  
  With gridVertices(2)
    .X = Collision_XUpper
    .Y = Collision_YLower
    .z = Collision_ZLower
  End With
  
  With gridVertices(3)
    .X = Collision_XUpper
    .Y = Collision_YUpper
    .z = Collision_ZLower
  End With
  
  With gridVertices(4)
    .X = Collision_XUpper
    .Y = Collision_YLower
    .z = Collision_ZUpper
  End With
  
  With gridVertices(5)
    .X = Collision_XUpper
    .Y = Collision_YUpper
    .z = Collision_ZUpper
    
    .Color = boxColour
  End With
  
  With gridVertices(6)
    .X = Collision_XLower
    .Y = Collision_YLower
    .z = Collision_ZUpper
    
    .Color = boxColour
  End With
  
  With gridVertices(7)
    .X = Collision_XLower
    .Y = Collision_YUpper
    .z = Collision_ZUpper
    
    .Color = boxColour
  End With
  
  dx_Direct3DDevice.DrawPrimitive D3DPT_LINELIST, D3DFVF_LVERTEX, gridVertices(0), 8, D3DDP_DEFAULT
End Sub

Public Function Test_Collision(ByVal targetObject As cls3D_Object, Optional ByVal targetMask As Long = &HFFFFFFFF) As Boolean
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
Public Function Get_Distance(ByVal targetXPos As Single, ByVal targetYPos As Single, ByVal targetZPos As Single) As Single
  Get_Distance = Sqr((targetXPos - BasePosX) ^ 2 + (targetYPos - BasePosY) ^ 2 + (targetZPos - BasePosZ) ^ 2)
End Function

'get the XZ direction between this object an a point in space (CW around +Y axis from +Z)
Public Function Get_Direction(ByVal targetXPos As Single, ByVal targetZPos As Single) As Single
  targetXPos = targetXPos - BasePosX
  targetZPos = targetZPos - BasePosZ
  
  If targetZPos < 0 Then
    Get_Direction = Pi + Atn(targetXPos / targetZPos)
  ElseIf targetZPos > 0 Then
    Get_Direction = Atn(targetXPos / targetZPos)
    
    If Get_Direction < 0 Then Get_Direction = Get_Direction + TwoPi
  ElseIf targetXPos <= 0 Then
    Get_Direction = ThreePiByTwo
  Else
    Get_Direction = PiByTwo
  End If
End Function

'get the angle of elevation (Y axis) between this object an a point in space (-PiByTwo to +PiByTwo)
Public Function Get_Elevation(ByVal targetXPos As Single, ByVal targetYPos As Single, ByVal targetZPos As Single) As Single
  targetXPos = Sqr((targetXPos - BasePosX) ^ 2 + (targetZPos - BasePosZ) ^ 2)
  targetYPos = targetYPos - BasePosY
  
  If targetXPos > 0 Then
    Get_Elevation = Atn(targetYPos / targetXPos)
  ElseIf targetYPos <= 0 Then
    Get_Elevation = -PiByTwo
  Else
    Get_Elevation = PiByTwo
  End If
End Function

'get speed, direction (CCW from +Y axis), and elevation (Z axis) based on component x, y, and z velocities
Public Sub Get_Speed(ByVal xVel As Single, ByVal yVel As Single, ByVal zVel As Single, ByRef Speed As Single, ByRef direction As Single, elevation As Single)
  Dim speedXY As Single
  
  speedXY = Sqr(xVel ^ 2 + yVel ^ 2)
  Speed = Sqr(speedXY + zVel ^ 2)
  
  If yVel < 0 Then
    direction = Pi + Atn(xVel / yVel)
  ElseIf yVel > 0 Then
    direction = Atn(xVel / yVel)
    
    If direction < 0 Then direction = direction + TwoPi
  ElseIf xVel <= 0 Then
    direction = ThreePiByTwo
  Else
    direction = PiByTwo
  End If
  
  If speedXY > 0 Then
    elevation = Atn(zVel / speedXY)
  ElseIf zVel <= 0 Then
    elevation = -PiByTwo
  Else
    elevation = PiByTwo
  End If
End Sub

'get component x, y, and z velocities based on speed, direction (CW around +Y axis from +Z), and elevation (Y axis)
Public Sub Get_Velocity(ByVal Speed As Single, ByVal direction As Single, ByVal elevation As Single, ByRef xVel As Single, ByRef yVel As Single, ByRef zVel As Single)
  xVel = Speed * Sin(direction) * Cos(elevation)
  zVel = Speed * Cos(direction) * Cos(elevation)
  yVel = Speed * Sin(elevation)
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

'Links surface into object's surface chain
Public Sub Add_Surface(ByVal SurfaceObject As cls3D_ObjectSurface)
  With SurfaceObject
    Set .ParentObject = Me
    Set .PreviousSurface = Nothing
    Set .NextSurface = SurfaceList
    
    If Not (SurfaceList Is Nothing) Then Set SurfaceList.PreviousSurface = SurfaceObject
    
    Set SurfaceList = SurfaceObject
  End With
End Sub

'Joins all attached surfaces together into one TriList surface - surface texture and primer are taken
'from first surface in chain
Public Sub Join_AllSurfaces(Optional ByVal newSurfaceID As Long = 1)
  Dim sObject1 As cls3D_ObjectSurface, sObject2 As cls3D_ObjectSurface
  
  Set sObject1 = SurfaceList
  
  If sObject1 Is Nothing Then Exit Sub
  
  sObject1.TypeID = newSurfaceID
  
  Set sObject2 = sObject1.NextSurface
  
  Do While Not (sObject2 Is Nothing)
    Remove_Surface sObject2
    
    sObject1.Join_Surfaces sObject2
    
    Set sObject2 = sObject1.NextSurface
  Loop
End Sub

'Removes the 3DSurface from the surface list and also
'releases references that the object may hold so it can be released by the system if not needed
Public Sub Remove_Surface(ByVal removeObject As cls3D_ObjectSurface)
  With removeObject
    If .PreviousSurface Is Nothing Then
      Set SurfaceList = .NextSurface
    Else
      Set .PreviousSurface.NextSurface = .NextSurface
    End If
    
    If Not (.NextSurface Is Nothing) Then
      Set .NextSurface.PreviousSurface = .PreviousSurface
      Set .NextSurface = Nothing
    End If
    
    Set .PreviousSurface = Nothing
    Set .ParentObject = Nothing
  End With
End Sub

'Removes and destroys all 3DSurfaces from the surface list and also
Public Sub DestroyAll_Surfaces()
  Dim sList As cls3D_ObjectSurface
  
  Set sList = SurfaceList
  Set SurfaceList = Nothing
  
  Do While Not (sList Is Nothing)
    With sList
      Set sList = .NextSurface
      
      .Destroy_Object
    End With
  Loop
End Sub

Public Function Get_Surface(ByVal SurfaceIndex As Long) As cls3D_ObjectSurface
  Dim sList As cls3D_ObjectSurface
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    SurfaceIndex = SurfaceIndex - 1
    
    If SurfaceIndex = 0 Then
      Set Get_Surface = sList
      
      Exit Function
    End If
    
    Set sList = sList.NextSurface
  Loop
End Function

Public Function Get_SurfaceID(ByVal SurfaceTypeID As Long) As cls3D_ObjectSurface
  Dim sList As cls3D_ObjectSurface
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      If SurfaceTypeID = .TypeID Then
        Set Get_SurfaceID = sList
        
        Exit Function
      End If
      
      Set sList = .NextSurface
    End With
  Loop
End Function

Public Function Get_SurfaceCount() As Long
  Dim sList As cls3D_ObjectSurface
  
  Get_SurfaceCount = 0
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    Get_SurfaceCount = Get_SurfaceCount + 1
    
    Set sList = sList.NextSurface
  Loop
End Function

Public Sub Add_Attached3DObject(ByVal addObject As cls3D_Object)
  With addObject
    Set .ParentObject = Me
    Set .chainPrevious = Nothing
    Set .chainNext = attachedList
  End With
  
  If Not (attachedList Is Nothing) Then Set attachedList.chainPrevious = addObject
  Set attachedList = addObject
End Sub

'Removes the object from an attached 3DObject list but
'the object still needs to be destoyed if it is not needed
Public Sub Remove_Attached3DObject(ByVal removeObject As cls3D_Object)
  With removeObject
    If .chainPrevious Is Nothing Then
      Set attachedList = .chainNext
    Else
      Set .chainPrevious.chainNext = .chainNext
    End If
    
    If Not (.chainNext Is Nothing) Then Set .chainNext.chainPrevious = .chainPrevious
    
    Set .chainNext = Nothing
    Set .chainPrevious = Nothing
    Set .ParentObject = Nothing
  End With
End Sub

'Destoys all attached objects and their surfaces
Public Sub DestroyAll_Attached3DObjects()
  Dim aObject As cls3D_Object
  
  Set aObject = attachedList
  Set attachedList = Nothing
  
  Do While Not (aObject Is Nothing)
    With aObject
      Set aObject = .chainNext
      
      .Destroy_Object
    End With
  Loop
End Sub

Public Function Get_Attached3DObject(ByVal attached3DObjectIndex As Long) As cls3D_Object
  Dim AList As cls3D_Object
  
  Set AList = attachedList
  
  Do While Not (AList Is Nothing)
    attached3DObjectIndex = attached3DObjectIndex - 1
    
    If attached3DObjectIndex = 0 Then
      Set Get_Attached3DObject = AList
      
      Exit Function
    End If
    
    Set AList = AList.chainNext
  Loop
End Function

Public Function Get_Attached3DObjectID(ByVal attached3DObjectID As Long) As cls3D_Object
  Dim AList As cls3D_Object
  
  Set AList = attachedList
  
  Do While Not (AList Is Nothing)
    With AList
      If attached3DObjectID = .TypeID Then
        Set Get_Attached3DObjectID = AList
        
        Exit Function
      End If
      
      Set AList = .chainNext
    End With
  Loop
End Function

Public Function Get_Attached3DObjectName(ByVal attached3DObjectName As String, Optional ByVal attached3DObjectID As Long = 0) As cls3D_Object
  Dim AList As cls3D_Object
  
  Set AList = attachedList
  
  If attached3DObjectID = 0 Then
    Do While Not (AList Is Nothing)
      With AList
        If attached3DObjectName = .objectName Then
          Set Get_Attached3DObjectName = AList
          
          Exit Function
        End If
        
        Set AList = .chainNext
      End With
    Loop
  Else
    Do While Not (AList Is Nothing)
      With AList
        If attached3DObjectName = .objectName Then
          If attached3DObjectID = .TypeID Then
            Set Get_Attached3DObjectName = AList
            
            Exit Function
          End If
        End If
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Function

Public Function Get_Attached3DObjectCount() As Long
  Dim AList As cls3D_Object
  
  Set AList = attachedList
  
  Do While Not (AList Is Nothing)
    Get_Attached3DObjectCount = Get_Attached3DObjectCount + 1
    
    Set AList = AList.chainNext
  Loop
End Function

Public Function Get_ChainObjectCount() As Long
  Dim AList As cls3D_Object
  
  Set AList = Me
  
  Do While Not (AList Is Nothing)
    Get_ChainObjectCount = Get_ChainObjectCount + 1
    
    Set AList = AList.chainNext
  Loop
End Function

Public Sub Render_Object(Optional ByVal enableTransparency As Boolean = True, Optional ByVal enableTranslucency As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
  
  On Error Resume Next
  
  If enableTranslucency = False Then dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, False
  If enableTransparency = False Then dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, False
  
  dx_Direct3DDevice.SetTransform D3DTRANSFORMSTATE_WORLD, worldMatrix
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Render_Surface
      
      Set sList = .NextSurface
    End With
  Loop
  
  Set AList = attachedList
  
  Do While Not (AList Is Nothing)
    With AList
      If .Visible Then .Render_Object True, True
      
      Set AList = .chainNext
    End With
  Loop
  
  If enableTranslucency = False Then dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, True
  If enableTransparency = False Then dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, True
End Sub

Public Sub ConvertTo_TriList(Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .ConvertTo_TriList
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .ConvertTo_TriList True
      
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Calculate_Normals(Optional ByVal surfacesAreCircular As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, sList2 As cls3D_ObjectSurface, AList As cls3D_Object
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Calculate_Normals surfacesAreCircular
      
      Set sList = .NextSurface
    End With
  Loop
    
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Calculate_Normals True
      
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Calculate_Vertices(Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Calculate_Vertices
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Calculate_Vertices True
      
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Calculate_VerticesXYZ(Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Calculate_VerticesXYZ
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Calculate_VerticesXYZ True
      
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Calculate_VerticesColour(Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Calculate_VerticesColour
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Calculate_VerticesColour True
      
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Calculate_VerticesUV(Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Calculate_VerticesUV
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Calculate_VerticesUV True
      
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Set_ObjectColour(ByVal objectRed As Single, ByVal objectGreen As Single, ByVal objectBlue As Single, Optional ByVal setAmbient As Boolean = False, Optional ByVal setEmissive As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      If setAmbient Then
        .SurfaceRed = objectRed
        .SurfaceGreen = objectGreen
        .SurfaceBlue = objectBlue
      End If
      
      If setEmissive Then
        .SurfaceEmissiveRed = objectRed
        .SurfaceEmissiveGreen = objectGreen
        .SurfaceEmissiveBlue = objectBlue
      End If
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Set_ObjectColour objectRed, objectGreen, objectBlue, setAmbient, setEmissive, True
      
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Normalize_AllSurfaces(Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Normalize_Surface
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Normalize_AllSurfaces True
      
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Double_ObjectComplexity(Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Double_SurfaceComplexity
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Double_ObjectComplexity True
      
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Apply_VertexAmbient(ByVal sourceX As Single, ByVal sourceY As Single, ByVal sourceZ As Single, _
        Optional ByVal lightRange As Single = 0, Optional ByVal lightFalloff As Single = 0, _
        Optional ByVal redColour As Single = 0, Optional ByVal greenColour As Single = 0, Optional ByVal blueColour As Single = 0, _
        Optional ByVal preclearSurfaceLighting As Boolean = True, Optional ByVal lightIsInWorldSpace As Boolean = False, _
        Optional ByVal includeAttachedObjects As Boolean = True)
        
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
  Dim matrixA As D3DMATRIX, matrixB As D3DMATRIX
  
  If lightIsInWorldSpace Then
    With dx_DirectX
      Select Case primaryAxis
        Case 0
          .RotateXMatrix matrixA, -PrimaryAxisRotation
          .RotateYMatrix matrixB, -PrimaryAxisTilt
          .MatrixMultiply matrixA, matrixA, matrixB
          .RotateXMatrix matrixB, -PrimaryAxisPrecession
          .MatrixMultiply matrixA, matrixA, matrixB
        Case 1
          .RotateYMatrix matrixA, -PrimaryAxisRotation
          .RotateZMatrix matrixB, -PrimaryAxisTilt
          .MatrixMultiply matrixA, matrixA, matrixB
          .RotateYMatrix matrixB, -PrimaryAxisPrecession
          .MatrixMultiply matrixA, matrixA, matrixB
        Case 2
          .RotateZMatrix matrixA, -PrimaryAxisRotation
          .RotateXMatrix matrixB, -PrimaryAxisTilt
          .MatrixMultiply matrixA, matrixA, matrixB
          .RotateZMatrix matrixB, -PrimaryAxisPrecession
          .MatrixMultiply matrixA, matrixA, matrixB
      End Select
      
      .IdentityMatrix matrixB
      
      With matrixB
        .rc41 = sourceX
        .rc42 = sourceY
        .rc43 = sourceZ
      End With
      
      .MatrixMultiply matrixA, matrixB, matrixA
      
      With matrixA
        sourceX = .rc41
        sourceY = .rc42
        sourceZ = .rc43
      End With
    End With
  End If
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Apply_VertexAmbient sourceX, sourceY, sourceZ, lightRange, lightFalloff, redColour, greenColour, blueColour, preclearSurfaceLighting
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Apply_VertexAmbient sourceX, sourceY, sourceZ, lightRange, lightFalloff, redColour, greenColour, blueColour, preclearSurfaceLighting, True, True
      
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Apply_VertexDiffuse(ByVal sourceX As Single, ByVal sourceY As Single, ByVal sourceZ As Single, _
        Optional ByVal lightRange As Single = 0, Optional ByVal lightFalloff As Single = 0, _
        Optional ByVal redColour As Single = 0, Optional ByVal greenColour As Single = 0, Optional ByVal blueColour As Single = 0, _
        Optional ByVal preclearSurfaceLighting As Boolean = True, Optional ByVal lightIsInWorldSpace As Boolean = False, _
        Optional ByVal includeAttachedObjects As Boolean = True)
        
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
  Dim matrixA As D3DMATRIX, matrixB As D3DMATRIX
  
  If lightIsInWorldSpace Then
    With dx_DirectX
      Select Case primaryAxis
        Case 0
          .RotateXMatrix matrixA, -PrimaryAxisRotation
          .RotateYMatrix matrixB, -PrimaryAxisTilt
          .MatrixMultiply matrixA, matrixA, matrixB
          .RotateXMatrix matrixB, -PrimaryAxisPrecession
          .MatrixMultiply matrixA, matrixA, matrixB
        Case 1
          .RotateYMatrix matrixA, -PrimaryAxisRotation
          .RotateZMatrix matrixB, -PrimaryAxisTilt
          .MatrixMultiply matrixA, matrixA, matrixB
          .RotateYMatrix matrixB, -PrimaryAxisPrecession
          .MatrixMultiply matrixA, matrixA, matrixB
        Case 2
          .RotateZMatrix matrixA, -PrimaryAxisRotation
          .RotateXMatrix matrixB, -PrimaryAxisTilt
          .MatrixMultiply matrixA, matrixA, matrixB
          .RotateZMatrix matrixB, -PrimaryAxisPrecession
          .MatrixMultiply matrixA, matrixA, matrixB
      End Select
      
      .IdentityMatrix matrixB
      
      With matrixB
        .rc41 = sourceX
        .rc42 = sourceY
        .rc43 = sourceZ
      End With
      
      .MatrixMultiply matrixA, matrixB, matrixA
      
      With matrixA
        sourceX = .rc41
        sourceY = .rc42
        sourceZ = .rc43
      End With
    End With
  End If
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Apply_VertexDiffuse sourceX, sourceY, sourceZ, lightRange, lightFalloff, redColour, greenColour, blueColour, preclearSurfaceLighting
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Apply_VertexDiffuse sourceX, sourceY, sourceZ, lightRange, lightFalloff, redColour, greenColour, blueColour, preclearSurfaceLighting, True, True
      
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Set_Opacity(Optional ByVal opacity As Single = 1#, Optional includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .SurfaceOpacity = opacity
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Set_Opacity opacity, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Scale_Object(ByVal xScale As Single, ByVal yScale As Single, ByVal zScale As Single, _
        Optional ByVal scaleWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
        
  Dim AList As cls3D_Object, sList As cls3D_ObjectSurface
  
  If scaleWorldPosition Then
    BasePosX = BasePosX * xScale
    BasePosY = BasePosY * yScale
    BasePosZ = BasePosZ * zScale
  End If
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Scale_Object xScale, yScale, zScale
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .BasePosX = .BasePosX * xScale
        .BasePosY = .BasePosY * yScale
        .BasePosZ = .BasePosZ * zScale
        
        .Scale_Object xScale, yScale, zScale
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Shift_Texture(ByVal uShift As Single, ByVal vShift As Single, Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Shift_Texture uShift, vShift
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Shift_Texture uShift, vShift, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Scale_Texture(ByVal uScale As Single, ByVal vScale As Single, Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Scale_Texture uScale, vScale
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Scale_Texture uScale, vScale, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Set_TextureAxisX(ByVal axisScale As Single, Optional ByVal wrapScale As Single = 2 * Pi, _
        Optional ByVal axisShift As Single = 0, Optional ByVal wrapShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
          
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Set_TextureAxisX axisScale, wrapScale, axisShift, wrapShift, FlipUV
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Set_TextureAxisX axisScale, wrapScale, axisShift, wrapShift, FlipUV, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Set_TextureAxisY(ByVal axisScale As Single, Optional ByVal wrapScale As Single = 2 * Pi, _
        Optional ByVal axisShift As Single = 0, Optional ByVal wrapShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
          
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Set_TextureAxisY axisScale, wrapScale, axisShift, wrapShift, FlipUV
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Set_TextureAxisY axisScale, wrapScale, axisShift, wrapShift, FlipUV, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Set_TextureAxisZ(ByVal axisScale As Single, Optional ByVal wrapScale As Single = 2 * Pi, _
        Optional ByVal axisShift As Single = 0, Optional ByVal wrapShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
          
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Set_TextureAxisZ axisScale, wrapScale, axisShift, wrapShift, FlipUV
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Set_TextureAxisZ axisScale, wrapScale, axisShift, wrapShift, FlipUV, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Set_TexturePlaneXY(Optional ByVal uBaseScale As Single = 1, Optional ByVal vBaseScale As Single = 1, _
        Optional ByVal uBaseShift As Single = 0, Optional ByVal vBaseShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False, Optional ByVal AutoFit As Boolean = True, Optional ByVal includeAttachedObjects As Boolean = True)
  
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Set_TexturePlaneXY uBaseScale, vBaseScale, uBaseShift, vBaseShift, FlipUV, AutoFit
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Set_TexturePlaneXY uBaseScale, vBaseScale, uBaseShift, vBaseShift, FlipUV, AutoFit, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Set_TexturePlaneXZ(Optional ByVal uBaseScale As Single = 1, Optional ByVal vBaseScale As Single = 1, _
        Optional ByVal uBaseShift As Single = 0, Optional ByVal vBaseShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False, Optional ByVal AutoFit As Boolean = True, Optional ByVal includeAttachedObjects As Boolean = True)
  
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Set_TexturePlaneXZ uBaseScale, vBaseScale, uBaseShift, vBaseShift, FlipUV, AutoFit
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Set_TexturePlaneXZ uBaseScale, vBaseScale, uBaseShift, vBaseShift, FlipUV, AutoFit, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Set_TexturePlaneYZ(Optional ByVal uBaseScale As Single = 1, Optional ByVal vBaseScale As Single = 1, _
        Optional ByVal uBaseShift As Single = 0, Optional ByVal vBaseShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False, Optional ByVal AutoFit As Boolean = True, Optional ByVal includeAttachedObjects As Boolean = True)
  
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Set_TexturePlaneYZ uBaseScale, vBaseScale, uBaseShift, vBaseShift, FlipUV, AutoFit
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Set_TexturePlaneYZ uBaseScale, vBaseScale, uBaseShift, vBaseShift, FlipUV, AutoFit, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Mirror_yzPlane(Optional ByVal mirrorWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
  
  If mirrorWorldPosition Then
    BasePosX = -BasePosX
    
    Select Case primaryAxis
      Case 0
        PrimaryAxisTilt = -PrimaryAxisTilt
      Case 1
        PrimaryAxisRotation = -PrimaryAxisRotation
        PrimaryAxisTilt = -PrimaryAxisTilt
        PrimaryAxisPrecession = -PrimaryAxisPrecession
      Case 2
        PrimaryAxisRotation = -PrimaryAxisRotation
        PrimaryAxisPrecession = -PrimaryAxisPrecession
    End Select
  End If
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Mirror_yzPlane
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Mirror_yzPlane True, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Mirror_xzPlane(Optional ByVal mirrorWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
    
  If mirrorWorldPosition Then
    BasePosY = -BasePosY
    
    Select Case primaryAxis
      Case 0
        PrimaryAxisRotation = -PrimaryAxisRotation
        PrimaryAxisPrecession = -PrimaryAxisPrecession
      Case 1
        PrimaryAxisTilt = -PrimaryAxisTilt
      Case 2
        PrimaryAxisRotation = -PrimaryAxisRotation
        PrimaryAxisTilt = -PrimaryAxisTilt
        PrimaryAxisPrecession = -PrimaryAxisPrecession
    End Select
  End If
    
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Mirror_xzPlane
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Mirror_xzPlane True, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Mirror_xyPlane(Optional ByVal mirrorWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
      
  If mirrorWorldPosition Then
    BasePosZ = -BasePosZ
    
    Select Case primaryAxis
      Case 0
        PrimaryAxisRotation = -PrimaryAxisRotation
        PrimaryAxisTilt = -PrimaryAxisTilt
        PrimaryAxisPrecession = -PrimaryAxisPrecession
      Case 1
        PrimaryAxisRotation = -PrimaryAxisRotation
        PrimaryAxisPrecession = -PrimaryAxisPrecession
      Case 2
        PrimaryAxisTilt = -PrimaryAxisTilt
    End Select
  End If
      
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Mirror_xyPlane
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Mirror_xyPlane True, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Flip_xAxis90(Optional ByVal flipWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object, tVal As Single
    
  If flipWorldPosition Then
    tVal = -BasePosY
    BasePosY = BasePosZ
    BasePosZ = tVal
    
    Select Case primaryAxis
      Case 0
        PrimaryAxisRotation = PrimaryAxisRotation - Pi / 2
        PrimaryAxisPrecession = PrimaryAxisPrecession + Pi / 2
      Case 1
        primaryAxis = 2
        PrimaryAxisRotation = -PrimaryAxisRotation + Pi / 2
        PrimaryAxisPrecession = -PrimaryAxisPrecession - Pi / 2
      Case 2
        primaryAxis = 1
        PrimaryAxisRotation = PrimaryAxisRotation + Pi / 2
        PrimaryAxisPrecession = PrimaryAxisPrecession - Pi / 2
    End Select
  End If
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Flip_xAxis90
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Flip_xAxis90 True, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Flip_xAxis180(Optional ByVal flipWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Flip_xAxis90 flipWorldPosition, includeAttachedObjects
  Flip_xAxis90 flipWorldPosition, includeAttachedObjects
End Sub

Public Sub Flip_xAxis270(Optional ByVal flipWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Flip_xAxis90 flipWorldPosition, includeAttachedObjects
  Flip_xAxis90 flipWorldPosition, includeAttachedObjects
  Flip_xAxis90 flipWorldPosition, includeAttachedObjects
End Sub

Public Sub Flip_yAxis90(Optional ByVal flipWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object, tVal As Single
    
  If flipWorldPosition Then
    tVal = BasePosX
    BasePosX = -BasePosZ
    BasePosZ = tVal
    
    Select Case primaryAxis
      Case 0
        primaryAxis = 2
        PrimaryAxisRotation = PrimaryAxisRotation + Pi / 2
        PrimaryAxisPrecession = PrimaryAxisPrecession - Pi / 2
      Case 1
        PrimaryAxisRotation = PrimaryAxisRotation - Pi / 2
        PrimaryAxisPrecession = PrimaryAxisPrecession + Pi / 2
      Case 2
        primaryAxis = 0
        PrimaryAxisRotation = -PrimaryAxisRotation + Pi / 2
        PrimaryAxisPrecession = -PrimaryAxisPrecession - Pi / 2
    End Select
  End If
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Flip_yAxis90
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Flip_yAxis90 True, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Flip_yAxis180(Optional ByVal flipWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Flip_yAxis90 flipWorldPosition, includeAttachedObjects
  Flip_yAxis90 flipWorldPosition, includeAttachedObjects
End Sub

Public Sub Flip_yAxis270(Optional ByVal flipWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Flip_yAxis90 flipWorldPosition, includeAttachedObjects
  Flip_yAxis90 flipWorldPosition, includeAttachedObjects
  Flip_yAxis90 flipWorldPosition, includeAttachedObjects
End Sub

Public Sub Flip_zAxis90(Optional ByVal flipWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object, tVal As Single
  
  If flipWorldPosition Then
    tVal = -BasePosX
    BasePosX = BasePosY
    BasePosY = tVal
    
    Select Case primaryAxis
      Case 0
        primaryAxis = 1
        PrimaryAxisRotation = -PrimaryAxisRotation + Pi / 2
        PrimaryAxisPrecession = -PrimaryAxisPrecession - Pi / 2
      Case 1
        primaryAxis = 0
        PrimaryAxisRotation = PrimaryAxisRotation + Pi / 2
        PrimaryAxisPrecession = PrimaryAxisPrecession - Pi / 2
      Case 2
        PrimaryAxisRotation = PrimaryAxisRotation - Pi / 2
        PrimaryAxisPrecession = PrimaryAxisPrecession + Pi / 2
    End Select
  End If
      
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .Flip_zAxis90
      
      Set sList = .NextSurface
    End With
  Loop
  
  If includeAttachedObjects Then
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      With AList
        .Flip_zAxis90 True, True
        
        Set AList = .chainNext
      End With
    Loop
  End If
End Sub

Public Sub Flip_zAxis180(Optional ByVal flipWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Flip_zAxis90 flipWorldPosition, includeAttachedObjects
  Flip_zAxis90 flipWorldPosition, includeAttachedObjects
End Sub

Public Sub Flip_zAxis270(Optional ByVal flipWorldPosition As Boolean = False, Optional ByVal includeAttachedObjects As Boolean = True)
  Flip_zAxis90 flipWorldPosition, includeAttachedObjects
  Flip_zAxis90 flipWorldPosition, includeAttachedObjects
  Flip_zAxis90 flipWorldPosition, includeAttachedObjects
End Sub

'Uses the passed object as a template and adds matching surfaces - uses the object's position and
'rotation values to orientate the surfaces first. This routine does not alter the passed object in any way
Public Sub Add_SurfacesFromTemplate(ByRef templateObject As cls3D_Object, Optional ByVal baseSurfaceID As Long = 0, Optional destroySourceObject As Boolean = False)
  Dim sList As cls3D_ObjectSurface, newSurface As cls3D_ObjectSurface
  Dim arrayTop As Long, loop1 As Long
  
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  Dim matrixA As D3DMATRIX, matrixB As D3DMATRIX
  Dim BasePosX As Single, BasePosY As Single, BasePosZ As Single
  
  On Error Resume Next
  
  With dx_DirectX
    BasePosX = templateObject.BasePosX
    BasePosY = templateObject.BasePosY
    BasePosZ = templateObject.BasePosZ
    
    Select Case templateObject.primaryAxis
      Case 0
        .RotateXMatrix matrixA, templateObject.PrimaryAxisRotation
        .RotateYMatrix matrixB, templateObject.PrimaryAxisTilt
        .MatrixMultiply matrixA, matrixA, matrixB
        .RotateXMatrix matrixB, templateObject.PrimaryAxisPrecession
        .MatrixMultiply matrixA, matrixA, matrixB
      Case 1
        .RotateYMatrix matrixA, templateObject.PrimaryAxisRotation
        .RotateZMatrix matrixB, templateObject.PrimaryAxisTilt
        .MatrixMultiply matrixA, matrixA, matrixB
        .RotateYMatrix matrixB, templateObject.PrimaryAxisPrecession
        .MatrixMultiply matrixA, matrixA, matrixB
      Case 2
        .RotateZMatrix matrixA, templateObject.PrimaryAxisRotation
        .RotateXMatrix matrixB, templateObject.PrimaryAxisTilt
        .MatrixMultiply matrixA, matrixA, matrixB
        .RotateZMatrix matrixB, templateObject.PrimaryAxisPrecession
        .MatrixMultiply matrixA, matrixA, matrixB
      Case Else 'billboards have no rotation
    End Select
  End With
  
  Set sList = templateObject.SurfaceList
  
  Do While Not (sList Is Nothing)
    Set newSurface = New cls3D_ObjectSurface
    
    With sList
      arrayTop = .NumberOfSurfaceVertices - 1
      
      ReDim nSurfaceVertexArrayX(0 To arrayTop)
      ReDim nSurfaceVertexArrayY(0 To arrayTop)
      ReDim nSurfaceVertexArrayZ(0 To arrayTop)
      
      ReDim nSurfaceVertexArrayRed(0 To arrayTop)
      ReDim nSurfaceVertexArrayGreen(0 To arrayTop)
      ReDim nSurfaceVertexArrayBlue(0 To arrayTop)
       
      ReDim nSurfaceVertexArrayNX(0 To arrayTop)
      ReDim nSurfaceVertexArrayNY(0 To arrayTop)
      ReDim nSurfaceVertexArrayNZ(0 To arrayTop)
      
      ReDim nSurfaceTextureArrayU(0 To arrayTop)
      ReDim nSurfaceTextureArrayV(0 To arrayTop)
      
      For loop1 = 0 To arrayTop
        dx_DirectX.IdentityMatrix matrixB
        
        matrixB.rc41 = .SurfaceVertexArrayX(loop1)
        matrixB.rc42 = .SurfaceVertexArrayY(loop1)
        matrixB.rc43 = .SurfaceVertexArrayZ(loop1)
        
        dx_DirectX.MatrixMultiply matrixB, matrixB, matrixA
        
        With matrixB
          nSurfaceVertexArrayX(loop1) = BasePosX + .rc41
          nSurfaceVertexArrayY(loop1) = BasePosY + .rc42
          nSurfaceVertexArrayZ(loop1) = BasePosZ + .rc43
        End With
        
        nSurfaceVertexArrayRed(loop1) = .SurfaceVertexArrayRed(loop1)
        nSurfaceVertexArrayGreen(loop1) = .SurfaceVertexArrayGreen(loop1)
        nSurfaceVertexArrayBlue(loop1) = .SurfaceVertexArrayBlue(loop1)
        
        dx_DirectX.IdentityMatrix matrixB
        
        matrixB.rc41 = .SurfaceVertexArrayNX(loop1)
        matrixB.rc42 = .SurfaceVertexArrayNY(loop1)
        matrixB.rc43 = .SurfaceVertexArrayNZ(loop1)
      
        dx_DirectX.MatrixMultiply matrixB, matrixB, matrixA
      
        nSurfaceVertexArrayNX(loop1) = matrixB.rc41
        nSurfaceVertexArrayNY(loop1) = matrixB.rc42
        nSurfaceVertexArrayNZ(loop1) = matrixB.rc43
        
        nSurfaceTextureArrayU(loop1) = .SurfaceTextureArrayU(loop1)
        nSurfaceTextureArrayV(loop1) = .SurfaceTextureArrayV(loop1)
      Next loop1
    End With
    
    With newSurface 'copy values into newSurface
      .Visible = sList.Visible
      
      .TypeID = sList.TypeID + baseSurfaceID
      .SurfaceType = sList.SurfaceType

      .MainAction = sList.MainAction
      .TicksPerMainAction = sList.TicksPerMainAction
      .MainActionTickCounter = sList.MainActionTickCounter
      .SubAction = sList.SubAction
      .TicksPerSubAction = sList.TicksPerSubAction
      .SubActionTickCounter = sList.SubActionTickCounter
      
      .TextureSurfaceIndex = sList.TextureSurfaceIndex

      .SurfaceRed = sList.SurfaceRed
      .SurfaceGreen = sList.SurfaceGreen
      .SurfaceBlue = sList.SurfaceBlue
      .SurfaceEmissiveRed = sList.SurfaceEmissiveRed
      .SurfaceEmissiveGreen = sList.SurfaceEmissiveGreen
      .SurfaceEmissiveBlue = sList.SurfaceEmissiveBlue
      
      .User1 = sList.User1
      
      .NumberOfSurfaceVertices = sList.NumberOfSurfaceVertices
      .SurfaceVertexArrayX = nSurfaceVertexArrayX
      .SurfaceVertexArrayY = nSurfaceVertexArrayY
      .SurfaceVertexArrayZ = nSurfaceVertexArrayZ
      
      .SurfaceVertexArrayRed = nSurfaceVertexArrayRed
      .SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
      .SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
       
      .SurfaceVertexArrayNX = nSurfaceVertexArrayNX
      .SurfaceVertexArrayNY = nSurfaceVertexArrayNY
      .SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
      
      .SurfaceTextureArrayU = nSurfaceTextureArrayU
      .SurfaceTextureArrayV = nSurfaceTextureArrayV
    End With
    
    Add_Surface newSurface 'join newSurface into object's surface chain
    
    Set sList = sList.NextSurface
  Loop
  
  If destroySourceObject Then templateObject.Destroy_Object
End Sub

'Uses the passed vertex arrays to create a new surface
Public Sub Add_SurfaceFromVertices(ByVal indexedTriListVertexArray As Variant, _
          ByVal vertexArrayX As Variant, ByVal vertexArrayY As Variant, ByVal vertexArrayZ As Variant, _
          Optional vertexArrayRed As Variant, Optional vertexArrayGreen As Variant, Optional vertexArrayBlue As Variant, _
          Optional vertexTextureArrayU As Variant, Optional vertexTextureArrayV As Variant, _
          Optional ByVal R_Primer As Single = 1, Optional ByVal G_Primer As Single = 1, Optional ByVal B_Primer As Single = 1, _
          Optional ByVal textureSurface As Long = 0, Optional ByVal primerEmissiveLevel As Single = 0, _
          Optional ByVal surfaceID As Long = 0)
  
  Dim newSurface As cls3D_ObjectSurface
  Dim arrayTop As Long, loop1 As Long
  
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  
  Set newSurface = New cls3D_ObjectSurface
  
  With newSurface
    arrayTop = UBound(indexedTriListVertexArray, 1)
    
    ReDim nSurfaceVertexArrayX(0 To arrayTop)
    ReDim nSurfaceVertexArrayY(0 To arrayTop)
    ReDim nSurfaceVertexArrayZ(0 To arrayTop)
    
    ReDim nSurfaceVertexArrayRed(0 To arrayTop)
    ReDim nSurfaceVertexArrayGreen(0 To arrayTop)
    ReDim nSurfaceVertexArrayBlue(0 To arrayTop)
     
    ReDim nSurfaceVertexArrayNX(0 To arrayTop)
    ReDim nSurfaceVertexArrayNY(0 To arrayTop)
    ReDim nSurfaceVertexArrayNZ(0 To arrayTop)
    
    ReDim nSurfaceTextureArrayU(0 To arrayTop)
    ReDim nSurfaceTextureArrayV(0 To arrayTop)
    
    For loop1 = 0 To arrayTop
      nSurfaceVertexArrayX(loop1) = vertexArrayX(indexedTriListVertexArray(loop1) - 1)
      nSurfaceVertexArrayY(loop1) = vertexArrayY(indexedTriListVertexArray(loop1) - 1)
      nSurfaceVertexArrayZ(loop1) = vertexArrayZ(indexedTriListVertexArray(loop1) - 1)
    Next loop1
    
    If Not IsMissing(vertexArrayRed) Then
      For loop1 = 0 To arrayTop
        nSurfaceVertexArrayRed(loop1) = vertexArrayRed(indexedTriListVertexArray(loop1) - 1)
      Next loop1
    End If
    
    If Not IsMissing(vertexArrayGreen) Then
      For loop1 = 0 To arrayTop
        nSurfaceVertexArrayGreen(loop1) = vertexArrayGreen(indexedTriListVertexArray(loop1) - 1)
      Next loop1
    End If
    
    If Not IsMissing(vertexArrayBlue) Then
      For loop1 = 0 To arrayTop
        nSurfaceVertexArrayBlue(loop1) = vertexArrayBlue(indexedTriListVertexArray(loop1) - 1)
      Next loop1
    End If
    
    If Not IsMissing(vertexTextureArrayU) Then
      For loop1 = 0 To arrayTop
        nSurfaceTextureArrayU(loop1) = vertexTextureArrayU(indexedTriListVertexArray(loop1) - 1)
      Next loop1
    End If
    
    If Not IsMissing(vertexTextureArrayV) Then
      For loop1 = 0 To arrayTop
        nSurfaceTextureArrayV(loop1) = vertexTextureArrayV(indexedTriListVertexArray(loop1) - 1)
      Next loop1
    End If
    
    .TypeID = surfaceID
    
    .SurfaceBlue = B_Primer
    .SurfaceGreen = G_Primer
    .SurfaceRed = R_Primer
    
    .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
    .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
    .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
    
    .NumberOfSurfaceVertices = arrayTop + 1
    
    .SurfaceVertexArrayX = nSurfaceVertexArrayX
    .SurfaceVertexArrayY = nSurfaceVertexArrayY
    .SurfaceVertexArrayZ = nSurfaceVertexArrayZ
    
    .SurfaceVertexArrayRed = nSurfaceVertexArrayRed
    .SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
    .SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
     
    .SurfaceVertexArrayNX = nSurfaceVertexArrayNX
    .SurfaceVertexArrayNY = nSurfaceVertexArrayNY
    .SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
    
    .SurfaceTextureArrayU = nSurfaceTextureArrayU
    .SurfaceTextureArrayV = nSurfaceTextureArrayV
    
    .Calculate_Normals
  End With
    
  Add_Surface newSurface 'join newSurface into object's surface chain
End Sub

Public Function Duplicate_Object(ByVal newObjectName As String, ByVal newObjectID As Long) As cls3D_Object
  Dim AList As cls3D_Object, sList As cls3D_ObjectSurface
  
  Set Duplicate_Object = New cls3D_Object
  
  With Duplicate_Object
    .objectName = newObjectName
    
    .Visible = Visible

    .TypeID = newObjectID
    
    .BasePosX = BasePosX
    .BasePosY = BasePosY
    .BasePosZ = BasePosZ
    .primaryAxis = primaryAxis
    .PrimaryAxisRotation = PrimaryAxisRotation
    .PrimaryAxisTilt = PrimaryAxisTilt
    .PrimaryAxisPrecession = PrimaryAxisPrecession
  
    .Collision_Mask = Collision_Mask
    .Collision_WidthX = Collision_WidthX
    .Collision_WidthY = Collision_WidthY
    .Collision_WidthZ = Collision_WidthZ
    
    Set sList = SurfaceList
    
    Do While Not (sList Is Nothing)
      .Add_Surface sList.Duplicate_Object
      
      Set sList = sList.NextSurface
    Loop
    
    Set AList = attachedList
    
    Do While Not (AList Is Nothing)
      .Add_Attached3DObject AList.Duplicate_Object(AList.objectName, AList.TypeID)
      
      Set AList = AList.chainNext
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
  
    .direction = direction
    .Range = Range
    .Speed = Speed
    
    .VelocityX = VelocityX
    .VelocityY = VelocityY
    .VelocityZ = VelocityZ
    .AccelerationX = AccelerationX
    .AccelerationY = AccelerationY
    .AccelerationZ = AccelerationZ
  
    .TargetMode = TargetMode
    .TargetX = TargetX
    .TargetY = TargetY
    .TargetZ = TargetZ
    
    .User1 = User1
    .User2 = User2
    .User3 = User3
    
    '**********************************************************************************************
  End With
End Function

Public Sub ReInitialize_D3D()
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object
  
  Set dx_DirectX = DXDraw.GetDirectX()
  Set dx_Direct3D = DXDraw.GetDirect3D()
  Set dx_Direct3DDevice = DXDraw.GetDirect3DDevice()
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    With sList
      .ReInitialize_D3D
      
      Set sList = .NextSurface
    End With
  Loop
  
  Set AList = SurfaceList
  
  Do While Not (AList Is Nothing)
    With AList
      .ReInitialize_D3D
      
      Set AList = .chainNext
    End With
  Loop
End Sub

Private Sub Class_Initialize()
  ReInitialize_D3D
  
  Visible = True
End Sub

'Releases references that the object may hold so it can be released by the system
'The object should not exist in any chains or that chain will become corrupt - first remove from chain
'Does NOT release attached light if one exists - call the DXDraw.RemoveFrom_3DLightChain()
'and DXDraw.Destroy_3DLight() routines first to do this
Public Sub Destroy_Object()
  On Error Resume Next
  
  DestroyAll_Surfaces
  DestroyAll_Attached3DObjects
  
  Set dx_DirectX = Nothing
  Set dx_Direct3D = Nothing
  Set dx_Direct3DDevice = Nothing
  
  Set chainNext = Nothing
  Set chainPrevious = Nothing
  Set ParentObject = Nothing
End Sub

'************************************  Pre-Generated Primitives ********************************

' Creates a 2 sided triangle pane using independent surfaces of type TriList - surfaces are identified by
' TypeID -> Front (ID1), Back (ID2)
Public Sub Create_TriPane(ByVal X_Size As Single, ByVal Y_Height As Single, Optional ByVal X_PeakOffset As Single = 0, _
          Optional ByVal R_Primer As Single = 1, Optional ByVal G_Primer As Single = 1, Optional ByVal B_Primer As Single = 1, _
          Optional ByVal textureSurface_Front As Long = 0, Optional ByVal textureSurface_Back As Long = 0, _
          Optional ByVal twoSided As Boolean = False, Optional ByVal primerEmissiveLevel As Single = 0)
  
  Dim sObject As cls3D_ObjectSurface, loop1 As Long, totalSurfaces As Long
  
  DestroyAll_Surfaces
  
  If twoSided Then
    totalSurfaces = 2
  Else
    totalSurfaces = 1
  End If
  
  X_Size = X_Size / 2
  
  For loop1 = 1 To totalSurfaces
    Set sObject = New cls3D_ObjectSurface
    
    With sObject
      .TypeID = loop1
      
      .SurfaceBlue = B_Primer
      .SurfaceGreen = G_Primer
      .SurfaceRed = R_Primer
    
      .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
      .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
      .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
      
      .SurfaceType = 0
      
      .NumberOfSurfaceVertices = 3
        
      .SurfaceTextureArrayU = Array(0, 0, 0)
      .SurfaceTextureArrayV = Array(0, 0, 0)
      
      If loop1 = 1 Then
        .TextureSurfaceIndex = textureSurface_Front
        
        .SurfaceVertexArrayX = Array(-X_Size, X_PeakOffset, X_Size)
        .SurfaceVertexArrayY = Array(0, Y_Height, 0)
        .SurfaceVertexArrayZ = Array(0, 0, 0)
        
        .SurfaceVertexArrayNX = Array(0, 0, 0)
        .SurfaceVertexArrayNY = Array(0, 0, 0)
        .SurfaceVertexArrayNZ = Array(-1, -1, -1)
          
        .Set_TexturePlaneXY 1, -1
      Else
        .TextureSurfaceIndex = textureSurface_Back
        
        .SurfaceVertexArrayX = Array(X_Size, X_PeakOffset, -X_Size)
        .SurfaceVertexArrayY = Array(0, Y_Height, 0)
        .SurfaceVertexArrayZ = Array(0, 0, 0)
        
        .SurfaceVertexArrayNX = Array(0, 0, 0)
        .SurfaceVertexArrayNY = Array(0, 0, 0)
        .SurfaceVertexArrayNZ = Array(1, 1, 1)
          
        .Set_TexturePlaneXY -1, -1
      End If
      
      .Normalize_Surface
    End With
      
    Add_Surface sObject
  Next loop1
End Sub

' Creates a 2 sided panel using independent surfaces of type TriStrips - surfaces are identified by
' TypeID -> Front (ID1), Back (ID2)
Public Sub Create_Panel(ByVal X_Size As Single, ByVal Y_Size As Single, _
          Optional ByVal R_Primer As Single = 1, Optional ByVal G_Primer As Single = 1, Optional ByVal B_Primer As Single = 1, _
          Optional ByVal textureSurface_Front As Long = 0, Optional ByVal textureSurface_Back As Long = 0, _
          Optional ByVal twoSided As Boolean = False, Optional ByVal primerEmissiveLevel As Single = 0)
  
  Dim sObject As cls3D_ObjectSurface, loop1 As Long, totalSurfaces As Long
  
  DestroyAll_Surfaces
  
  If twoSided Then
    totalSurfaces = 2
  Else
    totalSurfaces = 1
  End If
  
  X_Size = X_Size / 2
  Y_Size = Y_Size / 2
  
  For loop1 = 1 To totalSurfaces
    Set sObject = New cls3D_ObjectSurface
    
    With sObject
      .TypeID = loop1
      
      .SurfaceBlue = B_Primer
      .SurfaceGreen = G_Primer
      .SurfaceRed = R_Primer
      
      .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
      .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
      .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
      
      .SurfaceType = 1
      
      .NumberOfSurfaceVertices = 4
        
      .SurfaceTextureArrayU = Array(0, 0, 0, 0)
      .SurfaceTextureArrayV = Array(0, 0, 0, 0)
      
      If loop1 = 1 Then
        .TextureSurfaceIndex = textureSurface_Front
        
        .SurfaceVertexArrayX = Array(-X_Size, -X_Size, X_Size, X_Size)
        .SurfaceVertexArrayY = Array(-Y_Size, Y_Size, -Y_Size, Y_Size)
        .SurfaceVertexArrayZ = Array(0, 0, 0, 0)
        
        .SurfaceVertexArrayNX = Array(0, 0, 0, 0)
        .SurfaceVertexArrayNY = Array(0, 0, 0, 0)
        .SurfaceVertexArrayNZ = Array(-1, -1, -1, -1)
          
        .Set_TexturePlaneXY 1, -1
      Else
        .TextureSurfaceIndex = textureSurface_Back
        
        .SurfaceVertexArrayX = Array(X_Size, X_Size, -X_Size, -X_Size)
        .SurfaceVertexArrayY = Array(-Y_Size, Y_Size, -Y_Size, Y_Size)
        .SurfaceVertexArrayZ = Array(0, 0, 0, 0)
        
        .SurfaceVertexArrayNX = Array(0, 0, 0, 0)
        .SurfaceVertexArrayNY = Array(0, 0, 0, 0)
        .SurfaceVertexArrayNZ = Array(1, 1, 1, 1)
          
        .Set_TexturePlaneXY -1, -1
      End If
      
      .Normalize_Surface
    End With
      
    Add_Surface sObject
  Next loop1
End Sub

' Creates a box using independent surfaces of type TriStrips - surfaces are identified by TypeID
' Surface Order is Outside-> Bottom (ID1), Top (ID2), Right (ID3), Left (ID4), Back (ID5), Front (ID6)
'                           Inside-> Bottom (ID7), Top (ID8), Right (ID9), Left (ID10), Back (ID11), Front (ID12)
' Set yDiff = 0 to create a wedge, both xDiff and yDiff = 0 to create a 4 sided cone
Public Sub Create_Box(ByVal X_Size As Single, ByVal Y_Size As Single, ByVal Z_Size As Single, _
          Optional ByVal xDiff As Single = 1, Optional ByVal yDiff As Single = 1, Optional ByVal xSkew As Single = 0, Optional ySkew As Single = 0, _
          Optional ByVal R_Primer As Single = 1, Optional ByVal G_Primer As Single = 1, Optional ByVal B_Primer As Single = 1, _
          Optional ByVal textureSurface_Inside As Long = 0, Optional ByVal textureSurface_Outside As Long = 0, _
          Optional ByVal Outside As Boolean = True, Optional ByVal Inside As Boolean = False, Optional ByVal primerEmissiveLevel As Single = 0)
            
  Dim sObject As cls3D_ObjectSurface, loop1 As Long, startS As Long, stopS As Long
  Dim X_SizeR As Single, Y_SizeTop As Single, X_SizeL As Single, Y_SizeBase As Single, Skip As Boolean
  
  DestroyAll_Surfaces
  
  X_Size = X_Size / 2
  Y_Size = Y_Size / 2
  Z_Size = Z_Size / 2
  
  X_SizeR = X_Size * xDiff
  Y_SizeTop = Y_Size * yDiff
  
  X_SizeL = xSkew - X_SizeR
  Y_SizeBase = ySkew - Y_SizeTop
  
  X_SizeR = xSkew + X_SizeR
  Y_SizeTop = ySkew + Y_SizeTop
  
  If Outside Then
    startS = 1
    
    If Inside Then
      stopS = 12
    Else
      stopS = 6
    End If
  Else
    startS = 7
    stopS = 12
  End If
  
  For loop1 = startS To stopS
    If xDiff = 0 Or yDiff = 0 Then
      Select Case loop1
        Case 5, 11
          Skip = True
        Case Else
          Skip = False
      End Select
    End If
    
    If Skip = False Then
      Set sObject = New cls3D_ObjectSurface
      
      With sObject
        .TypeID = loop1
        
        .SurfaceBlue = B_Primer
        .SurfaceGreen = G_Primer
        .SurfaceRed = R_Primer
      
        .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
        .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
        .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
        
        .SurfaceType = 1
        
        If loop1 < 7 Then
          .TextureSurfaceIndex = textureSurface_Outside
        Else
          .TextureSurfaceIndex = textureSurface_Inside
        End If
        
        .NumberOfSurfaceVertices = 4
        
        .SurfaceTextureArrayU = Array(0, 0, 0, 0)
        .SurfaceTextureArrayV = Array(0, 0, 0, 0)
        
        Select Case loop1
        Case 1 'bottom surface outside
          If xDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(X_SizeL, -X_Size, X_Size)
            .SurfaceVertexArrayY = Array(Y_SizeBase, -Y_Size, -Y_Size)
            .SurfaceVertexArrayZ = Array(Z_Size, -Z_Size, -Z_Size)
            
            .SurfaceVertexArrayNX = Array(0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0)
            .SurfaceVertexArrayNY = Array(-1, -1, -1)
          Else
            .SurfaceVertexArrayX = Array(X_SizeL, -X_Size, X_SizeR, X_Size)
            .SurfaceVertexArrayY = Array(Y_SizeBase, -Y_Size, Y_SizeBase, -Y_Size)
            .SurfaceVertexArrayZ = Array(Z_Size, -Z_Size, Z_Size, -Z_Size)
            
            .SurfaceVertexArrayNX = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNY = Array(-1, -1, -1, -1)
          End If
          
          .Set_TexturePlaneXZ -1, -1
        Case 2 'top surface outside
          If xDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(-X_Size, X_SizeL, X_Size)
            .SurfaceVertexArrayY = Array(Y_Size, Y_SizeTop, Y_Size)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size)
           
            .SurfaceVertexArrayNX = Array(0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0)
            .SurfaceVertexArrayNY = Array(1, 1, 1)
          Else
            .SurfaceVertexArrayX = Array(-X_Size, X_SizeL, X_Size, X_SizeR)
            .SurfaceVertexArrayY = Array(Y_Size, Y_SizeTop, Y_Size, Y_SizeTop)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size, Z_Size)
           
            .SurfaceVertexArrayNX = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNY = Array(1, 1, 1, 1)
          End If
          
          .Set_TexturePlaneXZ , -1
        Case 3 'right side surface outside
          If yDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(X_Size, X_SizeR, X_Size)
            .SurfaceVertexArrayY = Array(Y_Size, Y_SizeBase, -Y_Size)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size)
            
            .SurfaceVertexArrayNY = Array(0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0)
            .SurfaceVertexArrayNX = Array(1, 1, 1)
          Else
            .SurfaceVertexArrayX = Array(X_Size, X_SizeR, X_Size, X_SizeR)
            .SurfaceVertexArrayY = Array(Y_Size, Y_SizeTop, -Y_Size, Y_SizeBase)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size, Z_Size)
            
            .SurfaceVertexArrayNY = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNX = Array(1, 1, 1, 1)
          End If
          
          .Set_TexturePlaneYZ , -1, , , True
        Case 4 'left side surface outside
          If yDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(-X_Size, X_SizeL, -X_Size)
            .SurfaceVertexArrayY = Array(-Y_Size, Y_SizeBase, Y_Size)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size)
            
            .SurfaceVertexArrayNY = Array(0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0)
            .SurfaceVertexArrayNX = Array(-1, -1, -1)
          Else
            .SurfaceVertexArrayX = Array(-X_Size, X_SizeL, -X_Size, X_SizeL)
            .SurfaceVertexArrayY = Array(-Y_Size, Y_SizeBase, Y_Size, Y_SizeTop)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size, Z_Size)
            
            .SurfaceVertexArrayNY = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNX = Array(-1, -1, -1, -1)
          End If
          
          .Set_TexturePlaneYZ -1, -1, , , True
        Case 5 'back surface outside
          .SurfaceVertexArrayX = Array(X_SizeR, X_SizeR, X_SizeL, X_SizeL)
          .SurfaceVertexArrayY = Array(Y_SizeBase, Y_SizeTop, Y_SizeBase, Y_SizeTop)
          .SurfaceVertexArrayZ = Array(Z_Size, Z_Size, Z_Size, Z_Size)
          
          .SurfaceVertexArrayNX = Array(0, 0, 0, 0)
          .SurfaceVertexArrayNY = Array(0, 0, 0, 0)
          .SurfaceVertexArrayNZ = Array(1, 1, 1, 1)
          
          .Set_TexturePlaneXY -1, -1
        Case 6 'front surface outside
          .SurfaceVertexArrayX = Array(-X_Size, -X_Size, X_Size, X_Size)
          .SurfaceVertexArrayY = Array(-Y_Size, Y_Size, -Y_Size, Y_Size)
          .SurfaceVertexArrayZ = Array(-Z_Size, -Z_Size, -Z_Size, -Z_Size)
            
          .SurfaceVertexArrayNX = Array(0, 0, 0, 0)
          .SurfaceVertexArrayNY = Array(0, 0, 0, 0)
          .SurfaceVertexArrayNZ = Array(-1, -1, -1, -1)
          
          .Set_TexturePlaneXY , -1
        Case 7 'bottom surface inside
          If xDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(-X_Size, X_SizeL, X_Size)
            .SurfaceVertexArrayY = Array(-Y_Size, Y_SizeBase, -Y_Size)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size)
           
            .SurfaceVertexArrayNX = Array(0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0)
            .SurfaceVertexArrayNY = Array(1, 1, 1)
          Else
            .SurfaceVertexArrayX = Array(-X_Size, X_SizeL, X_Size, X_SizeR)
            .SurfaceVertexArrayY = Array(-Y_Size, Y_SizeBase, -Y_Size, Y_SizeBase)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size, Z_Size)
           
            .SurfaceVertexArrayNX = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNY = Array(1, 1, 1, 1)
          End If
          
          .Set_TexturePlaneXZ , -1
        Case 8 'top surface inside
          If xDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(X_SizeL, -X_Size, X_Size)
            .SurfaceVertexArrayY = Array(Y_SizeTop, Y_Size, Y_Size)
            .SurfaceVertexArrayZ = Array(Z_Size, -Z_Size, -Z_Size)
            
            .SurfaceVertexArrayNX = Array(0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0)
            .SurfaceVertexArrayNY = Array(-1, -1, -1)
          Else
            .SurfaceVertexArrayX = Array(X_SizeL, -X_Size, X_SizeR, X_Size)
            .SurfaceVertexArrayY = Array(Y_SizeTop, Y_Size, Y_SizeTop, Y_Size)
            .SurfaceVertexArrayZ = Array(Z_Size, -Z_Size, Z_Size, -Z_Size)
            
            .SurfaceVertexArrayNX = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNY = Array(-1, -1, -1, -1)
          End If
          
          .Set_TexturePlaneXZ , -1
        Case 9 'right side surface inside
          If yDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(X_Size, X_SizeR, X_Size)
            .SurfaceVertexArrayY = Array(-Y_Size, Y_SizeBase, Y_Size)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size)
            
            .SurfaceVertexArrayNY = Array(0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0)
            .SurfaceVertexArrayNX = Array(-1, -1, -1)
          Else
            .SurfaceVertexArrayX = Array(X_Size, X_SizeR, X_Size, X_SizeR)
            .SurfaceVertexArrayY = Array(-Y_Size, Y_SizeBase, Y_Size, Y_SizeTop)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size, Z_Size)
            
            .SurfaceVertexArrayNY = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNX = Array(-1, -1, -1, -1)
          End If
          
          .Set_TexturePlaneYZ -1, -1, , , True
        Case 10 'left side surface inside
          If yDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(-X_Size, X_SizeL, -X_Size)
            .SurfaceVertexArrayY = Array(Y_Size, Y_SizeTop, -Y_Size)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size)
            
            .SurfaceVertexArrayNY = Array(0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0)
            .SurfaceVertexArrayNX = Array(1, 1, 1)
          Else
            .SurfaceVertexArrayX = Array(-X_Size, X_SizeL, -X_Size, X_SizeL)
            .SurfaceVertexArrayY = Array(Y_Size, Y_SizeTop, -Y_Size, Y_SizeBase)
            .SurfaceVertexArrayZ = Array(-Z_Size, Z_Size, -Z_Size, Z_Size)
            
            .SurfaceVertexArrayNY = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNZ = Array(0, 0, 0, 0)
            .SurfaceVertexArrayNX = Array(1, 1, 1, 1)
          End If
          
          .Set_TexturePlaneYZ , -1, , , True
        Case 11 'back surface inside
          .SurfaceVertexArrayX = Array(X_SizeL, X_SizeL, X_SizeR, X_SizeR)
          .SurfaceVertexArrayY = Array(Y_SizeBase, Y_SizeTop, Y_SizeBase, Y_SizeTop)
          .SurfaceVertexArrayZ = Array(Z_Size, Z_Size, Z_Size, Z_Size)
          
          .SurfaceVertexArrayNX = Array(0, 0, 0, 0)
          .SurfaceVertexArrayNY = Array(0, 0, 0, 0)
          .SurfaceVertexArrayNZ = Array(-1, -1, -1, -1)
          
          .Set_TexturePlaneXY , -1
        Case Else 'front surface inside
          .SurfaceVertexArrayX = Array(X_Size, X_Size, -X_Size, -X_Size)
          .SurfaceVertexArrayY = Array(-Y_Size, Y_Size, -Y_Size, Y_Size)
          .SurfaceVertexArrayZ = Array(-Z_Size, -Z_Size, -Z_Size, -Z_Size)
            
          .SurfaceVertexArrayNX = Array(0, 0, 0, 0)
          .SurfaceVertexArrayNY = Array(0, 0, 0, 0)
          .SurfaceVertexArrayNZ = Array(1, 1, 1, 1)
          
          .Set_TexturePlaneXY -1, -1
        End Select
        
        .Normalize_Surface
      End With
      
      Add_Surface sObject
    End If
  Next loop1
End Sub

' Creates a wedge using independent surfaces of type TriStrips - surfaces are identified by TypeID
' Surface Order is Outside-> Bottom (ID1), Right (ID2), Left (ID3), Back (ID4), Front (ID5)
'                           Inside-> Bottom (ID6), Right (ID7), Left (ID8), Back (ID9), Front (ID10)
' Set yDiff = 0 to create a wedge's wedge, both xDiff and yDiff = 0 to create a 3 sided cone
Public Sub Create_Wedge(ByVal X_Base As Single, ByVal Y_Height As Single, ByVal Z_Length As Single, Optional ByVal X_PeakOffset As Single = 0, _
          Optional ByVal xDiff As Single = 1, Optional ByVal yDiff As Single = 1, Optional ByVal xSkew As Single = 0, Optional ySkew As Single = 0, _
          Optional ByVal R_Primer As Single = 1, Optional ByVal G_Primer As Single = 1, Optional ByVal B_Primer As Single = 1, _
          Optional ByVal textureSurface_Inside As Long = 0, Optional ByVal textureSurface_Outside As Long = 0, _
          Optional ByVal Outside As Boolean = True, Optional ByVal Inside As Boolean = False, Optional ByVal primerEmissiveLevel As Single = 0)
          
  Dim sObject As cls3D_ObjectSurface, loop1 As Long, startS As Long, stopS As Long, X_PeakOffsetBack As Single
  Dim X_BaseR As Single, Y_HeightPeak As Single, X_BaseL As Single, Y_HeightBase As Single, Skip As Boolean
  
  DestroyAll_Surfaces
  
  X_Base = X_Base / 2
  Y_Height = Y_Height / 2
  Z_Length = Z_Length / 2
  
  X_BaseR = X_Base * xDiff
  Y_HeightPeak = Y_Height * yDiff
  
  X_BaseL = xSkew - X_BaseR
  Y_HeightBase = ySkew - Y_HeightPeak
  
  X_BaseR = xSkew + X_BaseR
  Y_HeightPeak = ySkew + Y_HeightPeak
  X_PeakOffsetBack = X_PeakOffset + xSkew
  
  If Outside Then
    startS = 1
    
    If Inside Then
      stopS = 10
    Else
      stopS = 5
    End If
  Else
    startS = 6
    stopS = 10
  End If
  
  For loop1 = startS To stopS
    If xDiff = 0 Or yDiff = 0 Then
      Select Case loop1
        Case 4, 9
          Skip = True
        Case Else
          Skip = False
      End Select
    End If
    
    If Skip = False Then
      Set sObject = New cls3D_ObjectSurface
      
      With sObject
        .TypeID = loop1
        
        .SurfaceBlue = B_Primer
        .SurfaceGreen = G_Primer
        .SurfaceRed = R_Primer
        
        .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
        .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
        .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
        
        .SurfaceType = 1
        
        If loop1 < 6 Then
          .TextureSurfaceIndex = textureSurface_Outside
        Else
          .TextureSurfaceIndex = textureSurface_Inside
        End If
        
        .NumberOfSurfaceVertices = 4
        
        .Normalize_Surface
        
        Select Case loop1
        Case 1 'bottom surface outside
          If xDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(X_BaseL, -X_Base, X_Base)
            .SurfaceVertexArrayY = Array(Y_HeightBase, -Y_Height, -Y_Height)
            .SurfaceVertexArrayZ = Array(Z_Length, -Z_Length, -Z_Length)
            
            .SurfaceVertexArrayNY = Array(-1, -1, -1)
          Else
            .SurfaceVertexArrayX = Array(X_BaseL, -X_Base, X_BaseR, X_Base)
            .SurfaceVertexArrayY = Array(Y_HeightBase, -Y_Height, Y_HeightBase, -Y_Height)
            .SurfaceVertexArrayZ = Array(Z_Length, -Z_Length, Z_Length, -Z_Length)
            
            .SurfaceVertexArrayNY = Array(-1, -1, -1, -1)
          End If
          
          .Set_TexturePlaneXZ -1, -1
        Case 2 'right side surface outside
          If yDiff = 0 And xDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(X_PeakOffset, X_PeakOffsetBack, X_Base)
            .SurfaceVertexArrayY = Array(Y_Height, Y_HeightPeak, -Y_Height)
            .SurfaceVertexArrayZ = Array(-Z_Length, Z_Length, -Z_Length)
            
            .SurfaceVertexArrayNX = Array(1, 1, 1)
          Else
            .SurfaceVertexArrayX = Array(X_PeakOffset, X_PeakOffsetBack, X_Base, X_BaseR)
            .SurfaceVertexArrayY = Array(Y_Height, Y_HeightPeak, -Y_Height, Y_HeightBase)
            .SurfaceVertexArrayZ = Array(-Z_Length, Z_Length, -Z_Length, Z_Length)
            
            .SurfaceVertexArrayNX = Array(1, 1, 1, 1)
          End If
          
          .Set_TexturePlaneYZ , -1, , , True
        Case 3 'left side surface outside
          If yDiff = 0 And xDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(X_PeakOffset, -X_Base, X_BaseL)
            .SurfaceVertexArrayY = Array(Y_Height, -Y_Height, Y_HeightBase)
            .SurfaceVertexArrayZ = Array(-Z_Length, -Z_Length, Z_Length)
            
            .SurfaceVertexArrayNX = Array(-1, -1, -1)
          Else
            .SurfaceVertexArrayX = Array(X_PeakOffset, -X_Base, X_PeakOffsetBack, X_BaseL)
            .SurfaceVertexArrayY = Array(Y_Height, -Y_Height, Y_HeightPeak, Y_HeightBase)
            .SurfaceVertexArrayZ = Array(-Z_Length, -Z_Length, Z_Length, Z_Length)
            
            .SurfaceVertexArrayNX = Array(-1, -1, -1, -1)
          End If
          
          .Set_TexturePlaneYZ -1, -1, , , True
        Case 4 'back surface outside
          .NumberOfSurfaceVertices = 3
          
          .SurfaceVertexArrayX = Array(X_BaseR, X_PeakOffsetBack, X_BaseL)
          .SurfaceVertexArrayY = Array(Y_HeightBase, Y_HeightPeak, Y_HeightBase)
          .SurfaceVertexArrayZ = Array(Z_Length, Z_Length, Z_Length)
          
          .SurfaceVertexArrayNZ = Array(1, 1, 1)
          
          .Set_TexturePlaneXY -1, -1
        Case 5 'front surface outside
          .NumberOfSurfaceVertices = 3
          
          .SurfaceVertexArrayX = Array(-X_Base, X_PeakOffset, X_Base)
          .SurfaceVertexArrayY = Array(-Y_Height, Y_Height, -Y_Height)
          .SurfaceVertexArrayZ = Array(-Z_Length, -Z_Length, -Z_Length)
            
          .SurfaceVertexArrayNZ = Array(-1, -1, -1)
          
          .Set_TexturePlaneXY , -1
        Case 6 'bottom surface inside
          If xDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(-X_Base, X_BaseL, X_Base)
            .SurfaceVertexArrayY = Array(-Y_Height, Y_HeightBase, -Y_Height)
            .SurfaceVertexArrayZ = Array(-Z_Length, Z_Length, -Z_Length)
           
            .SurfaceVertexArrayNY = Array(1, 1, 1)
          Else
            .SurfaceVertexArrayX = Array(-X_Base, X_BaseL, X_Base, X_BaseR)
            .SurfaceVertexArrayY = Array(-Y_Height, Y_HeightBase, -Y_Height, Y_HeightBase)
            .SurfaceVertexArrayZ = Array(-Z_Length, Z_Length, -Z_Length, Z_Length)
           
            .SurfaceVertexArrayNY = Array(1, 1, 1, 1)
          End If
          
          .Set_TexturePlaneXZ , -1
        Case 7 'right side surface inside
          If yDiff = 0 And xDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(X_PeakOffset, X_Base, X_PeakOffsetBack)
            .SurfaceVertexArrayY = Array(Y_Height, -Y_Height, Y_HeightPeak)
            .SurfaceVertexArrayZ = Array(-Z_Length, -Z_Length, Z_Length)
            
            .SurfaceVertexArrayNX = Array(-1, -1, -1)
          Else
            .SurfaceVertexArrayX = Array(X_PeakOffset, X_Base, X_PeakOffsetBack, X_BaseR)
            .SurfaceVertexArrayY = Array(Y_Height, -Y_Height, Y_HeightPeak, Y_HeightBase)
            .SurfaceVertexArrayZ = Array(-Z_Length, -Z_Length, Z_Length, Z_Length)
            
            .SurfaceVertexArrayNX = Array(-1, -1, -1, -1)
          End If
          
          .Set_TexturePlaneYZ -1, -1, , , True
        Case 8 'left side surface inside
          If yDiff = 0 And xDiff = 0 Then
            .NumberOfSurfaceVertices = 3
            
            .SurfaceVertexArrayX = Array(X_PeakOffset, X_PeakOffsetBack, -X_Base)
            .SurfaceVertexArrayY = Array(Y_Height, Y_HeightPeak, -Y_Height)
            .SurfaceVertexArrayZ = Array(-Z_Length, Z_Length, -Z_Length)
            
            .SurfaceVertexArrayNX = Array(1, 1, 1)
          Else
            .SurfaceVertexArrayX = Array(X_PeakOffset, X_PeakOffsetBack, -X_Base, X_BaseL)
            .SurfaceVertexArrayY = Array(Y_Height, Y_HeightPeak, -Y_Height, Y_HeightBase)
            .SurfaceVertexArrayZ = Array(-Z_Length, Z_Length, -Z_Length, Z_Length)
            
            .SurfaceVertexArrayNX = Array(1, 1, 1, 1)
          End If
          
          .Set_TexturePlaneYZ , -1, , , True
        Case 9 'back surface inside
          .NumberOfSurfaceVertices = 3
          
          .SurfaceVertexArrayX = Array(X_BaseL, X_PeakOffsetBack, X_BaseR)
          .SurfaceVertexArrayY = Array(Y_HeightBase, Y_HeightPeak, Y_HeightBase)
          .SurfaceVertexArrayZ = Array(Z_Length, Z_Length, Z_Length)
          
          .SurfaceVertexArrayNZ = Array(-1, -1, -1)
          
          .Set_TexturePlaneXY , -1
        Case Else 'front surface inside
          .NumberOfSurfaceVertices = 3
          
          .SurfaceVertexArrayX = Array(X_Base, X_PeakOffset, -X_Base)
          .SurfaceVertexArrayY = Array(-Y_Height, Y_Height, -Y_Height)
          .SurfaceVertexArrayZ = Array(-Z_Length, -Z_Length, -Z_Length)
            
          .SurfaceVertexArrayNZ = Array(1, 1, 1)
          
          .Set_TexturePlaneXY -1, -1
        End Select
        
        .Calculate_Normals
        
        .Normalize_Surface
      End With
      
      Add_Surface sObject
    End If
  Next loop1
End Sub

' Creates a rod using independent surfaces of type TriStrips - surfaces are identified by TypeID
' Inside Rod (ID1), Outside Rod (ID2), Inside Front (ID3), Outside Front (ID4), Inside Back(ID5)
' Outside Back (ID6)     ... NOTE: if numSides is less than 3 it is set to 3
Public Sub Create_Rod(ByVal numSides As Long, ByVal X_Radius As Single, ByVal Y_Radius As Single, ByVal Z_Length As Single, _
          Optional ByVal xDiff As Single = 1, Optional ByVal yDiff As Single = 1, Optional ByVal xSkew As Single = 0, Optional ySkew As Single = 0, _
          Optional ByVal frontPeak As Single = 0, Optional ByVal backPeak As Single = 0, _
          Optional ByVal R_Primer As Single = 1, Optional ByVal G_Primer As Single = 1, Optional ByVal B_Primer As Single = 1, _
          Optional ByVal textureSurface_Inside As Long = 0, Optional ByVal textureSurface_Outside As Long = 0, _
          Optional ByVal Outside As Boolean = True, Optional ByVal Inside As Boolean = False, Optional ByVal primerEmissiveLevel As Single = 0)
            
  Dim sObject As cls3D_ObjectSurface, loop1 As Long, vertAngle As Single, diffAngle As Single
  Dim X_RadiusBack As Single, Y_RadiusBack As Single, tempVertex As Long, numVertices As Long, numVerticesEnd As Long
  Dim bsObject As cls3D_ObjectSurface, fsObject As cls3D_ObjectSurface
  
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  
  Dim fSurfaceVertexArrayX() As Single, fSurfaceVertexArrayY() As Single, fSurfaceVertexArrayZ() As Single
  Dim fSurfaceVertexArrayNX() As Single, fSurfaceVertexArrayNY() As Single, fSurfaceVertexArrayNZ() As Single
  Dim fSurfaceTextureArrayU() As Single, fSurfaceTextureArrayV() As Single
  
  Dim bSurfaceVertexArrayX() As Single, bSurfaceVertexArrayY() As Single, bSurfaceVertexArrayZ() As Single
  Dim bSurfaceVertexArrayNX() As Single, bSurfaceVertexArrayNY() As Single, bSurfaceVertexArrayNZ() As Single
  Dim bSurfaceTextureArrayU() As Single, bSurfaceTextureArrayV() As Single
  
  DestroyAll_Surfaces
  
  Z_Length = Z_Length / 2
  
  If numSides < 3 Then numSides = 3
  
  diffAngle = (2 * Pi) / numSides
  numVertices = numSides * 2 + 2
  numVerticesEnd = numSides * 3
  
  tempVertex = numVertices - 1
      
  ReDim nSurfaceVertexArrayX(0 To tempVertex)
  ReDim nSurfaceVertexArrayY(0 To tempVertex)
  ReDim nSurfaceVertexArrayZ(0 To tempVertex)
  
  ReDim nSurfaceVertexArrayNX(0 To tempVertex)
  ReDim nSurfaceVertexArrayNY(0 To tempVertex)
  ReDim nSurfaceVertexArrayNZ(0 To tempVertex)
        
  ReDim nSurfaceTextureArrayU(0 To tempVertex)
  ReDim nSurfaceTextureArrayV(0 To tempVertex)
  
  tempVertex = numVerticesEnd - 1
  
  ReDim fSurfaceVertexArrayX(0 To tempVertex)
  ReDim fSurfaceVertexArrayY(0 To tempVertex)
  ReDim fSurfaceVertexArrayZ(0 To tempVertex)
  
  ReDim fSurfaceVertexArrayNX(0 To tempVertex)
  ReDim fSurfaceVertexArrayNY(0 To tempVertex)
  ReDim fSurfaceVertexArrayNZ(0 To tempVertex)
        
  ReDim fSurfaceTextureArrayU(0 To tempVertex)
  ReDim fSurfaceTextureArrayV(0 To tempVertex)
  
  ReDim bSurfaceVertexArrayX(0 To tempVertex)
  ReDim bSurfaceVertexArrayY(0 To tempVertex)
  ReDim bSurfaceVertexArrayZ(0 To tempVertex)
  
  ReDim bSurfaceVertexArrayNX(0 To tempVertex)
  ReDim bSurfaceVertexArrayNY(0 To tempVertex)
  ReDim bSurfaceVertexArrayNZ(0 To tempVertex)
        
  ReDim bSurfaceTextureArrayU(0 To tempVertex)
  ReDim bSurfaceTextureArrayV(0 To tempVertex)
  
  If Inside Then 'inside surface
    If Outside Then
      X_Radius = X_Radius - 0.25
      Y_Radius = Y_Radius - 0.25
    End If
      
    X_RadiusBack = X_Radius * xDiff
    Y_RadiusBack = Y_Radius * yDiff
    
    Set sObject = New cls3D_ObjectSurface
    Set fsObject = New cls3D_ObjectSurface
    Set bsObject = New cls3D_ObjectSurface
    
    With fsObject
      .TypeID = 3
      
      .SurfaceBlue = B_Primer
      .SurfaceGreen = G_Primer
      .SurfaceRed = R_Primer
      
      .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
      .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
      .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
      
      .SurfaceType = 0
      
      .TextureSurfaceIndex = textureSurface_Inside
        
      .NumberOfSurfaceVertices = numVerticesEnd
    End With
    
    With bsObject
      .TypeID = 5
      
      .SurfaceBlue = B_Primer
      .SurfaceGreen = G_Primer
      .SurfaceRed = R_Primer
      
      .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
      .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
      .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
      
      .SurfaceType = 0
      
      .TextureSurfaceIndex = textureSurface_Inside
        
      .NumberOfSurfaceVertices = numVerticesEnd
    End With
    
    With sObject
      .TypeID = 1
      
      .SurfaceBlue = B_Primer
      .SurfaceGreen = G_Primer
      .SurfaceRed = R_Primer
      
      .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
      .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
      .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
      
      .SurfaceType = 1
      
      .TextureSurfaceIndex = textureSurface_Inside
        
      .NumberOfSurfaceVertices = numVertices
    End With
    
    vertAngle = 0
    
    For loop1 = 0 To numVertices - 3 Step 2
      nSurfaceVertexArrayX(loop1) = Sin(vertAngle) * X_Radius
      nSurfaceVertexArrayY(loop1) = Cos(vertAngle) * Y_Radius
      nSurfaceVertexArrayZ(loop1) = -Z_Length
      nSurfaceVertexArrayX(loop1 + 1) = xSkew + Sin(vertAngle) * X_RadiusBack
      nSurfaceVertexArrayY(loop1 + 1) = ySkew + Cos(vertAngle) * Y_RadiusBack
      nSurfaceVertexArrayZ(loop1 + 1) = Z_Length
      
      vertAngle = vertAngle - diffAngle
    Next loop1
    
    ' make sure last vertices match first exactly
    nSurfaceVertexArrayX(numVertices - 2) = nSurfaceVertexArrayX(0)
    nSurfaceVertexArrayY(numVertices - 2) = nSurfaceVertexArrayY(0)
    nSurfaceVertexArrayZ(numVertices - 2) = nSurfaceVertexArrayZ(0)
    nSurfaceVertexArrayX(numVertices - 1) = nSurfaceVertexArrayX(1)
    nSurfaceVertexArrayY(numVertices - 1) = nSurfaceVertexArrayY(1)
    nSurfaceVertexArrayZ(numVertices - 1) = nSurfaceVertexArrayZ(1)
    
    For loop1 = 0 To numSides - 1
      fSurfaceVertexArrayZ(loop1 * 3) = -(Z_Length + frontPeak)
      fSurfaceVertexArrayX(loop1 * 3 + 1) = nSurfaceVertexArrayX(loop1 * 2)
      fSurfaceVertexArrayY(loop1 * 3 + 1) = nSurfaceVertexArrayY(loop1 * 2)
      fSurfaceVertexArrayZ(loop1 * 3 + 1) = -Z_Length
      fSurfaceVertexArrayX(loop1 * 3 + 2) = nSurfaceVertexArrayX(loop1 * 2 + 2)
      fSurfaceVertexArrayY(loop1 * 3 + 2) = nSurfaceVertexArrayY(loop1 * 2 + 2)
      fSurfaceVertexArrayZ(loop1 * 3 + 2) = -Z_Length
      
      bSurfaceVertexArrayX(loop1 * 3) = xSkew
      bSurfaceVertexArrayY(loop1 * 3) = ySkew
      bSurfaceVertexArrayZ(loop1 * 3) = Z_Length + backPeak
      bSurfaceVertexArrayX(loop1 * 3 + 1) = nSurfaceVertexArrayX(loop1 * 2 + 3)
      bSurfaceVertexArrayY(loop1 * 3 + 1) = nSurfaceVertexArrayY(loop1 * 2 + 3)
      bSurfaceVertexArrayZ(loop1 * 3 + 1) = Z_Length
      bSurfaceVertexArrayX(loop1 * 3 + 2) = nSurfaceVertexArrayX(loop1 * 2 + 1)
      bSurfaceVertexArrayY(loop1 * 3 + 2) = nSurfaceVertexArrayY(loop1 * 2 + 1)
      bSurfaceVertexArrayZ(loop1 * 3 + 2) = Z_Length
    Next loop1
    
    With sObject
      .SurfaceVertexArrayX = nSurfaceVertexArrayX
      .SurfaceVertexArrayY = nSurfaceVertexArrayY
      .SurfaceVertexArrayZ = nSurfaceVertexArrayZ
      
      .SurfaceVertexArrayNX = nSurfaceVertexArrayNX
      .SurfaceVertexArrayNY = nSurfaceVertexArrayNY
      .SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
    
      .SurfaceTextureArrayU = nSurfaceTextureArrayU
      .SurfaceTextureArrayV = nSurfaceTextureArrayV
      
      .Normalize_Surface
      .Calculate_Normals True
      .Set_TextureAxisZ -Z_Length * 2, -2 * Pi, Z_Length
    End With
    
    With fsObject
      .SurfaceVertexArrayX = fSurfaceVertexArrayX
      .SurfaceVertexArrayY = fSurfaceVertexArrayY
      .SurfaceVertexArrayZ = fSurfaceVertexArrayZ
      
      .SurfaceVertexArrayNX = fSurfaceVertexArrayNX
      .SurfaceVertexArrayNY = fSurfaceVertexArrayNY
      .SurfaceVertexArrayNZ = fSurfaceVertexArrayNZ
    
      .SurfaceTextureArrayU = fSurfaceTextureArrayU
      .SurfaceTextureArrayV = fSurfaceTextureArrayV
      
      .Normalize_Surface
      .Calculate_Normals True
      
      If frontPeak = 0 Then
        .Set_TexturePlaneXY -1, -1, X_Radius, Y_Radius
      Else
        .Set_TextureAxisZ frontPeak, -2 * Pi, frontPeak / 2
      End If
    End With
    
    With bsObject
      .SurfaceVertexArrayX = bSurfaceVertexArrayX
      .SurfaceVertexArrayY = bSurfaceVertexArrayY
      .SurfaceVertexArrayZ = bSurfaceVertexArrayZ
      
      .SurfaceVertexArrayNX = bSurfaceVertexArrayNX
      .SurfaceVertexArrayNY = bSurfaceVertexArrayNY
      .SurfaceVertexArrayNZ = bSurfaceVertexArrayNZ
    
      .SurfaceTextureArrayU = bSurfaceTextureArrayU
      .SurfaceTextureArrayV = bSurfaceTextureArrayV
      
      .Normalize_Surface
      .Calculate_Normals True
      
      If backPeak = 0 Then
        .Set_TexturePlaneXY , -1, X_Radius, Y_Radius
      Else
        .Set_TextureAxisZ backPeak, -2 * Pi, backPeak / 2
      End If
    End With
    
    Add_Surface sObject
    Add_Surface fsObject
    If xDiff <> 0 And yDiff <> 0 Then Add_Surface bsObject
  End If
  
  If Outside Then 'outside surface
    If Inside Then
      X_Radius = X_Radius + 0.25
      Y_Radius = Y_Radius + 0.25
    End If
      
    X_RadiusBack = X_Radius * xDiff
    Y_RadiusBack = Y_Radius * yDiff
    
    Set sObject = New cls3D_ObjectSurface
    Set fsObject = New cls3D_ObjectSurface
    Set bsObject = New cls3D_ObjectSurface
    
    With fsObject
      .TypeID = 4
      
      .SurfaceBlue = B_Primer
      .SurfaceGreen = G_Primer
      .SurfaceRed = R_Primer
      
      .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
      .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
      .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
      
      .SurfaceType = 0
      
      .TextureSurfaceIndex = textureSurface_Outside
        
      .NumberOfSurfaceVertices = numVerticesEnd
    End With
    
    With bsObject
      .TypeID = 6
      
      .SurfaceBlue = B_Primer
      .SurfaceGreen = G_Primer
      .SurfaceRed = R_Primer
      
      .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
      .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
      .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
      
      .SurfaceType = 0
      
      .TextureSurfaceIndex = textureSurface_Outside
        
      .NumberOfSurfaceVertices = numVerticesEnd
    End With
    
    With sObject
      .TypeID = 2
      
      .SurfaceBlue = B_Primer
      .SurfaceGreen = G_Primer
      .SurfaceRed = R_Primer
      
      .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
      .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
      .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
      
      .SurfaceType = 1
      
      .TextureSurfaceIndex = textureSurface_Outside
        
      .NumberOfSurfaceVertices = numVertices
    End With
    
    vertAngle = 0
      
    For loop1 = 0 To numVertices - 3 Step 2
      nSurfaceVertexArrayX(loop1) = Sin(vertAngle) * X_Radius
      nSurfaceVertexArrayY(loop1) = Cos(vertAngle) * Y_Radius
      nSurfaceVertexArrayZ(loop1) = -Z_Length
      nSurfaceVertexArrayX(loop1 + 1) = xSkew + Sin(vertAngle) * X_RadiusBack
      nSurfaceVertexArrayY(loop1 + 1) = ySkew + Cos(vertAngle) * Y_RadiusBack
      nSurfaceVertexArrayZ(loop1 + 1) = Z_Length
      
      vertAngle = vertAngle + diffAngle
    Next loop1
    
     ' make sure last vertices match first exactly
    nSurfaceVertexArrayX(numVertices - 2) = nSurfaceVertexArrayX(0)
    nSurfaceVertexArrayY(numVertices - 2) = nSurfaceVertexArrayY(0)
    nSurfaceVertexArrayZ(numVertices - 2) = nSurfaceVertexArrayZ(0)
    nSurfaceVertexArrayX(numVertices - 1) = nSurfaceVertexArrayX(1)
    nSurfaceVertexArrayY(numVertices - 1) = nSurfaceVertexArrayY(1)
    nSurfaceVertexArrayZ(numVertices - 1) = nSurfaceVertexArrayZ(1)
    
    For loop1 = 0 To numSides - 1
      fSurfaceVertexArrayZ(loop1 * 3) = -(Z_Length + frontPeak)
      fSurfaceVertexArrayX(loop1 * 3 + 1) = nSurfaceVertexArrayX(loop1 * 2)
      fSurfaceVertexArrayY(loop1 * 3 + 1) = nSurfaceVertexArrayY(loop1 * 2)
      fSurfaceVertexArrayZ(loop1 * 3 + 1) = -Z_Length
      fSurfaceVertexArrayX(loop1 * 3 + 2) = nSurfaceVertexArrayX(loop1 * 2 + 2)
      fSurfaceVertexArrayY(loop1 * 3 + 2) = nSurfaceVertexArrayY(loop1 * 2 + 2)
      fSurfaceVertexArrayZ(loop1 * 3 + 2) = -Z_Length
      
      bSurfaceVertexArrayX(loop1 * 3) = xSkew
      bSurfaceVertexArrayY(loop1 * 3) = ySkew
      bSurfaceVertexArrayZ(loop1 * 3) = Z_Length + backPeak
      bSurfaceVertexArrayX(loop1 * 3 + 1) = nSurfaceVertexArrayX(loop1 * 2 + 3)
      bSurfaceVertexArrayY(loop1 * 3 + 1) = nSurfaceVertexArrayY(loop1 * 2 + 3)
      bSurfaceVertexArrayZ(loop1 * 3 + 1) = Z_Length
      bSurfaceVertexArrayX(loop1 * 3 + 2) = nSurfaceVertexArrayX(loop1 * 2 + 1)
      bSurfaceVertexArrayY(loop1 * 3 + 2) = nSurfaceVertexArrayY(loop1 * 2 + 1)
      bSurfaceVertexArrayZ(loop1 * 3 + 2) = Z_Length
    Next loop1
    
    With sObject
      .SurfaceVertexArrayX = nSurfaceVertexArrayX
      .SurfaceVertexArrayY = nSurfaceVertexArrayY
      .SurfaceVertexArrayZ = nSurfaceVertexArrayZ
      
      .SurfaceVertexArrayNX = nSurfaceVertexArrayNX
      .SurfaceVertexArrayNY = nSurfaceVertexArrayNY
      .SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
    
      .SurfaceTextureArrayU = nSurfaceTextureArrayU
      .SurfaceTextureArrayV = nSurfaceTextureArrayV
      
      .Normalize_Surface
      .Calculate_Normals True
      .Set_TextureAxisZ -Z_Length * 2, 2 * Pi, Z_Length
    End With
    
    With fsObject
      .SurfaceVertexArrayX = fSurfaceVertexArrayX
      .SurfaceVertexArrayY = fSurfaceVertexArrayY
      .SurfaceVertexArrayZ = fSurfaceVertexArrayZ
      
      .SurfaceVertexArrayNX = fSurfaceVertexArrayNX
      .SurfaceVertexArrayNY = fSurfaceVertexArrayNY
      .SurfaceVertexArrayNZ = fSurfaceVertexArrayNZ
    
      .SurfaceTextureArrayU = fSurfaceTextureArrayU
      .SurfaceTextureArrayV = fSurfaceTextureArrayV
      
      .Normalize_Surface
      .Calculate_Normals True
      
      If frontPeak = 0 Then
        .Set_TexturePlaneXY , -1, X_Radius, Y_Radius
      Else
        .Set_TextureAxisZ frontPeak, -2 * Pi, frontPeak / 2
      End If
    End With
    
    With bsObject
      .SurfaceVertexArrayX = bSurfaceVertexArrayX
      .SurfaceVertexArrayY = bSurfaceVertexArrayY
      .SurfaceVertexArrayZ = bSurfaceVertexArrayZ
      
      .SurfaceVertexArrayNX = bSurfaceVertexArrayNX
      .SurfaceVertexArrayNY = bSurfaceVertexArrayNY
      .SurfaceVertexArrayNZ = bSurfaceVertexArrayNZ
    
      .SurfaceTextureArrayU = bSurfaceTextureArrayU
      .SurfaceTextureArrayV = bSurfaceTextureArrayV
      
      .Normalize_Surface
      .Calculate_Normals True
      
      If backPeak = 0 Then
        .Set_TexturePlaneXY -1, -1, X_Radius, Y_Radius
      Else
        .Set_TextureAxisZ backPeak, -2 * Pi, backPeak / 2
      End If
    End With
    
    Add_Surface sObject
    Add_Surface fsObject
    If xDiff <> 0 And yDiff <> 0 Then Add_Surface bsObject
  End If
End Sub

' Creates a sphere using independent surfaces of type TriStrips - surfaces are identified by TypeID
' Inside Sphere TypeIDs (1 to 999),  Outside Sphere TypeIDs (1001 to 1999)
' ... NOTE: if numSides is less than 3 it is set to 3, if numStripsPerHalfSphere is less than 2 it is set to 2
Public Sub Create_Sphere(ByVal numSides As Long, ByVal numStripsPerHalfSphere As Long, _
          ByVal X_Radius As Single, ByVal Y_Radius As Single, ByVal Z_Radius As Single, _
          Optional ByVal R_Primer As Single = 1, Optional ByVal G_Primer As Single = 1, Optional ByVal B_Primer As Single = 1, _
          Optional ByVal textureSurface_Inside As Long = 0, Optional ByVal textureSurface_Outside As Long = 0, _
          Optional ByVal Outside As Boolean = True, Optional ByVal Inside As Boolean = False, _
          Optional ByVal halfSphereOnly As Boolean = False, Optional ByVal primerEmissiveLevel As Single = 0)
          
  Dim loop1 As Long, loop2 As Long, vertAngle As Single, vertAngleZ As Single, diffAngle As Single, diffAngleZ As Single
  Dim vertAngleZNext As Single, X_RadiusTempNext As Single, Y_RadiusTempNext As Single
  Dim X_RadiusTemp As Single, Y_RadiusTemp As Single, tempVertex As Long, numVertices As Long
  Dim bsObject As cls3D_ObjectSurface, fsObject As cls3D_ObjectSurface
  
  Dim fSurfaceVertexArrayX() As Single, fSurfaceVertexArrayY() As Single, fSurfaceVertexArrayZ() As Single
  Dim fSurfaceVertexArrayNX() As Single, fSurfaceVertexArrayNY() As Single, fSurfaceVertexArrayNZ() As Single
  Dim fSurfaceTextureArrayU() As Single, fSurfaceTextureArrayV() As Single
  
  Dim bSurfaceVertexArrayX() As Single, bSurfaceVertexArrayY() As Single, bSurfaceVertexArrayZ() As Single
  Dim bSurfaceVertexArrayNX() As Single, bSurfaceVertexArrayNY() As Single, bSurfaceVertexArrayNZ() As Single
  Dim bSurfaceTextureArrayU() As Single, bSurfaceTextureArrayV() As Single
  
  DestroyAll_Surfaces
  
  If numSides < 3 Then numSides = 3
  If numStripsPerHalfSphere < 2 Then numStripsPerHalfSphere = 2
  
  diffAngle = (2 * Pi) / numSides
  diffAngleZ = (0.4999 * Pi) / numStripsPerHalfSphere
  numVertices = numSides * 2 + 2
  
  tempVertex = numVertices - 1
      
  ReDim fSurfaceVertexArrayX(0 To tempVertex)
  ReDim fSurfaceVertexArrayY(0 To tempVertex)
  ReDim fSurfaceVertexArrayZ(0 To tempVertex)
  
  ReDim fSurfaceVertexArrayNX(0 To tempVertex)
  ReDim fSurfaceVertexArrayNY(0 To tempVertex)
  ReDim fSurfaceVertexArrayNZ(0 To tempVertex)
        
  ReDim fSurfaceTextureArrayU(0 To tempVertex)
  ReDim fSurfaceTextureArrayV(0 To tempVertex)
  
  ReDim bSurfaceVertexArrayX(0 To tempVertex)
  ReDim bSurfaceVertexArrayY(0 To tempVertex)
  ReDim bSurfaceVertexArrayZ(0 To tempVertex)
  
  ReDim bSurfaceVertexArrayNX(0 To tempVertex)
  ReDim bSurfaceVertexArrayNY(0 To tempVertex)
  ReDim bSurfaceVertexArrayNZ(0 To tempVertex)
        
  ReDim bSurfaceTextureArrayU(0 To tempVertex)
  ReDim bSurfaceTextureArrayV(0 To tempVertex)
  
  If Inside Then 'inside surface
    If Outside Then
      X_Radius = X_Radius - 0.25
      Y_Radius = Y_Radius - 0.25
    End If
    
    vertAngleZNext = 0
    
    For loop2 = 1 To numStripsPerHalfSphere
      Set fsObject = New cls3D_ObjectSurface
      Set bsObject = New cls3D_ObjectSurface
      
      With fsObject
        .TypeID = loop2 * 2 - 1
        
        .SurfaceBlue = B_Primer
        .SurfaceGreen = G_Primer
        .SurfaceRed = R_Primer
      
        .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
        .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
        .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
        
        .SurfaceType = 1
        
        .TextureSurfaceIndex = textureSurface_Inside
        
        .NumberOfSurfaceVertices = numVertices
      End With
      
      With bsObject
        .TypeID = loop2 * 2
        
        .SurfaceBlue = B_Primer
        .SurfaceGreen = G_Primer
        .SurfaceRed = R_Primer
      
        .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
        .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
        .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
        
        .SurfaceType = 1
        
        .TextureSurfaceIndex = textureSurface_Inside
        
        .NumberOfSurfaceVertices = numVertices
      End With
      
      vertAngle = 0
      vertAngleZ = vertAngleZNext
      vertAngleZNext = vertAngleZ + diffAngleZ
      X_RadiusTemp = X_Radius * Cos(vertAngleZ)
      Y_RadiusTemp = Y_Radius * Cos(vertAngleZ)
      X_RadiusTempNext = X_Radius * Cos(vertAngleZNext)
      Y_RadiusTempNext = Y_Radius * Cos(vertAngleZNext)
      
      For loop1 = 0 To numVertices - 3 Step 2
        fSurfaceVertexArrayX(loop1) = Sin(vertAngle) * X_RadiusTemp
        fSurfaceVertexArrayY(loop1) = Cos(vertAngle) * Y_RadiusTemp
        fSurfaceVertexArrayZ(loop1) = -Sin(vertAngleZ) * Z_Radius
        fSurfaceVertexArrayX(loop1 + 1) = Sin(vertAngle) * X_RadiusTempNext
        fSurfaceVertexArrayY(loop1 + 1) = Cos(vertAngle) * Y_RadiusTempNext
        fSurfaceVertexArrayZ(loop1 + 1) = -Sin(vertAngleZNext) * Z_Radius
        
        bSurfaceVertexArrayX(loop1 + 1) = Sin(vertAngle) * X_RadiusTemp
        bSurfaceVertexArrayY(loop1 + 1) = Cos(vertAngle) * Y_RadiusTemp
        bSurfaceVertexArrayZ(loop1 + 1) = Sin(vertAngleZ) * Z_Radius
        bSurfaceVertexArrayX(loop1) = Sin(vertAngle) * X_RadiusTempNext
        bSurfaceVertexArrayY(loop1) = Cos(vertAngle) * Y_RadiusTempNext
        bSurfaceVertexArrayZ(loop1) = Sin(vertAngleZNext) * Z_Radius
        
        vertAngle = vertAngle + diffAngle
      Next loop1
      
      ' make sure last vertices match first exactly
      fSurfaceVertexArrayX(numVertices - 2) = fSurfaceVertexArrayX(0)
      fSurfaceVertexArrayY(numVertices - 2) = fSurfaceVertexArrayY(0)
      fSurfaceVertexArrayZ(numVertices - 2) = fSurfaceVertexArrayZ(0)
      fSurfaceVertexArrayX(numVertices - 1) = fSurfaceVertexArrayX(1)
      fSurfaceVertexArrayY(numVertices - 1) = fSurfaceVertexArrayY(1)
      fSurfaceVertexArrayZ(numVertices - 1) = fSurfaceVertexArrayZ(1)
       
      bSurfaceVertexArrayX(numVertices - 2) = bSurfaceVertexArrayX(0)
      bSurfaceVertexArrayY(numVertices - 2) = bSurfaceVertexArrayY(0)
      bSurfaceVertexArrayZ(numVertices - 2) = bSurfaceVertexArrayZ(0)
      bSurfaceVertexArrayX(numVertices - 1) = bSurfaceVertexArrayX(1)
      bSurfaceVertexArrayY(numVertices - 1) = bSurfaceVertexArrayY(1)
      bSurfaceVertexArrayZ(numVertices - 1) = bSurfaceVertexArrayZ(1)
      
      With fsObject
        .SurfaceVertexArrayX = fSurfaceVertexArrayX
        .SurfaceVertexArrayY = fSurfaceVertexArrayY
        .SurfaceVertexArrayZ = fSurfaceVertexArrayZ
        
        .SurfaceVertexArrayNX = fSurfaceVertexArrayNX
        .SurfaceVertexArrayNY = fSurfaceVertexArrayNY
        .SurfaceVertexArrayNZ = fSurfaceVertexArrayNZ
      
        .SurfaceTextureArrayU = fSurfaceTextureArrayU
        .SurfaceTextureArrayV = fSurfaceTextureArrayV
       
        .Normalize_Surface
        .Calculate_Normals True
        .Set_TextureAxisZ Z_Radius * 2, 2 * Pi
      End With
      
      With bsObject
        .SurfaceVertexArrayX = bSurfaceVertexArrayX
        .SurfaceVertexArrayY = bSurfaceVertexArrayY
        .SurfaceVertexArrayZ = bSurfaceVertexArrayZ
        
        .SurfaceVertexArrayNX = bSurfaceVertexArrayNX
        .SurfaceVertexArrayNY = bSurfaceVertexArrayNY
        .SurfaceVertexArrayNZ = bSurfaceVertexArrayNZ
      
        .SurfaceTextureArrayU = bSurfaceTextureArrayU
        .SurfaceTextureArrayV = bSurfaceTextureArrayV
        
        .Normalize_Surface
        .Calculate_Normals True
        .Set_TextureAxisZ Z_Radius * 2, 2 * Pi
      End With
      
      If halfSphereOnly = False Then Add_Surface fsObject
      Add_Surface bsObject
    Next loop2
  End If
  
  If Outside Then 'outside surface
    If Inside Then
      X_Radius = X_Radius + 0.25
      Y_Radius = Y_Radius + 0.25
    End If
      
    vertAngleZNext = 0
    
    For loop2 = 1 To numStripsPerHalfSphere
      Set fsObject = New cls3D_ObjectSurface
      Set bsObject = New cls3D_ObjectSurface
      
      With fsObject
        .TypeID = loop2 * 2 + 999
        
        .SurfaceBlue = B_Primer
        .SurfaceGreen = G_Primer
        .SurfaceRed = R_Primer
      
        .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
        .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
        .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
        
        .SurfaceType = 1
        
        .TextureSurfaceIndex = textureSurface_Outside
        
        .NumberOfSurfaceVertices = numVertices
      End With
      
      With bsObject
        .TypeID = loop2 * 2 + 1000
        
        .SurfaceBlue = B_Primer
        .SurfaceGreen = G_Primer
        .SurfaceRed = R_Primer
      
        .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
        .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
        .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
        
        .SurfaceType = 1
        
        .TextureSurfaceIndex = textureSurface_Outside
        
        .NumberOfSurfaceVertices = numVertices
      End With
      
      vertAngle = 0
      vertAngleZ = vertAngleZNext
      vertAngleZNext = vertAngleZ + diffAngleZ
      X_RadiusTemp = X_Radius * Cos(vertAngleZ)
      Y_RadiusTemp = Y_Radius * Cos(vertAngleZ)
      X_RadiusTempNext = X_Radius * Cos(vertAngleZNext)
      Y_RadiusTempNext = Y_Radius * Cos(vertAngleZNext)
      
      For loop1 = 0 To numVertices - 3 Step 2
        fSurfaceVertexArrayX(loop1) = Sin(vertAngle) * X_RadiusTemp
        fSurfaceVertexArrayY(loop1) = Cos(vertAngle) * Y_RadiusTemp
        fSurfaceVertexArrayZ(loop1) = -Sin(vertAngleZ) * Z_Radius
        fSurfaceVertexArrayX(loop1 + 1) = Sin(vertAngle) * X_RadiusTempNext
        fSurfaceVertexArrayY(loop1 + 1) = Cos(vertAngle) * Y_RadiusTempNext
        fSurfaceVertexArrayZ(loop1 + 1) = -Sin(vertAngleZNext) * Z_Radius
        
        bSurfaceVertexArrayX(loop1 + 1) = Sin(vertAngle) * X_RadiusTemp
        bSurfaceVertexArrayY(loop1 + 1) = Cos(vertAngle) * Y_RadiusTemp
        bSurfaceVertexArrayZ(loop1 + 1) = Sin(vertAngleZ) * Z_Radius
        bSurfaceVertexArrayX(loop1) = Sin(vertAngle) * X_RadiusTempNext
        bSurfaceVertexArrayY(loop1) = Cos(vertAngle) * Y_RadiusTempNext
        bSurfaceVertexArrayZ(loop1) = Sin(vertAngleZNext) * Z_Radius
        
        vertAngle = vertAngle - diffAngle
      Next loop1
      
      ' make sure last vertices match first exactly
      fSurfaceVertexArrayX(numVertices - 2) = fSurfaceVertexArrayX(0)
      fSurfaceVertexArrayY(numVertices - 2) = fSurfaceVertexArrayY(0)
      fSurfaceVertexArrayZ(numVertices - 2) = fSurfaceVertexArrayZ(0)
      fSurfaceVertexArrayX(numVertices - 1) = fSurfaceVertexArrayX(1)
      fSurfaceVertexArrayY(numVertices - 1) = fSurfaceVertexArrayY(1)
      fSurfaceVertexArrayZ(numVertices - 1) = fSurfaceVertexArrayZ(1)
       
      bSurfaceVertexArrayX(numVertices - 2) = bSurfaceVertexArrayX(0)
      bSurfaceVertexArrayY(numVertices - 2) = bSurfaceVertexArrayY(0)
      bSurfaceVertexArrayZ(numVertices - 2) = bSurfaceVertexArrayZ(0)
      bSurfaceVertexArrayX(numVertices - 1) = bSurfaceVertexArrayX(1)
      bSurfaceVertexArrayY(numVertices - 1) = bSurfaceVertexArrayY(1)
      bSurfaceVertexArrayZ(numVertices - 1) = bSurfaceVertexArrayZ(1)
      
      With fsObject
        .SurfaceVertexArrayX = fSurfaceVertexArrayX
        .SurfaceVertexArrayY = fSurfaceVertexArrayY
        .SurfaceVertexArrayZ = fSurfaceVertexArrayZ
        
        .SurfaceVertexArrayNX = fSurfaceVertexArrayNX
        .SurfaceVertexArrayNY = fSurfaceVertexArrayNY
        .SurfaceVertexArrayNZ = fSurfaceVertexArrayNZ
      
        .SurfaceTextureArrayU = fSurfaceTextureArrayU
        .SurfaceTextureArrayV = fSurfaceTextureArrayV
        
        .Normalize_Surface
        .Calculate_Normals True
        .Set_TextureAxisZ Z_Radius * 2, 2 * Pi
      End With
      
      With bsObject
        .SurfaceVertexArrayX = bSurfaceVertexArrayX
        .SurfaceVertexArrayY = bSurfaceVertexArrayY
        .SurfaceVertexArrayZ = bSurfaceVertexArrayZ
        
        .SurfaceVertexArrayNX = bSurfaceVertexArrayNX
        .SurfaceVertexArrayNY = bSurfaceVertexArrayNY
        .SurfaceVertexArrayNZ = bSurfaceVertexArrayNZ
      
        .SurfaceTextureArrayU = bSurfaceTextureArrayU
        .SurfaceTextureArrayV = bSurfaceTextureArrayV
        
        .Normalize_Surface
        .Calculate_Normals True
        .Set_TextureAxisZ Z_Radius * 2, 2 * Pi
      End With
      
      If halfSphereOnly = False Then Add_Surface fsObject
      Add_Surface bsObject
    Next loop2
  End If
End Sub

' Creates a surface that can be used for such things as terrain - includes a random number seed
' to replicate specific results, or 0 to use a system random seed - surface is composed of a TriList
' of (2 ^ (surfaceComplexity + 1)) triangles : allowed surfaceComplexity ranges from 1 to 9
' returns the random seed used to generate surface
Public Function Create_Terrain(ByVal surfaceLengthX As Single, ByVal surfaceWidthZ As Single, ByVal surfaceComplexity As Long, _
          Optional ByVal randSeed As Single = 0, Optional ByVal depthDev As Single = 0, Optional heightDev As Single = 0, _
          Optional ByVal R_Primer As Single = 1, Optional ByVal G_Primer As Single = 1, Optional ByVal B_Primer As Single = 1, _
          Optional ByVal textureSurface As Long = 0, Optional ByVal primerEmissiveLevel As Single = 0) As Single
  
  Dim loop1 As Long, loop2 As Long, sObject As cls3D_ObjectSurface, tDev As Single, temp1 As Single
  Dim xMin As Single, xMax As Single, zMin As Single, zMax As Single
  Dim nVertexY() As Single
  
  DestroyAll_Surfaces
  
  If surfaceComplexity < 1 Then
    surfaceComplexity = 1
  ElseIf surfaceComplexity > 9 Then
    surfaceComplexity = 9
  End If
  
  If randSeed = 0 Then
    Randomize
    
    randSeed = Rnd()
  End If
  
  Rnd -randSeed
  
  xMax = surfaceLengthX / 2
  xMin = -xMax
  zMax = surfaceWidthZ / 2
  zMin = -zMax
  
  tDev = depthDev + heightDev
  
  Set sObject = New cls3D_ObjectSurface
  
  With sObject
    .TypeID = 999
    
    .SurfaceBlue = B_Primer
    .SurfaceGreen = G_Primer
    .SurfaceRed = R_Primer
    
    .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
    .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
    .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
    
    .SurfaceType = 1
    .TextureSurfaceIndex = textureSurface
    
    .NumberOfSurfaceVertices = 4
    
    .SurfaceTextureArrayU = Array(0, 0, 1, 1)
    .SurfaceTextureArrayV = Array(1, 0, 1, 0)
    
    .SurfaceVertexArrayX = Array(xMin, xMin, xMax, xMax)
    .SurfaceVertexArrayY = Array(0, 0, 0, 0)
    .SurfaceVertexArrayZ = Array(zMin, zMax, zMin, zMax)
    
    .SurfaceVertexArrayNX = Array(0, 0, 0, 0)
    .SurfaceVertexArrayNY = Array(1, 1, 1, 1)
    .SurfaceVertexArrayNZ = Array(0, 0, 0, 0)
    .Normalize_Surface
    
    For loop1 = 1 To surfaceComplexity
      .Double_SurfaceComplexity
    Next loop1
    
    nVertexY = .SurfaceVertexArrayY
    
    If tDev > 0 Then
      For loop1 = 0 To .NumberOfSurfaceVertices - 1
        If nVertexY(loop1) = 0 Then
          temp1 = (Rnd() * tDev) - depthDev
          
          For loop2 = 0 To .NumberOfSurfaceVertices - 1
            If .SurfaceVertexArrayX(loop1) = .SurfaceVertexArrayX(loop2) Then
              If .SurfaceVertexArrayZ(loop1) = .SurfaceVertexArrayZ(loop2) Then nVertexY(loop2) = temp1
            End If
          Next loop2
        End If
      Next loop1
    End If
    
    .SurfaceVertexArrayY = nVertexY
    
    .Calculate_Normals
  End With
  
  Add_Surface sObject
  
  Create_Terrain = randSeed
End Function

' Creates a sphere of particles - type pointList
' shapeMode is 0 - block, 1 - rounded, 2 - cylinder on Y-axis, 3 - outside block on Y-axis,
'                         4 - outside rounded, 5 - outside cylinder on Y-axis, 6 - spiral, 7 - doughnut
Public Sub Create_Particles(ByVal X_Radius As Single, ByVal Y_Radius As Single, ByVal Z_Radius As Single, Optional ByVal particleDensity As Single = 0.1, _
          Optional ByVal R_Primer As Single = 1, Optional ByVal G_Primer As Single = 1, Optional ByVal B_Primer As Single = 1, _
          Optional ByVal R_VertexVariance As Single = 0, Optional ByVal G_VertexVariance As Single = 0, Optional ByVal B_VertexVariance As Single = 0, _
          Optional ByVal primerEmissiveLevel As Single = 0, Optional ByVal vertexEmissiveVariance As Single = 0, Optional ByVal shapeMode As Long = 0)
          
  Dim loop1 As Long, numberOfParticles As Long, arrayTop As Long, tVar As Single
  Dim sObject As cls3D_ObjectSurface, angularValue1 As Single, angularValue2 As Single, angularValue3 As Single
  Dim radiusValue As Single, tRed As Single, tGreen As Single, tBlue As Single
  
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  Dim vertexMode As Boolean
  
  DestroyAll_Surfaces
  
  numberOfParticles = (X_Radius * Y_Radius * Z_Radius) * particleDensity
  
  arrayTop = numberOfParticles - 1
  
  ReDim nSurfaceVertexArrayX(0 To arrayTop)
  ReDim nSurfaceVertexArrayY(0 To arrayTop)
  ReDim nSurfaceVertexArrayZ(0 To arrayTop)

  ReDim nSurfaceVertexArrayRed(0 To arrayTop)
  ReDim nSurfaceVertexArrayGreen(0 To arrayTop)
  ReDim nSurfaceVertexArrayBlue(0 To arrayTop)
    
  ReDim nSurfaceVertexArrayNX(0 To arrayTop)
  ReDim nSurfaceVertexArrayNY(0 To arrayTop)
  ReDim nSurfaceVertexArrayNZ(0 To arrayTop)
  
  ReDim nSurfaceTextureArrayU(0 To arrayTop)
  ReDim nSurfaceTextureArrayV(0 To arrayTop)
  
  Set sObject = New cls3D_ObjectSurface
  
  Randomize
  
  With sObject
    If (R_VertexVariance <> 0) Or (G_VertexVariance <> 0) Or (B_VertexVariance <> 0) Or (vertexEmissiveVariance <> 0) Then
      vertexMode = True
    Else
      vertexMode = False
    End If
    
    For loop1 = 0 To arrayTop
      If vertexMode Then
        tVar = primerEmissiveLevel + vertexEmissiveVariance - (Rnd() * vertexEmissiveVariance * 2)
        
        tRed = tVar * (R_Primer + R_VertexVariance - (Rnd() * R_VertexVariance * 2))
        If tRed < 0 Then tRed = 0
        If tRed > 1 Then tRed = 1
        
        tGreen = tVar * (G_Primer + G_VertexVariance - (Rnd() * G_VertexVariance * 2))
        If tGreen < 0 Then tGreen = 0
        If tGreen > 1 Then tGreen = 1
        
        tBlue = tVar * (B_Primer + B_VertexVariance - (Rnd() * B_VertexVariance * 2))
        If tBlue < 0 Then tBlue = 0
        If tBlue > 1 Then tBlue = 1
        
        nSurfaceVertexArrayRed(loop1) = tRed
        nSurfaceVertexArrayGreen(loop1) = tGreen
        nSurfaceVertexArrayBlue(loop1) = tBlue
      End If
      
      Select Case shapeMode
        Case 0
          nSurfaceVertexArrayX(loop1) = (Rnd() * 2 - 1) * X_Radius
          nSurfaceVertexArrayY(loop1) = (Rnd() * 2 - 1) * Y_Radius
          nSurfaceVertexArrayZ(loop1) = (Rnd() * 2 - 1) * Z_Radius
        Case 1
          angularValue1 = Rnd() * Pi
          angularValue2 = Rnd() * Pi
          radiusValue = Rnd() * 2 - 1
          
          nSurfaceVertexArrayX(loop1) = Cos(angularValue1) * Cos(angularValue2) * radiusValue * X_Radius
          nSurfaceVertexArrayY(loop1) = Sin(angularValue1) * radiusValue * Y_Radius
          nSurfaceVertexArrayZ(loop1) = Cos(angularValue1) * Sin(angularValue2) * radiusValue * Z_Radius
        Case 2
          angularValue1 = Rnd() * Pi
          radiusValue = Rnd() * 2 - 1
          
          nSurfaceVertexArrayX(loop1) = Cos(angularValue1) * X_Radius * radiusValue
          nSurfaceVertexArrayY(loop1) = Y_Radius * (Rnd() * 2 - 1)
          nSurfaceVertexArrayZ(loop1) = Sin(angularValue1) * Z_Radius * radiusValue
        Case 3
          radiusValue = Rnd() * 4
          
          If radiusValue > 3 Then
            nSurfaceVertexArrayX(loop1) = X_Radius
            nSurfaceVertexArrayZ(loop1) = (Rnd() * 2 - 1) * Z_Radius
          ElseIf radiusValue > 2 Then
            nSurfaceVertexArrayX(loop1) = -X_Radius
            nSurfaceVertexArrayZ(loop1) = (Rnd() * 2 - 1) * Z_Radius
          ElseIf radiusValue > 1 Then
            nSurfaceVertexArrayX(loop1) = (Rnd() * 2 - 1) * X_Radius
            nSurfaceVertexArrayZ(loop1) = Z_Radius
          Else
            nSurfaceVertexArrayX(loop1) = (Rnd() * 2 - 1) * X_Radius
            nSurfaceVertexArrayZ(loop1) = -Z_Radius
          End If
          nSurfaceVertexArrayY(loop1) = Y_Radius * (Rnd() * 2 - 1)
        Case 4
          angularValue1 = Rnd() * Pi / 2 - Rnd() * Pi / 2
          angularValue2 = Rnd() * 2 * Pi
          
          nSurfaceVertexArrayX(loop1) = Cos(angularValue1) * Cos(angularValue2) * X_Radius
          nSurfaceVertexArrayY(loop1) = Sin(angularValue1) * Y_Radius
          nSurfaceVertexArrayZ(loop1) = Cos(angularValue1) * Sin(angularValue2) * Z_Radius
        Case 5
          angularValue1 = Rnd() * 2 * Pi
          
          nSurfaceVertexArrayX(loop1) = Cos(angularValue1) * X_Radius
          nSurfaceVertexArrayY(loop1) = (Rnd() * 2 - 1) * Y_Radius
          nSurfaceVertexArrayZ(loop1) = Sin(angularValue1) * Z_Radius
        Case 6
          angularValue1 = Rnd() * Pi
          angularValue2 = Rnd() * Pi
          angularValue3 = angularValue2 * 4
          
          Do While angularValue3 > Pi
            angularValue3 = angularValue3 - Pi
          Loop
          
          radiusValue = Sin(angularValue3 / 2)
          
          nSurfaceVertexArrayX(loop1) = Cos(angularValue1) * Cos(angularValue2 * radiusValue ^ 4) * radiusValue * X_Radius
          nSurfaceVertexArrayY(loop1) = Sin(angularValue1) * (Rnd() * 2 - 1) * Y_Radius
          nSurfaceVertexArrayZ(loop1) = Cos(angularValue1) * Sin(angularValue2 * radiusValue ^ 4) * radiusValue * Z_Radius
        Case Else
          angularValue1 = Rnd() * Pi * 2
          angularValue2 = Rnd() * Pi * 2
          
          nSurfaceVertexArrayX(loop1) = Cos(angularValue2) * (X_Radius + Sin(angularValue1) * Y_Radius)
          nSurfaceVertexArrayY(loop1) = Cos(angularValue1) * Y_Radius
          nSurfaceVertexArrayZ(loop1) = Sin(angularValue2) * (Z_Radius + Sin(angularValue1) * Y_Radius)
      End Select
    Next loop1
    
    .NumberOfSurfaceVertices = numberOfParticles
    
    .SurfaceVertexArrayX = nSurfaceVertexArrayX
    .SurfaceVertexArrayY = nSurfaceVertexArrayY
    .SurfaceVertexArrayZ = nSurfaceVertexArrayZ
    
    .SurfaceVertexArrayRed = nSurfaceVertexArrayRed
    .SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
    .SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
      
    .SurfaceVertexArrayNX = nSurfaceVertexArrayNX
    .SurfaceVertexArrayNY = nSurfaceVertexArrayNY
    .SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
    
    .SurfaceTextureArrayU = nSurfaceTextureArrayU
    .SurfaceTextureArrayV = nSurfaceTextureArrayV
    
    If vertexMode = False Then
      .SurfaceBlue = B_Primer
      .SurfaceGreen = G_Primer
      .SurfaceRed = R_Primer
      
      .SurfaceEmissiveBlue = B_Primer * primerEmissiveLevel
      .SurfaceEmissiveGreen = G_Primer * primerEmissiveLevel
      .SurfaceEmissiveRed = R_Primer * primerEmissiveLevel
    End If
    
    .SurfaceType = 3
    
    .Normalize_Surface
  End With
  
  Add_Surface sObject
End Sub

' Saves 3D object definition to file
Public Sub Save_ToFile(filePathName As String)
  Dim fileHandle As Long
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Output As #fileHandle
    Write #fileHandle, "Logicon Enterprises cls3D_Object Description Ver2.6"
    
    Save_ToOpenFile fileHandle
  Close fileHandle
  
  Exit Sub
  
badFile:
  MsgBox "Error writing 3D object to -> " & filePathName, vbCritical, "ERROR"
End Sub

' Saves 3D render chain definition to file
Public Sub Save_ChainToFile(filePathName As String)
  Dim fileHandle As Long, objectList As cls3D_Object, tCount As Long
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Output As #fileHandle
    Write #fileHandle, "Logicon Enterprises cls3D_Chain Description Ver2.6"
    
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
  MsgBox "Error writing 3D render chain to -> " & filePathName, vbCritical, "ERROR"
End Sub

' This should not be called directly... instead call Save_ToFile()
Public Sub Save_ToOpenFile(ByVal fileHandle As Long)
  Dim sList As cls3D_ObjectSurface, AList As cls3D_Object, tCount As Long
  
  Write #fileHandle, objectName
  Write #fileHandle, Visible
  Write #fileHandle, TypeID
  
  Write #fileHandle, BasePosX
  Write #fileHandle, BasePosY
  Write #fileHandle, BasePosZ
  Write #fileHandle, primaryAxis
  Write #fileHandle, PrimaryAxisRotation
  Write #fileHandle, PrimaryAxisTilt
  Write #fileHandle, PrimaryAxisPrecession
  
  Write #fileHandle, Collision_Mask
  Write #fileHandle, Collision_WidthX
  Write #fileHandle, Collision_WidthY
  Write #fileHandle, Collision_WidthZ
  
  Write #fileHandle, Collision_XLower
  Write #fileHandle, Collision_XUpper
  Write #fileHandle, Collision_YLower
  Write #fileHandle, Collision_YUpper
  Write #fileHandle, Collision_ZLower
  Write #fileHandle, Collision_ZUpper
  
  tCount = Get_SurfaceCount() 'save attached surfaces
  Write #fileHandle, tCount
  
  Set sList = SurfaceList
  
  Do While Not (sList Is Nothing)
    sList.Save_ToOpenFile fileHandle
    
    Set sList = sList.NextSurface
  Loop

  tCount = Get_Attached3DObjectCount() ' save attached objects
  Write #fileHandle, tCount
  
  Set AList = attachedList
  
  Do While Not (AList Is Nothing)
    AList.Save_ToOpenFile fileHandle
    
    Set AList = AList.chainNext
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

  Write #fileHandle, direction
  Write #fileHandle, Range
  Write #fileHandle, Speed
  
  Write #fileHandle, VelocityX
  Write #fileHandle, VelocityY
  Write #fileHandle, VelocityZ
  Write #fileHandle, AccelerationX
  Write #fileHandle, AccelerationY
  Write #fileHandle, AccelerationZ

  Write #fileHandle, TargetMode
  Write #fileHandle, TargetX
  Write #fileHandle, TargetY
  Write #fileHandle, TargetZ

  Write #fileHandle, User1
  Write #fileHandle, User2
  Write #fileHandle, User3
  
  '**********************************************************************************************
End Sub

' Loads 3D object definition from file
Public Sub Load_FromFile(filePathName As String)
  Dim fileHandle As Long, versionString As String
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Input As #fileHandle
    Input #fileHandle, versionString
    
    Destroy_Object
    
    Load_FromOpenFile fileHandle
  Close fileHandle
  
  Exit Sub
  
badFile:
  MsgBox "Error loading 3D object -> " & filePathName, vbCritical, "ERROR"
End Sub

' Loads 3D render chain definition to file
Public Sub Load_ChainFromFile(filePathName As String)
  Dim fileHandle As Long, newObject As cls3D_Object, lastObject As cls3D_Object
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
      Set newObject = New cls3D_Object
      
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
  MsgBox "Error loading 3D render chain from -> " & filePathName, vbCritical, "ERROR"
End Sub

' This should not be called directly... instead call Load_FromFile()
Public Sub Load_FromOpenFile(ByVal fileHandle As Long)
  Dim sObject As cls3D_ObjectSurface, aObject As cls3D_Object, tCount As Long
  
  Input #fileHandle, objectName
  Input #fileHandle, Visible
  Input #fileHandle, TypeID

  Input #fileHandle, BasePosX
  Input #fileHandle, BasePosY
  Input #fileHandle, BasePosZ
  Input #fileHandle, primaryAxis
  Input #fileHandle, PrimaryAxisRotation
  Input #fileHandle, PrimaryAxisTilt
  Input #fileHandle, PrimaryAxisPrecession

  Input #fileHandle, Collision_Mask
  Input #fileHandle, Collision_WidthX
  Input #fileHandle, Collision_WidthY
  Input #fileHandle, Collision_WidthZ
  
  Input #fileHandle, Collision_XLower
  Input #fileHandle, Collision_XUpper
  Input #fileHandle, Collision_YLower
  Input #fileHandle, Collision_YUpper
  Input #fileHandle, Collision_ZLower
  Input #fileHandle, Collision_ZUpper
  
  DestroyAll_Surfaces
  Input #fileHandle, tCount
  
  Do While tCount > 0
    Set sObject = New cls3D_ObjectSurface
    
    sObject.Load_FromOpenFile fileHandle
    Add_Surface sObject
    
    tCount = tCount - 1
  Loop
  
  Input #fileHandle, tCount
  
  Do While tCount > 0
    Set aObject = New cls3D_Object
    
    aObject.Load_FromOpenFile fileHandle
    Add_Attached3DObject aObject
    
    tCount = tCount - 1
  Loop
  
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

  Input #fileHandle, direction
  Input #fileHandle, Range
  Input #fileHandle, Speed
  
  Input #fileHandle, VelocityX
  Input #fileHandle, VelocityY
  Input #fileHandle, VelocityZ
  Input #fileHandle, AccelerationX
  Input #fileHandle, AccelerationY
  Input #fileHandle, AccelerationZ

  Input #fileHandle, TargetMode
  Input #fileHandle, TargetX
  Input #fileHandle, TargetY
  Input #fileHandle, TargetZ
  
  Input #fileHandle, User1
  Input #fileHandle, User2
  Input #fileHandle, User3
  
  '**********************************************************************************************
End Sub

'**********************************************************************************************************

'******************************************** render chains *********************************************

'**********************************************************************************************************

'adds the 3D object to the 3D render chain
Public Sub AddTo_Chain(ByRef chainStart As cls3D_Object, Optional objectAlwaysOpaque As Boolean = True)
  If chainStart Is Nothing Then
    Set chainPrevious = Nothing
    Set chainNext = Nothing
    
    Set chainStart = Me
  ElseIf objectAlwaysOpaque Then
    Set chainPrevious = Nothing
    Set chainNext = chainStart
    Set chainStart.chainPrevious = Me
    
    Set chainStart = Me
  Else
    Dim renderList As cls3D_Object
    
    Set renderList = chainStart
    
    Do While Not (renderList.chainNext Is Nothing)
      Set renderList = renderList.chainNext
    Loop
    
    Set renderList.chainNext = Me
    Set chainPrevious = renderList
    Set chainNext = Nothing
  End If
End Sub

'removes the object from the render chain and returns the first objects in the render chain
'removing an object from the render chain that was not in the render chain will cause corruption in
'the render chain
Public Sub RemoveFrom_Chain(ByRef chainStart As cls3D_Object)
  If Not (chainNext Is Nothing) Then Set chainNext.chainPrevious = chainPrevious
    
  If chainPrevious Is Nothing Then
    Set chainStart = chainNext
  Else
    Set chainPrevious.chainNext = chainNext
  End If
    
  Set chainNext = Nothing
  Set chainPrevious = Nothing
End Sub

Public Function Get_ChainObjectID(ByVal objectID As Long) As cls3D_Object
  Dim chainList As cls3D_Object
  
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

Public Function Get_ChainObjectName(ByVal objectName As String, Optional ByVal objectID As Long = 0) As cls3D_Object
  Dim chainList As cls3D_Object
  
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

Public Sub Calculate_and_Update_Chain()
  Dim oList As cls3D_Object
  
  Set oList = Me
  
  Do While Not (oList Is Nothing)
    With oList
      .Update_ObjectPosition
      .Calculate_Vertices True
      
      Set oList = .chainNext
    End With
  Loop
End Sub

'renders the passed chain of 3DObjects
Public Sub ReDraw_Chain(Optional ByVal enableTransparency As Boolean = True, Optional ByVal enableTranslucency As Boolean = True)
  Dim objectList As cls3D_Object
  
  On Error Resume Next
  
  If enableTranslucency = False Then dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, False
  If enableTransparency = False Then dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, False
  
  Set objectList = Me
  
  Do While Not (objectList Is Nothing)
    With objectList
      If .Visible Then .Render_Object True, True
      
      Set objectList = .chainNext
    End With
  Loop
  
  If enableTranslucency = False Then dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, True
  If enableTransparency = False Then dx_Direct3DDevice.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, True
End Sub

Public Sub Calculate_Chain_CollisionBox()
  Dim chainList As cls3D_Object
  
  Set chainList = Me
  
  Do While Not (chainList Is Nothing)
    With chainList
      .Calculate_CollsionBox
      
      Set chainList = .chainNext
    End With
  Loop
End Sub

'Releases references that the objects in the chain may hold so they can be released by the system
Public Sub Destroy_Chain()
  Dim chainList As cls3D_Object
  
  Set chainList = Me
  
  Do While Not (chainList Is Nothing)
    With chainList
      Set chainList = .chainNext
      
      .Destroy_Object
    End With
  Loop
End Sub

