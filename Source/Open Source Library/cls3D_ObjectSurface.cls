VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls3D_ObjectSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************************
'
' cls3D_ObjectSurface class definition
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

'***************************************************************************************************************
'
' The following section contains the core member variables for this class - do not remove.
'
'***************************************************************************************************************

Private dx_DirectX As DirectX7
Private dx_Direct3D As Direct3D7
Private dx_Direct3DDevice As Direct3DDevice7

Public Visible As Boolean                                         ' surface's visibility flag (default TRUE)

Public TypeID As Long                                            ' user defined value

Public ParentObject As cls3D_Object                        ' parent object for surface
Public PreviousSurface As cls3D_ObjectSurface       ' previous surface in list
Public NextSurface As cls3D_ObjectSurface              ' next surface in list

Public TextureSurfaceIndex As Long                  ' set to 0 for no texture
                                                     
' all vertex arrays MUST be 0 based
Public SurfaceType As Long                     ' surface type = 0 (triList), 1 (triStrip), 2 (triFans)
                                                                   '                         3 (pointList), 4 (lineList), 5 (lineStrip)
Public NumberOfSurfaceVertices As Long        ' number of surfaces vertices
Public SurfaceVertexArrayX As Variant              ' array of vertex X co-ords
Public SurfaceVertexArrayY As Variant              ' array of vertex Y co-ords
Public SurfaceVertexArrayZ As Variant              ' array of vertex Z co-ords
Public SurfaceVertexArrayNX As Variant           ' array of vertex X normals
Public SurfaceVertexArrayNY As Variant           ' array of vertex Y normals
Public SurfaceVertexArrayNZ As Variant           ' array of vertex Z normals
Public SurfaceTextureArrayU As Variant            ' array of vertex texture u co-ords
Public SurfaceTextureArrayV As Variant            ' array of vertex texture v co-ords
Public SurfaceVertexArrayRed As Variant         ' array of vertex red colour
Public SurfaceVertexArrayGreen As Variant      ' array of vertex green colour
Public SurfaceVertexArrayBlue As Variant         ' array of vertex blue colour

Private Type D3DLitVertex
  X As Single
  Y As Single
  z As Single
  
  Colour As Long
  
  tU As Single
  tV As Single
End Type

Private VertexArray() As D3DLitVertex             ' array of calculated surface vertices
Private Const VERTEXTYPE As Long = D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1

Public SurfaceOpacity As Single                           ' the opacity of the surface (opposite of transparency)
Public SurfaceRed As Single                                  ' surface ambient/diffuse red property
Public SurfaceGreen As Single                               ' surface ambient/diffuse green property
Public SurfaceBlue As Single                                  ' surface ambient/diffuse blue property
Public SurfaceEmissiveRed As Single                       ' surface emissive red property
Public SurfaceEmissiveGreen As Single                    ' surface emissive green property
Public SurfaceEmissiveBlue As Single                       ' surface emissive blue property

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

Public MainAction As Long                                       ' user defined value
Public TicksPerMainAction As Long                         ' user defined value
Public MainActionTickCounter As Long                   ' user defined value
Public SubAction As Long                                        ' user defined value
Public TicksPerSubAction As Long                         ' user defined value
Public SubActionTickCounter As Long                    ' user defined value

Public User1 As Long                                            ' user defined value

'***************************************************************************************************************



' must be called once before surface can be rendered
' also re-call if the world light settings or the surface vertices are altered in any way
Public Sub Calculate_Vertices()
  Dim loop1 As Long, arrayTop As Long
  Dim tRedL As Long, tGreenL As Long, tBlueL As Long, tOpacity As Long
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim VertexArray(0 To arrayTop)
  
  tOpacity = SurfaceOpacity * 255
  If tOpacity > 255 Then tOpacity = 255
  
  If tOpacity > 127 Then
    tOpacity = (&H1000000 * (255 - tOpacity)) Xor &HFF000000
  Else
    tOpacity = &H1000000 * tOpacity
  End If
  
  For loop1 = 0 To arrayTop
    With VertexArray(loop1)
      .X = SurfaceVertexArrayX(loop1)
      .Y = SurfaceVertexArrayY(loop1)
      .z = SurfaceVertexArrayZ(loop1)
      
      tRedL = 255 * ((SurfaceVertexArrayRed(loop1) + SurfaceEmissiveRed + SurfaceRed * DXDraw.D3DWorldAmbient_Red) * DXDraw.D3DWorldIntensity + DXDraw.D3DWorldBrightnessAdjust)
      If tRedL > 255 Then tRedL = 255
      
      tGreenL = 255 * ((SurfaceVertexArrayGreen(loop1) + SurfaceEmissiveGreen + SurfaceGreen * DXDraw.D3DWorldAmbient_Green) * DXDraw.D3DWorldIntensity + DXDraw.D3DWorldBrightnessAdjust)
      If tGreenL > 255 Then tGreenL = 255
      
      tBlueL = 255 * ((SurfaceVertexArrayBlue(loop1) + SurfaceEmissiveBlue + SurfaceBlue * DXDraw.D3DWorldAmbient_Blue) * DXDraw.D3DWorldIntensity + DXDraw.D3DWorldBrightnessAdjust)
      If tBlueL > 255 Then tBlueL = 255
      
      .Colour = tOpacity Or (tRedL * &H10000) Or (tGreenL * &H100) Or tBlueL
      
      .tU = SurfaceTextureArrayU(loop1)
      .tV = SurfaceTextureArrayV(loop1)
    End With
  Next loop1
End Sub

' calculates only the XYZ positional vertex information - used for rapid updates in position
Public Sub Calculate_VerticesXYZ()
  Dim loop1 As Long, arrayTop As Long
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  For loop1 = 0 To arrayTop
    With VertexArray(loop1)
      .X = SurfaceVertexArrayX(loop1)
      .Y = SurfaceVertexArrayY(loop1)
      .z = SurfaceVertexArrayZ(loop1)
    End With
  Next loop1
End Sub

' calculates only the RGB colour vertex information - used for rapid updates in colour
Public Sub Calculate_VerticesColour()
  Dim loop1 As Long, arrayTop As Long
  Dim tRedL As Long, tGreenL As Long, tBlueL As Long, tOpacity As Long
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  tOpacity = SurfaceOpacity * 255
  If tOpacity > 255 Then tOpacity = 255
  
  If tOpacity > 127 Then
    tOpacity = (&H1000000 * (255 - tOpacity)) Xor &HFF000000
  Else
    tOpacity = &H1000000 * tOpacity
  End If
  
  For loop1 = 0 To arrayTop
    With VertexArray(loop1)
      tRedL = 255 * ((SurfaceVertexArrayRed(loop1) + SurfaceEmissiveRed + SurfaceRed * DXDraw.D3DWorldAmbient_Red) * DXDraw.D3DWorldIntensity + DXDraw.D3DWorldBrightnessAdjust)
      If tRedL > 255 Then tRedL = 255
      
      tGreenL = 255 * ((SurfaceVertexArrayGreen(loop1) + SurfaceEmissiveGreen + SurfaceGreen * DXDraw.D3DWorldAmbient_Green) * DXDraw.D3DWorldIntensity + DXDraw.D3DWorldBrightnessAdjust)
      If tGreenL > 255 Then tGreenL = 255
      
      tBlueL = 255 * ((SurfaceVertexArrayBlue(loop1) + SurfaceEmissiveBlue + SurfaceBlue * DXDraw.D3DWorldAmbient_Blue) * DXDraw.D3DWorldIntensity + DXDraw.D3DWorldBrightnessAdjust)
      If tBlueL > 255 Then tBlueL = 255
      
      .Colour = tOpacity Or (tRedL * &H10000) Or (tGreenL * &H100) Or tBlueL
    End With
  Next loop1
End Sub

' calculates only the UV texture vertex information - used for rapid updates in textures
Public Sub Calculate_VerticesUV()
  Dim loop1 As Long, arrayTop As Long
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  For loop1 = 0 To arrayTop
    With VertexArray(loop1)
      .tU = SurfaceTextureArrayU(loop1)
      .tV = SurfaceTextureArrayV(loop1)
    End With
  Next loop1
End Sub

' do not call directly - should be called through call to it's ParentObject's Render_Object()
Public Sub Render_Surface()
  On Error Resume Next
  
  If Visible Then
    If (TextureSurfaceIndex > 0) And DXDraw.D3DTextureEnable Then
      dx_Direct3DDevice.SetTexture 0, DXDraw.GetDirectDrawSurface(TextureSurfaceIndex)
    Else
      dx_Direct3DDevice.SetTexture 0, Nothing
    End If
    
    Select Case SurfaceType
      Case 0
        dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, VERTEXTYPE, VertexArray(0), NumberOfSurfaceVertices, D3DDP_DEFAULT
      Case 1
        dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, VERTEXTYPE, VertexArray(0), NumberOfSurfaceVertices, D3DDP_DEFAULT
      Case 2
        dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLEFAN, VERTEXTYPE, VertexArray(0), NumberOfSurfaceVertices, D3DDP_DEFAULT
      Case 3
        dx_Direct3DDevice.DrawPrimitive D3DPT_POINTLIST, VERTEXTYPE, VertexArray(0), NumberOfSurfaceVertices, D3DDP_DEFAULT
      Case 4
        dx_Direct3DDevice.DrawPrimitive D3DPT_LINELIST, VERTEXTYPE, VertexArray(0), NumberOfSurfaceVertices, D3DDP_DEFAULT
      Case Else
        dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, VERTEXTYPE, VertexArray(0), NumberOfSurfaceVertices, D3DDP_DEFAULT
    End Select
  End If
End Sub

Public Sub Mirror_yzPlane()
  Dim loop1 As Long, arrayTop As Long
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayNX() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceVertexArrayX(0 To arrayTop)
  ReDim nSurfaceVertexArrayNX(0 To arrayTop)
  
  For loop1 = 0 To arrayTop
    nSurfaceVertexArrayX(loop1) = -SurfaceVertexArrayX(loop1)
    nSurfaceVertexArrayNX(loop1) = -SurfaceVertexArrayNX(loop1)
  Next loop1
  
  SurfaceVertexArrayX = nSurfaceVertexArrayX
  SurfaceVertexArrayNX = nSurfaceVertexArrayNX
  
  Flip_VisibleSurface
End Sub

Public Sub Mirror_xzPlane()
  Dim loop1 As Long, arrayTop As Long
  Dim nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayNY() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceVertexArrayY(0 To arrayTop)
  ReDim nSurfaceVertexArrayNY(0 To arrayTop)
  
  For loop1 = 0 To arrayTop
    nSurfaceVertexArrayY(loop1) = -SurfaceVertexArrayY(loop1)
    nSurfaceVertexArrayNY(loop1) = -SurfaceVertexArrayNY(loop1)
  Next loop1
  
  SurfaceVertexArrayY = nSurfaceVertexArrayY
  SurfaceVertexArrayNY = nSurfaceVertexArrayNY
  
  Flip_VisibleSurface
End Sub

Public Sub Mirror_xyPlane()
  Dim loop1 As Long, arrayTop As Long
  Dim nSurfaceVertexArrayZ() As Single, nSurfaceVertexArrayNZ() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceVertexArrayZ(0 To arrayTop)
  ReDim nSurfaceVertexArrayNZ(0 To arrayTop)
  
  For loop1 = 0 To arrayTop
    nSurfaceVertexArrayZ(loop1) = -SurfaceVertexArrayZ(loop1)
    nSurfaceVertexArrayNZ(loop1) = -SurfaceVertexArrayNZ(loop1)
  Next loop1
  
  SurfaceVertexArrayZ = nSurfaceVertexArrayZ
  SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
  
  Flip_VisibleSurface
End Sub

Public Sub Flip_zAxis90()
  Dim loop1 As Long, arrayTop As Long
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceVertexArrayX(0 To arrayTop)
  ReDim nSurfaceVertexArrayY(0 To arrayTop)
  
  ReDim nSurfaceVertexArrayNX(0 To arrayTop)
  ReDim nSurfaceVertexArrayNY(0 To arrayTop)
  
  For loop1 = 0 To arrayTop
    nSurfaceVertexArrayX(loop1) = SurfaceVertexArrayY(loop1)
    nSurfaceVertexArrayY(loop1) = -SurfaceVertexArrayX(loop1)
    nSurfaceVertexArrayNX(loop1) = SurfaceVertexArrayNY(loop1)
    nSurfaceVertexArrayNY(loop1) = -SurfaceVertexArrayNX(loop1)
  Next loop1
  
  SurfaceVertexArrayX = nSurfaceVertexArrayX
  SurfaceVertexArrayY = nSurfaceVertexArrayY
  SurfaceVertexArrayNX = nSurfaceVertexArrayNX
  SurfaceVertexArrayNY = nSurfaceVertexArrayNY
End Sub

Public Sub Flip_yAxis90()
  Dim loop1 As Long, arrayTop As Long, holdValue As Single
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNZ() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceVertexArrayX(0 To arrayTop)
  ReDim nSurfaceVertexArrayZ(0 To arrayTop)
  
  ReDim nSurfaceVertexArrayNX(0 To arrayTop)
  ReDim nSurfaceVertexArrayNZ(0 To arrayTop)
  
  For loop1 = 0 To arrayTop
    nSurfaceVertexArrayX(loop1) = -SurfaceVertexArrayZ(loop1)
    nSurfaceVertexArrayZ(loop1) = SurfaceVertexArrayX(loop1)
    nSurfaceVertexArrayNX(loop1) = -SurfaceVertexArrayNZ(loop1)
    nSurfaceVertexArrayNZ(loop1) = SurfaceVertexArrayNX(loop1)
  Next loop1
  
  SurfaceVertexArrayX = nSurfaceVertexArrayX
  SurfaceVertexArrayZ = nSurfaceVertexArrayZ
  SurfaceVertexArrayNX = nSurfaceVertexArrayNX
  SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
End Sub

Public Sub Flip_xAxis90()
  Dim loop1 As Long, arrayTop As Long, holdValue As Single
  Dim nSurfaceVertexArrayZ() As Single, nSurfaceVertexArrayY() As Single
  Dim nSurfaceVertexArrayNZ() As Single, nSurfaceVertexArrayNY() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceVertexArrayY(0 To arrayTop)
  ReDim nSurfaceVertexArrayZ(0 To arrayTop)
  
  ReDim nSurfaceVertexArrayNY(0 To arrayTop)
  ReDim nSurfaceVertexArrayNZ(0 To arrayTop)
  
  For loop1 = 0 To arrayTop
    nSurfaceVertexArrayY(loop1) = SurfaceVertexArrayZ(loop1)
    nSurfaceVertexArrayZ(loop1) = -SurfaceVertexArrayY(loop1)
    nSurfaceVertexArrayNY(loop1) = SurfaceVertexArrayNZ(loop1)
    nSurfaceVertexArrayNZ(loop1) = -SurfaceVertexArrayNY(loop1)
  Next loop1
  
  SurfaceVertexArrayY = nSurfaceVertexArrayY
  SurfaceVertexArrayZ = nSurfaceVertexArrayZ
  SurfaceVertexArrayNY = nSurfaceVertexArrayNY
  SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
End Sub

Public Sub Flip_VisibleSurface()
  Dim loop1 As Long, arrayTop As Long
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
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
  
  If SurfaceType = 0 Then
    For loop1 = 0 To arrayTop
      nSurfaceVertexArrayX(loop1) = SurfaceVertexArrayX(arrayTop - loop1)
      nSurfaceVertexArrayY(loop1) = SurfaceVertexArrayY(arrayTop - loop1)
      nSurfaceVertexArrayZ(loop1) = SurfaceVertexArrayZ(arrayTop - loop1)
      
      nSurfaceTextureArrayU(loop1) = SurfaceTextureArrayU(arrayTop - loop1)
      nSurfaceTextureArrayV(loop1) = SurfaceTextureArrayV(arrayTop - loop1)
      
      nSurfaceVertexArrayRed(loop1) = SurfaceVertexArrayRed(arrayTop - loop1)
      nSurfaceVertexArrayGreen(loop1) = SurfaceVertexArrayGreen(arrayTop - loop1)
      nSurfaceVertexArrayBlue(loop1) = SurfaceVertexArrayBlue(arrayTop - loop1)
      
      nSurfaceVertexArrayNX(loop1) = -SurfaceVertexArrayNX(arrayTop - loop1)
      nSurfaceVertexArrayNY(loop1) = -SurfaceVertexArrayNY(arrayTop - loop1)
      nSurfaceVertexArrayNZ(loop1) = -SurfaceVertexArrayNZ(arrayTop - loop1)
    Next loop1
  ElseIf SurfaceType = 1 Then
    For loop1 = 0 To arrayTop Step 2
      nSurfaceVertexArrayX(loop1) = SurfaceVertexArrayX(loop1 + 1)
      nSurfaceVertexArrayX(loop1 + 1) = SurfaceVertexArrayX(loop1)
      nSurfaceVertexArrayY(loop1) = SurfaceVertexArrayY(loop1 + 1)
      nSurfaceVertexArrayY(loop1 + 1) = SurfaceVertexArrayY(loop1)
      nSurfaceVertexArrayZ(loop1) = SurfaceVertexArrayZ(loop1 + 1)
      nSurfaceVertexArrayZ(loop1 + 1) = SurfaceVertexArrayZ(loop1)
      
      nSurfaceTextureArrayU(loop1) = SurfaceTextureArrayU(loop1 + 1)
      nSurfaceTextureArrayU(loop1 + 1) = SurfaceTextureArrayU(loop1)
      nSurfaceTextureArrayV(loop1) = SurfaceTextureArrayV(loop1 + 1)
      nSurfaceTextureArrayV(loop1 + 1) = SurfaceTextureArrayV(loop1)
      
      nSurfaceVertexArrayRed(loop1) = SurfaceVertexArrayRed(loop1 + 1)
      nSurfaceVertexArrayRed(loop1 + 1) = SurfaceVertexArrayRed(loop1)
      nSurfaceVertexArrayGreen(loop1) = SurfaceVertexArrayGreen(loop1 + 1)
      nSurfaceVertexArrayGreen(loop1 + 1) = SurfaceVertexArrayGreen(loop1)
      nSurfaceVertexArrayBlue(loop1) = SurfaceVertexArrayBlue(loop1 + 1)
      nSurfaceVertexArrayBlue(loop1 + 1) = SurfaceVertexArrayBlue(loop1)
      
      nSurfaceVertexArrayNX(loop1) = -SurfaceVertexArrayNX(loop1 + 1)
      nSurfaceVertexArrayNX(loop1 + 1) = -SurfaceVertexArrayNX(loop1)
      nSurfaceVertexArrayNY(loop1) = -SurfaceVertexArrayNY(loop1 + 1)
      nSurfaceVertexArrayNY(loop1 + 1) = -SurfaceVertexArrayNY(loop1)
      nSurfaceVertexArrayNZ(loop1) = -SurfaceVertexArrayNZ(loop1 + 1)
      nSurfaceVertexArrayNZ(loop1 + 1) = -SurfaceVertexArrayNZ(loop1)
    Next loop1
  Else
    For loop1 = 1 To arrayTop
      nSurfaceVertexArrayX(loop1) = SurfaceVertexArrayX(arrayTop - loop1 + 1)
      nSurfaceVertexArrayY(loop1) = SurfaceVertexArrayY(arrayTop - loop1 + 1)
      nSurfaceVertexArrayZ(loop1) = SurfaceVertexArrayZ(arrayTop - loop1 + 1)
      
      nSurfaceTextureArrayU(loop1) = SurfaceTextureArrayU(arrayTop - loop1 + 1)
      nSurfaceTextureArrayV(loop1) = SurfaceTextureArrayV(arrayTop - loop1 + 1)
      
      nSurfaceVertexArrayRed(loop1) = SurfaceVertexArrayRed(arrayTop - loop1 + 1)
      nSurfaceVertexArrayGreen(loop1) = SurfaceVertexArrayGreen(arrayTop - loop1 + 1)
      nSurfaceVertexArrayBlue(loop1) = SurfaceVertexArrayBlue(arrayTop - loop1 + 1)
        
      nSurfaceVertexArrayNX(loop1) = -SurfaceVertexArrayNX(arrayTop - loop1 + 1)
      nSurfaceVertexArrayNY(loop1) = -SurfaceVertexArrayNY(arrayTop - loop1 + 1)
      nSurfaceVertexArrayNZ(loop1) = -SurfaceVertexArrayNZ(arrayTop - loop1 + 1)
    Next loop1
  End If
  
  SurfaceVertexArrayX = nSurfaceVertexArrayX
  SurfaceVertexArrayY = nSurfaceVertexArrayY
  SurfaceVertexArrayZ = nSurfaceVertexArrayZ
  
  SurfaceVertexArrayRed = nSurfaceVertexArrayRed
  SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
  SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
    
  SurfaceVertexArrayNX = nSurfaceVertexArrayNX
  SurfaceVertexArrayNY = nSurfaceVertexArrayNY
  SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

' translates surface vertices - this adjustment is relative
Public Sub Translate_Surface(Optional ByVal xShift As Long = 0, Optional ByVal yShift As Long = 0, Optional ByVal zShift As Long = 0, _
      Optional ByVal pAxis As Long = 0, Optional ByVal pAxisRotation As Single = 0, _
      Optional ByVal pAxisTilt As Single = 0, Optional ByVal pAxisPrecession As Single = 0)
      
  Dim loop1 As Long, arrayTop As Long, tMatrix As D3DMATRIX, matSpinA As D3DMATRIX, matSpinB As D3DMATRIX
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceVertexArrayX(0 To arrayTop)
  ReDim nSurfaceVertexArrayY(0 To arrayTop)
  ReDim nSurfaceVertexArrayZ(0 To arrayTop)
  ReDim nSurfaceVertexArrayNX(0 To arrayTop)
  ReDim nSurfaceVertexArrayNY(0 To arrayTop)
  ReDim nSurfaceVertexArrayNZ(0 To arrayTop)
  
  With dx_DirectX
    Select Case pAxis
      Case 0
        .RotateXMatrix matSpinA, pAxisRotation
        .RotateYMatrix matSpinB, pAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateXMatrix matSpinB, pAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
      Case 1
        .RotateYMatrix matSpinA, pAxisRotation
        .RotateZMatrix matSpinB, pAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateYMatrix matSpinB, pAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
      Case Else
        .RotateZMatrix matSpinA, pAxisRotation
        .RotateXMatrix matSpinB, pAxisTilt
        .MatrixMultiply matSpinA, matSpinA, matSpinB
        .RotateZMatrix matSpinB, pAxisPrecession
        .MatrixMultiply matSpinA, matSpinA, matSpinB
    End Select
    
    For loop1 = 0 To arrayTop
      .IdentityMatrix tMatrix
      
      With tMatrix
        .rc41 = SurfaceVertexArrayX(loop1)
        .rc42 = SurfaceVertexArrayY(loop1)
        .rc43 = SurfaceVertexArrayZ(loop1)
      End With
      
      .MatrixMultiply tMatrix, tMatrix, matSpinA
      
      With tMatrix
        nSurfaceVertexArrayX(loop1) = .rc41 + xShift
        nSurfaceVertexArrayY(loop1) = .rc42 + yShift
        nSurfaceVertexArrayZ(loop1) = .rc43 + zShift
      End With
      
      .IdentityMatrix tMatrix
      
      With tMatrix
        .rc41 = SurfaceVertexArrayNX(loop1)
        .rc42 = SurfaceVertexArrayNY(loop1)
        .rc43 = SurfaceVertexArrayNZ(loop1)
      End With
      
      .MatrixMultiply tMatrix, tMatrix, matSpinA
      
      With tMatrix
        nSurfaceVertexArrayNX(loop1) = .rc41
        nSurfaceVertexArrayNY(loop1) = .rc42
        nSurfaceVertexArrayNZ(loop1) = .rc43
      End With
    Next loop1
  End With
  
  SurfaceVertexArrayX = nSurfaceVertexArrayX
  SurfaceVertexArrayY = nSurfaceVertexArrayY
  SurfaceVertexArrayZ = nSurfaceVertexArrayZ
  SurfaceVertexArrayNX = nSurfaceVertexArrayNX
  SurfaceVertexArrayNY = nSurfaceVertexArrayNY
  SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
End Sub

' shifts texture vertices - this adjustment is relative
Public Sub Shift_Texture(ByVal uShift As Single, ByVal vShift As Single)
  Dim loop1 As Long, arrayTop As Long
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceTextureArrayU(0 To arrayTop)
  ReDim nSurfaceTextureArrayV(0 To arrayTop)
  
  For loop1 = 0 To arrayTop
    nSurfaceTextureArrayU(loop1) = SurfaceTextureArrayU(loop1) + uShift
    nSurfaceTextureArrayV(loop1) = SurfaceTextureArrayV(loop1) + vShift
  Next loop1
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

' scales texture vertices - this adjustment is relative
Public Sub Scale_Texture(ByVal uScale As Single, ByVal vScale As Single)
  Dim loop1 As Long, arrayTop As Long
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceTextureArrayU(0 To arrayTop)
  ReDim nSurfaceTextureArrayV(0 To arrayTop)
  
  For loop1 = 0 To arrayTop
    nSurfaceTextureArrayU(loop1) = SurfaceTextureArrayU(loop1) * uScale
    nSurfaceTextureArrayV(loop1) = SurfaceTextureArrayV(loop1) * vScale
  Next loop1
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

' recalculates texture co-ords, the base scale is the distance before texture is repeated
Public Sub Set_TexturePlaneXY(Optional ByVal uBaseScale As Single = 1, Optional ByVal vBaseScale As Single = 1, _
        Optional ByVal uBaseShift As Single = 0, Optional ByVal vBaseShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False, Optional ByVal AutoFit As Boolean = True)
  
  Dim loop1 As Long, arrayTop As Long
  Dim minX As Single, maxX As Single, minY As Single, maxY As Single, uTemp As Single, vTemp As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceTextureArrayU(0 To arrayTop)
  ReDim nSurfaceTextureArrayV(0 To arrayTop)
  
  If AutoFit Then
    minX = 9999999
    maxX = -9999999
    minY = 9999999
    maxY = -9999999
    
    For loop1 = 0 To arrayTop
      If SurfaceVertexArrayX(loop1) < minX Then minX = SurfaceVertexArrayX(loop1)
      If SurfaceVertexArrayX(loop1) > maxX Then maxX = SurfaceVertexArrayX(loop1)
      If SurfaceVertexArrayY(loop1) < minY Then minY = SurfaceVertexArrayY(loop1)
      If SurfaceVertexArrayY(loop1) > maxY Then maxY = SurfaceVertexArrayY(loop1)
    Next loop1
    
    If FlipUV Then
      uTemp = maxY - minY
      vTemp = maxX - minX
      
      uBaseShift = uBaseShift - minY / uTemp
      vBaseShift = vBaseShift - minX / vTemp
    Else
      uTemp = maxX - minX
      vTemp = maxY - minY
      
      uBaseShift = uBaseShift - minX / uTemp
      vBaseShift = vBaseShift - minY / vTemp
    End If
    
    uBaseScale = uBaseScale * uTemp
    vBaseScale = vBaseScale * vTemp
  End If
  
  If FlipUV Then
    For loop1 = 0 To arrayTop
      nSurfaceTextureArrayU(loop1) = uBaseShift + SurfaceVertexArrayY(loop1) / uBaseScale
      nSurfaceTextureArrayV(loop1) = vBaseShift + SurfaceVertexArrayX(loop1) / vBaseScale
    Next loop1
  Else
    For loop1 = 0 To arrayTop
      nSurfaceTextureArrayU(loop1) = uBaseShift + SurfaceVertexArrayX(loop1) / uBaseScale
      nSurfaceTextureArrayV(loop1) = vBaseShift + SurfaceVertexArrayY(loop1) / vBaseScale
    Next loop1
  End If
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

' recalculates texture co-ords, the base scale is the distance before texture is repeated
Public Sub Set_TexturePlaneXZ(Optional ByVal uBaseScale As Single = 1, Optional ByVal vBaseScale As Single = 1, _
        Optional ByVal uBaseShift As Single = 0, Optional ByVal vBaseShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False, Optional ByVal AutoFit As Boolean = True)
  
  Dim loop1 As Long, arrayTop As Long
  Dim minX As Single, maxX As Single, minZ As Single, maxZ As Single, uTemp As Single, vTemp As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceTextureArrayU(0 To arrayTop)
  ReDim nSurfaceTextureArrayV(0 To arrayTop)
  
  If AutoFit Then
    minX = 9999999
    maxX = -9999999
    minZ = 9999999
    maxZ = -9999999
    
    For loop1 = 0 To arrayTop
      If SurfaceVertexArrayX(loop1) < minX Then minX = SurfaceVertexArrayX(loop1)
      If SurfaceVertexArrayX(loop1) > maxX Then maxX = SurfaceVertexArrayX(loop1)
      If SurfaceVertexArrayZ(loop1) < minZ Then minZ = SurfaceVertexArrayZ(loop1)
      If SurfaceVertexArrayZ(loop1) > maxZ Then maxZ = SurfaceVertexArrayZ(loop1)
    Next loop1
    
    If FlipUV Then
      uTemp = maxZ - minZ
      vTemp = maxX - minX
      
      uBaseShift = uBaseShift - minZ / uTemp
      vBaseShift = vBaseShift - minX / vTemp
    Else
      uTemp = maxX - minX
      vTemp = maxZ - minZ
      
      uBaseShift = uBaseShift - minX / uTemp
      vBaseShift = vBaseShift - minZ / vTemp
    End If
    
    uBaseScale = uBaseScale * uTemp
    vBaseScale = vBaseScale * vTemp
  End If
  
  If FlipUV Then
    For loop1 = 0 To arrayTop
      nSurfaceTextureArrayU(loop1) = uBaseShift + SurfaceVertexArrayZ(loop1) / uBaseScale
      nSurfaceTextureArrayV(loop1) = vBaseShift + SurfaceVertexArrayX(loop1) / vBaseScale
    Next loop1
  Else
    For loop1 = 0 To arrayTop
      nSurfaceTextureArrayU(loop1) = uBaseShift + SurfaceVertexArrayX(loop1) / uBaseScale
      nSurfaceTextureArrayV(loop1) = vBaseShift + SurfaceVertexArrayZ(loop1) / vBaseScale
    Next loop1
  End If
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

' recalculates texture co-ords, the base scale is the distance before texture is repeated
Public Sub Set_TexturePlaneYZ(Optional ByVal uBaseScale As Single = 1, Optional ByVal vBaseScale As Single = 1, _
        Optional ByVal uBaseShift As Single = 0, Optional ByVal vBaseShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False, Optional ByVal AutoFit As Boolean = True)
  
  Dim loop1 As Long, arrayTop As Long
  Dim minZ As Single, maxZ As Single, minY As Single, maxY As Single, uTemp As Single, vTemp As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceTextureArrayU(0 To arrayTop)
  ReDim nSurfaceTextureArrayV(0 To arrayTop)
  
  If AutoFit Then
    minZ = 9999999
    maxZ = -9999999
    minY = 9999999
    maxY = -9999999
    
    For loop1 = 0 To arrayTop
      If SurfaceVertexArrayY(loop1) < minY Then minY = SurfaceVertexArrayY(loop1)
      If SurfaceVertexArrayY(loop1) > maxY Then maxY = SurfaceVertexArrayY(loop1)
      If SurfaceVertexArrayZ(loop1) < minZ Then minZ = SurfaceVertexArrayZ(loop1)
      If SurfaceVertexArrayZ(loop1) > maxZ Then maxZ = SurfaceVertexArrayZ(loop1)
    Next loop1
    
    If FlipUV Then
      uTemp = maxZ - minZ
      vTemp = maxY - minY
      
      uBaseShift = uBaseShift - minZ / uTemp
      vBaseShift = vBaseShift - minY / vTemp
    Else
      uTemp = maxY - minY
      vTemp = maxZ - minZ
      
      uBaseShift = uBaseShift - minY / uTemp
      vBaseShift = vBaseShift - minZ / vTemp
    End If
  
    uBaseScale = uBaseScale * uTemp
    vBaseScale = vBaseScale * vTemp
  End If
  
  If FlipUV Then
    For loop1 = 0 To arrayTop
      nSurfaceTextureArrayU(loop1) = uBaseShift + SurfaceVertexArrayZ(loop1) / uBaseScale
      nSurfaceTextureArrayV(loop1) = vBaseShift + SurfaceVertexArrayY(loop1) / vBaseScale
    Next loop1
  Else
    For loop1 = 0 To arrayTop
      nSurfaceTextureArrayU(loop1) = uBaseShift + SurfaceVertexArrayY(loop1) / uBaseScale
      nSurfaceTextureArrayV(loop1) = vBaseShift + SurfaceVertexArrayZ(loop1) / vBaseScale
    Next loop1
  End If
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

' recalculates texture co-ords, the base scale is the distance before texture is repeated
Public Sub Set_TextureAxisX(ByVal axisScale As Single, Optional ByVal wrapScale As Single = 2 * Pi, _
        Optional ByVal axisShift As Single = 0, Optional ByVal wrapShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False)
  
  Dim loop1 As Long, arrayTop As Long, lastAngle As Single, thisAngle As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  
  ReDim nSurfaceTextureArrayU(0 To arrayTop)
  ReDim nSurfaceTextureArrayV(0 To arrayTop)
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  For loop1 = 0 To arrayTop
    If SurfaceVertexArrayZ(loop1) = 0 Then
      If SurfaceVertexArrayY(loop1) > 0 Then
        thisAngle = 0
      Else
        thisAngle = Pi
      End If
    ElseIf SurfaceVertexArrayZ(loop1) > 0 Then
      thisAngle = 0.5 * Pi - Atn(SurfaceVertexArrayY(loop1) / SurfaceVertexArrayZ(loop1))
    Else
      thisAngle = 1.5 * Pi - Atn(SurfaceVertexArrayY(loop1) / SurfaceVertexArrayZ(loop1))
    End If
    
    If loop1 > 0 Then
      If thisAngle - lastAngle > Pi Then
        Do While thisAngle - lastAngle > Pi
          thisAngle = thisAngle - 2 * Pi
        Loop
      ElseIf lastAngle - thisAngle > Pi Then
        Do While lastAngle - thisAngle > Pi
          thisAngle = thisAngle + 2 * Pi
        Loop
      End If
    End If
    
    If FlipUV Then
      nSurfaceTextureArrayV(loop1) = wrapShift + thisAngle / wrapScale
      nSurfaceTextureArrayU(loop1) = axisShift + SurfaceVertexArrayX(loop1) / axisScale
    Else
      nSurfaceTextureArrayU(loop1) = wrapShift + thisAngle / wrapScale
      nSurfaceTextureArrayV(loop1) = axisShift + SurfaceVertexArrayX(loop1) / axisScale
    End If
    
    lastAngle = thisAngle
  Next loop1
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

' recalculates texture co-ords, the base scale is the distance before texture is repeated
Public Sub Set_TextureAxisY(ByVal axisScale As Single, Optional ByVal wrapScale As Single = 2 * Pi, _
        Optional ByVal axisShift As Single = 0, Optional ByVal wrapShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False)
  
  Dim loop1 As Long, arrayTop As Long, lastAngle As Single, thisAngle As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single

  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceTextureArrayU(0 To arrayTop)
  ReDim nSurfaceTextureArrayV(0 To arrayTop)
  
  For loop1 = 0 To arrayTop
    If SurfaceVertexArrayX(loop1) = 0 Then
      If SurfaceVertexArrayZ(loop1) > 0 Then
        thisAngle = 0
      Else
        thisAngle = Pi
      End If
    ElseIf SurfaceVertexArrayX(loop1) > 0 Then
      thisAngle = 0.5 * Pi - Atn(SurfaceVertexArrayZ(loop1) / SurfaceVertexArrayX(loop1))
    Else
      thisAngle = 1.5 * Pi - Atn(SurfaceVertexArrayZ(loop1) / SurfaceVertexArrayX(loop1))
    End If
    
    If loop1 > 0 Then
      If thisAngle - lastAngle > Pi Then
        Do While thisAngle - lastAngle > Pi
          thisAngle = thisAngle - 2 * Pi
        Loop
      ElseIf lastAngle - thisAngle > Pi Then
        Do While lastAngle - thisAngle > Pi
          thisAngle = thisAngle + 2 * Pi
        Loop
      End If
    End If
    
    If FlipUV Then
      nSurfaceTextureArrayV(loop1) = wrapShift + thisAngle / wrapScale
      nSurfaceTextureArrayU(loop1) = axisShift + SurfaceVertexArrayY(loop1) / axisScale
    Else
      nSurfaceTextureArrayU(loop1) = wrapShift + thisAngle / wrapScale
      nSurfaceTextureArrayV(loop1) = axisShift + SurfaceVertexArrayY(loop1) / axisScale
    End If
    
    lastAngle = thisAngle
  Next loop1
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

' recalculates texture co-ords, the base scale is the distance before texture is repeated
Public Sub Set_TextureAxisZ(ByVal axisScale As Single, Optional ByVal wrapScale As Single = 2 * Pi, _
        Optional ByVal axisShift As Single = 0, Optional ByVal wrapShift As Single = 0, _
        Optional ByVal FlipUV As Boolean = False)
  
  Dim loop1 As Long, arrayTop As Long, lastAngle As Single, thisAngle As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceTextureArrayU(0 To arrayTop)
  ReDim nSurfaceTextureArrayV(0 To arrayTop)
    
  For loop1 = 0 To arrayTop
    If SurfaceVertexArrayX(loop1) = 0 Then
      If SurfaceVertexArrayY(loop1) > 0 Then
        thisAngle = 0
      Else
        thisAngle = Pi
      End If
    ElseIf SurfaceVertexArrayX(loop1) > 0 Then
      thisAngle = 0.5 * Pi - Atn(SurfaceVertexArrayY(loop1) / SurfaceVertexArrayX(loop1))
    Else
      thisAngle = 1.5 * Pi - Atn(SurfaceVertexArrayY(loop1) / SurfaceVertexArrayX(loop1))
    End If
    
    If loop1 > 0 Then
      If thisAngle - lastAngle > Pi Then
        Do While thisAngle - lastAngle > Pi
          thisAngle = thisAngle - 2 * Pi
        Loop
      ElseIf lastAngle - thisAngle > Pi Then
        Do While lastAngle - thisAngle > Pi
          thisAngle = thisAngle + 2 * Pi
        Loop
      End If
    End If
    
    If FlipUV Then
      nSurfaceTextureArrayV(loop1) = wrapShift + thisAngle / wrapScale
      nSurfaceTextureArrayU(loop1) = axisShift + SurfaceVertexArrayZ(loop1) / axisScale
    Else
      nSurfaceTextureArrayU(loop1) = wrapShift + thisAngle / wrapScale
      nSurfaceTextureArrayV(loop1) = axisShift + SurfaceVertexArrayZ(loop1) / axisScale
    End If
    
    lastAngle = thisAngle
  Next loop1
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

'calculates vertex normals for the surface - surfaceIsCircular indicates that first and last vertices
'match (Tri-Strip & Tri-Fan) or that matching vertices should have their normals softened (Tri-List)
Public Sub Calculate_Normals(Optional ByVal surfaceIsCircular As Boolean = False)
  Dim loop1 As Long, loop2 As Long, v1 As D3DVECTOR, v2 As D3DVECTOR, vResult As D3DVECTOR
  Dim arrayTop As Long, vPeak As D3DVECTOR, effectiveType As Long
  Dim triCount As Long, triIndex As Long
  Dim nX() As Single, nY() As Single, nz() As Single
  Dim ntX() As Single, ntY() As Single, ntZ() As Single
  
  Dim vTestHit() As Boolean, nTotalX As Single, nTotalY As Single, nTotalZ As Single
  Dim vTestX As Single, vTestY As Single, vTestZ As Single, vTotal As Long
    
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nX(0 To arrayTop)
  ReDim nY(0 To arrayTop)
  ReDim nz(0 To arrayTop)
  ReDim ntX(0 To arrayTop)
  ReDim ntY(0 To arrayTop)
  ReDim ntZ(0 To arrayTop)
  ReDim vTestHit(0 To arrayTop)
  
  If NumberOfSurfaceVertices = 3 Then
    effectiveType = 0
  Else
    effectiveType = SurfaceType
  End If
  
  Select Case effectiveType
    Case 0 'Tri-List
      triCount = NumberOfSurfaceVertices / 3
      
      For loop1 = 0 To triCount - 1
        triIndex = loop1 * 3
        
        With v1
          .X = SurfaceVertexArrayX(triIndex + 1) - SurfaceVertexArrayX(triIndex)
          .Y = SurfaceVertexArrayY(triIndex + 1) - SurfaceVertexArrayY(triIndex)
          .z = SurfaceVertexArrayZ(triIndex + 1) - SurfaceVertexArrayZ(triIndex)
        End With
        
        With v2
          .X = SurfaceVertexArrayX(triIndex + 2) - SurfaceVertexArrayX(triIndex)
          .Y = SurfaceVertexArrayY(triIndex + 2) - SurfaceVertexArrayY(triIndex)
          .z = SurfaceVertexArrayZ(triIndex + 2) - SurfaceVertexArrayZ(triIndex)
        End With
        
        dx_DirectX.VectorCrossProduct vResult, v1, v2
        dx_DirectX.VectorNormalize vResult
        
        With vResult
          nX(triIndex) = .X
          nY(triIndex) = .Y
          nz(triIndex) = .z
          
          nX(triIndex + 1) = .X
          nY(triIndex + 1) = .Y
          nz(triIndex + 1) = .z
          
          nX(triIndex + 2) = .X
          nY(triIndex + 2) = .Y
          nz(triIndex + 2) = .z
        End With
      Next loop1
      
      If surfaceIsCircular Then 'soften vertices
        For loop1 = 0 To arrayTop
          If vTestHit(loop1) <> True Then
            nTotalX = nX(loop1)
            nTotalY = nY(loop1)
            nTotalZ = nz(loop1)
            vTotal = 1
            
            vTestX = SurfaceVertexArrayX(loop1)
            vTestY = SurfaceVertexArrayY(loop1)
            vTestZ = SurfaceVertexArrayZ(loop1)
            
            For loop2 = loop1 + 1 To arrayTop
              If SurfaceVertexArrayX(loop2) = vTestX Then
                If SurfaceVertexArrayY(loop2) = vTestY Then
                  If SurfaceVertexArrayZ(loop2) = vTestZ Then
                    vTotal = vTotal + 1
                    
                    nTotalX = nTotalX + nX(loop2)
                    nTotalY = nTotalY + nY(loop2)
                    nTotalZ = nTotalZ + nz(loop2)
                  End If
                End If
              End If
            Next loop2
            
            nTotalX = nTotalX / vTotal
            nTotalY = nTotalY / vTotal
            nTotalZ = nTotalZ / vTotal
            
            For loop2 = loop1 To arrayTop
              If SurfaceVertexArrayX(loop2) = vTestX Then
                If SurfaceVertexArrayY(loop2) = vTestY Then
                  If SurfaceVertexArrayZ(loop2) = vTestZ Then
                    vTestHit(loop2) = True
                    
                    nX(loop2) = nTotalX
                    nY(loop2) = nTotalY
                    nz(loop2) = nTotalZ
                  End If
                End If
              End If
            Next loop2
          End If
        Next loop1
      End If
    Case 1 'Tri-Strip
      triCount = NumberOfSurfaceVertices - 1
      
      For triIndex = 0 To triCount - 2
        With v1
          .X = SurfaceVertexArrayX(triIndex + 1) - SurfaceVertexArrayX(triIndex)
          .Y = SurfaceVertexArrayY(triIndex + 1) - SurfaceVertexArrayY(triIndex)
          .z = SurfaceVertexArrayZ(triIndex + 1) - SurfaceVertexArrayZ(triIndex)
        End With
        
        With v2
          .X = SurfaceVertexArrayX(triIndex + 2) - SurfaceVertexArrayX(triIndex)
          .Y = SurfaceVertexArrayY(triIndex + 2) - SurfaceVertexArrayY(triIndex)
          .z = SurfaceVertexArrayZ(triIndex + 2) - SurfaceVertexArrayZ(triIndex)
        End With
        
        If (triIndex Mod 2) = 0 Then
          dx_DirectX.VectorCrossProduct vResult, v1, v2
        Else
          dx_DirectX.VectorCrossProduct vResult, v2, v1
        End If
        
        dx_DirectX.VectorNormalize vResult
        
        With vResult
          ntX(triIndex) = .X
          ntY(triIndex) = .Y
          ntZ(triIndex) = .z
        End With
      Next triIndex
      
      If surfaceIsCircular Then
        ntX(triCount) = ntX(1)
        ntY(triCount) = ntY(1)
        ntZ(triCount) = ntZ(1)
        
        ntX(triCount - 1) = ntX(0)
        ntY(triCount - 1) = ntY(0)
        ntZ(triCount - 1) = ntZ(0)
      Else
        ntX(triCount) = ntX(triCount - 2)
        ntY(triCount) = ntY(triCount - 2)
        ntZ(triCount) = ntZ(triCount - 2)
        
        ntX(triCount - 1) = ntX(triCount - 3)
        ntY(triCount - 1) = ntY(triCount - 3)
        ntZ(triCount - 1) = ntZ(triCount - 3)
      End If
      
      'now average out the 3 vertex normals for each facing
      If surfaceIsCircular Then
        nX(0) = (ntX(0) + ntX(triCount) + ntX(triCount - 1)) / 3
        nY(0) = (ntY(0) + ntZ(triCount) + ntY(triCount - 1)) / 3
        nz(0) = (ntZ(0) + ntZ(triCount) + ntZ(triCount - 1)) / 3
        
        nX(1) = (ntX(0) + ntX(triCount) + ntX(1)) / 3
        nY(1) = (ntY(0) + ntZ(triCount) + ntY(1)) / 3
        nz(1) = (ntZ(0) + ntZ(triCount) + ntZ(1)) / 3
      Else
        nX(0) = ntX(0)
        nY(0) = ntY(0)
        nz(0) = ntZ(0)
        
        nX(1) = (ntX(1) + ntX(0)) / 2
        nY(1) = (ntY(1) + ntY(0)) / 2
        nz(1) = (ntZ(1) + ntZ(0)) / 2
      End If
      
      For triIndex = 2 To triCount
        nX(triIndex) = (ntX(triIndex) + ntX(triIndex - 1) + ntX(triIndex - 2)) / 3
        nY(triIndex) = (ntY(triIndex) + ntZ(triIndex - 1) + ntY(triIndex - 2)) / 3
        nz(triIndex) = (ntZ(triIndex) + ntZ(triIndex - 1) + ntZ(triIndex - 2)) / 3
      Next triIndex
    Case 2 'Tri-Fan
      triCount = NumberOfSurfaceVertices - 1
      
      For triIndex = 1 To triCount - 1
        With v1
          .X = SurfaceVertexArrayX(triIndex) - SurfaceVertexArrayX(0)
          .Y = SurfaceVertexArrayY(triIndex) - SurfaceVertexArrayY(0)
          .z = SurfaceVertexArrayZ(triIndex) - SurfaceVertexArrayZ(0)
        End With
        
        With v2
          .X = SurfaceVertexArrayX(triIndex + 1) - SurfaceVertexArrayX(0)
          .Y = SurfaceVertexArrayY(triIndex + 1) - SurfaceVertexArrayY(0)
          .z = SurfaceVertexArrayZ(triIndex + 1) - SurfaceVertexArrayZ(0)
        End With
        
        dx_DirectX.VectorCrossProduct vResult, v1, v2
        dx_DirectX.VectorNormalize vResult
        
        With vPeak
          .X = .X + vResult.X
          .Y = .Y + vResult.Y
          .z = .z + vResult.z
        End With
        
        With vResult
          nX(triIndex) = .X
          nY(triIndex) = .Y
          nz(triIndex) = .z
        End With
      Next triIndex
      
      dx_DirectX.VectorNormalize vPeak
      
      If surfaceIsCircular Then
        nX(triCount) = nX(1)
        nY(triCount) = nY(1)
        nz(triCount) = nz(1)
      Else
        nX(triCount) = nX(triCount - 1)
        nY(triCount) = nY(triCount - 1)
        nz(triCount) = nz(triCount - 1)
      End If
      
      With vPeak
        nX(0) = .X
        nY(0) = .Y
        nz(0) = .z
      End With
  End Select
  
  SurfaceVertexArrayNX = nX
  SurfaceVertexArrayNY = nY
  SurfaceVertexArrayNZ = nz
End Sub

' converts a surface that is type TriStrip or TriFan to type TriList - this is nescessary for certain
' functions that manipulate the surface in complex fashions such as those functions which follow
Public Sub ConvertTo_TriList()
  Dim loop1 As Long, loop2 As Long, sTemp As Long, tTemp As Long
  Dim nNumberOfTriangles As Long, nNumberOfSurfaceVertices As Long, arrayTop As Long
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  
  If SurfaceType = 1 Or SurfaceType = 2 Then
    nNumberOfTriangles = NumberOfSurfaceVertices - 2
    nNumberOfSurfaceVertices = nNumberOfTriangles * 3
    arrayTop = nNumberOfSurfaceVertices - 1
      
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
    
    tTemp = 0
    
    For loop1 = 0 To nNumberOfTriangles - 1
      For loop2 = 0 To 2
        Select Case SurfaceType
          Case 1 ' TriStrip
            If (loop1 Mod 2) = 0 Then
              sTemp = loop1 + loop2
            Else
              sTemp = loop1 - loop2 + 2
            End If
          Case 2 ' TriFan
            Select Case loop2
              Case 0
                sTemp = 0
              Case 1
                sTemp = loop1 + 1
              Case Else
                sTemp = loop1 + 2
            End Select
        End Select
          
        nSurfaceVertexArrayX(tTemp) = SurfaceVertexArrayX(sTemp)
        nSurfaceVertexArrayY(tTemp) = SurfaceVertexArrayY(sTemp)
        nSurfaceVertexArrayZ(tTemp) = SurfaceVertexArrayZ(sTemp)
        
        nSurfaceVertexArrayRed(tTemp) = SurfaceVertexArrayRed(sTemp)
        nSurfaceVertexArrayGreen(tTemp) = SurfaceVertexArrayGreen(sTemp)
        nSurfaceVertexArrayBlue(tTemp) = SurfaceVertexArrayBlue(sTemp)
          
        nSurfaceVertexArrayNX(tTemp) = SurfaceVertexArrayNX(sTemp)
        nSurfaceVertexArrayNY(tTemp) = SurfaceVertexArrayNY(sTemp)
        nSurfaceVertexArrayNZ(tTemp) = SurfaceVertexArrayNZ(sTemp)
        
        nSurfaceTextureArrayU(tTemp) = SurfaceTextureArrayU(sTemp)
        nSurfaceTextureArrayV(tTemp) = SurfaceTextureArrayV(sTemp)
        
        tTemp = tTemp + 1
      Next loop2
    Next loop1
    
    NumberOfSurfaceVertices = nNumberOfSurfaceVertices
    
    SurfaceVertexArrayX = nSurfaceVertexArrayX
    SurfaceVertexArrayY = nSurfaceVertexArrayY
    SurfaceVertexArrayZ = nSurfaceVertexArrayZ
    
    SurfaceVertexArrayRed = nSurfaceVertexArrayRed
    SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
    SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
      
    SurfaceVertexArrayNX = nSurfaceVertexArrayNX
    SurfaceVertexArrayNY = nSurfaceVertexArrayNY
    SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
    
    SurfaceTextureArrayU = nSurfaceTextureArrayU
    SurfaceTextureArrayV = nSurfaceTextureArrayV
    
    SurfaceType = 0
  End If
End Sub

Public Sub ConvertTo_LineList()
  Dim loop1 As Long, loop2 As Long, sTemp As Long, tTemp As Long
  Dim nNumberOfLines As Long, nNumberOfSurfaceVertices As Long, arrayTop As Long
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  
  If SurfaceType = 5 Then
    nNumberOfLines = NumberOfSurfaceVertices - 1
    nNumberOfSurfaceVertices = 2 * nNumberOfLines
    arrayTop = nNumberOfSurfaceVertices - 1
    
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
    
    tTemp = 0
    
    For loop1 = 1 To nNumberOfLines
      sTemp = loop1 - 1
      
      For loop2 = 0 To 1
        nSurfaceVertexArrayX(tTemp) = SurfaceVertexArrayX(sTemp)
        nSurfaceVertexArrayY(tTemp) = SurfaceVertexArrayY(sTemp)
        nSurfaceVertexArrayZ(tTemp) = SurfaceVertexArrayZ(sTemp)
        
        nSurfaceVertexArrayRed(tTemp) = SurfaceVertexArrayRed(sTemp)
        nSurfaceVertexArrayGreen(tTemp) = SurfaceVertexArrayGreen(sTemp)
        nSurfaceVertexArrayBlue(tTemp) = SurfaceVertexArrayBlue(sTemp)
        
        nSurfaceVertexArrayNX(tTemp) = SurfaceVertexArrayNX(sTemp)
        nSurfaceVertexArrayNY(tTemp) = SurfaceVertexArrayNY(sTemp)
        nSurfaceVertexArrayNZ(tTemp) = SurfaceVertexArrayNZ(sTemp)
        
        nSurfaceTextureArrayU(tTemp) = SurfaceTextureArrayU(sTemp)
        nSurfaceTextureArrayV(tTemp) = SurfaceTextureArrayV(sTemp)
        
        tTemp = tTemp + 1
        sTemp = sTemp + 1
      Next loop2
    Next loop1
    
    NumberOfSurfaceVertices = nNumberOfSurfaceVertices
    
    SurfaceVertexArrayX = nSurfaceVertexArrayX
    SurfaceVertexArrayY = nSurfaceVertexArrayY
    SurfaceVertexArrayZ = nSurfaceVertexArrayZ
    
    SurfaceVertexArrayRed = nSurfaceVertexArrayRed
    SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
    SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
      
    SurfaceVertexArrayNX = nSurfaceVertexArrayNX
    SurfaceVertexArrayNY = nSurfaceVertexArrayNY
    SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
    
    SurfaceTextureArrayU = nSurfaceTextureArrayU
    SurfaceTextureArrayV = nSurfaceTextureArrayV
    
    SurfaceType = 4
  End If
End Sub

' joins the surface description of joiningSurface with this surface - if either surfaces are not of
' type TriList they will be converted first - the joiningSurface is not altered in any other way
' and may be used mutliple times on different surfaces as a template
Public Sub Join_Surfaces(ByVal joiningSurface As cls3D_ObjectSurface)
  Dim loop1 As Long, tloop1 As Long
  Dim nNumberOfSurfaceVertices As Long, arrayTop As Long
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  
  If SurfaceType = 3 Then
    If joiningSurface.SurfaceType <> 3 Then Exit Sub
  ElseIf SurfaceType > 3 Then
    If joiningSurface.SurfaceType < 4 Then Exit Sub
    
    ConvertTo_LineList
    joiningSurface.ConvertTo_LineList
  End If
  
  ConvertTo_TriList
  
  With joiningSurface
    .ConvertTo_TriList
    
    nNumberOfSurfaceVertices = NumberOfSurfaceVertices + .NumberOfSurfaceVertices
    
    arrayTop = nNumberOfSurfaceVertices - 1
      
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
    
    For loop1 = 0 To NumberOfSurfaceVertices - 1
      nSurfaceVertexArrayX(loop1) = SurfaceVertexArrayX(loop1)
      nSurfaceVertexArrayY(loop1) = SurfaceVertexArrayY(loop1)
      nSurfaceVertexArrayZ(loop1) = SurfaceVertexArrayZ(loop1)
      
      nSurfaceVertexArrayRed(loop1) = SurfaceVertexArrayRed(loop1)
      nSurfaceVertexArrayGreen(loop1) = SurfaceVertexArrayGreen(loop1)
      nSurfaceVertexArrayBlue(loop1) = SurfaceVertexArrayBlue(loop1)
        
      nSurfaceVertexArrayNX(loop1) = SurfaceVertexArrayNX(loop1)
      nSurfaceVertexArrayNY(loop1) = SurfaceVertexArrayNY(loop1)
      nSurfaceVertexArrayNZ(loop1) = SurfaceVertexArrayNZ(loop1)
      
      nSurfaceTextureArrayU(loop1) = SurfaceTextureArrayU(loop1)
      nSurfaceTextureArrayV(loop1) = SurfaceTextureArrayV(loop1)
    Next loop1
    
    For loop1 = 0 To .NumberOfSurfaceVertices - 1
      tloop1 = NumberOfSurfaceVertices + loop1
      
      nSurfaceVertexArrayX(tloop1) = .SurfaceVertexArrayX(loop1)
      nSurfaceVertexArrayY(tloop1) = .SurfaceVertexArrayY(loop1)
      nSurfaceVertexArrayZ(tloop1) = .SurfaceVertexArrayZ(loop1)
      
      nSurfaceVertexArrayRed(tloop1) = SurfaceVertexArrayRed(loop1)
      nSurfaceVertexArrayGreen(tloop1) = SurfaceVertexArrayGreen(loop1)
      nSurfaceVertexArrayBlue(tloop1) = SurfaceVertexArrayBlue(loop1)
        
      nSurfaceVertexArrayNX(tloop1) = .SurfaceVertexArrayNX(loop1)
      nSurfaceVertexArrayNY(tloop1) = .SurfaceVertexArrayNY(loop1)
      nSurfaceVertexArrayNZ(tloop1) = .SurfaceVertexArrayNZ(loop1)
      
      nSurfaceTextureArrayU(tloop1) = .SurfaceTextureArrayU(loop1)
      nSurfaceTextureArrayV(tloop1) = .SurfaceTextureArrayV(loop1)
    Next loop1
  End With
  
  NumberOfSurfaceVertices = nNumberOfSurfaceVertices
  
  SurfaceVertexArrayX = nSurfaceVertexArrayX
  SurfaceVertexArrayY = nSurfaceVertexArrayY
  SurfaceVertexArrayZ = nSurfaceVertexArrayZ
  
  SurfaceVertexArrayRed = nSurfaceVertexArrayRed
  SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
  SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
    
  SurfaceVertexArrayNX = nSurfaceVertexArrayNX
  SurfaceVertexArrayNY = nSurfaceVertexArrayNY
  SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

' splits the all the surface triangles into two - doubling surface complexity - if surface is not of type
' TriList it is converted first
Public Sub Double_SurfaceComplexity()
  Dim loop1 As Long, loop2 As Long, sTemp As Long, tTemp As Long
  Dim nNumberOfTriangles As Long, nNumberOfSurfaceVertices As Long, arrayTop As Long
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  
  Dim tSurfaceVertexArrayX As Single, tSurfaceVertexArrayY As Single, tSurfaceVertexArrayZ As Single
  Dim tSurfaceVertexArrayNX As Single, tSurfaceVertexArrayNY As Single, tSurfaceVertexArrayNZ As Single
  Dim tSurfaceTextureArrayU As Single, tSurfaceTextureArrayV As Single
  Dim tSurfaceVertexArrayRed As Single, tSurfaceVertexArrayGreen As Single, tSurfaceVertexArrayBlue As Single
  
  ConvertTo_TriList
  
  nNumberOfTriangles = NumberOfSurfaceVertices / 3
  nNumberOfSurfaceVertices = nNumberOfTriangles * 6
  arrayTop = nNumberOfSurfaceVertices - 1
      
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
    
  tTemp = 0
    
  For loop1 = 0 To nNumberOfTriangles - 1
    sTemp = loop1 * 3
    
    tSurfaceVertexArrayX = (SurfaceVertexArrayX(sTemp + 1) + SurfaceVertexArrayX(sTemp + 2)) / 2
    tSurfaceVertexArrayY = (SurfaceVertexArrayY(sTemp + 1) + SurfaceVertexArrayY(sTemp + 2)) / 2
    tSurfaceVertexArrayZ = (SurfaceVertexArrayZ(sTemp + 1) + SurfaceVertexArrayZ(sTemp + 2)) / 2
    
    tSurfaceVertexArrayRed = (SurfaceVertexArrayRed(sTemp + 1) + SurfaceVertexArrayRed(sTemp + 2)) / 2
    tSurfaceVertexArrayGreen = (SurfaceVertexArrayGreen(sTemp + 1) + SurfaceVertexArrayGreen(sTemp + 2)) / 2
    tSurfaceVertexArrayBlue = (SurfaceVertexArrayBlue(sTemp + 1) + SurfaceVertexArrayBlue(sTemp + 2)) / 2
    
    tSurfaceVertexArrayNX = (SurfaceVertexArrayNX(sTemp + 1) + SurfaceVertexArrayNX(sTemp + 2)) / 2
    tSurfaceVertexArrayNY = (SurfaceVertexArrayNY(sTemp + 1) + SurfaceVertexArrayNY(sTemp + 2)) / 2
    tSurfaceVertexArrayNZ = (SurfaceVertexArrayNZ(sTemp + 1) + SurfaceVertexArrayNZ(sTemp + 2)) / 2
    
    tSurfaceTextureArrayU = (SurfaceTextureArrayU(sTemp + 1) + SurfaceTextureArrayU(sTemp + 2)) / 2
    tSurfaceTextureArrayV = (SurfaceTextureArrayV(sTemp + 1) + SurfaceTextureArrayV(sTemp + 2)) / 2
    
    For loop2 = 0 To 5
      Select Case loop2
        Case 0, 3
          nSurfaceVertexArrayX(tTemp) = tSurfaceVertexArrayX
          nSurfaceVertexArrayY(tTemp) = tSurfaceVertexArrayY
          nSurfaceVertexArrayZ(tTemp) = tSurfaceVertexArrayZ
          
          nSurfaceVertexArrayRed(tTemp) = tSurfaceVertexArrayRed
          nSurfaceVertexArrayGreen(tTemp) = tSurfaceVertexArrayGreen
          nSurfaceVertexArrayBlue(tTemp) = tSurfaceVertexArrayBlue
            
          nSurfaceVertexArrayNX(tTemp) = tSurfaceVertexArrayNX
          nSurfaceVertexArrayNY(tTemp) = tSurfaceVertexArrayNY
          nSurfaceVertexArrayNZ(tTemp) = tSurfaceVertexArrayNZ
          
          nSurfaceTextureArrayU(tTemp) = tSurfaceTextureArrayU
          nSurfaceTextureArrayV(tTemp) = tSurfaceTextureArrayV
        Case Else
          Select Case loop2
            Case 1, 5
              sTemp = loop1 * 3
            Case 2
              sTemp = loop1 * 3 + 1
            Case Else
              sTemp = loop1 * 3 + 2
          End Select
          
          nSurfaceVertexArrayX(tTemp) = SurfaceVertexArrayX(sTemp)
          nSurfaceVertexArrayY(tTemp) = SurfaceVertexArrayY(sTemp)
          nSurfaceVertexArrayZ(tTemp) = SurfaceVertexArrayZ(sTemp)
          
          nSurfaceVertexArrayRed(tTemp) = SurfaceVertexArrayRed(sTemp)
          nSurfaceVertexArrayGreen(tTemp) = SurfaceVertexArrayGreen(sTemp)
          nSurfaceVertexArrayBlue(tTemp) = SurfaceVertexArrayBlue(sTemp)
          
          nSurfaceVertexArrayNX(tTemp) = SurfaceVertexArrayNX(sTemp)
          nSurfaceVertexArrayNY(tTemp) = SurfaceVertexArrayNY(sTemp)
          nSurfaceVertexArrayNZ(tTemp) = SurfaceVertexArrayNZ(sTemp)
          
          nSurfaceTextureArrayU(tTemp) = SurfaceTextureArrayU(sTemp)
          nSurfaceTextureArrayV(tTemp) = SurfaceTextureArrayV(sTemp)
      End Select
      
      tTemp = tTemp + 1
    Next loop2
  Next loop1
  
  NumberOfSurfaceVertices = nNumberOfSurfaceVertices
  
  SurfaceVertexArrayX = nSurfaceVertexArrayX
  SurfaceVertexArrayY = nSurfaceVertexArrayY
  SurfaceVertexArrayZ = nSurfaceVertexArrayZ
  
  SurfaceVertexArrayRed = nSurfaceVertexArrayRed
  SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
  SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
    
  SurfaceVertexArrayNX = nSurfaceVertexArrayNX
  SurfaceVertexArrayNY = nSurfaceVertexArrayNY
  SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

Public Sub Scale_Object(ByVal xScale As Single, yScale As Single, zScale As Single)
  Dim loop1 As Long, arrayTop As Long
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceVertexArrayX(0 To arrayTop)
  ReDim nSurfaceVertexArrayY(0 To arrayTop)
  ReDim nSurfaceVertexArrayZ(0 To arrayTop)
  
  For loop1 = 0 To arrayTop
    nSurfaceVertexArrayX(loop1) = xScale * SurfaceVertexArrayX(loop1)
    nSurfaceVertexArrayY(loop1) = yScale * SurfaceVertexArrayY(loop1)
    nSurfaceVertexArrayZ(loop1) = zScale * SurfaceVertexArrayZ(loop1)
  Next loop1
  
  SurfaceVertexArrayX = nSurfaceVertexArrayX
  SurfaceVertexArrayY = nSurfaceVertexArrayY
  SurfaceVertexArrayZ = nSurfaceVertexArrayZ
End Sub

' loads the SurfaceVertexArrayColour() with the colours set by the passed ambient light
' if preclearSurfaceLighting is FALSE then the values will be added to existing values, enabling
' muliple light sources
Public Sub Apply_VertexAmbient(ByVal sourceX As Single, ByVal sourceY As Single, ByVal sourceZ As Single, _
        Optional ByVal lightRange As Single = 0, Optional ByVal lightFalloff As Single = 0, _
        Optional ByVal redColour As Single = 0, Optional ByVal greenColour As Single = 0, Optional ByVal blueColour As Single = 0, _
        Optional ByVal preclearVertexLighting As Boolean = True)
  
  Dim arrayTop As Long, loop1 As Long
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  Dim distance As Single, distMod As Single
  Dim tRed As Single, tGreen As Single, tBlue As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceVertexArrayRed(0 To arrayTop)
  ReDim nSurfaceVertexArrayGreen(0 To arrayTop)
  ReDim nSurfaceVertexArrayBlue(0 To arrayTop)
  
  redColour = redColour * SurfaceRed
  greenColour = greenColour * SurfaceGreen
  blueColour = blueColour * SurfaceBlue
  
  If preclearVertexLighting Then
    If lightRange = 0 Then
      For loop1 = 0 To arrayTop
        If redColour <= 0 Then
          nSurfaceVertexArrayRed(loop1) = 0
        ElseIf redColour >= 1 Then
          nSurfaceVertexArrayRed(loop1) = 1
        Else
          nSurfaceVertexArrayRed(loop1) = redColour
        End If
        
        If greenColour <= 0 Then
          nSurfaceVertexArrayGreen(loop1) = 0
        ElseIf tGreen >= 1 Then
          nSurfaceVertexArrayGreen(loop1) = 1
        Else
          nSurfaceVertexArrayGreen(loop1) = greenColour
        End If
        
        If blueColour <= 0 Then
          nSurfaceVertexArrayBlue(loop1) = 0
        ElseIf tBlue >= 1 Then
          nSurfaceVertexArrayBlue(loop1) = 1
        Else
          nSurfaceVertexArrayBlue(loop1) = blueColour
        End If
      Next loop1
    Else
      For loop1 = 0 To arrayTop
        distance = Sqr((sourceX - SurfaceVertexArrayX(loop1)) ^ 2 + (sourceY - SurfaceVertexArrayY(loop1)) ^ 2 + (sourceZ - SurfaceVertexArrayZ(loop1)) ^ 2)
        
        If distance <= lightRange Then
          distMod = 1 - (distance / lightRange) * lightFalloff
        Else
          distMod = 1 - lightFalloff
        End If
        
        tRed = distMod * redColour
        
        If tRed <= 0 Then
          nSurfaceVertexArrayRed(loop1) = 0
        ElseIf tRed >= 1 Then
          nSurfaceVertexArrayRed(loop1) = 1
        Else
          nSurfaceVertexArrayRed(loop1) = tRed
        End If
        
        tGreen = distMod * greenColour
        
        If tGreen <= 0 Then
          nSurfaceVertexArrayGreen(loop1) = 0
        ElseIf tGreen >= 1 Then
          nSurfaceVertexArrayGreen(loop1) = 1
        Else
          nSurfaceVertexArrayGreen(loop1) = tGreen
        End If
        
        tBlue = distMod * blueColour
        
        If tBlue <= 0 Then
          nSurfaceVertexArrayBlue(loop1) = 0
        ElseIf tBlue >= 1 Then
          nSurfaceVertexArrayBlue(loop1) = 1
        Else
          nSurfaceVertexArrayBlue(loop1) = tBlue
        End If
      Next loop1
    End If
  Else
    If lightRange = 0 Then
      For loop1 = 0 To arrayTop
        tRed = SurfaceVertexArrayRed(loop1) + redColour
        
        If tRed <= 0 Then
          nSurfaceVertexArrayRed(loop1) = 0
        ElseIf tRed >= 1 Then
          nSurfaceVertexArrayRed(loop1) = 1
        Else
          nSurfaceVertexArrayRed(loop1) = tRed
        End If
        
        tGreen = SurfaceVertexArrayGreen(loop1) + greenColour
        
        If tGreen <= 0 Then
          nSurfaceVertexArrayGreen(loop1) = 0
        ElseIf tGreen >= 1 Then
          nSurfaceVertexArrayGreen(loop1) = 1
        Else
          nSurfaceVertexArrayGreen(loop1) = tGreen
        End If
        
        tBlue = SurfaceVertexArrayBlue(loop1) + blueColour
        
        If tBlue <= 0 Then
          nSurfaceVertexArrayBlue(loop1) = 0
        ElseIf tBlue >= 1 Then
          nSurfaceVertexArrayBlue(loop1) = 1
        Else
          nSurfaceVertexArrayBlue(loop1) = tBlue
        End If
      Next loop1
    Else
      For loop1 = 0 To arrayTop
        distance = Sqr((sourceX - SurfaceVertexArrayX(loop1)) ^ 2 + (sourceY - SurfaceVertexArrayY(loop1)) ^ 2 + (sourceZ - SurfaceVertexArrayZ(loop1)) ^ 2)
        
        If distance <= lightRange Then
          distMod = 1 - (distance / lightRange) * lightFalloff
        Else
          distMod = 1 - lightFalloff
        End If
        
        tRed = SurfaceVertexArrayRed(loop1) + distMod * redColour
        
        If tRed <= 0 Then
          nSurfaceVertexArrayRed(loop1) = 0
        ElseIf tRed >= 1 Then
          nSurfaceVertexArrayRed(loop1) = 1
        Else
          nSurfaceVertexArrayRed(loop1) = tRed
        End If
        
        tGreen = SurfaceVertexArrayGreen(loop1) + distMod * greenColour
        
        If tGreen <= 0 Then
          nSurfaceVertexArrayGreen(loop1) = 0
        ElseIf tGreen >= 1 Then
          nSurfaceVertexArrayGreen(loop1) = 1
        Else
          nSurfaceVertexArrayGreen(loop1) = tGreen
        End If
        
        tBlue = SurfaceVertexArrayBlue(loop1) + distMod * blueColour
        
        If tBlue <= 0 Then
          nSurfaceVertexArrayBlue(loop1) = 0
        ElseIf tBlue >= 1 Then
          nSurfaceVertexArrayBlue(loop1) = 1
        Else
          nSurfaceVertexArrayBlue(loop1) = tBlue
        End If
      Next loop1
    End If
  End If
  
  SurfaceVertexArrayRed = nSurfaceVertexArrayRed
  SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
  SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
End Sub

' loads the SurfaceVertexArrayColour() with the colours set by the passed diffuse light
' if preclearSurfaceLighting is FALSE then the values will be added to existing values, enabling
' muliple light sources
Public Sub Apply_VertexDiffuse(ByVal sourceX As Single, ByVal sourceY As Single, ByVal sourceZ As Single, _
        Optional ByVal lightRange As Single = 0, Optional ByVal lightFalloff As Single = 0, _
        Optional ByVal redColour As Single = 0, Optional ByVal greenColour As Single = 0, Optional ByVal blueColour As Single = 0, _
        Optional ByVal preclearVertexLighting As Boolean = True)
  
  Dim arrayTop As Long, loop1 As Long
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  Dim distance1 As Single, distance2 As Single
  Dim dVector As D3DVECTOR, effMod As Single, angleMod As Single
  Dim tRed As Single, tGreen As Single, tBlue As Single
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  ReDim nSurfaceVertexArrayRed(0 To arrayTop)
  ReDim nSurfaceVertexArrayGreen(0 To arrayTop)
  ReDim nSurfaceVertexArrayBlue(0 To arrayTop)
  
  redColour = redColour * SurfaceRed
  greenColour = greenColour * SurfaceGreen
  blueColour = blueColour * SurfaceBlue
  
  If preclearVertexLighting Then
    If lightRange = 0 Then
      For loop1 = 0 To arrayTop
        With dVector
          .X = sourceX - SurfaceVertexArrayX(loop1)
          .Y = sourceY - SurfaceVertexArrayY(loop1)
          .z = sourceZ - SurfaceVertexArrayZ(loop1)
          
          distance1 = Sqr(.X ^ 2 + .Y ^ 2 + .z ^ 2) + 0.0000001      'prevents div by zero error
          distance2 = Sqr(SurfaceVertexArrayNX(loop1) ^ 2 + SurfaceVertexArrayNY(loop1) ^ 2 + SurfaceVertexArrayNZ(loop1) ^ 2) + 0.0000001
          
          'this is an aproximated equation that produces close enough results to the real thing but reduces
          'equation calculations by a couple of steps and thus improves speed
          effMod = 1 - 0.5 * ((.X / distance1 - SurfaceVertexArrayNX(loop1) / distance2) ^ 2 + (.Y / distance1 - SurfaceVertexArrayNY(loop1) / distance2) ^ 2 + (.z / distance1 - SurfaceVertexArrayNZ(loop1) / distance2) ^ 2)
        End With
        
        If effMod > 0 Then
          tRed = redColour * effMod
          
          If tRed >= 1 Then
            nSurfaceVertexArrayRed(loop1) = 1
          Else
            nSurfaceVertexArrayRed(loop1) = tRed
          End If
          
          tGreen = greenColour * effMod
          
          If tGreen >= 1 Then
            nSurfaceVertexArrayGreen(loop1) = 1
          Else
            nSurfaceVertexArrayGreen(loop1) = tGreen
          End If
          
          tBlue = blueColour * effMod
          
          If tBlue >= 1 Then
            nSurfaceVertexArrayBlue(loop1) = 1
          Else
            nSurfaceVertexArrayBlue(loop1) = tBlue
          End If
        End If
      Next loop1
    Else
      For loop1 = 0 To arrayTop
        With dVector
          .X = sourceX - SurfaceVertexArrayX(loop1)
          .Y = sourceY - SurfaceVertexArrayY(loop1)
          .z = sourceZ - SurfaceVertexArrayZ(loop1)
          
          distance1 = Sqr(.X ^ 2 + .Y ^ 2 + .z ^ 2) + 0.0000001      'prevents div by zero
        End With
        
        distance2 = Sqr(SurfaceVertexArrayNX(loop1) ^ 2 + SurfaceVertexArrayNY(loop1) ^ 2 + SurfaceVertexArrayNZ(loop1) ^ 2) + 0.0000001
        
        If distance1 <= lightRange Then
          With dVector
            effMod = (1 - (distance1 / lightRange) * lightFalloff) * (1 - 0.5 * ((.X / distance1 - SurfaceVertexArrayNX(loop1) / distance2) ^ 2 + (.Y / distance1 - SurfaceVertexArrayNY(loop1) / distance2) ^ 2 + (.z / distance1 - SurfaceVertexArrayNZ(loop1) / distance2) ^ 2))
          End With
        Else
          With dVector
            effMod = (1 - lightFalloff) * (1 - 0.5 * ((.X / distance1 - SurfaceVertexArrayNX(loop1) / distance2) ^ 2 + (.Y / distance1 - SurfaceVertexArrayNY(loop1) / distance2) ^ 2 + (.z / distance1 - SurfaceVertexArrayNZ(loop1) / distance2) ^ 2))
          End With
        End If
        
        If effMod > 0 Then
          tRed = effMod * redColour
          
          If tRed <= 0 Then
            nSurfaceVertexArrayRed(loop1) = 0
          ElseIf tRed >= 1 Then
            nSurfaceVertexArrayRed(loop1) = 1
          Else
            nSurfaceVertexArrayRed(loop1) = tRed
          End If
          
          tGreen = effMod * greenColour
          
          If tGreen <= 0 Then
            nSurfaceVertexArrayGreen(loop1) = 0
          ElseIf tGreen >= 1 Then
            nSurfaceVertexArrayGreen(loop1) = 1
          Else
            nSurfaceVertexArrayGreen(loop1) = tGreen
          End If
          
          tBlue = effMod * blueColour
          
          If tBlue <= 0 Then
            nSurfaceVertexArrayBlue(loop1) = 0
          ElseIf tBlue >= 1 Then
            nSurfaceVertexArrayBlue(loop1) = 1
          Else
            nSurfaceVertexArrayBlue(loop1) = tBlue
          End If
        End If
      Next loop1
    End If
  Else
    If lightRange = 0 Then
      For loop1 = 0 To arrayTop
        With dVector
          .X = sourceX - SurfaceVertexArrayX(loop1)
          .Y = sourceY - SurfaceVertexArrayY(loop1)
          .z = sourceZ - SurfaceVertexArrayZ(loop1)
          
          distance1 = Sqr(.X ^ 2 + .Y ^ 2 + .z ^ 2) + 0.0000001      'prevents div by zero
          distance2 = Sqr(SurfaceVertexArrayNX(loop1) ^ 2 + SurfaceVertexArrayNY(loop1) ^ 2 + SurfaceVertexArrayNZ(loop1) ^ 2) + 0.0000001
          
          effMod = 1 - 0.5 * ((.X / distance1 - SurfaceVertexArrayNX(loop1) / distance2) ^ 2 + (.Y / distance1 - SurfaceVertexArrayNY(loop1) / distance2) ^ 2 + (.z / distance1 - SurfaceVertexArrayNZ(loop1) / distance2) ^ 2)
        End With
        
        If effMod > 0 Then
          tRed = SurfaceVertexArrayRed(loop1) + redColour * effMod
          
          If tRed <= 0 Then
            nSurfaceVertexArrayRed(loop1) = 0
          ElseIf tRed >= 1 Then
            nSurfaceVertexArrayRed(loop1) = 1
          Else
            nSurfaceVertexArrayRed(loop1) = tRed
          End If
          
          tGreen = SurfaceVertexArrayGreen(loop1) + greenColour * effMod
          
          If tGreen <= 0 Then
            nSurfaceVertexArrayGreen(loop1) = 0
          ElseIf tGreen >= 1 Then
            nSurfaceVertexArrayGreen(loop1) = 1
          Else
            nSurfaceVertexArrayGreen(loop1) = tGreen
          End If
          
          tBlue = SurfaceVertexArrayBlue(loop1) + blueColour * effMod
          
          If tBlue <= 0 Then
            nSurfaceVertexArrayBlue(loop1) = 0
          ElseIf tBlue >= 1 Then
            nSurfaceVertexArrayBlue(loop1) = 1
          Else
            nSurfaceVertexArrayBlue(loop1) = tBlue
          End If
        Else
          nSurfaceVertexArrayRed(loop1) = SurfaceVertexArrayRed(loop1)
          nSurfaceVertexArrayGreen(loop1) = SurfaceVertexArrayGreen(loop1)
          nSurfaceVertexArrayBlue(loop1) = SurfaceVertexArrayBlue(loop1)
        End If
      Next loop1
    Else
      For loop1 = 0 To arrayTop
        With dVector
          .X = sourceX - SurfaceVertexArrayX(loop1)
          .Y = sourceY - SurfaceVertexArrayY(loop1)
          .z = sourceZ - SurfaceVertexArrayZ(loop1)
          
          distance1 = Sqr(.X ^ 2 + .Y ^ 2 + .z ^ 2) + 0.0000001      'prevents div by zero
        End With
        
        distance2 = Sqr(SurfaceVertexArrayNX(loop1) ^ 2 + SurfaceVertexArrayNY(loop1) ^ 2 + SurfaceVertexArrayNZ(loop1) ^ 2) + 0.0000001
        
        If distance1 <= lightRange Then
          With dVector
            effMod = (1 - (distance1 / lightRange) * lightFalloff) * (1 - 0.5 * ((.X / distance1 - SurfaceVertexArrayNX(loop1) / distance2) ^ 2 + (.Y / distance1 - SurfaceVertexArrayNY(loop1) / distance2) ^ 2 + (.z / distance1 - SurfaceVertexArrayNZ(loop1) / distance2) ^ 2))
          End With
        Else
          With dVector
            effMod = (1 - lightFalloff) * (1 - 0.5 * ((.X / distance1 - SurfaceVertexArrayNX(loop1) / distance2) ^ 2 + (.Y / distance1 - SurfaceVertexArrayNY(loop1) / distance2) ^ 2 + (.z / distance1 - SurfaceVertexArrayNZ(loop1) / distance2) ^ 2))
          End With
        End If
        
        If effMod > 0 Then
          tRed = SurfaceVertexArrayRed(loop1) + effMod * redColour
          
          If tRed <= 0 Then
            nSurfaceVertexArrayRed(loop1) = 0
          ElseIf tRed >= 1 Then
            nSurfaceVertexArrayRed(loop1) = 1
          Else
            nSurfaceVertexArrayRed(loop1) = tRed
          End If
          
          tGreen = SurfaceVertexArrayGreen(loop1) + effMod * greenColour
          
          If tGreen <= 0 Then
            nSurfaceVertexArrayGreen(loop1) = 0
          ElseIf tGreen >= 1 Then
            nSurfaceVertexArrayGreen(loop1) = 1
          Else
            nSurfaceVertexArrayGreen(loop1) = tGreen
          End If
          
          tBlue = SurfaceVertexArrayBlue(loop1) + effMod * blueColour
          
          If tBlue <= 0 Then
            nSurfaceVertexArrayBlue(loop1) = 0
          ElseIf tBlue >= 1 Then
            nSurfaceVertexArrayBlue(loop1) = 1
          Else
            nSurfaceVertexArrayBlue(loop1) = tBlue
          End If
        Else
          nSurfaceVertexArrayRed(loop1) = SurfaceVertexArrayRed(loop1)
          nSurfaceVertexArrayGreen(loop1) = SurfaceVertexArrayGreen(loop1)
          nSurfaceVertexArrayBlue(loop1) = SurfaceVertexArrayBlue(loop1)
        End If
      Next loop1
    End If
  End If
  
  SurfaceVertexArrayRed = nSurfaceVertexArrayRed
  SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
  SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
End Sub

'ensures that all vertex arrays have the correct number of elements and also corrects for any problems
'that may occur from using the Array() function to assign values - which creates an array that can
'have its values read but not changed
Public Sub Normalize_Surface()
  Dim arrayTop As Long, loop1 As Long
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  
  On Error Resume Next
  
  arrayTop = NumberOfSurfaceVertices - 1
  
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
    nSurfaceVertexArrayX(loop1) = SurfaceVertexArrayX(loop1)
    nSurfaceVertexArrayY(loop1) = SurfaceVertexArrayY(loop1)
    nSurfaceVertexArrayZ(loop1) = SurfaceVertexArrayZ(loop1)
    
    nSurfaceVertexArrayRed(loop1) = SurfaceVertexArrayRed(loop1)
    nSurfaceVertexArrayGreen(loop1) = SurfaceVertexArrayGreen(loop1)
    nSurfaceVertexArrayBlue(loop1) = SurfaceVertexArrayBlue(loop1)
    
    nSurfaceVertexArrayNX(loop1) = SurfaceVertexArrayNX(loop1)
    nSurfaceVertexArrayNY(loop1) = SurfaceVertexArrayNY(loop1)
    nSurfaceVertexArrayNZ(loop1) = SurfaceVertexArrayNZ(loop1)
    
    nSurfaceTextureArrayU(loop1) = SurfaceTextureArrayU(loop1)
    nSurfaceTextureArrayV(loop1) = SurfaceTextureArrayV(loop1)
  Next loop1
  
  SurfaceVertexArrayX = nSurfaceVertexArrayX
  SurfaceVertexArrayY = nSurfaceVertexArrayY
  SurfaceVertexArrayZ = nSurfaceVertexArrayZ
  
  SurfaceVertexArrayRed = nSurfaceVertexArrayRed
  SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
  SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
    
  SurfaceVertexArrayNX = nSurfaceVertexArrayNX
  SurfaceVertexArrayNY = nSurfaceVertexArrayNY
  SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
End Sub

'returns a new surface that is a copy of this surface
Public Function Duplicate_Object() As cls3D_ObjectSurface
  Set Duplicate_Object = New cls3D_ObjectSurface
  
  With Duplicate_Object
    .Visible = Visible

    .TypeID = TypeID
    .SurfaceType = SurfaceType
    
    .TextureSurfaceIndex = TextureSurfaceIndex
                                                     
    .NumberOfSurfaceVertices = NumberOfSurfaceVertices
    .SurfaceVertexArrayX = SurfaceVertexArrayX
    .SurfaceVertexArrayY = SurfaceVertexArrayY
    .SurfaceVertexArrayZ = SurfaceVertexArrayZ
    .SurfaceVertexArrayNX = SurfaceVertexArrayNX
    .SurfaceVertexArrayNY = SurfaceVertexArrayNY
    .SurfaceVertexArrayNZ = SurfaceVertexArrayNZ
    .SurfaceTextureArrayU = SurfaceTextureArrayU
    .SurfaceTextureArrayV = SurfaceTextureArrayV
    .SurfaceVertexArrayRed = SurfaceVertexArrayRed
    .SurfaceVertexArrayGreen = SurfaceVertexArrayGreen
    .SurfaceVertexArrayBlue = SurfaceVertexArrayBlue
    
    .SurfaceOpacity = SurfaceOpacity
    .SurfaceRed = SurfaceRed
    .SurfaceGreen = SurfaceGreen
    .SurfaceBlue = SurfaceBlue
    .SurfaceEmissiveRed = SurfaceEmissiveRed
    .SurfaceEmissiveGreen = SurfaceEmissiveGreen
    .SurfaceEmissiveBlue = SurfaceEmissiveBlue
    
    '**********************************************************************************************
    'use this section for custom member variables
    '**********************************************************************************************
    
    .MainAction = MainAction
    .TicksPerMainAction = TicksPerMainAction
    .MainActionTickCounter = MainActionTickCounter
    .SubAction = SubAction
    .TicksPerSubAction = TicksPerSubAction
    .SubActionTickCounter = SubActionTickCounter
    
    .User1 = User1
    
    '**********************************************************************************************
  End With
End Function

Public Sub ReInitialize_D3D()
  Set dx_DirectX = DXDraw.GetDirectX()
  Set dx_Direct3D = DXDraw.GetDirect3D()
  Set dx_Direct3DDevice = DXDraw.GetDirect3DDevice()
End Sub

Private Sub Class_Initialize()
  ReInitialize_D3D
  
  Visible = True
  SurfaceOpacity = 1#
End Sub

Public Sub Destroy_Object()
  Set dx_DirectX = Nothing
  Set dx_Direct3D = Nothing
  Set dx_Direct3DDevice = Nothing
  
  Set NextSurface = Nothing
  Set PreviousSurface = Nothing
  Set ParentObject = Nothing
End Sub

' Saves D3D object surface definition to file
Public Sub Save_ToFile(filePathName As String)
  Dim fileHandle As Long
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Output As #fileHandle
    Write #fileHandle, "Logicon Enterprises cls3D_ObjectSurface Description Ver2.6"
    
    Save_ToOpenFile fileHandle
  Close fileHandle
  
  Exit Sub
  
badFile:
  MsgBox "Error writing surface to file.", vbCritical, "ERROR"
End Sub

' This should not be called directly... instead call Save_ToFile()
Public Sub Save_ToOpenFile(ByVal fileHandle As Long)
  Dim loop1 As Long, arrayTop As Long
  Dim sWidth As Long, sHeight As Long, sTransparent As Long, sIsTexture As Boolean
  
  Normalize_Surface
  
  Write #fileHandle, Visible

  Write #fileHandle, TypeID
  Write #fileHandle, SurfaceType
  
  Write #fileHandle, TextureSurfaceIndex
  
  If TextureSurfaceIndex <> 0 Then
    Write #fileHandle, DXDraw.Get_SSurfaceFileName(TextureSurfaceIndex)
    
    DXDraw.Get_SSurfaceSettings TextureSurfaceIndex, sWidth, sHeight, sTransparent, sIsTexture
    
    Write #fileHandle, sWidth
    Write #fileHandle, sHeight
    Write #fileHandle, sTransparent
  End If
  
  Write #fileHandle, NumberOfSurfaceVertices
  
  arrayTop = NumberOfSurfaceVertices - 1
  
  For loop1 = 0 To arrayTop
    Write #fileHandle, SurfaceVertexArrayX(loop1)
    Write #fileHandle, SurfaceVertexArrayY(loop1)
    Write #fileHandle, SurfaceVertexArrayZ(loop1)
    Write #fileHandle, SurfaceVertexArrayNX(loop1)
    Write #fileHandle, SurfaceVertexArrayNY(loop1)
    Write #fileHandle, SurfaceVertexArrayNZ(loop1)
    Write #fileHandle, SurfaceTextureArrayU(loop1)
    Write #fileHandle, SurfaceTextureArrayV(loop1)
    Write #fileHandle, SurfaceVertexArrayRed(loop1)
    Write #fileHandle, SurfaceVertexArrayGreen(loop1)
    Write #fileHandle, SurfaceVertexArrayBlue(loop1)
  Next loop1
  
  Write #fileHandle, SurfaceOpacity
  Write #fileHandle, SurfaceRed
  Write #fileHandle, SurfaceGreen
  Write #fileHandle, SurfaceBlue
  Write #fileHandle, SurfaceEmissiveRed
  Write #fileHandle, SurfaceEmissiveGreen
  Write #fileHandle, SurfaceEmissiveBlue
  
  '**********************************************************************************************
  'use this section for custom member variables
  '**********************************************************************************************
  
  Write #fileHandle, MainAction
  Write #fileHandle, TicksPerMainAction
  Write #fileHandle, MainActionTickCounter
  Write #fileHandle, SubAction
  Write #fileHandle, TicksPerSubAction
  Write #fileHandle, SubActionTickCounter
  
  Write #fileHandle, User1
  
  '**********************************************************************************************
End Sub

' Loads D3D object surface definition from file
Public Sub Load_FromFile(filePathName As String)
  Dim fileHandle As Long, versionString As String
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Input As #fileHandle
    Input #fileHandle, versionString
    
    Load_FromOpenFile fileHandle
  Close fileHandle
  
  Exit Sub
  
badFile:
  MsgBox "Error loading surface from file.", vbCritical, "ERROR"
End Sub

' This should not be called directly... instead call Load_FromFile()
Public Sub Load_FromOpenFile(ByVal fileHandle As Long)
  Dim loop1 As Long, arrayTop As Long
  Dim textureFile As String, textureWidth As Long, textureHeight As Long, textureTrans As Long
  
  Dim nSurfaceVertexArrayX() As Single, nSurfaceVertexArrayY() As Single, nSurfaceVertexArrayZ() As Single
  Dim nSurfaceVertexArrayNX() As Single, nSurfaceVertexArrayNY() As Single, nSurfaceVertexArrayNZ() As Single
  Dim nSurfaceTextureArrayU() As Single, nSurfaceTextureArrayV() As Single
  Dim nSurfaceVertexArrayRed() As Single, nSurfaceVertexArrayGreen() As Single, nSurfaceVertexArrayBlue() As Single
  
  Input #fileHandle, Visible

  Input #fileHandle, TypeID
  Input #fileHandle, SurfaceType

  Input #fileHandle, TextureSurfaceIndex
  
  If TextureSurfaceIndex <> 0 Then ' find the matching texture surface if loaded otherwise load it
    Input #fileHandle, textureFile
    Input #fileHandle, textureWidth
    Input #fileHandle, textureHeight
    Input #fileHandle, textureTrans
    
    TextureSurfaceIndex = DXDraw.Get_SSurfaceIndex(textureFile)
    
    If TextureSurfaceIndex = 0 Then
      TextureSurfaceIndex = DXDraw.Get_SSurfaceFreeIndex()
    
      If TextureSurfaceIndex <> 0 Then
        DXDraw.Init_TextureSurface TextureSurfaceIndex, textureFile, textureWidth, textureHeight, textureTrans
      End If
    End If
  End If
  
  Input #fileHandle, NumberOfSurfaceVertices
  
  arrayTop = NumberOfSurfaceVertices - 1
  
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
    Input #fileHandle, nSurfaceVertexArrayX(loop1)
    Input #fileHandle, nSurfaceVertexArrayY(loop1)
    Input #fileHandle, nSurfaceVertexArrayZ(loop1)
    Input #fileHandle, nSurfaceVertexArrayNX(loop1)
    Input #fileHandle, nSurfaceVertexArrayNY(loop1)
    Input #fileHandle, nSurfaceVertexArrayNZ(loop1)
    Input #fileHandle, nSurfaceTextureArrayU(loop1)
    Input #fileHandle, nSurfaceTextureArrayV(loop1)
    Input #fileHandle, nSurfaceVertexArrayRed(loop1)
    Input #fileHandle, nSurfaceVertexArrayGreen(loop1)
    Input #fileHandle, nSurfaceVertexArrayBlue(loop1)
  Next loop1
  
  SurfaceVertexArrayX = nSurfaceVertexArrayX
  SurfaceVertexArrayY = nSurfaceVertexArrayY
  SurfaceVertexArrayZ = nSurfaceVertexArrayZ
  
  SurfaceVertexArrayRed = nSurfaceVertexArrayRed
  SurfaceVertexArrayGreen = nSurfaceVertexArrayGreen
  SurfaceVertexArrayBlue = nSurfaceVertexArrayBlue
    
  SurfaceVertexArrayNX = nSurfaceVertexArrayNX
  SurfaceVertexArrayNY = nSurfaceVertexArrayNY
  SurfaceVertexArrayNZ = nSurfaceVertexArrayNZ
  
  SurfaceTextureArrayU = nSurfaceTextureArrayU
  SurfaceTextureArrayV = nSurfaceTextureArrayV
  
  Input #fileHandle, SurfaceOpacity
  Input #fileHandle, SurfaceRed
  Input #fileHandle, SurfaceGreen
  Input #fileHandle, SurfaceBlue
  Input #fileHandle, SurfaceEmissiveRed
  Input #fileHandle, SurfaceEmissiveGreen
  Input #fileHandle, SurfaceEmissiveBlue
  
  '**********************************************************************************************
  'use this section for custom member variables
  '**********************************************************************************************
  
  Input #fileHandle, MainAction
  Input #fileHandle, TicksPerMainAction
  Input #fileHandle, MainActionTickCounter
  Input #fileHandle, SubAction
  Input #fileHandle, TicksPerSubAction
  Input #fileHandle, SubActionTickCounter
  
  Input #fileHandle, User1
  
  '**********************************************************************************************
End Sub
