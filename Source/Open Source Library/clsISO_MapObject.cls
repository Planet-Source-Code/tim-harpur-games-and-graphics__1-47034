VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsISO_MapObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************************
'
' clsISO_MapObject class definition
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

Private Type MapCell
  staticSurface As Long ' if this value is 0 then this map cell will be rendered as OutBounds
  image As Long
  value As Long
  altitude As Long
  flatLock As Boolean
  
  ' all remaining variables are system calculated - do not alter
  yOffset0 As Long 'vertex offsets
  yOffset1 As Long
  yOffset2 As Long
  yOffset3 As Long
  yOffsetPeak As Long '(highest vertex - used to determine if cell needs to be displayed)
  
  altitudeB_1000ths As Long 'bottom vertex
  altitudeT_1000ths As Long 'top vertex
  altitudeL_1000ths As Long 'left vertex
  altitudeR_1000ths As Long 'right vertex
  altitudeC_1000ths As Long 'center of cell
  
  cellVertices(0 To 3) As D3DTLVERTEX 'vertex order -> 0 Left, 1 Top, 2 Bottom, 3 Right
  
  fogOfWar(0 To 3) As Long 'vertex fogOfWar status unexplored < 0, explored but grey = 0, in view > 0
  fogOfWarStatus As Long
  
  isoObjectChain As clsISO_Object 'chain of ISOmetric objects located at this map cell (level = 1)
  ghostWorldObjects As Boolean
End Type

Private MapArray() As MapCell
Private isoObjectChain As clsISO_Object 'chain of floating ISOmetric objects for entire map (level = 2)
Private isoObjectGroundFXChain As clsISO_Object  'chain of ground effect ISOmetric objects for entire map (level = 0)

Public DisplayColumn_1000ths As Long       ' Column co-ordinates of map to display
Public DisplayRow_1000ths As Long            ' Row co-ordinates of map to display

Public worldBase As Long                                 ' base size (in pixels) of cell

Public OutBoundsColour As Long               'ARGB

Public SurfacePixelRatioU As Single
Public SurfacePixelRatioV As Single
Public ImageWidthU As Single
Public ImageHeightV As Single
Public ImagesPerRow As Long

Public MapWidth As Long                                 ' width of the image map in image cells (read only)
Public MapHeight As Long                                ' height of the image map in image cells (read only)
Private mapTWidth As Long
Private mapTHeight As Long

Public minBrightness As Single
Public maxBrightness As Single
Public gridColour As Long                            'ARGB

Private tempVertices(0 To 7) As D3DTLVERTEX
Private validateMap() As Boolean

Private Type displayParameters 'pre-calculated values
  BasePosX As Long
  BasePosY As Long
  
  firstRow As Long
  firstCol As Long
  mapShiftRow As Long
  mapShiftCol As Long
  
  displayOriginXCenter As Long
  displayOriginX As Long
  displayOriginY As Long
  MapDisplayWidth As Long
  MapDisplayHeight As Long
  
  lastRowPos As Long
  
  effWorldBase As Long
  effWorldHalfBase As Long
  effWorldDoubleBase As Long
  worldBase1000 As Long
End Type

Private dParams As displayParameters

'***************************************************************************************************************
'
' The following section can be used to add custom member variables to this class.
' It is stongly advised that a COPY of this file be placed in the project directory, and that only the copy
' be customized and used. In this way the original class definition is preserved, while ensuring that each
' project has its own customized class definition that is not at risk of being modified by another project.
'
' Be certain to make changes to the routines Load_FromOpenFile and Save_ToOpenFile
' to reflect any changes made to the member variables here.
'
'***************************************************************************************************************




'***************************************************************************************************************



'Initializes the selected map array for indicated dimensions
'this must be called before the map rows can be loaded
Public Sub Init_MapArray(ByVal Width As Long, ByVal Height As Long)
  Dim loop1 As Long, loop2 As Long
  
  MapWidth = Width
  MapHeight = Height
  
  mapTWidth = Width + 5
  mapTHeight = Height + 5
  
  ReDim MapArray(-5 To mapTHeight - 1, -5 To mapTWidth - 1)
  
  For loop1 = -5 To mapTHeight - 1
    For loop2 = -5 To mapTWidth - 1
      With MapArray(loop1, loop2)
        .cellVertices(0).rhw = 1
        .cellVertices(1).rhw = 1
        .cellVertices(2).rhw = 1
        .cellVertices(3).rhw = 1
        
        .fogOfWar(0) = -1
        .fogOfWar(1) = -1
        .fogOfWar(2) = -1
        .fogOfWar(3) = -1
        
        .fogOfWarStatus = -1
      End With
    Next loop2
  Next loop1
End Sub

Public Sub Set_MapCell(ByVal row As Long, ByVal column As Long, ByVal staticSurface As Long, ByVal imageNumber As Long, ByVal cellValue As Long, ByVal altitude As Long, ByVal flatLock As Long)
  Dim imageXlow As Single, imageXhigh As Single, imageYlow As Single, imageYhigh As Single
  
  With MapArray(row, column)
    .staticSurface = staticSurface
    .image = imageNumber
    .value = cellValue
    .altitude = altitude
    .flatLock = flatLock
    
    imageXlow = (imageNumber Mod ImagesPerRow) * ImageWidthU + SurfacePixelRatioU
    imageXhigh = imageXlow + ImageWidthU - 2 * SurfacePixelRatioU
    imageYlow = (imageNumber \ ImagesPerRow) * ImageHeightV + SurfacePixelRatioV
    imageYhigh = imageYlow + ImageHeightV - 2 * SurfacePixelRatioV
    
    If .staticSurface = 0 Then
      .cellVertices(0).Color = OutBoundsColour
      .cellVertices(1).Color = OutBoundsColour
      .cellVertices(2).Color = OutBoundsColour
      .cellVertices(3).Color = OutBoundsColour
    Else
      .cellVertices(0).tU = imageXlow 'vertex order -> 0 Left, 1 Top, 2 Bottom, 3 Right
      .cellVertices(0).tV = imageYhigh
      
      .cellVertices(1).tU = imageXlow
      .cellVertices(1).tV = imageYlow
      
      .cellVertices(2).tU = imageXhigh
      .cellVertices(2).tV = imageYhigh
      
      .cellVertices(3).tU = imageXhigh
      .cellVertices(3).tV = imageYlow
    End If
  End With
End Sub

Public Function Get_MapCell_StaticSurface(ByVal row As Long, ByVal column As Long) As Long
  Get_MapCell_StaticSurface = MapArray(row, column).staticSurface
End Function

Public Sub Set_MapCell_StaticSurface(ByVal row As Long, ByVal column As Long, ByVal staticSurface As Long)
  With MapArray(row, column)
    .staticSurface = staticSurface
    
    If .staticSurface = 0 Then
      .cellVertices(0).Color = OutBoundsColour
      .cellVertices(1).Color = OutBoundsColour
      .cellVertices(2).Color = OutBoundsColour
      .cellVertices(3).Color = OutBoundsColour
    End If
  End With
End Sub

Public Function Get_MapCell_Image(ByVal row As Long, ByVal column As Long) As Long
  Get_MapCell_Image = MapArray(row, column).image
End Function

Public Sub Set_MapCell_Image(ByVal row As Long, ByVal column As Long, ByVal imageNumber As Long)
  Dim imageXlow As Single, imageXhigh As Single, imageYlow As Single, imageYhigh As Single
  
  With MapArray(row, column)
    .image = imageNumber
    
    imageXlow = (imageNumber Mod ImagesPerRow) * ImageWidthU + SurfacePixelRatioU
    imageXhigh = imageXlow + ImageWidthU - 2 * SurfacePixelRatioU
    imageYlow = (imageNumber \ ImagesPerRow) * ImageHeightV + SurfacePixelRatioV
    imageYhigh = imageYlow + ImageHeightV - 2 * SurfacePixelRatioV
    
    .cellVertices(0).tU = imageXlow 'vertex order -> 0 Left, 1 Top, 2 Bottom, 3 Right
    .cellVertices(0).tV = imageYhigh
    
    .cellVertices(1).tU = imageXlow
    .cellVertices(1).tV = imageYlow
    
    .cellVertices(2).tU = imageXhigh
    .cellVertices(2).tV = imageYhigh
    
    .cellVertices(3).tU = imageXhigh
    .cellVertices(3).tV = imageYlow
  End With
End Sub

Public Function Get_MapCell_Altitude(ByVal row As Long, ByVal column As Long) As Long
  Get_MapCell_Altitude = MapArray(row, column).altitude
End Function

Public Sub Set_MapCell_Altitude(ByVal row As Long, ByVal column As Long, ByVal altitude As Long)
  MapArray(row, column).altitude = altitude
End Sub

Public Function Get_MapCell_FlatLock(ByVal row As Long, ByVal column As Long) As Boolean
  Get_MapCell_FlatLock = MapArray(row, column).flatLock
End Function

Public Sub Set_MapCell_FlatLock(ByVal row As Long, ByVal column As Long, ByVal flatLock As Boolean)
  MapArray(row, column).flatLock = flatLock
End Sub

Public Function Get_MapCell_Value(ByVal row As Long, ByVal column As Long) As Long
  Get_MapCell_Value = MapArray(row, column).value
End Function

Public Sub Set_MapCell_Value(ByVal row As Long, ByVal column As Long, ByVal value As Long)
  MapArray(row, column).value = value
End Sub

Public Function Get_Altitude_from_MapRC(ByVal row As Long, ByVal column As Long) As Long
  Get_Altitude_from_MapRC = MapArray(row, column).altitudeC_1000ths
End Function

Public Function Get_Altitude_from_WorldXY(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long) As Long
  Dim row As Long, column As Long, modX As Single, modY As Single
  
  worldPosX_1000ths = worldPosX_1000ths \ worldBase
  worldPosY_1000ths = worldPosY_1000ths \ worldBase
  
  row = worldPosY_1000ths \ 1000
  column = worldPosX_1000ths \ 1000
  
  modX = (worldPosX_1000ths Mod 1000) / 1000#
  modY = (worldPosY_1000ths Mod 1000) / 1000#
  
  With MapArray(row, column)
    Get_Altitude_from_WorldXY = (.altitudeT_1000ths * (1 - modX) * (1 - modY)) + (.altitudeR_1000ths * modX * (1 - modY)) + (.altitudeL_1000ths * (1 - modX) * modY) + (.altitudeB_1000ths * modX * modY)
  End With
End Function

Public Sub Get_Rise_from_MapRC(ByVal row As Long, ByVal column As Long, ByRef xRise_1000ths As Long, ByRef yRise_1000ths As Long)
  With MapArray(row, column)
    xRise_1000ths = .altitudeR_1000ths - .altitudeL_1000ths
    yRise_1000ths = .altitudeB_1000ths - .altitudeT_1000ths
  End With
End Sub

Public Sub Get_Rise_from_WorldXY(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long, ByRef xRise_1000ths As Long, ByRef yRise_1000ths As Long)
  Dim row As Long, column As Long
  
  row = worldPosX_1000ths \ dParams.worldBase1000
  column = worldPosY_1000ths \ dParams.worldBase1000
  
  With MapArray(row, column)
    xRise_1000ths = .altitudeR_1000ths - .altitudeL_1000ths
    yRise_1000ths = .altitudeB_1000ths - .altitudeT_1000ths
  End With
End Sub

Public Sub Get_MapRC_from_WorldXY(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long, _
        ByRef row As Long, ByRef column As Long)
        
  column = worldPosX_1000ths \ dParams.worldBase1000
  row = worldPosY_1000ths \ dParams.worldBase1000
End Sub

Public Sub Get_WorldXYZ_from_MapRC(ByVal row As Long, ByVal column As Long, _
        ByRef worldPosX_1000ths As Long, ByRef worldPosY_1000ths As Long, ByRef worldPosZ_1000ths As Long)
        
  worldPosX_1000ths = (column * worldBase + worldBase \ 2) * 1000
  worldPosY_1000ths = (row * worldBase + worldBase \ 2) * 1000
  worldPosZ_1000ths = MapArray(row, column).altitudeC_1000ths
End Sub

Public Sub Get_WorldXYZ_from_DisplayXY(ByVal xPos As Long, ByVal yPos As Long, ByVal maxAltitude As Long, ByRef worldPosX_1000ths As Long, ByRef worldPosY_1000ths As Long, ByRef worldPosZ_1000ths As Long)
  Dim deltaY As Long, baseX As Long, baseY As Long, row As Long, column As Long
  
  baseX = ((yPos + (xPos - dParams.displayOriginXCenter) \ 2) \ DXDraw.ISOmetricViewScale) + (DisplayColumn_1000ths * worldBase) \ 1000
  baseY = ((yPos + (dParams.displayOriginXCenter - xPos) \ 2) \ DXDraw.ISOmetricViewScale) + (DisplayRow_1000ths * worldBase) \ 1000
  
  For deltaY = 0 To maxAltitude
    row = (baseY + deltaY) \ worldBase
    column = (baseX + deltaY) \ worldBase
    
    If row >= -5 And column >= -5 And row < mapTHeight And column < mapTWidth Then
      With MapArray(row, column)
        If .altitude <= deltaY Then
          If .altitude > deltaY - worldBase Then
            worldPosX_1000ths = (baseX + deltaY) * 1000
            worldPosY_1000ths = (baseY + deltaY) * 1000
            worldPosZ_1000ths = .altitude * 1000
            
            Exit Sub
          End If
        End If
      End With
    ElseIf row >= mapTHeight Or column >= mapTWidth Then
      worldPosX_1000ths = -1000000000
      worldPosY_1000ths = -1000000000
      worldPosZ_1000ths = -1000000000
      
      Exit Sub
    End If
  Next deltaY
    
  worldPosX_1000ths = -1000000000
  worldPosY_1000ths = -1000000000
  worldPosZ_1000ths = -1000000000
End Sub

Public Sub Set_MapDisplay_from_WorldXY(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long)
  DisplayColumn_1000ths = worldPosX_1000ths \ worldBase
  DisplayRow_1000ths = worldPosY_1000ths \ worldBase
End Sub

Public Sub Get_WorldXY_from_MapDisplay(ByRef worldPosX_1000ths As Long, ByRef worldPosY_1000ths As Long)
  worldPosX_1000ths = DisplayColumn_1000ths * worldBase
  worldPosY_1000ths = DisplayRow_1000ths * worldBase
End Sub

Public Sub Ghost_MapCell(ByVal row As Long, ByVal column As Long)
  MapArray(row, column).ghostWorldObjects = True
End Sub

Public Function Get_FogOfWar_MapRC(ByVal row As Long, ByVal column As Long) As Long
  If row < -5 Or row >= mapTHeight Or column < -5 Or column >= mapTWidth Then
    Get_FogOfWar_MapRC = -1
    
    Exit Function
  End If
  
  Get_FogOfWar_MapRC = MapArray(row, column).fogOfWarStatus
End Function

Public Function Get_FogOfWar_WorldXY(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long) As Long
  Dim row As Long, column As Long
  
  column = worldPosX_1000ths \ dParams.worldBase1000
  row = worldPosY_1000ths \ dParams.worldBase1000
  
  If row < -5 Or row >= mapTHeight Or column < -5 Or column >= mapTWidth Then
    Get_FogOfWar_WorldXY = -1
    
    Exit Function
  End If
  
  Get_FogOfWar_WorldXY = MapArray(row, column).fogOfWarStatus
End Function

Public Sub ClearAll_Fog()
  Dim row As Long, col As Long
  
  For row = mapTHeight - 1 To -5 Step -1
    For col = mapTWidth - 1 To -5 Step -1
      With MapArray(row, col)
        .fogOfWar(0) = 1
        .fogOfWar(1) = 1
        .fogOfWar(2) = 1
        .fogOfWar(3) = 1
        
        .fogOfWarStatus = 1
      End With
    Next col
  Next row
End Sub

Public Sub Clear_Fog(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long, ByVal Range_1000ths As Long)
  Dim loop1 As Long, loop2 As Long, maxLoop As Long, baseRow As Long, baseCol As Long
  Dim rowS As Single, colS As Single, rangeS As Single
  Dim minRow As Long, maxRow As Long, minCol As Long, maxCol As Long
  Dim rS0 As Single, rS1 As Single, cS0 As Single, cS1 As Single
  
  maxLoop = (Range_1000ths \ dParams.worldBase1000) + 1
  baseRow = worldPosY_1000ths \ dParams.worldBase1000
  baseCol = worldPosX_1000ths \ dParams.worldBase1000
  
  rowS = worldPosY_1000ths / dParams.worldBase1000
  colS = worldPosX_1000ths / dParams.worldBase1000
  rangeS = (Range_1000ths / dParams.worldBase1000) ^ 2
  
  minRow = baseRow - maxLoop
  If minRow < -5 Then
    minRow = -5
  ElseIf minRow >= mapTHeight Then
    Exit Sub
  End If
  
  maxRow = baseRow + maxLoop
  If maxRow < -5 Then
    Exit Sub
  ElseIf maxRow >= mapTHeight Then
    maxRow = mapTHeight - 1
  End If
  
  minCol = baseCol - maxLoop
  If minCol < -5 Then
    minCol = -5
  ElseIf minCol >= mapTWidth Then
    Exit Sub
  End If
  
  maxCol = baseCol + maxLoop
  If maxCol < -5 Then
    Exit Sub
  ElseIf maxCol >= mapTWidth Then
    maxCol = mapTWidth - 1
  End If
  
  For loop1 = minRow To maxRow
    rS0 = (rowS - loop1) ^ 2
    rS1 = (rowS - (loop1 + 1)) ^ 2
    
    For loop2 = minCol To maxCol
      cS0 = (colS - loop2) ^ 2
      cS1 = (colS - (loop2 + 1)) ^ 2
      
      With MapArray(loop1, loop2)
        If (rS1 + cS0) <= rangeS Then  'vertex 0
          If .fogOfWar(0) <= 0 Then
            .fogOfWar(0) = 1
          Else
            .fogOfWar(0) = .fogOfWar(0) + 1
          End If
        End If
        
        If (rS0 + cS0) <= rangeS Then  'vertex 1
          If .fogOfWar(1) <= 0 Then
            .fogOfWar(1) = 1
          Else
            .fogOfWar(1) = .fogOfWar(1) + 1
          End If
        End If
        
        If (rS1 + cS1) <= rangeS Then  'vertex 2
          If .fogOfWar(2) < 0 Then
            .fogOfWar(2) = 1
          Else
            .fogOfWar(2) = .fogOfWar(2) + 1
          End If
        End If
        
        If (rS0 + cS1) <= rangeS Then  'vertex 3
          If .fogOfWar(3) < 0 Then
            .fogOfWar(3) = 1
          Else
            .fogOfWar(3) = .fogOfWar(3) + 1
          End If
        End If
        
        If (.fogOfWar(0) + .fogOfWar(1) + .fogOfWar(2) + .fogOfWar(3)) = -4 Then
          .fogOfWarStatus = -1
        ElseIf (.fogOfWar(0) > 0) Or (.fogOfWar(1) > 0) Or (.fogOfWar(2) > 0) Or (.fogOfWar(3) > 0) Then
          .fogOfWarStatus = 1
        Else
          .fogOfWarStatus = 0
        End If
      End With
    Next loop2
  Next loop1
End Sub

Public Sub ApplyAll_FogThin()
  Dim row As Long, col As Long
  
  For row = mapTHeight - 1 To -5 Step -1
    For col = mapTWidth - 1 To -5 Step -1
      With MapArray(row, col)
        .fogOfWar(0) = 0
        .fogOfWar(1) = 0
        .fogOfWar(2) = 0
        .fogOfWar(3) = 0
        
        .fogOfWarStatus = 0
      End With
    Next col
  Next row
End Sub

Public Sub Apply_FogThin(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long, ByVal Range_1000ths As Long)
  Dim loop1 As Long, loop2 As Long, maxLoop As Long, baseRow As Long, baseCol As Long
  Dim rowS As Single, colS As Single, rangeS As Single
  Dim minRow As Long, maxRow As Long, minCol As Long, maxCol As Long, fogMask As Long
  Dim rS0 As Single, rS1 As Single, cS0 As Single, cS1 As Single
  
  maxLoop = (Range_1000ths \ dParams.worldBase1000) + 1
  baseRow = worldPosY_1000ths \ dParams.worldBase1000
  baseCol = worldPosX_1000ths \ dParams.worldBase1000
  
  rowS = worldPosY_1000ths / dParams.worldBase1000
  colS = worldPosX_1000ths / dParams.worldBase1000
  rangeS = (Range_1000ths / dParams.worldBase1000) ^ 2
  
  minRow = baseRow - maxLoop
  If minRow < -5 Then
    minRow = -5
  ElseIf minRow >= mapTHeight Then
    Exit Sub
  End If
  
  maxRow = baseRow + maxLoop
  If maxRow < -5 Then
    Exit Sub
  ElseIf maxRow >= mapTHeight Then
    maxRow = mapTHeight - 1
  End If
  
  minCol = baseCol - maxLoop
  If minCol < -5 Then
    minCol = -5
  ElseIf minCol >= mapTWidth Then
    Exit Sub
  End If
  
  maxCol = baseCol + maxLoop
  If maxCol < -5 Then
    Exit Sub
  ElseIf maxCol >= mapTWidth Then
    maxCol = mapTWidth - 1
  End If
  
  For loop1 = minRow To maxRow
    rS0 = (rowS - loop1) ^ 2
    rS1 = (rowS - (loop1 + 1)) ^ 2
    
    For loop2 = minCol To maxCol
      cS0 = (colS - loop2) ^ 2
      cS1 = (colS - (loop2 + 1)) ^ 2
      
      With MapArray(loop1, loop2)
        If (rS1 + cS0) <= rangeS Then .fogOfWar(0) = 0 'vertex 0
        If (rS0 + cS0) <= rangeS Then .fogOfWar(1) = 0 'vertex 1
        If (rS1 + cS1) <= rangeS Then .fogOfWar(2) = 0 'vertex 2
        If (rS0 + cS1) <= rangeS Then .fogOfWar(3) = 0 'vertex 3
        
        If (.fogOfWar(0) + .fogOfWar(1) + .fogOfWar(2) + .fogOfWar(3)) = -4 Then
          .fogOfWarStatus = -1
        ElseIf (.fogOfWar(0) > 0) Or (.fogOfWar(1) > 0) Or (.fogOfWar(2) > 0) Or (.fogOfWar(3) > 0) Then
          .fogOfWarStatus = 1
        Else
          .fogOfWarStatus = 0
        End If
      End With
    Next loop2
  Next loop1
End Sub

Public Sub RestoreAll_Fog()
  Dim row As Long, col As Long
  
  For row = mapTHeight - 1 To -5 Step -1
    For col = mapTWidth - 1 To -5 Step -1
      With MapArray(row, col)
        If .fogOfWar(0) > 0 Then .fogOfWar(0) = 0
        If .fogOfWar(1) > 0 Then .fogOfWar(1) = 0
        If .fogOfWar(2) > 0 Then .fogOfWar(2) = 0
        If .fogOfWar(3) > 0 Then .fogOfWar(3) = 0
        
        If (.fogOfWar(0) + .fogOfWar(1) + .fogOfWar(2) + .fogOfWar(3)) = -4 Then
          .fogOfWarStatus = -1
        ElseIf (.fogOfWar(0) > 0) Or (.fogOfWar(1) > 0) Or (.fogOfWar(2) > 0) Or (.fogOfWar(3) > 0) Then
          .fogOfWarStatus = 1
        Else
          .fogOfWarStatus = 0
        End If
      End With
    Next col
  Next row
End Sub

Public Sub Restore_Fog(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long, ByVal Range_1000ths As Long)
  Dim loop1 As Long, loop2 As Long, maxLoop As Long, baseRow As Long, baseCol As Long
  Dim rowS As Single, colS As Single, rangeS As Single
  Dim minRow As Long, maxRow As Long, minCol As Long, maxCol As Long, fogMask As Long
  Dim rS0 As Single, rS1 As Single, cS0 As Single, cS1 As Single
  
  maxLoop = (Range_1000ths \ dParams.worldBase1000) + 1
  baseRow = worldPosY_1000ths \ dParams.worldBase1000
  baseCol = worldPosX_1000ths \ dParams.worldBase1000
  
  rowS = worldPosY_1000ths / dParams.worldBase1000
  colS = worldPosX_1000ths / dParams.worldBase1000
  rangeS = (Range_1000ths / dParams.worldBase1000) ^ 2
  
  minRow = baseRow - maxLoop
  If minRow < -5 Then
    minRow = -5
  ElseIf minRow >= mapTHeight Then
    Exit Sub
  End If
  
  maxRow = baseRow + maxLoop
  If maxRow < -5 Then
    Exit Sub
  ElseIf maxRow >= mapTHeight Then
    maxRow = mapTHeight - 1
  End If
  
  minCol = baseCol - maxLoop
  If minCol < -5 Then
    minCol = -5
  ElseIf minCol >= mapTWidth Then
    Exit Sub
  End If
  
  maxCol = baseCol + maxLoop
  If maxCol < -5 Then
    Exit Sub
  ElseIf maxCol >= mapTWidth Then
    maxCol = mapTWidth - 1
  End If
  
  For loop1 = minRow To maxRow
    rS0 = (rowS - loop1) ^ 2
    rS1 = (rowS - (loop1 + 1)) ^ 2
    
    For loop2 = minCol To maxCol
      cS0 = (colS - loop2) ^ 2
      cS1 = (colS - (loop2 + 1)) ^ 2
      
      With MapArray(loop1, loop2)
        If (rS1 + cS0) <= rangeS Then If .fogOfWar(0) > 0 Then .fogOfWar(0) = .fogOfWar(0) - 1 'vertex 0
        If (rS0 + cS0) <= rangeS Then If .fogOfWar(1) > 0 Then .fogOfWar(1) = .fogOfWar(1) - 1 'vertex 1
        If (rS1 + cS1) <= rangeS Then If .fogOfWar(2) > 0 Then .fogOfWar(2) = .fogOfWar(2) - 1 'vertex 2
        If (rS0 + cS1) <= rangeS Then If .fogOfWar(3) > 0 Then .fogOfWar(3) = .fogOfWar(3) - 1 'vertex 3
        
        If (.fogOfWar(0) + .fogOfWar(1) + .fogOfWar(2) + .fogOfWar(3)) = -4 Then
          .fogOfWarStatus = -1
        ElseIf (.fogOfWar(0) > 0) Or (.fogOfWar(1) > 0) Or (.fogOfWar(2) > 0) Or (.fogOfWar(3) > 0) Then
          .fogOfWarStatus = 1
        Else
          .fogOfWarStatus = 0
        End If
      End With
    Next loop2
  Next loop1
End Sub

Public Sub ApplyAll_FogThick()
  Dim row As Long, col As Long
  
  For row = mapTHeight - 1 To -5 Step -1
    For col = mapTWidth - 1 To -5 Step -1
      With MapArray(row, col)
        .fogOfWar(0) = -1 'vertex 0
        .fogOfWar(1) = -1 'vertex 1
        .fogOfWar(2) = -1 'vertex 2
        .fogOfWar(3) = -1 'vertex 3
        
        .fogOfWarStatus = -1
      End With
    Next col
  Next row
End Sub

Public Sub Apply_FogThick(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long, ByVal Range_1000ths As Long)
  Dim loop1 As Long, loop2 As Long, maxLoop As Long, baseRow As Long, baseCol As Long
  Dim rowS As Single, colS As Single, rangeS As Single
  Dim minRow As Long, maxRow As Long, minCol As Long, maxCol As Long
  Dim rS0 As Single, rS1 As Single, cS0 As Single, cS1 As Single
  
  maxLoop = (Range_1000ths \ dParams.worldBase1000) + 1
  baseRow = worldPosY_1000ths \ dParams.worldBase1000
  baseCol = worldPosX_1000ths \ dParams.worldBase1000
  
  rowS = worldPosY_1000ths / dParams.worldBase1000
  colS = worldPosX_1000ths / dParams.worldBase1000
  rangeS = (Range_1000ths / dParams.worldBase1000) ^ 2
  
  minRow = baseRow - maxLoop
  If minRow < -5 Then
    minRow = -5
  ElseIf minRow >= mapTHeight Then
    Exit Sub
  End If
  
  maxRow = baseRow + maxLoop
  If maxRow < -5 Then
    Exit Sub
  ElseIf maxRow >= mapTHeight Then
    maxRow = mapTHeight - 1
  End If
  
  minCol = baseCol - maxLoop
  If minCol < -5 Then
    minCol = -5
  ElseIf minCol >= mapTWidth Then
    Exit Sub
  End If
  
  maxCol = baseCol + maxLoop
  If maxCol < -5 Then
    Exit Sub
  ElseIf maxCol >= mapTWidth Then
    maxCol = mapTWidth - 1
  End If
  
  For loop1 = minRow To maxRow
    rS0 = (rowS - loop1) ^ 2
    rS1 = (rowS - (loop1 + 1)) ^ 2
    
    For loop2 = minCol To maxCol
      cS0 = (colS - loop2) ^ 2
      cS1 = (colS - (loop2 + 1)) ^ 2
      
      With MapArray(loop1, loop2)
        If (rS1 + cS0) <= rangeS Then .fogOfWar(0) = -1 'vertex 0
        If (rS0 + cS0) <= rangeS Then .fogOfWar(1) = -1 'vertex 1
        If (rS1 + cS1) <= rangeS Then .fogOfWar(2) = -1 'vertex 2
        If (rS0 + cS1) <= rangeS Then .fogOfWar(3) = -1 'vertex 3
        
        If (.fogOfWar(0) + .fogOfWar(1) + .fogOfWar(2) + .fogOfWar(3)) = -4 Then
          .fogOfWarStatus = -1
        ElseIf (.fogOfWar(0) > 0) Or (.fogOfWar(1) > 0) Or (.fogOfWar(2) > 0) Or (.fogOfWar(3) > 0) Then
          .fogOfWarStatus = 1
        Else
          .fogOfWarStatus = 0
        End If
      End With
    Next loop2
  Next loop1
End Sub

Public Sub Apply_Light(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long, ByVal Range_1000ths As Long, ByVal Colour As Long)
  Dim loop1 As Long, loop2 As Long, maxLoop As Long, baseRow As Long, baseCol As Long
  Dim redC As Long, blueC As Long, greenC As Long
  Dim rowS As Single, colS As Single, rangeS As Single
  Dim minRow As Long, maxRow As Long, minCol As Long, maxCol As Long
  Dim rS0 As Single, rS1 As Single, cS0 As Single, cS1 As Single
  Dim rColour As Long, gColour As Long, bColour As Long, lightEffect As Single
  
  maxLoop = (Range_1000ths \ dParams.worldBase1000) + 1
  baseRow = worldPosY_1000ths \ dParams.worldBase1000
  baseCol = worldPosX_1000ths \ dParams.worldBase1000
  
  rowS = worldPosY_1000ths / dParams.worldBase1000
  colS = worldPosX_1000ths / dParams.worldBase1000
  rangeS = Range_1000ths / dParams.worldBase1000
  
  redC = Colour \ 65536
  greenC = (Colour And &HFF00&) \ 256
  blueC = Colour And &HFF&
  
  minRow = baseRow - maxLoop
  If minRow < 0 Then
    minRow = 0
  ElseIf minRow >= MapHeight Then
    Exit Sub
  End If
  
  maxRow = baseRow + maxLoop
  If maxRow < 0 Then
    Exit Sub
  ElseIf maxRow >= MapHeight Then
    maxRow = MapHeight - 1
  End If
  
  minCol = baseCol - maxLoop
  If minCol < 0 Then
    minCol = 0
  ElseIf minCol >= MapWidth Then
    Exit Sub
  End If
  
  maxCol = baseCol + maxLoop
  If maxCol < 0 Then
    Exit Sub
  ElseIf maxCol >= MapWidth Then
    maxCol = MapWidth - 1
  End If
  
  For loop1 = minRow To maxRow
    rS0 = (rowS - loop1) ^ 2
    rS1 = (rowS - (loop1 + 1)) ^ 2
    
    For loop2 = minCol To maxCol
      cS0 = (colS - loop2) ^ 2
      cS1 = (colS - (loop2 + 1)) ^ 2
      
      With MapArray(loop1, loop2)
        If .staticSurface <> 0 Then
          lightEffect = 1# - (Sqr(rS0 + cS0) / rangeS)  'vertex 1
          
          If lightEffect > 0 Then
            If lightEffect > 1# Then lightEffect = 1#
            
            With .cellVertices(1)
              rColour = (.Color \ 65536) + redC * lightEffect
              If rColour > 255 Then rColour = 255
              
              gColour = ((.Color And &HFF00&) \ 256) + greenC * lightEffect
              If gColour > 255 Then gColour = 255
              
              bColour = (.Color And &HFF&) + blueC * lightEffect
              If bColour > 255 Then bColour = 255
              
              .Color = rColour * 65536 + gColour * 256 + bColour
            End With
          End If
          
          lightEffect = 1# - (Sqr(rS1 + cS0) / rangeS)  'vertex 0
          
          If lightEffect > 0 Then
            If lightEffect > 1# Then lightEffect = 1#
            
            With .cellVertices(0)
              rColour = (.Color \ 65536) + redC * lightEffect
              If rColour > 255 Then rColour = 255
              
              gColour = ((.Color And &HFF00&) \ 256) + greenC * lightEffect
              If gColour > 255 Then gColour = 255
              
              bColour = (.Color And &HFF&) + blueC * lightEffect
              If bColour > 255 Then bColour = 255
              
              .Color = rColour * 65536 + gColour * 256 + bColour
            End With
          End If
          
          lightEffect = 1# - (Sqr(rS1 + cS1) / rangeS)  'vertex 2
          
          If lightEffect > 0 Then
            If lightEffect > 1# Then lightEffect = 1#
            
            With .cellVertices(2)
              rColour = (.Color \ 65536) + redC * lightEffect
              If rColour > 255 Then rColour = 255
              
              gColour = ((.Color And &HFF00&) \ 256) + greenC * lightEffect
              If gColour > 255 Then gColour = 255
              
              bColour = (.Color And &HFF&) + blueC * lightEffect
              If bColour > 255 Then bColour = 255
              
              .Color = rColour * 65536 + gColour * 256 + bColour
            End With
          End If
          
          lightEffect = 1# - (Sqr(rS0 + cS1) / rangeS)  'vertex 3
          
          If lightEffect > 0 Then
            If lightEffect > 1# Then lightEffect = 1#
            
            With .cellVertices(3)
              rColour = (.Color \ 65536) + redC * lightEffect
              If rColour > 255 Then rColour = 255
              
              gColour = ((.Color And &HFF00&) \ 256) + greenC * lightEffect
              If gColour > 255 Then gColour = 255
              
              bColour = (.Color And &HFF&) + blueC * lightEffect
              If bColour > 255 Then bColour = 255
              
              .Color = rColour * 65536 + gColour * 256 + bColour
            End With
          End If
        End If
      End With
    Next loop2
  Next loop1
End Sub

Public Sub Scale_MapAltitude(ByVal altitudeScale As Single)
  Dim loop1 As Long, loop2 As Long
  
  If altitudeScale < 0 Then Exit Sub
  
  For loop1 = 0 To MapHeight - 1
    For loop2 = 0 To MapWidth - 1
      With MapArray(loop1, loop2)
        .altitude = .altitude * altitudeScale
      End With
    Next loop2
  Next loop1
End Sub

Private Sub mapPrecalc()
  dParams.BasePosX = DisplayColumn_1000ths * worldBase
  dParams.BasePosY = DisplayRow_1000ths * worldBase
  
  dParams.displayOriginXCenter = DXDraw.m_ClippingRectangleWidth \ 2
  dParams.lastRowPos = DXDraw.m_ClippingRectangleY + DXDraw.m_ClippingRectangleHeight
  
  dParams.MapDisplayWidth = (dParams.displayOriginXCenter \ dParams.effWorldDoubleBase) + 3
  dParams.MapDisplayHeight = 0
  
  dParams.firstCol = DisplayColumn_1000ths \ 1000 - dParams.MapDisplayWidth - 1
  dParams.firstRow = DisplayRow_1000ths \ 1000 + dParams.MapDisplayWidth
  
  dParams.mapShiftCol = DisplayColumn_1000ths Mod 1000
  dParams.mapShiftRow = DisplayRow_1000ths Mod 1000
  
  If dParams.mapShiftCol < 0 Then
    dParams.firstCol = dParams.firstCol - 1
    dParams.mapShiftCol = dParams.mapShiftCol + 1000
  End If
  
  If dParams.mapShiftRow < 0 Then
    dParams.firstRow = dParams.firstRow - 1
    dParams.mapShiftRow = dParams.mapShiftRow + 1000
  End If
    
  dParams.displayOriginX = DXDraw.m_ClippingRectangleX + dParams.displayOriginXCenter - (dParams.MapDisplayWidth * dParams.effWorldDoubleBase + dParams.effWorldBase) - ((dParams.mapShiftCol - dParams.mapShiftRow) * dParams.effWorldBase) \ 1000
  dParams.displayOriginY = DXDraw.m_ClippingRectangleY + dParams.effWorldHalfBase - ((dParams.mapShiftCol + dParams.mapShiftRow) * dParams.effWorldHalfBase) \ 1000
  
  dParams.MapDisplayWidth = dParams.MapDisplayWidth * 2
End Sub

'Redraw the ISOmetric map layer
Public Sub Render_Map()
  Dim RowX As Long, RowY As Long, ImageXOffset As Long, ImageYOffset As Long
  Dim MapCellX As Long, MapCellY As Long, MapStartCellX As Long, MapStartCellY As Long
  Dim BaseRowPosY As Long, BaseRowPosX As Long
  Dim waitingDisplay As Boolean, didDisplay As Boolean, lastTexture As Long
  
  On Error Resume Next
  
  mapPrecalc
  
  lastTexture = -1
  
  MapStartCellY = dParams.firstRow
  MapStartCellX = dParams.firstCol
  
  RowY = 0
  BaseRowPosY = dParams.displayOriginY
  BaseRowPosX = dParams.displayOriginX
  
  waitingDisplay = True
  
  dx_Direct3DDevice.BeginScene
  
  Do
    didDisplay = False
    
    dParams.MapDisplayHeight = dParams.MapDisplayHeight + 1
    
    MapCellX = MapStartCellX
    MapCellY = MapStartCellY
    
    For RowX = 1 To dParams.MapDisplayWidth
      If MapCellY >= -5 And MapCellY < mapTHeight And MapCellX >= -5 And MapCellX < mapTWidth Then
        With MapArray(MapCellY, MapCellX)
          If .staticSurface <> lastTexture Then
            lastTexture = .staticSurface
            
            dx_Direct3DDevice.SetTexture 0, DXDraw.GetDirectDrawSurface(lastTexture)
          End If
          
          If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
            .cellVertices(2).sx = BaseRowPosX
            .cellVertices(2).sy = BaseRowPosY + .yOffset0
            
            .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
            .cellVertices(0).sy = BaseRowPosY + .yOffset1
            
            .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
            .cellVertices(3).sy = BaseRowPosY + .yOffset2
            
            .cellVertices(1).sx = BaseRowPosX
            .cellVertices(1).sy = BaseRowPosY + .yOffset3 'vertex order -> 0 Left, 1 Top, 2 Bottom, 3 Right
            
            dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
            
            didDisplay = True
          End If
        End With
      Else
        If lastTexture <> 0 Then
          lastTexture = 0
          
          dx_Direct3DDevice.SetTexture 0, Nothing
        End If
        
        If MapCellY < 0 Then
          If MapCellX < 0 Then
            With MapArray(-1, -1)
              If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          ElseIf MapCellX >= MapWidth Then
            With MapArray(-1, MapWidth)
               If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
               
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          Else
            With MapArray(-1, MapCellX)
              If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          End If
        ElseIf MapCellY >= MapHeight Then
          If MapCellX < 0 Then
            With MapArray(MapHeight, -1)
              If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          ElseIf MapCellX >= MapWidth Then
            With MapArray(MapHeight, MapWidth)
              If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          Else
            With MapArray(MapHeight, MapCellX)
              If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          End If
        ElseIf MapCellX < 0 Then
          With MapArray(MapCellY, -1)
            If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
              .cellVertices(2).sy = BaseRowPosY + .yOffset0
              .cellVertices(0).sy = BaseRowPosY + .yOffset1
              .cellVertices(3).sy = BaseRowPosY + .yOffset2
              .cellVertices(1).sy = BaseRowPosY + .yOffset3
              .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
              .cellVertices(1).sx = BaseRowPosX
              .cellVertices(2).sx = BaseRowPosX
              .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
              
              dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
              
              didDisplay = True
            End If
          End With
        ElseIf MapCellX >= MapWidth Then
          With MapArray(MapCellY, MapWidth)
            If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
              .cellVertices(2).sy = BaseRowPosY + .yOffset0
              .cellVertices(0).sy = BaseRowPosY + .yOffset1
              .cellVertices(3).sy = BaseRowPosY + .yOffset2
              .cellVertices(1).sy = BaseRowPosY + .yOffset3
              .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
              .cellVertices(1).sx = BaseRowPosX
              .cellVertices(2).sx = BaseRowPosX
              .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
              
              dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
              
              didDisplay = True
            End If
          End With
        End If
      End If
      
      BaseRowPosX = BaseRowPosX + dParams.effWorldDoubleBase
      MapCellX = MapCellX + 1
      MapCellY = MapCellY - 1
    Next RowX
    
    BaseRowPosY = BaseRowPosY + dParams.effWorldHalfBase
    
    If (dParams.MapDisplayHeight Mod 2) = 1 Then
      BaseRowPosX = dParams.displayOriginX + dParams.effWorldBase
      
      MapStartCellX = MapStartCellX + 1
    Else
      BaseRowPosX = dParams.displayOriginX
      
      MapStartCellY = MapStartCellY + 1
    End If
    
    RowY = RowY + 1
    
    If didDisplay Then
      waitingDisplay = False
    ElseIf waitingDisplay Then
      didDisplay = True
    End If
  Loop While didDisplay
  
  dx_Direct3DDevice.EndScene
End Sub

'Redraw the ISOmetric map layer with a grid
Public Sub Render_Map_with_Grid()
  Dim RowX As Long, RowY As Long, ImageXOffset As Long, ImageYOffset As Long
  Dim MapCellX As Long, MapCellY As Long, MapStartCellX As Long, MapStartCellY As Long
  Dim BaseRowPosY As Long, BaseRowPosX As Long
  Dim waitingDisplay As Boolean, didDisplay As Boolean
  
  Dim outBaseX As Long, outBaseY As Long
  Dim outBaseWidthX As Long, outBaseWidthY As Long, outBaseWidthXY As Long
  Dim outBaseHeightX As Long, outBaseHeightY As Long, outBaseHeightXY As Long
  
  Dim tempVertices(0 To 4) As D3DTLVERTEX, loop1 As Long
  
  On Error Resume Next
  
  mapPrecalc
  
  tempVertices(0).Color = gridColour 'faster than a loop
  tempVertices(1).Color = gridColour
  tempVertices(2).Color = gridColour
  tempVertices(3).Color = gridColour
  tempVertices(4).Color = gridColour
  
  MapStartCellY = dParams.firstRow
  MapStartCellX = dParams.firstCol
  
  RowY = 0
  BaseRowPosY = dParams.displayOriginY
  BaseRowPosX = dParams.displayOriginX
  
  waitingDisplay = True
  
  dx_Direct3DDevice.BeginScene
  
  Do
    didDisplay = False
    
    dParams.MapDisplayHeight = dParams.MapDisplayHeight + 1
    
    MapCellX = MapStartCellX
    MapCellY = MapStartCellY
    
    For RowX = 1 To dParams.MapDisplayWidth
      If MapCellY >= -5 And MapCellY < mapTHeight And MapCellX >= -5 And MapCellX < mapTWidth Then
        With MapArray(MapCellY, MapCellX)
          dx_Direct3DDevice.SetTexture 0, DXDraw.GetDirectDrawSurface(.staticSurface)
          
          If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
            .cellVertices(2).sx = BaseRowPosX
            .cellVertices(2).sy = BaseRowPosY + .yOffset0
            
            .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
            .cellVertices(0).sy = BaseRowPosY + .yOffset1
            
            .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
            .cellVertices(3).sy = BaseRowPosY + .yOffset2
            
            .cellVertices(1).sx = BaseRowPosX
            .cellVertices(1).sy = BaseRowPosY + .yOffset3 'vertex order -> 0 Left, 1 Top, 2 Bottom, 3 Right
          
            dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
            
            dx_Direct3DDevice.SetTexture 0, Nothing
            
            tempVertices(0).sx = .cellVertices(2).sx
            tempVertices(0).sy = .cellVertices(2).sy
            tempVertices(1).sx = .cellVertices(0).sx
            tempVertices(1).sy = .cellVertices(0).sy
            tempVertices(2).sx = .cellVertices(1).sx
            tempVertices(2).sy = .cellVertices(1).sy
            tempVertices(3).sx = .cellVertices(3).sx
            tempVertices(3).sy = .cellVertices(3).sy
            tempVertices(4).sx = .cellVertices(2).sx
            tempVertices(4).sy = .cellVertices(2).sy
            
            dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
            
            didDisplay = True
          End If
        End With
      Else
        dx_Direct3DDevice.SetTexture 0, Nothing
        
        If MapCellY < 0 Then
          If MapCellX < 0 Then
            With MapArray(-1, -1)
              If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                tempVertices(0).sx = .cellVertices(2).sx
                tempVertices(0).sy = .cellVertices(2).sy
                tempVertices(1).sx = .cellVertices(0).sx
                tempVertices(1).sy = .cellVertices(0).sy
                tempVertices(2).sx = .cellVertices(1).sx
                tempVertices(2).sy = .cellVertices(1).sy
                tempVertices(3).sx = .cellVertices(3).sx
                tempVertices(3).sy = .cellVertices(3).sy
                tempVertices(4).sx = .cellVertices(2).sx
                tempVertices(4).sy = .cellVertices(2).sy
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          ElseIf MapCellX >= MapWidth Then
            With MapArray(-1, MapWidth)
              If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                tempVertices(0).sx = .cellVertices(2).sx
                tempVertices(0).sy = .cellVertices(2).sy
                tempVertices(1).sx = .cellVertices(0).sx
                tempVertices(1).sy = .cellVertices(0).sy
                tempVertices(2).sx = .cellVertices(1).sx
                tempVertices(2).sy = .cellVertices(1).sy
                tempVertices(3).sx = .cellVertices(3).sx
                tempVertices(3).sy = .cellVertices(3).sy
                tempVertices(4).sx = .cellVertices(2).sx
                tempVertices(4).sy = .cellVertices(2).sy
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          Else
            With MapArray(-1, MapCellX)
              If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                tempVertices(0).sx = .cellVertices(2).sx
                tempVertices(0).sy = .cellVertices(2).sy
                tempVertices(1).sx = .cellVertices(0).sx
                tempVertices(1).sy = .cellVertices(0).sy
                tempVertices(2).sx = .cellVertices(1).sx
                tempVertices(2).sy = .cellVertices(1).sy
                tempVertices(3).sx = .cellVertices(3).sx
                tempVertices(3).sy = .cellVertices(3).sy
                tempVertices(4).sx = .cellVertices(2).sx
                tempVertices(4).sy = .cellVertices(2).sy
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          End If
        ElseIf MapCellY >= MapHeight Then
          If MapCellX < 0 Then
            With MapArray(MapHeight, -1)
              If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                tempVertices(0).sx = .cellVertices(2).sx
                tempVertices(0).sy = .cellVertices(2).sy
                tempVertices(1).sx = .cellVertices(0).sx
                tempVertices(1).sy = .cellVertices(0).sy
                tempVertices(2).sx = .cellVertices(1).sx
                tempVertices(2).sy = .cellVertices(1).sy
                tempVertices(3).sx = .cellVertices(3).sx
                tempVertices(3).sy = .cellVertices(3).sy
                tempVertices(4).sx = .cellVertices(2).sx
                tempVertices(4).sy = .cellVertices(2).sy
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          ElseIf MapCellX >= MapWidth Then
            With MapArray(MapHeight, MapWidth)
              If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                tempVertices(0).sx = .cellVertices(2).sx
                tempVertices(0).sy = .cellVertices(2).sy
                tempVertices(1).sx = .cellVertices(0).sx
                tempVertices(1).sy = .cellVertices(0).sy
                tempVertices(2).sx = .cellVertices(1).sx
                tempVertices(2).sy = .cellVertices(1).sy
                tempVertices(3).sx = .cellVertices(3).sx
                tempVertices(3).sy = .cellVertices(3).sy
                tempVertices(4).sx = .cellVertices(2).sx
                tempVertices(4).sy = .cellVertices(2).sy
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          Else
            With MapArray(MapHeight, MapCellX)
              If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
                .cellVertices(2).sy = BaseRowPosY + .yOffset0
                .cellVertices(0).sy = BaseRowPosY + .yOffset1
                .cellVertices(3).sy = BaseRowPosY + .yOffset2
                .cellVertices(1).sy = BaseRowPosY + .yOffset3
                .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
                .cellVertices(1).sx = BaseRowPosX
                .cellVertices(2).sx = BaseRowPosX
                .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
                
                tempVertices(0).sx = .cellVertices(2).sx
                tempVertices(0).sy = .cellVertices(2).sy
                tempVertices(1).sx = .cellVertices(0).sx
                tempVertices(1).sy = .cellVertices(0).sy
                tempVertices(2).sx = .cellVertices(1).sx
                tempVertices(2).sy = .cellVertices(1).sy
                tempVertices(3).sx = .cellVertices(3).sx
                tempVertices(3).sy = .cellVertices(3).sy
                tempVertices(4).sx = .cellVertices(2).sx
                tempVertices(4).sy = .cellVertices(2).sy
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          End If
        ElseIf MapCellX < 0 Then
          With MapArray(MapCellY, -1)
            If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
              .cellVertices(2).sy = BaseRowPosY + .yOffset0
              .cellVertices(0).sy = BaseRowPosY + .yOffset1
              .cellVertices(3).sy = BaseRowPosY + .yOffset2
              .cellVertices(1).sy = BaseRowPosY + .yOffset3
              .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
              .cellVertices(1).sx = BaseRowPosX
              .cellVertices(2).sx = BaseRowPosX
              .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
              
              dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
              
              tempVertices(0).sx = .cellVertices(2).sx
              tempVertices(0).sy = .cellVertices(2).sy
              tempVertices(1).sx = .cellVertices(0).sx
              tempVertices(1).sy = .cellVertices(0).sy
              tempVertices(2).sx = .cellVertices(1).sx
              tempVertices(2).sy = .cellVertices(1).sy
              tempVertices(3).sx = .cellVertices(3).sx
              tempVertices(3).sy = .cellVertices(3).sy
              tempVertices(4).sx = .cellVertices(2).sx
              tempVertices(4).sy = .cellVertices(2).sy
              
              dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
              
              didDisplay = True
            End If
          End With
        ElseIf MapCellX >= MapWidth Then
          With MapArray(MapCellY, MapWidth)
            If BaseRowPosY + .yOffsetPeak < dParams.lastRowPos Then
              .cellVertices(2).sy = BaseRowPosY + .yOffset0
              .cellVertices(0).sy = BaseRowPosY + .yOffset1
              .cellVertices(3).sy = BaseRowPosY + .yOffset2
              .cellVertices(1).sy = BaseRowPosY + .yOffset3
              .cellVertices(0).sx = BaseRowPosX - dParams.effWorldBase
              .cellVertices(1).sx = BaseRowPosX
              .cellVertices(2).sx = BaseRowPosX
              .cellVertices(3).sx = BaseRowPosX + dParams.effWorldBase
              
              dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, .cellVertices(0), 4, D3DDP_DEFAULT
              
              tempVertices(0).sx = .cellVertices(2).sx
              tempVertices(0).sy = .cellVertices(2).sy
              tempVertices(1).sx = .cellVertices(0).sx
              tempVertices(1).sy = .cellVertices(0).sy
              tempVertices(2).sx = .cellVertices(1).sx
              tempVertices(2).sy = .cellVertices(1).sy
              tempVertices(3).sx = .cellVertices(3).sx
              tempVertices(3).sy = .cellVertices(3).sy
              tempVertices(4).sx = .cellVertices(2).sx
              tempVertices(4).sy = .cellVertices(2).sy
              
              dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
              
              didDisplay = True
            End If
          End With
        End If
      End If
      
      BaseRowPosX = BaseRowPosX + dParams.effWorldDoubleBase
      MapCellX = MapCellX + 1
      MapCellY = MapCellY - 1
    Next RowX
    
    BaseRowPosY = BaseRowPosY + dParams.effWorldHalfBase
    
    If (dParams.MapDisplayHeight Mod 2) = 1 Then
      BaseRowPosX = dParams.displayOriginX + dParams.effWorldBase
      
      MapStartCellX = MapStartCellX + 1
    Else
      BaseRowPosX = dParams.displayOriginX
      
      MapStartCellY = MapStartCellY + 1
    End If
    
    RowY = RowY + 1
    
    If didDisplay Then
      waitingDisplay = False
    ElseIf waitingDisplay Then
      didDisplay = True
    End If
  Loop While didDisplay
  
  dx_Direct3DDevice.EndScene
End Sub

' Redraws the ISOmetric objects for the visible portion of the map
Public Sub Render_WorldObjects(Optional ByVal useFogOfWar As Boolean = False, Optional ByVal targetISOmetricObject As clsISO_Object = Nothing)
  Dim effX As Long, effY As Long, startRow As Long, startCol As Long, FOWstatus As Long
  Dim RowY  As Long, RowX As Long, chainList As clsISO_Object, MapCellX As Long, MapCellY As Long
  Dim imageBase As Long, temp1 As Long
  
  On Error Resume Next
  
  dx_Direct3DDevice.BeginScene
  
  startRow = dParams.firstRow
  startCol = dParams.firstCol
  
  If Not (targetISOmetricObject Is Nothing) Then targetISOmetricObject.wasDisplayed = 0
  
  Set chainList = isoObjectGroundFXChain ' display ground effects
    
  Do While Not (chainList Is Nothing)
    With chainList
      If .Visible Then
        .wasDisplayed = 3
        
        effX = .PosX_1000ths - dParams.BasePosX
        effY = .PosY_1000ths - dParams.BasePosY
      
        .displayPosX = dParams.displayOriginXCenter + ((effX - effY) * DXDraw.ISOmetricViewScale) \ 1000
        .displayPosY = (((effX + effY) \ 2 - .PosZ_1000ths) * DXDraw.ISOmetricViewScale) \ 1000
        
        .Render_Object
        
        If .CollisionBoxVisible Then .Draw_CollisionBox
      Else
        .wasDisplayed = 0
      End If
      
      Set chainList = .mapChainNext
    End With
  Loop
  
  If useFogOfWar Then
    For RowY = 1 To dParams.MapDisplayHeight
      MapCellY = startRow
      MapCellX = startCol
      
      For RowX = 1 To dParams.MapDisplayWidth
        If MapCellY >= -5 And MapCellY < mapTHeight And MapCellX >= -5 And MapCellX < mapTWidth Then
          With MapArray(MapCellY, MapCellX)
            Set chainList = .isoObjectChain
            
            If Not (chainList Is Nothing) Then
              FOWstatus = .fogOfWarStatus
              
              If .ghostWorldObjects Then
                .ghostWorldObjects = False
              
                Do
                  With chainList
                    If .Visible Then
                      If (FOWstatus = 1) Or (.visibleInFOW = True) Then
                        .wasDisplayed = FOWstatus + 2
                        
                        effX = .PosX_1000ths - dParams.BasePosX
                        effY = .PosY_1000ths - dParams.BasePosY
                        
                        .displayPosX = dParams.displayOriginXCenter + ((effX - effY) * DXDraw.ISOmetricViewScale) \ 1000
                        .displayPosY = (((effX + effY) \ 2 - .PosZ_1000ths) * DXDraw.ISOmetricViewScale) \ 1000
                        
                        .Render_ObjectGhosted
                        
                        If .CollisionBoxVisible Then .Draw_CollisionBox
                      Else
                        .wasDisplayed = 0
                      End If
                    Else
                      .wasDisplayed = 0
                    End If
                    
                    Set chainList = .mapChainNext
                  End With
                Loop While Not (chainList Is Nothing)
              Else
                Do
                  With chainList
                    If .Visible Then
                      If (FOWstatus = 1) Or (.visibleInFOW = True) Then
                        .wasDisplayed = FOWstatus + 2
                        
                        effX = .PosX_1000ths - dParams.BasePosX
                        effY = .PosY_1000ths - dParams.BasePosY
                        
                        .displayPosX = dParams.displayOriginXCenter + ((effX - effY) * DXDraw.ISOmetricViewScale) \ 1000
                        .displayPosY = (((effX + effY) \ 2 - .PosZ_1000ths) * DXDraw.ISOmetricViewScale) \ 1000
                        
                        .Render_Object
                        
                        If .CollisionBoxVisible Then .Draw_CollisionBox
                      Else
                        .wasDisplayed = 0
                      End If
                    Else
                      .wasDisplayed = 0
                    End If
                    
                    Set chainList = .mapChainNext
                  End With
                Loop While Not (chainList Is Nothing)
              End If
            End If
          End With
        End If
        
        MapCellX = MapCellX + 1
        MapCellY = MapCellY - 1
      Next RowX
      
      If (RowY Mod 2) = 1 Then
        startCol = startCol + 1
      Else
        startRow = startRow + 1
      End If
    Next RowY
    
    Set chainList = isoObjectChain ' display floater object drop shadows
    
    Do While Not (chainList Is Nothing)
      With chainList
        If .Visible Then
          MapCellX = .PosX_1000ths \ dParams.worldBase1000
          MapCellY = .PosY_1000ths \ dParams.worldBase1000
          
          If MapCellY >= -5 And MapCellY < mapTHeight And MapCellX >= -5 And MapCellX < mapTWidth Then
            FOWstatus = MapArray(MapCellY, MapCellX).fogOfWarStatus
          Else
            FOWstatus = -1
          End If
          
          If (FOWstatus = 1) Or (.visibleInFOW = True) Then
            .wasDisplayed = FOWstatus + 2
            
            effX = .PosX_1000ths - dParams.BasePosX
            effY = .PosY_1000ths - dParams.BasePosY
          
            .displayPosX = dParams.displayOriginXCenter + ((effX - effY) * DXDraw.ISOmetricViewScale) \ 1000
            .displayPosY = (((effX + effY) \ 2) * DXDraw.ISOmetricViewScale) \ 1000
            
            .Render_ObjectShadowed
            
            .displayPosY = (((effX + effY) \ 2 - .PosZ_1000ths) * DXDraw.ISOmetricViewScale) \ 1000
          Else
            .wasDisplayed = 0
          End If
        Else
          .wasDisplayed = 0
        End If
        
        Set chainList = .mapChainNext
      End With
    Loop
    
   Set chainList = isoObjectChain ' display floater objects
    
    Do While Not (chainList Is Nothing)
      With chainList
        If .wasDisplayed > 0 Then
          .Render_Object
          
          If .CollisionBoxVisible Then .Draw_CollisionBox
        End If
        
        Set chainList = .mapChainNext
      End With
    Loop
  Else 'no FOW
    For RowY = 1 To dParams.MapDisplayHeight
      MapCellY = startRow
      MapCellX = startCol
      
      For RowX = 1 To dParams.MapDisplayWidth
        If MapCellY >= -5 And MapCellY < mapTHeight And MapCellX >= -5 And MapCellX < mapTWidth Then
          With MapArray(MapCellY, MapCellX)
            Set chainList = .isoObjectChain
            
            If .ghostWorldObjects Then
              .ghostWorldObjects = False
              
              Do While Not (chainList Is Nothing)
                With chainList
                  If .Visible Then
                    .wasDisplayed = 3
                    
                    effX = .PosX_1000ths - dParams.BasePosX
                    effY = .PosY_1000ths - dParams.BasePosY
                    
                    .displayPosX = dParams.displayOriginXCenter + ((effX - effY) * DXDraw.ISOmetricViewScale) \ 1000
                    .displayPosY = (((effX + effY) \ 2 - .PosZ_1000ths) * DXDraw.ISOmetricViewScale) \ 1000
                    
                    .Render_ObjectGhosted
                      
                    If .CollisionBoxVisible Then .Draw_CollisionBox
                  Else
                    .wasDisplayed = 0
                  End If
                  
                  Set chainList = .mapChainNext
                End With
              Loop
            Else
              Do While Not (chainList Is Nothing)
                With chainList
                  If .Visible Then
                    .wasDisplayed = 3
                    
                    effX = .PosX_1000ths - dParams.BasePosX
                    effY = .PosY_1000ths - dParams.BasePosY
                    
                    .displayPosX = dParams.displayOriginXCenter + ((effX - effY) * DXDraw.ISOmetricViewScale) \ 1000
                    .displayPosY = (((effX + effY) \ 2 - .PosZ_1000ths) * DXDraw.ISOmetricViewScale) \ 1000
                    
                    .Render_Object
                      
                    If .CollisionBoxVisible Then .Draw_CollisionBox
                  Else
                    .wasDisplayed = 0
                  End If
                  
                  Set chainList = .mapChainNext
                End With
              Loop
            End If
          End With
        End If
        
        MapCellX = MapCellX + 1
        MapCellY = MapCellY - 1
      Next RowX
      
      If (RowY Mod 2) = 1 Then
        startCol = startCol + 1
      Else
        startRow = startRow + 1
      End If
    Next RowY
    
    Set chainList = isoObjectChain ' display floater object drop shadows
    
    Do While Not (chainList Is Nothing)
      With chainList
        If .Visible Then
          .wasDisplayed = 3
          
          effX = .PosX_1000ths - dParams.BasePosX
          effY = .PosY_1000ths - dParams.BasePosY
        
          .displayPosX = dParams.displayOriginXCenter + ((effX - effY) * DXDraw.ISOmetricViewScale) \ 1000
          .displayPosY = (((effX + effY) \ 2) * DXDraw.ISOmetricViewScale) \ 1000
          
          .Render_ObjectShadowed
          
          .displayPosY = (((effX + effY) \ 2 - .PosZ_1000ths) * DXDraw.ISOmetricViewScale) \ 1000
        Else
          .wasDisplayed = 0
        End If
        
        Set chainList = .mapChainNext
      End With
    Loop
    
   Set chainList = isoObjectChain ' display floater objects
    
    Do While Not (chainList Is Nothing)
      With chainList
        If .wasDisplayed > 0 Then
          .Render_Object
          
          If .CollisionBoxVisible Then .Draw_CollisionBox
        End If
        
        Set chainList = .mapChainNext
      End With
    Loop
  End If
  
  If Not (targetISOmetricObject Is Nothing) Then
    With targetISOmetricObject
      If .wasDisplayed > 0 Then
        .Render_ObjectGhosted
        
        If .CollisionBoxVisible Then
          temp1 = .CollisionBoxColour
          .CollisionBoxColour = .CollisionBoxColour And &H80FFFFFF
          
          .Draw_CollisionBox
          
          .CollisionBoxColour = temp1
        End If
      End If
    End With
  End If
  
  dx_Direct3DDevice.EndScene
End Sub

'Redraw the fog of war
Public Sub Render_FogOfWar()
  Dim RowX As Long, RowY As Long, rowLoop As Long
  Dim MapCellX As Long, MapCellY As Long
  Dim mapShiftX As Long, mapShiftY As Long, MapStartCellX As Long, MapStartCellY As Long
  Dim BaseRowPosY As Long, BaseRowPosX As Long
  Dim waitingDisplay As Boolean, didDisplay As Boolean
  Dim fogVertex(0 To 3) As D3DTLVERTEX
  
  On Error Resume Next
    
  fogVertex(0).rhw = 1
  fogVertex(1).rhw = 1
  fogVertex(2).rhw = 1
  fogVertex(3).rhw = 1
  
  MapStartCellX = dParams.firstCol
  MapStartCellY = dParams.firstRow
  
  RowY = 0
  BaseRowPosY = dParams.displayOriginY
  waitingDisplay = True
  
  dx_Direct3DDevice.BeginScene
  dx_Direct3DDevice.SetTexture 0, Nothing
  
  Do
    For rowLoop = 1 To 2
      RowX = 0
      MapCellY = MapStartCellY
      
      If rowLoop = 1 Then 'render odd display row
        BaseRowPosX = dParams.displayOriginX
        MapCellX = MapStartCellX
        
        didDisplay = False
      Else 'render even display row
        BaseRowPosX = dParams.displayOriginX + dParams.effWorldBase
        BaseRowPosY = BaseRowPosY + dParams.effWorldHalfBase
        MapCellX = MapStartCellX + 1
      End If
      
      Do While RowX < dParams.MapDisplayWidth
        If MapCellY >= -5 And MapCellY < mapTHeight And MapCellX >= -5 And MapCellX < mapTWidth Then
          With MapArray(MapCellY, MapCellX)
            fogVertex(1).sy = BaseRowPosY + .yOffset3 'vertex order -> 0 Left, 1 Top, 2 Bottom, 3 Right
            
            If fogVertex(1).sy < dParams.lastRowPos Then
              fogVertex(2).sx = BaseRowPosX
              fogVertex(2).sy = BaseRowPosY + .yOffset0
              
              fogVertex(0).sx = BaseRowPosX - dParams.effWorldBase
              fogVertex(0).sy = BaseRowPosY + .yOffset1
              
              fogVertex(3).sx = BaseRowPosX + dParams.effWorldBase
              fogVertex(3).sy = BaseRowPosY + .yOffset2
              
              fogVertex(1).sx = BaseRowPosX
              
              If .fogOfWarStatus = -1 Then
                fogVertex(0).Color = &HFF000000
                fogVertex(1).Color = &HFF000000
                fogVertex(2).Color = &HFF000000
                fogVertex(3).Color = &HFF000000
              Else
                Select Case .fogOfWar(0)
                  Case -1
                    fogVertex(0).Color = &HFF000000
                  Case 0
                    fogVertex(0).Color = &HC0202428
                  Case Else
                    fogVertex(0).Color = 0
                End Select
                
                Select Case .fogOfWar(1)
                  Case -1
                    fogVertex(1).Color = &HFF000000
                  Case 0
                    fogVertex(1).Color = &HC0202428
                  Case Else
                    fogVertex(1).Color = 0
                End Select
                
                Select Case .fogOfWar(2)
                  Case -1
                    fogVertex(2).Color = &HFF000000
                  Case 0
                    fogVertex(2).Color = &HC0202428
                  Case Else
                    fogVertex(2).Color = 0
                End Select
                
                Select Case .fogOfWar(3)
                  Case -1
                    fogVertex(3).Color = &HFF000000
                  Case 0
                    fogVertex(3).Color = &HC0202428
                  Case Else
                    fogVertex(3).Color = 0
                End Select
              End If
              
              dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, fogVertex(0), 4, D3DDP_DEFAULT
            
              didDisplay = True
            End If
          End With
        Else
          fogVertex(0).Color = &HFF000000
          fogVertex(1).Color = &HFF000000
          fogVertex(2).Color = &HFF000000
          fogVertex(3).Color = &HFF000000
          
          If MapCellY < 0 Then
            If MapCellX < 0 Then
              With MapArray(-1, -1)
                fogVertex(1).sy = BaseRowPosY + .yOffset3
                
                If fogVertex(1).sy < dParams.lastRowPos Then
                  fogVertex(2).sy = BaseRowPosY + .yOffset0
                  fogVertex(0).sy = BaseRowPosY + .yOffset1
                  fogVertex(3).sy = BaseRowPosY + .yOffset2
                  fogVertex(0).sx = BaseRowPosX - dParams.effWorldBase
                  fogVertex(1).sx = BaseRowPosX
                  fogVertex(2).sx = BaseRowPosX
                  fogVertex(3).sx = BaseRowPosX + dParams.effWorldBase
                  
                  dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, fogVertex(0), 4, D3DDP_DEFAULT
                  
                  didDisplay = True
                End If
              End With
            ElseIf MapCellX >= MapWidth Then
              With MapArray(-1, MapWidth)
                fogVertex(1).sy = BaseRowPosY + .yOffset3
                
                If fogVertex(1).sy < dParams.lastRowPos Then
                  fogVertex(2).sy = BaseRowPosY + .yOffset0
                  fogVertex(0).sy = BaseRowPosY + .yOffset1
                  fogVertex(3).sy = BaseRowPosY + .yOffset2
                  fogVertex(0).sx = BaseRowPosX - dParams.effWorldBase
                  fogVertex(1).sx = BaseRowPosX
                  fogVertex(2).sx = BaseRowPosX
                  fogVertex(3).sx = BaseRowPosX + dParams.effWorldBase
                  
                  dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, fogVertex(0), 4, D3DDP_DEFAULT
                  
                  didDisplay = True
                End If
              End With
            Else
              With MapArray(-1, MapCellX)
                fogVertex(1).sy = BaseRowPosY + .yOffset3
                
                If fogVertex(1).sy < dParams.lastRowPos Then
                  fogVertex(2).sy = BaseRowPosY + .yOffset0
                  fogVertex(0).sy = BaseRowPosY + .yOffset1
                  fogVertex(3).sy = BaseRowPosY + .yOffset2
                  fogVertex(0).sx = BaseRowPosX - dParams.effWorldBase
                  fogVertex(1).sx = BaseRowPosX
                  fogVertex(2).sx = BaseRowPosX
                  fogVertex(3).sx = BaseRowPosX + dParams.effWorldBase
                  
                  dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, fogVertex(0), 4, D3DDP_DEFAULT
                  
                  didDisplay = True
                End If
              End With
            End If
          ElseIf MapCellY >= MapHeight Then
            If MapCellX < 0 Then
              With MapArray(MapHeight, -1)
                fogVertex(1).sy = BaseRowPosY + .yOffset3
                
                If fogVertex(1).sy < dParams.lastRowPos Then
                  fogVertex(2).sy = BaseRowPosY + .yOffset0
                  fogVertex(0).sy = BaseRowPosY + .yOffset1
                  fogVertex(3).sy = BaseRowPosY + .yOffset2
                  fogVertex(0).sx = BaseRowPosX - dParams.effWorldBase
                  fogVertex(1).sx = BaseRowPosX
                  fogVertex(2).sx = BaseRowPosX
                  fogVertex(3).sx = BaseRowPosX + dParams.effWorldBase
                  
                  dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, fogVertex(0), 4, D3DDP_DEFAULT
                  
                  didDisplay = True
                End If
              End With
            ElseIf MapCellX >= MapWidth Then
              With MapArray(MapHeight, MapWidth)
                fogVertex(1).sy = BaseRowPosY + .yOffset3
                
                If fogVertex(1).sy < dParams.lastRowPos Then
                  fogVertex(2).sy = BaseRowPosY + .yOffset0
                  fogVertex(0).sy = BaseRowPosY + .yOffset1
                  fogVertex(3).sy = BaseRowPosY + .yOffset2
                  fogVertex(0).sx = BaseRowPosX - dParams.effWorldBase
                  fogVertex(1).sx = BaseRowPosX
                  fogVertex(2).sx = BaseRowPosX
                  fogVertex(3).sx = BaseRowPosX + dParams.effWorldBase
                  
                  dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, fogVertex(0), 4, D3DDP_DEFAULT
                  
                  didDisplay = True
                End If
              End With
            Else
              With MapArray(MapHeight, MapCellX)
                fogVertex(1).sy = BaseRowPosY + .yOffset3
                
                If fogVertex(1).sy < dParams.lastRowPos Then
                  fogVertex(2).sy = BaseRowPosY + .yOffset0
                  fogVertex(0).sy = BaseRowPosY + .yOffset1
                  fogVertex(3).sy = BaseRowPosY + .yOffset2
                  fogVertex(0).sx = BaseRowPosX - dParams.effWorldBase
                  fogVertex(1).sx = BaseRowPosX
                  fogVertex(2).sx = BaseRowPosX
                  fogVertex(3).sx = BaseRowPosX + dParams.effWorldBase
                  
                  dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, fogVertex(0), 4, D3DDP_DEFAULT
                  
                  didDisplay = True
                End If
              End With
            End If
          ElseIf MapCellX < 0 Then
            With MapArray(MapCellY, -1)
              fogVertex(1).sy = BaseRowPosY + .yOffset3
              
              If fogVertex(1).sy < dParams.lastRowPos Then
                fogVertex(2).sy = BaseRowPosY + .yOffset0
                fogVertex(0).sy = BaseRowPosY + .yOffset1
                fogVertex(3).sy = BaseRowPosY + .yOffset2
                fogVertex(0).sx = BaseRowPosX - dParams.effWorldBase
                fogVertex(1).sx = BaseRowPosX
                fogVertex(2).sx = BaseRowPosX
                fogVertex(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, fogVertex(0), 4, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          ElseIf MapCellX >= MapWidth Then
            With MapArray(MapCellY, MapWidth)
              fogVertex(1).sy = BaseRowPosY + .yOffset3
              
              If fogVertex(1).sy < dParams.lastRowPos Then
                fogVertex(2).sy = BaseRowPosY + .yOffset0
                fogVertex(0).sy = BaseRowPosY + .yOffset1
                fogVertex(3).sy = BaseRowPosY + .yOffset2
                fogVertex(0).sx = BaseRowPosX - dParams.effWorldBase
                fogVertex(1).sx = BaseRowPosX
                fogVertex(2).sx = BaseRowPosX
                fogVertex(3).sx = BaseRowPosX + dParams.effWorldBase
                
                dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, fogVertex(0), 4, D3DDP_DEFAULT
                
                didDisplay = True
              End If
            End With
          End If
        End If
        
        RowX = RowX + 1
        BaseRowPosX = BaseRowPosX + dParams.effWorldDoubleBase
        MapCellX = MapCellX + 1
        MapCellY = MapCellY - 1
      Loop
    Next rowLoop
    
    RowY = RowY + 1
    BaseRowPosY = BaseRowPosY + dParams.effWorldHalfBase
    
    MapStartCellX = MapStartCellX + 1
    MapStartCellY = MapStartCellY + 1
    
    If didDisplay Then
      waitingDisplay = False
    ElseIf waitingDisplay Then
      didDisplay = True
    End If
  Loop While didDisplay
  
  dx_Direct3DDevice.EndScene
End Sub

Public Sub Draw_TerrainCursor(ByVal row As Long, ByVal column As Long, ByVal cursorColour As Long, Optional ByVal cursorIs3D As Boolean = False, Optional ByVal showShadowCursor As Boolean = False)
  Dim displayBaseX As Long, displayBaseY As Long
  Dim baseT As Long, baseL As Long, baseR As Long, baseB As Long
  Dim effX As Long, effY As Long
  
  On Error Resume Next
  
  If row >= 0 And row < MapHeight Then
    If column >= 0 And column < MapWidth Then
      dx_Direct3DDevice.BeginScene
      dx_Direct3DDevice.SetTexture 0, Nothing
      
      effX = (column * 1000 - DisplayColumn_1000ths) * dParams.effWorldBase
      effY = (row * 1000 - DisplayRow_1000ths) * dParams.effWorldBase
      
      displayBaseX = DXDraw.m_ClippingRectangleX + DXDraw.m_ClippingRectangleWidth \ 2 + (effX - effY) \ 1000
      displayBaseY = DXDraw.m_ClippingRectangleY + (effX + effY) \ 2000
      
      With MapArray(row, column)
        baseT = (.altitudeT_1000ths * DXDraw.ISOmetricViewScale) \ 1000
        baseR = (.altitudeR_1000ths * DXDraw.ISOmetricViewScale) \ 1000
        baseL = (.altitudeL_1000ths * DXDraw.ISOmetricViewScale) \ 1000
        baseB = (.altitudeB_1000ths * DXDraw.ISOmetricViewScale) \ 1000
        
        If cursorIs3D Then
          tempVertices(0).Color = &H80000000
          tempVertices(0).sx = displayBaseX
          tempVertices(0).sy = displayBaseY
          
          tempVertices(1).Color = &H80000000
          tempVertices(1).sx = displayBaseX - dParams.effWorldBase
          tempVertices(1).sy = displayBaseY + dParams.effWorldHalfBase
          
          tempVertices(2).Color = &H80000000
          tempVertices(2).sx = displayBaseX
          tempVertices(2).sy = displayBaseY + dParams.effWorldBase
          
          tempVertices(3).Color = &H80000000
          tempVertices(3).sx = displayBaseX + dParams.effWorldBase
          tempVertices(3).sy = displayBaseY + dParams.effWorldHalfBase
          
          tempVertices(4).Color = &H80000000
          tempVertices(4).sx = displayBaseX
          tempVertices(4).sy = displayBaseY
            
          dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
          
          tempVertices(0).Color = cursorColour
          tempVertices(0).sy = displayBaseY - baseT
          
          tempVertices(1).sx = displayBaseX
          tempVertices(1).sy = displayBaseY
          
          tempVertices(2).Color = cursorColour
          tempVertices(2).sx = displayBaseX - dParams.effWorldBase
          tempVertices(2).sy = displayBaseY - baseL + dParams.effWorldHalfBase
          
          tempVertices(3).sx = displayBaseX - dParams.effWorldBase
          tempVertices(3).sy = displayBaseY + dParams.effWorldHalfBase
          
          tempVertices(4).Color = cursorColour
          tempVertices(4).sx = displayBaseX + dParams.effWorldBase
          tempVertices(4).sy = displayBaseY - baseR + dParams.effWorldHalfBase
          
          tempVertices(5).Color = &H80000000
          tempVertices(5).sx = displayBaseX + dParams.effWorldBase
          tempVertices(5).sy = displayBaseY + dParams.effWorldHalfBase
          
          tempVertices(6).Color = cursorColour
          tempVertices(6).sx = displayBaseX
          tempVertices(6).sy = displayBaseY - baseB + dParams.effWorldBase
          
          tempVertices(7).Color = &H80000000
          tempVertices(7).sx = displayBaseX
          tempVertices(7).sy = displayBaseY + dParams.effWorldBase
          
          dx_Direct3DDevice.DrawPrimitive D3DPT_LINELIST, D3DFVF_TLVERTEX, tempVertices(0), 8, D3DDP_DEFAULT
        End If
        
        If showShadowCursor Then
          tempVertices(1).Color = &H80000000
          tempVertices(1).sx = displayBaseX
          tempVertices(1).sy = displayBaseY - baseT
          
          tempVertices(0).Color = &H80000000
          tempVertices(0).sx = displayBaseX - dParams.effWorldBase
          tempVertices(0).sy = displayBaseY + dParams.effWorldHalfBase - baseL
          
          tempVertices(2).Color = &H80000000
          tempVertices(2).sx = displayBaseX
          tempVertices(2).sy = displayBaseY + dParams.effWorldBase - baseB
          
          tempVertices(3).Color = &H80000000
          tempVertices(3).sx = displayBaseX + dParams.effWorldBase
          tempVertices(3).sy = displayBaseY + dParams.effWorldHalfBase - baseR
          
          dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 4, D3DDP_DEFAULT
        End If
        
        tempVertices(0).Color = cursorColour
        tempVertices(0).sx = displayBaseX
        tempVertices(0).sy = displayBaseY - baseT
        
        tempVertices(1).Color = cursorColour
        tempVertices(1).sx = displayBaseX - dParams.effWorldBase
        tempVertices(1).sy = displayBaseY + dParams.effWorldHalfBase - baseL
        
        tempVertices(2).Color = cursorColour
        tempVertices(2).sx = displayBaseX
        tempVertices(2).sy = displayBaseY + dParams.effWorldBase - baseB
        
        tempVertices(3).Color = cursorColour
        tempVertices(3).sx = displayBaseX + dParams.effWorldBase
        tempVertices(3).sy = displayBaseY + dParams.effWorldHalfBase - baseR
        
        tempVertices(4).Color = cursorColour
        tempVertices(4).sx = displayBaseX
        tempVertices(4).sy = displayBaseY - baseT
        
        dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
      End With
      
      dx_Direct3DDevice.EndScene
    End If
  End If
End Sub

Public Sub Draw_FloatingCursor(ByVal row As Long, ByVal column As Long, ByVal cursorColour As Long, ByVal zBase As Long, ByVal zHeight As Long, Optional ByVal showShadowCursor As Boolean = False)
  Dim displayBaseX As Long, displayBaseY As Long
  Dim baseT As Long, baseL As Long, baseR As Long, baseB As Long
  Dim effX As Long, effY As Long
  
  On Error Resume Next
  
  If row >= 0 And row < MapHeight Then
    If column >= 0 And column < MapWidth Then
      dx_Direct3DDevice.BeginScene
      dx_Direct3DDevice.SetTexture 0, Nothing
      
      effX = (column * 1000 - DisplayColumn_1000ths) * dParams.effWorldBase
      effY = (row * 1000 - DisplayRow_1000ths) * dParams.effWorldBase
      
      displayBaseX = DXDraw.m_ClippingRectangleX + DXDraw.m_ClippingRectangleWidth \ 2 + (effX - effY) \ 1000
      displayBaseY = DXDraw.m_ClippingRectangleY + (effX + effY) \ 2000
      
      zBase = zBase * DXDraw.ISOmetricViewScale
      zHeight = zHeight * DXDraw.ISOmetricViewScale + zBase
      
      If showShadowCursor Then
        With MapArray(row, column)
          baseT = (.altitudeT_1000ths * DXDraw.ISOmetricViewScale) \ 1000
          baseR = (.altitudeR_1000ths * DXDraw.ISOmetricViewScale) \ 1000
          baseL = (.altitudeL_1000ths * DXDraw.ISOmetricViewScale) \ 1000
          baseB = (.altitudeB_1000ths * DXDraw.ISOmetricViewScale) \ 1000
          
          tempVertices(1).Color = &H80000000
          tempVertices(1).sx = displayBaseX
          tempVertices(1).sy = displayBaseY - baseT
          
          tempVertices(0).Color = &H80000000
          tempVertices(0).sx = displayBaseX - dParams.effWorldBase
          tempVertices(0).sy = displayBaseY + dParams.effWorldHalfBase - baseL
          
          tempVertices(2).Color = &H80000000
          tempVertices(2).sx = displayBaseX
          tempVertices(2).sy = displayBaseY + dParams.effWorldBase - baseB
          
          tempVertices(3).Color = &H80000000
          tempVertices(3).sx = displayBaseX + dParams.effWorldBase
          tempVertices(3).sy = displayBaseY + dParams.effWorldHalfBase - baseR
          
          dx_Direct3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 4, D3DDP_DEFAULT
          
          tempVertices(0).Color = cursorColour
          tempVertices(0).sx = displayBaseX
          tempVertices(0).sy = displayBaseY - baseT
          
          tempVertices(1).Color = cursorColour
          tempVertices(1).sx = displayBaseX - dParams.effWorldBase
          tempVertices(1).sy = displayBaseY + dParams.effWorldHalfBase - baseL
          
          tempVertices(2).Color = cursorColour
          tempVertices(2).sx = displayBaseX
          tempVertices(2).sy = displayBaseY + dParams.effWorldBase - baseB
          
          tempVertices(3).Color = cursorColour
          tempVertices(3).sx = displayBaseX + dParams.effWorldBase
          tempVertices(3).sy = displayBaseY + dParams.effWorldHalfBase - baseR
          
          tempVertices(4).Color = cursorColour
          tempVertices(4).sx = displayBaseX
          tempVertices(4).sy = tempVertices(0).sy
          
          dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
        End With
      End If
      
      tempVertices(0).Color = cursorColour
      tempVertices(0).sx = displayBaseX
      tempVertices(0).sy = displayBaseY - zBase
            
      tempVertices(1).Color = cursorColour
      tempVertices(1).sx = displayBaseX - dParams.effWorldBase
      tempVertices(1).sy = displayBaseY - zBase + dParams.effWorldHalfBase
      
      tempVertices(2).Color = cursorColour
      tempVertices(2).sx = displayBaseX
      tempVertices(2).sy = displayBaseY - zBase + dParams.effWorldBase
      
      tempVertices(3).Color = cursorColour
      tempVertices(3).sx = displayBaseX + dParams.effWorldBase
      tempVertices(3).sy = displayBaseY - zBase + dParams.effWorldHalfBase
      
      tempVertices(4).Color = cursorColour
      tempVertices(4).sx = displayBaseX
      tempVertices(4).sy = displayBaseY - zBase
      
      dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
      
      tempVertices(0).sx = displayBaseX
      tempVertices(0).sy = displayBaseY - zBase
      
      tempVertices(1).sx = displayBaseX
      tempVertices(1).sy = displayBaseY - zHeight
      
      tempVertices(2).sx = displayBaseX - dParams.effWorldBase
      tempVertices(2).sy = displayBaseY - zBase + dParams.effWorldHalfBase
      
      tempVertices(3).sx = displayBaseX - dParams.effWorldBase
      tempVertices(3).sy = displayBaseY - zHeight + dParams.effWorldHalfBase
       
      tempVertices(4).sx = displayBaseX + dParams.effWorldBase
      tempVertices(4).sy = displayBaseY - zBase + dParams.effWorldHalfBase
      
      tempVertices(5).Color = cursorColour
      tempVertices(5).sx = displayBaseX + dParams.effWorldBase
      tempVertices(5).sy = displayBaseY - zHeight + dParams.effWorldHalfBase
      
      tempVertices(6).Color = cursorColour
      tempVertices(6).sx = displayBaseX
      tempVertices(6).sy = displayBaseY - zBase + dParams.effWorldBase
      
      tempVertices(7).Color = cursorColour
      tempVertices(7).sx = displayBaseX
      tempVertices(7).sy = displayBaseY - zHeight + dParams.effWorldBase
      
      dx_Direct3DDevice.DrawPrimitive D3DPT_LINELIST, D3DFVF_TLVERTEX, tempVertices(0), 8, D3DDP_DEFAULT
      
      tempVertices(0).Color = cursorColour
      tempVertices(0).sx = displayBaseX
      tempVertices(0).sy = displayBaseY - zHeight
            
      tempVertices(1).Color = cursorColour
      tempVertices(1).sx = displayBaseX - dParams.effWorldBase
      tempVertices(1).sy = displayBaseY - zHeight + dParams.effWorldHalfBase
      
      tempVertices(2).Color = cursorColour
      tempVertices(2).sx = displayBaseX
      tempVertices(2).sy = displayBaseY - zHeight + dParams.effWorldBase
      
      tempVertices(3).Color = cursorColour
      tempVertices(3).sx = displayBaseX + dParams.effWorldBase
      tempVertices(3).sy = displayBaseY - zHeight + dParams.effWorldHalfBase
      
      tempVertices(4).Color = cursorColour
      tempVertices(4).sx = displayBaseX
      tempVertices(4).sy = displayBaseY - zHeight
      
      dx_Direct3DDevice.DrawPrimitive D3DPT_LINESTRIP, D3DFVF_TLVERTEX, tempVertices(0), 5, D3DDP_DEFAULT
      
      dx_Direct3DDevice.EndScene
    End If
  End If
End Sub

Public Sub Ensure_Visible_WorldXYZ(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long, ByVal worldPosZ_1000ths As Long, Optional ByVal forceCenter As Boolean = False)
  Dim displayX As Long, displayY As Long, halfHeight As Long
  Dim effX As Long, effY As Long
  
  If forceCenter = False Then
    effX = worldPosX_1000ths - DisplayColumn_1000ths * worldBase
    effY = worldPosY_1000ths - DisplayRow_1000ths * worldBase
   
    displayX = DXDraw.m_ClippingRectangleWidth \ 2 + ((effX - effY) * DXDraw.ISOmetricViewScale) \ 1000
    displayY = (((effX + effY) \ 2 - worldPosZ_1000ths) * DXDraw.ISOmetricViewScale) \ 1000
    
    If displayX < 0 Or displayY < 0 Then
      forceCenter = True
    ElseIf displayX >= DXDraw.m_ClippingRectangleWidth Or displayY >= DXDraw.m_ClippingRectangleHeight Then
      forceCenter = True
    End If
  End If
  
  If forceCenter Then
    halfHeight = ((DXDraw.m_ClippingRectangleHeight * 500) \ DXDraw.ISOmetricViewScale) + worldPosZ_1000ths
    
    DisplayRow_1000ths = (worldPosY_1000ths - halfHeight) \ worldBase
    DisplayColumn_1000ths = (worldPosX_1000ths - halfHeight) \ worldBase
  End If
End Sub

Public Sub Ensure_Visible_MapRC(ByVal row As Long, ByVal column As Long, Optional ByVal forceCenter As Boolean = False)
  Dim worldPosX_1000ths As Long, worldPosY_1000ths As Long, worldPosZ_1000ths As Long
  Dim displayX As Long, displayY As Long, halfHeight As Long
  Dim effX As Long, effY As Long
  
  worldPosX_1000ths = (column * worldBase + worldBase \ 2) * 1000
  worldPosY_1000ths = (row * worldBase + worldBase \ 2) * 1000
  worldPosZ_1000ths = MapArray(row, column).altitudeC_1000ths
    
  If forceCenter = False Then
    effX = worldPosX_1000ths - DisplayColumn_1000ths * worldBase
    effY = worldPosY_1000ths - DisplayRow_1000ths * worldBase
    
    displayX = DXDraw.m_ClippingRectangleWidth \ 2 + ((effX - effY) * DXDraw.ISOmetricViewScale) \ 1000
    displayY = (((effX + effY) \ 2 - worldPosZ_1000ths) * DXDraw.ISOmetricViewScale) \ 1000
  
    If displayX < 0 Or displayY < 0 Then
      forceCenter = True
    ElseIf displayX >= DXDraw.m_ClippingRectangleWidth Or displayY >= DXDraw.m_ClippingRectangleHeight Then
      forceCenter = True
    End If
  End If
  
  If forceCenter Then
    halfHeight = ((DXDraw.m_ClippingRectangleHeight * 500) \ DXDraw.ISOmetricViewScale) + worldPosZ_1000ths
    
    DisplayRow_1000ths = (worldPosY_1000ths - halfHeight) \ worldBase
    DisplayColumn_1000ths = (worldPosX_1000ths - halfHeight) \ worldBase
  End If
End Sub

Public Sub Save_ToFile(filePathName As String)
  Dim fileHandle As Long, versionString As String
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Output As #fileHandle
    Write #fileHandle, "Logicon Enterprises clsISO_MapObject Description Ver2.6"
    
    Save_ToOpenFile fileHandle
  Close fileHandle
  
  Exit Sub
  
badFile:
  MsgBox "Error saving map object -> " & filePathName, vbCritical, "ERROR"
End Sub

Public Sub Calculate_Map()
  Dim loop1 As Long, loop2 As Long
  Dim bValue As Long, bBase As Single, bRange As Single
  
  dParams.effWorldBase = worldBase * DXDraw.ISOmetricViewScale
  dParams.effWorldHalfBase = dParams.effWorldBase \ 2
  dParams.effWorldDoubleBase = dParams.effWorldBase * 2
  dParams.worldBase1000 = worldBase * 1000
  
  'set outbounds altitudes
  For loop1 = -5 To mapTWidth - 1
    If loop1 < 0 Then
      With MapArray(0, 0)
        MapArray(-5, loop1).altitude = .altitude
        MapArray(-4, loop1).altitude = .altitude
        MapArray(-3, loop1).altitude = .altitude
        MapArray(-2, loop1).altitude = .altitude
        MapArray(-1, loop1).altitude = .altitude
      End With
      
      With MapArray(MapHeight - 1, 0)
        MapArray(MapHeight, loop1).altitude = .altitude
        MapArray(MapHeight + 1, loop1).altitude = .altitude
        MapArray(MapHeight + 2, loop1).altitude = .altitude
        MapArray(MapHeight + 3, loop1).altitude = .altitude
        MapArray(MapHeight + 4, loop1).altitude = .altitude
      End With
    ElseIf loop1 >= MapWidth Then
      With MapArray(0, MapWidth - 1)
        MapArray(-5, loop1).altitude = .altitude
        MapArray(-4, loop1).altitude = .altitude
        MapArray(-3, loop1).altitude = .altitude
        MapArray(-2, loop1).altitude = .altitude
        MapArray(-1, loop1).altitude = .altitude
      End With
      
      With MapArray(MapHeight - 1, MapWidth - 1)
        MapArray(MapHeight, loop1).altitude = .altitude
        MapArray(MapHeight + 1, loop1).altitude = .altitude
        MapArray(MapHeight + 2, loop1).altitude = .altitude
        MapArray(MapHeight + 3, loop1).altitude = .altitude
        MapArray(MapHeight + 4, loop1).altitude = .altitude
      End With
    Else
      With MapArray(0, loop1)
        MapArray(-5, loop1).altitude = .altitude
        MapArray(-4, loop1).altitude = .altitude
        MapArray(-3, loop1).altitude = .altitude
        MapArray(-2, loop1).altitude = .altitude
        MapArray(-1, loop1).altitude = .altitude
      End With
      
      With MapArray(MapHeight - 1, loop1)
        MapArray(MapHeight, loop1).altitude = .altitude
        MapArray(MapHeight + 1, loop1).altitude = .altitude
        MapArray(MapHeight + 2, loop1).altitude = .altitude
        MapArray(MapHeight + 3, loop1).altitude = .altitude
        MapArray(MapHeight + 4, loop1).altitude = .altitude
      End With
    End If
  Next loop1
  
  For loop1 = -5 To mapTHeight - 1
    If loop1 < 0 Then
      With MapArray(0, 0)
        MapArray(loop1, -5).altitude = .altitude
        MapArray(loop1, -4).altitude = .altitude
        MapArray(loop1, -3).altitude = .altitude
        MapArray(loop1, -2).altitude = .altitude
        MapArray(loop1, -1).altitude = .altitude
      End With
      
      With MapArray(0, MapWidth - 1)
        MapArray(loop1, MapWidth).altitude = .altitude
        MapArray(loop1, MapWidth + 1).altitude = .altitude
        MapArray(loop1, MapWidth + 2).altitude = .altitude
        MapArray(loop1, MapWidth + 3).altitude = .altitude
        MapArray(loop1, MapWidth + 4).altitude = .altitude
      End With
    ElseIf loop1 >= MapWidth Then
      With MapArray(MapHeight - 1, 0)
        MapArray(loop1, -5).altitude = .altitude
        MapArray(loop1, -4).altitude = .altitude
        MapArray(loop1, -3).altitude = .altitude
        MapArray(loop1, -2).altitude = .altitude
        MapArray(loop1, -1).altitude = .altitude
      End With
      
      With MapArray(MapHeight - 1, MapWidth - 1)
        MapArray(loop1, MapWidth).altitude = .altitude
        MapArray(loop1, MapWidth + 1).altitude = .altitude
        MapArray(loop1, MapWidth + 2).altitude = .altitude
        MapArray(loop1, MapWidth + 3).altitude = .altitude
        MapArray(loop1, MapWidth + 4).altitude = .altitude
      End With
    Else
      With MapArray(loop1, 0)
        MapArray(loop1, -5).altitude = .altitude
        MapArray(loop1, -4).altitude = .altitude
        MapArray(loop1, -3).altitude = .altitude
        MapArray(loop1, -2).altitude = .altitude
        MapArray(loop1, -1).altitude = .altitude
      End With
      
      With MapArray(loop1, MapWidth - 1)
        MapArray(loop1, MapWidth).altitude = .altitude
        MapArray(loop1, MapWidth + 1).altitude = .altitude
        MapArray(loop1, MapWidth + 2).altitude = .altitude
        MapArray(loop1, MapWidth + 3).altitude = .altitude
        MapArray(loop1, MapWidth + 4).altitude = .altitude
      End With
    End If
  Next loop1
  
  bBase = (maxBrightness + minBrightness) / 2
  bRange = (maxBrightness - minBrightness) / worldBase
  
  'calculate map cell vertice offsets and colours
  For loop1 = 0 To MapHeight - 1
    For loop2 = 0 To MapWidth - 1
      With MapArray(loop1, loop2) 'vertex order -> 0 Left, 1 Top, 2 Bottom, 3 Right
        .yOffset0 = -(.altitude * DXDraw.ISOmetricViewScale)
        .altitudeB_1000ths = .altitude * 1000
        .yOffset1 = -(MapArray(loop1, loop2 - 1).altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
        .altitudeL_1000ths = MapArray(loop1, loop2 - 1).altitude * 1000
        
        If loop2 = 0 Then
          .cellVertices(0).Color = 255# * bBase
        Else
          bValue = 255# * (bBase + bRange * (MapArray(loop1, loop2 - 1).altitude - .altitude))
          
          If bValue < 0 Then
            bValue = 0
          ElseIf bValue > 255 Then
            bValue = 255
          End If
          
          .cellVertices(0).Color = bValue
        End If
        
        If loop2 = MapWidth - 1 Then
          .cellVertices(2).Color = 255# * bBase
        Else
          bValue = 255# * (bBase + bRange * (.altitude - MapArray(loop1, loop2 + 1).altitude))
          
          If bValue < 0 Then
            bValue = 0
          ElseIf bValue > 255 Then
            bValue = 255
          End If
          
          .cellVertices(2).Color = bValue
        End If
        
        .cellVertices(3).Color = .cellVertices(2).Color
        .cellVertices(1).Color = .cellVertices(0).Color
        
        .altitudeR_1000ths = MapArray(loop1 - 1, loop2).altitude * 1000
        .yOffset2 = -(MapArray(loop1 - 1, loop2).altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
          
        .altitudeT_1000ths = MapArray(loop1 - 1, loop2 - 1).altitude * 1000
        .yOffset3 = -(MapArray(loop1 - 1, loop2 - 1).altitude * DXDraw.ISOmetricViewScale + dParams.effWorldBase)
        
        .altitudeC_1000ths = (.altitudeB_1000ths + .altitudeL_1000ths + .altitudeR_1000ths + .altitudeT_1000ths) \ 4
        
        .yOffsetPeak = .yOffset3
        
        If .yOffset1 < .yOffsetPeak Then .yOffsetPeak = .yOffset1
        If .yOffset2 < .yOffsetPeak Then .yOffsetPeak = .yOffset2
      End With
    Next loop2
  Next loop1
  
  'calculate outbounds map cell vertices
  For loop1 = -5 To -1
    For loop2 = -5 To mapTWidth - 1
      With MapArray(loop1, loop2)
        .yOffset0 = -(.altitude * DXDraw.ISOmetricViewScale)
        
        If loop2 >= 0 Then
          .yOffset1 = -(MapArray(0, loop2 - 1).altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
          .yOffset3 = -(MapArray(0, loop2 - 1).altitude * DXDraw.ISOmetricViewScale + dParams.effWorldBase)
        Else
          .yOffset1 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
          .yOffset3 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldBase)
        End If
        
        .yOffset2 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
        
        .cellVertices(0).Color = OutBoundsColour
        .cellVertices(1).Color = OutBoundsColour
        .cellVertices(2).Color = OutBoundsColour
        .cellVertices(3).Color = OutBoundsColour
        
        .yOffsetPeak = .yOffset3
        
        If .yOffset1 < .yOffsetPeak Then .yOffsetPeak = .yOffset1
        If .yOffset2 < .yOffsetPeak Then .yOffsetPeak = .yOffset2
      End With
    Next loop2
  Next loop1
  
  For loop1 = -5 To -1
    For loop2 = -5 To mapTHeight - 1
      With MapArray(loop2, loop1)
        .yOffset0 = -(.altitude * DXDraw.ISOmetricViewScale)
        .yOffset1 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
        
        If loop2 >= 0 Then
          .yOffset2 = -(MapArray(loop2 - 1, 0).altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
          .yOffset3 = -(MapArray(loop2 - 1, 0).altitude * DXDraw.ISOmetricViewScale + dParams.effWorldBase)
        Else
          .yOffset2 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
          .yOffset3 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldBase)
        End If
        
        .cellVertices(0).Color = OutBoundsColour
        .cellVertices(1).Color = OutBoundsColour
        .cellVertices(2).Color = OutBoundsColour
        .cellVertices(3).Color = OutBoundsColour
        
        .yOffsetPeak = .yOffset3
        
        If .yOffset1 < .yOffsetPeak Then .yOffsetPeak = .yOffset1
        If .yOffset2 < .yOffsetPeak Then .yOffsetPeak = .yOffset2
      End With
    Next loop2
  Next loop1
  
  For loop1 = mapTHeight - 5 To mapTHeight - 1
    For loop2 = -5 To mapTWidth - 1
      With MapArray(loop1, loop2)
        .yOffset0 = -(.altitude * DXDraw.ISOmetricViewScale)
        
        If loop2 >= 0 Then
          .yOffset1 = -(MapArray(MapHeight - 1, loop2 - 1).altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
          .yOffset3 = -(MapArray(MapHeight - 1, loop2 - 1).altitude * DXDraw.ISOmetricViewScale + dParams.effWorldBase)
        Else
          .yOffset1 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
          .yOffset3 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldBase)
        End If
        
        .yOffset2 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
        
        .cellVertices(0).Color = OutBoundsColour
        .cellVertices(1).Color = OutBoundsColour
        .cellVertices(2).Color = OutBoundsColour
        .cellVertices(3).Color = OutBoundsColour
        
        .yOffsetPeak = .yOffset3
        
        If .yOffset1 < .yOffsetPeak Then .yOffsetPeak = .yOffset1
        If .yOffset2 < .yOffsetPeak Then .yOffsetPeak = .yOffset2
      End With
    Next loop2
  Next loop1
  
  For loop1 = mapTWidth - 5 To mapTWidth - 1
    For loop2 = -5 To mapTHeight - 1
      With MapArray(loop2, loop1)
        .yOffset0 = -(.altitude * DXDraw.ISOmetricViewScale)
        .yOffset1 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
        
        If loop2 >= 0 Then
          .yOffset2 = -(MapArray(loop2 - 1, MapWidth - 1).altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
          .yOffset3 = -(MapArray(loop2 - 1, MapWidth - 1).altitude * DXDraw.ISOmetricViewScale + dParams.effWorldBase)
        Else
          .yOffset2 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldHalfBase)
          .yOffset3 = -(.altitude * DXDraw.ISOmetricViewScale + dParams.effWorldBase)
        End If
        
        .cellVertices(0).Color = OutBoundsColour
        .cellVertices(1).Color = OutBoundsColour
        .cellVertices(2).Color = OutBoundsColour
        .cellVertices(3).Color = OutBoundsColour
        
        .yOffsetPeak = .yOffset3
        
        If .yOffset1 < .yOffsetPeak Then .yOffsetPeak = .yOffset1
        If .yOffset2 < .yOffsetPeak Then .yOffsetPeak = .yOffset2
      End With
    Next loop2
  Next loop1
  
  With MapArray(0, 0) ' calculate first row/col lighting
    bValue = .cellVertices(0).Color
    If bValue > 255 Then bValue = 255
    bValue = (((bValue * 256&) + bValue) * 256&) + bValue
    
    .cellVertices(0).Color = &HFF000000 Or bValue
    
    bValue = .cellVertices(1).Color
    If bValue > 255 Then bValue = 255
    bValue = (((bValue * 256&) + bValue) * 256&) + bValue
    
    .cellVertices(1).Color = &HFF000000 Or bValue
    
    bValue = .cellVertices(3).Color
    If bValue > 255 Then bValue = 255
    bValue = (((bValue * 256&) + bValue) * 256&) + bValue
    
    .cellVertices(3).Color = &HFF000000 Or bValue
  End With
  
  For loop1 = 1 To MapHeight - 1
    With MapArray(loop1, 0)
      bValue = .cellVertices(0).Color
      If bValue > 255 Then bValue = 255
      bValue = (((bValue * 256&) + bValue) * 256&) + bValue
      
      .cellVertices(0).Color = &HFF000000 Or bValue
      
      bValue = .cellVertices(1).Color
      If bValue > 255 Then bValue = 255
      bValue = (((bValue * 256&) + bValue) * 256&) + bValue
    
      .cellVertices(1).Color = &HFF000000 Or bValue
    End With
  Next loop1
  
  For loop2 = 1 To MapWidth - 1
    With MapArray(0, loop2)
      bValue = .cellVertices(3).Color
      If bValue > 255 Then bValue = 255
      bValue = (((bValue * 256&) + bValue) * 256&) + bValue
      
      .cellVertices(3).Color = &HFF000000 Or bValue
      
      bValue = .cellVertices(1).Color
      If bValue > 255 Then bValue = 255
      bValue = (((bValue * 256&) + bValue) * 256&) + bValue
    
      .cellVertices(1).Color = &HFF000000 Or bValue
    End With
  Next loop2
  
  For loop1 = 0 To MapHeight - 1 'calculate map lighting
    For loop2 = 0 To MapWidth - 1
      With MapArray(loop1, loop2)
        If loop1 = MapHeight - 1 Then
          If loop2 = MapWidth - 1 Then
            bValue = .cellVertices(2).Color
            If bValue > 255 Then bValue = 255
            bValue = &HFF000000 Or ((((bValue * 256&) + bValue) * 256&) + bValue)
            
            .cellVertices(2).Color = bValue
          Else
            bValue = (.cellVertices(2).Color + MapArray(loop1, loop2 + 1).cellVertices(0).Color) \ 2
            If bValue > 255 Then bValue = 255
            bValue = &HFF000000 Or ((((bValue * 256&) + bValue) * 256&) + bValue)
            
            .cellVertices(2).Color = bValue
            MapArray(loop1, loop2 + 1).cellVertices(0).Color = bValue
          End If
        Else
          If loop2 = MapWidth - 1 Then
            bValue = (.cellVertices(2).Color + MapArray(loop1 + 1, loop2).cellVertices(3).Color) \ 2
            If bValue > 255 Then bValue = 255
            bValue = &HFF000000 Or ((((bValue * 256&) + bValue) * 256&) + bValue)
            
            .cellVertices(2).Color = bValue
            MapArray(loop1 + 1, loop2).cellVertices(3).Color = bValue
          Else
            bValue = (.cellVertices(2).Color + MapArray(loop1, loop2 + 1).cellVertices(0).Color + MapArray(loop1 + 1, loop2).cellVertices(3).Color + MapArray(loop1 + 1, loop2 + 1).cellVertices(1).Color) \ 4
            If bValue > 255 Then bValue = 255
            bValue = &HFF000000 Or ((((bValue * 256&) + bValue) * 256&) + bValue)
            
            .cellVertices(2).Color = bValue
            MapArray(loop1, loop2 + 1).cellVertices(0).Color = bValue
            MapArray(loop1 + 1, loop2).cellVertices(3).Color = bValue
            MapArray(loop1 + 1, loop2 + 1).cellVertices(1).Color = bValue
          End If
        End If
      End With
    Next loop2
  Next loop1
  
  For loop1 = 0 To MapHeight - 1 'ensure that all blank cells are OutBoundsColour
    For loop2 = 0 To MapWidth - 1
      With MapArray(loop1, loop2)
        If .staticSurface = 0 Then
          .cellVertices(0).Color = OutBoundsColour
          .cellVertices(1).Color = OutBoundsColour
          .cellVertices(2).Color = OutBoundsColour
          .cellVertices(3).Color = OutBoundsColour
        End If
      End With
    Next loop2
  Next loop1
End Sub

Public Sub Validate_Map(ByVal startRow As Long, ByVal startColumn As Long)
  Dim loop1 As Long, loop2 As Long, maxRing As Long, temp1 As Long, temp2 As Long
  Dim rowHold As Long, colHold As Long, validationError As Boolean
  Dim trackRow(1 To 8) As Long, trackCol(1 To 8) As Long
  Dim loop3 As Long, loop4 As Long
  
  ReDim validateMap(0 To MapHeight - 1, 0 To MapWidth - 1)
  
  rowHold = startRow
  colHold = startColumn
  
  startRow = rowHold
  startColumn = colHold
  
  For loop1 = 1 To 8
    trackRow(loop1) = -1
    trackCol(loop1) = -1
  Next loop1
  
  Do
    maxRing = MapHeight - (startRow + 1)
    If maxRing < startRow Then maxRing = startRow
    If maxRing < MapWidth - (startColumn + 1) Then maxRing = MapWidth - (startColumn + 1)
    If maxRing < startColumn Then maxRing = startColumn
    
    For loop1 = 0 To MapHeight - 1
      For loop2 = 0 To MapWidth - 1
        validateMap(loop1, loop2) = False
      Next loop2
    Next loop1
    
    validateTest rowHold, colHold
    
    For loop1 = 0 To maxRing
      For loop2 = -loop1 To loop1
        validationError = validateTest(startRow + loop2, startColumn - loop1)
        
        If validationError Then
          startRow = startRow + loop2
          startColumn = startColumn - loop1
          
          Exit For
        End If
         
        validationError = validateTest(startRow + loop2, startColumn + loop1)
        
        If validationError Then
          startRow = startRow + loop2
          startColumn = startColumn + loop1
          
          Exit For
        End If
        
        validationError = validateTest(startRow - loop1, startColumn + loop2)
        
        If validationError Then
          startRow = startRow - loop1
          startColumn = startColumn + loop2
          
          Exit For
        End If
        
        validationError = validateTest(startRow + loop1, startColumn + loop2)
        
        If validationError Then
          startRow = startRow + loop1
          startColumn = startColumn + loop2
          
          Exit For
        End If
      Next loop2
       
      If validationError Then 'logic to prevent an endless loop
        If startRow = trackRow(8) And startColumn = trackCol(8) Then 'last ditch prevention
          For loop3 = 0 To MapHeight - 1
            For loop4 = 0 To MapWidth - 1
              MapArray(loop3, loop4).altitude = (MapArray(rowHold, colHold).altitude + MapArray(loop3, loop4).altitude) \ 2
            Next loop4
          Next loop3
        ElseIf (startRow = trackRow(6) And startColumn = trackCol(6)) Or (startRow = trackRow(7) And startColumn = trackCol(7)) Then
          For loop3 = -4 To 4
            temp1 = startRow + loop3
            
            If temp1 > 0 And temp1 < MapHeight Then
              For loop4 = -4 To 4
                temp2 = startColumn + loop4
                
                If temp2 > 0 And temp2 < MapWidth Then MapArray(temp1, temp2).altitude = (MapArray(rowHold, colHold).altitude + MapArray(temp1, temp2).altitude) \ 2
              Next loop4
            End If
          Next loop3
        ElseIf (startRow = trackRow(4) And startColumn = trackCol(4)) Or (startRow = trackRow(5) And startColumn = trackCol(5)) Then
          For loop3 = -2 To 2
            temp1 = startRow + loop3
            
            If temp1 > 0 And temp1 < MapHeight Then
              For loop4 = -2 To 2
                temp2 = startColumn + loop4
                
                If temp2 > 0 And temp2 < MapWidth Then MapArray(temp1, temp2).altitude = (MapArray(rowHold, colHold).altitude + MapArray(temp1, temp2).altitude) \ 2
              Next loop4
            End If
          Next loop3
        ElseIf (startRow = trackRow(2) And startColumn = trackCol(2)) Or (startRow = trackRow(3) And startColumn = trackCol(3)) Then
          For loop3 = -1 To 1
            temp1 = startRow + loop3
            
            If temp1 > 0 And temp1 < MapHeight Then
              For loop4 = -1 To 1
                temp2 = startColumn + loop4
                
                If temp2 > 0 And temp2 < MapWidth Then MapArray(temp1, temp2).altitude = (MapArray(rowHold, colHold).altitude + MapArray(temp1, temp2).altitude) \ 2
              Next loop4
            End If
          Next loop3
        End If
        
        For loop3 = 8 To 2 Step -1
          trackRow(loop3) = trackRow(loop3 - 1)
          trackCol(loop3) = trackCol(loop3 - 1)
        Next loop3
       
        trackRow(1) = startRow
        trackCol(1) = startColumn
        
        Exit For
      End If
    Next loop1
  Loop While validationError
End Sub

Private Function validateTest(ByVal sRow As Long, ByVal sCol As Long) As Boolean
  Dim maxAlt As Long, minAlt As Long, validAlt As Boolean, forcedAlt As Long, validForcedAlt As Boolean
  Dim forceFlatX As Boolean, forceFlatY As Boolean, forceFlatXY As Boolean
  Dim altDiff As Long, allowedAltitude As Long
  
  validateTest = False
  
  If sRow < 0 Or sCol < 0 Or sRow >= MapHeight Or sCol >= MapWidth Then Exit Function
  If validateMap(sRow, sCol) = True Then Exit Function
  
  allowedAltitude = worldBase - 2
  
  validAlt = False
  validForcedAlt = False
  forceFlatX = False
  forceFlatY = False
  forceFlatXY = False
  
  With MapArray(sRow, sCol)
    If sRow > 0 Then
      With MapArray(sRow - 1, sCol)
        If validateMap(sRow - 1, sCol) Then
          If validAlt Then
            If minAlt > .altitude Then minAlt = .altitude
            If maxAlt < .altitude Then maxAlt = .altitude
          Else
            validAlt = True
            
            minAlt = .altitude
            maxAlt = .altitude
          End If
        ElseIf .flatLock Then
          forceFlatY = True
        End If
      End With
          
      If sCol > 0 Then
        With MapArray(sRow - 1, sCol - 1)
          If validateMap(sRow - 1, sCol - 1) Then
            If validAlt Then
              If minAlt > .altitude Then minAlt = .altitude
              If maxAlt < .altitude Then maxAlt = .altitude
            Else
              validAlt = True
              
              minAlt = .altitude
              maxAlt = .altitude
            End If
          ElseIf .flatLock Then
            forceFlatXY = True
          End If
        End With
      End If
      
      If sCol < MapWidth - 1 Then
        If validateMap(sRow - 1, sCol + 1) Then
          With MapArray(sRow - 1, sCol + 1)
            If validAlt Then
              If minAlt > .altitude Then minAlt = .altitude
              If maxAlt < .altitude Then maxAlt = .altitude
            Else
              validAlt = True
              
              minAlt = .altitude
              maxAlt = .altitude
            End If
          End With
        End If
      End If
    End If
    
    If sRow < MapHeight - 1 Then
      If validateMap(sRow + 1, sCol) Then
        With MapArray(sRow + 1, sCol)
          If validAlt Then
            If minAlt > .altitude Then minAlt = .altitude
            If maxAlt < .altitude Then maxAlt = .altitude
          Else
            validAlt = True
            
            minAlt = .altitude
            maxAlt = .altitude
          End If
          
          If .flatLock Then
            If validForcedAlt Then
              If forcedAlt <> .altitude Then validateTest = True
            Else
              validForcedAlt = True
              forcedAlt = .altitude
            End If
          End If
        End With
      End If
        
      If sCol > 0 Then
        If validateMap(sRow + 1, sCol - 1) Then
          With MapArray(sRow + 1, sCol - 1)
            If validAlt Then
              If minAlt > .altitude Then minAlt = .altitude
              If maxAlt < .altitude Then maxAlt = .altitude
            Else
              validAlt = True
              
              minAlt = .altitude
              maxAlt = .altitude
            End If
          End With
        End If
      End If
      
      If sCol < MapWidth - 1 Then
        If validateMap(sRow + 1, sCol + 1) Then
          With MapArray(sRow + 1, sCol + 1)
            If validAlt Then
              If minAlt > .altitude Then minAlt = .altitude
              If maxAlt < .altitude Then maxAlt = .altitude
            Else
              validAlt = True
              
              minAlt = .altitude
              maxAlt = .altitude
            End If
            
            If .flatLock Then
              If validForcedAlt Then
                If forcedAlt <> .altitude Then validateTest = True
              Else
                validForcedAlt = True
                forcedAlt = .altitude
              End If
            End If
          End With
        End If
      End If
    End If
    
    If sCol > 0 Then
      With MapArray(sRow, sCol - 1)
        If validateMap(sRow, sCol - 1) Then
          If validAlt Then
            If minAlt > .altitude Then minAlt = .altitude
            If maxAlt < .altitude Then maxAlt = .altitude
          Else
            validAlt = True
            
            minAlt = .altitude
            maxAlt = .altitude
          End If
        ElseIf .flatLock Then
          forceFlatX = True
        End If
      End With
    End If
    
    If sCol < MapWidth - 1 Then
      If validateMap(sRow, sCol + 1) Then
        With MapArray(sRow, sCol + 1)
          If validAlt Then
            If minAlt > .altitude Then minAlt = .altitude
            If maxAlt < .altitude Then maxAlt = .altitude
          Else
            validAlt = True
            
            minAlt = .altitude
            maxAlt = .altitude
          End If
          
          If .flatLock Then
            If validForcedAlt Then
              If forcedAlt <> .altitude Then validateTest = True
            Else
              validForcedAlt = True
              forcedAlt = .altitude
            End If
          End If
        End With
      End If
    End If
    
    If validAlt Then
      If validForcedAlt Then .altitude = forcedAlt
      
      If .altitude > minAlt + allowedAltitude Then
        .altitude = minAlt + allowedAltitude
      ElseIf .altitude < maxAlt - allowedAltitude Then
        .altitude = maxAlt - allowedAltitude
      End If
      
      If maxAlt - minAlt > 2 * allowedAltitude Then
        validateTest = True
      ElseIf .altitude > minAlt + allowedAltitude Then
        validateTest = True
      ElseIf .altitude < maxAlt - allowedAltitude Then
        validateTest = True
      End If
      
      If validForcedAlt Then If .altitude <> forcedAlt Then validateTest = True
    End If
    
    If validateTest = False Then validateImage sRow, sCol, .altitude, True
    
    If validateTest = False Then
      If forceFlatX Then validateTest sRow, sCol - 1
      If forceFlatY Then validateTest sRow - 1, sCol
      If forceFlatXY Then validateTest sRow - 1, sCol - 1
    End If
  End With
End Function

'worm around map testing all map format 0 cells that are interconnected and locking altitudes
Private Sub validateImage(ByVal sRow As Long, ByVal sCol As Long, ByVal sAlt As Long, Optional ByVal forceValidation As Boolean = False)
  If validateMap(sRow, sCol) Then Exit Sub
  
  With MapArray(sRow, sCol)
    If .flatLock Then
      validateMap(sRow, sCol) = True
      .altitude = sAlt
      
      If sRow > 0 Then
        validateImage sRow - 1, sCol, sAlt
        
        If sCol > 0 Then validateImage sRow - 1, sCol - 1, sAlt
        If sCol < MapWidth - 1 Then validateImage sRow - 1, sCol + 1, sAlt
      End If
      
      If sRow < MapHeight - 1 Then
        validateImage sRow + 1, sCol, sAlt
        
        If sCol > 0 Then validateImage sRow + 1, sCol - 1, sAlt
        If sCol < MapWidth - 1 Then validateImage sRow + 1, sCol + 1, sAlt
      End If
      
      If sCol > 0 Then validateImage sRow, sCol - 1, sAlt
      If sCol < MapWidth - 1 Then validateImage sRow, sCol + 1, sAlt
    ElseIf forceValidation Then
      validateMap(sRow, sCol) = True
      .altitude = sAlt
      
      If sRow < MapHeight - 1 Then
        validateImage sRow + 1, sCol, sAlt
        
        If sCol < MapWidth - 1 Then validateImage sRow + 1, sCol + 1, sAlt
      End If
      
      If sCol < MapWidth - 1 Then validateImage sRow, sCol + 1, sAlt
    End If
  End With
End Sub

Public Sub ReInitialize_D3D()
  Set dx_DirectX = DXDraw.GetDirectX()
  Set dx_Direct3D = DXDraw.GetDirect3D()
  Set dx_Direct3DDevice = DXDraw.GetDirect3DDevice()
End Sub

Private Sub Class_Initialize()
  Dim loop1 As Long
  
  ReInitialize_D3D
  
  For loop1 = 0 To 7
    tempVertices(loop1).rhw = 1
  Next loop1
  
  minBrightness = 0.4
  maxBrightness = 1#
  
  gridColour = &H80000000
  OutBoundsColour = &HFF000000
End Sub

Public Sub Save_ToOpenFile(fileHandle As Long)
  Dim rowNum As Long, colNum As Long
  Dim bitmapSurfaceName As String, bitmapSurfaceWidth As Long, bitmapIsTexture As Boolean
  Dim bitmapSurfaceHeight As Long, bitmapSurfaceTransparency As Long, numberOfBitmapSurfaces As Long
  Dim loop1 As Long, loop2 As Long, staticSurfaceInUse() As Boolean
    
  numberOfBitmapSurfaces = DXDraw.Get_TotalStaticSurfaces()
  
  ReDim staticSurfaceInUse(0 To numberOfBitmapSurfaces)
  
  For loop1 = MapHeight - 1 To 0 Step -1 'only need static surfaces in use by map
    For loop2 = MapWidth - 1 To 0 Step -1
      With MapArray(loop1, loop2)
        staticSurfaceInUse(.staticSurface) = True
      End With
    Next loop2
  Next loop1
  
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
  
  Write #fileHandle, DisplayColumn_1000ths
  Write #fileHandle, DisplayRow_1000ths
  
  Write #fileHandle, OutBoundsColour
  
  Write #fileHandle, SurfacePixelRatioU
  Write #fileHandle, SurfacePixelRatioV
  Write #fileHandle, ImageWidthU
  Write #fileHandle, ImageHeightV
  Write #fileHandle, ImagesPerRow
  
  Write #fileHandle, worldBase
  
  Write #fileHandle, MapWidth
  Write #fileHandle, MapHeight
  
  Write #fileHandle, minBrightness
  Write #fileHandle, maxBrightness
  Write #fileHandle, gridColour
  
  For rowNum = -5 To mapTHeight - 1
    For colNum = -5 To mapTWidth - 1
      With MapArray(rowNum, colNum)
        Write #fileHandle, .staticSurface
        Write #fileHandle, .image
        Write #fileHandle, .altitude
        Write #fileHandle, .value
        Write #fileHandle, .flatLock
        Write #fileHandle, .ghostWorldObjects
        
        Write #fileHandle, .fogOfWar(0)
        Write #fileHandle, .fogOfWar(1)
        Write #fileHandle, .fogOfWar(2)
        Write #fileHandle, .fogOfWar(3)
        Write #fileHandle, .fogOfWarStatus
      End With
    Next colNum
  Next rowNum
  
  '**********************************************************************************************
  'use this section for custom member variables
  '**********************************************************************************************
  
  
  
  
  
  '**********************************************************************************************
End Sub

Public Sub Load_FromFile(filePathName As String)
  Dim fileHandle As Long, versionString As String
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Input As #fileHandle
    Input #fileHandle, versionString
    
    Load_FromOpenFile fileHandle
  Close fileHandle
  
  Calculate_Map
  
  Exit Sub
  
badFile:
  MsgBox "Error loading map object -> " & filePathName, vbCritical, "ERROR"
End Sub

Public Sub Load_FromOpenFile(ByVal fileHandle As Long)
  Dim rowNum As Long, colNum As Long
  Dim bitmapSurfaceIndex() As Long, bitmapSurfaceName As String, bitmapSurfaceWidth As Long
  Dim bitmapSurfaceHeight As Long, bitmapSurfaceTransparency As Long, numberOfBitmapSurfaces As Long
  Dim loop1 As Long, numberOfSurfacesInUse As Long, surfaceNumber As Long
  
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
        DXDraw.Init_TextureSurface bitmapSurfaceIndex(surfaceNumber), bitmapSurfaceName, bitmapSurfaceWidth, bitmapSurfaceHeight, bitmapSurfaceTransparency
      End If
    End If
  Next loop1
  
  Input #fileHandle, DisplayColumn_1000ths
  Input #fileHandle, DisplayRow_1000ths
  
  Input #fileHandle, OutBoundsColour
  
  Input #fileHandle, SurfacePixelRatioU
  Input #fileHandle, SurfacePixelRatioV
  Input #fileHandle, ImageWidthU
  Input #fileHandle, ImageHeightV
  Input #fileHandle, ImagesPerRow
  
  Input #fileHandle, worldBase
  
  Input #fileHandle, MapWidth
  Input #fileHandle, MapHeight
  
  Input #fileHandle, minBrightness
  Input #fileHandle, maxBrightness
  Input #fileHandle, gridColour
  
  mapTWidth = MapWidth + 5
  mapTHeight = MapHeight + 5
  
  ReDim MapArray(-5 To mapTHeight - 1, -5 To mapTWidth - 1)
  
  For rowNum = -5 To mapTHeight - 1
    For colNum = -5 To mapTWidth - 1
      With MapArray(rowNum, colNum)
        Input #fileHandle, .staticSurface
        
        .staticSurface = bitmapSurfaceIndex(.staticSurface) 'remap the static surface
        
        Input #fileHandle, .image
        
        Set_MapCell_Image rowNum, colNum, .image
        
        Input #fileHandle, .altitude
        Input #fileHandle, .value
        Input #fileHandle, .flatLock
        Input #fileHandle, .ghostWorldObjects
        
        Input #fileHandle, .fogOfWar(0)
        Input #fileHandle, .fogOfWar(1)
        Input #fileHandle, .fogOfWar(2)
        Input #fileHandle, .fogOfWar(3)
        Input #fileHandle, .fogOfWarStatus
      End With
    Next colNum
  Next rowNum
  
  '**********************************************************************************************
  'use this section for custom member variables
  '**********************************************************************************************
  
  
  
  
  
  '**********************************************************************************************
End Sub

'translates the world XYZ coordinates to the display ISOmetric coordinates for the passed coordinates
Public Sub Translate_WorldXYZ_to_ISODisplay(ByVal BasePosX_1000ths As Long, ByVal BasePosY_1000ths As Long, ByVal BasePosZ_1000ths As Long, _
        ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long, ByVal worldPosZ_1000ths As Long, _
        ByRef displayX As Long, ByRef displayY As Long)
      
  Dim effX As Long, effY As Long
  
  effX = worldPosX_1000ths - BasePosX_1000ths
  effY = worldPosY_1000ths - BasePosY_1000ths
  
  displayX = DXDraw.m_ClippingRectangleWidth \ 2 + ((effX - effY) * DXDraw.ISOmetricViewScale) \ 1000
  displayY = (((effX + effY) \ 2 - (worldPosZ_1000ths - BasePosZ_1000ths)) * DXDraw.ISOmetricViewScale) \ 1000
End Sub

'***********************************************************************
' ISOmetric objects
'***********************************************************************

Public Sub Move_ISOmetricObject(ByVal ISOmetricObject As clsISO_Object)
  Dim mapLevel As Long
  
  mapLevel = ISOmetricObject.mapChainLevel
  
  Remove_ISOmetricObject ISOmetricObject
  Add_ISOmetricObject ISOmetricObject, mapLevel
End Sub

Public Sub Add_ISOmetricObject(ByVal ISOmetricObject As clsISO_Object, ByVal mapLevel As Long)
  Dim chainList As clsISO_Object, renderPriority As Long
  
  With ISOmetricObject
    renderPriority = .PosX_1000ths + .PosY_1000ths
    .renderPriority = renderPriority
    
    Select Case mapLevel
      Case 0
        .mapChainLevel = 0
        
        If isoObjectGroundFXChain Is Nothing Then
          Set .mapChainNext = Nothing
          Set .mapChainPrevious = Nothing
          
          Set isoObjectGroundFXChain = ISOmetricObject
        Else
          Set chainList = isoObjectGroundFXChain
          
          Do While Not (chainList Is Nothing)
            With chainList
              If renderPriority <= .renderPriority Then
                Set ISOmetricObject.mapChainNext = chainList
                Set ISOmetricObject.mapChainPrevious = .mapChainPrevious
                
                If .mapChainPrevious Is Nothing Then
                  Set isoObjectGroundFXChain = ISOmetricObject
                Else
                  Set .mapChainPrevious.mapChainNext = ISOmetricObject
                End If
                
                Set .mapChainPrevious = ISOmetricObject
                
                Set chainList = Nothing
              ElseIf .mapChainNext Is Nothing Then
                Set .mapChainNext = ISOmetricObject
                
                Set ISOmetricObject.mapChainPrevious = chainList
                Set ISOmetricObject.mapChainNext = Nothing
                
                Set chainList = Nothing
              Else
                Set chainList = .mapChainNext
              End If
            End With
          Loop
        End If
      Case 1
        .mapChainLevel = 1
        .mapChainRow = .PosY_1000ths \ (1000 * worldBase)
        .mapChainColumn = .PosX_1000ths \ (1000 * worldBase)
        
        If .mapChainRow < -5 Or .mapChainColumn < -5 Or .mapChainRow >= mapTHeight Or .mapChainColumn >= mapTWidth Then Exit Sub
        
        Set chainList = MapArray(.mapChainRow, .mapChainColumn).isoObjectChain
        
        If chainList Is Nothing Then
          Set .mapChainNext = Nothing
          Set .mapChainPrevious = Nothing
          
          Set MapArray(.mapChainRow, .mapChainColumn).isoObjectChain = ISOmetricObject
        Else
          Do While Not (chainList Is Nothing)
            With chainList
              If renderPriority <= .renderPriority Then
                Set ISOmetricObject.mapChainNext = chainList
                Set ISOmetricObject.mapChainPrevious = .mapChainPrevious
                
                If .mapChainPrevious Is Nothing Then
                  Set MapArray(.mapChainRow, .mapChainColumn).isoObjectChain = ISOmetricObject
                Else
                  Set .mapChainPrevious.mapChainNext = ISOmetricObject
                End If
                
                Set .mapChainPrevious = ISOmetricObject
                
                Set chainList = Nothing
              ElseIf .mapChainNext Is Nothing Then
                Set .mapChainNext = ISOmetricObject
                
                Set ISOmetricObject.mapChainPrevious = chainList
                Set ISOmetricObject.mapChainNext = Nothing
                
                Set chainList = Nothing
              Else
                Set chainList = .mapChainNext
              End If
            End With
          Loop
        End If
      Case 2
        .mapChainLevel = 2
        
        If isoObjectChain Is Nothing Then
          Set .mapChainNext = Nothing
          Set .mapChainPrevious = Nothing
          
          Set isoObjectChain = ISOmetricObject
        Else
          Set chainList = isoObjectChain
          
          Do While Not (chainList Is Nothing)
            With chainList
              If renderPriority <= .renderPriority Then
                Set ISOmetricObject.mapChainNext = chainList
                Set ISOmetricObject.mapChainPrevious = .mapChainPrevious
                
                If .mapChainPrevious Is Nothing Then
                  Set isoObjectChain = ISOmetricObject
                Else
                  Set .mapChainPrevious.mapChainNext = ISOmetricObject
                End If
                
                Set .mapChainPrevious = ISOmetricObject
                
                Set chainList = Nothing
              ElseIf .mapChainNext Is Nothing Then
                Set .mapChainNext = ISOmetricObject
                
                Set ISOmetricObject.mapChainPrevious = chainList
                Set ISOmetricObject.mapChainNext = Nothing
                
                Set chainList = Nothing
              Else
                Set chainList = .mapChainNext
              End If
            End With
          Loop
        End If
    End Select
  End With
End Sub

Public Sub Remove_ISOmetricObject(ByVal ISOmetricObject As clsISO_Object)
  With ISOmetricObject
    Select Case .mapChainLevel
      Case 0
        If Not (.mapChainNext Is Nothing) Then Set .mapChainNext.mapChainPrevious = .mapChainPrevious
        
        If .mapChainPrevious Is Nothing Then
          Set isoObjectGroundFXChain = .mapChainNext
        Else
          Set .mapChainPrevious.mapChainNext = .mapChainNext
          Set .mapChainPrevious = Nothing
        End If
        
        Set .mapChainNext = Nothing
      Case 1
        If .mapChainRow < -5 Or .mapChainColumn < -5 Or .mapChainRow >= mapTHeight Or .mapChainColumn >= mapTWidth Then Exit Sub
        
        If Not (.mapChainNext Is Nothing) Then Set .mapChainNext.mapChainPrevious = .mapChainPrevious
        
        If .mapChainPrevious Is Nothing Then
          Set MapArray(.mapChainRow, .mapChainColumn).isoObjectChain = .mapChainNext
        Else
          Set .mapChainPrevious.mapChainNext = .mapChainNext
          Set .mapChainPrevious = Nothing
        End If
        
        Set .mapChainNext = Nothing
      Case 2
        If Not (.mapChainNext Is Nothing) Then Set .mapChainNext.mapChainPrevious = .mapChainPrevious
        
        If .mapChainPrevious Is Nothing Then
          Set isoObjectChain = .mapChainNext
        Else
          Set .mapChainPrevious.mapChainNext = .mapChainNext
          Set .mapChainPrevious = Nothing
        End If
        
        Set .mapChainNext = Nothing
    End Select
    
    .mapChainLevel = -1
  End With
End Sub

Public Sub Add_ISOmetricObjectChain(ByVal chainStart As clsISO_Object, ByVal mapLevel As Long)
  Do While Not (chainStart Is Nothing)
    Add_ISOmetricObject chainStart, mapLevel
    
    Set chainStart = chainStart.chainNext
  Loop
End Sub

Public Sub Remove_ISOmetricObjectChain(ByVal chainStart As clsISO_Object)
  Do While Not (chainStart Is Nothing)
    Remove_ISOmetricObject chainStart
    
    Set chainStart = chainStart.chainNext
  Loop
End Sub

Public Sub RemoveAll_ISOmetricObjects()
  Dim chainList As clsISO_Object, row As Long, col As Long
  
  Set chainList = isoObjectChain
  
  Do While Not (chainList Is Nothing)
    With chainList
      Set chainList = .mapChainNext
      
      Set .mapChainNext = Nothing
      Set .mapChainPrevious = Nothing
      
      .mapChainLevel = -1
    End With
  Loop
  
  Set isoObjectChain = Nothing
  
  Set chainList = isoObjectGroundFXChain
  
  Do While Not (chainList Is Nothing)
    With chainList
      Set chainList = .mapChainNext
      
      Set .mapChainNext = Nothing
      Set .mapChainPrevious = Nothing
      
      .mapChainLevel = -1
    End With
  Loop
  
  Set isoObjectGroundFXChain = Nothing
  
  For row = mapTHeight - 1 To -5 Step -1
    For col = mapTWidth - 1 To -5 Step -1
      With MapArray(row, col)
        Set chainList = .isoObjectChain
        
        Do While Not (chainList Is Nothing)
          With chainList
            Set chainList = .mapChainNext
            
            Set .mapChainNext = Nothing
            Set .mapChainPrevious = Nothing
            
            .mapChainLevel = -1
          End With
        Loop
        
        Set .isoObjectChain = Nothing
      End With
    Next col
  Next row
End Sub

Public Function Get_MapFirstISOmetricObject(ByVal row As Long, ByVal column As Long, ByVal mapLevel As Long) As clsISO_Object
  Select Case mapLevel
    Case 0
      Set Get_MapFirstISOmetricObject = isoObjectGroundFXChain
    Case 1
      Set Get_MapFirstISOmetricObject = MapArray(row, column).isoObjectChain
    Case 2
      Set Get_MapFirstISOmetricObject = isoObjectChain
  End Select
End Function

Public Function Get_DisplayedISOmetricObject(ByVal xPos As Long, ByVal yPos As Long, ByVal mapLevels As Long, Optional ByVal targetMask As Long = &HFFFFFFFF) As clsISO_Object
  Dim objectList As clsISO_Object, temp1 As Long, loop1 As Long, loop2 As Long
  Dim MapCellX As Long, MapCellY As Long, mapCellXT As Long
  
  If mapLevels And 1 Then
    Set objectList = isoObjectGroundFXChain
    
    Do While Not (objectList Is Nothing)
      With objectList
        If targetMask And .Collision_Mask Then
          If .wasDisplayed Then
            If .displayPosY > yPos Then
              If .displayPosY - (.Collision_WidthZ_1000ths * DXDraw.ISOmetricViewScale) \ 1000 < yPos Then
                temp1 = ((.Collision_WidthX_1000ths + .Collision_WidthY_1000ths) * DXDraw.ISOmetricViewScale) \ 2000
                
                If .displayPosX - temp1 < xPos Then
                  If .displayPosX + temp1 > xPos Then Set Get_DisplayedISOmetricObject = objectList
                End If
              End If
            End If
          End If
        End If
        
        Set objectList = .mapChainNext
      End With
    Loop
  End If
  
  If mapLevels And 2 Then
    temp1 = yPos Mod dParams.effWorldBase
    
    If xPos > dParams.displayOriginXCenter Then
      MapCellX = (((temp1 + (xPos - dParams.displayOriginXCenter + dParams.effWorldBase) \ 2) \ dParams.effWorldBase) + DisplayColumn_1000ths \ 1000) - 1
      MapCellY = ((temp1 + (dParams.displayOriginXCenter - xPos - dParams.effWorldBase) \ 2) \ dParams.effWorldBase) + DisplayRow_1000ths \ 1000
    Else
      MapCellX = (((temp1 + (xPos - dParams.displayOriginXCenter - dParams.effWorldBase) \ 2) \ dParams.effWorldBase) + DisplayColumn_1000ths \ 1000) - 1
      MapCellY = ((temp1 + (dParams.displayOriginXCenter - xPos + dParams.effWorldBase) \ 2) \ dParams.effWorldBase) + DisplayRow_1000ths \ 1000
    End If
    
    For loop1 = (dParams.MapDisplayHeight + 1) \ 2 To 1 Step -1
      If MapCellY >= -5 And MapCellY < mapTHeight Then
        mapCellXT = MapCellX
        
        For loop2 = 1 To 3
          If mapCellXT >= -5 And mapCellXT < mapTWidth Then
            Set objectList = MapArray(MapCellY, mapCellXT).isoObjectChain
            
            Do While Not (objectList Is Nothing)
              With objectList
                If targetMask And .Collision_Mask Then
                  If .wasDisplayed Then
                    If .displayPosY > yPos Then
                      If .displayPosY - (.Collision_WidthZ_1000ths * DXDraw.ISOmetricViewScale) \ 1000 < yPos Then
                        temp1 = ((.Collision_WidthX_1000ths + .Collision_WidthY_1000ths) * DXDraw.ISOmetricViewScale) \ 2000
                        
                        If .displayPosX - temp1 < xPos Then
                          If .displayPosX + temp1 > xPos Then Set Get_DisplayedISOmetricObject = objectList
                        End If
                      End If
                    End If
                  End If
                End If
                
                Set objectList = .mapChainNext
              End With
            Loop
          End If
          
          mapCellXT = mapCellXT + 1
        Next loop2
      End If
      
      MapCellX = MapCellX + 1
      MapCellY = MapCellY + 1
    Next loop1
  End If
  
  If mapLevels And 4 Then
    Set objectList = isoObjectChain
    
    Do While Not (objectList Is Nothing)
      With objectList
        If targetMask And .Collision_Mask Then
          If .wasDisplayed Then
            If .displayPosY > yPos Then
              If .displayPosY - (.Collision_WidthZ_1000ths * DXDraw.ISOmetricViewScale) \ 1000 < yPos Then
                temp1 = ((.Collision_WidthX_1000ths + .Collision_WidthY_1000ths) * DXDraw.ISOmetricViewScale) \ 2000
                
                If .displayPosX - temp1 < xPos Then
                  If .displayPosX + temp1 > xPos Then Set Get_DisplayedISOmetricObject = objectList
                End If
              End If
            End If
          End If
        End If
    
        Set objectList = .mapChainNext
      End With
    Loop
  End If
End Function

Public Sub Destroy_Object()
  On Error Resume Next
  
  Set dx_DirectX = Nothing
  Set dx_Direct3D = Nothing
  Set dx_Direct3DDevice = Nothing
  
  RemoveAll_ISOmetricObjects
  
  Erase MapArray
  
  MapWidth = 0
  MapHeight = 0
End Sub
