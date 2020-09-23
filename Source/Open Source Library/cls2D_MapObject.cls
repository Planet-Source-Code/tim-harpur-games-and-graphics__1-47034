VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls2D_MapObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************************
'
' cls2D_MapObject class definition
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

Private Type MapCell
  staticSurface As Long
  image As Long
  value As Long
  
  ImageXOffset As Long
  ImageYOffset As Long
End Type

Private MapArray() As MapCell

Public DisplayColumn_1000ths As Long       ' Column co-ordinates of map to display
Public DisplayRow_1000ths As Long            ' Row co-ordinates of map to display

Public worldBase As Long        ' base world size (in pixels) of map's cells
                                                  ' this is also the width and height of the source images (in pixels)

Public OutBoundsColour As Long
Public OutBoundsClear As Boolean

Public ImagesPerRow As Long                       ' number of images per row on the image surface(s)

Public MapWidth As Long                                 ' width of the image map in image cells
Public MapHeight As Long                                ' height of the image map in image cells

Public gridColour As Long

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
  MapWidth = Width
  MapHeight = Height
  
  ReDim MapArray(0 To MapHeight - 1, 0 To MapWidth - 1)
End Sub

Public Sub Set_MapCell(ByVal row As Long, ByVal column As Long, ByVal staticSurface As Long, ByVal imageNumber As Long, ByVal cellValue As Long)
  With MapArray(row, column)
    .staticSurface = staticSurface
    .image = imageNumber
    .value = cellValue
    
    .ImageYOffset = (.image \ ImagesPerRow) * worldBase
    .ImageXOffset = (.image Mod ImagesPerRow) * worldBase
  End With
End Sub

Public Function Get_MapCell_StaticSurface(ByVal row As Long, ByVal column As Long) As Long
  Get_MapCell_StaticSurface = MapArray(row, column).staticSurface
End Function

Public Sub Set_MapCell_StaticSurface(ByVal row As Long, ByVal column As Long, ByVal staticSurface As Long)
  MapArray(row, column).staticSurface = staticSurface
End Sub

Public Function Get_MapCell_Image(ByVal row As Long, ByVal column As Long) As Long
  Get_MapCell_Image = MapArray(row, column).image
End Function

Public Sub Set_MapCell_Image(ByVal row As Long, ByVal column As Long, ByVal imageNumber As Long)
  With MapArray(row, column)
    .image = imageNumber
    
    .ImageYOffset = (.image \ ImagesPerRow) * worldBase
    .ImageXOffset = (.image Mod ImagesPerRow) * worldBase
  End With
End Sub

Public Function Get_MapCell_Value(ByVal row As Long, ByVal column As Long) As Long
  Get_MapCell_Value = MapArray(row, column).value
End Function

Public Sub Set_MapCell_Value(ByVal row As Long, ByVal column As Long, ByVal value As Long)
  MapArray(row, column).value = value
End Sub

Public Sub Get_WorldXY_from_MapRC(ByVal row As Long, ByVal column As Long, _
        ByRef worldPosX_1000ths As Long, ByRef worldPosY_1000ths As Long)
        
  worldPosX_1000ths = (column * worldBase + worldBase \ 2) * 1000
  worldPosY_1000ths = (row * worldBase + worldBase \ 2) * 1000
End Sub

Public Sub Get_MapRC_from_WorldXY(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long, _
        ByRef row As Long, ByRef column As Long)
        
  column = worldPosX_1000ths \ (worldBase * 1000)
  row = worldPosY_1000ths \ (worldBase * 1000)
End Sub

Public Sub Set_MapDisplay_from_WorldXY(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long)
  DisplayColumn_1000ths = worldPosX_1000ths \ worldBase
  DisplayRow_1000ths = worldPosY_1000ths \ worldBase
End Sub

Public Sub Get_WorldXY_from_MapDisplay(ByRef worldPosX_1000ths As Long, ByRef worldPosY_1000ths As Long)
  worldPosX_1000ths = DisplayColumn_1000ths * worldBase
  worldPosY_1000ths = DisplayRow_1000ths * worldBase
End Sub

Public Sub Ensure_Visible_WorldXY(ByVal worldPosX_1000ths As Long, ByVal worldPosY_1000ths As Long, Optional ByVal forceCenter As Boolean = False)
  If forceCenter = False Then
    If worldPosX_1000ths < DisplayColumn_1000ths Or worldPosY_1000ths < DisplayRow_1000ths Then
      forceCenter = True
    ElseIf (worldPosX_1000ths \ 1000) >= DXDraw.m_ClippingRectangleWidth Or (worldPosY_1000ths \ 1000) >= DXDraw.m_ClippingRectangleHeight Then
      forceCenter = True
    End If
  End If
  
  If forceCenter Then
    DisplayRow_1000ths = (worldPosY_1000ths - DXDraw.m_ClippingRectangleHeight * 500) \ worldBase
    DisplayColumn_1000ths = (worldPosX_1000ths - DXDraw.m_ClippingRectangleWidth * 500) \ worldBase
  End If
End Sub

Public Sub Ensure_Visible_MapRC(ByVal row As Long, ByVal column As Long, Optional ByVal forceCenter As Boolean = False)
  Dim worldPosX_1000ths As Long, worldPosY_1000ths As Long
  
  worldPosX_1000ths = (column * worldBase + worldBase \ 2) * 1000
  worldPosY_1000ths = (row * worldBase + worldBase \ 2) * 1000
  
  If forceCenter = False Then
    If worldPosX_1000ths < DisplayColumn_1000ths Or worldPosY_1000ths < DisplayRow_1000ths Then
      forceCenter = True
    ElseIf (worldPosX_1000ths \ 1000) >= DXDraw.m_ClippingRectangleWidth Or (worldPosY_1000ths \ 1000) >= DXDraw.m_ClippingRectangleHeight Then
      forceCenter = True
    End If
  End If
  
  If forceCenter Then
    DisplayRow_1000ths = (worldPosY_1000ths - DXDraw.m_ClippingRectangleHeight * 500) \ worldBase
    DisplayColumn_1000ths = (worldPosX_1000ths - DXDraw.m_ClippingRectangleWidth * 500) \ worldBase
  End If
End Sub

'Redraw the XY map layer
Public Sub Render_Map()
  Dim RowX As Long, RowY As Long
  Dim MapDisplayWidth As Long, MapDisplayHeight As Long
  Dim MapDisplayShiftX As Long, MapDisplayShiftY As Long
  Dim MapStartCellX As Long, MapStartCellY As Long, MapLayer As Long
  Dim MapCellX As Long, MapCellY As Long, BasePosX As Long, BasePosY As Long
  
  MapDisplayWidth = DXDraw.m_ClippingRectangleWidth \ worldBase + 2
  MapDisplayHeight = DXDraw.m_ClippingRectangleHeight \ worldBase + 2
  
  MapStartCellX = DisplayColumn_1000ths \ 1000
  MapStartCellY = DisplayRow_1000ths \ 1000
  
  MapDisplayShiftX = ((DisplayColumn_1000ths Mod 1000) * worldBase) \ 1000
  MapDisplayShiftY = ((DisplayRow_1000ths Mod 1000) * worldBase) \ 1000
  
  If MapDisplayShiftX < 0 Then
    MapStartCellX = MapStartCellX - 1
    MapDisplayShiftX = MapDisplayShiftX + worldBase
  End If
  
  If MapDisplayShiftY < 0 Then
    MapStartCellY = MapStartCellY - 1
    MapDisplayShiftY = MapDisplayShiftY + worldBase
  End If
  
  RowY = 0
  BasePosY = -MapDisplayShiftY
  
  Do While RowY < MapDisplayHeight
    MapCellY = RowY + MapStartCellY
    
    RowX = 0
    BasePosX = -MapDisplayShiftX
    
    If MapCellY < 0 Or MapCellY >= MapHeight Then
      If OutBoundsClear Then
        Do While RowX < MapDisplayWidth
          DXDraw.BlitClear BasePosX, BasePosY, worldBase, worldBase, OutBoundsColour
          
          RowX = RowX + 1
          BasePosX = BasePosX + worldBase
        Loop
      End If
    Else
      Do While RowX < MapDisplayWidth
        MapCellX = RowX + MapStartCellX
        
        If MapCellX < 0 Or MapCellX >= MapWidth Then
          If OutBoundsClear Then DXDraw.BlitClear BasePosX, BasePosY, worldBase, worldBase, OutBoundsColour
        Else
          With MapArray(MapCellY, MapCellX)
            DXDraw.BlitTransparent .staticSurface, BasePosX, BasePosY, .ImageXOffset, .ImageYOffset, worldBase, worldBase
          End With
        End If
        
        RowX = RowX + 1
        BasePosX = BasePosX + worldBase
      Loop
    End If
    
    RowY = RowY + 1
    BasePosY = BasePosY + worldBase
  Loop
End Sub

Public Sub Render_Map_with_Grid()
  Dim RowX As Long, RowY As Long
  Dim MapDisplayWidth As Long, MapDisplayHeight As Long
  Dim MapDisplayShiftX As Long, MapDisplayShiftY As Long
  Dim MapStartCellX As Long, MapStartCellY As Long, MapLayer As Long
  Dim MapCellX As Long, MapCellY As Long, BasePosX As Long, BasePosY As Long
  Dim gridX As Long, gridY As Long, gridWidth As Long, gridHeight As Long
  
  Dim backSurface As DirectDrawSurface7
  
  On Error Resume Next
  
  Set backSurface = DXDraw.GetDirectDrawBackSurface
  
  MapDisplayWidth = DXDraw.m_ClippingRectangleWidth \ worldBase + 2
  MapDisplayHeight = DXDraw.m_ClippingRectangleHeight \ worldBase + 2
  
  MapStartCellX = DisplayColumn_1000ths \ 1000
  MapStartCellY = DisplayRow_1000ths \ 1000
  
  MapDisplayShiftX = ((DisplayColumn_1000ths Mod 1000) * worldBase) \ 1000
  MapDisplayShiftY = ((DisplayRow_1000ths Mod 1000) * worldBase) \ 1000
  
  If MapDisplayShiftX < 0 Then
    MapStartCellX = MapStartCellX - 1
    MapDisplayShiftX = MapDisplayShiftX + worldBase
  End If
  
  If MapDisplayShiftY < 0 Then
    MapStartCellY = MapStartCellY - 1
    MapDisplayShiftY = MapDisplayShiftY + worldBase
  End If
  
  With backSurface
    .SetForeColor gridColour
    .setDrawStyle LineStyles.LSSolid
    .setDrawWidth 1
    .SetFillColor 0
    .SetFillStyle FillStyles.FSNoFill
  End With
  
  RowY = 0
  BasePosY = -MapDisplayShiftY
  
  Do While RowY < MapDisplayHeight
    MapCellY = RowY + MapStartCellY
    
    RowX = 0
    BasePosX = -MapDisplayShiftX
    
    If BasePosY < 0 Then
      gridY = 0
      gridHeight = worldBase + BasePosY
    Else
      gridY = BasePosY
      gridHeight = worldBase + gridY
    End If
    
    If gridHeight >= DXDraw.m_ClippingRectangleHeight Then
      gridHeight = DXDraw.m_ClippingRectangleHeight + DXDraw.m_ClippingRectangleY
    Else
      gridHeight = gridHeight + DXDraw.m_ClippingRectangleY
    End If
    
    gridY = gridY + DXDraw.m_ClippingRectangleY
    
    If MapCellY < 0 Or MapCellY >= MapHeight Then
      If OutBoundsClear Then
        Do While RowX < MapDisplayWidth
          DXDraw.BlitClear BasePosX, BasePosY, worldBase, worldBase, OutBoundsColour
          
          RowX = RowX + 1
          BasePosX = BasePosX + worldBase
        Loop
      End If
    Else
      Do While RowX < MapDisplayWidth
        MapCellX = RowX + MapStartCellX
        
        If MapCellX < 0 Or MapCellX >= MapWidth Then
          If OutBoundsClear Then DXDraw.BlitClear BasePosX, BasePosY, worldBase, worldBase, OutBoundsColour
        Else
          With MapArray(MapCellY, MapCellX)
            DXDraw.BlitTransparent .staticSurface, BasePosX, BasePosY, .ImageXOffset, .ImageYOffset, worldBase, worldBase
            
            If BasePosX < 0 Then
              gridX = 0
              gridWidth = worldBase + BasePosX
            Else
              gridX = BasePosX
              gridWidth = worldBase + gridX
            End If
            
            If gridWidth >= DXDraw.m_ClippingRectangleWidth Then
              gridWidth = DXDraw.m_ClippingRectangleWidth + DXDraw.m_ClippingRectangleX
            Else
              gridWidth = gridWidth + DXDraw.m_ClippingRectangleX
            End If
            
            gridX = gridX + DXDraw.m_ClippingRectangleX
            
            backSurface.DrawBox gridX, gridY, gridWidth, gridHeight
          End With
        End If
        
        RowX = RowX + 1
        BasePosX = BasePosX + worldBase
      Loop
    End If
    
    RowY = RowY + 1
    BasePosY = BasePosY + worldBase
  Loop
End Sub

Public Sub Save_ToFile(filePathName As String)
  Dim fileHandle As Long, versionString As String
  
  On Error GoTo badFile
  
  fileHandle = FreeFile()
  
  Open filePathName For Output As #fileHandle
    Write #fileHandle, "Logicon Enterprises cls2D_MapObject Description Ver2.6"
    
    Save_ToOpenFile fileHandle
  Close fileHandle
  
  Exit Sub
  
badFile:
  MsgBox "Error saving 2D map object -> " & filePathName, vbCritical, "ERROR"
End Sub

Public Sub Save_ToOpenFile(fileHandle As Long)
  Dim rowNum As Long, colNum As Long
  Dim bitmapSurfaceName As String, bitmapSurfaceWidth As Long, bitmapIsTexture As Boolean
  Dim bitmapSurfaceHeight As Long, bitmapSurfaceTransparency As Long, numberOfBitmapSurfaces As Long
  Dim loop1 As Long, loop2 As Long, staticSurfaceInUse() As Boolean
    
  numberOfBitmapSurfaces = DXDraw.Get_TotalStaticSurfaces()
  
  'I am using 0 even though it shouldn't occur just to prevent possible error
  ReDim staticSurfaceInUse(0 To numberOfBitmapSurfaces)
  
  For loop1 = MapHeight - 1 To 0 Step -1 'only need static surfaces in use by map
    For loop2 = MapWidth - 1 To 0 Step -1
      staticSurfaceInUse(MapArray(loop1, loop2).staticSurface) = True
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
  
  Write #fileHandle, worldBase
  
  Write #fileHandle, OutBoundsColour
  Write #fileHandle, OutBoundsClear
  
  Write #fileHandle, gridColour
  
  Write #fileHandle, ImagesPerRow
  
  Write #fileHandle, MapWidth
  Write #fileHandle, MapHeight
  
  For rowNum = 0 To MapHeight - 1
    For colNum = 0 To MapWidth - 1
      With MapArray(rowNum, colNum)
        Write #fileHandle, .staticSurface
        Write #fileHandle, .image
        Write #fileHandle, .value
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
  
  Exit Sub
  
badFile:
  MsgBox "Error loading 2D map object -> " & filePathName, vbCritical, "ERROR"
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
        DXDraw.Init_StaticSurface bitmapSurfaceIndex(surfaceNumber), bitmapSurfaceName, bitmapSurfaceWidth, bitmapSurfaceHeight, bitmapSurfaceTransparency
      End If
    End If
  Next loop1
  
  Input #fileHandle, DisplayColumn_1000ths
  Input #fileHandle, DisplayRow_1000ths
  
  Input #fileHandle, worldBase
  
  Input #fileHandle, OutBoundsColour
  Input #fileHandle, OutBoundsClear
  
  Input #fileHandle, gridColour
  
  Input #fileHandle, ImagesPerRow
  
  Input #fileHandle, MapWidth
  Input #fileHandle, MapHeight
  
  ReDim MapArray(0 To MapHeight - 1, 0 To MapWidth - 1)
  
  For rowNum = 0 To MapHeight - 1
    For colNum = 0 To MapWidth - 1
      With MapArray(rowNum, colNum)
        Input #fileHandle, .staticSurface
        
        .staticSurface = bitmapSurfaceIndex(.staticSurface) 'remap the static surface
        
        Input #fileHandle, .image
        Input #fileHandle, .value
        
        Set_MapCell_Image rowNum, colNum, .image
      End With
    Next colNum
  Next rowNum
  
  
  '**********************************************************************************************
  'use this section for custom member variables
  '**********************************************************************************************
  
  
  
  
  
  '**********************************************************************************************
End Sub

Public Sub Destroy_Object()
  Erase MapArray
  
  MapWidth = 0
  MapHeight = 0
End Sub

