VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsISO_ObjectSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************************
'
' clsISO_ObjectSurface class definition
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

Public Visible As Boolean                                                            ' Surface's visibility flag (default TRUE)

Public ParentObject As clsISO_Object                               ' parent object for this Surface
Public PrevSurface As clsISO_ObjectSurface     ' previous Surface in list
Public NextSurface As clsISO_ObjectSurface     ' next Surface in list

Public ZPriority As Long                   ' z render priority value of this Surface

Public TypeID As Long                   ' object type identifier

Public DisplayOffsetX As Long         ' x display offset from ISOmetric Object's center (usually -ve)
Public DisplayOffsetY As Long         ' y display offset from ISOmetric Object's center (usually -ve)
Public DisplayWidth As Long
Public DisplayHeight As Long

Public ActionSequence As Variant    ' array of image indexes for each action frame (0 based array)
Public ActionLast As Long                ' UBOUND of ActionSequence - automatically set by Set_AnimationSequence()
Public ActionFrame As Long            ' current index into the array - determines which frame to display

Public SurfaceIndex As Long                                ' static texture surface index for image list
Public SurfaceOffsetU As Single                            ' u offset to start of image list on texture surface
Public SurfaceOffsetV As Single                            ' v offset to start of image list on texture surface
Public SurfacePixelRatioU As Single
Public SurfacePixelRatioV As Single
Public ImageWidthU As Single                                  ' width of image as a ratio of texture surface width
Public ImageHeightV As Single                                 ' height of image as a ratio of texture surface height
Public ImagesPerRow As Long                            ' number of images per row on surface

Public SurfaceColour As Long                 ' surface ARGB colour

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

Public TicksPerFrame As Long                                  ' user defined value
Public FrameTickCounter As Long                             ' user defined value

Public MainAction As Long                                       ' user defined value
Public TicksPerMainAction As Long                         ' user defined value
Public MainActionTickCounter As Long                   ' user defined value
Public SubAction As Long                                        ' user defined value
Public TicksPerSubAction As Long                         ' user defined value
Public SubActionTickCounter As Long                    ' user defined value

Public User1 As Long                   ' user defined value

'***************************************************************************************************************



Public Sub Set_AnimationSequence(actionSequenceArray As Variant, Optional ByVal resetSequence As Boolean = True)
  ActionLast = UBound(actionSequenceArray)
    
  ActionSequence = actionSequenceArray
  
  If resetSequence = True Then ActionFrame = 0
End Sub

Private Sub Class_Initialize()
  ActionFrame = -1
  ActionLast = -1
  SurfaceColour = &HFFFFFFFF
  
  Visible = True
End Sub

Public Function Duplicate_Object(ByVal newObjectID As Long) As clsISO_ObjectSurface
  Set Duplicate_Object = New clsISO_ObjectSurface
  
  With Duplicate_Object
    .Visible = Visible
     
    .ZPriority = ZPriority
     
    .TypeID = TypeID
     
    .DisplayOffsetX = DisplayOffsetX
    .DisplayOffsetY = DisplayOffsetY
    .DisplayWidth = DisplayWidth
    .DisplayHeight = DisplayHeight
    
    .ActionLast = ActionLast
    .ActionSequence = ActionSequence
           
    .ActionFrame = ActionFrame
    
    .SurfaceIndex = SurfaceIndex
    .SurfaceOffsetU = SurfaceOffsetU
    .SurfaceOffsetV = SurfaceOffsetV
    .SurfacePixelRatioU = SurfacePixelRatioU
    .SurfacePixelRatioV = SurfacePixelRatioV
    .ImageWidthU = ImageWidthU
    .ImageHeightV = ImageHeightV
    .ImagesPerRow = ImagesPerRow
    
    .SurfaceColour = SurfaceColour
    
    '**********************************************************************************************
    'use this section for custom member variables
    '**********************************************************************************************
    
    .TicksPerFrame = TicksPerFrame
    .FrameTickCounter = FrameTickCounter
     
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

Public Sub Save_ToOpenFile(ByVal fileHandle As Long)
  Dim loop1 As Long
  
  Write #fileHandle, Visible
      
  Write #fileHandle, ZPriority
  
  Write #fileHandle, TypeID
  
  Write #fileHandle, DisplayOffsetX
  Write #fileHandle, DisplayOffsetY
  Write #fileHandle, DisplayWidth
  Write #fileHandle, DisplayHeight

  Write #fileHandle, ActionLast
  
  For loop1 = 0 To ActionLast
    Write #fileHandle, ActionSequence(loop1)
  Next loop1
      
  Write #fileHandle, ActionFrame
  
  Write #fileHandle, SurfaceIndex
  Write #fileHandle, SurfaceOffsetU
  Write #fileHandle, SurfaceOffsetV
  Write #fileHandle, SurfacePixelRatioU
  Write #fileHandle, SurfacePixelRatioV
  Write #fileHandle, ImageWidthU
  Write #fileHandle, ImageHeightV
  Write #fileHandle, ImagesPerRow
  
  Write #fileHandle, SurfaceColour
  
  '**********************************************************************************************
  'use this section for custom member variables
  '**********************************************************************************************
  
  Write #fileHandle, TicksPerFrame
  Write #fileHandle, FrameTickCounter
      
  Write #fileHandle, MainAction
  Write #fileHandle, TicksPerMainAction
  Write #fileHandle, MainActionTickCounter
  Write #fileHandle, SubAction
  Write #fileHandle, TicksPerSubAction
  Write #fileHandle, SubActionTickCounter

  Write #fileHandle, User1
  
  '**********************************************************************************************
End Sub

Public Sub Load_FromOpenFile(ByVal fileHandle As Long)
  Dim loop1 As Long
  
  Input #fileHandle, Visible
  
  Input #fileHandle, ZPriority
  
  Input #fileHandle, TypeID

  Input #fileHandle, DisplayOffsetX
  Input #fileHandle, DisplayOffsetY
  Input #fileHandle, DisplayWidth
  Input #fileHandle, DisplayHeight
  
  Input #fileHandle, ActionLast
  
  If ActionLast > -1 Then
    ReDim ActionSequence(0 To ActionLast)
    
    For loop1 = 0 To ActionLast
      Input #fileHandle, ActionSequence(loop1)
    Next loop1
  End If
  
  Input #fileHandle, ActionFrame

  Input #fileHandle, SurfaceIndex
  Input #fileHandle, SurfaceOffsetU
  Input #fileHandle, SurfaceOffsetV
  Input #fileHandle, SurfacePixelRatioU
  Input #fileHandle, SurfacePixelRatioV
  Input #fileHandle, ImageWidthU
  Input #fileHandle, ImageHeightV
  Input #fileHandle, ImagesPerRow
  
  Input #fileHandle, SurfaceColour
  
  '**********************************************************************************************
  'use this section for custom member variables
  '**********************************************************************************************
  
  Input #fileHandle, TicksPerFrame
  Input #fileHandle, FrameTickCounter
  
  Input #fileHandle, MainAction
  Input #fileHandle, TicksPerMainAction
  Input #fileHandle, MainActionTickCounter
  Input #fileHandle, SubAction
  Input #fileHandle, TicksPerSubAction
  Input #fileHandle, SubActionTickCounter
  
  Input #fileHandle, User1
  
  '**********************************************************************************************
End Sub

