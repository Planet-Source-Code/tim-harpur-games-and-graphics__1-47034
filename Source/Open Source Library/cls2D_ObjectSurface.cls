VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls2D_ObjectSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************************
'
' cls2D_ObjectSurface class definition
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

Public ParentObject As cls2D_Object                               ' parent object for this Surface
Public PrevSurface As cls2D_ObjectSurface     ' previous Surface in list
Public NextSurface As cls2D_ObjectSurface     ' next Surface in list

Public ZPriority As Long                   ' z render priority value of this Surface

Public TypeID As Long                   ' object identifier

Public DisplayOffsetX As Long         ' x display offset from 2DObject's center (usually -ve)
Public DisplayOffsetY As Long         ' y display offset from 2DObject's center (usually -ve)
  
Public ActionSequence As Variant    ' array of image indexes for each action frame (0 based array)
Public ActionLast As Long                ' UBOUND of ActionSequence - automatically set by Set_AnimationSequence()
Public ActionFrame As Long            ' current index into the array - determines which frame to display

Public SurfaceIndex As Long                                ' static surface index for image list
Public SurfaceOffsetX As Long                            ' x offset to start of image list on surface
Public SurfaceOffsetY As Long                            ' y offset to start of image list on surface
Public ImageWidth As Long                                  ' width of image
Public ImageHeight As Long                                 ' height of image
Public ImagesPerRow As Long                            ' number of images per row on surface

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
  
  Visible = True
End Sub

Public Function Duplicate_Object(ByVal newObjectID As Long) As cls2D_ObjectSurface
  Set Duplicate_Object = New cls2D_ObjectSurface
  
  With Duplicate_Object
    .Visible = Visible
     
    .ZPriority = ZPriority
     
    .TypeID = TypeID
     
    .DisplayOffsetX = DisplayOffsetX
    .DisplayOffsetY = DisplayOffsetY
     
    .ActionLast = ActionLast
    .ActionSequence = ActionSequence
           
    .ActionFrame = ActionFrame
    
    .SurfaceIndex = SurfaceIndex
    .SurfaceOffsetX = SurfaceOffsetX
    .SurfaceOffsetY = SurfaceOffsetY
    .ImageWidth = ImageWidth
    .ImageHeight = ImageHeight
    .ImagesPerRow = ImagesPerRow
    
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
      
  Write #fileHandle, ActionLast
  
  For loop1 = 0 To ActionLast
    Write #fileHandle, ActionSequence(loop1)
  Next loop1
      
  Write #fileHandle, ActionFrame

  Write #fileHandle, SurfaceIndex
  Write #fileHandle, SurfaceOffsetX
  Write #fileHandle, SurfaceOffsetY
  Write #fileHandle, ImageWidth
  Write #fileHandle, ImageHeight
  Write #fileHandle, ImagesPerRow
  
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
  
  Input #fileHandle, ActionLast
  
  If ActionLast > -1 Then
    ReDim ActionSequence(0 To ActionLast)
    
    For loop1 = 0 To ActionLast
      Input #fileHandle, ActionSequence(loop1)
    Next loop1
  End If
  
  Input #fileHandle, ActionFrame

  Input #fileHandle, SurfaceIndex
  Input #fileHandle, SurfaceOffsetX
  Input #fileHandle, SurfaceOffsetY
  Input #fileHandle, ImageWidth
  Input #fileHandle, ImageHeight
  Input #fileHandle, ImagesPerRow
  
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

