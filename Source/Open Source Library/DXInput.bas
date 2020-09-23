Attribute VB_Name = "DXInput"
'***************************************************************************************************************
'
' DirectX VisualBASIC Interface for Input Support
'                                                     - written by Tim Harpur for Logicon Enterprises
'
' Don't forget to add the appropriate Project->Reference to the DirectX7 library
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
Option Compare Text
Option Base 0

Private dx_DirectX As New DirectX7

'Direct Input variables - all these should be treated as read-only outside of this module
Private dx_DirectInput As DirectInput

Private dx_DirectKeyboard As DirectInputDevice
Private dx_DirectMouse As DirectInputDevice
Private dx_DirectJoystick(1 To 2) As DirectInputDevice

Public dx_KeyboardState As DIKEYBOARDSTATE
Public dx_MouseState As DIMOUSESTATE
Public dx_ControllerState(1 To 2) As DIJOYSTATE

Private dx_EnumJoysticks As DirectInputEnumDevices

Private windowHandle As Long

Public Type ControllerDesc
  description As String
  
  buttons As Long
  povs As Long
  
  X As Boolean
  Y As Boolean
  z As Boolean
  
  deadzone_x As Long
  deadzone_y As Long
  deadzone_z As Long
  
  saturation_x As Long
  saturation_y As Long
  saturation_z As Long
  
  range_xMin As Long
  range_xMax As Long
  range_yMin As Long
  range_yMax As Long
  range_zMin As Long
  range_zMax As Long
  
  rx As Boolean
  ry As Boolean
  rz As Boolean
  
  deadzone_rx As Long
  deadzone_ry As Long
  deadzone_rz As Long
  
  saturation_rx As Long
  saturation_ry As Long
  saturation_rz As Long
  
  range_rxMin As Long
  range_rxMax As Long
  range_ryMin As Long
  range_ryMax As Long
  range_rzMin As Long
  range_rzMax As Long
  
  slider0 As Boolean
  slider1 As Boolean
End Type

Public dx_ControllerDesc(1 To 2) As ControllerDesc

Public Enum InputMODE
  IM_Foreground = 1
  IM_Background
  IM_ForegroundExclusive
End Enum

'Initialize Direct Input for use with system keyboard, system mouse and any attached joysticks/controllers
Public Sub Init_DXInput(parentForm As Object) ', Optional ByVal keyboardMode As InputMODE = IM_Foreground, Optional ByVal mouseMode As InputMODE = IM_Foreground, Optional ByVal activeController1 As Long = 0, Optional ByVal activeController2 As Long = 0)
  Dim dx_DirectMouse_Property As DIPROPLONG
  
  On Error Resume Next
  
  CleanUp_DXInput
  
  windowHandle = parentForm.hWnd
  
  Set dx_DirectInput = dx_DirectX.DirectInputCreate
  
  Set dx_DirectKeyboard = dx_DirectInput.CreateDevice("GUID_SysKeyboard")
  dx_DirectKeyboard.SetCommonDataFormat DIFORMAT_KEYBOARD
  
  Set dx_DirectMouse = dx_DirectInput.CreateDevice("GUID_SysMouse")
  dx_DirectMouse.SetCommonDataFormat DIFORMAT_MOUSE
  
  With dx_DirectMouse_Property
    .lData = DIPROPAXISMODE_REL
    .lHow = DIPH_DEVICE
    .lObj = 0
    .lSize = Len(dx_DirectMouse_Property)
    
    dx_DirectMouse.SetProperty "DIPROP_AXISMODE", dx_DirectMouse_Property
  
    .lData = DIPROPAXISMODE_REL
    .lHow = DIPH_DEVICE
    .lObj = 0
    .lSize = Len(dx_DirectMouse_Property)
    
    dx_DirectMouse.SetProperty "DIPROP_AXISMODE", dx_DirectMouse_Property
  End With
  
  'enumerate attached joysticks
  Set dx_EnumJoysticks = dx_DirectInput.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
End Sub
 
'Acquire keyboard in selected mode
Public Sub Acquire_Keyboard(Optional ByVal accessMode As InputMODE = IM_Foreground)
  On Error GoTo badKeyMode
  
  Select Case accessMode
    Case IM_ForegroundExclusive
      dx_DirectMouse.SetCooperativeLevel windowHandle, DISCL_EXCLUSIVE Or DISCL_FOREGROUND
    Case IM_Background
      dx_DirectMouse.SetCooperativeLevel windowHandle, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
  End Select
  
resumeKeyboard:
  On Error Resume Next
  
  If accessMode = IM_Foreground Then dx_DirectKeyboard.SetCooperativeLevel windowHandle, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND
  
  dx_DirectKeyboard.Acquire
  
  Exit Sub
  
badKeyMode:
  accessMode = IM_Foreground
  
  Resume resumeKeyboard
End Sub
  
'Acquire mouse in selected mode
Public Sub Acquire_Mouse(Optional ByVal accessMode As InputMODE = IM_Foreground)
  On Error GoTo badMouseMode
  
  Select Case accessMode
    Case IM_ForegroundExclusive
      dx_DirectMouse.SetCooperativeLevel windowHandle, DISCL_EXCLUSIVE Or DISCL_FOREGROUND
    Case IM_Background
      dx_DirectMouse.SetCooperativeLevel windowHandle, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
  End Select
  
resumeMouse:
  On Error Resume Next
  
  If accessMode = IM_Foreground Then dx_DirectMouse.SetCooperativeLevel windowHandle, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND
  
  dx_DirectMouse.Acquire
  
  Exit Sub
  
badMouseMode:
  accessMode = IM_Foreground
  
  Resume resumeMouse
End Sub

'select available joystick from enumerated list
Public Sub Acquire_Controller(Optional ByVal controllerNum As Long = 1, Optional ByVal deviceNum As Long = 1)
  Dim didoEnum As DirectInputEnumDeviceObjects, dido As DirectInputDeviceObjectInstance, loop1 As Long
  Dim joystickCaps As DIDEVCAPS, accessMode As Long
  
  On Error Resume Next
  
  If Not (dx_DirectJoystick(controllerNum) Is Nothing) Then
    dx_DirectJoystick(controllerNum).Unacquire
    Set dx_DirectJoystick(controllerNum) = Nothing
  End If
  
  If deviceNum > 0 And deviceNum <= dx_EnumJoysticks.GetCount() Then
    Set dx_DirectJoystick(controllerNum) = dx_DirectInput.CreateDevice(dx_EnumJoysticks.GetItem(deviceNum).GetGuidInstance)
    
    If dx_DirectJoystick(controllerNum) Is Nothing Then Exit Sub
    
    dx_DirectJoystick(controllerNum).SetCommonDataFormat DIFORMAT_JOYSTICK
    
    On Error GoTo badMode
    
    accessMode = IM_Background
    
    dx_DirectJoystick(controllerNum).SetCooperativeLevel windowHandle, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    
resumeController:
    On Error Resume Next
    
    If accessMode = IM_Foreground Then dx_DirectJoystick(controllerNum).SetCooperativeLevel windowHandle, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND
    
    dx_DirectJoystick(controllerNum).GetCapabilities joystickCaps
    
    With dx_ControllerDesc(controllerNum)
      .description = dx_EnumJoysticks.GetItem(deviceNum).GetInstanceName
      
      .X = False
      .Y = False
      .z = False
      .rx = False
      .ry = False
      .rz = False
      .slider0 = False
      .slider1 = False
      
      Set didoEnum = dx_DirectJoystick(controllerNum).GetDeviceObjectsEnum(DIDFT_AXIS)
      
      For loop1 = 1 To didoEnum.GetCount
        Set dido = didoEnum.GetItem(loop1)
        
        Select Case dido.GetOfs
          Case DIJOFS_X
            .X = True
          Case DIJOFS_Y
            .Y = True
          Case DIJOFS_Z
            .z = True
          Case DIJOFS_RX
            .rx = True
          Case DIJOFS_RY
            .ry = True
          Case DIJOFS_RZ
            .rz = True
          Case DIJOFS_SLIDER0
            .slider0 = True
          Case DIJOFS_SLIDER1
            .slider1 = True
        End Select
      Next loop1
    
      .buttons = joystickCaps.lButtons
      .povs = joystickCaps.lPOVs
    End With
    
    Dim DiProp_Abs As DIPROPLONG
  
    With DiProp_Abs
      .lData = DIPROPAXISMODE_ABS
      .lSize = Len(DiProp_Abs)
      .lHow = DIPH_DEVICE
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_AXISMODE", DiProp_Abs
    End With
    
    dx_DirectJoystick(controllerNum).Acquire
    
    Get_ControllerDeadZoneSat controllerNum
    Get_ControllerRange controllerNum
  End If
  
  Exit Sub
  
badMode:
  accessMode = IM_Foreground
  
  Resume resumeController
End Sub

'gets description of joystick from enumerated list
Public Function Get_ControllerDescription(Optional ByVal deviceNum As Long = 1) As String
  On Error Resume Next
  
  If dx_DirectInput Is Nothing Then
    Set dx_DirectInput = dx_DirectX.DirectInputCreate
    
    Set dx_EnumJoysticks = dx_DirectInput.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
  End If
  
  Get_ControllerDescription = dx_EnumJoysticks.GetItem(deviceNum).GetInstanceName
End Function

'gets joystick matching passed description from enumerated list
Public Function Get_ControllerNumber(deviceDescription As String) As Long
  Dim loop1 As Long
  
  On Error Resume Next
  
  If dx_DirectInput Is Nothing Then
    Set dx_DirectInput = dx_DirectX.DirectInputCreate
    
    Set dx_EnumJoysticks = dx_DirectInput.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
  End If
  
  For loop1 = 1 To dx_EnumJoysticks.GetCount()
    If deviceDescription = dx_EnumJoysticks.GetItem(loop1).GetInstanceName Then
      Get_ControllerNumber = loop1
      
      Exit Function
    End If
  Next loop1
  
  Get_ControllerNumber = 0
End Function

'returns the number of attached controllers
Public Function Get_ControllerCount() As Long
  On Error Resume Next
  
  If dx_DirectInput Is Nothing Then
    Set dx_DirectInput = dx_DirectX.DirectInputCreate
    
    Set dx_EnumJoysticks = dx_DirectInput.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
  End If
  
  Get_ControllerCount = dx_EnumJoysticks.GetCount()
End Function

'Release all Direct Input control
Public Sub CleanUp_DXInput()
  On Error Resume Next
  
  If Not (dx_DirectKeyboard Is Nothing) Then
    dx_DirectKeyboard.Unacquire
    Set dx_DirectKeyboard = Nothing
  End If
  
  If Not (dx_DirectMouse Is Nothing) Then
    dx_DirectMouse.Unacquire
    Set dx_DirectMouse = Nothing
  End If
  
  If Not (dx_DirectJoystick(1) Is Nothing) Then
    dx_DirectJoystick(1).Unacquire
    Set dx_DirectJoystick(1) = Nothing
  End If
  
  If Not (dx_DirectJoystick(2) Is Nothing) Then
    dx_DirectJoystick(2).Unacquire
    Set dx_DirectJoystick(2) = Nothing
  End If
  
  Set dx_EnumJoysticks = Nothing
  Set dx_DirectInput = Nothing
End Sub

'Sets active joystick's dead zone and saturation for selected axis
Public Sub Set_ControllerDeadZoneSat(Optional ByVal controllerNum As Long = 1, Optional ByVal deadZone As Long = 200, Optional ByVal Saturation As Long = 8000, Optional ByVal Apply_x As Boolean = False, Optional ByVal Apply_y As Boolean = False, Optional ByVal Apply_z As Boolean = False, Optional ByVal Apply_rx As Boolean = False, Optional ByVal Apply_ry As Boolean = False, Optional ByVal Apply_rz As Boolean = False)
  Dim DiProp_Dead As DIPROPLONG
  
  On Error Resume Next
  
  If dx_DirectJoystick(controllerNum) Is Nothing Then Exit Sub
  
  dx_DirectJoystick(controllerNum).Unacquire
  
  With DiProp_Dead
    .lData = deadZone
    .lSize = Len(DiProp_Dead)
    .lHow = DIPH_BYOFFSET
    
    If Apply_x And dx_ControllerDesc(controllerNum).X Then
      .lObj = DIJOFS_X
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    If Apply_y And dx_ControllerDesc(controllerNum).Y Then
      .lObj = DIJOFS_Y
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    If Apply_z And dx_ControllerDesc(controllerNum).z Then
      .lObj = DIJOFS_Z
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    If Apply_rx And dx_ControllerDesc(controllerNum).rx Then
      .lObj = DIJOFS_RX
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    If Apply_ry And dx_ControllerDesc(controllerNum).ry Then
      .lObj = DIJOFS_RY
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    If Apply_rz And dx_ControllerDesc(controllerNum).rz Then
      .lObj = DIJOFS_RZ
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    .lData = Saturation
    .lSize = Len(DiProp_Dead)
    .lHow = DIPH_BYOFFSET
    
    If Apply_x And dx_ControllerDesc(controllerNum).X Then
      .lObj = DIJOFS_X
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
    
    If Apply_y And dx_ControllerDesc(controllerNum).Y Then
      .lObj = DIJOFS_Y
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
    
    If Apply_z And dx_ControllerDesc(controllerNum).z Then
      .lObj = DIJOFS_Z
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
    
    If Apply_rx And dx_ControllerDesc(controllerNum).rx Then
      .lObj = DIJOFS_RX
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
    
    If Apply_ry And dx_ControllerDesc(controllerNum).ry Then
      .lObj = DIJOFS_RY
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
    
    If Apply_rz And dx_ControllerDesc(controllerNum).rz Then
      .lObj = DIJOFS_RZ
      
      dx_DirectJoystick(controllerNum).SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
  End With
  
  dx_DirectJoystick(controllerNum).Acquire
  
  Get_ControllerDeadZoneSat controllerNum 'get the actual values back
End Sub

'Sets active joystick's range for selected axis
Public Sub Set_ControllerRange(Optional ByVal controllerNum As Long = 1, Optional ByVal rangeMin As Long = -10000, Optional ByVal rangeMax As Long = 10000, Optional Apply_x As Boolean = False, Optional ByVal Apply_y As Boolean = False, Optional Apply_z As Boolean = False, Optional ByVal Apply_rx As Boolean = False, Optional ByVal Apply_ry As Boolean = False, Optional ByVal Apply_rz As Boolean = False)
  Dim DiProp_Range As DIPROPRANGE
 
  On Error Resume Next
  
  If dx_DirectJoystick(controllerNum) Is Nothing Then Exit Sub
  
  dx_DirectJoystick(controllerNum).Unacquire
  
  With DiProp_Range
    .lMin = rangeMin
    .lMax = rangeMax
    .lSize = Len(DiProp_Range)
    .lHow = DIPH_BYOFFSET
    
    If Apply_x And dx_ControllerDesc(controllerNum).X Then
      .lObj = DIJOFS_X
      
      dx_DirectJoystick(controllerNum).SetProperty "DiProp_Range", DiProp_Range
    End If
    
    If Apply_y And dx_ControllerDesc(controllerNum).Y Then
      .lObj = DIJOFS_Y
      
      dx_DirectJoystick(controllerNum).SetProperty "DiProp_Range", DiProp_Range
    End If
    
    If Apply_z And dx_ControllerDesc(controllerNum).z Then
      .lObj = DIJOFS_Z
      
      dx_DirectJoystick(controllerNum).SetProperty "DiProp_Range", DiProp_Range
    End If
    
    If Apply_rx And dx_ControllerDesc(controllerNum).rx Then
      .lObj = DIJOFS_RX
      
      dx_DirectJoystick(controllerNum).SetProperty "DiProp_Range", DiProp_Range
    End If
    
    If Apply_ry And dx_ControllerDesc(controllerNum).ry Then
      .lObj = DIJOFS_RY
      
      dx_DirectJoystick(controllerNum).SetProperty "DiProp_Range", DiProp_Range
    End If
    
    If Apply_rz And dx_ControllerDesc(controllerNum).rz Then
      .lObj = DIJOFS_RZ
      
      dx_DirectJoystick(controllerNum).SetProperty "DiProp_Range", DiProp_Range
    End If
  End With
  
  dx_DirectJoystick(controllerNum).Acquire
  
  Get_ControllerRange controllerNum 'get the actual values back
End Sub

'Gets active joystick's deadzones and saturation levels for selected axis and loads into dx_ControllerDesc
Private Sub Get_ControllerDeadZoneSat(Optional ByVal controllerNum As Long = 1)
  Dim DiProp_Dead As DIPROPLONG
  
  On Error Resume Next
  
  If dx_DirectJoystick(controllerNum) Is Nothing Then Exit Sub
  
  With DiProp_Dead
    .lSize = Len(DiProp_Dead)
    .lHow = DIPH_BYOFFSET
    
    If dx_ControllerDesc(controllerNum).X Then
      .lObj = DIJOFS_X
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).deadzone_x = .lData
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).saturation_x = .lData
    End If
    
    If dx_ControllerDesc(controllerNum).Y Then
      .lObj = DIJOFS_Y
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).deadzone_y = .lData
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).saturation_y = .lData
    End If
    
    If dx_ControllerDesc(controllerNum).z Then
      .lObj = DIJOFS_Z
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).deadzone_z = .lData
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).saturation_z = .lData
    End If
    
    If dx_ControllerDesc(controllerNum).rx Then
      .lObj = DIJOFS_RX
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).deadzone_rx = .lData
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).saturation_rx = .lData
    End If
    
    If dx_ControllerDesc(controllerNum).ry Then
      .lObj = DIJOFS_RY
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).deadzone_ry = .lData
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).saturation_ry = .lData
    End If
    
    If dx_ControllerDesc(controllerNum).rz Then
      .lObj = DIJOFS_RZ
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).deadzone_rz = .lData
      
      dx_DirectJoystick(controllerNum).GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_ControllerDesc(controllerNum).saturation_rz = .lData
    End If
  End With
End Sub

'Gets active joystick's ranges for selected axis and loads into dx_ControllerDesc
Private Sub Get_ControllerRange(Optional ByVal controllerNum As Long = 1)
  Dim DiProp_Range As DIPROPRANGE
 
 On Error Resume Next
 
 If dx_DirectJoystick(controllerNum) Is Nothing Then Exit Sub
 
  With DiProp_Range
    .lSize = Len(DiProp_Range)
    .lHow = DIPH_BYOFFSET
    
    If dx_ControllerDesc(controllerNum).X Then
      .lObj = DIJOFS_X
      
      dx_DirectJoystick(controllerNum).GetProperty "DiProp_Range", DiProp_Range
      
      dx_ControllerDesc(controllerNum).range_xMin = .lMin
      dx_ControllerDesc(controllerNum).range_xMax = .lMax
    End If
    
    If dx_ControllerDesc(controllerNum).Y Then
      .lObj = DIJOFS_Y
      
      dx_DirectJoystick(controllerNum).GetProperty "DiProp_Range", DiProp_Range
      
      dx_ControllerDesc(controllerNum).range_yMin = .lMin
      dx_ControllerDesc(controllerNum).range_yMax = .lMax
    End If
    
    If dx_ControllerDesc(controllerNum).z Then
      .lObj = DIJOFS_Z
      
      dx_DirectJoystick(controllerNum).GetProperty "DiProp_Range", DiProp_Range
      
      dx_ControllerDesc(controllerNum).range_zMin = .lMin
      dx_ControllerDesc(controllerNum).range_zMax = .lMax
    End If
    
    If dx_ControllerDesc(controllerNum).rx Then
      .lObj = DIJOFS_RX
      
      dx_DirectJoystick(controllerNum).GetProperty "DiProp_Range", DiProp_Range
      
      dx_ControllerDesc(controllerNum).range_rxMin = .lMin
      dx_ControllerDesc(controllerNum).range_rxMax = .lMax
    End If
    
    If dx_ControllerDesc(controllerNum).ry Then
      .lObj = DIJOFS_RY
      
      dx_DirectJoystick(controllerNum).GetProperty "DiProp_Range", DiProp_Range
      
      dx_ControllerDesc(controllerNum).range_ryMin = .lMin
      dx_ControllerDesc(controllerNum).range_ryMax = .lMax
    End If
    
    If dx_ControllerDesc(controllerNum).rz Then
      .lObj = DIJOFS_RZ
      
      dx_DirectJoystick(controllerNum).GetProperty "DiProp_Range", DiProp_Range
      
      dx_ControllerDesc(controllerNum).range_rzMin = .lMin
      dx_ControllerDesc(controllerNum).range_rzMax = .lMax
    End If
  End With
End Sub

'Get immediate device state for keyboard
Public Sub Poll_Keyboard()
  On Error GoTo lostAcquire
  
  dx_DirectKeyboard.Poll 'this is usually not nescessary with the keyboard but its so quick it doesn't hurt
  dx_DirectKeyboard.GetDeviceStateKeyboard dx_KeyboardState
  
  Exit Sub

reacquire:
  On Error Resume Next
  
  dx_DirectKeyboard.Acquire
  dx_DirectKeyboard.Poll
  dx_DirectKeyboard.GetDeviceStateKeyboard dx_KeyboardState
  
  Exit Sub
  
lostAcquire:
  Resume reacquire
End Sub

'Get immediate device state for mouse
Public Sub Poll_Mouse()
  On Error GoTo lostAcquire
  
  dx_DirectMouse.Poll 'this is usually not nescessary with the mouse but its so quick it doesn't hurt
  dx_DirectMouse.GetDeviceStateMouse dx_MouseState
  
  Exit Sub

reacquire:
  On Error Resume Next
  
  dx_DirectMouse.Acquire
  dx_DirectMouse.Poll
  dx_DirectMouse.GetDeviceStateMouse dx_MouseState
  
  Exit Sub
  
lostAcquire:
  Resume reacquire
End Sub

'Get immediate device state for joystick
Public Sub Poll_Controller(Optional ByVal controllerNum As Long = 1)
  On Error GoTo lostAcquire
  
  If dx_DirectJoystick(controllerNum) Is Nothing Then Exit Sub
  
  With dx_DirectJoystick(controllerNum)
    .Poll
    .GetDeviceStateJoystick dx_ControllerState(controllerNum)
  End With
  
  Exit Sub

reacquire:
  On Error Resume Next
  
  With dx_DirectJoystick(controllerNum)
    .Acquire
    .Poll
    .GetDeviceStateJoystick dx_ControllerState(controllerNum)
  End With
  
  Exit Sub
  
lostAcquire:
  Resume reacquire
End Sub

Public Function GetDirectInput() As DirectInput
  Set GetDirectInput = dx_DirectInput
End Function

