Attribute VB_Name = "DXSound"
'***************************************************************************************************************
'
' DirectX VisualBASIC Interface for Sound & Music Support
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

'Direct Music & Sound variables
Private dx_DirectSound As DirectSound

Private dx_Primary3DSoundBuffer As DirectSoundBuffer
Private dx_Sound3DListener As DirectSound3DListener
Private dx_SoundBuffer() As DirectSoundBuffer
Private dx_Sound3DBuffer() As DirectSound3DBuffer
Private m_BufferIs3D() As Boolean
Private m_SoundBufferFileName() As String
Private m_MasterSoundVolume() As Long
Private m_TotalSoundBuffers As Long
Private m_SoundMuted As Boolean
Private m_SoundPath As String

Private m_TotalMusicChannels As Long
Private dx_DirectMusicLoader As DirectMusicLoader
Private dx_DirectMusicPerformance() As DirectMusicPerformance
Private dx_DirectMusicSegment() As DirectMusicSegment
Private m_MusicChannelFileName() As String
Private m_MasterMusicVolume() As Long
Private m_MusicMuted As Boolean
Private m_MusicPath As String

'Initialize DirectX Sound and Music
Public Sub Init_DXSound(parentForm As Object, Optional ByVal numSoundChannels As Long = 0, _
      Optional ByVal numMusicChannels As Long = 0, Optional ByVal enable3DSound As Boolean = False)
  
  Dim loop1 As Long
  Dim dx_PrimarySoundBufferDesc As DSBUFFERDESC, dx_PrimaryWaveFormat As WAVEFORMATEX

  On Error Resume Next
  
  CleanUp_DXSound
  
  Set dx_DirectSound = dx_DirectX.DirectSoundCreate("")
  dx_DirectSound.SetCooperativeLevel parentForm.hWnd, DSSCL_PRIORITY
  
  m_TotalSoundBuffers = numSoundChannels
  
  If numSoundChannels > 0 Then
    ReDim dx_SoundBuffer(1 To numSoundChannels)
    ReDim dx_Sound3DBuffer(1 To numSoundChannels)
    ReDim m_BufferIs3D(1 To numSoundChannels)
    ReDim m_SoundBufferFileName(1 To numSoundChannels)
    ReDim m_MasterSoundVolume(1 To numSoundChannels)
    
    If enable3DSound Then
      dx_PrimarySoundBufferDesc.lFlags = DSBCAPS_CTRL3D Or DSBCAPS_PRIMARYBUFFER
      
      Set dx_Primary3DSoundBuffer = dx_DirectSound.CreateSoundBuffer(dx_PrimarySoundBufferDesc, dx_PrimaryWaveFormat)
      Set dx_Sound3DListener = dx_Primary3DSoundBuffer.GetDirectSound3DListener()
    End If
  End If
    
  m_SoundMuted = False
  Set_SoundPath "\"
  
  m_TotalMusicChannels = numMusicChannels
  
  If numMusicChannels > 0 Then
    ReDim dx_DirectMusicPerformance(1 To numMusicChannels)
    ReDim dx_DirectMusicSegment(1 To numMusicChannels)
    ReDim m_MusicChannelFileName(1 To numMusicChannels)
    ReDim m_MasterMusicVolume(1 To numSoundChannels)
    
    Set dx_DirectMusicLoader = dx_DirectX.DirectMusicLoaderCreate()
    
    For loop1 = 1 To numMusicChannels
      Set dx_DirectMusicPerformance(loop1) = dx_DirectX.DirectMusicPerformanceCreate()
      dx_DirectMusicPerformance(loop1).Init dx_DirectSound, 0
      
      dx_DirectMusicPerformance(loop1).SetPort -1, 1
      dx_DirectMusicPerformance(loop1).SetMasterVolume 0
    Next loop1
  End If
  
  m_MusicMuted = False
  Set_MusicPath "\"
End Sub

Public Sub Set_MusicPath(ByVal musicFilePath As String)
  If Left$(musicFilePath, 1) = "\" Then musicFilePath = App.Path & musicFilePath
  
  If Right$(musicFilePath, 1) <> "\" Then
    m_MusicPath = musicFilePath & "\"
  Else
    m_MusicPath = musicFilePath
  End If
End Sub

Public Function Get_MusicPath() As String
  Get_MusicPath = m_MusicPath
End Function

Public Function Get_TotalMusicChannels() As Long
  Get_TotalMusicChannels = m_TotalMusicChannels
End Function

Public Function Get_FreeMusicChannel() As Long
  Dim loop1 As Long
  
  For loop1 = 1 To m_TotalMusicChannels
    If m_MusicChannelFileName(loop1) = "" Then
      Get_FreeMusicChannel = loop1
      
      Exit Function
    End If
  Next loop1
  
  Get_FreeMusicChannel = 0
End Function

Public Sub Set_MidiPort(Optional ByVal channelNum As Long = 0, Optional ByVal midiPort As Long = -1)
  Dim loop1 As Long
  
  On Error Resume Next
  
  If channelNum = 0 Then
    UnloadAll_MusicChannels
    
    For loop1 = 1 To m_TotalMusicChannels
      dx_DirectMusicPerformance(loop1).SetPort midiPort, 1
    Next loop1
  Else
    Unload_MusicChannel channelNum
    
    dx_DirectMusicPerformance(channelNum).SetPort midiPort, 1
  End If
End Sub

Public Function Get_TotalMidiPorts() As Long
  On Error Resume Next
  
  Get_TotalMidiPorts = dx_DirectMusicPerformance(1).GetPortCount
End Function

Public Function Get_MidiPortDescription(Optional ByVal midiPort As Long = 1) As String
  On Error GoTo noPort
  
  If midiPort > 0 And midiPort <= dx_DirectMusicPerformance(1).GetPortCount Then
    Get_MidiPortDescription = dx_DirectMusicPerformance(1).GetPortName(midiPort)
    
    Exit Function
  End If
  
noPort:
  Get_MidiPortDescription = ""
End Function

Public Function Get_MidiPortNumber(portDescription As String) As Long
  Dim loop1 As Long
  
  On Error Resume Next
  
  ' just use the first dx_DirectMusicPerformance object as they all return the same information on
  ' what ports are available
  For loop1 = 1 To dx_DirectMusicPerformance(1).GetPortCount
    If portDescription = dx_DirectMusicPerformance(1).GetPortName(loop1) Then
      Get_MidiPortNumber = loop1
      
      Exit Function
    End If
  Next loop1
  
  Get_MidiPortNumber = -1
End Function

Public Sub Play_Music(ByVal channelNum As Long)
  On Error Resume Next
  
  dx_DirectMusicPerformance(channelNum).PlaySegment dx_DirectMusicSegment(channelNum), 0, 0
End Sub

Public Function IsMusicPlaying(ByVal channelNum As Long) As Boolean
  On Error Resume Next
  
  IsMusicPlaying = dx_DirectMusicPerformance(channelNum).IsPlaying(dx_DirectMusicSegment(channelNum), Nothing)
End Function

Public Sub Stop_Music(ByVal channelNum As Long)
  On Error Resume Next
  
  dx_DirectMusicPerformance(channelNum).Stop Nothing, Nothing, 0, 0
End Sub

Public Sub StopAll_Music()
  Dim loop1 As Long
  
  On Error Resume Next
  
  For loop1 = 1 To m_TotalMusicChannels
    dx_DirectMusicPerformance(loop1).Stop Nothing, Nothing, 0, 0
  Next loop1
End Sub

Public Sub MuteAll_Music()
  Dim loop1 As Long
  
  On Error Resume Next
  
  If m_MusicMuted = True Then Exit Sub
  
  For loop1 = 1 To m_TotalMusicChannels
    m_MasterMusicVolume(loop1) = dx_DirectMusicPerformance(loop1).GetMasterVolume
    
    dx_DirectMusicPerformance(loop1).SetMasterVolume -10000
  Next loop1
  
  m_MusicMuted = True
End Sub

Public Sub UnmuteAll_Music()
  Dim loop1 As Long
  
  On Error Resume Next
  
  If m_MusicMuted = False Then Exit Sub
  
  For loop1 = 1 To m_TotalMusicChannels
    dx_DirectMusicPerformance(loop1).SetMasterVolume m_MasterMusicVolume(loop1)
  Next loop1
  
  m_MusicMuted = False
End Sub

Public Sub Load_MusicFromMidi(ByVal channelNum As Long, midiFileName As String)
  On Error GoTo badLoad
  
  Unload_MusicChannel channelNum
  
  Set dx_DirectMusicSegment(channelNum) = dx_DirectMusicLoader.LoadSegment(m_MusicPath & midiFileName)
  
  With dx_DirectMusicSegment(channelNum)
    .SetStandardMidiFile
    .Download dx_DirectMusicPerformance(channelNum)
  End With
  
  m_MusicChannelFileName(channelNum) = midiFileName
  
  Exit Sub
  
badLoad:
  Unload_MusicChannel channelNum
End Sub

'Return Music Channel fileName
Public Function Get_MusicChannelFileName(ByVal channelNum As Long) As String
  On Error Resume Next
  
  Get_MusicChannelFileName = m_MusicChannelFileName(channelNum)
End Function

Public Sub Change_MusicSettings(ByVal channelNum As Long, Optional ByVal Volume As Variant, Optional ByVal Tempo As Variant)
  On Error Resume Next
  
  'Set any parameters
  'Volume adjustment, in hundredths of a decibel -range is port specific
  'Tempo is from (.25) to (2) with (1) being the default
  If Not IsMissing(Volume) Then
    If m_MusicMuted Then 'if music is muted then just store new value until unmuted
      m_MasterMusicVolume(channelNum) = Volume
    Else
      dx_DirectMusicPerformance(channelNum).SetMasterVolume Volume
    End If
  End If
  
  If Not IsMissing(Tempo) Then dx_DirectMusicPerformance(channelNum).SetMasterTempo Tempo
End Sub

Public Sub Unload_MusicChannel(ByVal channelNum As Long)
  On Error Resume Next
  
  dx_DirectMusicPerformance(channelNum).Stop Nothing, Nothing, 0, 0
  dx_DirectMusicSegment(channelNum).Unload dx_DirectMusicPerformance(channelNum)
  
  m_MusicChannelFileName(channelNum) = ""
End Sub

Public Sub UnloadAll_MusicChannels()
  Dim loop1 As Long
  
  On Error Resume Next
  
  For loop1 = 1 To m_TotalMusicChannels
    dx_DirectMusicPerformance(loop1).Stop Nothing, Nothing, 0, 0
    dx_DirectMusicSegment(loop1).Unload dx_DirectMusicPerformance(loop1)
    
    m_MusicChannelFileName(loop1) = ""
  Next loop1
End Sub

Public Sub Set_SoundPath(ByVal soundFilePath As String)
  If Left$(soundFilePath, 1) = "\" Then soundFilePath = App.Path & soundFilePath
  
  If Right$(soundFilePath, 1) <> "\" Then
    m_SoundPath = soundFilePath & "\"
  Else
    m_SoundPath = soundFilePath
  End If
End Sub

Public Function Get_SoundPath() As String
  Get_SoundPath = m_SoundPath
End Function

Public Function Get_TotalSoundBuffers() As Long
  Get_TotalSoundBuffers = m_TotalSoundBuffers
End Function

Public Function Get_FreeSoundBuffer() As Long
  Dim loop1 As Long
  
  For loop1 = 1 To m_TotalSoundBuffers
    If m_SoundBufferFileName(loop1) = "" Then
      Get_FreeSoundBuffer = loop1
      
      Exit Function
    End If
  Next loop1
  
  Get_FreeSoundBuffer = 0
End Function

'Create the sound buffer from file
Public Sub Load_SoundBuffer(ByVal channelNum As Long, soundFile As String)
  Dim dx_SoundBufferDesc As DSBUFFERDESC, dx_WaveFormat As WAVEFORMATEX
  
  On Error GoTo badLoad
  
  If channelNum > 0 And channelNum <= m_TotalSoundBuffers Then
    Unload_SoundBuffer channelNum
    
    dx_SoundBufferDesc.lFlags = DSBCAPS_STICKYFOCUS Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_STATIC
        
    Set dx_SoundBuffer(channelNum) = dx_DirectSound.CreateSoundBufferFromFile(m_SoundPath & soundFile, dx_SoundBufferDesc, dx_WaveFormat)
    
    m_SoundBufferFileName(channelNum) = soundFile
    m_BufferIs3D(channelNum) = False
    
    If m_SoundMuted = True Then 'if sound is muted then mute new buffer too
      m_MasterSoundVolume(channelNum) = dx_SoundBuffer(channelNum).GetVolume
      
      dx_SoundBuffer(channelNum).SetVolume -10000
    End If
  End If
  
  Exit Sub
  
badLoad:
  Unload_SoundBuffer channelNum
End Sub

'Create the 3D sound buffer from file
Public Sub Load_3DSoundBuffer(ByVal channelNum As Long, soundFile As String)
  Dim dx_SoundBufferDesc As DSBUFFERDESC, dx_WaveFormat As WAVEFORMATEX
  
  On Error GoTo badLoad
  
  If channelNum > 0 And channelNum <= m_TotalSoundBuffers Then
    Unload_SoundBuffer channelNum
    
    dx_SoundBufferDesc.lFlags = DSBCAPS_STICKYFOCUS Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_STATIC
        
    Set dx_SoundBuffer(channelNum) = dx_DirectSound.CreateSoundBufferFromFile(m_SoundPath & soundFile, dx_SoundBufferDesc, dx_WaveFormat)
    Set dx_Sound3DBuffer(channelNum) = dx_SoundBuffer(channelNum).GetDirectSound3DBuffer
    
    m_SoundBufferFileName(channelNum) = soundFile
    m_BufferIs3D(channelNum) = True
    
    If m_SoundMuted = True Then 'if sound is muted then mute new buffer too
      m_MasterSoundVolume(channelNum) = dx_SoundBuffer(channelNum).GetVolume
      
      dx_SoundBuffer(channelNum).SetVolume -10000
    End If
  End If
  
  Exit Sub
  
badLoad:
  Unload_SoundBuffer channelNum
End Sub

'Return Sound Buffer fileName
Public Function Get_SoundBufferFileName(ByVal channelNum As Long) As String
  On Error Resume Next
  
  Get_SoundBufferFileName = m_SoundBufferFileName(channelNum)
End Function

'Create the sound buffer - duplicate existing
Public Sub Duplicate_SoundBuffer(ByVal channelNum As Long, ByVal channelNumSource As Long)
  On Error GoTo badLoad
  
  If channelNum <> channelNumSource Then 'don't allow any attempt to duplicate into self
    Unload_SoundBuffer channelNum
    
    Set dx_SoundBuffer(channelNum) = dx_DirectSound.DuplicateSoundBuffer(dx_SoundBuffer(channelNumSource))
    If m_BufferIs3D(channelNumSource) Then Set dx_Sound3DBuffer(channelNum) = dx_SoundBuffer(channelNum).GetDirectSound3DBuffer
    
    m_SoundBufferFileName(channelNum) = m_SoundBufferFileName(channelNumSource)
    m_BufferIs3D(channelNum) = m_BufferIs3D(channelNumSource)
    m_MasterSoundVolume(channelNum) = m_MasterSoundVolume(channelNumSource)
  End If
  
  Exit Sub
  
badLoad:
  Unload_SoundBuffer channelNum
End Sub

'Plays the selected sound buffer
Public Sub Play_Sound(ByVal channelNum As Long, Optional ByVal loopMode As Boolean = False, Optional ByVal forcedPlay As Boolean = True)
  On Error Resume Next
  
  If forcedPlay = True Then
    dx_SoundBuffer(channelNum).Stop
    dx_SoundBuffer(channelNum).SetCurrentPosition 0
  End If
  
  If loopMode Then
    dx_SoundBuffer(channelNum).Play DSBPLAY_LOOPING 'Play the sound buffer repeating
  Else
    dx_SoundBuffer(channelNum).Play DSBPLAY_DEFAULT 'Play the sound buffer once
  End If
End Sub

'Resumes playing the selected sound buffer
Public Sub Resume_Sound(ByVal channelNum As Long, Optional ByVal loopMode As Boolean = False)
  On Error Resume Next
  
  If loopMode Then
    dx_SoundBuffer(channelNum).Play DSBPLAY_LOOPING 'Play the sound buffer repeating
  Else
    dx_SoundBuffer(channelNum).Play DSBPLAY_DEFAULT 'Play the sound buffer once
  End If
End Sub

'Stops the selected sound buffer
Public Sub Stop_Sound(ByVal channelNum As Long)
  On Error Resume Next
  
  dx_SoundBuffer(channelNum).Stop
End Sub

Public Sub MuteAll_Sound()
  Dim loop1 As Long
  
  On Error Resume Next
  
  If m_SoundMuted = True Then Exit Sub 'this routine can only be called once until sound is unmuted
  
  For loop1 = 1 To m_TotalSoundBuffers
    If Not (dx_SoundBuffer(loop1) Is Nothing) Then
      m_MasterSoundVolume(loop1) = dx_SoundBuffer(loop1).GetVolume
      
      dx_SoundBuffer(loop1).SetVolume -10000
    End If
  Next loop1
  
  m_SoundMuted = True
End Sub

Public Sub UnmuteAll_Sound()
  Dim loop1 As Long
  
  On Error Resume Next
  
  If m_SoundMuted = False Then Exit Sub
  
  For loop1 = 1 To m_TotalSoundBuffers
    If Not (dx_SoundBuffer(loop1) Is Nothing) Then dx_SoundBuffer(loop1).SetVolume m_MasterSoundVolume(loop1)
  Next loop1
  
  m_SoundMuted = False
End Sub

'Stops all sound buffers
Public Sub StopAll_Sound()
  Dim loop1 As Long
  
  On Error Resume Next
  
  For loop1 = 1 To m_TotalSoundBuffers
    If Not (dx_SoundBuffer(loop1) Is Nothing) Then dx_SoundBuffer(loop1).Stop
  Next loop1
End Sub

'Changes the settings for the selected sound buffer - if sound buffer is 3D then panLeftRight is ignored
Public Sub Change_SoundSettings(ByVal channelNum As Long, Optional ByVal playbackFrequency As Variant, Optional ByVal playbackVolume As Variant, Optional ByVal panLeftRight As Variant)
  On Error Resume Next
  
  'Set any dx_SoundBuffer() parameters
  'Frequency is in Hz
  'Volume is rated in 1/100ths of a dB from max (0) to min (-10000)
  'Pan is is rated in reduction in 1/100ths of a dB with -ve reducing left channel and +ve reducing right
  'channel (-10000) to (10000) with (0) being centered
  With dx_SoundBuffer(channelNum)
    If Not IsMissing(playbackFrequency) Then .SetFrequency playbackFrequency
    
    If Not IsMissing(playbackVolume) Then
      If m_SoundMuted = True Then 'if sound is muted then just store new value until unmuted
        m_MasterSoundVolume(channelNum) = playbackVolume
      Else
        .SetVolume playbackVolume
      End If
    End If
    
    If Not IsMissing(panLeftRight) Then .SetPan panLeftRight
  End With
End Sub

Public Function Get_SoundFrequency(ByVal channelNum As Long) As Long
  On Error Resume Next
  
  'get the sound buffer's current playback frequency in Hz
  Get_SoundFrequency = dx_SoundBuffer(channelNum).GetFrequency
End Function

'Changes the (static) settings for the selected 3D sound buffer
Public Sub Change_3DSoundSettings(ByVal channelNum As Long, ByVal minDistance As Single, ByVal maxDistance As Single, _
      Optional ByVal innerConeAngle As Long = 360, Optional ByVal outerConeAngle As Long = 360, Optional ByVal outsideConeAttenuation As Long = 0)
      
  On Error Resume Next
  
  'Set any dx_Sound3DBuffer() parameters other than position/velocity/orientation
  With dx_Sound3DBuffer(channelNum)
    .SetMinDistance minDistance, DS3D_DEFERRED
    .SetMaxDistance maxDistance, DS3D_DEFERRED
    .SetConeAngles innerConeAngle, outerConeAngle, DS3D_DEFERRED
    .SetConeOutsideVolume outsideConeAttenuation, DS3D_DEFERRED
  End With
End Sub

'Changes the (dynamic) position/velocity/orientation settings for the selected 3D sound buffer
Public Sub Change_3DSoundPosition(ByVal channelNum As Long, ByVal xPos As Single, ByVal yPos As Single, ByVal zPos As Single, _
        Optional ByVal xVelocity As Single = 0, Optional ByVal yVelocity As Single = 0, Optional ByVal zVelocity As Single = 0, _
        Optional ByVal xConeVector As Single = 0, Optional ByVal yConeVector As Single = 1, Optional ByVal zConeVector As Single = 0)
        
  On Error Resume Next
  
  With dx_Sound3DBuffer(channelNum)
    .SetPosition xPos, yPos, zPos, DS3D_DEFERRED
    .SetVelocity xVelocity, yVelocity, zVelocity, DS3D_DEFERRED
    .SetConeOrientation xConeVector, yConeVector, zConeVector, DS3D_DEFERRED
  End With
End Sub

'Changes the (dynamic) position/velocity/orientation settings for the 3D listener
Public Sub Change_3DListenerPosition(ByVal xPos As Single, ByVal yPos As Single, ByVal zPos As Single, _
        Optional ByVal xVelocity As Single = 0, Optional ByVal yVelocity As Single = 0, Optional ByVal zVelocity As Single = 0, _
        Optional ByVal xFront As Single = 0, Optional ByVal yFront As Single = 0, Optional ByVal zFront As Single = 0, _
        Optional ByVal xTop As Single = 0, Optional ByVal yTop As Single = 0, Optional ByVal zTop As Single = 0)
        
  On Error Resume Next
  
  With dx_Sound3DListener
    .SetPosition xPos, yPos, zPos, DS3D_DEFERRED
    .SetVelocity xVelocity, yVelocity, zVelocity, DS3D_DEFERRED
    .SetOrientation xFront, yFront, zFront, xTop, yTop, zTop, DS3D_DEFERRED
  End With
End Sub

Public Sub Commit_3DSoundChanges()
  On Error Resume Next
  
  dx_Sound3DListener.CommitDeferredSettings
End Sub

Public Sub Unload_SoundBuffer(ByVal channelNum As Long)
  On Error Resume Next
  
  If Not (dx_SoundBuffer(channelNum) Is Nothing) Then dx_SoundBuffer(channelNum).Stop
  
  Set dx_SoundBuffer(channelNum) = Nothing
  Set dx_Sound3DBuffer(channelNum) = Nothing
  
  m_SoundBufferFileName(channelNum) = ""
End Sub

Public Sub UnloadAll_SoundBuffers()
  Dim loop1 As Long
  
  On Error Resume Next
  
  'Perform for each allocated buffer
  For loop1 = 1 To m_TotalSoundBuffers
    If Not (dx_SoundBuffer(loop1) Is Nothing) Then dx_SoundBuffer(loop1).Stop
    
    Set dx_SoundBuffer(loop1) = Nothing
    Set dx_Sound3DBuffer(loop1) = Nothing
    
    m_SoundBufferFileName(loop1) = ""
  Next loop1
End Sub

'Releases all sound buffers and Direct Sound
Public Sub CleanUp_DXSound()
  Dim loop1 As Long
  
  On Error Resume Next
  
  UnloadAll_SoundBuffers
  
  m_TotalSoundBuffers = 0
  
  Set dx_Sound3DListener = Nothing
  Set dx_Primary3DSoundBuffer = Nothing
  
  UnloadAll_MusicChannels
  
  For loop1 = 1 To m_TotalMusicChannels
    dx_DirectMusicPerformance(loop1).CloseDown
    
    Set dx_DirectMusicPerformance(loop1) = Nothing
    Set dx_DirectMusicSegment(loop1) = Nothing
  Next loop1
  
  Set dx_DirectMusicLoader = Nothing
  m_TotalMusicChannels = 0
  
  Set dx_DirectSound = Nothing
End Sub

Public Function GetDirectSound() As DirectSound
  Set GetDirectSound = dx_DirectSound
End Function

