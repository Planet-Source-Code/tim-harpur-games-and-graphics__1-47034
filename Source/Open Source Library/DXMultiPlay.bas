Attribute VB_Name = "DXMultiPlay"
'***************************************************************************************************************
'
' DirectX VisualBASIC Interface for DirectPlay MultiPlayer Support
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
Private dx_DirectPlay As DirectPlay4
Private dx_DirectPlayConnections As DirectPlayEnumConnections
Private dx_DirectPlaySessions As DirectPlayEnumSessions

Private m_Connection As Long
Private m_SessionActive As Boolean
Private m_PlayerID As Long, m_PlayerName As String, m_PlayerHandle As String

Public Type DXMP_SystemMessage
  messageType As CONST_DPSYSMSGTYPES
  playerType As CONST_DPPLAYERTYPEFLAGS
  
  parentID As Long
  playerID As Long
End Type

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNull Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Function Get_ComputerName() As String
  Dim hKey As Long, sValue As String, vSize As Long, vType As Long
  
  On Error GoTo badKey
  
  If RegOpenKeyEx(&H80000002, "System\CurrentControlSet\Control\ComputerName\ComputerName", 0, &H3F, hKey) = 0 Then
    If RegQueryValueExNull(hKey, "ComputerName", 0&, vType, 0&, vSize) = 0 Then
      If vType = 1 Then
        sValue = String(vSize, 0)
        
        If RegQueryValueExString(hKey, "ComputerName", 0&, vType, sValue, vSize) = 0 Then Get_ComputerName = Left$(sValue, vSize - 1)
      End If
    End If
    
    RegCloseKey hKey
  End If
  
badKey:
End Function

Public Sub Init_DXMultiPlay()
  On Error GoTo noDPlay
  
  CleanUp_DXMultiPlay
  
  Set dx_DirectPlay = dx_DirectX.DirectPlayCreate("")
  
noDPlay:
  
End Sub

Public Function Refresh_AvailableConnections() As Long
  On Error GoTo noConnections
  
  Set dx_DirectPlayConnections = dx_DirectPlay.GetDPEnumConnections("", DPCONNECTION_DIRECTPLAY)
  
  Refresh_AvailableConnections = dx_DirectPlayConnections.GetCount()
  
noConnections:
  
End Function

Public Function Get_ConnectionInfo(ByVal connectionNum) As String
  On Error GoTo noConnections
  
  Get_ConnectionInfo = dx_DirectPlayConnections.GetName(connectionNum)
  
noConnections:
  
End Function

Public Function Get_ActiveConnectionInfo() As String
  On Error GoTo noConnections
  
  If m_Connection > 0 Then Get_ActiveConnectionInfo = dx_DirectPlayConnections.GetName(m_Connection)
  
noConnections:
  
End Function

Public Function Set_ActiveConnection(ByVal connectionNum As Long) As Long
  Dim connectAddress As DirectPlayAddress
  
  On Error GoTo failedToConnect
  
  If m_Connection > 0 Then
    Init_DXMultiPlay 'must re-initialize DirectPlay to change connections
    
    Set dx_DirectPlayConnections = dx_DirectPlay.GetDPEnumConnections("", DPCONNECTION_DIRECTPLAY)
  End If
  
  Set connectAddress = dx_DirectPlayConnections.GetAddress(connectionNum)
  
  dx_DirectPlay.InitializeConnection connectAddress
  
  m_Connection = connectionNum
  Set_ActiveConnection = connectionNum

failedToConnect:
  
End Function

Public Function Refresh_AvailableSessions(Optional ByVal applicationGUID As String = "") As Long
  Dim mSessionData As DirectPlaySessionData
  
  On Error GoTo noSessions
  
  'cannot refresh available sessions if not connected or if already in a session
  If m_Connection > 0 And m_SessionActive = False Then
    Set mSessionData = dx_DirectPlay.CreateSessionData()
    
    mSessionData.SetGuidApplication applicationGUID
    
    Set dx_DirectPlaySessions = dx_DirectPlay.GetDPEnumSessions(mSessionData, 0, DPENUMSESSIONS_ALL Or DPENUMSESSIONS_ASYNC)
    
    Refresh_AvailableSessions = dx_DirectPlaySessions.GetCount()
    
    Exit Function
  End If
  
noSessions:
  Set dx_DirectPlaySessions = Nothing
End Function

Public Function Get_SessionInfo(ByVal sessionNum As Long, ByRef sessionName As String, ByRef sessionPlayers As Long, _
        ByRef sessionMaxPlayers As Long, ByRef sessionStillOpen As Boolean) As Boolean
  
  Dim mSessionData As DirectPlaySessionData
  
  On Error GoTo noSessions
  
  If Not (dx_DirectPlaySessions Is Nothing) Then
    Set mSessionData = dx_DirectPlaySessions.GetItem(sessionNum)
    
    With mSessionData
      sessionName = .GetSessionName()
      sessionPlayers = .GetCurrentPlayers()
      sessionMaxPlayers = .GetMaxPlayers()
      
      If (.GetFlags() And DPSESSION_NEWPLAYERSDISABLED) = DPSESSION_NEWPLAYERSDISABLED Then
        sessionStillOpen = False
      Else
        sessionStillOpen = True
      End If
    End With
    
    Get_SessionInfo = True
  End If
  
noSessions:
  
End Function

Public Function Get_ActiveSessionInfo(ByRef sessionName As String, ByRef sessionPlayers As Long, _
        ByRef sessionMaxPlayers As Long, ByRef sessionStillOpen As Boolean, _
        ByRef playerID() As Long, ByRef playerName() As String, ByRef playerHandle() As String) As Long
        
  Dim mSessionData As DirectPlaySessionData, mDirectPlayPlayers As DirectPlayEnumPlayers
  Dim loop1 As Long
  
  On Error GoTo noSession
  
  'cannot see active session info until connected and in a session
  If m_Connection > 0 And m_SessionActive = True Then
    Set mSessionData = dx_DirectPlay.CreateSessionData()
    
    dx_DirectPlay.GetSessionDesc mSessionData
    
    With mSessionData
      sessionName = .GetSessionName()
      sessionPlayers = .GetCurrentPlayers()
      sessionMaxPlayers = .GetMaxPlayers()
      
      If (.GetFlags() And DPSESSION_NEWPLAYERSDISABLED) = DPSESSION_NEWPLAYERSDISABLED Then
        sessionStillOpen = False
      Else
        sessionStillOpen = True
      End If
    End With
    
    ReDim playerID(1 To sessionPlayers)
    ReDim playerName(1 To sessionPlayers)
    ReDim playerHandle(1 To sessionPlayers)
    
    On Error Resume Next
    
    Set mDirectPlayPlayers = dx_DirectPlay.GetDPEnumPlayers("", DPENUMPLAYERS_ALL)
    
    With mDirectPlayPlayers
      For loop1 = 1 To sessionPlayers
        playerID(loop1) = .GetDPID(loop1)
        playerName(loop1) = .GetLongName(loop1)
        playerHandle(loop1) = .GetShortName(loop1)
      Next loop1
    End With
    
    Get_ActiveSessionInfo = m_PlayerID
    
    Exit Function
  End If
  
noSession:
  sessionName = ""
  sessionPlayers = 0
  sessionMaxPlayers = 0
  sessionStillOpen = False
  
  Erase playerID
  Erase playerName
  Erase playerHandle
End Function

Public Function Join_ActiveSession(ByVal sessionNum As Long, ByVal playerName As String, ByVal playerHandle As String) As Long
      
  Dim mSessionData As DirectPlaySessionData
  
  On Error GoTo noSession
  
  Leave_ActiveSession
  
  Set mSessionData = dx_DirectPlaySessions.GetItem(sessionNum)
  
  dx_DirectPlay.Open mSessionData, DPOPEN_JOIN
  
  m_PlayerID = dx_DirectPlay.CreatePlayer(playerHandle, playerName, 0, DPPLAYER_DEFAULT)
  
  m_PlayerName = playerName
  m_PlayerHandle = playerHandle
  m_SessionActive = True
  
  Join_ActiveSession = m_PlayerID
  
noSession:
  
End Function

Public Function Create_ActiveSession(ByVal playerName As String, ByVal playerHandle As String, _
          ByVal sessionName As String, ByVal sessionMaxPlayers As Long, _
          Optional ByVal applicationGUID As String = "") As Long
        
  Dim mSessionData As DirectPlaySessionData
  
  On Error GoTo noSession
  
  Leave_ActiveSession
  
  Set mSessionData = dx_DirectPlay.CreateSessionData()
  
  With mSessionData
    .SetGuidApplication applicationGUID
    .SetSessionName sessionName
    .SetMaxPlayers sessionMaxPlayers
    
    .SetFlags DPSESSION_KEEPALIVE Or DPSESSION_MIGRATEHOST Or DPSESSION_DIRECTPLAYPROTOCOL Or DPSESSION_NODATAMESSAGES
  End With
  
  dx_DirectPlay.Open mSessionData, DPOPEN_CREATE
  
  m_PlayerID = dx_DirectPlay.CreatePlayer(playerHandle, playerName, 0, DPPLAYER_DEFAULT)
  
  m_PlayerName = playerName
  m_PlayerHandle = playerHandle
  m_SessionActive = True
  
  Create_ActiveSession = m_PlayerID
  
noSession:
  
End Function

Public Sub Leave_ActiveSession()
  Dim mCount As Long, mSize As Long, toPlayer As Long, fromPlayer As Long
  
  On Error Resume Next
  
  If m_SessionActive Then
    Do 'ensure all messages in queue are sent before removing player from the session
      DoEvents
      
      mCount = 0
      fromPlayer = m_PlayerID
      toPlayer = 0
      
      dx_DirectPlay.GetMessageQueue fromPlayer, toPlayer, DPMESSAGEQUEUE_SEND, mCount, mSize
    Loop While mCount > 0
    
    dx_DirectPlay.Close
    
    m_PlayerID = 0
    m_PlayerName = ""
    m_PlayerHandle = ""
    
    m_SessionActive = False
  End If
End Sub

Public Sub Lock_ActiveSession() 'only works for session host
  Dim mSessionData As DirectPlaySessionData
  
  On Error Resume Next
  
  dx_DirectPlay.GetSessionDesc mSessionData
  
  mSessionData.SetFlags DPSESSION_NEWPLAYERSDISABLED Or DPSESSION_KEEPALIVE Or DPSESSION_MIGRATEHOST Or DPSESSION_DIRECTPLAYPROTOCOL Or DPSESSION_NODATAMESSAGES
  
  dx_DirectPlay.SetSessionDesc mSessionData
End Sub

Public Sub Unlock_ActiveSession() 'only works for session host
  Dim mSessionData As DirectPlaySessionData
  
  On Error Resume Next
  
  dx_DirectPlay.GetSessionDesc mSessionData
  
  mSessionData.SetFlags DPSESSION_KEEPALIVE Or DPSESSION_MIGRATEHOST Or DPSESSION_DIRECTPLAYPROTOCOL Or DPSESSION_NODATAMESSAGES
  
  dx_DirectPlay.SetSessionDesc mSessionData
End Sub

Public Function Get_ActivePlayerInfo(ByRef playerID As Long, ByRef playerName As String, ByRef playerHandle As String) As Boolean
  If m_SessionActive Then
    playerID = m_PlayerID
    playerName = m_PlayerName
    playerHandle = m_PlayerHandle
    
    Get_ActivePlayerInfo = True
  Else
    playerID = 0
    playerName = ""
    playerHandle = ""
  End If
End Function

Public Function Get_PlayerName(ByVal playerID As Long) As String
  On Error Resume Next
  
  Get_PlayerName = dx_DirectPlay.GetPlayerFormalName(playerID)
End Function

Public Function Get_PlayerHandle(ByVal playerID As Long) As String
  On Error Resume Next
  
  Get_PlayerHandle = dx_DirectPlay.GetPlayerFriendlyName(playerID)
End Function

Public Function Get_GroupName(ByVal groupID As Long) As String
  On Error Resume Next
  
  Get_GroupName = dx_DirectPlay.GetGroupLongName(groupID)
End Function

Public Function Get_GroupHandle(ByVal groupID As Long) As String
  On Error Resume Next
  
  Get_GroupHandle = dx_DirectPlay.GetGroupShortName(groupID)
End Function

Public Function Create_Group(ByVal groupName As String, ByVal groupHandle As String) As Long
  On Error Resume Next
  
  Create_Group = dx_DirectPlay.CreateGroup(groupHandle, groupName, DPGROUP_DEFAULT)
End Function

Public Function Destroy_Group(ByVal groupID As Long) As Boolean
  On Error GoTo failed
  
  dx_DirectPlay.DestroyGroup groupID
  
  Destroy_Group = True
  
failed:
  
End Function

Public Function Add_PlayerToGroup(ByVal groupID As Long, ByVal playerID As Long) As Boolean
  On Error GoTo failed
  
  dx_DirectPlay.AddPlayerToGroup groupID, playerID
  
  Add_PlayerToGroup = True
  
failed:
  
End Function

Public Function Remove_PlayerFromGroup(ByVal groupID As Long, ByVal playerID As Long) As Boolean
  On Error GoTo failed
  
  dx_DirectPlay.DeletePlayerFromGroup groupID, playerID
  
  Remove_PlayerFromGroup = True
  
failed:
  
End Function

Public Function Add_GroupToGroup(ByVal groupID As Long, ByVal subGroupID As Long) As Boolean
  On Error GoTo failed
  
  dx_DirectPlay.AddGroupToGroup groupID, subGroupID
  
  Add_GroupToGroup = True
  
failed:
  
End Function

Public Function Remove_GroupFromGroup(ByVal groupID As Long, ByVal subGroupID As Long) As Boolean
  On Error GoTo failed
  
  dx_DirectPlay.DeleteGroupFromGroup groupID, subGroupID
  
  Remove_GroupFromGroup = True
  
failed:
  
End Function

'a value of 0 (DPID_SYSMSG) returned in fromPlayerID indicates this is a system message
Public Function Get_Message(ByRef getMessage As DirectPlayMessage, ByRef fromPlayerID As Long) As Boolean
  Dim mToPlayerID As Long
  
  On Error GoTo noMessages
  
  If dx_DirectPlay.GetMessageCount(m_PlayerID) > 0 Then
    mToPlayerID = m_PlayerID
    
    Set getMessage = dx_DirectPlay.Receive(fromPlayerID, mToPlayerID, DPRECEIVE_TOPLAYER)
    
    Get_Message = True
  End If
  
noMessages:
  
End Function

' this routine should only be called on messages that have been flagged as system messages
Public Function Read_SystemMessage(ByVal sysMessage As DirectPlayMessage) As DXMP_SystemMessage
  On Error Resume Next
  
  With Read_SystemMessage
    .messageType = sysMessage.ReadLong()
    
    Select Case .messageType
      Case DPSYS_ADDGROUPTOGROUP, DPSYS_DELETEGROUPFROMGROUP
        'group has been added or removed from parent group
        .playerType = DPPLAYERTYPE_GROUP
        
        .parentID = sysMessage.ReadLong()
        .playerID = sysMessage.ReadLong()
        
      Case DPSYS_ADDPLAYERTOGROUP, DPSYS_DELETEPLAYERFROMGROUP
        'player has been added or removed from parent group
        .playerType = DPPLAYERTYPE_PLAYER
        
        .parentID = sysMessage.ReadLong()
        .playerID = sysMessage.ReadLong()
        
      Case DPSYS_CREATEPLAYERORGROUP, DPSYS_DESTROYPLAYERORGROUP
        'player or group has been created or destroyed
        .playerType = sysMessage.ReadLong()
        
        .playerID = sysMessage.ReadLong()
        
      'Case DPSYS_HOST
        'computer receiving this message is the new host for the session
        
        'there is no need to process this one here
        
      'Case DPSYS_SESSIONLOST
        'session has been lost - perform emergency recovery and cleanup
        
        'there is no need to process this one here
        
    End Select
  End With
End Function

Public Function Start_Message(ByRef newMessage As DirectPlayMessage) As Boolean
  On Error GoTo failed
  
  Set newMessage = dx_DirectPlay.CreateMessage()
  
  Start_Message = True
  
failed:
  
End Function

'a value of 0 (DPID_ALLPLAYERS) is used to send a message to all players in the session
Public Function Send_Message(ByVal sendMessage As DirectPlayMessage, Optional ByVal toPlayerGroupID As Long = 0) As Boolean
  On Error GoTo failed
 
  dx_DirectPlay.SendEx m_PlayerID, toPlayerGroupID, DPSEND_ASYNC Or DPSEND_GUARANTEED Or DPSEND_NOSENDCOMPLETEMSG, sendMessage, 0, 0, 0
  
  Send_Message = True
  
failed:
  
End Function

Public Sub CleanUp_DXMultiPlay()
  Leave_ActiveSession

  m_Connection = 0
  
  Set dx_DirectPlaySessions = Nothing
  Set dx_DirectPlayConnections = Nothing
  Set dx_DirectPlay = Nothing
End Sub


