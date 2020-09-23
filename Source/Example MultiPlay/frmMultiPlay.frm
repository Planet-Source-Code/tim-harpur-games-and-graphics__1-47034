VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMultiPlay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Example MultiPlay"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9900
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMultiPlay.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send All"
      Enabled         =   0   'False
      Height          =   350
      Index           =   1
      Left            =   8250
      TabIndex        =   12
      Top             =   7110
      Width           =   1485
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear Messages"
      Height          =   1425
      Left            =   8340
      TabIndex        =   14
      Top             =   7560
      Width           =   1425
   End
   Begin VB.TextBox txtMessages 
      Height          =   1395
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   7560
      Width           =   8085
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   7140
      TabIndex        =   11
      Top             =   7110
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   350
      Left            =   5160
      TabIndex        =   10
      Top             =   7110
      Width           =   1065
   End
   Begin VB.TextBox txtMessage 
      Height          =   975
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   6060
      Width           =   4575
   End
   Begin VB.Timer tmrSession 
      Interval        =   1000
      Left            =   120
      Top             =   2250
   End
   Begin VB.ListBox lstNames 
      Height          =   1410
      Left            =   150
      TabIndex        =   8
      Top             =   6030
      Width           =   4875
   End
   Begin VB.TextBox txtPlayerHandle 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   6750
      TabIndex        =   2
      Top             =   1890
      Width           =   3000
   End
   Begin VB.TextBox txtPlayerName 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   6750
      TabIndex        =   1
      Top             =   1380
      Width           =   3000
   End
   Begin MSComctlLib.ListView lsvSessions 
      Height          =   1395
      Left            =   150
      TabIndex        =   3
      Top             =   2700
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   2461
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   7938
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Players"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "MaxPlayers"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Open"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.CheckBox chkLocked 
      Caption         =   "Lock Session"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7980
      TabIndex        =   7
      Top             =   4530
      Width           =   1785
   End
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave Session"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7800
      TabIndex        =   6
      Top             =   4140
      Width           =   1965
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Session"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2160
      TabIndex        =   5
      Top             =   4140
      Width           =   1965
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join Session"
      Enabled         =   0   'False
      Height          =   350
      Left            =   150
      TabIndex        =   4
      Top             =   4140
      Width           =   1965
   End
   Begin VB.ListBox lstConnections 
      Height          =   1680
      Left            =   150
      TabIndex        =   0
      Top             =   540
      Width           =   4815
   End
   Begin MSComctlLib.ListView lsvSession 
      Height          =   765
      Left            =   150
      TabIndex        =   20
      Top             =   4890
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   1349
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Players"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "MaxPlayers"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Open"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Session Players"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   23
      Top             =   5730
      Width           =   1875
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Player Handle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5130
      TabIndex        =   22
      Top             =   1950
      Width           =   1605
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Player Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   21
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblConnection 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5070
      TabIndex        =   19
      Top             =   540
      Width           =   4695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Selected MultiPlayer Connection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5580
      TabIndex        =   18
      Top             =   210
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Current Session"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4020
      TabIndex        =   17
      Top             =   4590
      Width           =   1875
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Available MultiPlayer Sessions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3210
      TabIndex        =   16
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Available MultiPlayer Connections"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   15
      Top             =   210
      Width           =   3915
   End
End
Attribute VB_Name = "frmMultiPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public currentSessionName As String, appGUID As String
Public isHost As Boolean, playerID As Long, selectedID As Long

Private Sub chkLocked_Click() 'lock/unlock the active session
  If chkLocked.Value = 0 Then
    DXMultiPlay.Unlock_ActiveSession
  Else
    DXMultiPlay.Lock_ActiveSession
  End If
End Sub

Private Sub cmdClear_Click() 'clear send message
  txtMessage.Text = ""
End Sub

Private Sub cmdClearAll_Click() 'clear receive messages
  txtMessages.Text = ""
End Sub

Private Sub cmdCreate_Click() 'create a new session
  If txtPlayerName.Text = "" Or txtPlayerHandle.Text = "" Then
    MsgBox "You must provide a name and handle first!", vbCritical, "Name/Handle Missing"
  ElseIf txtPlayerName.Text = txtPlayerHandle.Text Then
    MsgBox "Player name and handle must be different!", vbCritical, "Name/Handle Error"
  Else
    'activate the session creation form
    frmCreateSession.Show vbModal
  End If
End Sub

Private Sub cmdJoin_Click() 'join an existing session
  Dim sItem As Long
  
  tmrSession.Enabled = False
  
  If txtPlayerName.Text = "" Or txtPlayerHandle.Text = "" Then
    MsgBox "You must provide a name and handle first!", vbCritical, "Name/Handle Missing"
  ElseIf txtPlayerName.Text = txtPlayerHandle.Text Then
    MsgBox "Player name and handle must be different!", vbCritical, "Name/Handle Error"
  Else
    isHost = False
    
    'ensure that a session has been selected
    If Not (lsvSessions.SelectedItem Is Nothing) Then
      sItem = lsvSessions.SelectedItem.Index
      
      'join the selected session
      playerID = DXMultiPlay.Join_ActiveSession(sItem, txtPlayerName.Text, txtPlayerHandle.Text)
    End If
  End If
  
  tmrSession.Enabled = True
End Sub

Private Sub cmdLeave_Click() 'leave the active session
  isHost = False
  playerID = 0
  
  'leave the active session
  DXMultiPlay.Leave_ActiveSession
End Sub

Private Sub cmdSend_Click(Index As Integer) 'send a message
  Dim mMessage As DirectPlayMessage
  
  'aquire a new DirectPlayMessage object
  If DXMultiPlay.Start_Message(mMessage) Then
    With mMessage 'prepare the message
      .WriteString txtPlayerName.Text
      .WriteString txtMessage.Text
    End With
    
    If Index = 0 Then 'send the message to the selected player
      DXMultiPlay.Send_Message mMessage, selectedID
    Else 'send the message to all players
      DXMultiPlay.Send_Message mMessage
    End If
  End If
End Sub

Private Sub Form_Load()
  Dim loop1 As Long, maxCount As Long, conNames() As String
  
  Me.Caption = "Example MultiPlay on : " & DXMultiPlay.Get_ComputerName()
  
  'any Guid will do - use a Guid generator to create a new one when needed
  appGUID = "{66455400-78CC-11D7-867A-B2C49B9E315F}"
  
  'initialize the multiplayer interface
  DXMultiPlay.Init_DXMultiPlay
  
  lstConnections.Clear
  
  'enumerate and list the different connections (DirectX DirectPlay services) to chose from
  maxCount = DXMultiPlay.Refresh_AvailableConnections()
  
  For loop1 = 1 To maxCount
    lstConnections.AddItem DXMultiPlay.Get_ConnectionInfo(loop1)
  Next loop1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'cleanup the multiplayer interface
  DXMultiPlay.CleanUp_DXMultiPlay
End Sub

Private Sub lstConnections_Click()
  'select one of the available DirectX DirectPlay service providers
  DXMultiPlay.Set_ActiveConnection lstConnections.ListIndex + 1
  
  'display the descriptive name of the active connection
  lblConnection.Caption = DXMultiPlay.Get_ActiveConnectionInfo()
  
  If lblConnection.Caption = "" Then MsgBox "Failed to initialize the connection!", vbCritical, "Connection Failed"
  
  chkLocked.Value = 0
End Sub

Private Sub lstNames_Click()
  'this command will be re-enabled by the timer if another player's name has been selected
  cmdSend(0).Enabled = False
End Sub

Private Sub tmrSession_Timer()
  Dim sessionName As String, sessionPlayers As Long
  Dim sessionMaxPlayers As Long, sessionStillOpen As Boolean
  Dim xplayerID() As Long, xplayerName() As String, xplayerHandle() As String
  Dim loop1 As Long, mItem As ListItem, maxCount As Long
  Dim currentS As Long, currentP As Long
  Dim mMessage As DirectPlayMessage, mSysMessage As DXMP_SystemMessage
  
  'disable the timer until all sessions and messages updated
  tmrSession.Enabled = False
  
  selectedID = 0
  
  Set mItem = lsvSessions.SelectedItem
  
  If Not (mItem Is Nothing) Then currentS = mItem.Index
  
  currentP = lstNames.ListIndex
  
  lstNames.Clear
  lsvSession.ListItems.Clear
  lsvSessions.ListItems.Clear
  
  'refresh the list of available sessions - only those that were created with the matching Guid
  'this will return a maxcount of 0 if already in a session
  maxCount = DXMultiPlay.Refresh_AvailableSessions(appGUID)
  
  For loop1 = 1 To maxCount
    'loop through each enumerated session and get description
    DXMultiPlay.Get_SessionInfo loop1, sessionName, sessionPlayers, sessionMaxPlayers, sessionStillOpen
    
    'post each available session into the listview
    With lsvSessions
      Set mItem = .ListItems.Add(, , loop1)
      
      mItem.SubItems(1) = sessionName
      mItem.SubItems(2) = sessionPlayers
      mItem.SubItems(3) = sessionMaxPlayers
      
      If sessionStillOpen Then
        mItem.SubItems(4) = "YES"
      Else
        mItem.SubItems(4) = "NO"
      End If
      
      If currentS = loop1 Then mItem.Selected = True
    End With
  Next loop1
  
  'get description of the active session - if no active session then a value of 0 will be returned
  If DXMultiPlay.Get_ActiveSessionInfo(sessionName, sessionPlayers, sessionMaxPlayers, sessionStillOpen, xplayerID, xplayerName, xplayerHandle) <> 0 Then
    'post the active session description
    With lstNames
      For loop1 = 1 To sessionPlayers
        .AddItem loop1 & ") " & xplayerName(loop1) & " (" & xplayerHandle(loop1) & ")  [" & xplayerID(loop1) & "]"
        
        If loop1 = currentP + 1 Then selectedID = xplayerID(loop1)
      Next loop1
      
      If currentP >= 0 And currentP < .ListCount Then .ListIndex = currentP
    End With
    
    With lsvSession
      Set mItem = .ListItems.Add(, , sessionName)
      
      mItem.SubItems(1) = sessionPlayers
      mItem.SubItems(2) = sessionMaxPlayers
      
      If sessionStillOpen Then
        mItem.SubItems(3) = "YES"
      Else
        mItem.SubItems(3) = "NO"
      End If
    End With
    
    'enable/disable certain commands based on session status
    cmdLeave.Enabled = True
    
    If isHost Then
      chkLocked.Enabled = True
    Else
      chkLocked.Enabled = False
    End If
    
    If selectedID > 0 And selectedID <> playerID Then
      cmdSend(0).Enabled = True
    Else
      cmdSend(0).Enabled = False
    End If
    
    cmdSend(1).Enabled = True
    
    txtPlayerName.Enabled = False
    txtPlayerHandle.Enabled = False
    
    cmdJoin.Enabled = False
    cmdCreate.Enabled = False
  Else
    'enable/disable certain commands based on session status
    If DXMultiPlay.Get_ActiveConnectionInfo() <> "" Then
      cmdJoin.Enabled = True
      cmdCreate.Enabled = True
    Else
      cmdJoin.Enabled = False
      cmdCreate.Enabled = False
    End If
    
    txtPlayerName.Enabled = True
    txtPlayerHandle.Enabled = True
    
    cmdLeave.Enabled = False
    
    chkLocked.Enabled = False
    
    cmdSend(0).Enabled = False
    cmdSend(1).Enabled = False
  End If
  
  'process incoming messages - continue until no more messages
  Do While DXMultiPlay.Get_Message(mMessage, currentP)
    If currentP = 0 Then 'system message
      mSysMessage = DXMultiPlay.Read_SystemMessage(mMessage)
      
      With mSysMessage
        Select Case .messageType
          Case DPSYS_HOST
            'perform action(s) to prepare to host (if needed)
            
            ' ...
            ' ...
            ' ...
            
            txtMessages.Text = "System Message : Host moved here." & vbNewLine & vbNewLine & txtMessages.Text
            
            isHost = True
            
          Case DPSYS_SESSIONLOST
            'perform action(s) to recover and exit
            
            ' ...
            ' ...
            ' ...
            
            txtMessages.Text = "System Message : Session lost." & vbNewLine & vbNewLine & txtMessages.Text
            
          Case DPSYS_CREATEPLAYERORGROUP
            'perform action(s) to initialize data for the new player or group
            
            ' ...
            ' ...
            ' ...
            
            txtMessages.Text = "System Message : New player -> " & DXMultiPlay.Get_PlayerName(.playerID) & " - " & DXMultiPlay.Get_PlayerHandle(.playerID) & " [" & .playerID & "]" & _
                    vbNewLine & vbNewLine & txtMessages.Text
                    
          Case DPSYS_DESTROYPLAYERORGROUP
            'perform action(s) to release data that was used by the leaving player or group
            
            ' ...
            ' ...
            ' ...
            
            'you could provide the name and handle of the departing player if you had kept
            'a copy in a player data structure/type variable
            txtMessages.Text = "System Message : Lost player ->  [" & .playerID & "]" & _
                    vbNewLine & vbNewLine & txtMessages.Text
                    
        End Select
      End With
    Else
      'perform action(s) to handle the application defined message
      
      ' ...
      ' ...
      ' ...
      
      'we are just exchanging simple text messages that includes the sending player's name
      With mMessage
        txtMessages.Text = "Message From Player " & .ReadString() & vbNewLine & .ReadString() & vbNewLine & vbNewLine & txtMessages.Text
      End With
    End If
  Loop
  
  're-enable the timer
  tmrSession.Enabled = True
End Sub
