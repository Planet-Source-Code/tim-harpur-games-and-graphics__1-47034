VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Games & Graphics - Music Demo - Press Esc to Quit"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Music Channel 2"
      Height          =   2355
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   2790
      Width           =   9315
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   1860
         Width           =   945
      End
      Begin VB.CheckBox chkRepeat 
         Caption         =   "Repeat"
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   810
         Width           =   1185
      End
      Begin VB.ListBox lstMidiDev 
         Height          =   1950
         Index           =   1
         Left            =   4830
         TabIndex        =   12
         Top             =   270
         Width           =   4395
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   945
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   225
         Index           =   1
         Left            =   1500
         Max             =   0
         Min             =   -2000
         TabIndex        =   10
         Top             =   540
         Width           =   3285
      End
      Begin VB.HScrollBar scrTempo 
         Height          =   225
         Index           =   1
         Left            =   1500
         Max             =   200
         Min             =   25
         TabIndex        =   9
         Top             =   1200
         Value           =   100
         Width           =   3285
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Volume"
         Height          =   255
         Index           =   1
         Left            =   2340
         TabIndex        =   15
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Tempo"
         Height          =   315
         Index           =   1
         Left            =   2490
         TabIndex        =   14
         Top             =   840
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Music Channel 1"
      Height          =   2355
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   300
      Width           =   9315
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1830
         Width           =   945
      End
      Begin VB.HScrollBar scrTempo 
         Height          =   225
         Index           =   0
         Left            =   1500
         Max             =   200
         Min             =   25
         TabIndex        =   5
         Top             =   1200
         Value           =   100
         Width           =   3285
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   225
         Index           =   0
         Left            =   1500
         Max             =   0
         Min             =   -2000
         TabIndex        =   4
         Top             =   540
         Width           =   3285
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   945
      End
      Begin VB.ListBox lstMidiDev 
         Height          =   1950
         Index           =   0
         Left            =   4830
         TabIndex        =   2
         Top             =   270
         Width           =   4395
      End
      Begin VB.CheckBox chkRepeat 
         Caption         =   "Repeat"
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   810
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Tempo"
         Height          =   315
         Index           =   0
         Left            =   2490
         TabIndex        =   7
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Volume"
         Height          =   255
         Index           =   0
         Left            =   2340
         TabIndex        =   6
         Top             =   240
         Width           =   1845
      End
   End
   Begin VB.Timer musicTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   9060
      Top             =   2400
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************************
'
' Games & Graphics - Music Demo
'                                                     - written by Tim Harpur for Logicon Enterprises @logicon.biz
'
' ----------- User Licensing Notice -----------
'
' This file and all source code herein is property of Logicon Enterprises.
' Whether in its original or modified form, Logicon Enterprises retains ownership of this file.
'
'***************************************************************************************************************

Option Explicit
Option Base 0

'we need 2 flags, 1 flag to keep track of when a channel has been requested to play (IsPlaying),
'and 1 flag to keep track of when a channel has actually started playing (HasStarted)
'this is nescessary as there may be a noticable lag time between the request to start a channel
'playing and when it actually starts playing
Private IsPlaying(1 To 2) As Boolean, HasStarted(1 To 2) As Boolean

Private Sub cmdPlay_Click(Index As Integer)
  'start playing selected channel - only if it is not already playing
  If IsPlaying(Index + 1) = True Then Exit Sub
  
  IsPlaying(Index + 1) = True
  HasStarted(Index + 1) = False
  
  DXSound.Play_Music Index + 1
End Sub

Private Sub cmdStop_Click(Index As Integer)
  'stop playing selected channel - only if it is playing
  If IsPlaying(Index + 1) = False Then Exit Sub
  
  IsPlaying(Index + 1) = False
  HasStarted(Index + 1) = False
  
  DXSound.Stop_Music Index + 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Sub Form_Load()
  'don't activate the timer until the form has loaded
  musicTimer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'shut down the timer and clean up the input input routines
  musicTimer.Enabled = False
  
  DXSound.CleanUp_DXSound
End Sub

Private Sub lstMidiDev_Click(Index As Integer)
  'user has selected a new midi port
  
  'turn of timer while changing midi port and reloading midi sample
  musicTimer.Enabled = False
  
  'change midi ports - this also stops the channel and dumps any loaded midi file
  DXSound.Set_MidiPort Index + 1, lstMidiDev(Index).ListIndex + 1
  
  'since midi port was changed - channel was dumped and needs to be reloaded
  If Index = 0 Then
    DXSound.Load_MusicFromMidi 1, "MUSIC Sample1.mid"
  Else
    DXSound.Load_MusicFromMidi 2, "MUSIC Sample2.mid"
  End If
  
  'channel is not playing (if it was, then it has been stopped by being dumped)
  HasStarted(Index + 1) = False
  IsPlaying(Index + 1) = False
  
  'all done - re-enable the timer
  musicTimer.Enabled = True
End Sub

Private Sub musicTimer_Timer()
  Dim loop1 As Long
  
  On Error Resume Next
  
  'check if DXSound has already been initialized - if not then initialize
  If DXSound.GetDirectSound() Is Nothing Then
    DXSound.Init_DXSound Me, 0, 2
    
    lstMidiDev(0).Clear
    lstMidiDev(1).Clear
    
    'list all the available midi devices
    'note that some of these devices may not be designed for playback and will not produce sound
    'the default midi device is usually the Microsoft Synthesizer
    For loop1 = 1 To DXSound.Get_TotalMidiPorts()
      lstMidiDev(0).AddItem DXSound.Get_MidiPortDescription(loop1)
      lstMidiDev(1).AddItem DXSound.Get_MidiPortDescription(loop1)
    Next loop1
    
    DXSound.Load_MusicFromMidi 1, "Sample1.mid"
    DXSound.Load_MusicFromMidi 2, "Sample2.mid"
  End If
  
  'check to see if music channel 1 has is stopped
  If DXSound.IsMusicPlaying(1) = False Then
    If IsPlaying(1) And HasStarted(1) Then 'if channel 1 is not playing now but was started
      'the channel is no longer player - so it is no longer started - turn off the HasStarted flag
      HasStarted(1) = False
      
      If chkRepeat(0).Value = 0 Then 'check for repeat mode
        IsPlaying(1) = False 'if no repeat then turn off the IsPlaying flag
      Else
        DXSound.Play_Music 1 'if repeat then start playing channel 1 again
      End If
    End If
    
    'channel 1 is currently not playing (either stopped or preparing to start)
    'so make sure back colour of play button is reset
    If cmdPlay(0).BackColor <> &H8000000F Then cmdPlay(0).BackColor = &H8000000F
  ElseIf HasStarted(1) = False Then
    'channel 1 has started playing
    HasStarted(1) = True
    
    'so change back colour of play button
    cmdPlay(0).BackColor = &H7F000
  End If
  
  'check to see if music channel 2 has is stopped
  If DXSound.IsMusicPlaying(2) = False Then
    If IsPlaying(2) And HasStarted(2) Then 'if channel 2 is not playing now but was started
      'the channel is no longer player - so it is no longer started - turn off the HasStarted flag
      HasStarted(2) = False
      
      If chkRepeat(1).Value = 0 Then 'check for repeat mode
        IsPlaying(2) = False 'if no repeat then turn off the IsPlaying flag
      Else
        DXSound.Play_Music 2 'if repeat then start playing channel 2 again
      End If
    End If
    
    'channel 2 is currently not playing (either stopped or preparing to start)
    'so make sure back colour of play button is reset
    If cmdPlay(1).BackColor <> &H8000000F Then cmdPlay(1).BackColor = &H8000000F
  ElseIf HasStarted(2) = False Then
    'channel 2 has started playing
    HasStarted(2) = True
    
    'so change back colour of play button
    cmdPlay(1).BackColor = &H7F000
  End If
End Sub

Private Sub scrTempo_Change(Index As Integer)
  'user has changed tempo for one of the channels
  
  DXSound.Change_MusicSettings Index + 1, , scrTempo(Index).Value / 100#
End Sub

Private Sub scrVolume_Change(Index As Integer)
  'user has changed volume for one of the channels
  
  DXSound.Change_MusicSettings Index + 1, scrVolume(Index).Value
End Sub

