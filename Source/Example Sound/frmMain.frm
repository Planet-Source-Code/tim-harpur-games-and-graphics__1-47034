VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Games & Graphics - Sound Demo - Press Esc to Quit"
   ClientHeight    =   5790
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
   ScaleHeight     =   5790
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer soundTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   9060
      Top             =   210
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sound Channel 3 (3D)"
      Height          =   5385
      Left            =   5130
      TabIndex        =   16
      Top             =   300
      Width           =   4275
      Begin VB.CommandButton cmdPlay 
         Caption         =   "RepeatPlay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   810
         Width           =   1275
      End
      Begin VB.HScrollBar scrVelocity 
         Enabled         =   0   'False
         Height          =   225
         Index           =   2
         Left            =   1500
         Max             =   500
         Min             =   -500
         TabIndex        =   30
         Top             =   1950
         Width           =   2625
      End
      Begin VB.HScrollBar scrVelocity 
         Height          =   225
         Index           =   1
         Left            =   1500
         Max             =   500
         Min             =   -500
         TabIndex        =   29
         Top             =   1650
         Width           =   2625
      End
      Begin VB.HScrollBar scrVelocity 
         Height          =   225
         Index           =   0
         Left            =   1500
         Max             =   500
         Min             =   -500
         TabIndex        =   27
         Top             =   1350
         Width           =   2625
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   375
         Index           =   2
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   390
         Width           =   945
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   225
         Index           =   2
         Left            =   1500
         Max             =   0
         Min             =   -2000
         TabIndex        =   24
         Top             =   690
         Width           =   2625
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Index           =   2
         Left            =   90
         TabIndex        =   23
         Top             =   1800
         Width           =   945
      End
      Begin VB.CheckBox chkForce 
         Caption         =   "Force"
         Height          =   345
         Index           =   2
         Left            =   90
         TabIndex        =   22
         Top             =   1260
         Width           =   1155
      End
      Begin VB.PictureBox pic3DSound 
         BackColor       =   &H00000000&
         Height          =   2985
         Left            =   120
         ScaleHeight     =   195
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   265
         TabIndex        =   21
         Top             =   2280
         Width           =   4035
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            X1              =   138
            X2              =   148
            Y1              =   98
            Y2              =   82
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   102
            X2              =   114
            Y1              =   84
            Y2              =   98
         End
         Begin VB.Image imgSource 
            Height          =   480
            Left            =   1710
            Picture         =   "frmMain.frx":014A
            Top             =   210
            Width           =   480
         End
         Begin VB.Image imgListener 
            Height          =   480
            Left            =   1650
            Picture         =   "frmMain.frx":0594
            Top             =   1200
            Width           =   480
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Source Velocity x/y/z"
         Height          =   285
         Index           =   3
         Left            =   1470
         TabIndex        =   28
         Top             =   1020
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Volume"
         Height          =   255
         Index           =   2
         Left            =   1860
         TabIndex        =   26
         Top             =   390
         Width           =   1845
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sound Channel 2"
      Height          =   2355
      Index           =   1
      Left            =   90
      TabIndex        =   6
      Top             =   3330
      Width           =   4995
      Begin VB.CommandButton cmdPlay 
         Caption         =   "RepeatPlay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   780
         Width           =   1275
      End
      Begin VB.HScrollBar scrFreq 
         Height          =   225
         Index           =   1
         Left            =   1500
         Max             =   4410
         Min             =   500
         TabIndex        =   19
         Top             =   1950
         Value           =   2205
         Width           =   3285
      End
      Begin VB.CheckBox chkForce 
         Caption         =   "Force"
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1155
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1860
         Width           =   945
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   945
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   225
         Index           =   1
         Left            =   1530
         Max             =   0
         Min             =   -2000
         TabIndex        =   8
         Top             =   540
         Width           =   3285
      End
      Begin VB.HScrollBar scrPan 
         Height          =   225
         Index           =   1
         Left            =   1530
         Max             =   1000
         Min             =   -1000
         TabIndex        =   7
         Top             =   1200
         Width           =   3285
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Frequency (5000 - 44100)"
         Height          =   315
         Index           =   3
         Left            =   1620
         TabIndex        =   20
         Top             =   1590
         Width           =   3075
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Volume"
         Height          =   255
         Index           =   1
         Left            =   2340
         TabIndex        =   11
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Pan Left/Right"
         Height          =   315
         Index           =   1
         Left            =   2490
         TabIndex        =   10
         Top             =   840
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sound Channel 1"
      Height          =   2355
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   300
      Width           =   4965
      Begin VB.CommandButton cmdPlay 
         Caption         =   "RepeatPlay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   780
         Width           =   1275
      End
      Begin VB.HScrollBar scrFreq 
         Height          =   225
         Index           =   0
         Left            =   1530
         Max             =   4410
         Min             =   500
         TabIndex        =   17
         Top             =   1920
         Value           =   2205
         Width           =   3285
      End
      Begin VB.CheckBox chkForce 
         Caption         =   "Force"
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1230
         Width           =   1155
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1830
         Width           =   945
      End
      Begin VB.HScrollBar scrPan 
         Height          =   225
         Index           =   0
         Left            =   1530
         Max             =   1000
         Min             =   -1000
         TabIndex        =   3
         Top             =   1200
         Width           =   3285
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   225
         Index           =   0
         Left            =   1530
         Max             =   0
         Min             =   -2000
         TabIndex        =   2
         Top             =   540
         Width           =   3285
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Frequency (5000 - 44100)"
         Height          =   315
         Index           =   2
         Left            =   1680
         TabIndex        =   18
         Top             =   1560
         Width           =   3075
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Pan Left/Right"
         Height          =   315
         Index           =   0
         Left            =   2490
         TabIndex        =   5
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Volume"
         Height          =   255
         Index           =   0
         Left            =   2340
         TabIndex        =   4
         Top             =   240
         Width           =   1845
      End
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

Private Const FreqMultiplier As Long = 10

Private Sub cmdPlay_Click(Index As Integer)
  'start playing selected channel
  
  'is channel to repeat
  If Index >= 3 Then
    'check if channel is to be forced
    If chkForce(Index - 3).Value <> 0 Then
      DXSound.Play_Sound Index - 2, True
    Else
      DXSound.Play_Sound Index - 2, True, False
    End If
  Else
    'check if channel is to be forced
    If chkForce(Index).Value <> 0 Then
      DXSound.Play_Sound Index + 1
    Else
      DXSound.Play_Sound Index + 1, , False
    End If
  End If
End Sub

Private Sub cmdStop_Click(Index As Integer)
  'stop playing selected channel
 
  DXSound.Stop_Sound Index + 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Sub Form_Load()
  'don't activate the timer until the form has loaded
  soundTimer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'clean up the sound routines
  DXSound.CleanUp_DXSound
End Sub

Private Sub pic3DSound_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  imgSource.Left = x - imgSource.Width / 2
  imgSource.Top = y - imgSource.Height / 2
  
  'change 3D source position
  DXSound.Change_3DSoundPosition 3, x, y, 0, scrVelocity(0).Value / 10#, scrVelocity(1).Value / 10#, scrVelocity(2).Value / 10#
  
  'then commit all 3D changes
  DXSound.Commit_3DSoundChanges
End Sub

Private Sub scrFreq_Change(Index As Integer)
  'user has changed playback frequency for one of the channels
  
  DXSound.Change_SoundSettings Index + 1, FreqMultiplier * scrFreq(Index).Value
End Sub

Private Sub scrVelocity_Change(Index As Integer)
  'update the source's position and velocity
  DXSound.Change_3DSoundPosition 3, imgSource.Left + imgSource.Width / 2, imgSource.Top + imgSource.Height / 2, 0, scrVelocity(0).Value / 10#, scrVelocity(1).Value / 10#, scrVelocity(2).Value / 10#
  
  'then commit all 3D changes
  DXSound.Commit_3DSoundChanges
End Sub

Private Sub soundTimer_Timer()
  On Error Resume Next
  
  soundTimer.Enabled = False
  
  'Initialize DXSound for 3 channels with 3D enabled
  DXSound.Init_DXSound Me, 3, , True
  
  'load wav samples into the 3 channels - using channel #3 as a 3D channel
  DXSound.Load_SoundBuffer 1, "SOUND Sample1.wav"
  DXSound.Load_SoundBuffer 2, "SOUND Sample2.wav"
  DXSound.Load_3DSoundBuffer 3, "SOUND Sample3.wav"
  
  'set the freq sliders to match the frequency of the first 2 samples
  scrFreq(0).Value = DXSound.Get_SoundFrequency(1) / FreqMultiplier
  scrFreq(1).Value = DXSound.Get_SoundFrequency(2) / FreqMultiplier
  
  'set 3D listener in center of window with front facing being in the -ve Y (up)
  DXSound.Change_3DListenerPosition imgListener.Left + imgListener.Width / 2, imgListener.Top + imgListener.Height / 2, 0, , , , , -1
  
  'set 3D source (channel 3) to match source image in window
  DXSound.Change_3DSoundPosition 3, imgSource.Left + imgSource.Width / 2, imgSource.Top + imgSource.Height / 2, 0
  
  'set 3D source (channel 3) min distance to 20 and max distance to 120
  'the volume will be at max at distance of 20 or less, at half at 40, at quarter at 80, etc
  'the volume will also be at max from front (top) and drop off past the inner cone angle of 90
  'to least volume outside outer cone at 180 degrees (volume drop off will be -1000ths (-10) dbs)
  DXSound.Change_3DSoundSettings 3, 20, 120, 90, 180, -1000
  
  'then commit all 3D changes
  DXSound.Commit_3DSoundChanges
End Sub

Private Sub scrPan_Change(Index As Integer)
  'user has changed pan left/right for one of the channels
  
  DXSound.Change_SoundSettings Index + 1, , , scrPan(Index).Value
End Sub

Private Sub scrVolume_Change(Index As Integer)
  'user has changed volume for one of the channels
  
  DXSound.Change_SoundSettings Index + 1, , scrVolume(Index).Value
End Sub

