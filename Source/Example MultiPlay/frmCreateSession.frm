VERSION 5.00
Begin VB.Form frmCreateSession 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create a New Session"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCreateSession.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMaxPlayers 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   4800
      TabIndex        =   1
      Top             =   630
      Width           =   1365
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   150
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   585
      Left            =   180
      TabIndex        =   2
      Top             =   1290
      Width           =   1815
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Default         =   -1  'True
      Height          =   585
      Left            =   4350
      TabIndex        =   3
      Top             =   1290
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Max Players"
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
      Left            =   4800
      TabIndex        =   5
      Top             =   300
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Session Name"
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
      Left            =   1500
      TabIndex        =   4
      Top             =   270
      Width           =   1650
   End
End
Attribute VB_Name = "frmCreateSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdCreate_Click()
  If txtName.Text <> "" Then
    If Val(txtMaxPlayers.Text) > 1 Then
      'create a new session
      frmMultiPlay.playerID = DXMultiPlay.Create_ActiveSession(frmMultiPlay.txtPlayerName.Text, frmMultiPlay.txtPlayerHandle.Text, txtName.Text, Val(txtMaxPlayers.Text), frmMultiPlay.appGUID)
      
      'if successful, the playerID is returned - otherwise 0 is returned
      If frmMultiPlay.playerID <> 0 Then
        frmMultiPlay.isHost = True
        
        Unload Me
      Else
        MsgBox "There was an error creating the session", vbCritical, "Failed To Create Session"
      End If
    Else
      MsgBox "You must select at least 2 players for a multiplayer session!", vbCritical, "Invalid Value"
    End If
  Else
    MsgBox "You must give a name for your multiplayer session!", vbCritical, "Invalid Value"
  End If
End Sub
