VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Games & Graphics - 2DX Demo - Press Esc to Quit"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7290
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   524
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   486
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDisplay 
      Height          =   7200
      Left            =   60
      ScaleHeight     =   476
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   476
      TabIndex        =   0
      Top             =   120
      Width           =   7200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Use cursor keys to adjust the speed of the rotors. Space bar launches missiles."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   1
      Top             =   7440
      Width           =   7005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape And masterMode <> -2 Then
    masterMode = -1
    
    KeyCode = 0
  ElseIf KeyCode = vbKeyLeft Then
    bodySpeed = bodySpeed + 0.005
    If bodySpeed > 0.05 Then bodySpeed = 0.05
    
    KeyCode = 0
  ElseIf KeyCode = vbKeyRight Then
    bodySpeed = bodySpeed - 0.005
    If bodySpeed < -0.05 Then bodySpeed = -0.05
    
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    rotorSpeed = rotorSpeed + 0.01
    If rotorSpeed > 0.23 Then rotorSpeed = 0.23
    
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    rotorSpeed = rotorSpeed - 0.01
    If rotorSpeed < 0 Then rotorSpeed = 0
    
    KeyCode = 0
  ElseIf KeyCode = vbKeySpace Then
    createMissile
    
    KeyCode = 0
  End If
End Sub

Private Sub Form_Load() 'wait till for has loaded before proceeding to initialization
  masterMode = 1
  
  picDisplay.Width = 480
  picDisplay.Height = 480
End Sub

Public Sub Initialize_Display()
  'DXDraw.Init_DXDrawWindow Me, picDisplay, 4
  DXDraw.Init_DXDrawScreen Me, , , , 4
  DXDraw.Set_D3DFilterMode True
  
  masterMode = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'prevent unloading this form until cleanup has occurred
  If masterMode <> -2 Then
    'signal that the program should clean up and shut down
    masterMode = -1
    Cancel = 1
  End If
End Sub

