VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Games & Graphics - 2D Demo - Press Esc to Quit"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9330.001
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
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   622
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCBoxes 
      Caption         =   "Show Object Collision Boxes"
      Height          =   855
      Left            =   7350
      TabIndex        =   7
      Top             =   2910
      Width           =   1965
   End
   Begin VB.CommandButton cmdFire 
      BackColor       =   &H000000FF&
      Caption         =   "Fire"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   7410
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   1785
   End
   Begin VB.Frame Frame1 
      Caption         =   "Backdrop"
      Height          =   2385
      Left            =   7410
      TabIndex        =   1
      Top             =   90
      Width           =   1815
      Begin VB.CheckBox chkMapGrid 
         Caption         =   "Grid"
         Height          =   345
         Left            =   390
         TabIndex        =   5
         Top             =   1920
         Width           =   1155
      End
      Begin VB.OptionButton optBackdrop 
         Caption         =   "Imagemap"
         Height          =   465
         Index           =   2
         Left            =   60
         TabIndex        =   4
         Top             =   1500
         Width           =   1635
      End
      Begin VB.OptionButton optBackdrop 
         Caption         =   "Scrolling Bitmap"
         Height          =   795
         Index           =   1
         Left            =   60
         TabIndex        =   3
         Top             =   750
         Width           =   1665
      End
      Begin VB.OptionButton optBackdrop 
         Caption         =   "None"
         Height          =   465
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1305
      End
   End
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCBoxes_Click()
  Dim oList As cls2D_Object
  
  Set oList = objectListStart
  
  Do While Not (oList Is Nothing)
    With oList
      If chkCBoxes.value = 0 Then
        .CollisionBoxVisible = False
      ElseIf .MainAction <> 2 Then
        .CollisionBoxVisible = True
      End If
      
      Set oList = .chainNext
    End With
  Loop
End Sub

Private Sub cmdFire_Click()
  'user pressed the fire button
  
  Dim fObject As cls2D_Object
  
  Set fObject = objectListStart.Get_ChainObjectID(2)
  
  createMissiles fObject.PosX_1000ths, fObject.PosY_1000ths - 10000
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape And masterMode <> -2 Then masterMode = -1
End Sub

Private Sub Form_Load() 'wait till for has loaded before proceeding to initialization
  masterMode = 1
  
  picDisplay.Width = 480
  picDisplay.Height = 480
End Sub

Public Sub Initialize_Display()
  DXDraw.Init_DXDrawWindow Me, picDisplay, 4
  
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
