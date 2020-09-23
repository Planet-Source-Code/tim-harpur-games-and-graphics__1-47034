VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game & Graphics - Game Flow Demo - Press Esc to Quit"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8310
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   8310
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":014A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1035
      Index           =   7
      Left            =   2760
      TabIndex        =   15
      Top             =   5700
      Visible         =   0   'False
      Width           =   4845
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":0214
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2025
      Index           =   6
      Left            =   2880
      TabIndex        =   14
      Top             =   2250
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":038E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1305
      Index           =   5
      Left            =   540
      TabIndex        =   13
      Top             =   1590
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":04EA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1275
      Index           =   4
      Left            =   480
      TabIndex        =   12
      Top             =   4350
      Visible         =   0   'False
      Width           =   5265
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":060D
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1755
      Index           =   3
      Left            =   2490
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":07AB
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1545
      Index           =   2
      Left            =   1110
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":08BF
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1065
      Index           =   1
      Left            =   2430
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":09B4
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1305
      Index           =   0
      Left            =   1470
      TabIndex        =   8
      Top             =   1110
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000000&
      X1              =   4380
      X2              =   4380
      Y1              =   4710
      Y2              =   5310
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00000000&
      BorderStyle     =   3  'Dot
      X1              =   1920
      X2              =   2460
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000000&
      X1              =   5190
      X2              =   6180
      Y1              =   3450
      Y2              =   3450
   End
   Begin VB.Line Line11 
      BorderStyle     =   3  'Dot
      X1              =   2220
      X2              =   5910
      Y1              =   5790
      Y2              =   5790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Splash Screen and Program Cleanup Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Index           =   7
      Left            =   2100
      TabIndex        =   7
      Top             =   6900
      Width           =   4245
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Central Timing, Input Monitor, and Game State Controller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1515
      Index           =   6
      Left            =   180
      TabIndex        =   6
      Top             =   3210
      Width           =   2385
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level Ended (Player Died or Level Completed)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1395
      Index           =   5
      Left            =   3420
      TabIndex        =   5
      Top             =   3270
      Width           =   1755
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Main In-Game Loop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   765
      Index           =   4
      Left            =   6090
      TabIndex        =   4
      Top             =   5070
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game and Level Initialization"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1155
      Index           =   3
      Left            =   6270
      TabIndex        =   3
      Top             =   2850
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Settings (Video/Audio)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   795
      Index           =   2
      Left            =   2460
      TabIndex        =   2
      Top             =   1080
      Width           =   2085
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Main Game Splash Screen and Game Options (New/Load/Difficulty)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1395
      Index           =   1
      Left            =   4770
      TabIndex        =   1
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Introductory Splash Screen and Program Initialization Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Index           =   0
      Left            =   2190
      TabIndex        =   0
      Top             =   150
      Width           =   4245
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000000&
      X1              =   4380
      X2              =   5850
      Y1              =   5310
      Y2              =   5310
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      X1              =   7050
      X2              =   7050
      Y1              =   4260
      Y2              =   3990
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      X1              =   6990
      X2              =   6990
      Y1              =   2550
      Y2              =   2790
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      X1              =   4530
      X2              =   4710
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      X1              =   1350
      X2              =   1350
      Y1              =   6570
      Y2              =   6840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderStyle     =   3  'Dot
      X1              =   2550
      X2              =   3390
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderStyle     =   3  'Dot
      X1              =   2280
      X2              =   4710
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   4860
      X2              =   4860
      Y1              =   2550
      Y2              =   3210
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   6270
      X2              =   6270
      Y1              =   930
      Y2              =   1170
   End
   Begin VB.Shape shpFlow 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   7
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6810
      Width           =   8085
   End
   Begin VB.Shape shpFlow 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1515
      Index           =   5
      Left            =   3390
      Shape           =   4  'Rounded Rectangle
      Top             =   3210
      Width           =   1815
   End
   Begin VB.Shape shpFlow 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   5415
      Index           =   6
      Left            =   150
      Shape           =   2  'Oval
      Top             =   1170
      Width           =   2415
   End
   Begin VB.Shape shpFlow 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   2355
      Index           =   4
      Left            =   5310
      Shape           =   3  'Circle
      Top             =   4260
      Width           =   3405
   End
   Begin VB.Shape shpFlow 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   1245
      Index           =   3
      Left            =   6180
      Top             =   2760
      Width           =   1725
   End
   Begin VB.Shape shpFlow 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   2
      Left            =   2460
      Top             =   1020
      Width           =   2085
   End
   Begin VB.Shape shpFlow 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   1395
      Index           =   1
      Left            =   4710
      Top             =   1170
      Width           =   3195
   End
   Begin VB.Shape shpFlow 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   8085
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'declare any private member variables here
'                         :
'                         :
'                         :



Private Sub Form_Load()
  'once the form has loaded signal that the initialization can continue
  gameState = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'prevent any attempt to unload this form unless cleanup has occurred
  If gameState <> -2 Then
    Cancel = 1
    
    MsgBox "Press Esc to quit", vbOKOnly, "This button has been blocked"
  End If
End Sub

Public Sub gameInitialize()
  'initialize the display
  '                         :
  '                         :
  '                         :
  
  'initialize the input
  DXInput.Init_DXInput Me
  '                         :
  '                         :
  '                         :
  
  'intialize the sound and music
  '                         :
  '                         :
  '                         :
  
  'load any global use sound, music, and graphics (level specific loading should be done just
  'before each level as part of level initialization and not here to save on available memory and
  'reduce load times)
  '                         :
  '                         :
  '                         :
    
  'signal that the intialization is complete and prepare to enter gameState 2
  gameState = 2
  gameSubState = -1
End Sub

