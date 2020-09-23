VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Games & Graphics - Input Demo - Press Esc to Quit"
   ClientHeight    =   7185
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Controller #1 Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   9225
      Begin VB.CheckBox chkButton 
         Caption         =   "B32"
         Height          =   285
         Index           =   31
         Left            =   8310
         TabIndex        =   84
         Top             =   3030
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B31"
         Height          =   285
         Index           =   30
         Left            =   8310
         TabIndex        =   83
         Top             =   2760
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B30"
         Height          =   285
         Index           =   29
         Left            =   7470
         TabIndex        =   82
         Top             =   3300
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B29"
         Height          =   285
         Index           =   28
         Left            =   7470
         TabIndex        =   81
         Top             =   3030
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B28"
         Height          =   285
         Index           =   27
         Left            =   7470
         TabIndex        =   80
         Top             =   2760
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B27"
         Height          =   285
         Index           =   26
         Left            =   6630
         TabIndex        =   79
         Top             =   3300
         Width           =   825
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B26"
         Height          =   285
         Index           =   25
         Left            =   6630
         TabIndex        =   78
         Top             =   3030
         Width           =   885
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B25"
         Height          =   285
         Index           =   24
         Left            =   6630
         TabIndex        =   77
         Top             =   2760
         Width           =   915
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B24"
         Height          =   285
         Index           =   23
         Left            =   5850
         TabIndex        =   76
         Top             =   3300
         Width           =   1000
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B23"
         Height          =   285
         Index           =   22
         Left            =   5850
         TabIndex        =   75
         Top             =   3030
         Width           =   1000
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B22"
         Height          =   285
         Index           =   21
         Left            =   5850
         TabIndex        =   74
         Top             =   2760
         Width           =   1000
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B21"
         Height          =   285
         Index           =   20
         Left            =   5070
         TabIndex        =   73
         Top             =   3300
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B20"
         Height          =   285
         Index           =   19
         Left            =   5070
         TabIndex        =   72
         Top             =   3030
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B19"
         Height          =   285
         Index           =   18
         Left            =   5070
         TabIndex        =   71
         Top             =   2760
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B18"
         Height          =   285
         Index           =   17
         Left            =   4230
         TabIndex        =   70
         Top             =   3300
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B17"
         Height          =   285
         Index           =   16
         Left            =   4230
         TabIndex        =   69
         Top             =   3030
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B16"
         Height          =   285
         Index           =   15
         Left            =   4230
         TabIndex        =   68
         Top             =   2760
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B15"
         Height          =   285
         Index           =   14
         Left            =   3390
         TabIndex        =   67
         Top             =   3300
         Width           =   825
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B14"
         Height          =   285
         Index           =   13
         Left            =   3390
         TabIndex        =   66
         Top             =   3030
         Width           =   885
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B13"
         Height          =   285
         Index           =   12
         Left            =   3390
         TabIndex        =   65
         Top             =   2760
         Width           =   915
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B12"
         Height          =   285
         Index           =   11
         Left            =   2550
         TabIndex        =   64
         Top             =   3300
         Width           =   1000
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B11"
         Height          =   285
         Index           =   10
         Left            =   2550
         TabIndex        =   63
         Top             =   3030
         Width           =   1000
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B10"
         Height          =   285
         Index           =   9
         Left            =   2550
         TabIndex        =   62
         Top             =   2760
         Width           =   1000
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B9"
         Height          =   285
         Index           =   8
         Left            =   1770
         TabIndex        =   61
         Top             =   3300
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B8"
         Height          =   285
         Index           =   7
         Left            =   1770
         TabIndex        =   60
         Top             =   3030
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B7"
         Height          =   285
         Index           =   6
         Left            =   1770
         TabIndex        =   59
         Top             =   2760
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B6"
         Height          =   285
         Index           =   5
         Left            =   990
         TabIndex        =   58
         Top             =   3300
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B5"
         Height          =   285
         Index           =   4
         Left            =   990
         TabIndex        =   57
         Top             =   3030
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B4"
         Height          =   285
         Index           =   3
         Left            =   990
         TabIndex        =   56
         Top             =   2760
         Width           =   800
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B3"
         Height          =   285
         Index           =   2
         Left            =   210
         TabIndex        =   55
         Top             =   3300
         Width           =   825
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B2"
         Height          =   285
         Index           =   1
         Left            =   210
         TabIndex        =   54
         Top             =   3030
         Width           =   885
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "B1"
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   53
         Top             =   2760
         Width           =   915
      End
      Begin VB.OptionButton optPOV 
         Height          =   315
         Index           =   8
         Left            =   7770
         TabIndex        =   52
         Top             =   960
         Width           =   285
      End
      Begin VB.OptionButton optPOV 
         Height          =   315
         Index           =   7
         Left            =   7680
         TabIndex        =   51
         Top             =   1350
         Width           =   285
      End
      Begin VB.OptionButton optPOV 
         Height          =   315
         Index           =   6
         Left            =   7800
         TabIndex        =   50
         Top             =   1770
         Width           =   285
      End
      Begin VB.OptionButton optPOV 
         Height          =   315
         Index           =   5
         Left            =   8160
         TabIndex        =   49
         Top             =   1920
         Width           =   285
      End
      Begin VB.OptionButton optPOV 
         Height          =   315
         Index           =   4
         Left            =   8550
         TabIndex        =   48
         Top             =   1800
         Width           =   285
      End
      Begin VB.OptionButton optPOV 
         Height          =   315
         Index           =   3
         Left            =   8640
         TabIndex        =   47
         Top             =   1350
         Width           =   285
      End
      Begin VB.OptionButton optPOV 
         Height          =   315
         Index           =   2
         Left            =   8520
         TabIndex        =   46
         Top             =   930
         Width           =   285
      End
      Begin VB.OptionButton optPOV 
         Height          =   315
         Index           =   1
         Left            =   8160
         TabIndex        =   45
         Top             =   780
         Width           =   315
      End
      Begin VB.OptionButton optPOV 
         Height          =   315
         Index           =   0
         Left            =   8160
         TabIndex        =   44
         Top             =   1350
         Width           =   285
      End
      Begin VB.ListBox lstControllers 
         Height          =   2220
         Left            =   180
         TabIndex        =   36
         Top             =   390
         Width           =   4965
      End
      Begin VB.Line Line7 
         X1              =   8340
         X2              =   8490
         Y1              =   1620
         Y2              =   1770
      End
      Begin VB.Line Line6 
         X1              =   8160
         X2              =   8010
         Y1              =   1620
         Y2              =   1800
      End
      Begin VB.Line Line5 
         X1              =   8160
         X2              =   7950
         Y1              =   1410
         Y2              =   1200
      End
      Begin VB.Line Line4 
         X1              =   7920
         X2              =   8130
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line3 
         X1              =   8370
         X2              =   8610
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line2 
         X1              =   8340
         X2              =   8490
         Y1              =   1410
         Y2              =   1200
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   8250
         X2              =   8250
         Y1              =   1620
         Y2              =   1890
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   8250
         X2              =   8250
         Y1              =   1080
         Y2              =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "POV 1"
         Height          =   345
         Index           =   15
         Left            =   7710
         TabIndex        =   43
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblAxis 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   5
         Left            =   6120
         TabIndex        =   42
         Top             =   2310
         Width           =   1155
      End
      Begin VB.Label lblAxis 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   4
         Left            =   6120
         TabIndex        =   41
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label lblAxis 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   3
         Left            =   6120
         TabIndex        =   40
         Top             =   1530
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "rZ Axis"
         Height          =   345
         Index           =   14
         Left            =   5310
         TabIndex        =   39
         Top             =   2310
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "rY Axis"
         Height          =   345
         Index           =   13
         Left            =   5310
         TabIndex        =   38
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "rX Axis"
         Height          =   345
         Index           =   12
         Left            =   5310
         TabIndex        =   37
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblAxis 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   2
         Left            =   6120
         TabIndex        =   35
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblAxis 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   6120
         TabIndex        =   34
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label lblAxis 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   0
         Left            =   6120
         TabIndex        =   33
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Z Axis"
         Height          =   345
         Index           =   11
         Left            =   5310
         TabIndex        =   32
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Y Axis"
         Height          =   345
         Index           =   10
         Left            =   5310
         TabIndex        =   31
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "X Axis"
         Height          =   345
         Index           =   9
         Left            =   5310
         TabIndex        =   30
         Top             =   390
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mouse Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   9225
      Begin VB.Label lblMButton 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   2
         Left            =   7770
         TabIndex        =   29
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label lblMButton 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   7770
         TabIndex        =   28
         Top             =   930
         Width           =   1155
      End
      Begin VB.Label lblMButton 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   0
         Left            =   7770
         TabIndex        =   27
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label lblRelative 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   2
         Left            =   4140
         TabIndex        =   26
         Top             =   1350
         Width           =   1155
      End
      Begin VB.Label lblRelative 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   4140
         TabIndex        =   25
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label lblRelative 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   0
         Left            =   4140
         TabIndex        =   24
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label lblAbsolute 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   2
         Left            =   1320
         TabIndex        =   23
         Top             =   1350
         Width           =   1155
      End
      Begin VB.Label lblAbsolute 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   1320
         TabIndex        =   22
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label lblAbsolute 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   0
         Left            =   1320
         TabIndex        =   21
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Button 3 (Center)"
         Height          =   345
         Index           =   8
         Left            =   5820
         TabIndex        =   20
         Top             =   1350
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Button 2 (Right)"
         Height          =   345
         Index           =   7
         Left            =   5820
         TabIndex        =   19
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Z Relative"
         Height          =   345
         Index           =   6
         Left            =   3030
         TabIndex        =   18
         Top             =   1350
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Z Absolute"
         Height          =   345
         Index           =   5
         Left            =   150
         TabIndex        =   17
         Top             =   1350
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Button 1 (Left)"
         Height          =   345
         Index           =   4
         Left            =   5820
         TabIndex        =   16
         Top             =   570
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Y Relative"
         Height          =   345
         Index           =   3
         Left            =   3030
         TabIndex        =   15
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "X Relative"
         Height          =   345
         Index           =   2
         Left            =   3030
         TabIndex        =   14
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Y Absolute"
         Height          =   345
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "X Absolute"
         Height          =   345
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   600
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Keyboard Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   9225
      Begin VB.Label keyDown 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   8
         Left            =   8220
         TabIndex        =   11
         Top             =   510
         Width           =   615
      End
      Begin VB.Label keyDown 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   7
         Left            =   7260
         TabIndex        =   10
         Top             =   510
         Width           =   615
      End
      Begin VB.Label keyDown 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   6
         Left            =   6300
         TabIndex        =   9
         Top             =   510
         Width           =   615
      End
      Begin VB.Label keyDown 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   5
         Left            =   5310
         TabIndex        =   8
         Top             =   510
         Width           =   615
      End
      Begin VB.Label keyDown 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   4
         Left            =   4320
         TabIndex        =   7
         Top             =   510
         Width           =   615
      End
      Begin VB.Label keyDown 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   3
         Left            =   3360
         TabIndex        =   6
         Top             =   510
         Width           =   615
      End
      Begin VB.Label keyDown 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   2
         Left            =   2370
         TabIndex        =   5
         Top             =   510
         Width           =   615
      End
      Begin VB.Label keyDown 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   1
         Left            =   1380
         TabIndex        =   4
         Top             =   510
         Width           =   615
      End
      Begin VB.Label keyDown 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   0
         Left            =   420
         TabIndex        =   3
         Top             =   510
         Width           =   615
      End
   End
   Begin VB.Timer inputTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   9180
      Top             =   4020
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************************
'
' Games & Graphics - Input Demo
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

Private mouseXAbs As Long, mouseYAbs As Long, mouseZAbs As Long

Sub Form_Load()
  'don't activate the timer until the form has loaded
  inputTimer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'shut down the timer and clean up the input input routines
  inputTimer.Enabled = False
  
  DXInput.CleanUp_DXInput
End Sub

Private Sub inputTimer_Timer()
  Dim loop1 As Long, tempCount As Long
  
  On Error Resume Next
  
  'check if DXInput has already been initialized - if not then initialize
  If DXInput.GetDirectInput() Is Nothing Then
    DXInput.Init_DXInput Me
    
    lstControllers.Clear
    
    'list all available controllers
    For loop1 = 1 To DXInput.Get_ControllerCount()
      lstControllers.AddItem DXInput.Get_ControllerDescription(loop1)
    Next loop1
    
    DXInput.Acquire_Keyboard IM_Background
    DXInput.Acquire_Mouse IM_Background
    
    If DXInput.Get_ControllerCount() > 0 Then
      DXInput.Acquire_Controller
      
      'try and set range for all axis from -10,000 to 10,000
      DXInput.Set_ControllerRange , , , True, True, True, True, True, True
      
      'try and set deadzone to 200 and saturation to 8000 for all axis
      DXInput.Set_ControllerDeadZoneSat , , , True, True, True, True, True, True
    End If
  End If
  
  'update the contents of DXInput.dx_KeyboardState
  DXInput.Poll_Keyboard
  
  With DXInput.dx_KeyboardState
    If .Key(DIK_ESCAPE) <> 0 Then 'if the escape key is pressed then unload the program
      Unload Me
      
      Exit Sub
    End If
    
    tempCount = 0
    
    'check all possible keys
    For loop1 = 0 To 255
      If .Key(loop1) <> 0 Then
        keyDown(tempCount) = loop1
        
        tempCount = tempCount + 1
        
        If tempCount >= 9 Then Exit For
      End If
    Next loop1
  End With
  
  For loop1 = tempCount To 8
    keyDown(loop1) = ""
  Next loop1
  
  'update the contents of DXInput.dx_MouseState
  DXInput.Poll_Mouse
  
  With DXInput.dx_MouseState
    mouseXAbs = mouseXAbs + .X 'left/right
    mouseYAbs = mouseYAbs + .Y 'up/down
    mouseZAbs = mouseZAbs + .z 'scroll wheel
    
    lblAbsolute(0).Caption = mouseXAbs
    lblAbsolute(1).Caption = mouseYAbs
    lblAbsolute(2).Caption = mouseZAbs
    
    lblRelative(0).Caption = .X
    lblRelative(1).Caption = .Y
    lblRelative(2).Caption = .z
    
    'check the left, right and middle mouse button
    For loop1 = 0 To 2
      If .buttons(loop1) <> 0 Then
        lblMButton(loop1).Caption = "DOWN"
      Else
        lblMButton(loop1).Caption = "UP"
      End If
    Next loop1
  End With
  
  'update the contents of DXInput.dx_ControllerState
  DXInput.Poll_Controller
  
  With DXInput.dx_ControllerState(1)
    'check the axis
    lblAxis(0).Caption = .X
    lblAxis(1).Caption = .Y
    lblAxis(2).Caption = .z
    lblAxis(3).Caption = .rx 'the next 3 are rotational axis
    lblAxis(4).Caption = .ry
    lblAxis(5).Caption = .rz
    
    'instead of checking for exact values, I am checking for ranges of values as some controllers
    'may return more precise measurements than 45 degrees for their POVs
    If .POV(0) = -1 Or .POV(0) = 65535 Then 'centered
      optPOV(0).Value = True
    ElseIf .POV(0) > 33750 Or .POV(0) < 2250 Then 'up
      optPOV(1).Value = True
    ElseIf .POV(0) < 6750 Then '45
      optPOV(2).Value = True
    ElseIf .POV(0) < 11250 Then '90
      optPOV(3).Value = True
    ElseIf .POV(0) < 15750 Then '135
      optPOV(4).Value = True
    ElseIf .POV(0) < 20250 Then '180
      optPOV(5).Value = True
    ElseIf .POV(0) < 24750 Then '225
      optPOV(6).Value = True
    ElseIf .POV(0) < 29250 Then '270
      optPOV(7).Value = True
    Else '315
      optPOV(8).Value = True
    End If
    
    'check all possible joystick buttons
    For loop1 = 0 To DXInput.dx_ControllerDesc(1).buttons - 1
      If .buttons(loop1) <> 0 Then
        chkButton(loop1).Value = 1
      Else
        chkButton(loop1).Value = 0
      End If
    Next loop1
    
    'set remaining buttons off as they don't exist on this controller
    For loop1 = DXInput.dx_ControllerDesc(1).buttons To 31
      chkButton(loop1).Value = 0
    Next loop1
  End With
End Sub

Private Sub lstControllers_Click()
  'try and Acquire selected controller
  DXInput.Acquire_Controller , lstControllers.ListIndex + 1
  
  'try and set range for all axis from -10,000 to 10,000
  DXInput.Set_ControllerRange , , , True, True, True, True, True, True
    
  'try and set deadzone to 200 and saturation to 8000 for all axis
  DXInput.Set_ControllerDeadZoneSat , , , True, True, True, True, True, True
End Sub
