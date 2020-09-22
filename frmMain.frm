VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "M-209 Sim"
   ClientHeight    =   8055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":030A
   ScaleHeight     =   8055
   ScaleMode       =   0  'User
   ScaleWidth      =   8050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   120
      Top             =   2760
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   120
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgLug 
      Height          =   495
      Index           =   2
      Left            =   3550
      MouseIcon       =   "frmMain.frx":0614
      MousePointer    =   99  'Custom
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgLug 
      Height          =   495
      Index           =   8
      Left            =   6129
      MouseIcon       =   "frmMain.frx":091E
      MousePointer    =   99  'Custom
      Top             =   3255
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgLug 
      Height          =   495
      Index           =   7
      Left            =   5854
      MouseIcon       =   "frmMain.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgLug 
      Height          =   500
      Index           =   6
      Left            =   5568
      MouseIcon       =   "frmMain.frx":0F32
      MousePointer    =   99  'Custom
      Top             =   3250
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Image imgLug 
      Height          =   500
      Index           =   5
      Left            =   5028
      MouseIcon       =   "frmMain.frx":123C
      MousePointer    =   99  'Custom
      Top             =   3250
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Image imgLug 
      Height          =   500
      Index           =   4
      Left            =   4433
      MouseIcon       =   "frmMain.frx":1546
      MousePointer    =   99  'Custom
      Top             =   3250
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Image imgLug 
      Height          =   500
      Index           =   3
      Left            =   3902
      MouseIcon       =   "frmMain.frx":1850
      MousePointer    =   99  'Custom
      Top             =   3250
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Image imgLug 
      Height          =   495
      Index           =   1
      Left            =   3270
      MouseIcon       =   "frmMain.frx":1B5A
      MousePointer    =   99  'Custom
      Top             =   3255
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgReset 
      Height          =   495
      Left            =   6360
      MouseIcon       =   "frmMain.frx":1E64
      MousePointer    =   99  'Custom
      ToolTipText     =   " Reset Wheels and Counter "
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Press F1 for help and more information on the M-209"
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   2880
      TabIndex        =   13
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Image imgSoundOff 
      Height          =   240
      Left            =   1680
      Picture         =   "frmMain.frx":216E
      Top             =   7560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSoundOn 
      Height          =   240
      Left            =   1200
      Picture         =   "frmMain.frx":24B0
      Top             =   7560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSound 
      Height          =   240
      Left            =   6120
      MouseIcon       =   "frmMain.frx":27F2
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":2AFC
      ToolTipText     =   " Sound Off "
      Top             =   45
      Width           =   240
   End
   Begin VB.Image imgGallerie 
      Height          =   255
      Left            =   6440
      MouseIcon       =   "frmMain.frx":2E3E
      MousePointer    =   99  'Custom
      ToolTipText     =   " Picture Gallery "
      Top             =   50
      Width           =   255
   End
   Begin VB.Image imgLug2 
      Height          =   120
      Left            =   6120
      Picture         =   "frmMain.frx":3148
      Stretch         =   -1  'True
      Top             =   3440
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgLug1 
      Height          =   120
      Left            =   3285
      Picture         =   "frmMain.frx":336A
      Stretch         =   -1  'True
      Top             =   3435
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgPin 
      Height          =   60
      Index           =   5
      Left            =   5434
      MouseIcon       =   "frmMain.frx":358C
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":3896
      Stretch         =   -1  'True
      Top             =   6620
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgPin 
      Height          =   60
      Index           =   6
      Left            =   5990
      MouseIcon       =   "frmMain.frx":3928
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":3C32
      Stretch         =   -1  'True
      Top             =   6620
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgPin 
      Height          =   60
      Index           =   4
      Left            =   4878
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":3CC4
      Stretch         =   -1  'True
      Top             =   6620
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgPin 
      Height          =   60
      Index           =   3
      Left            =   4322
      MouseIcon       =   "frmMain.frx":3D56
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":4060
      Stretch         =   -1  'True
      Top             =   6620
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgPin 
      Height          =   60
      Index           =   2
      Left            =   3766
      MouseIcon       =   "frmMain.frx":40F2
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":43FC
      Stretch         =   -1  'True
      Top             =   6620
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgPin 
      Height          =   60
      Index           =   1
      Left            =   3210
      MouseIcon       =   "frmMain.frx":448E
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":4798
      Stretch         =   -1  'True
      Top             =   6620
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgPinChange 
      Height          =   375
      Index           =   6
      Left            =   5990
      MouseIcon       =   "frmMain.frx":482A
      MousePointer    =   99  'Custom
      Top             =   6450
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgPinChange 
      Height          =   375
      Index           =   5
      Left            =   5432
      MouseIcon       =   "frmMain.frx":4B34
      MousePointer    =   99  'Custom
      Top             =   6450
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgPinChange 
      Height          =   375
      Index           =   4
      Left            =   4874
      MouseIcon       =   "frmMain.frx":4E3E
      MousePointer    =   99  'Custom
      Top             =   6450
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgPinChange 
      Height          =   375
      Index           =   3
      Left            =   4316
      MouseIcon       =   "frmMain.frx":5148
      MousePointer    =   99  'Custom
      Top             =   6450
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgPinChange 
      Height          =   375
      Index           =   2
      Left            =   3758
      MouseIcon       =   "frmMain.frx":5452
      MousePointer    =   99  'Custom
      Top             =   6450
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgPinChange 
      Height          =   375
      Index           =   1
      Left            =   3200
      MouseIcon       =   "frmMain.frx":575C
      MousePointer    =   99  'Custom
      Top             =   6450
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imghandle 
      Height          =   495
      Index           =   3
      Left            =   4560
      Picture         =   "frmMain.frx":5A66
      Stretch         =   -1  'True
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imghandle 
      Height          =   495
      Index           =   2
      Left            =   3960
      Picture         =   "frmMain.frx":B340
      Stretch         =   -1  'True
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imghandle 
      Height          =   495
      Index           =   1
      Left            =   3360
      Picture         =   "frmMain.frx":10C1A
      Stretch         =   -1  'True
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgCoverClosed 
      Height          =   375
      Left            =   2760
      Picture         =   "frmMain.frx":164F4
      Stretch         =   -1  'True
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCoverOpen 
      Height          =   375
      Left            =   2160
      Picture         =   "frmMain.frx":1B5E1
      Stretch         =   -1  'True
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAbout 
      Height          =   255
      Left            =   6800
      MouseIcon       =   "frmMain.frx":253DC
      MousePointer    =   99  'Custom
      ToolTipText     =   " About M-209 Sim "
      Top             =   50
      Width           =   255
   End
   Begin VB.Image imgHelp 
      Height          =   255
      Left            =   7050
      MouseIcon       =   "frmMain.frx":256E6
      MousePointer    =   99  'Custom
      ToolTipText     =   " Helpfile "
      Top             =   50
      Width           =   255
   End
   Begin VB.Image imgExit 
      Height          =   255
      Left            =   7320
      MouseIcon       =   "frmMain.frx":259F0
      MousePointer    =   99  'Custom
      ToolTipText     =   " Exit M-209 "
      Top             =   50
      Width           =   255
   End
   Begin VB.Image imgCover 
      Height          =   615
      Left            =   0
      MouseIcon       =   "frmMain.frx":25CFA
      MousePointer    =   99  'Custom
      ToolTipText     =   " Open Cover "
      Top             =   7440
      Width           =   8055
   End
   Begin VB.Label lblBarNr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6340
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgBarUp 
      Height          =   615
      Left            =   6360
      MouseIcon       =   "frmMain.frx":26004
      MousePointer    =   99  'Custom
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgBarDn 
      Height          =   855
      Left            =   6360
      MouseIcon       =   "frmMain.frx":2630E
      MousePointer    =   99  'Custom
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPrint 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   975
      Left            =   960
      TabIndex        =   10
      Top             =   5695
      Width           =   255
   End
   Begin VB.Image imgWheelDn 
      Height          =   1215
      Index           =   2
      Left            =   3765
      MouseIcon       =   "frmMain.frx":26618
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image imgWheelDn 
      Height          =   1215
      Index           =   1
      Left            =   3195
      MouseIcon       =   "frmMain.frx":26922
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image imgWheelDn 
      Height          =   1215
      Index           =   3
      Left            =   4320
      MouseIcon       =   "frmMain.frx":26C2C
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image imgWheelDn 
      Height          =   1215
      Index           =   4
      Left            =   4875
      MouseIcon       =   "frmMain.frx":26F36
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image imgWheelDn 
      Height          =   1215
      Index           =   5
      Left            =   5445
      MouseIcon       =   "frmMain.frx":27240
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image imgWheelDn 
      Height          =   1215
      Index           =   6
      Left            =   6000
      MouseIcon       =   "frmMain.frx":2754A
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image imgIndicatorUp 
      Height          =   735
      Left            =   120
      MouseIcon       =   "frmMain.frx":27854
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   615
   End
   Begin VB.Image imgIndicatorDn 
      Height          =   735
      Left            =   120
      MouseIcon       =   "frmMain.frx":27B5E
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   615
   End
   Begin VB.Image imgSelect 
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmMain.frx":27E68
      MousePointer    =   99  'Custom
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image imgAllDn 
      Height          =   375
      Left            =   7200
      MouseIcon       =   "frmMain.frx":28172
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   700
   End
   Begin VB.Image igmAllUp 
      Height          =   375
      Left            =   7200
      MouseIcon       =   "frmMain.frx":2847C
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   700
   End
   Begin VB.Image imgPower 
      Height          =   2835
      Left            =   7230
      MouseIcon       =   "frmMain.frx":28786
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   585
   End
   Begin VB.Image imgWheelUp 
      Height          =   495
      Index           =   6
      Left            =   6000
      MouseIcon       =   "frmMain.frx":28A90
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   495
   End
   Begin VB.Image imgWheelUp 
      Height          =   495
      Index           =   5
      Left            =   5445
      MouseIcon       =   "frmMain.frx":28D9A
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   495
   End
   Begin VB.Image imgWheelUp 
      Height          =   495
      Index           =   4
      Left            =   4875
      MouseIcon       =   "frmMain.frx":290A4
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   495
   End
   Begin VB.Image imgWheelUp 
      Height          =   495
      Index           =   3
      Left            =   4320
      MouseIcon       =   "frmMain.frx":293AE
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   495
   End
   Begin VB.Image imgWheelUp 
      Height          =   495
      Index           =   2
      Left            =   3765
      MouseIcon       =   "frmMain.frx":296B8
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   495
   End
   Begin VB.Image imgWheelUp 
      Height          =   495
      Index           =   1
      Left            =   3195
      MouseIcon       =   "frmMain.frx":299C2
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Index           =   5
      Left            =   5545
      TabIndex        =   8
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Index           =   4
      Left            =   4990
      TabIndex        =   7
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Index           =   3
      Left            =   4435
      TabIndex        =   6
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Index           =   2
      Left            =   3880
      TabIndex        =   5
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Index           =   1
      Left            =   3325
      TabIndex        =   4
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Index           =   6
      Left            =   6100
      TabIndex        =   3
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label lblCounter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1901
      TabIndex        =   2
      Top             =   5850
      Width           =   555
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   230
      TabIndex        =   1
      Top             =   5075
      Width           =   255
   End
   Begin VB.Label lblIndicator 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1575
      Left            =   720
      TabIndex        =   9
      Top             =   5485
      Width           =   255
   End
   Begin VB.Image picTitlebar 
      Height          =   1700
      Left            =   0
      MousePointer    =   15  'Size All
      ToolTipText     =   " Move M-209 "
      Top             =   0
      Width           =   6000
   End
   Begin VB.Label lblInput 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   675
      Width           =   6135
   End
   Begin VB.Label lblOutput 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1125
      Width           =   6135
   End
   Begin VB.Image imgBackGround 
      Height          =   690
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bMoveFrom As Boolean
Private KeyIsDown As Boolean

Private Sub Form_Load()
Dim a As Long
Dim m As Long
Dim rn As Long
' shape rounded form
rn = 96
rn = (rn / iTPPX&) * 15
a = CreateRoundRectRgn(0, 0, Me.Width / iTPPX&, Me.Height / iTPPY&, rn, rn)
m = SetWindowRgn(Me.hwnd, a, True)
DeleteObject m
End Sub

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MENU'S <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub imgAbout_Click()
Call PlaySound(2)
frmInfo.Show (vbModal)
End Sub

Private Sub imgExit_Click()
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = ExitProgram
End Sub

Private Sub imgGallerie_Click()
Call PlaySound(2)
frmGallerie.Show
End Sub

Private Sub imgHelp_Click()
'show helpfile
Call ShowHelpFile
End Sub

Private Sub imgSound_Click()
If gblnSound = True Then
    Call PlaySound(2)
    gblnSound = False
    Me.imgSound.ToolTipText = " Sound On "
    Me.imgSound.Picture = Me.imgSoundOff
    Else
    Me.imgSound.Picture = Me.imgSoundOn
    gblnSound = True
    Me.imgSound.ToolTipText = " Sound Off "
    Call PlaySound(2)
End If
End Sub

' >>>>>>>>>>>>>>>>>>>>>>>>>> BUTTONS AND HANDLES <<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If gstrAutoType = True Then gstrAutoType = False
'use keyboard to enter
If KeyIsDown = True Then Exit Sub
KeyIsDown = True
If CoverOpen = True Then Exit Sub
If KeyCode < 65 Or KeyCode > 90 Then KeyCode = 0: Exit Sub
Indicator = KeyCode - 64
Call PlaySound(2)
Call SetIndicatorView
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyIsDown = False Then Exit Sub
KeyIsDown = False
'interrupt autotyping
If gstrAutoType = True Then gstrAutoType = False: Exit Sub
'use keyboard to enter
Select Case KeyCode
Case 112
    Call ShowHelpFile '(F1)
    Exit Sub
Case 121
    'alignment (F10)
    frmAlign.Show (vbModal)
    Exit Sub
Case 33
    'all wheel 1 up (PAGE UP)
    Call TurnbackWheels
Case 34
    'all wheel 1 down (PAGE DOWN)
    Call AdvanceWheels
    If gstrAutoType = False Or (gstrAutoType = True And frmQuick.cmbSpeed <> "Fast") Then
        Call PlaySound(2)
    End If
Case 46
    'delete text (DEL)
    frmMain.lblOutput.Caption = ""
    frmMain.lblInput.Caption = ""
    gstrClipOutput = ""
    gstrClipInput = ""
    OutLen = 0
Case 45
    'memorize wheels(INS)
    Call SetWheelsMemo
Case 36
    'get memorized wheels
    Call GetWheelsMemo
Case 8
    'reset wheels (BACKSPACE)
    Call ResetAll
    Exit Sub
Case 109
    'indicator -
    Call imgIndicatorDn_MouseUp(1, 1, 1, 1)
    Exit Sub
Case 107
    'indicator +
    Call imgIndicatorUp_MouseUp(1, 1, 1, 1)
    Exit Sub
Case 123 'erase (F12)
    Call DeleteAllSettings
    Exit Sub
End Select

If CoverOpen = True Then
    'set wheel pins
    If KeyCode > 48 And KeyCode < 55 Then Call imgPinChange_MouseUp(KeyCode - 48, 0, 0, 0, 0)
    Exit Sub
End If

Select Case KeyCode
Case 116
    frmClipBoard.Show (vbModal)
    Exit Sub
Case 117
    frmQuick.Show (vbModal)
    Exit Sub
Case 13
    If Indicator = PreviousIndic Then Exit Sub
    Call EncodeChar(Indicator)
    Exit Sub
End Select
If (KeyCode < 65 Or KeyCode > 90) And KeyCode <> 32 Then Exit Sub
If KeyCode = 32 Then
    'when in cipher mode, replace keyboard spaces by Z
    If SetCipher = True Then
        KeyCode = 90
        Else
        KeyCode = 0
        Exit Sub
        End If
    End If
Indicator = KeyCode - 64
Call EncodeChar(Indicator)
End Sub

Private Sub igmAllUp_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
'all wheels plus 1
Call AdvanceWheels
If gstrAutoType = False Or (gstrAutoType = True And frmQuick.cmbSpeed <> "Fast") Then
    Call PlaySound(2)
End If
End Sub

Private Sub imgAllDn_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
'all wheels minus 1
Call TurnbackWheels
End Sub

Private Sub imgReset_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
'reset all wheels and counter
Call ResetAll
End Sub

Private Sub imgPower_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
'turn handle to encipher/decipher a letter
Dim tmp As String
If CoverOpen = True Then Exit Sub
If Indicator = PreviousIndic Then Exit Sub
Call EncodeChar(Indicator)
End Sub

Private Sub imgIndicatorDn_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
'indicator one down (B to A)
Indicator = Indicator + 1
If Indicator > 26 Then Indicator = 1
Call SetIndicatorView
'set power handle free
PreviousIndic = 0
Call PlaySound(2)
End Sub

Private Sub imgIndicatorUp_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
'indicator one up (A to B)
Indicator = Indicator - 1
If Indicator < 1 Then Indicator = 26
Call SetIndicatorView
'set power handle free
PreviousIndic = 0
Call PlaySound(2)
End Sub

Private Sub imgWheelDn_MouseUp(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
'wheel one down (B to A)
Wpos(Index) = Wpos(Index) - 1
If Wpos(Index) < 1 Then Wpos(Index) = Wlenght(Index)
Call SetWheelsView(Index)
Call PlaySound(2)
End Sub

Private Sub imgWheelUp_MouseUp(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
'wheel one up (A to B)
Wpos(Index) = Wpos(Index) + 1
If Wpos(Index) > Wlenght(Index) Then Wpos(Index) = 1
Call SetWheelsView(Index)
Call PlaySound(2)
End Sub

Private Sub imgSelect_Click()
'set cipher/decipher handle
If SetCipher = True Then
    SetCipher = False
    Me.lblSelect.Caption = "D"
    Else
    SetCipher = True
    Me.lblSelect.Caption = "C"
    End If
Call PlaySound(2)
End Sub

Private Sub imgCover_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
'open/close cover
Dim k As Integer
Dim act As Boolean
Call PlaySound(2)
With Me
If CoverOpen = False Then
    CoverOpen = True
    If DpiDefault = True Then
        .Picture = .imgCoverOpen.Picture
        Else
        .imgBackGround = .imgCoverOpen.Picture
        End If
    .lblBarNr.Visible = True
    .imgBarUp.Visible = True
    .imgBarDn.Visible = True
    .imgLug1.Visible = True
    .imgLug2.Visible = True
    .imgReset.Visible = False
    .imgCover.ToolTipText = " Close Cover "
    For k = 1 To 8
        .imgLug(k).Visible = True
    Next
    'set view pins
    For k = 1 To 6
        .imgPin(k).Visible = True
        Call SetWheelsView(k)
    Next k
    For k = 1 To 6
        .imgPinChange(k).Visible = True
    Next
    CurrentBar = 1
    SetBarView (CurrentBar)
    Else '>>>>>>>>cover closed
    CoverOpen = False
    If DpiDefault = True Then
        .Picture = .imgCoverClosed.Picture
        Else
        .imgBackGround = .imgCoverClosed.Picture
        End If
    .lblBarNr.Visible = False
    .imgBarUp.Visible = False
    .imgBarDn.Visible = False
    .imgLug1.Visible = False
    .imgLug2.Visible = False
    .imgReset.Visible = True
    .imgCover.ToolTipText = " Open Cover "
    For k = 1 To 8
        .imgLug(k).Visible = False
    Next
    For k = 1 To 6
        .imgPin(k).Visible = False
        Call SetWheelsView(k)
    Next k
    For k = 1 To 6
        .imgPinChange(k).Visible = False
    Next
    End If
End With
Call SetIndicatorView
End Sub

' >>>>>>>>>>>>>>>>>>>>>>>>>> INTERNAL SETTINGS <<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub imgPinChange_MouseUp(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
'change position of a pin
If GetPin(Wpos(Index), Index) = "-" Then
    Call SetPin(Wpos(Index), Index, True)
    Call SetPinView(Wpos(Index), Index, True)
    Else
    Call SetPin(Wpos(Index), Index, False)
    Call SetPinView(Wpos(Index), Index, False)
    End If
Call PlaySound(2)
End Sub

Private Sub imgLug_MouseUp(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
'move a lug on drum
Dim ImgToVal(8) As Integer
ImgToVal(1) = 1
ImgToVal(2) = 0
ImgToVal(3) = 2
ImgToVal(4) = 3
ImgToVal(5) = 4
ImgToVal(6) = 5
ImgToVal(7) = 0
ImgToVal(8) = 6

If ActiveLug = 0 Then ActiveLug = Index: Call PlaySound(2): Exit Sub
If ActiveLug = Index Then Call PlaySound(2): Exit Sub
Call GetLugPositions(CurrentBar)
If Index = LugPos1 Or Index = LugPos2 Then
    ActiveLug = Index
    Call PlaySound(2)
    Exit Sub
    Else
    If ActiveLug = LugPos1 And Index < LugPos2 Then
        LugPos1 = Index
        ActiveLug = Index
        Call PlaySound(2)
    ElseIf ActiveLug = LugPos2 And Index > LugPos1 Then
        LugPos2 = Index
        ActiveLug = Index
        Call PlaySound(2)
    Else
        Exit Sub
    End If
End If
Bar(CurrentBar) = Trim(Str(ImgToVal(LugPos1))) & Trim(Str(ImgToVal(LugPos2)))
frmMain.imgLug1.Left = LugLeft(LugPos1)
frmMain.imgLug2.Left = LugLeft(LugPos2)
End Sub

Private Sub imgBarDn_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
'turn drum, one bar down
CurrentBar = CurrentBar - 1
If CurrentBar < 1 Then CurrentBar = 1: Exit Sub
Call SetBarView(CurrentBar)
ActiveLug = 0
Me.lblBarNr.Caption = Trim(Str(CurrentBar))
Call PlaySound(2)
End Sub


Private Sub imgBarUp_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
'turn drum, one bar up
CurrentBar = CurrentBar + 1
If CurrentBar > 27 Then CurrentBar = 27: Exit Sub
Call SetBarView(CurrentBar)
ActiveLug = 0
Me.lblBarNr.Caption = Trim(Str(CurrentBar))
Call PlaySound(2)
End Sub



' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MOVE FORM <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub picTitleBar_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
'get mouse movement
Dim POINT As POINTAPI
GetCursorPos POINT
LastPoint.X = POINT.X
LastPoint.Y = POINT.Y
bMoveFrom = True
End Sub

Private Sub picTitleBar_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
'if mouse is down, move the form
Dim iDX As Long, iDY As Long
Dim POINT As POINTAPI
If Not bMoveFrom Then
    Exit Sub
    End If
GetCursorPos POINT
iDX& = (POINT.X - LastPoint.X) * iTPPX&
iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
LastPoint.X = POINT.X
LastPoint.Y = POINT.Y
Me.Move Me.Left + iDX&, Me.Top + iDY&
End Sub

Private Sub picTitleBar_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
'release form
bMoveFrom = False
End Sub

Private Sub Timer1_Timer()
' display flashing message at start pointing to readme
Static flashCount As Integer
flashCount = flashCount + 1
Select Case flashCount
    Case 1, 3, 5, 7
        Me.lblInfo.Visible = False
    Case 2, 4, 6, 8
        Me.lblInfo.Visible = True
    Case 24
        Me.Timer1.Enabled = False
        Me.lblInfo.Visible = False
End Select
End Sub

