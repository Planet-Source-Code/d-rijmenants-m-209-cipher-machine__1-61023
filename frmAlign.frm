VERSION 5.00
Begin VB.Form frmAlign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Text Alignment"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   50
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   5055
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   600
      TabIndex        =   9
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   7
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAlign.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Wheels (align with A)"
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   14
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Counter"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   13
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Indicator (align right side Z with the left side Y)"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   12
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Indicator (align left side with A)"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   11
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "frmAlign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Call LoadAlignment
Me.Hide
End Sub

Private Sub cmdOk_Click()
Call SaveAlignment
Me.Hide
End Sub

Private Sub cmdPlus_Click(Index As Integer)
Dim k As Integer
Select Case Index
Case 0
    frmMain.lblIndicator.Top = frmMain.lblIndicator.Top + 10
Case 1
    frmMain.lblPrint.Top = frmMain.lblPrint.Top + 10
Case 2
    frmMain.lblCounter.Top = frmMain.lblCounter.Top + 10
Case 3
    frmMain.lblWindow(1).Top = frmMain.lblWindow(1).Top + 10
    For k = 2 To 6
        frmMain.lblWindow(k).Top = frmMain.lblWindow(1).Top
    Next k
End Select
End Sub

Private Sub cmdMin_Click(Index As Integer)
Dim k As Integer
Select Case Index
Case 0
    frmMain.lblIndicator.Top = frmMain.lblIndicator.Top - 10
Case 1
    frmMain.lblPrint.Top = frmMain.lblPrint.Top - 10
Case 2
    frmMain.lblCounter.Top = frmMain.lblCounter.Top - 10
Case 3
    frmMain.lblWindow(1).Top = frmMain.lblWindow(1).Top - 10
    For k = 2 To 6
        frmMain.lblWindow(k).Top = frmMain.lblWindow(1).Top
    Next k
End Select
End Sub

Private Sub cmdReset_Click()
SaveSetting App.EXEName, "config", "Vind", ""
SaveSetting App.EXEName, "config", "Vprt", ""
SaveSetting App.EXEName, "config", "Vcnt", ""
SaveSetting App.EXEName, "config", "Vwhl", ""
Call LoadAlignment
End Sub

Private Sub Form_Activate()
Call ResetAll
Indicator = 1
Call SetIndicatorView
End Sub
