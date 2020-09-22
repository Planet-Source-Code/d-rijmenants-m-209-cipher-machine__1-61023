VERSION 5.00
Begin VB.Form frmExit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Exit M-209"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   Icon            =   "frmExit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdErase 
      Caption         =   "&Erase All"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdKeep 
      Caption         =   "&Keep Old Settings"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save New Settings"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmExit.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Label1.Caption = "You are about to exit the M-209 Simulator." & vbCrLf & vbCrLf & "Do you want to save the current machine settings?"
End Sub

Private Sub Form_Activate()
Me.cmdSave.SetFocus
End Sub

Private Sub cmdCancel_Click()
Me.Hide
gstrExitVal = "cancel"
End Sub

Private Sub cmdErase_Click()
Me.Hide
gstrExitVal = "erase"
End Sub

Private Sub cmdKeep_Click()
Me.Hide
gstrExitVal = "keep"
End Sub

Private Sub cmdSave_Click()
Me.Hide
gstrExitVal = "save"
End Sub
