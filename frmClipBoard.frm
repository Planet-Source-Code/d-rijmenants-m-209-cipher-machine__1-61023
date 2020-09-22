VERSION 5.00
Begin VB.Form frmClipBoard 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " M-209 Smart Clipboard"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply &New Format"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox txtFormat 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4215
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   240
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtLF 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   175
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "5"
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton optOutput 
         BackColor       =   &H00C0C0C0&
         Caption         =   "M-209 &Output"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optInput 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Typed &Input"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   650
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Groups per line"
         Height          =   255
         Left            =   645
         TabIndex        =   8
         Top             =   1020
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&To Clipboard"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "frmClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strClipText As String

Private Sub cmdApply_Click()
Call ApplyFormat
Me.txtFormat.Text = strClipText
End Sub

Private Sub Form_Activate()
Me.cmdOk.SetFocus
Call ApplyFormat
Me.txtFormat.Text = strClipText
End Sub

Private Sub Form_Load()
Me.optOutput.Value = True
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText Me.txtFormat.Text
Me.Hide
End Sub

Private Sub ApplyFormat()
'apply new format
On Error Resume Next
If gstrClipInput = "" Then Exit Sub
Screen.MousePointer = 11
'check wath to copy
If Me.optInput.Value = True Then
    'show input
    strClipText = gstrClipInput
    If SetCipher = True Then
        'cipher mode, input is clear text
        strClipText = MakeText(gstrClipInput)
        Else
        'decipher mode, input is code
        strClipText = MakeGroups(gstrClipInput)
        End If
    Else
    'show output
    strClipText = gstrClipOutput
    If SetCipher = True Then
        'cipher mode, output is code
        strClipText = MakeGroups(gstrClipOutput)
        Else
        strClipText = MakeText(gstrClipOutput)
        'decipher mode, output is clear text
        End If
    End If
Me.txtFormat.Text = strClipText
Screen.MousePointer = 0
End Sub


Private Function MakeGroups(aText As String) As String

Dim k As Long
Dim strTmp As String
Dim Groups As Integer
Dim G As Integer

Me.txtLF.Enabled = True
Me.txtLF.BackColor = &HFFFFFF

strTmp = ""
Groups = 1
G = 1
For k = 1 To Len(aText)
    G = G + 1
    strTmp = strTmp & Mid(aText, k, 1)
    If G = 6 Then
        G = 1
        strTmp = strTmp & " "
        Groups = Groups + 1
        If Groups = Val(Me.txtLF) + 1 Then strTmp = strTmp & vbCrLf: Groups = 1
    End If
Next k
MakeGroups = strTmp
End Function

Private Function MakeText(aText As String) As String
Dim k As Long
Dim tmpChar As String
Dim strTmp As String

Me.txtLF.Enabled = False
Me.txtLF.BackColor = &HC0C0C0

strTmp = ""
For k = 1 To Len(aText)
    tmpChar = Mid(aText, k, 1)
    If tmpChar <> "Z" Then
        strTmp = strTmp & tmpChar
        Else
        'replace letter Z by a space
        strTmp = strTmp & " "
    End If
Next
MakeText = strTmp
End Function

Private Sub optInput_Click()
Call ApplyFormat
End Sub

Private Sub optOutput_Click()
Call ApplyFormat
End Sub

Private Sub txtLF_KeyPress(KeyAscii As Integer)
'limit input groups to figures
Select Case KeyAscii
Case 8, 9
    Exit Sub
Case Is < 48, Is > 57
    KeyAscii = 0
End Select
End Sub

Private Sub txtLF_GotFocus()
Me.txtLF.SelStart = 0
Me.txtLF.SelLength = Len(Me.txtLF.Text)
End Sub




