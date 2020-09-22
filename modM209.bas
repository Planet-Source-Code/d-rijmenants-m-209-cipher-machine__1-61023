Attribute VB_Name = "modM209"
Option Explicit

Public W1(25) As Byte '26 letters A-Z
Public W2(24) As Byte '25 letters A-Z exept W
Public W3(22) As Byte '23 letters A-X exept W
Public W4(20) As Byte '21 letters A-U
Public W5(18) As Byte '19 letters A-S
Public W6(16) As Byte '17 letters A-Q

Public Wstring(6) As String
Public Wlenght(6) As Byte
Public WpinPos(6) As Integer
Public Wpins(6) As String
Public Wpos(6) As Integer
Public Wmemo(6) As Integer

Public Bar(27) As String
Public LugLeft(8) As Integer
Public PinLeft(6) As Integer
Public LugPos1 As Integer
Public LugPos2 As Integer
Public CurrentBar As Integer
Public ActiveLug As Integer

Public Indicator As Integer
Public PreviousIndic As Integer

Public Counter As Integer
Public OutLen As Long

Public SetCipher As Boolean
Public gblnSound As Boolean

Public gstrClipInput As String
Public gstrClipOutput As String

Public CoverOpen As Boolean
Public gstrExitVal As String
Public DpiDefault As Boolean
Public Vdefault As Integer
Public Hdefault As Integer


Public gstrStopFlag As Boolean
Public gstrAutoType As Boolean

'cursor functions to move forms
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public LastPoint As POINTAPI
Public iTPPY As Long
Public iTPPX As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type

'sound api
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1     ' Play asynchronously
Public Const SND_NODEFAULT = &H2 ' Don't use default sound
Public Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file
Public SoundBuffer As String

'time function
Public Declare Function GetTickCount Lib "kernel32" () As Long

'form shape functions
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const PrintWheel = "ZYXWVUTSRQPONMLKJIHGFEDCBA"

Public Sub Main()
Dim k As Integer
Dim j As Integer

iTPPX& = Screen.TwipsPerPixelX
iTPPY& = Screen.TwipsPerPixelY

gblnSound = True
CurrentBar = 1
CoverOpen = False
SetCipher = True

Load frmMain
Vdefault = 8050
Hdefault = 8050
With frmMain
    .Height = Vdefault
    .Width = Hdefault
    .imgBackGround.Height = Vdefault
    .imgBackGround.Width = Hdefault
    'check for default dpi settings or corrections
    Call SetDpiCorrection
    .imgBackGround.Picture = .imgCoverClosed.Picture
    .imgPower.Picture = .imghandle(1).Picture
    .Timer1.Enabled = True
End With

'indications on wheels
Wstring(1) = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" '26 letters A-Z
Wstring(2) = "ABCDEFGHIJKLMNOPQRSTUVXYZ"  '25 letters A-Z exept W
Wstring(3) = "ABCDEFGHIJKLMNOPQRSTUVX"    '23 letters A-X exept W
Wstring(4) = "ABCDEFGHIJKLMNOPQRSTU"      '21 letters A-U
Wstring(5) = "ABCDEFGHIJKLMNOPQRS"        '19 letters A-S
Wstring(6) = "ABCDEFGHIJKLMNOPQ"          '17 letters A-Q

'lenght of each wheel
For k = 1 To 6
    Wlenght(k) = Len(Wstring(k))
Next k

'set all wheel to "A"
Indicator = 1
For k = 1 To 6
    Wpos(k) = 1
    Call SetWheelsView(k)
Next
'set indicator
Call SetIndicatorView

'lugs on bars (0 = inactive, 1-6 = active)
For k = 1 To 27
Bar(k) = "00"
Next k

'X-positions for lugs
LugLeft(1) = 3240
LugLeft(2) = 3520
LugLeft(3) = 3875
LugLeft(4) = 4440
LugLeft(5) = 4995
LugLeft(6) = 5600
LugLeft(7) = 5870
LugLeft(8) = 6100

'x-position pins (inactive)
PinLeft(1) = 3220
PinLeft(2) = 3766
PinLeft(3) = 4322
PinLeft(4) = 4878
PinLeft(5) = 5434
PinLeft(6) = 5990

Call LoadConfiguration
Call LoadAlignment

frmMain.Timer1.Enabled = True
frmMain.Show
    
End Sub

Public Function CodeLetter(ByRef Letter As Integer) As String
Dim PinString As String
Dim k As Integer
Dim Offset As Integer
Dim PrintPos As Integer

'get pin positions, relative to viewed letter
WpinPos(1) = Wpos(1) + 15: If WpinPos(1) > 26 Then WpinPos(1) = WpinPos(1) - 26
WpinPos(2) = Wpos(2) + 14: If WpinPos(2) > 25 Then WpinPos(2) = WpinPos(2) - 25
WpinPos(3) = Wpos(3) + 13: If WpinPos(3) > 23 Then WpinPos(3) = WpinPos(3) - 23
WpinPos(4) = Wpos(4) + 12: If WpinPos(4) > 21 Then WpinPos(4) = WpinPos(4) - 21
WpinPos(5) = Wpos(5) + 11: If WpinPos(5) > 19 Then WpinPos(5) = WpinPos(5) - 19
WpinPos(6) = Wpos(6) + 10: If WpinPos(6) > 17 Then WpinPos(6) = WpinPos(6) - 17

'get pins
For k = 1 To 6
    If Mid(Wpins(k), WpinPos(k), 1) <> "-" Then
        PinString = PinString & Trim(Str(k))
        End If
Next k

'check lugs to pins
If Len(PinString) <> 0 Then
    For k = 1 To 27
        If InStr(1, PinString, Left(Bar(k), 1)) <> 0 Or InStr(1, PinString, Right(Bar(k), 1)) <> 0 Then
            Offset = Offset + 1
        End If
    Next
End If
       
'get offset letter on print drum
PrintPos = Letter - Offset
If PrintPos < 1 Then PrintPos = PrintPos + 26
If PrintPos < 1 Then PrintPos = PrintPos + 26

CodeLetter = Mid(PrintWheel, PrintPos, 1)
Indicator = PrintPos
PreviousIndic = Indicator
Call AdvanceWheels
Call SetIndicatorView
End Function

Public Sub AdvanceWheels()
'advance all wheels one position
Dim k As Integer
For k = 1 To 6
    Wpos(k) = Wpos(k) + 1
    If Wpos(k) > Wlenght(k) Then Wpos(k) = 1
    Call SetWheelsView(k)
Next
Counter = Counter + 1
If Counter > 9999 Then Counter = 0
frmMain.lblCounter.Caption = Format(Counter, "0000")
End Sub

Public Sub TurnbackWheels()
'turn back all wheels one position
Dim k As Integer
For k = 1 To 6
    Wpos(k) = Wpos(k) - 1
    If Wpos(k) < 1 Then Wpos(k) = Wlenght(k)
    Call SetWheelsView(k)
Next
Counter = Counter - 1
If Counter < 0 Then Counter = 9999
frmMain.lblCounter.Caption = Format(Counter, "0000")
Call PlaySound(2)
End Sub

Public Sub SetWheelsView(wheel As Integer)
'set view of all wheels
Dim k As Integer
Dim j As Integer
Dim P As Integer
Dim BB As Integer

If CoverOpen = False Then
    BB = 2
    Else
    BB = 6
    End If
'set wheels view
frmMain.lblWindow(wheel) = ""
If CoverOpen = False Then frmMain.lblWindow(wheel) = vbCrLf & vbCrLf & vbCrLf & vbCrLf
For j = Wpos(wheel) + BB To Wpos(wheel) - 2 Step -1
    If j < 1 Then
        P = j + Wlenght(wheel)
    ElseIf j > Wlenght(wheel) Then
        P = j - Wlenght(wheel)
    Else
        P = j
    End If
    frmMain.lblWindow(wheel) = frmMain.lblWindow(wheel) & Mid(Wstring(wheel), P, 1) & vbCrLf
Next j

If CoverOpen = False Then Exit Sub

If GetPin(Wpos(wheel), wheel) = "-" Then
    Call SetPinView(Wpos(wheel), wheel, False)
    Else
    Call SetPinView(Wpos(wheel), wheel, True)
    End If

End Sub

Public Sub SetIndicatorView()
'set indicator view
Dim j As Integer
Dim P As Integer
Dim BB As Integer

frmMain.lblIndicator = ""
For j = Indicator - 5 To Indicator + 1
    If j < 1 Then
        P = j + 26
    ElseIf j > 26 Then
        P = j - 26
    Else
        P = j
    End If
    frmMain.lblIndicator = frmMain.lblIndicator & Chr(P + 64) & vbCrLf
Next j

If CoverOpen = False Then
    BB = 0
    Else
    BB = 2
    End If
'set indicator print view
frmMain.lblPrint = ""
For j = Indicator - 2 To Indicator + BB
    If j < 1 Then
        P = j + 26
    ElseIf j > 26 Then
        P = j - 26
    Else
        P = j
    End If
    frmMain.lblPrint = frmMain.lblPrint & Mid(PrintWheel, P, 1) & vbCrLf
Next j
frmMain.lblPrint.Refresh
frmMain.lblIndicator.Refresh
End Sub

Public Sub GetLugPositions(curBar As Integer)
'get lug positions on a given bar
Dim L1 As Integer
Dim L2 As Integer
Dim ValToImg(6)
ValToImg(1) = 1
ValToImg(2) = 3
ValToImg(3) = 4
ValToImg(4) = 5
ValToImg(5) = 6
ValToImg(6) = 8

L1 = Val(Left(Bar(curBar), 1))
L2 = Val(Right(Bar(curBar), 1))
If L1 = 0 Then
    LugPos1 = 2
    Else
    LugPos1 = ValToImg(L1)
    End If
If L2 = 0 Then
    LugPos2 = 7
    Else
    LugPos2 = ValToImg(L2)
    End If
End Sub

Public Sub SetBarView(curBar As Integer)
'set position of lugs on the given bar
Call GetLugPositions(curBar)
frmMain.imgLug1.Left = LugLeft(LugPos1)
frmMain.imgLug2.Left = LugLeft(LugPos2)
frmMain.lblBarNr.Caption = Str(Trim(curBar))
End Sub

Public Sub SetPin(Pos As Integer, wheel As Integer, Active As Boolean)
'save a pin
If Active = True Then
    Mid(Wpins(wheel), Pos, 1) = Mid(Wstring(wheel), Pos, 1)
    Else
    Mid(Wpins(wheel), Pos, 1) = "-"
    End If
End Sub

Public Function GetPin(Pos As Integer, wheel As Integer) As String
'read a pin
GetPin = Mid(Wpins(wheel), Pos, 1)
End Function

Public Sub SetPinView(Pos As Integer, wheel As Integer, Active As Boolean)
'set a pin view
If Active = True Then
    frmMain.imgPin(wheel).Left = PinLeft(wheel) + 370
    Else
    frmMain.imgPin(wheel).Left = PinLeft(wheel)
    End If
End Sub

Public Sub EncodeChar(Key As Integer)
'encode a character
Dim kIn As String
Dim tmp As String
If gstrAutoType = False Or (gstrAutoType = True And frmQuick.cmbSpeed <> "Fast") Then
    Call PlaySound(1)
End If
Call SetIndicatorView
Call TurnHandle
kIn = Chr(Key + 64)
tmp = CodeLetter(Key)
OutLen = OutLen + 1
If SetCipher = False Then
    'on enciphering, Z is used as a space !
    'so: on deciphering, replace Z by a space
    If tmp = "Z" Then tmp = " "
    End If
gstrClipOutput = gstrClipOutput & tmp
gstrClipInput = gstrClipInput & kIn
With frmMain
.lblOutput.Caption = .lblOutput.Caption & tmp
.lblInput.Caption = .lblInput.Caption & kIn
If SetCipher = True And OutLen <> 1 And OutLen Mod 5 = 0 Then
    .lblOutput.Caption = .lblOutput.Caption & " "
    .lblInput.Caption = .lblInput.Caption & " "
    End If
If Len(.lblOutput.Caption) > 57 Then
    .lblOutput.Caption = Right(.lblOutput.Caption, 57)
    .lblInput.Caption = Right(.lblInput.Caption, 57)
    End If
End With
End Sub

Public Sub AutoTyping()
'start coding/decoding the tet
Dim tmpQuick As String
Dim tmpChar As Integer
Dim i As Long
Dim tm As Long
'delet all but alphabet
tmpQuick = frmQuick.txtQuick.Text
If tmpQuick = "" Then Exit Sub
gstrAutoType = True
Select Case frmQuick.cmbSpeed.Text
Case "Slow"
    tm = 2000
Case "Normal"
    tm = 500
Case "Fast"
    tm = 0
End Select
If SetCipher = True Then tmpQuick = Trim(tmpQuick)
For i = 1 To Len(tmpQuick)
    'if cipher mode, replace spaces by 'Z'
    If SetCipher = True And Mid(tmpQuick, i, 1) = " " Then
        tmpChar = 26
        Else
        tmpChar = Asc(UCase(Mid(tmpQuick, i, 1))) - 64
        End If
    If tmpChar > 0 And tmpChar < 27 Then
        'encode
        PauzeTime (tm)
        Indicator = tmpChar
        If tm = 2000 Then
            Call SetIndicatorView
            PauzeTime (2000)
            End If
        Call EncodeChar(Indicator)
        DoEvents
        If gstrAutoType = False Then
            MsgBox "Auto Typing aborted.", vbInformation, " M-209"
            Exit For
            End If
    End If
Next
gstrAutoType = False
End Sub

Public Sub PauzeTime(TimeToWait As Long)
'pauze the program
Dim currTime As Long
Dim passedTime As Long
currTime = GetTickCount()
Do
    passedTime = Abs(GetTickCount() - currTime)
Loop While passedTime < TimeToWait
End Sub

Public Sub ResetAll()
'reset all wheels and counter
Dim k As Integer
Indicator = 1
Call SetIndicatorView
For k = 1 To 6
    Wpos(k) = 1
    Call SetWheelsView(k)
Next
frmMain.lblOutput.Caption = ""
frmMain.lblInput.Caption = ""
gstrClipOutput = ""
gstrClipInput = ""
Counter = 0
OutLen = 0
frmMain.lblCounter.Caption = Format(Counter, "0000")
Call PlaySound(2)
End Sub

Public Sub TurnHandle()
'turn power handle
If gstrAutoType = True And frmQuick.cmbSpeed.Text = "Fast" Then Exit Sub
With frmMain
PauzeTime (75)
.imgPower.Picture = .imghandle(2).Picture
.imgPower.Refresh
PauzeTime (75)
.imgPower.Picture = .imghandle(3).Picture
.imgPower.Refresh
PauzeTime (75)
.imgPower.Picture = .imghandle(2).Picture
.imgPower.Refresh
PauzeTime (75)
.imgPower.Picture = .imghandle(1).Picture
.imgPower.Refresh
End With
End Sub

Public Sub SetWheelsMemo()
'save current external message indicator
Dim k As Integer
If Wmemo(1) <> 0 Then
    k = MsgBox("Do You want to overwrite the currently memorized Message Indicator?", vbQuestion + vbYesNo, "M-209 Memory")
    If k = vbNo Then Exit Sub
    End If
Call PlaySound(2)
For k = 1 To 6
    Wmemo(k) = Wpos(k)
Next k
End Sub

Public Sub GetWheelsMemo()
'get memorized message indicator
Dim k As Integer
If Wmemo(1) = 0 Then Exit Sub
Call PlaySound(2)
Call ResetAll
For k = 1 To 6
    Wpos(k) = Wmemo(k)
    Call SetWheelsView(k)
Next k
End Sub

Public Sub ShowHelpFile()
'show the helpfile
Call PlaySound(2)
frmMain.Dialog1.InitDir = App.Path
frmMain.Dialog1.HelpFile = "M-209 HELP.HLP"
frmMain.Dialog1.HelpCommand = cdlHelpContents
frmMain.Dialog1.ShowHelp
End Sub

Public Sub PlaySound(aSound As Integer)
'play sound
Dim Ret
Select Case aSound
Case 0
    Exit Sub
Case 1
    If gblnSound = False Then Exit Sub
    SoundBuffer = StrConv(LoadResData(1, "Sounds"), vbUnicode)
Case 2
    If gblnSound = False Then Exit Sub
    SoundBuffer = StrConv(LoadResData(2, "Sounds"), vbUnicode)
Case 3
    SoundBuffer = StrConv(LoadResData(3, "Sounds"), vbUnicode)
End Select
Ret = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
End Sub

Public Function ExitProgram() As Boolean
'exit the program
Call PlaySound(3)
frmExit.Show (vbModal)
If gstrExitVal = "cancel" Or gstrExitVal = "" Then ExitProgram = True: Exit Function
If gstrExitVal = "save" Then
    Call SaveConfiguration
ElseIf gstrExitVal = "erase" Then
    Call EraseConfiguration
End If
Unload frmClipBoard
Unload frmExit
Unload frmGallerie
Unload frmInfo
Unload frmQuick
Unload frmAlign
ExitProgram = False
End
End Function

Public Sub SaveConfiguration()
'save current key settings to registry
Dim k As Integer
Dim tmp As String
For k = 1 To 27
    tmp = tmp & Bar(k)
Next k
SaveSetting App.EXEName, "config", "lugs", tmp
SaveSetting App.EXEName, "config", "W1", Wpins(1)
SaveSetting App.EXEName, "config", "W2", Wpins(2)
SaveSetting App.EXEName, "config", "W3", Wpins(3)
SaveSetting App.EXEName, "config", "W4", Wpins(4)
SaveSetting App.EXEName, "config", "W5", Wpins(5)
SaveSetting App.EXEName, "config", "W6", Wpins(6)
End Sub

Public Sub LoadConfiguration()
'load key settings from registry
Dim tmp As String
Dim k As Integer

tmp = GetSetting(App.EXEName, "config", "lugs", "")
If Len(tmp) <> 54 Then tmp = String(54, "0")
For k = 1 To 27
    Bar(k) = Mid(tmp, ((k - 1) * 2) + 1, 2)
Next

For k = 1 To 6
    tmp = GetSetting(App.EXEName, "config", "W" & Trim(Str(k)), "")
    If Len(tmp) = Wlenght(k) Then
        Wpins(k) = tmp
        Else
        Wpins(k) = String(Wlenght(k), "-")
        End If
Next k
End Sub

Public Sub EraseConfiguration()
'erase all registry settings
SaveSetting App.EXEName, "config", "lugs", ""
SaveSetting App.EXEName, "config", "W1", ""
SaveSetting App.EXEName, "config", "W2", ""
SaveSetting App.EXEName, "config", "W3", ""
SaveSetting App.EXEName, "config", "W4", ""
SaveSetting App.EXEName, "config", "W5", ""
SaveSetting App.EXEName, "config", "W6", ""
End Sub


Public Sub DeleteAllSettings()
Dim k As Integer
k = MsgBox("Are you sure you want to erase all pin and lug settings?", vbYesNo + vbDefaultButton2 + vbExclamation, "M-209")
If k = vbNo Then Exit Sub
'clear lugs on bars
For k = 1 To 27
Bar(k) = "00"
Next k
For k = 1 To 6
    Wpins(k) = String(Wlenght(k), "-")
Next k
Call ResetAll
Call SetBarView(1)
End Sub

Public Sub SetDpiCorrection()
With frmMain
If iTPPY& <> 15 Then
    DpiDefault = False
    If GetSetting(App.EXEName, "config", "DPIfirst") <> "1" Then
        'warn dpi alignment
        SaveSetting App.EXEName, "config", "DPIfirst", "1"
        MsgBox "You are not using the default 96 Dpi (100%) screen settings. The alignment of graphics and text in this program could be distorted when using another Dpi setting." & vbCrLf & vbCrLf & "Please press F10 in this program to adjust the text alignment to your dpi settings.", vbExclamation, "M-209 Dpi Settings Change"
        End If
    'adjust settings
    .imgBackGround.Visible = True
    .imgBackGround.Height = Vdefault
    If CoverOpen = True Then
        .imgBackGround.Picture = .imgCoverOpen.Picture
        Else
        .imgBackGround.Picture = .imgCoverClosed.Picture
    End If
    Else
    'use form background
    DpiDefault = True
    If GetSetting(App.EXEName, "config", "DPIfirst") = "1" Then
        'warn dpi alignment
        SaveSetting App.EXEName, "config", "DPIfirst", "0"
        MsgBox "You have changed to the default 96 Dpi (100%) screen settings. The alignment of graphics and text in this program could be distorted when using another Dpi setting." & vbCrLf & vbCrLf & "Please press F10 in this program to adjust the text alignment to your dpi settings.", vbExclamation, "M-209 Dpi Settings Change"
        End If
    .imgBackGround.Visible = False
    .Height = Vdefault
    If CoverOpen = True Then
        .Picture = .imgCoverOpen.Picture
        Else
        .Picture = .imgCoverClosed.Picture
    End If
    
End If
End With
End Sub

Public Sub LoadAlignment()
Dim k As Integer
frmMain.lblIndicator.Top = Val(GetSetting(App.EXEName, "config", "Vind", "5485"))
If frmMain.lblIndicator.Top = 0 Then frmMain.lblIndicator.Top = 5485

frmMain.lblPrint.Top = Val(GetSetting(App.EXEName, "config", "Vprt"))
If frmMain.lblPrint.Top = 0 Then frmMain.lblPrint.Top = 5695

frmMain.lblCounter.Top = Val(GetSetting(App.EXEName, "config", "Vcnt"))
If frmMain.lblCounter.Top = 0 Then frmMain.lblCounter.Top = 5850

frmMain.lblWindow(1).Top = Val(GetSetting(App.EXEName, "config", "Vwhl"))
If frmMain.lblWindow(1).Top = 0 Then frmMain.lblWindow(1).Top = 5280
For k = 2 To 6
    frmMain.lblWindow(k).Top = frmMain.lblWindow(1).Top
Next k
End Sub

Public Sub SaveAlignment()
SaveSetting App.EXEName, "config", "Vind", frmMain.lblIndicator.Top
SaveSetting App.EXEName, "config", "Vprt", frmMain.lblPrint.Top
SaveSetting App.EXEName, "config", "Vcnt", frmMain.lblCounter.Top
SaveSetting App.EXEName, "config", "Vwhl", frmMain.lblWindow(1).Top
End Sub

