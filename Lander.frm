VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Space Moonlander"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lander.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4545
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.Timer WarningBeep 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   600
         Top             =   960
      End
      Begin VB.Timer JoystickTimer 
         Interval        =   1
         Left            =   120
         Top             =   960
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   600
         Top             =   480
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   120
         Top             =   480
      End
      Begin VB.PictureBox Ship 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   560
         Left            =   1320
         Picture         =   "Lander.frx":030A
         ScaleHeight     =   525
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   3160
         Width           =   520
      End
      Begin VB.PictureBox Shape1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1080
         ScaleHeight     =   105
         ScaleWidth      =   945
         TabIndex        =   16
         Top             =   3720
         Width           =   975
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -6360
         Picture         =   "Lander.frx":1354
         ScaleHeight     =   825
         ScaleWidth      =   10425
         TabIndex        =   17
         Top             =   3720
         Width           =   10455
      End
      Begin VB.Image ship2 
         Height          =   360
         Left            =   2760
         Picture         =   "Lander.frx":45886
         Stretch         =   -1  'True
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image TheLander 
         Height          =   405
         Left            =   2760
         Picture         =   "Lander.frx":45B90
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image T1 
         Height          =   480
         Left            =   2640
         Picture         =   "Lander.frx":46BDA
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image T3 
         Height          =   480
         Left            =   2640
         Picture         =   "Lander.frx":46EE4
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image T2 
         Height          =   480
         Left            =   2640
         Picture         =   "Lander.frx":471EE
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image RocketCrash 
         Height          =   360
         Left            =   2760
         Picture         =   "Lander.frx":474F8
         Stretch         =   -1  'True
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "G A M E   O V E R"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   2520
         Picture         =   "Lander.frx":48542
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.Label lblHighScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "High: 0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Speed:"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Image life3 
      Height          =   255
      Left            =   4200
      Picture         =   "Lander.frx":49B8C
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   240
   End
   Begin VB.Image life2 
      Height          =   255
      Left            =   3900
      Picture         =   "Lander.frx":4A456
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   240
   End
   Begin VB.Image life1 
      Height          =   255
      Left            =   3600
      Picture         =   "Lander.frx":4AD20
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score: 0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Level 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Speed: 0"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblAlt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alt: 0"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New Game"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblLand 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LAND"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblFuel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F U E L"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblThrust 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T H R U S T"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetTickCount& Lib "kernel32" ()

Private Const SND_SYNC = &H1
Private Const SND_ASYNC = &H2
Private Const SND_LOOP = &H19
Private Const SND_STOP = &H5

Private Const Fuel1 = 200
Private Const Fuel2 = 200
Private Const Fuel3 = 150
Private Const Fuel4 = 100
Private Const Fuel5 = 80
Private Const Fuel6 = 80
Private Const Fuel7 = 80
Private Const Fuel8 = 90
Private Const fuel9 = 90
Private Const fuel10 = 80



Private Const Gap1 = 35
Private Const Gap2 = 35
Private Const Gap3 = 25
Private Const Gap4 = 10
Private Const Gap5 = 6
Private Const Gap6 = 10
Private Const Gap7 = 6
Private Const Gap8 = 10
Private Const gap9 = 10
Private Const gap10 = 6




Private Const ShipTop = 240
Private Const ShipLeft = 1320

Private Const Ship2Top = 720
Private Const Ship2Left = 1440


Private Const PadTop = 3720 '3720 '3840
Private Const PadLeft = 1080

Private Const OffPadDrop = 3740


Dim Direct1 As Integer

Dim Gap As Integer

Dim Level As Integer

Dim Lifes As Integer

Dim Score As Integer

Dim OnPad As Boolean

Dim HighScore As Long

Dim FuelWarned As Boolean

Dim UsingJoystick As Boolean


Dim Thrusting As Boolean
Dim Landed As Boolean
Dim Fuel As Integer
Dim DownSpeed As Long
Dim dd As Integer
Dim Cheat As String

Dim JoyX As Long
Dim JoyY As Long
Dim JoyButton As Integer

Dim JoyStillX As Long
Dim JoyStillY As Long


Dim Joystick As JOYINFO


Private Function Short_Name(Long_Path As String) As String

    Dim Short_Path As String
    Dim PathLength As Long
    Short_Path = Space(250)
    PathLength = GetShortPathName(Long_Path, Short_Path, Len(Short_Path))


    If PathLength Then
        Short_Name = Left$(Short_Path, PathLength)
        
    End If
End Function




Private Sub Form_Load()
Fuel = Fuel1
Gap = Gap1
Landed = True
lblFuel = "F U E L" & vbCrLf & Fuel
dd = 1
Level = 1
Lifes = 3

lblHighScore = "High: " & GetSetting("Moonlander", "Options", "HighScore", 0)
HighScore = GetSetting("Moonlander", "Options", "HighScore", 0)

Dim TempInt3 As String

TempInt3 = GetSetting("Moonlander", "Options", "GameSpeed", 2)

If TempInt3 = 3 Then
Timer1.Interval = 10
Label5.ForeColor = vbBlack
Label4.ForeColor = vbBlack
Label6.ForeColor = &HC0C0C0
End If

If TempInt3 = 2 Then
Timer1.Interval = 25
Label5.ForeColor = &HC0C0C0
Label4.ForeColor = vbBlack
Label6.ForeColor = vbBlack
End If

If TempInt3 = 1 Then
Timer1.Interval = 50
Label5.ForeColor = vbBlack
Label4.ForeColor = &HC0C0C0
Label6.ForeColor = vbBlack
End If





Randomize Timer
Dim TempInt As Integer



Do Until X = 100
TempInt = Int(Rnd * 5) + 1
If TempInt <= 4 Then Picture1.DrawWidth = 1 Else Picture1.DrawWidth = 2
Picture1.PSet (Int(Rnd * Picture1.Width) + 1, Int(Rnd * Picture1.Height) + 1), vbWhite
X = X + 1
Loop

If JoyStickMod.IsJoyPresent = 1 Then
Form1.Hide
Form2.Show
End If




End Sub


Private Sub JoystickTimer_Timer()
If UsingJoystick = True Then

If JoyStickMod.IsJoyPresent = 0 Then MsgBox "Communication with game controller has been lost" & vbCrLf & "Press OK to proceed using the keyboard", vbExclamation, "Error": UsingJoystick = False: JoystickTimer.Enabled = False: Exit Sub

JoyStickMod.GetJoystick 0, Joystick


If Landed = False Then


If Joystick.Buttons = 1 Then
If JoyButton = 0 Or JoyButton = 2 Or JoyButton = 3 Then
Debug.Print "Thrust"
DoThrust

End If
End If


If Joystick.Buttons = 0 Then
If Thrusting = True Then
Debug.Print "NO THRUST"
StopThrust
End If
End If

'If Joystick.Y > JoyStillY + 5000 Then
'DoThrust
'Else
'StopThrust
'End If

If Joystick.X > JoyStillX + 5000 Then Direct1 = 2: Debug.Print "RIGHT"

If Joystick.X < JoyStillX - 5000 Then Direct1 = 1: Debug.Print "LEFT"

If Joystick.X > JoyStillX - 5000 And Joystick.X < JoyStillX + 5000 Then Direct1 = 0









End If
End If

If Joystick.Buttons = 2 Then
If JoyButton = 0 Or JoyButton = 1 Or JoyButton = 3 Then
Debug.Print "New Game"
Label1_Click
End If
End If

JoyButton = Joystick.Buttons

End Sub

Private Sub Label1_Click()

WarningBeep.Enabled = False
sndPlaySound Short_Name(App.Path & "\Beep.wav"), SND_SYNC

Randomize Timer
Dim TempInt As Integer
lblStatus.Visible = False
Picture2.Visible = False
17 Picture2.Left = -Int(Rnd * Picture2.Width - Picture1.Width) + 1
If Picture2.Left >= 0 Then GoTo 17
Picture1.Cls
Picture2.Visible = True

FuelWarned = False

ship2.Visible = False

Do Until X = 100
TempInt = Int(Rnd * 5) + 1
If TempInt <= 4 Then Picture1.DrawWidth = 1 Else Picture1.DrawWidth = 2
Picture1.PSet (Int(Rnd * Picture1.Width) + 1, Int(Rnd * Picture1.Height) + 1), vbWhite
X = X + 1
Loop
lblHighScore.BackColor = &HFFC0C0
Ship.Picture = TheLander.Picture

Ship.Left = ShipLeft
Ship.Top = ShipTop
ship2.Top = Ship2Top
ship2.Left = Ship2Left

Shape1.Left = PadLeft
Shape1.Top = PadTop

Score = 0
lblScore = "Score: " & Score

Ship.Top = 240
Thrusting = False
Landed = False
Fuel = Fuel1
Gap = Gap1
Level = 1
lblLevel = "Level 1"
Lifes = 3
life1.Visible = True
life2.Visible = True
life3.Visible = True

DownSpeed = 0
lblLand = "L A N D"
lblFuel.BackColor = &H80FF80
lblFuel = "F U E L" & vbCrLf & Fuel
Timer1.Enabled = True


End Sub

Private Sub NewLevel(NewLevel As Integer)
Dim TempInt As Integer
WarningBeep.Enabled = False

Picture2.Visible = False
17 Picture2.Left = -Int(Rnd * Picture2.Width - Picture1.Width) + 1
If Picture2.Left >= 0 Then GoTo 17
Picture2.Visible = True

Picture1.Cls

FuelWarned = False

Do Until X = 100
TempInt = Int(Rnd * 5) + 1
If TempInt <= 4 Then Picture1.DrawWidth = 1 Else Picture1.DrawWidth = 2
Picture1.PSet (Int(Rnd * Picture1.Width) + 1, Int(Rnd * Picture1.Height) + 1), vbWhite
X = X + 1
Loop

lblScore = "Score: " & Score

Ship.Top = 240
Thrusting = False
Landed = False

If NewLevel = 1 Then
Fuel = Fuel1: Gap = Gap1: lblLevel = "Level 1": Level = 1
Ship.Left = ShipLeft
Ship.Top = ShipTop
ship2.Top = Ship2Top
ship2.Left = Ship2Left

Shape1.Left = PadLeft
Shape1.Top = PadTop
End If

If NewLevel = 2 Then Fuel = Fuel2: Gap = Gap2: lblLevel = "Level 2": Level = 2
If NewLevel = 3 Then Fuel = Fuel3: Gap = Gap3: lblLevel = "Level 3": Level = 3
If NewLevel = 4 Then Fuel = Fuel4: Gap = Gap4: lblLevel = "Level 4": Level = 4
If NewLevel = 5 Then Fuel = Fuel5: Gap = Gap5: lblLevel = "Level 5": Level = 5
If NewLevel = 6 Then Fuel = Fuel6: Gap = Gap6: lblLevel = "Level 6": Level = 6
If NewLevel = 7 Then Fuel = Fuel7: Gap = Gap7: lblLevel = "Level 7": Level = 7
If NewLevel = 8 Then Fuel = Fuel8: Gap = Gap8: lblLevel = "Level 8": Level = 8
If NewLevel = 9 Then Fuel = fuel9: Gap = gap9: lblLevel = "Level 9": Level = 9
If NewLevel = 10 Then Fuel = fuel10: Gap = gap10: lblLevel = "Level 10": Level = 10


DownSpeed = 0
Ship.Picture = TheLander.Picture

lblLand = "L A N D"
lblFuel.BackColor = &H80FF80
lblFuel = "F U E L" & vbCrLf & Fuel

End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub Label4_Click()
Timer1.Interval = 50
Label5.ForeColor = vbBlack
Label4.ForeColor = &HC0C0C0
Label6.ForeColor = vbBlack
SaveSetting "Moonlander", "Options", "GameSpeed", "1"
End Sub

Private Sub Label5_Click()
Timer1.Interval = 25
Label5.ForeColor = &HC0C0C0
Label4.ForeColor = vbBlack
Label6.ForeColor = vbBlack
SaveSetting "Moonlander", "Options", "GameSpeed", "2"
End Sub

Private Sub Label6_Click()
Timer1.Interval = 10
Label5.ForeColor = vbBlack
Label4.ForeColor = vbBlack
Label6.ForeColor = &HC0C0C0
SaveSetting "Moonlander", "Options", "GameSpeed", "3"
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If UsingJoystick = True Then Exit Sub
If Landed = False Then

If KeyCode = 37 Then Direct1 = 1
If KeyCode = 39 Then Direct1 = 2

If Thrusting = False Then
If Fuel <= 0 Then Exit Sub
If KeyCode = 40 Then lblThrust.BackColor = vbRed: lblThrust.Font.Bold = True: ship2.Visible = True: Thrusting = True: sndPlaySound Short_Name(App.Path & "\" & "Thrust.wav"), SND_LOOP

End If
End If
End Sub
Private Sub DoThrust()
If Thrusting = False Then
If Fuel <= 0 Then Exit Sub
lblThrust.BackColor = vbRed: lblThrust.Font.Bold = True: ship2.Visible = True: Thrusting = True: sndPlaySound Short_Name(App.Path & "\" & "Thrust.wav"), SND_LOOP
End If
End Sub

Private Sub StopThrust()
lblThrust.BackColor = &H80FFFF: lblThrust.Font.Bold = False: Thrusting = False: ship2.Visible = False: sndPlaySound Short_Name(App.Path & "\" & "Thrust.wav"), SND_STOP

End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
If UsingJoystick = True Then Exit Sub

If KeyAscii = 110 Then Label1_Click
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
If UsingJoystick = True Then Exit Sub
If KeyCode = 40 Then lblThrust.BackColor = &H80FFFF: lblThrust.Font.Bold = False: Thrusting = False: ship2.Visible = False: sndPlaySound Short_Name(App.Path & "\" & "Thrust.wav"), SND_STOP
If KeyCode = 37 Then Direct1 = 0
If KeyCode = 39 Then Direct1 = 0
End Sub

Private Sub Timer1_Timer()
If Landed = True Then DownSpeed = 0: Exit Sub

If Fuel <= 0 And FuelWarned = False Then GoTo 19
If Thrusting = True Then
If Fuel <= 0 Then
19
sndPlaySound Short_Name(App.Path & "\" & "Thrust.wav"), SND_STOP

Thrusting = False
lblThrust.BackColor = &H80FFFF
WarningBeep.Enabled = True
lblFuel.BackColor = vbRed
ship2.Visible = False
sndPlaySound Short_Name(App.Path & "\Beep2.wav"), SND_SYNC
FuelWarned = True
GoTo 10
End If

DownSpeed = DownSpeed - 1
Fuel = Fuel - 1
lblFuel = "F U E L" & vbCrLf & Fuel
If Fuel <= 10 Then lblFuel.BackColor = &H80FF&
If DownSpeed <= 20 Then DownSpeed = DownSpeed - 2
Else
DownSpeed = DownSpeed + 1
If DownSpeed >= 20 Then DownSpeed = DownSpeed + 3
End If
10

Ship.Top = Ship.Top + DownSpeed
ship2.Top = Ship.Top + Ship.Height

If Direct1 = 1 Then Ship.Left = Ship.Left - 6: ship2.Left = ship2.Left - 6
If Direct1 = 2 Then Ship.Left = Ship.Left + 6: ship2.Left = ship2.Left + 6

If Not Level = 1 Then
If dd = 1 Then
Shape1.Left = Shape1.Left - 15
If Level = 6 Then Shape1.Left = Shape1.Left - 8
If Level = 7 Then Shape1.Left = Shape1.Left - 16
If Level = 8 Then Shape1.Left = Shape1.Left - 20
If Level = 9 Then Shape1.Left = Shape1.Left - 25
If Level = 10 Then Shape1.Left = Shape1.Left - 30

End If

If dd = 2 Then
Shape1.Left = Shape1.Left + 15
If Level = 6 Then Shape1.Left = Shape1.Left + 8
If Level = 7 Then Shape1.Left = Shape1.Left + 16
If Level = 8 Then Shape1.Left = Shape1.Left + 20
If Level = 9 Then Shape1.Left = Shape1.Left + 25
If Level = 10 Then Shape1.Left = Shape1.Left + 30

End If

End If


If Shape1.Left <= 0 Then dd = 2
If Shape1.Left >= Picture1.Width - Shape1.Width Then dd = 1
'Debug.Print Ship.Left + Ship.Width & " - "; Shape1.Width + Shape1.Left

lblAlt = "Alt: " & (Ship.Height - Ship.Top) + (Shape1.Height + Shape1.Top) - 1120
lblSpeed = "Speed: " & DownSpeed


If Ship.Left + Ship.Width > Shape1.Left + (Ship.Width - 100) And Ship.Left + (Ship.Width - 100) < Shape1.Left + Shape1.Width Then OnPad = True Else OnPad = False

'If DownSpeed >= 80 And Ship.Top <= 5000 Then If Not WarningBeep.Enabled = True Then WarningBeep.Enabled = True


If DownSpeed < Gap Then
If OnPad = True Then lblLand.BackColor = &H80FF80 Else lblLand.BackColor = vbRed
Else
lblLand.BackColor = vbRed
End If

If OnPad = True Then
If Not Shape1.BackColor = &HC0C0C0 Then Shape1.BackColor = &HC0C0C0
Else
If Not Shape1.BackColor = &H808080 Then Shape1.BackColor = &H808080
End If


If Ship.Top >= Shape1.Top - Ship.Height Then
If OnPad = False Then GoTo 25
If DownSpeed < Gap Then
lblLand = "L A N D E D"
Landed = True
lblAlt = "ALT: 0"
If Level = 10 Then
lblStatus = "C O M P L E T E D !"
lblStatus.Visible = True

If Level = 10 Then Score = Score + 10000
Score = Score + Fuel - DownSpeed
lblScore = "Score: " & Score

If Score > HighScore Then
SaveSetting "Moonlander", "Options", "HighScore", Score
HighScore = Score
lblHighScore = "High: " & Score
lblHighScore.BackColor = vbYellow
End If
End If


ship2.Visible = False
lblThrust.BackColor = &H80FFFF: lblThrust.Font.Bold = False
If Level = 1 Then Score = Score + 10
If Level = 2 Then Score = Score + 100
If Level = 3 Then Score = Score + 200
If Level = 4 Then Score = Score + 400
If Level = 5 Then Score = Score + 600
If Level = 6 Then Score = Score + 1000
If Level = 7 Then Score = Score + 2000
If Level = 8 Then Score = Score + 2500
If Level = 9 Then Score = Score + 5000


If Not Level = 10 Then
Score = Score + Fuel - DownSpeed
lblScore = "Score: " & Score
End If




ship2.Visible = False
'sndPlaySound Short_Name(App.Path & "\" & "Thrust.wav"), SND_STOP
Shape1.BackColor = vbGreen
Sleep 10
sndPlaySound Short_Name(App.Path & "\Applause.wav"), SND_ASYNC
If Not Level = 10 Then NewLevel (Level + 1)

Else
GoTo 26
25 lblLand = "O F F   P A D": Ship.Top = OffPadDrop - Ship.Height: GoTo 27
26
lblLand = "T O O  F A S T"
27
Landed = True
lblAlt = "ALT: 0"
ship2.Visible = False
Ship.Picture = RocketCrash.Picture
WarningBeep.Enabled = False

lblThrust.BackColor = &H80FFFF: lblThrust.Font.Bold = False
ship2.Visible = False
'sndPlaySound Short_Name(App.Path & "\" & "Thrust.wav"), SND_STOP

'ON GAME OVER
If Lifes = 1 Then
lblStatus = "G A M E   O V E R"
lblStatus.Visible = True
life1.Visible = False
If Score > HighScore Then
SaveSetting "Moonlander", "Options", "HighScore", Score
HighScore = Score
lblHighScore = "High: " & Score
lblHighScore.BackColor = vbYellow
End If
End If


Sleep 10
sndPlaySound Short_Name(App.Path & "\Explode.wav"), SND_ASYNC
If Lifes = 1 Then Exit Sub
Lifes = Lifes - 1
If Lifes = 2 Then life3.Visible = False: NewLevel (Level)
If Lifes = 1 Then life2.Visible = False: NewLevel (Level)
End If
End If







End Sub

Private Sub Timer2_Timer()
If Thrusting = True Then
Dim dd5 As Integer
dd5 = Int(Rnd * 3) + 1
If dd5 = 1 Then ship2.Picture = T1
If dd5 = 2 Then ship2.Picture = T2
If dd5 = 3 Then ship2.Picture = T3
End If

End Sub

Private Sub WarningBeep_Timer()
sndPlaySound Short_Name(App.Path & "\Beep2.wav"), SND_SYNC
If Fuel <= 0 Then
If lblFuel.BackColor = vbRed Then lblFuel.BackColor = &H80FF80 Else lblFuel.BackColor = vbRed
End If

End Sub




Private Function WinTime() As String
Dim lngReturn As Long
Dim tmp, tmpTime, tmpHours, tmpMinutes, tmpSeconds, tmpMilliSeconds As String
lngReturn = GetTickCount()
WinTime = lngReturn
End Function

Public Function Sleep(Time As Integer, Optional Freeze As Boolean) As Boolean
Dim StartSleepTime As Long
Dim StopSleepTime As Long
StopSleepTime = WinTime + Time

Do
If Freeze = False Then DoEvents
If WinTime >= StopSleepTime Then GoTo 10
Loop

10 Sleep = True
End Function


Public Sub ControllerType(CType As Integer)
If CType = 1 Then
UsingJoystick = False
Else
JoyStickMod.GetJoystick 0, Joystick
UsingJoystick = True
JoyStillX = Joystick.X
JoyStillY = Joystick.Y
End If

End Sub

