VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Controller"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4725
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Controller.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4200
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3120
      Picture         =   "Controller.frx":0442
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"Controller.frx":0D0C
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   1680
      Picture         =   "Controller.frx":0DA6
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Controller"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   240
      Picture         =   "Controller.frx":3578
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Keyboard"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Joystick As JOYINFO

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then StartKeyboard
End Sub

Private Sub Image2_Click()
StartKeyboard

End Sub

Private Sub Image3_Click()
StartJoystick
End Sub

Private Sub Label1_Click()
Unload Form1
Unload Form2
End

End Sub

Private Sub Label3_Click()
StartKeyboard
End Sub

Private Sub Label4_Click()
StartJoystick
End Sub

Private Sub Timer1_Timer()
JoyStickMod.GetJoystick 0, Joystick

If Joystick.Buttons = 1 Then StartJoystick

End Sub

Private Sub StartJoystick()
Form1.ControllerType (2)
Form1.Show
Unload Form2
End Sub

Private Sub StartKeyboard()
Form1.ControllerType (1)
Form1.Show
Unload Form2
End Sub

