VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Space Rockz"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   600
      Top             =   600
   End
   Begin VB.PictureBox Ship3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   7320
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   525
      ScaleWidth      =   510
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Timer TmrRock1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer TmrB 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   600
      Top             =   120
   End
   Begin VB.PictureBox Boom1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   720
      Picture         =   "FrmMain.frx":0E7A
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox Ship2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   120
      Picture         =   "FrmMain.frx":106C
      ScaleHeight     =   525
      ScaleWidth      =   510
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox Ship 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   3720
      Picture         =   "FrmMain.frx":1EE6
      ScaleHeight     =   525
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   5520
      Width           =   510
   End
   Begin VB.Shape b 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   960
      Top             =   5880
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape Rock1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      Height          =   855
      Index           =   0
      Left            =   6960
      Shape           =   1  'Square
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cur As Integer

Private Sub Form_Load()
CreateSpace
For i = 1 To 10
    Load b(i)
    Load TmrB(i)
Next i
For r1 = 1 To 15
    Load Rock1(r1)
    Load TmrRock1(r1)
Next r1
Me.Show
RockLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Ship_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
    Call Shoot
End If
End Sub

Private Sub Ship_KeyUp(KeyCode As Integer, Shift As Integer)
Ship.Picture = LoadPicture(App.Path & "\ship.bmp")
End Sub

Private Sub Timer1_Timer()
Dim CurCount As Integer ' redo stars every 20 counts of this timer
' key events
If GetKey(vbKeyLeft) = True Then
    Ship.Left = Ship.Left - ShipSpeed
End If
If GetKey(vbKeyRight) = True Then
    Ship.Left = Ship.Left + ShipSpeed
End If
If GetKey(vbKeyUp) = True Then
    Ship.Top = Ship.Top - ShipSpeed
    Ship.Picture = Ship2.Picture
End If
If GetKey(vbKeyDown) = True Then
     Ship.Top = Ship.Top + ShipSpeed
End If
wCollision

RedoSpace
Me.Caption = "Space Rockz - Score: " & Score
End Sub

Public Sub Shoot()
Cur = -1
For i = 0 To 10
    If b(i).Visible = False Then
        Cur = i
    End If
Next i
If Cur = -1 Then Exit Sub
b(Cur).Left = (Ship.Left + 205)
b(Cur).Top = (Ship.Top - b(Cur).Height) - 50
b(Cur).Visible = True
TmrB(Cur).Enabled = True
End Sub

Private Sub Timer2_Timer()
BulletCollision
End Sub

Private Sub Timer3_Timer()
BulletCollision
End Sub

Private Sub TmrB_Timer(Index As Integer)
b(Index).Top = b(Index).Top - bSpeed
If b(Index).Top < 0 Then
    TmrB(Index).Enabled = False
    b(Index).Visible = False
    Boom1.Top = 0
    Boom1.Left = (b(Index).Left - 50)
    Boom1.Visible = True
    Pause 0.1
    Boom1.Visible = False
End If
End Sub

Private Sub TmrRock1_Timer(Index As Integer)
If ArrayRock(Index).rX < -0 Then
    Rock1(Index).Left = Rock1(Index).Left - ArrayRock(Index).rX
Else
    Rock1(Index).Left = Rock1(Index).Left + ArrayRock(Index).rX
End If
Rock1(Index).Top = Rock1(Index).Top + ArrayRock(Index).rY
End Sub
