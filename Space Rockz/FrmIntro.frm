VERSION 5.00
Begin VB.Form FrmIntro 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Space Rockz - Main Menu"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "FrmIntro"
   MaxButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start Game!"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label blah 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "FrmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FrmMain.Show
Me.Hide
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Dim Msg As String
Show
Pause 1#
Msg$ = "As you know, for the last few weeks, Earth has been struck by hundreds of Asteroids.  Scientist Chris Hodges has descovered that this surely isn't the end of the destruction."
Msg$ = Msg$ & "  Knowing that our planet can't hold out very much longer, we have suited you up for ''Mission SaveOurAsses.''  We have put together the best machinery in todays knowledge--don't let us down..."
Call TypeWriterThing(Msg$)
End Sub
Public Sub TypeWriterThing(sText As String)
On Error Resume Next
For i = 0 To Len(sText$)
    blah.Caption = Left(sText$, i)
    blah.Refresh
    Pause 0.1
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
Do: DoEvents
Loop
End Sub
