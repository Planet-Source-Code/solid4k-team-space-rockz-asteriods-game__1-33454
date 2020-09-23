Attribute VB_Name = "ModRocks"
Dim Cur As Integer
Public Score As Integer
Public Type RockM
    rX As Integer
    rY As Integer
End Type
Public ArrayRock(50) As RockM
Public Sub CreateRock()
Cur = -1
For i = 0 To 15
    If FrmMain.Rock1(i).Visible = False Then
        Let Cur = i
    End If
Next i
If Cur = -1 Then Exit Sub
Randomize

FrmMain.Rock1(Cur).Top = -5000

FrmMain.Rock1(Cur).Left = Int(Rnd * FrmMain.ScaleWidth)
FrmMain.Rock1(Cur).Visible = True
FrmMain.TmrRock1(Cur).Enabled = True
ArrayRock(Cur).rX = Int(Rnd * 30) - 15
ArrayRock(Cur).rY = Int(Rnd * 20) + 5
End Sub
Public Sub RockLoop()
Do: DoEvents
    Pause Int(Rnd * 3)
    CreateRock
    wRockCollision
    sRockCollision
Loop
End Sub
Public Sub wRockCollision()
For i = 0 To 15
    If FrmMain.Rock1(i).Top > (FrmMain.ScaleHeight + FrmMain.Rock1(i).Height) Then
        FrmMain.Rock1(i).Visible = False
        FrmMain.TmrRock1(i).Enabled = False
    End If
    If FrmMain.Rock1(i).Left > 8055 Then
        FrmMain.Rock1(i).Left = 0
    End If
    If FrmMain.Rock1(i).Left < 0 Then
        FrmMain.Rock1(i).Left = FrmMain.ScaleWidth + FrmMain.Rock1(i).Width
    End If
Next i
End Sub
Public Sub sRockCollision()
For i = 0 To 15
    If FrmMain.Ship.Left > (FrmMain.Rock1(i).Left - 100) And FrmMain.Ship.Left < (FrmMain.Rock1(i).Left + FrmMain.Rock1(i).Width) + 100 Then
        If FrmMain.Ship.Top > (FrmMain.Rock1(i).Top - 100) And FrmMain.Ship.Top < (FrmMain.Rock1(i).Top + FrmMain.Rock1(i).Height) + 100 Then
            For boom = 0 To 5
                FrmMain.Ship.Picture = FrmMain.Ship3.Picture
                Pause 0.2
                FrmMain.Ship.Picture = LoadPicture(App.Path & "\ship.bmp")
                Pause 0.2
            Next boom
            Call ResetGame
            Score = Score - 50
        End If
    End If
Next i
End Sub
Public Sub BulletCollision()
For i = 0 To 15
    If FrmMain.Rock1(i).Visible = False Then GoTo jump1:
    DoEvents
    For b = 0 To 10
        If FrmMain.b(b).Visible = False Then GoTo jump2:
        DoEvents
        If FrmMain.b(b).Left > FrmMain.Rock1(i).Left And FrmMain.b(b).Left < (FrmMain.Rock1(i).Left + FrmMain.Rock1(i).Width) Then
            If FrmMain.b(b).Top > FrmMain.Rock1(i).Top And FrmMain.b(b).Top < (FrmMain.Rock1(i).Top + FrmMain.Rock1(i).Height) Then
                Call PostExplosion(CInt(b))
                FrmMain.Rock1(i).Visible = False
                FrmMain.Rock1(i).Left = 0
                FrmMain.Rock1(i).Top = 0
                FrmMain.TmrRock1(i).Enabled = False
                FrmMain.b(b).Visible = False
                FrmMain.TmrB(b).Enabled = False
                FrmMain.b(b).Left = 0
                FrmMain.b(b).Top = FrmMain.ScaleHeight
                Score = Score + 10
            End If
        End If
jump2:
    Next b
jump1:
Next i
End Sub
Public Sub PostExplosion(bIndex As Long)
FrmMain.Boom1.Left = (FrmMain.b(bIndex).Left - 50)
FrmMain.Boom1.Top = FrmMain.b(bIndex).Top
FrmMain.Boom1.Visible = True
Pause 0.1
FrmMain.Boom1.Visible = False
End Sub
