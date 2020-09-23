Attribute VB_Name = "ModMain"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const ShipSpeed = 25
Public Const bSpeed = 75
Public Function GetKey(lngKey As Long) As Boolean
If GetAsyncKeyState(lngKey&) < 0 Then
    GetKey = True
Else
    GetKey = False
End If
End Function
Public Sub Pause(Length As Double)
Dim StartTime
Let StartTime = Timer
Do While Timer - StartTime < Length
    DoEvents
Loop
End Sub
Public Sub wCollision()
If FrmMain.Ship.Left < 0 Then
    FrmMain.Ship.Left = 0
ElseIf FrmMain.Ship.Left > (FrmMain.ScaleWidth - FrmMain.Ship.ScaleWidth) Then
    FrmMain.Ship.Left = (FrmMain.ScaleWidth - FrmMain.Ship.ScaleWidth)
ElseIf FrmMain.Ship.Top < 0 Then
    FrmMain.Ship.Top = 0
ElseIf FrmMain.Ship.Top > (FrmMain.ScaleHeight - FrmMain.Ship.ScaleHeight) Then
    FrmMain.Ship.Top = (FrmMain.ScaleHeight - FrmMain.Ship.ScaleHeight)
End If
End Sub
Public Sub ResetGame()
For i = 0 To 15
    FrmMain.Rock1(i).Visible = False
    FrmMain.TmrRock1(i).Enabled = False
    FrmMain.Rock1(i).Left = 0
    FrmMain.Rock1(i).Top = 0
Next i
FrmMain.Ship.Left = 3720
FrmMain.Ship.Top = 5520
End Sub
