Attribute VB_Name = "ModSpace"
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Type Star
    sX As Long
    sY As Long
End Type
Public StarArrays(1000) As Star
Public Sub CreateSpace()
Dim rX As Integer, rY As Integer
FrmMain.Show
Randomize
For i = 0 To 1000
    rX = Int(Rnd * FrmMain.ScaleWidth)
    rY = Int(Rnd * FrmMain.ScaleHeight)
    FrmMain.CurrentX = rX - 1
    FrmMain.CurrentY = rY - 1
    FrmMain.Line -(rX, rY), RGB(255, 255, 255)
    StarArrays(i).sX = rX
    StarArrays(i).sY = rY
Next i
End Sub
Public Sub RedoSpace()
Dim curStarX As Long, curStarY As Long
For i = 0 To 1000
    Let curStarX = StarArrays(i).sX
    Let curStarY = StarArrays(i).sY
    FrmMain.CurrentX = curStarX - 1
    FrmMain.CurrentY = curStarY - 1
    FrmMain.Line -(curStarX, curStarY), RGB(255, 255, 255)
Next i
End Sub
