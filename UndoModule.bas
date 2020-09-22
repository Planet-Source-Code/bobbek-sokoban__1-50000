Attribute VB_Name = "UndoModule"
Option Explicit
Global UndoBoxFrom(1 To 5000)
Global UndoBoxTo(1 To 5000)
Global UndoSokoFrom(1 To 5000)
Global UndoSokoTo(1 To 5000)
Global Totalmoves As Integer
Function UndoMove()

If Totalmoves > 0 Then
'Debug.Print UndoSokoFrom(Totalmoves) - UndoSokoTo(Totalmoves)
If UndoSokoFrom(Totalmoves) - UndoSokoTo(Totalmoves) = -1 Then
Form1.sokoban.Picture = Form1.imgSokoRIGHT.Picture
End If
If UndoSokoFrom(Totalmoves) - UndoSokoTo(Totalmoves) = 1 Then
Form1.sokoban.Picture = Form1.imgSokoLEFT.Picture
End If
If UndoSokoFrom(Totalmoves) - UndoSokoTo(Totalmoves) = -20 Then
Form1.sokoban.Picture = Form1.imgSokoDOWN.Picture
End If
If UndoSokoFrom(Totalmoves) - UndoSokoTo(Totalmoves) = 20 Then
Form1.sokoban.Picture = Form1.imgSokoUP.Picture
End If

If UndoSokoFrom(Totalmoves) <> 0 And UndoSokoTo(Totalmoves) <> 0 Then
    GetOriginalPic (UndoSokoTo(Totalmoves))
    Form1.brick(UndoSokoFrom(Totalmoves)).Picture = Form1.sokoban.Picture
    SokoPos = UndoSokoFrom(Totalmoves)
    UndoSokoFrom(Totalmoves) = 0
    UndoSokoTo(Totalmoves) = 0
End If
If UndoBoxFrom(Totalmoves) <> 0 And UndoBoxTo(Totalmoves) <> 0 Then
    GetOriginalPic (UndoBoxTo(Totalmoves))
    If Form1.brick(UndoBoxFrom(Totalmoves)).Picture = Form1.golv.Picture Then
        Form1.brick(UndoBoxFrom(Totalmoves)).Picture = Form1.Låda.Picture
    Else
        Form1.brick(UndoBoxFrom(Totalmoves)).Picture = Form1.LådaF.Picture
    End If
    UndoBoxFrom(Totalmoves) = 0
    UndoBoxTo(Totalmoves) = 0
End If
    Totalmoves = Totalmoves - 1
    Form1.lblMoves = Totalmoves
End If
End Function
