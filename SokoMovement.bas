Attribute VB_Name = "Movement"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Global SokoPos
Dim DontMove
Global NrTargets As Integer
Global BoxTarget(1 To 50)

Public Function chkMove(direction As String)
If direction = "DOWN" And SokoPos < 381 Then
    moveSoko (20)
End If
If direction = "UP" And SokoPos > 20 Then
    moveSoko (-20)
End If

If direction = "RIGHT" And SokoPos Mod 20 <> 0 Then
    moveSoko (1)
End If
If direction = "LEFT" And SokoPos Mod 20 <> 1 Then
    moveSoko (-1)
End If

End Function

Public Function moveSoko(Steps As Integer)
DontMove = False

If Form1.brick(SokoPos + Steps).Picture = Form1.vägg.Picture Then
    DontMove = True
End If
If Not DontMove Then Totalmoves = Totalmoves + 1

If Form1.brick(SokoPos + Steps).Picture = Form1.Låda.Picture Or _
   Form1.brick(SokoPos + Steps).Picture = Form1.LådaF.Picture Then
    If Form1.brick(SokoPos + (2 * Steps)).Picture = Form1.golv.Picture Then
        Form1.brick(SokoPos + (2 * Steps)).Picture = Form1.Låda.Picture
        GetOriginalPic (SokoPos + Steps)
        UndoBoxFrom(Totalmoves) = SokoPos + Steps
        UndoBoxTo(Totalmoves) = SokoPos + (2 * Steps)
    ElseIf Form1.brick(SokoPos + (2 * Steps)).Picture = Form1.target.Picture Then
        Form1.brick(SokoPos + (2 * Steps)).Picture = Form1.LådaF.Picture
        GetOriginalPic (SokoPos + Steps)
        UndoBoxFrom(Totalmoves) = SokoPos + Steps
        UndoBoxTo(Totalmoves) = SokoPos + (2 * Steps)
    Else
    DontMove = True
    Totalmoves = Totalmoves - 1
    End If
End If

If Not DontMove Then        'Flytta SokoGubben
    Form1.brick(SokoPos + Steps).Picture = Form1.sokoban.Picture
    GetOriginalPic (SokoPos)
    UndoSokoFrom(Totalmoves) = SokoPos
    UndoSokoTo(Totalmoves) = SokoPos + Steps
    SokoPos = SokoPos + Steps
    Form1.lblMoves = Totalmoves
    Call checkwin
End If
If Totalmoves > 4995 Then
    MsgBox "The UndoModule can not handle more than 5000 moves," & vbNewLine & _
    "PROGRAM WILL BE TERMINATED", vbCritical + vbOKOnly, "YOU TOTALLY SUCK! ;-) Please Try Again!"
    End
End If

End Function

Function checkwin()
Dim färdiga
färdiga = 0
For i = 1 To NrTargets
    If Form1.brick(BoxTarget(i)).Picture = Form1.LådaF.Picture Then
        färdiga = färdiga + 1
     
    End If
Next i
If färdiga = NrTargets Then
    Call checkScore
    Form1.cboLevelSelect.Enabled = True
End If
End Function











