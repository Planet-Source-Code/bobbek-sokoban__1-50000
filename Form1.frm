VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Box Kicking Soko, The Angry Sokoban"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAutoPlay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8400
      Top             =   6120
   End
   Begin VB.CommandButton cmdAutoPlay 
      Caption         =   "Auto Play!"
      Enabled         =   0   'False
      Height          =   435
      Left            =   7800
      TabIndex        =   8
      Top             =   5700
      Width           =   1035
   End
   Begin VB.CommandButton cmdHiscore 
      Caption         =   "Display Hiscore"
      Height          =   495
      Left            =   7800
      TabIndex        =   7
      Top             =   1620
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "Level"
      Height          =   615
      Left            =   7800
      TabIndex        =   5
      Top             =   180
      Width           =   1035
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moves"
      Height          =   615
      Left            =   7800
      TabIndex        =   3
      Top             =   900
      Width           =   1035
      Begin VB.Label lblMoves 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo Move"
      Enabled         =   0   'False
      Height          =   435
      Left            =   7800
      TabIndex        =   2
      Top             =   4740
      Width           =   1035
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   4320
      Width           =   1035
   End
   Begin VB.Timer tmrGetKeyPress 
      Interval        =   100
      Left            =   8040
      Top             =   2700
   End
   Begin VB.ComboBox cboLevelSelect 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7800
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Intro"
      Top             =   3900
      Width           =   1035
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7800
      Top             =   1620
      Width           =   1035
      Visible         =   0   'False
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7800
      Top             =   4740
      Width           =   1035
      Visible         =   0   'False
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7800
      Top             =   4320
      Width           =   1035
      Visible         =   0   'False
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7800
      Top             =   3900
      Width           =   1035
      Visible         =   0   'False
   End
   Begin VB.Image imgSokoLEFT 
      Height          =   375
      Left            =   9300
      Picture         =   "Form1.frx":08CA
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgSokoRIGHT 
      Height          =   375
      Left            =   9300
      Picture         =   "Form1.frx":1078
      Top             =   4380
      Width           =   375
   End
   Begin VB.Image imgSokoDOWN 
      Height          =   375
      Left            =   9300
      Picture         =   "Form1.frx":1826
      Top             =   3540
      Width           =   375
   End
   Begin VB.Image imgSokoUP 
      Height          =   375
      Left            =   9300
      Picture         =   "Form1.frx":1FD4
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image brick 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   1
      Left            =   240
      Top             =   240
      Width           =   375
      Visible         =   0   'False
   End
   Begin VB.Image tom 
      Height          =   375
      Left            =   9180
      Picture         =   "Form1.frx":2782
      Top             =   180
      Width           =   375
   End
   Begin VB.Image LådaF 
      Height          =   375
      Left            =   9180
      Picture         =   "Form1.frx":2F30
      Top             =   600
      Width           =   375
   End
   Begin VB.Image Låda 
      Height          =   375
      Left            =   9180
      Picture         =   "Form1.frx":36DE
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image target 
      Height          =   375
      Left            =   9180
      Picture         =   "Form1.frx":3E8C
      Top             =   1500
      Width           =   375
   End
   Begin VB.Image golv 
      Height          =   375
      Left            =   9180
      Picture         =   "Form1.frx":463A
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image vägg 
      Height          =   375
      Left            =   9180
      Picture         =   "Form1.frx":4DE8
      Top             =   2760
      Width           =   375
   End
   Begin VB.Image sokoban 
      Height          =   375
      Left            =   9180
      Picture         =   "Form1.frx":5596
      Top             =   2340
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lvl As Integer
Dim APMoves
Dim clkTimer

Private Sub cmdHiscore_Click()
Dim i
Dim Level As String
Dim MaxScore
Dim stringen
For i = 1 To 20
    Level = "Level" & Trim$(Str$(i))
    MaxScore = Val(GetINIKeyStr("HISCORE", Level))
    stringen = stringen & "Level " & Trim$(Str$(i)) & vbTab & vbTab & "Score: " & MaxScore & vbTab & vbNewLine
Next i
MsgBox stringen, vbOKOnly, "  *  *  *  Hiscore List  *  *  *  "
End Sub


Private Sub cboLevelSelect_Click()
lvl = Val(cboLevelSelect.ListIndex + 1)
cmdAutoPlay.Enabled = False
End Sub



Private Sub cmdStart_Click()
Dim i
Dim j
Dim a$
Dim tile
NrTargets = 0
For i = 1 To 50
    BoxTarget(i) = 0
Next i
lblLevel.Caption = lvl
If lvl > 0 Then
    GetLevel (lvl)
    For i = 1 To 20
        a$ = LevelData(i)
        For j = 1 To 20
            tile = Mid$(a$, j, 1)
            If tile = "X" Then brick(((i - 1) * 20) + j).Picture = tom.Picture
            If tile = "L" Then brick(((i - 1) * 20) + j).Picture = Låda.Picture
            If tile = "F" Or tile = "f" Then
                brick(((i - 1) * 20) + j).Picture = target.Picture
                NrTargets = NrTargets + 1
                BoxTarget(NrTargets) = ((i - 1) * 20) + j
                If tile = "f" Then brick(((i - 1) * 20) + j).Picture = LådaF.Picture ' Target with box ontop
            End If
            If tile = " " Then brick(((i - 1) * 20) + j).Picture = golv.Picture
            If tile = "V" Then brick(((i - 1) * 20) + j).Picture = vägg.Picture
        Next j
    Next i
    brick(Val(LevelData(21))).Picture = sokoban.Picture
    SokoPos = Val(LevelData(21))
End If
'For i = 1 To NrTargets
'    Debug.Print i, BoxTarget(i)
'Next i
Totalmoves = 0
lblMoves.Caption = "0"
For i = 1 To 5000
    UndoBoxFrom(i) = 0
    UndoBoxTo(i) = 0
    UndoSokoFrom(i) = 0
    UndoSokoTo(i) = 0
Next i
If lvl > 0 Then
    cmdAutoPlay.Enabled = True
    tmrAutoPlay.Enabled = False
    If FreeFile > 1 Then Close #1 'Still Open, Start Pressed While Autoplay
End If
End Sub

Private Sub cmdUndo_Click()
UndoMove
End Sub

Private Sub cmdAutoPlay_Click()
Dim i
Dim poop As String
If lvl > 0 Then
    Open App.Path & "\Bana" & Trim(Str$(lvl)) & ".txt" For Input As #1
    tmrAutoPlay.Enabled = True
    cmdAutoPlay.Enabled = False
End If
End Sub


Private Sub Form_Load()
'------- Code From Roger Gilchrist ------------------------
Dim i
Dim j
Dim k
On Error Resume Next             'Just a safety thing when using Load
  For i = 0 To 19
    cboLevelSelect.AddItem "Level " & (i + 1)
    For j = 0 To 19
      k = k + 1                  'Create an index for the new Image
      Load brick(k)
      With brick(k)
        .Left = 120 + .Width * j 'set brick into the grid
        .Top = 120 + .Width * i
        .Visible = True          'make it visible
      End With
    Next j
  Next i
'------- END Roger's Code --------------------------------
'Link to Roger's uploads at PSCode
'http://www.pscode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=Roger+Gilchrist&lngWId=1&B1=Quick+Search
'
INIFile = App.Path & "\SAVES.INI"
Call intro
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Visible = True
    Image3.Visible = True
    Image4.Visible = True
    Image5.Visible = True
    cboLevelSelect.Enabled = False
    cmdStart.Enabled = False
    cmdUndo.Enabled = False
    cmdHiscore.Enabled = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = False
cboLevelSelect.Enabled = True
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
cmdStart.Enabled = True
End Sub
Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
cmdUndo.Enabled = True
End Sub
Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
cmdHiscore.Enabled = True
End Sub

Private Sub tmrGetKeyPress_Timer()
If Val(lblMoves.Caption) > 0 Then cmdAutoPlay.Enabled = False

If Me.hWnd = GetActiveWindow() And lvl > 0 And tmrAutoPlay.Enabled = False Then
    If GetAsyncKeyState(vbKeyDown) < 0 Then
        sokoban.Picture = imgSokoDOWN.Picture
        chkMove ("DOWN")
        Exit Sub
    End If
    If GetAsyncKeyState(vbKeyUp) < 0 Then
        sokoban.Picture = imgSokoUP.Picture
        chkMove ("UP")
        Exit Sub
    End If
    If GetAsyncKeyState(vbKeyRight) < 0 Then
        sokoban.Picture = imgSokoRIGHT.Picture
        chkMove ("RIGHT")
        Exit Sub
     End If
    If GetAsyncKeyState(vbKeyLeft) < 0 Then
        sokoban.Picture = imgSokoLEFT.Picture
        chkMove ("LEFT")
        Exit Sub
    End If
End If
End Sub
Private Sub MovaSoko(toPlace As Integer)
Dim apan
apan = SokoPos - toPlace

If lvl > 0 Then
    If apan = -20 Then
        sokoban.Picture = imgSokoDOWN.Picture
        chkMove ("DOWN")
    End If
    If apan = 20 Then
        sokoban.Picture = imgSokoUP.Picture
        chkMove ("UP")
    End If
    If apan = -1 Then
        sokoban.Picture = imgSokoRIGHT.Picture
        chkMove ("RIGHT")
     End If
    If apan = 1 Then
        sokoban.Picture = imgSokoLEFT.Picture
        chkMove ("LEFT")
    End If
End If
End Sub

Private Sub tmrAutoPlay_Timer()
APMoves = APMoves + 1
If Not EOF(1) Then
    Input #1, APMoves
    MovaSoko (Val(Trim(APMoves)))
End If
If EOF(1) Then
    Close #1
    tmrAutoPlay.Enabled = False
End If
End Sub

