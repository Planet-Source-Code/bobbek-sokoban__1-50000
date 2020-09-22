Attribute VB_Name = "ScoreAndSave"
Option Explicit

Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)
Public Declare Function GetPrivateProfileInt& Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String)
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String)
Public INIFile As String

Public Function GetINIKeyStr(strSection As String, strKey As String)
    Dim strInfo As String
    Dim lngRet As Long
    
    strInfo = String(260, " ")
    GetPrivateProfileString strSection, strKey, "", strInfo, 260, INIFile
    GetINIKeyStr = Trim(strInfo)
End Function

Public Function GetINIKeyInt(strSection As String, strKey As String)
    GetINIKeyInt = GetPrivateProfileInt(strSection, strKey, 0, INIFile)
End Function

Public Sub SetINIKey(strSection As String, strKey As String, strData As String)
    WritePrivateProfileString strSection, strKey, Trim$(strData), INIFile
End Sub

Public Sub DeleteINISection(strSection As String)
    WritePrivateProfileString strSection, vbNullString, vbNullString, INIFile
End Sub

Public Sub DeleteINIKey(strSection As String, strKey As String)
    WritePrivateProfileString strSection, strKey, vbNullString, INIFile
End Sub

Public Sub checkScore()
Dim CurrLevl
Dim CurrMoves
Dim MaxScore
Dim Level As String
Dim answer
Dim i

CurrLevl = Val(Form1.lblLevel.Caption)
CurrMoves = Val(Form1.lblMoves.Caption)
Level = "Level" & Trim$(Str$(CurrLevl))
'Debug.Print Level
MaxScore = Val(GetINIKeyStr("HISCORE", Level))
If CurrMoves < MaxScore Then
    SetINIKey "HISCORE", Level, Trim$(Str$(CurrMoves))
    MsgBox "Last HiScore: " & MaxScore & vbNewLine & "Your Score: " & CurrMoves, vbOKOnly, "* * * * NEW HISCORE * * * *"
    answer = MsgBox("Do You Want to SAVE moves to Autoplay?", vbQuestion + vbYesNo + vbDefaultButton1, "GOOD ONE!")
    If answer = vbYes Then
        Open App.Path & "\Bana" & Trim(Form1.lblLevel.Caption) & ".txt" For Output As #1
        For i = 1 To Totalmoves
        Print #1, UndoSokoTo(i)
        Next i
        Close #1
    End If
ElseIf Not Form1.tmrAutoPlay Then
    MsgBox "That Wasn't too good..." & vbNewLine & "Current HiScore: " & MaxScore & vbNewLine & "Your Score: " & CurrMoves, vbOKOnly, "* * * * NOT A NEW HISCORE * * * *"
End If
 
End Sub
