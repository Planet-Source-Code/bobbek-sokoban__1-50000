Attribute VB_Name = "Levels"
Global LevelData(1 To 21)
Public Function GetLevel(Level As Integer)
If Level = 0 Then
 LevelData(1) = "VVVVVVVVVVVVVVVVVVVV"
 LevelData(2) = "VXXXXXXXXXXXXXXXXXXV"
 LevelData(3) = "VXVVVVVXVVVVVXVXXVXV"
 LevelData(4) = "VXVXXXXXVXXXVXVXVXXV"
 LevelData(5) = "VXVVVVVXVXXXVXVVXXXV"
 LevelData(6) = "VXXXXXVXVXXXVXVXVXXV"
 LevelData(7) = "VXVVVVVXVVVVVXVXXVXV"
 LevelData(8) = "VXXXXXXXXXXXXXXXXXXV"
 LevelData(9) = "VXXXXXXVVVVVXXXXXXXV"
LevelData(10) = "VXXXXXXVXXXVXXXXXXXV"
LevelData(11) = "VXXXXXXVXXXVXXXXXXXV"
LevelData(12) = "VXXXXXXVVVVVXXXXXXXV"
LevelData(13) = "VXXXXXXXXXXXXXXXXXXV"
LevelData(14) = "VXVVVXXXXVXXXVXXXVXV"
LevelData(15) = "VXVXXVXXVXVXXVVXXVXV"
LevelData(16) = "VXVXVXXVXXXVXVXVXVXV"
LevelData(17) = "VXVXXVXVVVVVXVXXVVXV"
LevelData(18) = "VXVVVXXVXXXVXVXXXVXV"
LevelData(19) = "VXXXXXXXXXXXXXXXXXXV"
LevelData(20) = "VVVVVVVVVVVVVVVVVVVV"
End If
If Level = 1 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(5) = "XXXXVVVVVXXXXXXXXXXX"
 LevelData(6) = "XXXXV   VXXXXXXXXXXX"
 LevelData(7) = "XXXXVL  VXXXXXXXXXXX"
 LevelData(8) = "XXVVV  LVVXXXXXXXXXX"
 LevelData(9) = "XXV  L L VXXXXXXXXXX"
LevelData(10) = "VVV V VV VXXXVVVVVVX"
LevelData(11) = "V   V VV VVVVV  FFVX"
LevelData(12) = "V L  L          FFVX"
LevelData(13) = "VVVVV VVV V VV  FFVX"
LevelData(14) = "XXXXV     VVVVVVVVVX"
LevelData(15) = "XXXXVVVVVVVXXXXXXXXX"
LevelData(16) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(17) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "252" 'sokostart
End If
If Level = 2 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXVVVVVVVVVVVVXXXXX"
 LevelData(5) = "XXXVFF  V     VVVXXX"
 LevelData(6) = "XXXVFF  V L  L  VXXX"
 LevelData(7) = "XXXVFF  VLVVVV  VXXX"
 LevelData(8) = "XXXVFF      VV  VXXX"
 LevelData(9) = "XXXVFF  V V  L VVXXX"
LevelData(10) = "XXXVVVVVV VVL L VXXX"
LevelData(11) = "XXXXXV L  L L L VXXX"
LevelData(12) = "XXXXXV    V     VXXX"
LevelData(13) = "XXXXXVVVVVVVVVVVVXXX"
LevelData(14) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(15) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(16) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(17) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "151" 'SokoStart
End If
If Level = 3 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(5) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(6) = "XXXXXXXXXXVVVVVVVVXX"
 LevelData(7) = "XXXXXXXXXXV      VXX"
 LevelData(8) = "XXXXXXXXXXV LVL VXXX"
 LevelData(9) = "XXXXXXXXXXV L  LVXXX"
LevelData(10) = "XXXXXXXXXXVVL L VXXX"
LevelData(11) = "XXVVVVVVVVV L V VVVX"
LevelData(12) = "XXVFFFF  VV L  L  VX"
LevelData(13) = "XXVVFFF    L  L   VX"
LevelData(14) = "XXVFFFF  VVVVVVVVVVX"
LevelData(15) = "XXVVVVVVVVXXXXXXXXXX"
LevelData(16) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(17) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "137" 'SokoStart
End If
If Level = 4 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXVVVVVVVVV"
 LevelData(4) = "XXXXXXXXXXXV  FFFFFV"
 LevelData(5) = "VVVVVVVVVVVV  FFFFFV"
 LevelData(6) = "V    V  L L   FFFFFV"
 LevelData(7) = "V LLLVL  L V  FFFFFV"
 LevelData(8) = "V  L     L VVVVVVVVV"
 LevelData(9) = "V LL VL L LVXXXXXXXX"
LevelData(10) = "V  L V     VXXXXXXXX"
LevelData(11) = "VV VVVVVVVVVXXXXXXXX"
LevelData(12) = "V    V    VVXXXXXXXX"
LevelData(13) = "V     L    VXXXXXXXX"
LevelData(14) = "V  LLVLL  VVXXXXXXXX"
LevelData(15) = "V    V   VVVXXXXXXXX"
LevelData(16) = "VVVVVVVVVVXXXXXXXXXX"
LevelData(17) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "251" 'SokoStart
End If
If Level = 5 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXXXXXXXVVVVVXXXXXX"
 LevelData(5) = "XXXXXXXXXV   VVVVVXX"
 LevelData(6) = "XXXXXXXXXV VLVV  VXX"
 LevelData(7) = "XXXXXXXXXV     L VXX"
 LevelData(8) = "XVVVVVVVVV VVV   VXX"
 LevelData(9) = "XVFFFF  VV L  LVVVXX"
LevelData(10) = "XVFFFF    L LL VVXXX"
LevelData(11) = "XVFFFF  VVL  L  VXXX"
LevelData(12) = "XVVVVVVVVV  L  VVXXX"
LevelData(13) = "XXXXXXXXXV L L  VXXX"
LevelData(14) = "XXXXXXXXXVVV VV VXXX"
LevelData(15) = "XXXXXXXXXXXV    VXXX"
LevelData(16) = "XXXXXXXXXXXVVVVVVXXX"
LevelData(17) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "216"
End If
If Level = 6 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(5) = "XXXXVVVVVVXXVVVXXXXX"
 LevelData(6) = "XXXXVFF  VXVV VVXXXX"
 LevelData(7) = "XXXXVFF  VVV   VXXXX"
 LevelData(8) = "XXXXVFF     LL VXXXX"
 LevelData(9) = "XXXXVFF  V V L VXXXX"
LevelData(10) = "XXXXVFFVVV V L VXXXX"
LevelData(11) = "XXXXVVVV L VL  VXXXX"
LevelData(12) = "XXXXXXXV  LV L VXXXX"
LevelData(13) = "XXXXXXXV L  L  VXXXX"
LevelData(14) = "XXXXXXXV  VV   VXXXX"
LevelData(15) = "XXXXXXXVVVVVVVVVXXXX"
LevelData(16) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(17) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "114"
End If
If Level = 7 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(5) = "XXXXXXXXXXVVVVVXXXXX"
 LevelData(6) = "XXXXVVVVVVV   VVXXXX"
 LevelData(7) = "XXXVV V  VV LL VXXXX"
 LevelData(8) = "XXXV    L      VXXXX"
 LevelData(9) = "XXXV  L  VVV   VXXXX"
LevelData(10) = "XXXVVV VVVVVLVVVXXXX"
LevelData(11) = "XXXV L  VVV FFVXXXXX"
LevelData(12) = "XXXV L L L FFFVXXXXX"
LevelData(13) = "XXXV    VVVFFFVXXXXX"
LevelData(14) = "XXXV LL VXVFFFVXXXXX"
LevelData(15) = "XXXV  VVVXVVVVVXXXXX"
LevelData(16) = "XXXVVVVXXXXXXXXXXXXX"
LevelData(17) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "129"
End If
If Level = 8 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXVVVVXXXXXXXXXXXX"
 LevelData(4) = "XXXXV  VVVVVVVVVVVXX"
 LevelData(5) = "XXXXV    L   L L VXX"
 LevelData(6) = "XXXXV LV L V  L  VXX"
 LevelData(7) = "XXXXV  L L  V    VXX"
 LevelData(8) = "XXVVV LV V  VVVV VXX"
 LevelData(9) = "XXV VL L L  VV   VXX"
LevelData(10) = "XXV    L VLV   V VXX"
LevelData(11) = "XXV   L    L L L VXX"
LevelData(12) = "XXVVVVV  VVVVVVVVVXX"
LevelData(13) = "XXXXV      VXXXXXXXX"
LevelData(14) = "XXXXV      VXXXXXXXX"
LevelData(15) = "XXXXVFFFFFFVXXXXXXXX"
LevelData(16) = "XXXXVFFFFFFVXXXXXXXX"
LevelData(17) = "XXXXVFFFFFFVXXXXXXXX"
LevelData(18) = "XXXXVVVVVVVVXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "164"
End If
If Level = 9 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXXXXXXXXXXVVVVVVVX"
 LevelData(5) = "XXXXXXXXXXXXV  FFFVX"
 LevelData(6) = "XXXXXXXXVVVVV  FFFVX"
 LevelData(7) = "XXXXXXXXV      F FVX"
 LevelData(8) = "XXXXXXXXV  VV  FFFVX"
 LevelData(9) = "XXXXXXXXVV VV  FFFVX"
LevelData(10) = "XXXXXXXVVV VVVVVVVVX"
LevelData(11) = "XXXXXXXV LLL VVXXXXX"
LevelData(12) = "XXXVVVVV  L L VVVVVX"
LevelData(13) = "XXVV   VL L   V   VX"
LevelData(14) = "XXV  L  L    L  L VX"
LevelData(15) = "XXVVVVVV LL L VVVVVX"
LevelData(16) = "XXXXXXXV      VXXXXX"
LevelData(17) = "XXXXXXXVVVVVVVVXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "264"
End If
If Level = 10 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XVVVXXVVVVVVVVVVVVVX"
 LevelData(4) = "VV VVVV       V   VX"
 LevelData(5) = "V LL   LL  L L FFFVX"
 LevelData(6) = "V  LLLV    L  VFFFVX"
 LevelData(7) = "V L   V LL LL VFFFVX"
 LevelData(8) = "VVV   V  L    VFFFVX"
 LevelData(9) = "V     V L L L VFFFVX"
LevelData(10) = "V    VVVVVV VVVFFFVX"
LevelData(11) = "VV V  V  L L  VFFFVX"
LevelData(12) = "V  VV V LL L LVVFFVX"
LevelData(13) = "V FFV V  L      VFVX"
LevelData(14) = "V FFV V LLL LLL VFVX"
LevelData(15) = "VVVVV V       V VFVX"
LevelData(16) = "XXXXV VVVVVVVVV VFVX"
LevelData(17) = "XXXXV           VFVX"
LevelData(18) = "XXXXVVVVVVVVVVVVVVVX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "63"
End If
If Level = 11 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXXXXXXXXXVVVVXXXXX"
 LevelData(5) = "XXXXXXVVVVXV  VXXXXX"
 LevelData(6) = "XXXXVVV  VVVL VXXXXX"
 LevelData(7) = "XXXVV      L  VXXXXX"
 LevelData(8) = "XXVV  L LLVV VVXXXXX"
 LevelData(9) = "XXV  VLVV     VXXXXX"
LevelData(10) = "XXV V L LL V VVVXXXX"
LevelData(11) = "XXV   L V  V L VVVVV"
LevelData(12) = "XVVVV    V  LL V   V"
LevelData(13) = "XVVVV VV L         V"
LevelData(14) = "XVF    VVV  VVVVVVVV"
LevelData(15) = "XVFF FFVXVVVVXXXXXXX"
LevelData(16) = "XVFFFVFVXXXXXXXXXXXX"
LevelData(17) = "XVFFFFFVXXXXXXXXXXXX"
LevelData(18) = "XVVVVVVVXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "129"
End If
If Level = 12 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXVVVVVVVVVVVVVVVVXX"
 LevelData(5) = "XXV              VXX"
 LevelData(6) = "XXV V VVVVVV     VXX"
 LevelData(7) = "XXV V  L L L LV  VXX"
 LevelData(8) = "XXV V   L L   VV VVX"
 LevelData(9) = "XXV V VL L LVVVFFFVX"
LevelData(10) = "XXV V   L L  VVFFFVX"
LevelData(11) = "XXV VVVLLL L VVFFFVX"
LevelData(12) = "XXV     V VV VVFFFVX"
LevelData(13) = "XXVVVVV   VV VVFFFVX"
LevelData(14) = "XXXXXXVVVVV     VVVX"
LevelData(15) = "XXXXXXXXXXV     VXXX"
LevelData(16) = "XXXXXXXXXXVVVVVVVXXX"
LevelData(17) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "150"
End If
If Level = 13 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(5) = "XXXVVVVVVVVVXXXXXXXX"
 LevelData(6) = "XXVV   VV  VVVVVVXXX"
 LevelData(7) = "VVV     V  V    VVVX"
 LevelData(8) = "V  L VL V  V  FFF VX"
 LevelData(9) = "V V LV LVV V VFVF VX"
LevelData(10) = "V  V VL  V    F F VX"
LevelData(11) = "V L    L V V VFVF VX"
LevelData(12) = "V   VV  VVL L F F VX"
LevelData(13) = "V L V   V  VLVFVF VX"
LevelData(14) = "VV L  L   L  LFFF VX"
LevelData(15) = "XVL VVVVVV    VV  VX"
LevelData(16) = "XV  VXXXXVVVVVVVVVVX"
LevelData(17) = "XVVVVXXXXXXXXXXXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "167"
End If
If Level = 14 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXVVVVVVVXXXXX"
 LevelData(4) = "XXVVVVVVV     VXXXXX"
 LevelData(5) = "XXV     V L L VXXXXX"
 LevelData(6) = "XXVLL V   VVVVVVVVVX"
 LevelData(7) = "XXV VVVFFFFFFVV   VX"
 LevelData(8) = "XXV   LFFFFFFVV V VX"
 LevelData(9) = "XXV VVVFFFFFF     VX"
LevelData(10) = "XVV   VVVV VVV VLVVX"
LevelData(11) = "XV  VL   V  L  V VXX"
LevelData(12) = "XV  L LLL  V LVV VXX"
LevelData(13) = "XV   L L VVVLL V VXX"
LevelData(14) = "XVVVVV     L   V VXX"
LevelData(15) = "XXXXXVVV VVV   V VXX"
LevelData(16) = "XXXXXXXV     V   VXX"
LevelData(17) = "XXXXXXXVVVVVVVV  VXX"
LevelData(18) = "XXXXXXXXXXXXXXVVVVXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "92"
End If
If Level = 15 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXVVVVVVVVXXXXXXX"
 LevelData(4) = "XXXXXV   V  VXXXXXXX"
 LevelData(5) = "XXXXXV  L   VXXXXXXX"
 LevelData(6) = "XXXVVV VL   VVVVXXXX"
 LevelData(7) = "XXXV  L  VVL   VXXXX"
 LevelData(8) = "XXXV  V   L V LVXXXX"
 LevelData(9) = "XXXV  V      L VVVVX"
LevelData(10) = "XXXVV VVVVLVV     VX"
LevelData(11) = "XXXV LVFFFFFV V   VX"
LevelData(12) = "XXXV  LFFffF LV VVVX"
LevelData(13) = "XXVV  VFFFFFV   VXXX"
LevelData(14) = "XXV   VVV VVVVVVVXXX"
LevelData(15) = "XXV LL  V  VXXXXXXXX"
LevelData(16) = "XXV  V     VXXXXXXXX"
LevelData(17) = "XXVVVVVV   VXXXXXXXX"
LevelData(18) = "XXXXXXXVVVVVXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "149"
End If
If Level = 16 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXVVVVVXXXXXXXXXXXX"
 LevelData(5) = "XXXV   VVXXXXXXXXXXX"
 LevelData(6) = "XXXV    VXXVVVVXXXXX"
 LevelData(7) = "XXXV L  VVVV  VXXXXX"
 LevelData(8) = "XXXV  LL L   LVXXXXX"
 LevelData(9) = "XXXVVV  VL    VVXXXX"
LevelData(10) = "XXXXV  VV  L L VVXXX"
LevelData(11) = "XXXXV L  VV VV FVXXX"
LevelData(12) = "XXXXV  VLVVL  VFVXXX"
LevelData(13) = "XXXXVVV   LFFVVFVXXX"
LevelData(14) = "XXXXXV    VFfFFFVXXX"
LevelData(15) = "XXXXXV LL VFFFFFVXXX"
LevelData(16) = "XXXXXV  VVVVVVVVVXXX"
LevelData(17) = "XXXXXV  VXXXXXXXXXXX"
LevelData(18) = "XXXXXVVVVXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "167"
End If
If Level = 17 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXXVVVVVVVVVVXXXXXX"
 LevelData(5) = "XXXXVFF  V   VXXXXXX"
 LevelData(6) = "XXXXVFF      VXXXXXX"
 LevelData(7) = "XXXXVFF  V  VVVVXXXX"
 LevelData(8) = "XXXVVVVVVV  V  VVXXX"
 LevelData(9) = "XXXV            VXXX"
LevelData(10) = "XXXV  V  VV  V  VXXX"
LevelData(11) = "XVVVV VV  VVVV VVXXX"
LevelData(12) = "XV  L  VVVVV V  VXXX"
LevelData(13) = "XV V L  L  V L  VXXX"
LevelData(14) = "XV  L  L   V   VVXXX"
LevelData(15) = "XVVVV VV VVVVVVVXXXX"
LevelData(16) = "XXXXV    VXXXXXXXXXX"
LevelData(17) = "XXXXVVVVVVXXXXXXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "264"
End If
If Level = 18 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(4) = "XXXXXVVVVVVVVVVVXXXX"
 LevelData(5) = "XXXXXV  F  V   VXXXX"
 LevelData(6) = "XXXXXV VF      VXXXX"
 LevelData(7) = "XVVVVV VVFFV VVVVXXX"
 LevelData(8) = "VV  V FFVVV     VVVX"
 LevelData(9) = "V L VFFF   L V  L VX"
LevelData(10) = "V    FF VV  VV VV VX"
LevelData(11) = "VVVVLVVLV L V   V VX"
LevelData(12) = "XXVV V    VL LL V VX"
LevelData(13) = "XXV  L V V  V LVV VX"
LevelData(14) = "XXV               VX"
LevelData(15) = "XXV  VVVVVVVVVVV  VX"
LevelData(16) = "XXVVVVXXXXXXXXXVVVVX"
LevelData(17) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(18) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "114"
End If
If Level = 19 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXVVVVVVXXXXXXXXXXXX"
 LevelData(4) = "XXV    VVVVXXXXXXXXX"
 LevelData(5) = "VVVVV L   VXXXXXXXXX"
 LevelData(6) = "V   VV    VVVVXXXXXX"
 LevelData(7) = "V L V  VV    VXXXXXX"
 LevelData(8) = "V L V  VVVVV VXXXXXX"
 LevelData(9) = "VV L  L    V VXXXXXX"
LevelData(10) = "VV L L VVV V VXXXXXX"
LevelData(11) = "VV V  L  V V VXXXXXX"
LevelData(12) = "VV V VLV   V VXXXXXX"
LevelData(13) = "VV VVV   V V VVVVVVX"
LevelData(14) = "V  L  VVVV V VFFFFVX"
LevelData(15) = "V    L    L   FFVFVX"
LevelData(16) = "VVVVL  LV L   FFFFVX"
LevelData(17) = "V       V  VV FFFFVX"
LevelData(18) = "VVVVVVVVVVVVVVVVVVVX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "67"
End If
If Level = 20 Then
 LevelData(1) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(2) = "XXXXXXXXXXXXXXXXXXXX"
 LevelData(3) = "XXXXVVVVVVVVVVXXXXXX"
 LevelData(4) = "VVVVV        VVVVXXX"
 LevelData(5) = "V     V   L  V  VXXX"
 LevelData(6) = "V VVVVVVVLVVVV  VVVX"
 LevelData(7) = "V V    VV V  VL FFVX"
 LevelData(8) = "V V L  L  V  V  VFVX"
 LevelData(9) = "V V L  V     VL FFVX"
LevelData(10) = "V V  VVV VV     VFVX"
LevelData(11) = "V VVV  V  V  VL FFVX"
LevelData(12) = "V V    V LVVVV  VFVX"
LevelData(13) = "V VL   L  L  Vf FFVX"
LevelData(14) = "V    L V L L V  VFVX"
LevelData(15) = "VVVV LVVV    Vf FFVX"
LevelData(16) = "XXXV    LL VVVFFFFVX"
LevelData(17) = "XXXV      VVXVVVVVVX"
LevelData(18) = "XXXVVVVVVVVXXXXXXXXX"
LevelData(19) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(20) = "XXXXXXXXXXXXXXXXXXXX"
LevelData(21) = "95"
End If
End Function

Function GetOriginalPic(Pos As Integer)
posRow = ((Pos - (Pos Mod 20)) / 20) + 1
posCol = Pos Mod 20
tile = Mid(LevelData(posRow), posCol, 1)
If tile = "X" Then Form1.brick(Pos).Picture = Form1.tom.Picture
If tile = "L" Then Form1.brick(Pos).Picture = Form1.golv.Picture ' inte rita låda igen
If tile = "F" Then Form1.brick(Pos).Picture = Form1.target.Picture
If tile = "f" Then Form1.brick(Pos).Picture = Form1.target.Picture
If tile = " " Then Form1.brick(Pos).Picture = Form1.golv.Picture
If tile = "V" Then Form1.brick(Pos).Picture = Form1.vägg.Picture
End Function

Public Sub intro()
GetLevel (0)
    For i = 1 To 20
        a$ = LevelData(i)
        For j = 1 To 20
            tile = Mid$(a$, j, 1)
            If tile = "X" Then Form1.brick(((i - 1) * 20) + j).Picture = Form1.tom.Picture
            If tile = "V" Then Form1.brick(((i - 1) * 20) + j).Picture = Form1.vägg.Picture
        Next j
    Next i

End Sub