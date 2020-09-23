Attribute VB_Name = "modGraph"
Option Explicit

'------------------------------------------------Windows API
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
'-------------------------------------------------------Vars
Public ToolTips As Boolean

Private Const RGN_OR = 2

Public Enum OnOff
  setON = 1
  setOFF = 0
End Enum

'------------------------------------------------DigitalText
Public Sub DigitalText(ByVal pX As Long, ByVal pY As Long, ByVal Text As String, ByVal modo As Byte, Optional color As Long)
  Dim posChar As Long
  Dim StrTam As Long
  Dim CharVal As Long
  Dim PosX As Long
  
  StrTam = Len(Text)
  PosX = 0
  For posChar = 1 To StrTam
    CharVal = Asc(Mid(Text, posChar, 1))
    If (CharVal < 32) Or (CharVal > 90) Then CharVal = 45
    Select Case modo
      Case 1:
        Central.PaintPicture Central.Letters(0).Picture, pX + PosX, pY, 5, 8, (CharVal - 32) * 5, 0, 5, 8, vbSrcCopy
        PosX = PosX + 5
      Case 2:
        Central.PaintPicture Central.Letters(1).Picture, pX + PosX, pY, 5, 5, (CharVal - 32) * 5, 0, 5, 5, vbSrcCopy
        PosX = PosX + 5
      Case 3:
        If (CharVal < 48) Or (CharVal > 57) Then
          PosX = PosX + 2
        Else
          Central.PaintPicture Central.Numbers.Picture, pX + PosX, pY, 4, 5, (CharVal - 48) * 4, 0, 4, 5, vbSrcCopy
          PosX = PosX + 4
        End If
      Case 4:
        Call DigitalBitText(Central, pX, pY, Text, RGB(255, 0, 0), RGB(15, 40, 47), 1)
      Case 5:
        Call DigitalBitText(Central, pX, pY, Text, color, RGB(15, 40, 47), 1)
    End Select
  Next posChar
End Sub

'---------------------------------------------RegionFromMask
Public Function RegionFromMask(picSource As PictureBox, ByVal modo As Long, Optional lngTransColor As Long = -1) As Long
  Dim wndRgn As Long, wndRgnTmp As Long, wndRgnAux As Long
  Dim pX As Long, pY As Long
  Dim tX As Long, tY As Long
  Dim pixVal As Long
  Dim rX1 As Long, rX2 As Long
  Dim shiftY As Long
  
  If lngTransColor = -1 Then lngTransColor = RGB(255, 255, 255)
  wndRgn = 0
  'set form shape
  Select Case modo
    Case 0: 'accelerate process for Central view
      shiftY = 252
      tY = 405
      tX = picSource.Width
      wndRgn = CreateRectRgn(1, 1, tX + 1, 33)
    Case 1: 'accelerate process for Surface view
      shiftY = 0
      tY = 310
      tX = picSource.Width
      wndRgn = CreateRectRgn(1, 1, tX + 1, 285)
    Case 2: 'accelerate process for Small view
      shiftY = 252
      tY = 310
      tX = picSource.Width
      wndRgn = CreateRectRgn(1, 1, tX + 1, 33)
    Case 3: 'accelerate process for Full view
      shiftY = 0
      tY = 405
      tX = picSource.Width
      wndRgn = CreateRectRgn(1, 1, tX + 1, 285)
  End Select
  'get mask pixels
  For pY = 285 To tY
    pX = 1
    Do While pX <= tX
      Do While (GetPixel(picSource.hDC, pX - 1, pY - 1) = lngTransColor) And (pX <= tX)
        pX = pX + 1
      Loop
      If pX <= tX Then
        rX1 = pX
        Do While (GetPixel(picSource.hDC, pX - 1, pY - 1) <> lngTransColor) And (pX <= tX)
          pX = pX + 1
        Loop
        rX2 = pX - 1
        wndRgnTmp = CreateRectRgn(rX1, pY - shiftY, rX2 + 1, pY + 1 - shiftY)
        wndRgnAux = CombineRgn(wndRgn, wndRgn, wndRgnTmp, RGN_OR)
        Call DeleteObject(wndRgnTmp)
      End If
    Loop
  Next pY
  RegionFromMask = wndRgn
End Function

'---------------------------------------------BrilhoPicClose
Public Sub BrilhoPicClose(ByVal modo As OnOff)
  If modo = setON Then
    Central.PicClose.Picture = LoadResPicture(102, vbResBitmap)
  Else
    Central.PicClose.Picture = LoadResPicture(101, vbResBitmap)
  End If
End Sub

'----------------------------------------------BrilhoPicDisk
Public Sub BrilhoPicDisk(ByVal modo As OnOff)
  If modo = setON Then
    Central.PicDisk.Picture = LoadResPicture(113, vbResBitmap)
  Else
    Central.PicDisk.Picture = LoadResPicture(112, vbResBitmap)
  End If
End Sub

'--------------------------------------------BrilhoPicWindow
Public Sub BrilhoPicWindow(ByVal Index As Integer, ByVal modo As OnOff)
  If modo = setON Then
    Central.PicWindow(Index).Picture = LoadResPicture(104 + Index * 2, vbResBitmap)
  Else
    Central.PicWindow(Index).Picture = LoadResPicture(103 + Index * 2, vbResBitmap)
  End If
End Sub

'--------------------------------------DisplayCentralSurface
Public Sub DisplayCentralSurface(ByVal Index As Integer)
  Dim Pic As New StdPicture
  
  Select Case Index
    Case 0:
      Set Pic = Central.CentralPics(3).Picture
      Central.PaintPicture Pic, 0, 0, 572, 32, 0, 0, 572, 32, vbSrcCopy
    Case 1:
      Set Pic = Central.CentralPics(2).Picture
      Central.PaintPicture Pic, 0, 0, 572, 284, 0, 0, 572, 284, vbSrcCopy
    Case 2:
      Set Pic = Central.CentralPics(3).Picture
      Central.PaintPicture Pic, 0, 0, 572, 32, 0, 0, 572, 32, vbSrcCopy
    Case 3:
      Set Pic = Central.CentralPics(2).Picture
      Central.PaintPicture Pic, 0, 0, 572, 284, 0, 0, 572, 284, vbSrcCopy
  End Select
  Set Pic = Nothing
End Sub

'---------------------------------------------DisplayCentral
Public Sub DisplayCentral(ByVal Index As Integer)
  Dim Pic As New StdPicture
  
  Select Case Index
    Case 0:
      Set Pic = Central.CentralPics(0).Picture
      Central.PaintPicture Pic, 0, 32, 572, 121, 0, 0, 572, 121, vbSrcCopy
    Case 1:
      Set Pic = LoadResPicture(114, vbResBitmap)
      Central.PaintPicture Pic, 0, 284, 572, 26, 0, 0, 572, 26, vbSrcCopy
    Case 2:
      Set Pic = LoadResPicture(114, vbResBitmap)
      Central.PaintPicture Pic, 0, 32, 572, 26, 0, 0, 572, 26, vbSrcCopy
    Case 3:
      Set Pic = Central.CentralPics(0).Picture
      Central.PaintPicture Pic, 0, 284, 572, 121, 0, 0, 572, 121, vbSrcCopy
  End Select
  oldNow = -1
  Set Pic = Nothing
End Sub

'-------------------------------------------RadioButtonCheck
Public Sub RadioButtonCheck(ByVal Aceso As Boolean, ByVal X As Long, ByVal Y As Long)
  Dim Pic As New StdPicture
  
  If Aceso Then
    Set Pic = LoadResPicture(116, vbResBitmap)
  Else
    Set Pic = LoadResPicture(115, vbResBitmap)
  End If
  Central.PaintPicture Pic, X, Y, 12, 11, 0, 0, 12, 11, vbSrcCopy
  Set Pic = Nothing
End Sub

'------------------------------------------------ControlDown
Public Sub ControlDown(ByVal modo As Byte, ByVal MainOp As Byte, ByVal SubOp As Byte)
  Dim Y As Long
  
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Y = Central.PicCentral.Top - 21
  
  Select Case MainOp
    Case 1: 'Scan (Mark/Jump/Depth/Copy)
      Call DrawBox3D(Central, modo, 214, 31 + Y + (SubOp - 7) * 15, 32, 15)
    Case 2: 'Format (Mark/Jump)
      Call DrawBox3D(Central, modo, 214, 31 + Y + (SubOp - 3) * 15, 32, 15)
    Case 3: 'Recover (Mark/Jump/Depth/Up/Down/File)
      Select Case SubOp
        Case 3, 4, 5: 'Mark/Jump/Depth
          Call DrawBox3D(Central, modo, 214, 31 + Y + (SubOp - 3) * 15, 32, 15)
        Case 6: 'Up
          Call DrawBox3D(Central, modo, 149, 65 + Y, 11, 6)
        Case 7: 'Down
          Call DrawBox3D(Central, modo, 138, 65 + Y, 11, 6)
        Case 8: 'File
          Call DrawBox3D(Central, modo, 223, 78 + Y, 19, 10)
      End Select
    Case 4: 'Edit (Format/Overwrite/Mark/Unmark)
      Call DrawBox3D(Central, modo, 219, 31 + Y + (SubOp - 6) * 15, 27, 15)
   'Case 5: N Read
    Case 6: 'Main
      Call DrawBox3D(Central, modo, 264, 28 + Y + (SubOp - 1) * 18, 51, 17)
  End Select
End Sub

'------------------------------------------DisplayEditValues
Public Sub DisplayEditValues()
  Dim Y As Long
  Dim Bad As Long, Good As Long, Avail As Long, Percent As Long
    
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Y = Central.PicCentral.Top - 21
  Call CountSectors(Bad, Good, Avail, Percent)
  Call DigitalText(181, 47 + Y, Str0N(Good, 4), 3)
  Call DigitalText(177, 53 + Y, Str0N(Good * 512, 7), 3)
  Call DigitalText(177, 59 + Y, Str0N(Avail, 7), 3)
  Call DigitalText(177, 65 + Y, Str0N(Bad * 512, 7), 3)
  Call DigitalText(181, 71 + Y, Str0N(Bad, 4), 3)
End Sub
  
'-----------------------------------------DisplayEditOpLight
Public Sub DisplayEditOpLight()
  Dim Y As Long
  Dim i As Long
  Dim Cor As Long
  
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Y = Central.PicCentral.Top - 21
  For i = 1 To 4
    Cor = 0
    If mLightEdit(i + 5) = True Then Cor = RGB(143, 250, 207)
    Central.ForeColor = Cor
    Central.FillColor = Cor
    Central.Line (239, 20 + Y + i * 15)-(241, 25 + Y + i * 15), , BF
  Next i
End Sub

'----------------------------------------------DisplayScanOp
Public Sub DisplayScanOp()
  Dim i As Byte
  Dim Y As Long
  Dim Cor As Long
  
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Y = Central.PicCentral.Top - 21
  'Light Copy/Depth/Mark/Jump
  For i = 1 To 4
    Cor = 0
    If mLightScan(i + 6) = True Then Cor = RGB(143, 250, 207)
    Central.ForeColor = Cor
    Central.FillColor = Cor
    Central.Line (239, 20 + Y + i * 15)-(241, 25 + Y + i * 15), , BF
  Next i
  'Repair/Check/User
  Call RadioButtonCheck(mLightScan(1), 143, 44 + Y)
  Call RadioButtonCheck(mLightScan(2), 143, 61 + Y)
  Call RadioButtonCheck(mLightScan(3), 143, 78 + Y)
  'Read/Write/Verify
  Call RadioButtonCheck(mLightScan(4), 178, 51 + Y)
  Call RadioButtonCheck(mLightScan(5), 175, 65 + Y)
  Call RadioButtonCheck(mLightScan(6), 188, 61 + Y)
End Sub

'----------------------------------------------DisplayFormatOp
Public Sub DisplayFormatOp()
  Dim i As Byte
  Dim Y As Long
  Dim Cor As Long
  Dim Bad As Long, Good As Long, Avail As Long, Percent As Long
  
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Y = Central.PicCentral.Top - 21
  'Light Mark/Jump/Read/Verify
  For i = 1 To 4
    Cor = 0
    If mLightFormat(i + 2) = True Then Cor = RGB(143, 250, 207)
    Central.ForeColor = Cor
    Central.FillColor = Cor
    Central.Line (239, 20 + Y + i * 15)-(241, 25 + Y + i * 15), , BF
  Next i
  'Full/Quick
  Call RadioButtonCheck(mLightFormat(1), 195, 44 + Y)
  Call RadioButtonCheck(mLightFormat(2), 195, 61 + Y)
  'Values
  Call CountSectors(Bad, Good, Avail, Percent)
  Call DigitalINT(Central, 135, 42 + Y, Bad, 2, 4)
  Call DigitalINT(Central, 135, 57 + Y, Good, 2, 4)
  Call DigitalINT(Central, 169, 79 + Y, Avail, 2, 7)
  Call DigitalText(127, 79 + Y, Str03(Percent), 3)
End Sub

'---------------------------------------------DisplayRecoverOp
Public Sub DisplayRecoverOp()
  Dim i As Byte
  Dim Y As Long
  Dim Cor As Long
  Dim sleft As Long
  
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Y = Central.PicCentral.Top - 21
  'Light Mark/Jump/Depth
  For i = 1 To 3
    Cor = 0
    If mLightRecover(i + 2) = True Then Cor = RGB(143, 250, 207)
    Central.ForeColor = Cor
    Central.FillColor = Cor
    Central.Line (239, 20 + Y + i * 15)-(241, 25 + Y + i * 15), , BF
  Next i
  'Save/Load
  Call RadioButtonCheck(mLightRecover(1), 196, 45 + Y)
  Call RadioButtonCheck(mLightRecover(2), 196, 62 + Y)
  'Next Jump
  Call DigitalINT(Central, 135, 54 + Y, mJumpNext, 2, 4)
  'File
  Call DigitalText(115, 80 + Y, mSaveName, 2)
  'Sectors left
  Call DigitalINT(Central, 135, 41 + Y, CountSecCopy(), 2, 4)
End Sub

'------------------------------------------------DisplayEditOp
Public Sub DisplayEditOp()
  Dim i As Byte
  Dim Y As Long
  Dim Cor As Long
  
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Y = Central.PicCentral.Top - 21
  'Light Format/Overwrite/Mark/Unmark
  Call DisplayEditOpLight
  'Mark/Copy/Read/Write/Verify
  Call RadioButtonCheck(mLightEdit(1), 115, 40 + Y)
  Call RadioButtonCheck(mLightEdit(2), 115, 50 + Y)
  Call RadioButtonCheck(mLightEdit(3), 115, 60 + Y)
  Call RadioButtonCheck(mLightEdit(4), 115, 70 + Y)
  Call RadioButtonCheck(mLightEdit(5), 115, 80 + Y)
  'Values
  Call DisplayEditValues
End Sub

'------------------------------------------DisplayReadSlider
Public Sub DisplayReadSlider()
  Dim i As Byte
  Dim Y As Long
  
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Y = Central.PicCentral.Top - 21
  For i = 1 To 4
    If mLightRead <> i Then
      Central.FillColor = 0
      Central.ForeColor = 0
      Central.Line (19 + (4 - i) * 25, 106 + Y)-(32 + (4 - i) * 25, 108 + Y), , BF
    Else
      Central.PaintPicture LoadResPicture(111, vbResBitmap), 19 + (4 - i) * 25, 106 + Y, 14, 3, 0, 0, 14, 3, vbSrcCopy
    End If
  Next i
End Sub

'---------------------------------------------DisplayGeralOp
Public Sub DisplayGeralOp()
  Dim Y As Long
  Dim i As Byte
  
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Y = Central.PicCentral.Top - 21
  For i = 1 To 4
    If mOperation = i Then
      Central.ForeColor = RGB(143, 250, 207)
      FilledTriangle Central.hDC, RGB(143, 250, 207), 268, Y + 32 + (i - 1) * 18, 276, Y + 32 + (i - 1) * 18, 268, Y + 40 + (i - 1) * 18
      Call DrawBox3D(Central, 1, 264, 28 + Y + (i - 1) * 18, 51, 17)
    Else
      Central.ForeColor = 0
      FilledTriangle Central.hDC, 0, 268, Y + 32 + (i - 1) * 18, 276, Y + 32 + (i - 1) * 18, 268, Y + 40 + (i - 1) * 18
      Call DrawBox3D(Central, 2, 264, 28 + Y + (i - 1) * 18, 51, 17)
    End If
  Next i
End Sub

'-------------------------------------------DisplayCentralOp
Public Sub DisplayCentralOp()
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Select Case mOperation
    Case 1: Central.PaintPicture Central.CentralPics(4).Picture, 105, 27 + Central.PicCentral.Top - 21, 145, 68, 0, 0, 145, 68, vbSrcCopy
    Case 2: Central.PaintPicture Central.CentralPics(5).Picture, 105, 27 + Central.PicCentral.Top - 21, 145, 68, 0, 0, 145, 68, vbSrcCopy
    Case 3: Central.PaintPicture Central.CentralPics(6).Picture, 105, 27 + Central.PicCentral.Top - 21, 145, 68, 0, 0, 145, 68, vbSrcCopy
    Case 4: Central.PaintPicture Central.CentralPics(7).Picture, 105, 27 + Central.PicCentral.Top - 21, 145, 68, 0, 0, 145, 68, vbSrcCopy
  End Select
End Sub

'----------------------------------------------DisplayGoText
Public Sub DisplayGoText()
  Dim Y As Long
  
  If PosGO <> 0 Then Exit Sub
  Y = 32
  If (mModWin = 1) Or (mModWin = 3) Then Y = 284
  If mWork = 0 Then
    Call DigitalText(273, 12 + Y, "       ", 1)
    Select Case mOperation
      Case 1: Call DigitalText(280, 12 + Y, "SCAN", 1)
      Case 2: Call DigitalText(275, 12 + Y, "FORMAT", 1)
      Case 3: Call DigitalText(273, 12 + Y, "RECOVER", 1)
      Case 4: Call DigitalText(280, 12 + Y, "EDIT", 1)
    End Select
  Else
    Call DigitalText(273, 12 + Y, "       ", 1)
    Select Case mOperation
      Case 1, 2, 3, 4: Call DigitalText(280, 12 + Y, "STOP", 4)
    End Select
  End If
End Sub

'---------------------------------------------DisplayTextTip
Public Sub DisplayTextTip(ByVal Texto As String)
  Dim Just As Boolean
  
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Just = False
  Do While Len(Texto) < 40
    If Just = False Then Texto = Texto & " "
    If Just = True Then Texto = " " & Texto
    Just = Not (Just)
  Loop
  Call DigitalText(187, 106 + Central.PicCentral.Top - 21, Texto, 1)
End Sub

'---------------------------------------------DisplayToolTip
Public Sub DisplayToolTip(ByVal mOp As Byte, ByVal sOp As Byte)
  Dim Y As Long
  
  If ToolTips = False Then Exit Sub
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Select Case mOp
    Case 1:  'Scan
      Select Case sOp
        Case 1: 'Repair
          Call DisplayTextTip("SIMPLE ATTEMPT TO REPAIR FLOPPY DISK")
        Case 2: 'Check
          Call DisplayTextTip("CHECK FOR POSSIBLE BAD SECTORS")
        Case 3: 'User
          Call DisplayTextTip("SCAN DISK WITH USER DEFINED OPTIONS ")
        Case 4: 'Read
          Call DisplayTextTip("PERFORM READ OPERATION")
        Case 5: 'Write
          Call DisplayTextTip("PERFORM WRITE AND READ OPERATION")
        Case 6: 'Verify
          Call DisplayTextTip("PERFORM DISK VERIFY OPERATION")
        Case 7: 'Mark
          Call DisplayTextTip("MARK BAD SECTORS WHEN BAD SECTOR FOUND")
        Case 8: 'Jump
          Call DisplayTextTip("JUMP BAD SECTORS (DO NOT TEST)")
        Case 9: 'Depth
          Call DisplayTextTip("SEARCH BAD SECTOR POSITION")
        Case 10: 'Copy
          Call DisplayTextTip("COPY GOOD SECTOR TO MEMORY FOR LATER USE")
      End Select
    Case 2:  'Format
      Select Case sOp
        Case 1: 'Full
          Call DisplayTextTip("FULL FORMAT - DISK INITIALIZATION")
        Case 2: 'Quick
          Call DisplayTextTip("QUICK FORMAT - ERASE ALL INFORMATION")
        Case 3: 'Mark
          Call DisplayTextTip("MARK BAD TRACK WHEN BAD TRACK FOUND")
        Case 4: 'Jump
          Call DisplayTextTip("JUMP BAD SECTORS (DO NOT FORMAT)")
        Case 5: 'Read
          Call DisplayTextTip("PERFORM READ OPERATION")
        Case 6: 'Verify
          Call DisplayTextTip("PERFORM DISK VERIFY OPERATION")
        Case 7: 'nº Bad
          Call DisplayTextTip("NUMBER OF BAD SECTORS")
        Case 8: 'nº Data
          Call DisplayTextTip("NUMBER OF DATA SECTORS")
        Case 9: '% Free
          Call DisplayTextTip("PERCENTAGE OF FREE SPACE")
        Case 10: 'Data Space
          Call DisplayTextTip("DATA SPACE IN BYTES")
      End Select
    Case 3:  'Recover
      Select Case sOp
        Case 1: 'Save
          Call DisplayTextTip("SAVE DISK IMAGE TO DISK")
        Case 2: 'Load
          Call DisplayTextTip("LOAD SAVED FILE TO DISK")
        Case 3: 'Mark
          Call DisplayTextTip("MARK BAD SECTORS AS IN ORIGINAL")
        Case 4: 'Jump
          Call DisplayTextTip("JUMP BAD SECTORS (DO NOT READ)")
        Case 5: 'Depth
          Call DisplayTextTip("SEARCH BAD SECTOR POSITION")
        Case 6: 'Up
          Call DisplayTextTip("INCREASE NUMBER OF RETRY")
        Case 7: 'Down
          Call DisplayTextTip("DECREASE NUMBER OF RETRY")
        Case 8: 'File
          Call DisplayTextTip("NAME OF SAVED FILE")
        Case 9: 'Sectors Left
          Call DisplayTextTip("SECTORS LEFT TO HAVE FULL DISK")
        Case 10: 'Next Jump
          Call DisplayTextTip("READS LEFT TO ABANDON SECTOR")
      End Select
    Case 4:  'Edit
      Select Case sOp
        Case 1: 'Mark
          Call DisplayTextTip("MARK BAD SECTORS WHEN FOUND")
        Case 2: 'Copy
          Call DisplayTextTip("COPY GOOD SECTOR TO MEMORY FOR LATER USE")
        Case 3: 'Read
          Call DisplayTextTip("PERFORM READ OPERATION")
        Case 4: 'Write
          Call DisplayTextTip("PERFORM WRITE AND READ OPERATION")
        Case 5: 'Verify
          Call DisplayTextTip("PERFORM VERIFY DISK OPERATION")
        Case 6: 'Format
          Call DisplayTextTip("FORMAT TRACK - ERASE DATA")
        Case 7: 'Overwrite
          Call DisplayTextTip("OVERWRITE SECTORS - ERASE DATA")
        Case 8: 'Mark
          Call DisplayTextTip("SET SECTORS AS BAD")
        Case 9: 'Unmark
          Call DisplayTextTip("SET SECTORS AS GOOD")
        Case 10: 'nº Free
          Call DisplayTextTip("NUMBER OF FREE SECTORS")
        Case 11: 'Spc. Free
          Call DisplayTextTip("FREE SPACE IN BYTES")
        Case 12: 'Available
          Call DisplayTextTip("AVAILABLE SPACE IN BYTES")
        Case 13: 'Spc. Bad
          Call DisplayTextTip("BAD SPACE IN BYTES")
        Case 14: 'nº Bad
          Call DisplayTextTip("NUMBER OF BAD SECTORS")
      End Select
    Case 5: 'Read N slider
      Call DisplayTextTip("NUMBER OF SECTORS TO TEST EACH TIME")
    Case 6: 'Main Operations
      Select Case sOp
        Case 1: 'Scan
          Call DisplayTextTip("DISK SCAN - SEARCH AND MARK SECTORS")
        Case 2: 'Format
          Call DisplayTextTip("FORMAT DISK - INITIALIZATION")
        Case 3: 'Recover
          Call DisplayTextTip("RECOVER DATA IN BAD SECTORS")
        Case 4: 'Edit
          Call DisplayTextTip("EDIT BAD SECTOR TABLE")
      End Select
    Case 7: 'Others
      Select Case sOp
        Case 1: 'Surface
          Call DisplayTextTip("DISK SURFACE")
        Case 2: 'Tealtech
          Call DisplayTextTip("CLICK TO SEE ABOUT BOX")
        Case 3: 'Exit
          Call DisplayTextTip("CLICK TO LEAVE DISKTEST PRO")
        Case 4: 'Disk
          Call DisplayTextTip("CLICK TO CHANGE/REFRESH DISK")
        Case 5: 'GO
          Call DisplayTextTip("CLICK TO START ACTIVE OPERATION")
        Case 6: 'DTPRO
          Call DisplayTextTip("CLICK TO GET HELP WITH DISKTEST PRO")
        Case 7: 'Position
          Call DisplayTextTip("CURRENT DISK TEST POSITION")
        Case 8: 'Wave
          Call DisplayTextTip("AVERAGE READ TIME")
      End Select
    Case 8: 'Views
      Select Case sOp
        Case 1: 'Central
          Call DisplayTextTip("CENTRAL VIEW")
        Case 2: 'Surface
          Call DisplayTextTip("SURFACE VIEW")
        Case 3: 'Small
          Call DisplayTextTip("SMALL VIEW")
        Case 4: 'Full
          Call DisplayTextTip("FULL VIEW")
      End Select
    Case 9: 'Time
      Select Case sOp
        Case 1: 'Current
          Call DisplayTextTip("CURRENT TIME")
        Case 2: 'Finish
          Call DisplayTextTip("FINISH TIME")
        Case 3: 'Elapsed
          Call DisplayTextTip("TIME ELAPSED")
        Case 4: 'Left
          Call DisplayTextTip("TIME LEFT")
        Case 5: 'Total
          Call DisplayTextTip("TOTAL TIME")
      End Select
    Case Else
      Y = 106 + Central.PicCentral.Top - 21
      Call DigitalText(187, Y, StrN(40, " "), 1)
  End Select
End Sub
