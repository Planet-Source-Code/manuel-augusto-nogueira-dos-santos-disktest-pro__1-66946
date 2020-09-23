Attribute VB_Name = "modCentral"
Option Explicit

'-----------------------------------------------------Public
Public mModWin As Long                  'modo para tamanho janela 0-Central View,1-Surface View,2-Small View,3-Full View
Public mLightScan(1 To 10) As Boolean   'ON/OFF para opções Scan
Public mLightFormat(1 To 6) As Boolean  'ON/OFF para opções Format
Public mLightRecover(1 To 5) As Boolean 'ON/OFF para opções Recover
Public mLightEdit(1 To 9) As Boolean    'ON/OFF para opções Scan
Public mUserOp(1 To 10) As Boolean      'ON/OFF para user pref
Public mOperation As Byte               'Posição para operações
Public mLightRead As Byte               'Posição para Read N
Public mJumpNext As Long                'number to Jump Next Sector
Public oldNow As Long                   'Control old clock display
Public PosGO As Long                    'Go Animation
Public MouseGO As Boolean               'Mouse in GO button
Public mSaveName As String              'Image name for Recover Save/Load

'----------------------------------------------ReDisplayTool
Public Sub ReDisplayTool()
  Select Case mOperation
    Case 1: Call DisplayScanOp
    Case 2: Call DisplayFormatOp
    Case 3: Call DisplayRecoverOp
    Case 4: Call DisplayEditOp
  End Select
End Sub

'---------------------------------------------ToolTipAtMouse
Public Sub ToolTipAtMouse(ByRef MainOp As Byte, ByRef SubOp As Byte, ByVal X As Long, ByVal Y As Long)
  Dim Y1 As Long
  
  Y1 = Y + Central.PicCentral.Top - 21
  Select Case mOperation
    Case 1:  'Scan
      'Repair/Check/User
      If IsInsideBox(X, Y, 114, 40, 41, 13) Then  'Repair
        MainOp = 1: SubOp = 1: Exit Sub
      End If
      If IsInsideBox(X, Y, 114, 57, 41, 13) Then  'Check
        MainOp = 1: SubOp = 2: Exit Sub
      End If
      If IsInsideBox(X, Y, 114, 74, 41, 13) Then  'User
        MainOp = 1: SubOp = 3: Exit Sub
      End If
      'Read/Write/Verify
      If IsInsideBox(X, Y, 173, 39, 14, 20) Then 'Read
        MainOp = 1: SubOp = 4: Exit Sub
      End If
      If IsInsideBox(X, Y, 165, 62, 18, 19) Then 'Write
        MainOp = 1: SubOp = 5: Exit Sub
      End If
      If IsInsideBox(X, Y, 186, 59, 19, 13) Then 'Verify
        MainOp = 1: SubOp = 6: Exit Sub
      End If
      'Mark/Jump/Depth/Copy
      If IsInsideBox(X, Y, 214, 31, 32, 15) Then  'Mark
        MainOp = 1: SubOp = 7: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 46, 32, 15) Then  'Jump
        MainOp = 1: SubOp = 8: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 61, 32, 15) Then  'Depth
        MainOp = 1: SubOp = 9: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 76, 32, 15) Then  'Copy
        MainOp = 1: SubOp = 10: Exit Sub
      End If
    Case 2:  'Format
      'Full/Quick
      If IsInsideBox(X, Y, 166, 40, 41, 13) Then  'Full
        MainOp = 2: SubOp = 1: Exit Sub
      End If
      If IsInsideBox(X, Y, 166, 57, 41, 13) Then  'Quick
        MainOp = 2: SubOp = 2: Exit Sub
      End If
      'Mark/Jump
      If IsInsideBox(X, Y, 214, 31, 32, 15) Then  'Mark
        MainOp = 2: SubOp = 3: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 46, 32, 15) Then  'Jump
        MainOp = 2: SubOp = 4: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 61, 32, 15) Then  'Read
        MainOp = 2: SubOp = 5: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 76, 32, 15) Then  'Verify
        MainOp = 2: SubOp = 6: Exit Sub
      End If
      If IsInsideBox(X, Y, 112, 39, 50, 15) Then  'Bad Sectors
        MainOp = 2: SubOp = 7: Exit Sub
      End If
      If IsInsideBox(X, Y, 112, 54, 50, 15) Then  'Data Sectors
        MainOp = 2: SubOp = 8: Exit Sub
      End If
      If IsInsideBox(X, Y, 111, 75, 30, 13) Then  '% Free
        MainOp = 2: SubOp = 9: Exit Sub
      End If
      If IsInsideBox(X, Y, 146, 76, 68, 15) Then  'Data Space
        MainOp = 2: SubOp = 10: Exit Sub
      End If
    Case 3:  'Recover
      'Save/Load
      If IsInsideBox(X, Y, 167, 41, 41, 13) Then   'Save
        MainOp = 3: SubOp = 1: Exit Sub
      End If
      If IsInsideBox(X, Y, 167, 58, 41, 13) Then   'Load
        MainOp = 3: SubOp = 2: Exit Sub
      End If
      'Mark/Jump/Depth
      If IsInsideBox(X, Y, 214, 31, 32, 15) Then  'Mark
        MainOp = 3: SubOp = 3: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 46, 32, 15) Then  'Jump
        MainOp = 3: SubOp = 4: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 61, 32, 15) Then  'Depth
        MainOp = 3: SubOp = 5: Exit Sub
      End If
      'Up/Down
      If IsInsideBox(X, Y, 149, 65, 11, 6) Then   'Up
        MainOp = 3: SubOp = 6: Exit Sub
      End If
      If IsInsideBox(X, Y, 138, 65, 11, 6) Then   'Down
        MainOp = 3: SubOp = 7: Exit Sub
      End If
      'File
      If IsInsideBox(X, Y, 109, 69, 25, 7) Then  'File
        MainOp = 3: SubOp = 8: Exit Sub
      End If
      If IsInsideBox(X, Y, 109, 76, 137, 15) Then 'File
        MainOp = 3: SubOp = 8: Exit Sub
      End If
      'Sectors Left
      If IsInsideBox(X, Y, 114, 39, 46, 13) Then 'Sectors left
        MainOp = 3: SubOp = 9: Exit Sub
      End If
      'Next Jump
      If IsInsideBox(X, Y, 114, 52, 46, 13) Then 'Next Jump
        MainOp = 3: SubOp = 10: Exit Sub
      End If
    Case 4:  'Edit
      'Mark/Copy/Read/Write/Verify
      If IsInsideBox(X, Y, 112, 39, 37, 9) Then  'Mark
        MainOp = 4: SubOp = 1: Exit Sub
      End If
      If IsInsideBox(X, Y, 112, 49, 37, 9) Then  'Copy
        MainOp = 4: SubOp = 2: Exit Sub
      End If
      If IsInsideBox(X, Y, 112, 59, 37, 9) Then  'Read
        MainOp = 4: SubOp = 3: Exit Sub
      End If
      If IsInsideBox(X, Y, 112, 69, 37, 9) Then  'Write
        MainOp = 4: SubOp = 4: Exit Sub
      End If
      If IsInsideBox(X, Y, 112, 79, 37, 9) Then  'Verify
        MainOp = 4: SubOp = 5: Exit Sub
      End If
      'Format/Overwrite/Mark/Unmark
      If IsInsideBox(X, Y, 219, 31, 27, 15) Then  'Format
        MainOp = 4: SubOp = 6: Exit Sub
      End If
      If IsInsideBox(X, Y, 219, 46, 27, 15) Then  'Overwrite
        MainOp = 4: SubOp = 7: Exit Sub
      End If
      If IsInsideBox(X, Y, 219, 61, 27, 15) Then  'Mark
        MainOp = 4: SubOp = 8: Exit Sub
      End If
      If IsInsideBox(X, Y, 219, 76, 27, 15) Then  'Unmark
        MainOp = 4: SubOp = 9: Exit Sub
      End If
      'Space
      If IsInsideBox(X, Y, 181, 47, 15, 5) Then   'Nº Free
        MainOp = 4: SubOp = 10: Exit Sub
      End If
      If IsInsideBox(X, Y, 177, 53, 27, 5) Then   'Free in Bytes
        MainOp = 4: SubOp = 11: Exit Sub
      End If
      If IsInsideBox(X, Y, 177, 59, 27, 5) Then   'Available
        MainOp = 4: SubOp = 12: Exit Sub
      End If
      If IsInsideBox(X, Y, 177, 65, 27, 5) Then   'Bad in Bytes
        MainOp = 4: SubOp = 13: Exit Sub
      End If
      If IsInsideBox(X, Y, 181, 71, 15, 5) Then   'Nº Bad
        MainOp = 4: SubOp = 14: Exit Sub
      End If
  End Select
  'Read N slider
  If IsInsideBox(X, Y, 19, 104, 89, 13) Then   'n read
    MainOp = 5
    SubOp = 4 - ((X - 19) * 4) \ 89
    Exit Sub
  End If
  'Main actions
  If IsInsideBox(X, Y, 264, 28, 51, 17) Then
    MainOp = 6: SubOp = 1: Exit Sub
  End If
  If IsInsideBox(X, Y, 264, 46, 51, 17) Then
    MainOp = 6: SubOp = 2: Exit Sub
  End If
  If IsInsideBox(X, Y, 264, 64, 51, 17) Then
    MainOp = 6: SubOp = 3: Exit Sub
  End If
  If IsInsideBox(X, Y, 264, 82, 51, 17) Then
    MainOp = 6: SubOp = 4: Exit Sub
  End If
  'Surface
  If IsInsideBox(X, Y1, 0, 0, Central.ScaleWidth, Central.PicCentral.Top - 21) Then
    MainOp = 7: SubOp = 1: Exit Sub
  End If
  'TEALTECH
  If IsInsideBox(X, Y, 6, 6, 99, 13) Then
    MainOp = 7: SubOp = 2: Exit Sub
  End If
  'Exit
  If IsInsideBox(X, Y, 173, 7, 46, 11) Then
    MainOp = 7: SubOp = 3: Exit Sub
  End If
  'Disk
  If IsInsideBox(X, Y, 233, 7, 14, 11) Then
    MainOp = 7: SubOp = 4: Exit Sub
  End If
  'GO
  If IsInsideBox(X, Y, 260, 6, 58, 19) Then
    MainOp = 7: SubOp = 5: Exit Sub
  End If
  'Disktest PRO
  If IsInsideBox(X, Y, 426, 6, 138, 13) Then
    MainOp = 7: SubOp = 6: Exit Sub
  End If
  'Disk Position
  If IsInsideBox(X, Y, 482, 31, 84, 44) Then
    MainOp = 7: SubOp = 7: Exit Sub
  End If
  'Disk Wave
  If IsInsideBox(X, Y, 482, 78, 84, 41) Then
    MainOp = 7: SubOp = 8: Exit Sub
  End If
  If IsInsideBox(X, Y, 450, 105, 32, 14) Then
    MainOp = 7: SubOp = 8: Exit Sub
  End If
  'Views
  If IsInsideBox(X, Y, 337, 7, 11, 11) Then  'Central View
    MainOp = 8: SubOp = 1: Exit Sub
  End If
  If IsInsideBox(X, Y, 353, 7, 11, 11) Then  'Surface View
    MainOp = 8: SubOp = 2: Exit Sub
  End If
  If IsInsideBox(X, Y, 371, 7, 11, 11) Then  'Small View
    MainOp = 8: SubOp = 3: Exit Sub
  End If
  If IsInsideBox(X, Y, 388, 7, 11, 11) Then  'Full View
    MainOp = 8: SubOp = 4: Exit Sub
  End If
  'Time
  If IsInsideBox(X, Y, 374, 35, 54, 51) Then 'Current
    MainOp = 9: SubOp = 1: Exit Sub
  End If
  If IsInsideBox(X, Y, 332, 34, 36, 18) Then 'Finish
    MainOp = 9: SubOp = 2: Exit Sub
  End If
  If IsInsideBox(X, Y, 332, 70, 36, 18) Then 'Elapsed
    MainOp = 9: SubOp = 3: Exit Sub
  End If
  If IsInsideBox(X, Y, 428, 34, 35, 18) Then 'Left
    MainOp = 9: SubOp = 4: Exit Sub
  End If
  If IsInsideBox(X, Y, 428, 70, 35, 18) Then 'Total
    MainOp = 9: SubOp = 5: Exit Sub
  End If
End Sub

'---------------------------------------------ControlAtMouse
Public Sub ControlAtMouse(ByRef MainOp As Byte, ByRef SubOp As Byte, ByVal X As Long, ByVal Y As Long)
  Select Case mOperation
    Case 1:  'Scan
      'Repair/Check/User
      If IsInsideBox(X, Y, 117, 43, 35, 8) Then   'Repair
        MainOp = 1: SubOp = 1: Exit Sub
      End If
      If IsInsideBox(X, Y, 117, 60, 35, 8) Then   'Check
        MainOp = 1: SubOp = 2: Exit Sub
      End If
      If IsInsideBox(X, Y, 117, 77, 35, 8) Then   'User
        MainOp = 1: SubOp = 3: Exit Sub
      End If
      'Read/Write/Verify
      If IsInsideBox(X, Y, 178, 51, 8, 8) Then   'Read
        MainOp = 1: SubOp = 4: Exit Sub
      End If
      If IsInsideBox(X, Y, 175, 65, 8, 8) Then   'Write
        MainOp = 1: SubOp = 5: Exit Sub
      End If
      If IsInsideBox(X, Y, 188, 61, 8, 8) Then   'Verify
        MainOp = 1: SubOp = 6: Exit Sub
      End If
      'Mark/Jump/Depth/Copy
      If IsInsideBox(X, Y, 214, 31, 32, 15) Then  'Mark
        MainOp = 1: SubOp = 7: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 46, 32, 15) Then  'Jump
        MainOp = 1: SubOp = 8: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 61, 32, 15) Then  'Depth
        MainOp = 1: SubOp = 9: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 76, 32, 15) Then  'Copy
        MainOp = 1: SubOp = 10: Exit Sub
      End If
    Case 2:  'Format
      'Full/Quick
      If IsInsideBox(X, Y, 169, 43, 35, 8) Then   'Full
        MainOp = 2: SubOp = 1: Exit Sub
      End If
      If IsInsideBox(X, Y, 169, 60, 35, 8) Then   'Quick
        MainOp = 2: SubOp = 2: Exit Sub
      End If
      'Mark/Jump
      If IsInsideBox(X, Y, 214, 31, 32, 15) Then  'Mark
        MainOp = 2: SubOp = 3: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 46, 32, 15) Then  'Jump
        MainOp = 2: SubOp = 4: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 61, 32, 15) Then  'Read
        MainOp = 2: SubOp = 5: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 76, 32, 15) Then  'Verify
        MainOp = 2: SubOp = 6: Exit Sub
      End If
    Case 3:  'Recover
      'Save/Load
      If IsInsideBox(X, Y, 170, 44, 35, 8) Then   'Save
        MainOp = 3: SubOp = 1: Exit Sub
      End If
      If IsInsideBox(X, Y, 170, 61, 35, 8) Then   'Load
        MainOp = 3: SubOp = 2: Exit Sub
      End If
      'Mark/Jump/Depth
      If IsInsideBox(X, Y, 214, 31, 32, 15) Then  'Mark
        MainOp = 3: SubOp = 3: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 46, 32, 15) Then  'Jump
        MainOp = 3: SubOp = 4: Exit Sub
      End If
      If IsInsideBox(X, Y, 214, 61, 32, 15) Then  'Depth
        MainOp = 3: SubOp = 5: Exit Sub
      End If
      'Up/Down
      If IsInsideBox(X, Y, 149, 65, 11, 6) Then   'Up
        MainOp = 3: SubOp = 6: Exit Sub
      End If
      If IsInsideBox(X, Y, 138, 65, 11, 6) Then   'Down
        MainOp = 3: SubOp = 7: Exit Sub
      End If
      'File
      If IsInsideBox(X, Y, 223, 78, 19, 10) Then  'File
        MainOp = 3: SubOp = 8: Exit Sub
      End If
    Case 4:  'Edit
      'Mark/Copy/Read/Write/Verify
      If IsInsideBox(X, Y, 114, 39, 30, 8) Then  'Mark
        MainOp = 4: SubOp = 1: Exit Sub
      End If
      If IsInsideBox(X, Y, 114, 49, 30, 8) Then  'Copy
        MainOp = 4: SubOp = 2: Exit Sub
      End If
      If IsInsideBox(X, Y, 114, 59, 30, 8) Then  'Read
        MainOp = 4: SubOp = 3: Exit Sub
      End If
      If IsInsideBox(X, Y, 114, 69, 30, 8) Then  'Write
        MainOp = 4: SubOp = 4: Exit Sub
      End If
      If IsInsideBox(X, Y, 114, 79, 30, 8) Then  'Verify
        MainOp = 4: SubOp = 5: Exit Sub
      End If
      'Format/Overwrite/Mark/Unmark
      If IsInsideBox(X, Y, 219, 31, 27, 15) Then  'Format
        MainOp = 4: SubOp = 6: Exit Sub
      End If
      If IsInsideBox(X, Y, 219, 46, 27, 15) Then  'Overwrite
        MainOp = 4: SubOp = 7: Exit Sub
      End If
      If IsInsideBox(X, Y, 219, 61, 27, 15) Then  'Mark
        MainOp = 4: SubOp = 8: Exit Sub
      End If
      If IsInsideBox(X, Y, 219, 76, 27, 15) Then  'Unmark
        MainOp = 4: SubOp = 9: Exit Sub
      End If
  End Select
  'Read N slider
  If IsInsideBox(X, Y, 19, 104, 89, 13) Then   'n read
    MainOp = 5
    SubOp = 4 - ((X - 19) * 4) \ 89
    Exit Sub
  End If
  'Main actions
  If IsInsideBox(X, Y, 264, 28, 51, 17) Then
    MainOp = 6: SubOp = 1: Exit Sub
  End If
  If IsInsideBox(X, Y, 264, 46, 51, 17) Then
    MainOp = 6: SubOp = 2: Exit Sub
  End If
  If IsInsideBox(X, Y, 264, 64, 51, 17) Then
    MainOp = 6: SubOp = 3: Exit Sub
  End If
  If IsInsideBox(X, Y, 264, 82, 51, 17) Then
    MainOp = 6: SubOp = 4: Exit Sub
  End If
End Sub

'-----------------------------------------StillInsideControl
Public Function StillInsideControl(ByRef pos As Point, ByVal MainOp As Byte, ByVal SubOp As Byte) As Boolean
  Dim novoMain As Byte
  Dim novoSub As Byte
  
  Call ControlAtMouse(novoMain, novoSub, pos.X, pos.Y - Central.PicCentral.Top + 21)
  If (novoMain <> MainOp) Or (novoSub <> SubOp) Then
    StillInsideControl = False
  Else
    StillInsideControl = True
  End If
End Function

'------------------------------------------------VerifyCheck
Public Sub VerifyCheck(ByVal Op As Byte)
  Dim i As Byte
  
  mLightScan(Op) = Not (mLightScan(Op))
  'testa combinações ilegais
  If (mLightScan(4) = False) And (mLightScan(5) = False) And (mLightScan(6) = False) Then
    mLightScan(4) = True
  End If
  If (mLightScan(4) = False) And (mLightScan(5) = True) And (mLightScan(6) = False) Then
    mLightScan(4) = True
  End If
  If (mLightScan(4) = False) And (mLightScan(5) = True) And (mLightScan(6) = True) Then
    mLightScan(4) = True
  End If
End Sub

'---------------------------------------------VerifyControls
Public Sub VerifyControls()
  If (mWork = 2) And (mLightRead <> 1) Then
    mLightRead = 1
    Call DisplayReadSlider
  End If
  If (mModWin = 1) Or (mModWin = 3) Then Exit Sub 'Large surface
  mLightRead = 1  '18 sectors
  mLightScan(9) = False
  mLightRecover(5) = False
End Sub

'-------------------------------------------ReDisplayCentral
Public Sub ReDisplayCentral()
  Call DisplayCentralSurface(mModWin)
  Call DisplayCentral(mModWin)
  Call DisplayCentralOp
  Call ReDisplayTool
  Call DisplayReadSlider
  Call DisplayGeralOp
  Call DisplaySurface
  Central.StartEnd.DrawCursor
  DoEvents
End Sub

'-----------------------------------------CentralScanOpCheck
Public Sub CentralScanOpCheck(ByVal SubOp As Byte)
  Dim i As Byte
  
  Select Case SubOp
    Case 1: 'Repair
      If mLightScan(3) Then
        For i = 1 To 10: mUserOp(i) = mLightScan(i): Next i
      End If
      mLightScan(4) = True: mLightScan(5) = False: mLightScan(6) = False
      mLightScan(7) = True:  mLightScan(8) = True
      mLightScan(9) = False: mLightScan(10) = True
      Call DisplayScanOp
    Case 2: 'Check
      If mLightScan(3) Then
        For i = 1 To 10: mUserOp(i) = mLightScan(i): Next i
      End If
      mLightScan(4) = True: mLightScan(5) = True: mLightScan(6) = True
      mLightScan(7) = False: mLightScan(8) = False
      mLightScan(9) = True:  mLightScan(10) = True
      Call DisplayScanOp
    Case 3:  'user
      If mLightScan(3) Then
        For i = 1 To 10: mUserOp(i) = mLightScan(i): Next i
      End If
      For i = 1 To 10: mLightScan(i) = mUserOp(i): Next i
      Call DisplayScanOp
  End Select
End Sub

'----------------------------------------CentralSurfaceScan
Public Sub CentralSurfaceScan(ByVal modo As Byte)
  If mWork = 0 Then
    If modo = 1 Then 'just click and change
      mOperation = 1
      Call DisplayGeralOp
      Call DisplayCentralOp
      Call DisplayScanOp
      DoEvents
    Else
      mWork = 1
      Call DisplayGoText
      If PrepareDisk() = True Then
        Call ReDisplayTool
        Call SurfaceScan
      End If
      mWork = 0
      Call DisplayGoText
    End If
  End If
End Sub

'-----------------------------------------CentralFormatDisk
Public Sub CentralFormatDisk(ByVal modo As Byte)
  If mWork = 0 Then
    If modo = 1 Then 'just click and change
      mOperation = 2
      Call DisplayGeralOp
      Call DisplayCentralOp
      Call DisplayFormatOp
      DoEvents
    Else
      mWork = 2
      Call DisplayGoText
      If mLightFormat(1) = True Then
        Call FormatFullDisk
        Call ReloadDisk
      Else
        Central.StartEnd.StartPosition = 1
        Call TestDiskChange
        Call FormatQuickDisk
        Call ReloadDisk
      End If
      mWork = 0
      Call DisplayGoText
    End If
  End If
End Sub

'----------------------------------------CentralRecoverDisk
Public Sub CentralRecoverDisk(ByVal modo As Byte)
  If mWork = 0 Then
    If modo = 1 Then 'just click and change
      mOperation = 3
      Call DisplayGeralOp
      Call DisplayCentralOp
      Call DisplayRecoverOp
      DoEvents
    Else
      mWork = 3
      Call DisplayGoText
      Central.StartEnd.StartPosition = 1
      If PrepareDisk() = True Then
        Call ReDisplayTool
        If mLightRecover(1) = True Then
          Call RecoverSaveDisk
        Else
          Call RecoverLoadDisk
          Call ReloadDisk
        End If
      Else
        Call RecoverLoadDisk
        Call ReloadDisk
      End If
      mWork = 0
      Call DisplayGoText
    End If
  End If
End Sub

'-------------------------------------------CentralEditMode
Public Sub CentralEditMode(ByVal modo As Byte)
  If mWork = 0 Then
    If modo = 1 Then 'just click and change
      mOperation = 4
      Call DisplayGeralOp
      Call DisplayCentralOp
      Call DisplayEditOp
      DoEvents
    Else
      mWork = 4
      Call DisplayGoText
      If PrepareDisk() = True Then
        Call ReDisplayTool
        Call EditDisk(eoInit)
      End If
    End If
  End If
End Sub

