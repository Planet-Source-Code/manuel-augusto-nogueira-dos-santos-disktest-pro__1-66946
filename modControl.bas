Attribute VB_Name = "modControl"
Option Explicit

'------------------------------------------------Windows API
Private Declare Sub GetCursorPos Lib "user32" (lpPoint As Point)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Point)
'------------------------------------------Private Variables
Private mPicClose As Boolean          'EXIT brilha?
Private mPicDisk As Boolean           'Disk brilha?
Private mPicWindow(0 To 3) As Boolean 'PicWindow brilha?
Private mMainOp As Byte               'numero da operação
Private mSubOp As Byte                'numero da sub-operação
Private mFastB As Long                'ticks para up/down
Private mFastN As Long                'jump para up/down
'-------------------------------------------Public Variables
Public CronoMode As TimedModeConst
'-----------------------------------------------Public Const
Public Const defInterval = 20
'-----------------------------------------------Public Enums
Public Enum TimedModeConst
  tmNothing = 0
  tmDragForm = 1
  tmOverExit = 2
  tmOverChgWin1 = 3
  tmOverChgWin2 = 4
  tmOverChgWin3 = 5
  tmOverChgWin4 = 6
  tmCtrlButton = 7
  tmUpDown = 8
  tmOverDisk = 9
  tmOverGO = 10
End Enum

'----------------------------------------------InTimeControl
Public Sub InTimeControl()
  Dim Mpos As Point
  Dim Fpos As Point
  
  Select Case CronoMode
    '------------------------Nothing
    Case tmNothing:
      Central.Crono.Enabled = False
      Central.Crono.Interval = 0
    '------------------------Over Exit
    Case tmOverExit:
      Call GetCursorPos(Mpos)
      Call GetFormCursorPos(Mpos, Central.Left, Central.Top, Fpos)
      If IsInsideImage(Fpos, Central.PicClose) Then
        If mPicClose = False Then Call BrilhoPicClose(setON)
        mPicClose = True
      Else
        Call BrilhoPicClose(setOFF)
        Central.Crono.Enabled = False
        Central.Crono.Interval = 0
        mPicClose = False
        CronoMode = tmNothing
      End If
    '------------------------Over Change Window
    Case tmOverChgWin1, tmOverChgWin2, tmOverChgWin3, tmOverChgWin4:
      Call GetCursorPos(Mpos)
      Call GetFormCursorPos(Mpos, Central.Left, Central.Top, Fpos)
      If IsInsideImage(Fpos, Central.PicWindow(CronoMode - 3)) Then
        If mPicWindow(CronoMode - 3) = False Then Call BrilhoPicWindow(CronoMode - 3, setON)
        mPicWindow(CronoMode - 3) = True
      Else
        Call BrilhoPicWindow(CronoMode - 3, setOFF)
        Central.Crono.Enabled = False
        Central.Crono.Interval = 0
        mPicWindow(CronoMode - 3) = False
        CronoMode = tmNothing
      End If
    '------------------------Control Button
    Case tmCtrlButton:
        Call GetCursorPos(Mpos)
        Call GetFormCursorPos(Mpos, Central.Left, Central.Top, Fpos)
        If StillInsideControl(Fpos, mMainOp, mSubOp) = False Then
          If mMainOp = 6 Then
            If mSubOp <> mOperation Then Call ControlDown(2, mMainOp, mSubOp)
          Else
            Call ControlDown(4, mMainOp, mSubOp)
          End If
        Else
          If mMainOp = 6 Then
            If mSubOp <> mOperation Then Call ControlDown(1, mMainOp, mSubOp)
          Else
            Call ControlDown(3, mMainOp, mSubOp)
          End If
        End If
    '------------------------
    Case tmUpDown:
        Call GetCursorPos(Mpos)
        Call GetFormCursorPos(Mpos, Central.Left, Central.Top, Fpos)
        If mSubOp = 6 Then
          If mJumpNext < 9999 Then
             mJumpNext = mJumpNext + mFastN
             BkJump = BkJump + mFastN
          End If
          If mJumpNext > 9999 Then mJumpNext = 9999
          If BkJump > 9999 Then BkJump = 9999
        Else
          If mJumpNext > 1 Then
            mJumpNext = mJumpNext - mFastN
            BkJump = BkJump - mFastN
          End If
          If mJumpNext < 1 Then mJumpNext = 1
          If BkJump < 1 Then BkJump = 1
        End If
        Call DigitalINT(Central, 135, 54 + Central.PicCentral.Top - 21, mJumpNext, 2, 4)
        If StillInsideControl(Fpos, mMainOp, mSubOp) = False Then
          Call ControlDown(4, mMainOp, mSubOp)
          Central.Crono.Interval = 0
          Central.Crono.Enabled = False
          CronoMode = tmNothing
        End If
        mFastB = mFastB + 1
        Select Case mFastB
          Case 10: Central.Crono.Interval = 100
          Case 30: Central.Crono.Interval = 50
          Case 50: Central.Crono.Interval = 10
          Case 100: mFastN = 5
          Case 150: mFastN = 20
          Case 200: mFastN = 100
        End Select
    '------------------------Over Disk
    Case tmOverDisk:
      Call GetCursorPos(Mpos)
      Call GetFormCursorPos(Mpos, Central.Left, Central.Top, Fpos)
      If IsInsideImage(Fpos, Central.PicDisk) Then
        If mPicDisk = False Then Call BrilhoPicDisk(setON)
        mPicDisk = True
      Else
        Call BrilhoPicDisk(setOFF)
        Central.Crono.Enabled = False
        Central.Crono.Interval = 0
        mPicDisk = False
        CronoMode = tmNothing
      End If
    '------------------------Over GO button
    Case tmOverGO:
      Call GetCursorPos(Mpos)
      Call GetFormCursorPos(Mpos, Central.Left, Central.Top, Fpos)
      If IsInsideImage(Fpos, Central.PicGO) Then
        MouseGO = True
      Else
        MouseGO = False
        Central.Crono.Enabled = False
        Central.Crono.Interval = 0
        CronoMode = tmNothing
      End If
  End Select
End Sub

'-----------------------------------------StartControlAction
Public Sub StartControlAction(ByVal X As Long, ByVal Y As Long)
  Dim MainOp As Byte
  Dim SubOp As Byte
  Dim i As Byte
  
  MainOp = 0: SubOp = 0
  Call ControlAtMouse(MainOp, SubOp, _
       X \ Screen.TwipsPerPixelX + Central.PicCentral.Left, _
       Y \ Screen.TwipsPerPixelY + 21)
  If CronoMode <> tmNothing Then Exit Sub
  mMainOp = MainOp: mSubOp = SubOp
  Select Case MainOp
    Case 1: 'Scan
      Select Case SubOp
        Case 1, 2, 3: 'Repair/Check/User
          Call CentralScanOpCheck(SubOp)
          For i = 1 To 3: mLightScan(i) = False: Next i
          mLightScan(SubOp) = True
          Call DisplayScanOp
        Case 4, 5, 6: 'Read/Write/Verify
          Call VerifyCheck(SubOp)
          Call DisplayScanOp
          If mLightScan(3) = False Then
            For i = 1 To 10: mUserOp(i) = mLightScan(i): Next i
            Call CentralScanOpCheck(3) 'set user
            mLightScan(1) = False: mLightScan(2) = False: mLightScan(3) = True
            Call DisplayScanOp
          End If
        Case 7, 8, 9, 10: 'Mark/Jump/Depth/Copy
          CronoMode = tmCtrlButton
          Call ControlDown(3, MainOp, SubOp)
          Central.Crono.Interval = defInterval
          Central.Crono.Enabled = True
      End Select
    Case 2: 'Format
      Select Case SubOp
        Case 1, 2: 'Full/Quick
          If mWork = 0 Then
            mLightFormat(1) = False: mLightFormat(2) = False
            mLightFormat(SubOp) = True
            Call DisplayFormatOp
          End If
        Case 3, 4: 'Mark/Jump
          CronoMode = tmCtrlButton
          Call ControlDown(3, MainOp, SubOp)
          Central.Crono.Interval = defInterval
          Central.Crono.Enabled = True
      End Select
    Case 3: 'Recover
      Select Case SubOp
        Case 1, 2: 'Save/Load
          If mWork = 0 Then
            mLightRecover(1) = False: mLightRecover(2) = False
            mLightRecover(SubOp) = True
            Call DisplayRecoverOp
          End If
        Case 3, 4, 5, 8: 'Mark/Jump/Depth / File
          If (mWork = 0) Or (SubOp <> 8) Then
            CronoMode = tmCtrlButton
            Call ControlDown(3, MainOp, SubOp)
            Central.Crono.Interval = defInterval
            Central.Crono.Enabled = True
          End If
        Case 6, 7: 'Up/Down
          CronoMode = tmUpDown
          Call ControlDown(3, MainOp, SubOp)
          Central.Crono.Interval = 300
          Central.Crono.Enabled = True
          mFastB = 1
          mFastN = 1
          If mSubOp = 6 Then
            If mJumpNext < 9999 Then mJumpNext = mJumpNext + 1
            If BkJump < 9999 Then BkJump = BkJump + 1
          Else
            If mJumpNext > 1 Then mJumpNext = mJumpNext - 1
            If BkJump > 1 Then BkJump = BkJump - 1
          End If
          Call DigitalINT(Central, 135, 54 + Central.PicCentral.Top - 21, mJumpNext, 2, 4)
      End Select
    Case 4: 'Edit
      Select Case SubOp
        Case 1, 2, 3, 4, 5: 'Mark/Copy/Read/Write/Verify
          mLightEdit(SubOp) = Not (mLightEdit(SubOp))
          Call DisplayEditOp
        Case 6, 7, 8, 9: 'Format/Overwrite/Mark/Unmark
          CronoMode = tmCtrlButton
          Call ControlDown(3, MainOp, SubOp)
          Central.Crono.Interval = defInterval
          Central.Crono.Enabled = True
      End Select
    Case 5: 'N Read
      If (mWork = 0) Or (mWork = 4) Then
        If mWork = 4 Then Call DisplaySectors(EditTrack, EditSide, EditSector, SectorStat(SectorNumber(EditTrack, EditSide, EditSector)))
        mLightRead = SubOp
        Call VerifyControls
        Call DisplayReadSlider
        If mWork = 4 Then
          Call EditDisk(eoResetPos)
          Call EditDisk(eoMove)
        End If
      End If
    Case 6: 'Main Action
      If (mWork = 0) Or (mWork = SubOp) Then
        CronoMode = tmCtrlButton
        Call ControlDown(1, MainOp, SubOp)
        Central.Crono.Interval = defInterval
        Central.Crono.Enabled = True
      End If
  End Select
End Sub

'-------------------------------------------EndControlAction
Public Sub EndControlAction(ByVal X As Long, ByVal Y As Long)
  Dim MainOp As Byte
  Dim SubOp As Byte
  Dim i As Byte
          
  Central.Crono.Interval = 0
  Central.Crono.Enabled = False
  CronoMode = tmNothing
  MainOp = 0: SubOp = 0
  Call ControlAtMouse(MainOp, SubOp, _
       X \ Screen.TwipsPerPixelX + Central.PicCentral.Left, _
       Y \ Screen.TwipsPerPixelY + 21)
  If (MainOp = mMainOp) And (SubOp = mSubOp) Then
    Select Case MainOp
      Case 1: 'Scan
        Select Case SubOp
          Case 7, 8, 9, 10: 'Mark/Jump/Depth/Copy
            Call ControlDown(4, MainOp, SubOp)
            mLightScan(SubOp) = Not mLightScan(SubOp)
            If mLightScan(3) = False Then
              For i = 1 To 10: mUserOp(i) = mLightScan(i): Next i
              Call CentralScanOpCheck(3) 'set user
              mLightScan(1) = False: mLightScan(2) = False: mLightScan(3) = True
            End If
            Call VerifyControls
            Call DisplayScanOp
        End Select
      Case 2: 'Format
        Select Case SubOp
          Case 3, 4: 'Mark/Jump
            Call ControlDown(4, MainOp, SubOp)
            mLightFormat(SubOp) = Not mLightFormat(SubOp)
            Call DisplayFormatOp
        End Select
      Case 3: 'Recover
        Select Case SubOp
          Case 3, 4, 5: 'Mark/Jump/Depth
            Call ControlDown(4, MainOp, SubOp)
            mLightRecover(SubOp) = Not mLightRecover(SubOp)
            Call VerifyControls
            Call DisplayRecoverOp
          Case 6, 7: 'Up/Down
            Call ControlDown(4, MainOp, SubOp)
          Case 8: 'File
            If mWork = 0 Then
              Call ControlDown(4, MainOp, SubOp)
              Call AskForRecoverFile
              Call DisplayRecoverOp
            End If
        End Select
      Case 4: 'Edit
        Select Case SubOp
          Case 6, 7, 8, 9: 'Format/Overwrite/Mark/Unmark
            Call ControlDown(4, MainOp, SubOp)
        End Select
        If mWork = 4 Then
          Select Case SubOp
            Case 6: Call EditDisk(eoFormat)
            Case 7: Call EditDisk(eoOverwrite)
            Case 8: Call EditDisk(eoMarkBad)
            Case 9: Call EditDisk(eoUnmark)
          End Select
        End If
     'Case 5: N Read
      Case 6: 'Main
        If mWork = 0 Then
          Select Case SubOp
            Case 1: Call CentralSurfaceScan(1)
            Case 2: Call CentralFormatDisk(1)
            Case 3: Call CentralRecoverDisk(1)
            Case 4: Call CentralEditMode(1)
          End Select
        End If
    End Select
  End If
End Sub

'------------------------------------------AskForRecoverFile
Private Sub AskForRecoverFile()
  Dim newName As String
  
  newName = InputBox("New File Name:", "Change File Name", mSaveName)
  If newName = "" Then Exit Sub
  mSaveName = UCase(newName)
  If Len(mSaveName) > 20 Then mSaveName = Mid(mSaveName, 1, 20)
End Sub
