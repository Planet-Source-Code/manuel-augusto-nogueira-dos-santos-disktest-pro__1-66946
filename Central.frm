VERSION 5.00
Begin VB.Form Central 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003E3E00&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6075
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Central.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   572
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer ToolTipTimer 
      Interval        =   500
      Left            =   2250
      Top             =   360
   End
   Begin VB.PictureBox FocusCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   1800
      ScaleHeight     =   105
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   1.50000e5
      Width           =   195
   End
   Begin VB.Timer EditTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1710
      Top             =   360
   End
   Begin VB.PictureBox Numbers 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   4950
      Picture         =   "Central.frx":0E42
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   600
   End
   Begin VB.PictureBox Letters 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   1
      Left            =   3420
      Picture         =   "Central.frx":10DC
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   296
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   4440
   End
   Begin VB.PictureBox Letters 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   0
      Left            =   3420
      Picture         =   "Central.frx":2276
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   295
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   4425
   End
   Begin VB.Timer GoTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   360
   End
   Begin VB.PictureBox CentralPics 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   4560
      Picture         =   "Central.frx":3E78
      ScaleHeight     =   195
      ScaleWidth      =   375
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   375
   End
   Begin VB.PictureBox CentralPics 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   3960
      Picture         =   "Central.frx":5F06
      ScaleHeight     =   195
      ScaleWidth      =   375
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   375
   End
   Begin VB.PictureBox CentralPics 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   3360
      Picture         =   "Central.frx":7D13
      ScaleHeight     =   195
      ScaleWidth      =   375
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   375
   End
   Begin VB.PictureBox CentralPics 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   4
      Left            =   2760
      Picture         =   "Central.frx":9BB7
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   1.50000e5
      Width           =   1155
   End
   Begin DiskTestPro.DragPos StartEnd 
      Height          =   135
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Start & End Position"
      Top             =   120
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   238
      Positions       =   80
      EndPosition     =   80
      Picture         =   "Central.frx":BF6D
      ForeColor       =   255
   End
   Begin DiskTestPro.TimedWave TimedWave1 
      Height          =   375
      Left            =   7260
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BackColor       =   2237442
      DrawWidth       =   2
      ForeColor       =   10156628
   End
   Begin VB.PictureBox CentralPics 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   2460
      Picture         =   "Central.frx":D7B7
      ScaleHeight     =   195
      ScaleWidth      =   375
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   375
   End
   Begin VB.PictureBox CentralPics 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   1920
      Picture         =   "Central.frx":E4B6
      ScaleHeight     =   195
      ScaleWidth      =   375
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   375
   End
   Begin VB.PictureBox CentralPics 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   6075
      Index           =   1
      Left            =   -750
      Picture         =   "Central.frx":10F39
      ScaleHeight     =   6075
      ScaleWidth      =   8580
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   8580
   End
   Begin VB.PictureBox CentralPics 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   960
      Picture         =   "Central.frx":11A38
      ScaleHeight     =   495
      ScaleWidth      =   1335
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   1335
   End
   Begin VB.Timer Tempo 
      Interval        =   1000
      Left            =   720
      Top             =   360
   End
   Begin VB.Timer Crono 
      Left            =   240
      Top             =   360
   End
   Begin VB.Image HelpDTP 
      Height          =   195
      Left            =   6420
      MouseIcon       =   "Central.frx":19D18
      MousePointer    =   99  'Custom
      Top             =   4350
      Width           =   2055
   End
   Begin VB.Image TealTech 
      Height          =   195
      Left            =   120
      MouseIcon       =   "Central.frx":1A15A
      MousePointer    =   99  'Custom
      Top             =   4350
      Width           =   1455
   End
   Begin VB.Image PicGO 
      Height          =   255
      Left            =   3930
      Top             =   4365
      Width           =   825
   End
   Begin VB.Image PicDisk 
      Height          =   195
      Left            =   3495
      Top             =   4350
      Width           =   210
   End
   Begin VB.Image EndCursor 
      Height          =   225
      Left            =   540
      MouseIcon       =   "Central.frx":1A464
      MousePointer    =   99  'Custom
      ToolTipText     =   "Test end track"
      Top             =   60
      Width           =   240
   End
   Begin VB.Image StartCursor 
      Height          =   225
      Left            =   240
      MouseIcon       =   "Central.frx":1A5B6
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Test start track "
      Top             =   60
      Width           =   240
   End
   Begin VB.Image PicWindow 
      Height          =   165
      Index           =   3
      Left            =   5820
      Top             =   4365
      Width           =   165
   End
   Begin VB.Image PicWindow 
      Height          =   165
      Index           =   2
      Left            =   5565
      Top             =   4365
      Width           =   165
   End
   Begin VB.Image PicWindow 
      Height          =   165
      Index           =   1
      Left            =   5310
      Top             =   4365
      Width           =   165
   End
   Begin VB.Image PicWindow 
      Height          =   165
      Index           =   0
      Left            =   5055
      Top             =   4365
      Width           =   165
   End
   Begin VB.Image PicClose 
      Height          =   165
      Left            =   2595
      Top             =   4365
      Width           =   705
   End
   Begin VB.Image PicCentral 
      Height          =   1470
      Left            =   45
      Top             =   4575
      Width           =   8490
   End
End
Attribute VB_Name = "Central"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+--------------------------------------------------------+
'| (c) Manuel Augusto N. dos Santos - July 2000           |
'+--------------------------------------------------------+
Option Explicit
'--------------------------------------Windows API Functions
Private Declare Sub GetCursorPos Lib "user32" (lpPoint As Point)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Point)
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseWinCapture Lib "user32" Alias "ReleaseCapture" () As Long
Private Declare Function SendWinMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function WinHelp Lib "user32.dll" Alias "WinHelpA" (ByVal hWndMain As Long, ByVal lpHelpFile As String, ByVal uCommand As Long, dwData As Any) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'----------------------------------------------------Private
Private mInterval As Long    'intervalo (ms) para crono timer
'------------------------------------------------------Const
Private Const WM_MOVE = &HF012
Private Const WM_SYSCOMMAND = &H112
Private Const SRCCOPY = &HCC0020
Private Const HELP_CONTENTS = &H3&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'---------------------------------------------Control Events
Private Sub EditTimer_Timer()
  If mWork = 4 Then
    Call EditDisk(eoReading)
  Else
    EditTimer.Enabled = False
  End If
End Sub

Private Sub FocusCenter_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim aux As Long
    
  If mWork <> 4 Then Exit Sub
  aux = SectorStat(SectorNumber(EditTrack, EditSide, EditSector))
  Call DisplaySectors(EditTrack, EditSide, EditSector, aux)
  Select Case KeyCode
    Case vbKeyLeft:
      If EditTrack > 0 Then EditTrack = EditTrack - 1 Else EditTrack = 79
    Case vbKeyRight:
      If EditTrack < 79 Then EditTrack = EditTrack + 1 Else EditTrack = 0
    Case vbKeyUp:
      aux = NumSectors()
      If EditSector > 1 Then
        EditSector = EditSector - aux
      Else
        If EditSide = 0 Then EditSide = 1 Else EditSide = 0
        EditSector = 19 - aux
      End If
    Case vbKeyDown:
      aux = NumSectors()
      If EditSector + aux <= 18 Then
        EditSector = EditSector + aux
      Else
        If EditSide = 0 Then EditSide = 1 Else EditSide = 0
        EditSector = 1
      End If
  End Select
  Call EditDisk(eoMove)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Mpos As Point
  Dim Wpos As Point
  Dim res As Long
  
  res = ReleaseWinCapture()
  'res = SendWinMessage(Me.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)  'does not work on NT
  res = SendWinMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  FocusCenter.SetFocus
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  CronoMode = tmNothing
  Crono.Enabled = False
  Crono.Interval = 0
  FocusCenter.SetFocus
End Sub

Private Sub Crono_Timer()
  Call InTimeControl
End Sub

Private Sub Form_Load()
  Dim i As Long
  Dim MyStr As String
 
  'Common defaults
  Crono.Enabled = False
  CronoMode = tmNothing
  mWork = 0
  mSaveName = "DTPRO.SAV"
  StartEnd.SetForm Central, StartEnd.Left, StartEnd.Top
  mJumpNext = 200
  mOperation = 1
  mLightRead = 1
  ToolTips = True
  Editting = False
  PosGO = 5
  MouseGO = False
  For i = 1 To 2880
    SectorStat(i) = statNormal
    SectorInfo(i) = IOempty
    SecCopy(i) = False
  Next i
  'defaults - Scan
  mLightScan(1) = False: mLightScan(2) = False: mLightScan(3) = True
  mLightScan(4) = True:  mLightScan(5) = False: mLightScan(6) = False
  mLightScan(7) = False: mLightScan(8) = True
  mLightScan(9) = False: mLightScan(10) = False
  For i = 1 To 10: mUserOp(i) = mLightScan(i): Next i
  'defaults - Format
  mLightFormat(1) = True: mLightFormat(2) = False
  mLightFormat(3) = True: mLightFormat(4) = True
  'defaults - Recover
  mLightRecover(1) = True:  mLightRecover(2) = False
  mLightRecover(3) = False: mLightRecover(4) = True: mLightRecover(5) = True
  'defaults - Edit
  mLightEdit(1) = False:  mLightEdit(2) = False
  mLightEdit(3) = True:   mLightEdit(4) = False
  mLightEdit(5) = False
  mLightEdit(6) = False:  mLightEdit(7) = False
  mLightEdit(8) = False:  mLightEdit(9) = False
  'Display all
  Call PicWindow_Click(3)  'Full View
  'Set focus
  MyStr = String(20, Chr$(0))
  MyStr = " Disktest PRO"
  SetWindowText Me.hWnd, MyStr
  'verifying Windows Version
  IsWinNT = False
  If GetWindowsVersion() = 2 Then IsWinNT = True
End Sub

Private Sub GoTimer_Timer()
  Dim Pic As StdPicture
  
  If MouseGO Then
    If PosGO > 0 Then
      PosGO = PosGO - 1
      If (PosGO > 0) And (mWork > 0) Then PosGO = 0
      Set Pic = LoadResPicture(206 - PosGO, vbResBitmap)
      Central.PaintPicture Pic, PicGO.Left, PicGO.Top, 57, 17, 0, 0, 57, 17, vbSrcCopy
      If PosGO = 0 Then Call DisplayGoText
    End If
  Else
    If PosGO < 5 Then
      PosGO = PosGO + 1
      If (PosGO < 5) And (mWork > 0) Then PosGO = 5
      Set Pic = LoadResPicture(206 - PosGO, vbResBitmap)
      Central.PaintPicture Pic, PicGO.Left, PicGO.Top, 57, 17, 0, 0, 57, 17, vbSrcCopy
    Else
      GoTimer.Enabled = False
    End If
  End If
End Sub

Private Sub HelpDTP_Click()
  WinHelp Me.hWnd, "DTPRO.HLP", HELP_CONTENTS, ByVal 0
End Sub

Private Sub PicDisk_Click()
  Dim i As Long
  
  DoEvents
  If mWork > 0 Then Exit Sub
  If TestDiskReady = True Then
    Call PrepareDisk
    For i = 1 To 2880
      SecCopy(i) = False
    Next i
    Call ReDisplayTool
  End If
End Sub

Private Sub PicDisk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If CronoMode = tmNothing Then
    CronoMode = tmOverDisk
    Crono.Interval = defInterval
    Crono.Enabled = True
  End If
End Sub

Private Sub PicGO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If CronoMode = tmNothing Then
    CronoMode = tmOverGO
    Crono.Interval = defInterval
    Crono.Enabled = True
    If PosGO > 0 Then
      GoTimer.Enabled = True
      MouseGO = True
    End If
  End If
End Sub

Private Sub PicGO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If mWork = 0 Then
    If PosGO = 0 Then
      Select Case mOperation
        Case 1: Call CentralSurfaceScan(2)
        Case 2: Call CentralFormatDisk(2)
        Case 3: Call CentralRecoverDisk(2)
        Case 4: Call CentralEditMode(2)
      End Select
    End If
  Else
    If mWork = 4 Then Call EditDisk(eoEndEdit)
    mWork = 0
    Call DisplayGoText
  End If
End Sub

Private Sub StartCursor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StartEnd.MouseOp 1
End Sub

Private Sub EndCursor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StartEnd.MouseOp 2
End Sub

Private Sub StartCursor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StartEnd.MouseOp 0
End Sub

Private Sub EndCursor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StartEnd.MouseOp 0
End Sub

Private Sub PicCentral_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call StartControlAction(X, Y)
End Sub

Private Sub PicCentral_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call EndControlAction(X, Y)
End Sub

Private Sub PicClose_Click()
  If MarkBad = True Then
    Call WriteDiskDATA
    MarkBad = False
  End If
  Call CloseDiskIO
  Call DiskSystemReset
  End
End Sub

Private Sub PicClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If CronoMode = tmNothing Then
    CronoMode = tmOverExit
    Crono.Interval = defInterval
    Crono.Enabled = True
  End If
End Sub

Private Sub PicWindow_Click(Index As Integer)
  Dim i As Integer
  Dim wndReg As Long
  Dim wndRet As Long
  
  mModWin = Index
  'Set Shaped form
  wndReg = RegionFromMask(CentralPics(1), mModWin, RGB(255, 255, 255))
  wndRet = SetWindowRgn(Me.hWnd, wndReg, True)
  Select Case Index
    Case 0:  'Central View
      Central.Height = 154 * Screen.TwipsPerPixelY
      PicClose.Top = 39
      PicWindow(0).Top = 39: PicWindow(1).Top = 39
      PicWindow(2).Top = 39: PicWindow(3).Top = 39
      PicDisk.Top = 38
      TealTech.Top = 38
      HelpDTP.Top = 38
      PicCentral.Top = 53
      PicGO.Top = 39
    Case 1:  'Surface View
      Central.Height = 311 * Screen.TwipsPerPixelY
      PicClose.Top = 291
      PicWindow(0).Top = 291: PicWindow(1).Top = 291
      PicWindow(2).Top = 291: PicWindow(3).Top = 291
      PicDisk.Top = 290
      TealTech.Top = 290
      HelpDTP.Top = 290
      PicCentral.Top = 10000
      PicGO.Top = 291
    Case 2:  'Small View
      Central.Height = 59 * Screen.TwipsPerPixelY
      PicClose.Top = 39
      PicWindow(0).Top = 39: PicWindow(1).Top = 39
      PicWindow(2).Top = 39: PicWindow(3).Top = 39
      PicDisk.Top = 38
      TealTech.Top = 38
      HelpDTP.Top = 38
      PicCentral.Top = 10000
      PicGO.Top = 39
    Case 3:  'Full View
      Central.Height = 406 * Screen.TwipsPerPixelY
      PicClose.Top = 291
      PicWindow(0).Top = 291: PicWindow(1).Top = 291
      PicWindow(2).Top = 291: PicWindow(3).Top = 291
      PicDisk.Top = 290
      TealTech.Top = 290
      HelpDTP.Top = 290
      PicCentral.Top = 305
      PicGO.Top = 291
  End Select
  HelpDTP.Left = 428
  PicCentral.Left = 3
  PicClose.Left = 173
  PicDisk.Left = 233
  PicGO.Left = 262
  PicWindow(0).Left = 337: PicWindow(1).Left = 354
  PicWindow(2).Left = 371: PicWindow(3).Left = 388
  StartEnd.Left = 4: StartEnd.Top = 8
  TealTech.Left = 8
  Call VerifyControls
  Call ReDisplayCentral
End Sub

Private Sub PicWindow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If CronoMode = tmNothing Then
    CronoMode = tmOverChgWin1 + Index
    Crono.Interval = defInterval
    Crono.Enabled = True
  End If
End Sub

Private Sub TealTech_Click()
  MousePointer = 0
  ToolTips = False
  DoEvents
  About.Show vbModal, Me
  Unload About
  ToolTips = True
End Sub

Private Sub Tempo_Timer()
  Dim NovoNow As Long
  
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  NovoNow = CalcNowSeconds(H24)
  If NovoNow <> oldNow Then
    If oldNow >= 0 Then
      Call Ponteiros(Central, oldNow, 0, 398, 60 + Central.PicCentral.Top - 21, 15)
    End If
    Call Ponteiros(Central, NovoNow, 1, 398, 60 + Central.PicCentral.Top - 21, 15)
    oldNow = NovoNow
  End If
End Sub

Private Sub TimedWave1_Added()
  Select Case mModWin
    Case 1, 2: Exit Sub
    Case 0: BitBlt Central.hDC, 484, 112, TimedWave1.Width, TimedWave1.Height, TimedWave1.hDC, 0, 0, SRCCOPY
    Case 3: BitBlt Central.hDC, 484, 364, TimedWave1.Width, TimedWave1.Height, TimedWave1.hDC, 0, 0, SRCCOPY
  End Select
End Sub

Private Sub ToolTipTimer_Timer()
  Dim tMain As Byte, tSub As Byte
  Dim Mpos As Point, Fpos As Point
  
  Call GetCursorPos(Mpos)
  Call GetFormCursorPos(Mpos, Central.Left, Central.Top, Fpos)
  Call ToolTipAtMouse(tMain, tSub, Fpos.X, Fpos.Y - Central.PicCentral.Top + 21)
  Call DisplayToolTip(tMain, tSub)
End Sub
