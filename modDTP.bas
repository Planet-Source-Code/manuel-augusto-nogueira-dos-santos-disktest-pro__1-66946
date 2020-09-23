Attribute VB_Name = "modDTP"
Option Explicit

'--------------------------------------Windows API Functions
Private Declare Function GetTickCount Lib "kernel32" () As Long
'-------------------------------------------Public Variables
Public mWork As Byte
Public SecCopy(1 To 2880) As Boolean
Public EditTrack As Byte
Public EditSide As Byte
Public EditSector As Byte
Public MarkBad As Boolean
Public Editting As Boolean
Public BkJump As Long
'------------------------------------------Private Variables
Private oldNsec As Byte
Private ReadTick As Long
Private StartTick As Long
Private SecList(1 To 2880, 1 To 512) As Byte
Private EOFdisk As Boolean
Private DepthScan As Boolean
Private StartSec As Long, CurrentSec As Long, EndSec As Long
    
'------------------------------------------Public Structures
Public Enum EditOperation
  eoInit = 1
  eoFormat = 2
  eoOverwrite = 3
  eoMarkBad = 4
  eoUnmark = 5
  eoReading = 6
  eoMove = 7
  eoResetPos = 8
  eoEndEdit = 9
End Enum
    
'-----------------------------------------------CountSecCopy
Public Function CountSecCopy() As Long
  Dim res As Long
  Dim i As Long
  
  res = 0
  For i = 1 To 2880
    If SecCopy(i) = False Then res = res + 1
  Next i
  CountSecCopy = res
End Function

'-----------------------------------------------SectorNumber
Public Function SectorNumber(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte) As Long
  SectorNumber = Track * 36 + Side * 18 + Sector
End Function

'-------------------------------------------------NumSectors
Public Function NumSectors() As Byte
  Select Case mLightRead
    Case 1: NumSectors = 18
    Case 2: NumSectors = 9
    Case 3: NumSectors = 3
    Case 4: NumSectors = 1
  End Select
End Function

'-----------------------------------------GetSideTrackInside
Public Function GetSideTrackInside(ByVal Track As Byte, ByVal Side As Byte, Optional Sector As Byte = 0) As SectorType
  Dim num As Long
  Dim Info As SectorType
  Dim i As Byte
  Dim nSec As Long
  
  Info = 0
  If Sector = 0 Then
    For i = 1 To 18
      num = SectorNumber(Track, Side, i)
      If SectorInfo(num) > Info Then
        Info = SectorInfo(num)
      End If
    Next i
  Else
    nSec = NumSectors() - 1
    For i = Sector To Sector + nSec
      num = SectorNumber(Track, Side, i)
      If SectorInfo(num) > Info Then
        Info = SectorInfo(num)
      End If
    Next i
  End If
  GetSideTrackInside = Info
End Function

'-----------------------------------------GetSideTrackSector
Public Function GetSideTrackSector(ByVal Track As Byte, ByVal Side As Byte) As StatType
  Dim num As Long
  Dim Info As StatType
  Dim i As Byte
  
  Info = 0
  For i = 1 To 18
    num = SectorNumber(Track, Side, i)
    If SectorStat(num) > Info Then
      Info = SectorStat(num)
    End If
  Next i
  GetSideTrackSector = Info
End Function

'---------------------------------------------DisplaySurface
Public Sub DisplaySurface()
  Dim Track As Byte
  Dim Side As Byte
  Dim Sector As Byte
  Dim num As Long
  Dim InfoSEC As SectorType
  Dim InfoSTA As StatType
  
  Call DisplayCentralSurface(mModWin)
  Central.StartEnd.DrawCursor
  For Track = 0 To 79
    For Side = 0 To 1
      If (mModWin = 1) Or (mModWin = 3) Then
        For Sector = 1 To 18
          num = SectorNumber(Track, Side, Sector)
          If SectorInfo(num) <> IOempty Then
            Call DisplaySectorInside(SectorInfo(num), Track, Side, Sector)
          End If
          If SectorStat(num) <> statNormal Then
            Call DisplaySector(Track, Side, Sector, SectorStat(num))
          End If
        Next Sector
      Else
        InfoSEC = GetSideTrackInside(Track, Side)
        InfoSTA = GetSideTrackSector(Track, Side)
        num = SectorNumber(Track, 0, Side + 1)
        If InfoSEC <> IOempty Then
          Call DisplaySectorInside(InfoSEC, Track, 0, Side + 1)
        End If
        If InfoSTA <> statNormal Then
          Call DisplaySector(Track, 0, Side + 1, InfoSTA)
        End If
      End If
    Next Side
  Next Track
End Sub

'----------------------------------------DisplaySectorInside
Public Sub DisplaySectorInside(ByVal Info As SectorType, ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte)
  Dim X As Long
  Dim Y As Long
  
  X = 8 + Track * 7
  Y = 18 + Side * 130 + (Sector - 1) * 7
  Select Case Info
    Case IOempty: Central.ForeColor = RGB(4, 54, 52)
    Case IObad:   Central.ForeColor = RGB(252, 2, 84)
    Case IOboot:  Central.ForeColor = RGB(252, 250, 4)
    Case IOfat1:  Central.ForeColor = RGB(4, 166, 252)
    Case IOfat2:  Central.ForeColor = RGB(4, 166, 4)
    Case IOdir:   Central.ForeColor = RGB(164, 166, 164)
    Case IOdata:  Central.ForeColor = RGB(4, 2, 4)
    Case Else:    Central.ForeColor = RGB(255, 255, 255)
  End Select
  Central.FillColor = Central.ForeColor
  Central.FillStyle = 0
  Central.Line (X, Y)-(X + 3, Y + 3), , BF
End Sub

'---------------------------------------------DisplaySectors
Public Sub DisplaySector(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal modo As StatType)
  Dim X As Long
  Dim Y As Long
  
  X = 7 + Track * 7
  Y = 17 + Side * 130 + (Sector - 1) * 7
  Select Case modo
    Case statError:  Central.ForeColor = RGB(252, 250, 252)
    Case statOk:     Central.ForeColor = RGB(84, 86, 84)
    Case statRead:   Central.ForeColor = RGB(0, 150, 150)
    Case statWrite:  Central.ForeColor = RGB(150, 0, 0)
    Case statVerify: Central.ForeColor = RGB(150, 150, 0)
    Case statNormal: Central.ForeColor = RGB(4, 2, 4)
    Case statEdit:   Central.ForeColor = RGB(94, 128, 142)
  End Select
  Central.FillStyle = 1
  Central.Line (X, Y)-(X + 5, Y + 5), , B
End Sub

Public Sub DisplaySectors(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal modo As StatType)
  Dim i As Byte
  Dim soma As Byte
  
  soma = NumSectors() - 1
  If (mModWin = 1) Or (mModWin = 3) Then
    For i = Sector To Sector + soma
      Call DisplaySector(Track, Side, i, modo)
    Next i
  Else
    Call DisplaySector(Track, 0, Side + 1, GetSideTrackSector(Track, Side))
  End If
End Sub

'-------------------------------------------------ReloadDisk
Public Sub ReloadDisk()
  Dim i As Long
  
  Call DiskSystemReset
  For i = 1 To 2880
    SecCopy(i) = False
  Next i
  Call ReadDiskDATA
  Call DisplaySurface
End Sub

'------------------------------------------------PrepareDisk
Public Function PrepareDisk() As Boolean
  Dim i As Long
  
  If TestDiskChange = True Then
    For i = 1 To 2880
      SecCopy(i) = False
    Next i
  End If
  If TestDiskReady = True Then
    Call ReadDiskDATA
    Call DisplaySurface
    PrepareDisk = True
  End If
End Function

'----------------------------------------------------JumpBad
Public Function JumpBad(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal Light As Boolean) As Boolean
  Dim nSect As Byte
  Dim i As Byte
  Dim num As Long
  
  If Light = False Then
    JumpBad = False
    Exit Function
  End If
  nSect = NumSectors() - 1
  For i = Sector To Sector + nSect
    num = SectorNumber(Track, Side, i)
    If SectorInfo(num) <> IObad Then
      JumpBad = False
      Exit Function
    End If
  Next i
  JumpBad = True
End Function

'-----------------------------------------------JumpOnlyData
Public Function JumpOnlyData(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal Light As Boolean) As Boolean
  Dim nSect As Byte
  Dim i As Byte
  Dim num As Long
  
  If Light = False Then
    JumpOnlyData = False
    Exit Function
  End If
  nSect = NumSectors() - 1
  For i = Sector To Sector + nSect
    num = SectorNumber(Track, Side, i)
    If (SectorInfo(num) = IOdata) Or (SectorInfo(num) = IOboot) Or _
       (SectorInfo(num) = IOdir) Or (SectorInfo(num) = IOfat1) Or _
       (SectorInfo(num) = IOfat2) Then
      JumpOnlyData = False
      Exit Function
    End If
  Next i
  JumpOnlyData = True
End Function

'----------------------------------------------AdvanceSector
Public Function AdvanceSector(ByRef Track As Byte, ByRef Side As Byte, ByRef Sector As Byte) As Boolean
  AdvanceSector = True
  Sector = Sector + NumSectors()
  If NumSectors < 18 Then Central.TimedWave1.Add
  If Sector > 18 Then
    Sector = 1
    Side = Side + 1
  End If
  If Side = 2 Then
    Side = 0
    Track = Track + 1
    If NumSectors = 18 Then Central.TimedWave1.Add
  End If
  If (Track = 80) Or (Track = Central.StartEnd.EndPosition) Then
    AdvanceSector = False
    Track = Track - 1: Side = 1: Sector = 18 'set last sector
  End If
End Function

'--------------------------------------------DisplayPosition
Public Sub DisplayPosition(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte)
  Call DigitalINT(Central, 550, 37 + Central.PicCentral.Top - 21, Track, 2, 2)
  Call DigitalINT(Central, 538, 49 + Central.PicCentral.Top - 21, Side, 2, 1)
  Call DigitalINT(Central, 550, 49 + Central.PicCentral.Top - 21, Sector, 2, 2)
  Call DigitalINT(Central, 538, 61 + Central.PicCentral.Top - 21, SectorNumber(Track, Side, Sector), 2, 4)
End Sub

'----------------------------------------------DisplayTiming
Public Sub DisplayTiming()
  Dim EndTick As Long
  Dim TickLeft As Long
  Dim CurrentTick As Long
  Dim aux As Long
  Dim TempT As Double
  Const r1X = 337, r1Y = 43  'Ending
  Const r2X = 431, r2Y = 43  'Left
  Const r3X = 337, r3Y = 74  'Elapsed
  Const r4X = 431, r4Y = 74  'Predicted
   
  'calculate ticks
  CurrentTick = GetTickCount()
  If CurrentSec <> StartSec Then
    TempT = (CurrentTick - StartTick) / (CurrentSec - StartSec)
    If TempT > 30000 Then TempT = 30000 'overflow check
    TickLeft = (EndSec - CurrentSec) * TempT
  End If
  EndTick = CurrentTick + TickLeft
  'Ending time
  aux = CalcNowSeconds(H24) + (EndTick - CurrentTick) / 1000
  Call DigitalText(r1X, r1Y + Central.PicCentral.Top - 21, StrClock(aux), 3)
  'Time left in seconds
  aux = TickLeft / 1000
  Call DigitalText(r2X, r2Y + Central.PicCentral.Top - 21, StrClock(aux), 3)
  'Seconds elapsed
  aux = (CurrentTick - StartTick) / 1000
  Call DigitalText(r3X, r3Y + Central.PicCentral.Top - 21, StrClock(aux), 3)
  'Predicted time
  aux = (EndTick - StartTick) / 1000
  Call DigitalText(r4X, r4Y + Central.PicCentral.Top - 21, StrClock(aux), 3)
End Sub

'------------------------------------------------DepthScanIn
Private Function DepthScanIn(ByVal IOres As Long, ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal Light As Boolean) As Boolean
  If (Light = False) And (DepthScan = False) Then
    DepthScanIn = False
    Exit Function
  End If
  'no error, no depth going - nothing to do
  If (IOres = 0) And (DepthScan = False) Then
    DepthScanIn = False
    Exit Function
  End If
  'error - start depth scan or more depth scan
  If IOres <> 0 Then
    If mLightRead < 4 Then
      Call DisplaySectors(Track, Side, Sector, statNormal)
      mLightRead = mLightRead + 1
    Else
      DepthScanIn = False
      Exit Function
    End If
    Call DisplayReadSlider
    DepthScan = True
    DepthScanIn = True
    Exit Function
  End If
  'default
  DepthScanIn = False
End Function

'-----------------------------------------------DepthScanOut
Private Sub DepthScanOut(ByVal IOres As Long, ByVal Sector As Byte)
  If DepthScan = False Then oldNsec = mLightRead
  If (IOres = 0) And (DepthScan = True) Then
    Select Case Sector
      Case 1:
        mLightRead = 1
        DepthScan = False
      Case 10: mLightRead = 2
      Case 4, 7, 13, 16: mLightRead = 3
      Case Else: mLightRead = 4
    End Select
    Call DisplayReadSlider
  End If
  If (mLightRead < oldNsec) Or (EOFdisk = True) Then
    mLightRead = oldNsec
    DepthScan = False
    Call DisplayReadSlider
  End If
End Sub

'---------------------------------------------------AutoCopy
Private Sub AutoCopy(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal nSect As Byte)
  Dim SecNum As Long
  Dim i As Byte
  Dim j As Long
  Dim sKey As String
  
  For i = 1 To nSect
    SecNum = SectorNumber(Track, Side, Sector + (i - 1))
    If SecCopy(SecNum) = False Then
      For j = 1 To 512
        SecList(SecNum, j) = IOdados((i - 1) * 512 + j - 1)
      Next j
      SecCopy(SecNum) = True
    End If
  Next i
End Sub

'-------------------------------------------TransferAutoCopy
Private Sub TransferAutoCopy(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal nSect As Byte)
  Dim SecNum As Long
  Dim i As Byte
  Dim j As Long
  Dim sKey As String
  
  For i = 1 To nSect
    SecNum = SectorNumber(Track, Side, Sector + (i - 1))
    For j = 1 To 512
      IOdados((i - 1) * 512 + j - 1) = SecList(SecNum, j)
    Next j
  Next i
End Sub

'-----------------------------------------------TestAutoCopy
Private Function TestAutoCopy(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal nSect As Byte)
  Dim i As Byte
  Dim curSec As Long
  Dim res As Boolean
  
  res = True
  For i = Sector To Sector + nSect - 1
    curSec = SectorNumber(Track, Side, i)
    If SecCopy(curSec) = False Then res = False
  Next i
  TestAutoCopy = res
End Function

'-------------------------------------------ReplaceByCopyFAT
Private Function ReplaceByCopyFAT(ByVal nSect As Long) As Long
  Dim i As Long
  
  If (nSect < 2) Or (nSect > 19) Then
    ReplaceByCopyFAT = 1
    Exit Function
  End If
  Select Case nSect
    Case 2, 3, 4, 5, 6, 7, 8, 9, 10:          'Get from FAT2
      If SecCopy(nSect + 9) = False Then
        ReplaceByCopyFAT = 1
        Exit Function
      End If
      For i = 1 To 512
        SecList(nSect, i) = SecList(nSect + 9, i)
      Next i
      SecCopy(nSect) = True
    Case 11, 12, 13, 14, 15, 16, 17, 18, 19:  'Get from FAT1
      If SecCopy(nSect - 9) = False Then
        ReplaceByCopyFAT = 1
        Exit Function
      End If
      For i = 1 To 512
        SecList(nSect, i) = SecList(nSect - 9, i)
      Next i
      SecCopy(nSect) = True
  End Select
  ReplaceByCopyFAT = 0
End Function

'-----------------------------------------MarkBadReservation
Private Sub MarkBadReservation(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal nSect As Byte)
  Dim i As Byte
  Dim curSec As Long
  
  For i = Sector To Sector + nSect - 1
    curSec = SectorNumber(Track, Side, i)
    If SectorInfo(curSec) = IOempty Then
      SectorInfo(curSec) = IObad
      SectorVal(curSec) = &HFF7 'bad
      MarkBad = True
      Call DisplaySectorInside(IObad, Track, Side, i)
    End If
  Next i
End Sub

'---------------------------------------UnMarkBadReservation
Private Sub UnMarkBadReservation(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal nSect As Byte)
  Dim i As Byte
  Dim curSec As Long
  
  For i = Sector To Sector + nSect - 1
    curSec = SectorNumber(Track, Side, i)
    If SectorInfo(curSec) = IObad Then
      SectorInfo(curSec) = IOempty
      SectorVal(curSec) = 0
      MarkBad = True
      Call DisplaySectorInside(IOempty, Track, Side, i)
    End If
  Next i
End Sub

'--------------------------------------------SetSectorStatus
Public Sub SetSectorStatus(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal nSect As Byte, ByVal IOResult As Long)
  Dim i As Byte
  Dim curSec As Long

  For i = Sector To Sector + nSect - 1
    curSec = SectorNumber(Track, Side, i)
    If IOResult = 0 Then
      SectorStat(curSec) = statOk
      If (mModWin = 1) Or (mModWin = 3) Then Call DisplaySector(Track, Side, i, statOk)
    Else
      SectorStat(curSec) = statError
      If (mModWin = 1) Or (mModWin = 3) Then Call DisplaySector(Track, Side, i, statError)
    End If
  Next i
  If (mModWin = 0) Or (mModWin = 2) Then
    Call DisplaySector(Track, 0, Side + 1, GetSideTrackSector(Track, Side))
  End If
End Sub
      
'---------------------------------------------------InitScan
Private Sub InitScan()
  Dim i As Long
  
  EOFdisk = False
  MarkBad = False
  DepthScan = False
  For i = 1 To 2880: SectorStat(i) = statNormal: Next i
  Call DisplaySurface
  StartSec = SectorNumber(Central.StartEnd.StartPosition - 1, 0, 1)
  CurrentSec = StartSec
  EndSec = SectorNumber(Central.StartEnd.EndPosition - 1, 1, 18)
  Call DisplayTiming
  Central.TimedWave1.Clear
  Call InitializeDiskIO
  DoEvents
  StartTick = GetTickCount()
End Sub

'------------------------------------------------SurfaceScan
Public Sub SurfaceScan()
  Dim Track As Byte
  Dim Side As Byte
  Dim Sector As Byte
  Dim IOResult As Long
  Dim nSect As Byte
  
  'prepare
  Track = Central.StartEnd.StartPosition - 1
  Side = 0
  Sector = 1
  Call InitScan
  'Scan
  Do While (EOFdisk = False) And (mWork = 1)
    nSect = NumSectors()
    Call DisplayPosition(Track, Side, Sector)
    If JumpBad(Track, Side, Sector, mLightScan(8)) = False Then
      IOResult = 0
      '-----------------------------------------------------
      'read
      If mLightScan(4) = True Then
        Call DisplaySectors(Track, Side, Sector, statRead)
        DoEvents
        IOResult = DiskIO(IOReadDisk, IOFloppyA, nSect, Track, Side, Sector)
        If (IOResult = 0) And (mLightScan(10) = True) Then Call AutoCopy(Track, Side, Sector, nSect)
      End If
      'write
      If mLightScan(5) = True Then
        If (IOResult = 0) Or (GetSideTrackInside(Track, Side, Sector) = IOempty) Or (GetSideTrackInside(Track, Side, Sector) = IObad) Then
          Call DisplaySectors(Track, Side, Sector, statWrite)
          DoEvents
          IOResult = DiskIO(IOWriteDisk, IOFloppyA, nSect, Track, Side, Sector)
          If IOResult = 0 Then
            Call DisplaySectors(Track, Side, Sector, statRead)
            DoEvents
            IOResult = DiskIO(IOReadDisk, IOFloppyA, nSect, Track, Side, Sector)
          End If
        End If
      End If
      'verify
      If (IOResult = 0) And (mLightScan(6) = True) Then
        Call DisplaySectors(Track, Side, Sector, statVerify)
        DoEvents
        IOResult = DiskIO(IOVerifyDisk, IOFloppyA, nSect, Track, Side, Sector)
      End If
      '----------------------------------------------------
      'Depth Scan (IN)
      If DepthScanIn(IOResult, Track, Side, Sector, mLightScan(9)) = True Then GoTo ContinueScan
      'set sector status
      Call SetSectorStatus(Track, Side, Sector, nSect, IOResult)
      'Mark Bad reservation
      If (IOResult <> 0) And (mLightScan(7) = True) Then Call MarkBadReservation(Track, Side, Sector, nSect)
    End If
    DoEvents
    'next sector
    If AdvanceSector(Track, Side, Sector) = False Then EOFdisk = True
    'Depth Scan (OUT)
    If DepthScan = False Then oldNsec = mLightRead
    Call DepthScanOut(IOResult, Sector)
    'Check time
    CurrentSec = SectorNumber(Track, Side, Sector)
    Call DisplayTiming
ContinueScan:
  Loop
  'Save FAT if marked bad
  If MarkBad = True Then
    Call WriteDiskDATA
    MarkBad = False
  End If
  Call CloseDiskIO
End Sub

'--------------------------------------------RecoverSaveDisk
Public Sub RecoverSaveDisk()
  Dim Track As Byte
  Dim Side As Byte
  Dim Sector As Byte
  Dim IOResult As Long
  Dim nSect As Byte
  Dim i As Long
  Dim num As Long
  
  'prepare
  Track = Central.StartEnd.StartPosition - 1
  Side = 0
  Sector = 1
  BkJump = mJumpNext
  Call InitScan
  Call CreateIdFile(mSaveName, "DTPRO-Saved Disk Image", 30)
  'Save
  Do While (EOFdisk = False) And (mWork = 3)
    If (mModWin = 0) Or (mModWin = 3) Then
      Call DigitalINT(Central, 135, 41 + Central.PicCentral.Top - 21, CountSecCopy(), 2, 4)
    End If
    nSect = NumSectors()
    Call DisplayPosition(Track, Side, Sector)
    If (JumpBad(Track, Side, Sector, mLightRecover(4)) = False) And _
       (JumpOnlyData(Track, Side, Sector, mLightRecover(4)) = False) Then
      IOResult = 0
      '-----------------------------------------------------
      'test auto copy
      If TestAutoCopy(Track, Side, Sector, nSect) = False Then
        Call DisplaySectors(Track, Side, Sector, statRead)
        DoEvents
        IOResult = DiskIO(IOReadDisk, IOFloppyA, nSect, Track, Side, Sector)
        'test if FAT problem
        If (nSect = 1) And (IOResult <> 0) Then
          num = SectorNumber(Track, Side, Sector)
          IOResult = ReplaceByCopyFAT(num)
        End If
        If IOResult = 0 Then Call AutoCopy(Track, Side, Sector, nSect)
      End If
      'Jump next after n readings
      If (IOResult <> 0) And (mJumpNext > 0) Then
        If DepthScanIn(IOResult, Track, Side, Sector, mLightRecover(5)) = False Then
          mJumpNext = mJumpNext - 1
          Call DigitalINT(Central, 135, 54 + Central.PicCentral.Top - 21, mJumpNext, 2, 4)
        End If
        GoTo ContinueRecover
      Else
        mJumpNext = BkJump
        Call DigitalINT(Central, 135, 54 + Central.PicCentral.Top - 21, mJumpNext, 2, 4)
      End If
      'transfer to IO buffer
      If IOResult = 0 Then Call TransferAutoCopy(Track, Side, Sector, nSect)
      '-----------------------------------------------------
      'set sector status
      Call SetSectorStatus(Track, Side, Sector, nSect, IOResult)
      'Mark Bad reservation
      If (IOResult <> 0) And (mLightRecover(3) = True) Then Call MarkBadReservation(Track, Side, Sector, nSect)
    Else
      For i = 1 To 512 * nSect
        IOdados(i - 1) = 0
      Next i
      Call AutoCopy(Track, Side, Sector, nSect)
      Call SetSectorStatus(Track, Side, Sector, nSect, 0)
    End If
    DoEvents
    'save data
    Call WriteIOData(nSect)
    'next sector
    If AdvanceSector(Track, Side, Sector) = False Then EOFdisk = True
    'Depth Scan (OUT)
    If DepthScan = False Then oldNsec = mLightRead
    Call DepthScanOut(IOResult, Sector)
    'Check time
    CurrentSec = SectorNumber(Track, Side, Sector)
    Call DisplayTiming
ContinueRecover:
  Loop
  'Save FAT if marked bad
  If MarkBad = True Then
    Call WriteDiskDATA
    MarkBad = False
  End If
  Call CloseIdFile
  Call CloseDiskIO
  If (mModWin = 0) Or (mModWin = 3) Then
    Call DigitalINT(Central, 135, 41 + Central.PicCentral.Top - 21, CountSecCopy(), 2, 4)
  End If
End Sub

'--------------------------------------------RecoverLoadDisk
Public Sub RecoverLoadDisk()
  Dim CopyVal() As Long
  Dim i As Long
  Dim Track As Byte
  Dim Side As Byte
  Dim Sector As Byte
  Dim IOResult As Long
  Dim nSect As Long
  Dim toCopy As Boolean
  
  'Ensure no Data on destination
  For i = 34 To 2880
    If SectorInfo(i) = IOdata Then
      i = MsgBox("Loading image with data on floppy disk." & Chr(13) & Chr(10) & "OK to continue?", vbExclamation Or vbOKCancel, "Error")
      If i = vbCancel Then Exit Sub Else Exit For
    End If
  Next i
  'Ensure valid file
  i = OpenIdFile(mSaveName, "DTPRO-Saved Disk Image", 30)
  If i = -2 Then
    MsgBox "The file provided was not saved with DiskTest PRO.", vbExclamation Or vbOKOnly, "Error"
    Call CloseIdFile
    Exit Sub
  End If
  If i = -1 Then
    MsgBox "Can't read the file provided.", vbExclamation Or vbOKOnly, "Error"
    Call CloseIdFile
    Exit Sub
  End If
  'Ensure compatibility - filesize
  If isExpectedSize(31, 512, 0) = False Then
    MsgBox "The disk image size is not compatible with a floppy disk sector.", vbExclamation Or vbOKOnly, "Error"
    Call CloseIdFile
    Exit Sub
  End If
  'Ensure compatibility - end position
  If isExpectedSize(0, 0, 31 + 512 * 36 * Central.StartEnd.EndPosition) = False Then
    MsgBox "The disk image size is not compatible" & Chr(13) & Chr(10) & "with the END cursor position.", vbExclamation Or vbOKOnly, "Error"
    Call CloseIdFile
    Exit Sub
  End If
  'Ensure compatibility - Bad sectors
  CopyVal() = GetImageFAT()
  For i = 34 To 2880
    'Dest:Bad; Src:not Bad; "Only data" is off -> error
    'Dest:Bad; Src:data -> error
    If (SectorInfo(i) = IObad) And ((CopyVal(i) < &HFF0) Or (CopyVal(i) > &HFF7)) Then
      If (mLightRecover(4) = False) Or (CopyVal(i) > 0) Then
        MsgBox "The floppy disk has bad sectors where data should be. (Sector " & i & ")", vbExclamation Or vbOKOnly, "Error"
        Call CloseIdFile
        Exit Sub
      End If
    End If
  Next i
  'prepare
  Track = Central.StartEnd.StartPosition - 1
  Side = 0
  Sector = 1
  Call InitScan
  mLightRead = 1
  Call DisplayReadSlider
  'Load
  Do While (EOFdisk = False) And (mWork = 3)
    Call DisplayPosition(Track, Side, Sector)
    IOResult = 0
    toCopy = True
    'only data ON - empty or bad - jump
    If mLightRecover(4) = True Then
      For i = 1 To 18
        nSect = SectorNumber(Track, Side, Sector + i - 1)
        If (CopyVal(nSect) = 0) Or ((CopyVal(nSect) >= &HFF0) And (CopyVal(nSect) <= &HFF7)) Then
          If Track > 1 Then toCopy = False
        End If
      Next i
    End If
    'read from image, write to floppy
    If toCopy = True Then
      Call ReadIOData(18)
      Call DisplaySectors(Track, Side, Sector, statWrite)
      DoEvents
      IOResult = DiskIO(IOWriteDisk, IOFloppyA, 18, Track, Side, Sector)
      'check/alter FAT
      For i = 1 To 18
        nSect = SectorNumber(Track, Side, Sector + i - 1)
        If (mLightRecover(3) = False) And (SectorInfo(nSect) = IOempty) And (CopyVal(nSect) >= &HFF0) And (CopyVal(nSect) <= &HFF7) Then
         'do nothing = do not mark bad
        Else
          SectorVal(nSect) = CopyVal(nSect)
          Call DisplaySectorInside(GetSecType(Track, Side, i), Track, Side, i)
        End If
      Next i
    End If
    'set sector status
    Call SetSectorStatus(Track, Side, Sector, 18, IOResult)
    DoEvents
    'next sector
    If AdvanceSector(Track, Side, Sector) = False Then EOFdisk = True
    'Check time
    CurrentSec = SectorNumber(Track, Side, Sector)
    Call DisplayTiming
ContinueRecover:
  Loop
  Call WriteDiskDATA
  Call CloseIdFile
  Call CloseDiskIO
End Sub

'---------------------------------------------FormatFullDisk
Public Sub FormatFullDisk()
  Dim i As Long
  Dim Track As Byte
  Dim Side As Byte
  Dim IOResult As Long
  Dim Bad As Long, Good As Long, Avail As Long, Percent As Long
  Dim Y As Long
  
  If (TestDiskChange = True) Then Call ClearDiskData
  If TestDiskReady = False Then
    mWork = 0
    Exit Sub
  End If
  'prepare
  Track = Central.StartEnd.StartPosition - 1
  Side = 0
  Call InitScan
  mLightRead = 1
  Call DisplayReadSlider
  'Format Full
  Do While (EOFdisk = False) And (mWork = 2)
    Call DisplayPosition(Track, Side, 1)
    If JumpBad(Track, Side, 1, mLightFormat(4)) = False Then
      IOResult = 0
      Call DisplaySectors(Track, Side, 1, statWrite)
      DoEvents
      'format
      IOResult = FormatTrack(IOFloppyA, Track, Side, mLightFormat(3))
      'set sector status
      Call SetSectorStatus(Track, Side, 1, 18, IOResult)
      'Mark Bad reservation
      If (IOResult <> 0) And (mLightFormat(3) = True) Then Call MarkBadReservation(Track, Side, 1, 18)
    End If
    If (mModWin = 0) Or (mModWin = 3) Then
      Y = Central.PicCentral.Top - 21
      Call CountSectors(Bad, Good, Avail, Percent)
      Call DigitalINT(Central, 135, 42 + Y, Bad, 2, 4)
      Call DigitalINT(Central, 135, 57 + Y, Good, 2, 4)
      Call DigitalINT(Central, 169, 79 + Y, Avail, 2, 7)
      Call DigitalText(127, 79 + Y, Str03(Percent), 3)
    End If
    DoEvents
    'next sector
    If AdvanceSector(Track, Side, 1) = False Then EOFdisk = True
    'Check time
    CurrentSec = SectorNumber(Track, Side, 1)
    Call DisplayTiming
ContinueScan:
  Loop
  'Save FAT
  Call WriteDiskDATA
  MarkBad = False
  Call CloseDiskIO
  Call DiskSystemReset
End Sub

'--------------------------------------------FormatQuickDisk
Public Sub FormatQuickDisk()
  Dim i As Long, j As Long
  Dim IOResult As Long
  Dim Bad As Long, Good As Long, Avail As Long, Percent As Long
  Dim Y As Long
  
  Call InitScan
  Call WriteBootSector
  For i = 0 To 4607: IOdados(i) = 0: Next i
  For i = 34 To 2880
    If SectorInfo(i) <> IObad Then
      SectorInfo(i) = IOempty
      SectorVal(i) = 0
    End If
    If (SectorInfo(i) = IObad) And (mLightFormat(3) = False) Then
      SectorInfo(i) = IOempty
      SectorVal(i) = 0
    End If
  Next i
  Call WriteDiskDATA
  For i = 0 To 3583: IOdados(i) = 0: Next i
  Call DiskIO(IOWriteDisk, IOFloppyA, 7, 0, 1, 2)
  Call DiskIO(IOWriteDisk, IOFloppyA, 7, 0, 1, 9)
  Call CloseDiskIO
  If (mModWin = 1) Or (mModWin = 2) Then Exit Sub
  Y = Central.PicCentral.Top - 21
  Call CountSectors(Bad, Good, Avail, Percent)
  Call DigitalINT(Central, 135, 42 + Y, Bad, 2, 4)
  Call DigitalINT(Central, 135, 57 + Y, Good, 2, 4)
  Call DigitalINT(Central, 169, 79 + Y, Avail, 2, 7)
  Call DigitalText(127, 79 + Y, Str03(Percent), 3)
End Sub

'---------------------------------------------------EditDisk
Public Sub EditDisk(ByVal Operation As EditOperation)
  Dim i As Long
  Dim IOResult As Long
  Dim oldRead As Byte
  
  If Editting = True Then Exit Sub
  Editting = True
  Select Case Operation
    Case eoInit:
      For i = 1 To 2880: SectorStat(i) = statNormal: Next i
      Call DisplaySurface
      Central.TimedWave1.Clear
      Call InitializeDiskIO
      EditTrack = Central.StartEnd.StartPosition - 1
      EditSide = 0
      EditSector = 1
      Call DisplaySectors(EditTrack, EditSide, EditSector, statEdit)
      DoEvents
      Central.EditTimer.Enabled = True
    Case eoFormat:
      mLightEdit(6) = True
      Call DisplayEditOpLight
      oldRead = mLightRead
      mLightRead = 1
      Call DisplaySectors(EditTrack, EditSide, 1, statWrite)
      DoEvents
      IOResult = FormatTrack(IOFloppyA, EditTrack, EditSide, mLightEdit(1))
      Call SetSectorStatus(EditTrack, EditSide, 1, 18, IOResult)
      If (IOResult <> 0) And (mLightEdit(1) = True) Then Call MarkBadReservation(EditTrack, EditSide, 1, 18)
      Call DisplayEditValues
      mLightRead = oldRead
      mLightEdit(6) = False
      Call DisplayEditOpLight
      Call DisplaySectors(EditTrack, EditSide, EditSector, statEdit)
    Case eoOverwrite:
      mLightEdit(7) = True
      Call DisplayEditOpLight
      Call DisplaySectors(EditTrack, EditSide, EditSector, statWrite)
      DoEvents
      Call SetDiskSystemSectorData(EditTrack, EditSide, EditSector, NumSectors(), mLightEdit(1))
      i = SectorNumber(EditTrack, EditSide, EditSector)
      If i = 1 Then IOResult = DiskIO(IOWriteDisk, IOFloppyA, 1, EditTrack, EditSide, EditSector)
      If i > 18 Then IOResult = DiskIO(IOWriteDisk, IOFloppyA, NumSectors(), EditTrack, EditSide, EditSector)
      Call SetSectorStatus(EditTrack, EditSide, EditSector, NumSectors(), IOResult)
      If (IOResult <> 0) And (mLightEdit(1) = True) Then Call MarkBadReservation(EditTrack, EditSide, EditSector, NumSectors())
      Call DisplayEditValues
      mLightEdit(7) = False
      Call DisplayEditOpLight
      Call DisplaySectors(EditTrack, EditSide, EditSector, statEdit)
    Case eoMarkBad:
      mLightEdit(8) = True
      Call DisplayEditOpLight
      DoEvents
      Call MarkBadReservation(EditTrack, EditSide, EditSector, NumSectors())
      Call DisplayEditValues
      mLightEdit(8) = False
      Call DisplayEditOpLight
      Call DisplaySectors(EditTrack, EditSide, EditSector, statEdit)
    Case eoUnmark:
      mLightEdit(9) = True
      Call DisplayEditOpLight
      DoEvents
      Call UnMarkBadReservation(EditTrack, EditSide, EditSector, NumSectors())
      Call DisplayEditValues
      mLightEdit(9) = False
      Call DisplayEditOpLight
      Call DisplaySectors(EditTrack, EditSide, EditSector, statEdit)
    Case eoReading:
      IOResult = 0
      For i = 1 To 9216: IOdados(i) = &HF6: Next i
      If mLightEdit(3) = True Then
        Call DisplaySectors(EditTrack, EditSide, EditSector, statRead)
        DoEvents
        IOResult = DiskIO(IOReadDisk, IOFloppyA, NumSectors(), EditTrack, EditSide, EditSector)
        If (IOResult = 0) And (mLightEdit(2) = True) Then Call AutoCopy(EditTrack, EditSide, EditSector, NumSectors())
      End If
      If mLightEdit(4) = True Then
        If ((IOResult = 0) And (mLightEdit(3) = True)) Or _
           (GetSideTrackInside(EditTrack, EditSide, EditSector) = IOempty) Or _
           (GetSideTrackInside(EditTrack, EditSide, EditSector) = IObad) Then
          Call DisplaySectors(EditTrack, EditSide, EditSector, statWrite)
          DoEvents
          IOResult = DiskIO(IOWriteDisk, IOFloppyA, NumSectors(), EditTrack, EditSide, EditSector)
        End If
      End If
      If (IOResult = 0) And (mLightEdit(5) = True) Then
        Call DisplaySectors(EditTrack, EditSide, EditSector, statVerify)
        DoEvents
        IOResult = DiskIO(IOVerifyDisk, IOFloppyA, NumSectors(), EditTrack, EditSide, EditSector)
      End If
      If (IOResult <> 0) And (mLightEdit(1) = True) Then Call MarkBadReservation(EditTrack, EditSide, EditSector, NumSectors())
      If (mLightEdit(3) = True) Or (mLightEdit(4) = True) Or (mLightEdit(5) = True) Then
        Call SetSectorStatus(EditTrack, EditSide, EditSector, NumSectors(), IOResult)
        Call DisplaySectors(EditTrack, EditSide, EditSector, statEdit)
        Call DisplayEditValues
      End If
    Case eoMove:
      Call DisplayPosition(EditTrack, EditSide, EditSector)
      Call DisplaySectors(EditTrack, EditSide, EditSector, statEdit)
      DoEvents
    Case eoResetPos:
      i = NumSectors()
      Select Case i
        Case 18: EditSector = 1
        Case 9:  If EditSector + 4 >= 10 Then EditSector = 10 Else EditSector = 1
        Case 3:  EditSector = 1 + ((EditSector - 1) \ 3) * 3
      End Select
      Call DisplayPosition(EditTrack, EditSide, EditSector)
    Case eoEndEdit:
      Central.EditTimer.Enabled = False
      If MarkBad = True Then
        Call WriteDiskDATA
        MarkBad = False
      End If
      Call CloseDiskIO
      Call ReloadDisk
  End Select
  Editting = False
End Sub
