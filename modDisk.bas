Attribute VB_Name = "modDisk"
Option Explicit

'-----------------------------------------------Windows APIs
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function LockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Private Declare Function UnlockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'------------------------------------------------------Const
Private Const VWIN32_DIOC_DOS_IOCTL = 1
Private Const VWIN32_DIOC_DOS_INT13 = 4
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = 1
Private Const FILE_SHARE_WRITE = 2
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FILE_BEGIN = 0

Private Const BytesPerSector = 512

'------------------------------------------------------Types
Private Type DIOC_REGISTERS
  EBX As Long
  EDX As Long
  ECX As Long
  EAX As Long
  EDI As Long
  ESI As Long
  Flags As Long
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type DiskControlBlock
  StartSector As Long
  SectorRead As Integer
  Data(1 To 9216) As Byte
End Type

Private Type OVERLAPPED
  Internal As Long
  InternalHigh As Long
  offset As Long
  OffsetHigh As Long
  hEvent As Long
End Type

'------------------------------------------------------Enums
Public Enum DiskFunction
  IOReadDisk = 2
  IOWriteDisk = 3
  IOVerifyDisk = 4
  IOFormatDisk = 5
  IOResetSystem = 0
End Enum

Public Enum FloppyNumber
  IOFloppyA = 0
  IOFloppyB = 1
End Enum

Public Enum SectorType
  IOboot = 1
  IOfat2 = 2
  IOempty = 3
  IOdata = 4
  IOfat1 = 5
  IOdir = 6
  IObad = 7
End Enum

Public Enum StatType
  statNormal = 1
  statOk = 2
  statError = 3
  statRead = 4
  statWrite = 5
  statVerify = 6
  statEdit = 7
End Enum
'--------------------------------------------------Variables
Public SectorVal(1 To 2880) As Long        'sector value in FAT
Public SectorInfo(1 To 2880) As SectorType 'type of sector
Public SectorStat(1 To 2880) As StatType   'sector status
Public IOdados(0 To 19216) As Byte         'sector data
Public IsWinNT As Boolean                  'true if WinNT 2000 or XP
Private auxDTA1(1 To 512) As Byte          '1 sector data
Private auxDTA3(1 To 1536) As Byte         '3 sector data
Private auxDTA9(1 To 4608) As Byte         '9 sector data
Private auxDTA18(1 To 9216) As Byte        '18 sector data
Private FileHandle As Long
Private FileNumber As Long
Private FileChunk As Long


'-------------------------------------------InitializaDiskIO
Public Sub InitializeDiskIO()
  If IsWinNT = False Then
    FileHandle = CreateFile("\\.\VWIN32", 0, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
  Else
    FileHandle = CreateFile("\\.\A:", GENERIC_READ Or GENERIC_WRITE, _
           FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
  End If
End Sub

'------------------------------------------------CloseDiskIO
Public Sub CloseDiskIO()
  Call CloseHandle(FileHandle)
End Sub

'-----------------------------------------------------DiskIO
Public Function DiskIO(ByVal IOfunc As DiskFunction, ByVal IOdrive As FloppyNumber, ByVal IOnsec As Byte, ByVal IOtrack As Byte, ByVal IOside As Byte, ByVal IOsector As Byte) As Long
  Dim fResult As Long
  Dim BytesReturned As Long
  Dim Reg As DIOC_REGISTERS
  Dim res As Long
  Dim mByte As Long
  
  If IsWinNT = False Then
  
    'set Bios registers for int 13h
    Reg.EAX = IOfunc * 256 + IOnsec    ' INT 13 Function
    Reg.EBX = VarPtr(IOdados(0))       ' 32bit pointer to data
    Reg.ECX = IOtrack * 256 + IOsector ' Track & Sector
    Reg.EDX = IOside * 256 + IOdrive   ' Side & Drive
    Reg.Flags = 0
    'floppy disk IO
    fResult = DeviceIoControl(FileHandle, VWIN32_DIOC_DOS_INT13, _
        Reg, Len(Reg), Reg, Len(Reg), BytesReturned, 0)
    DiskIO = (Reg.EAX And &HFF00) / 256
  
  Else
    
    mByte = (SectorNumber(IOtrack, IOside, IOsector) - 1) * BytesPerSector
    Call SetFilePointer(FileHandle, mByte, 0, FILE_BEGIN)
    
    If IOfunc = IOReadDisk Then
      Call ReadFile(FileHandle, IOdados(0), IOnsec * BytesPerSector, res, 0&)
    End If
    If IOfunc = IOWriteDisk Then
      Call LockFile(FileHandle, LoWord(mByte), HiWord(mByte), LoWord(IOnsec * BytesPerSector), HiWord(IOnsec * BytesPerSector))
      Call WriteFile(FileHandle, IOdados(0), IOnsec * BytesPerSector, res, 0&)
      Call FlushFileBuffers(FileHandle)
      Call UnlockFile(FileHandle, LoWord(mByte), HiWord(mByte), LoWord(IOnsec * BytesPerSector), HiWord(IOnsec * BytesPerSector))
    End If
    If IOfunc = IOVerifyDisk Then
      Call ReadFile(FileHandle, IOdados(0), IOnsec * BytesPerSector, res, 0&)
    End If
    If IOfunc = IOFormatDisk Then
      Call LockFile(FileHandle, LoWord(mByte), HiWord(mByte), LoWord(IOnsec * BytesPerSector), HiWord(IOnsec * BytesPerSector))
      Call WriteFile(FileHandle, IOdados(0), IOnsec * BytesPerSector, res, 0&)
      Call FlushFileBuffers(FileHandle)
      Call UnlockFile(FileHandle, LoWord(mByte), HiWord(mByte), LoWord(IOnsec * BytesPerSector), HiWord(IOnsec * BytesPerSector))
    End If
    If IOfunc = IOResetSystem Then
      res = IOnsec * BytesPerSector
    End If
    
    If res = IOnsec * BytesPerSector Then DiskIO = 0 Else DiskIO = 255
    
  End If
End Function

'------------------------------------------------FormatTrack
Public Function FormatTrack(ByVal IOdrive As FloppyNumber, ByVal IOtrack As Byte, ByVal IOside As Byte, ByVal Light As Boolean) As Long
  Dim fResult As Long
  Dim BytesReturned As Long
  Dim Reg As DIOC_REGISTERS
  Dim res As Long
  Dim i As Long
  
  'Format Track
  For i = 1 To 18
    IOdados(0 + (i - 1) * 4) = IOtrack
    IOdados(1 + (i - 1) * 4) = IOside
    IOdados(2 + (i - 1) * 4) = i
    IOdados(3 + (i - 1) * 4) = 2
  Next i
  Reg.EAX = 5 * 256 + 18
  Reg.EBX = VarPtr(IOdados(0))
  Reg.ECX = IOtrack * 256
  Reg.EDX = IOside * 256 + IOdrive
  Reg.Flags = 0
  fResult = DeviceIoControl(FileHandle, VWIN32_DIOC_DOS_INT13, _
      Reg, Len(Reg), Reg, Len(Reg), BytesReturned, 0)
  res = (Reg.EAX And &HFF00) / 256
  If res <> 0 Then
    'Reset Disk
    Reg.EAX = 0                        ' Reset Disk system
    Reg.EBX = 0: Reg.ECX = 0
    Reg.EDX = IOdrive                  ' Drive
    Reg.Flags = 0
    fResult = DeviceIoControl(FileHandle, VWIN32_DIOC_DOS_INT13, _
        Reg, Len(Reg), Reg, Len(Reg), BytesReturned, 0)
  End If
  'Set Disk System area
  If IOtrack = 0 Then
    If IOside = 0 Then
      Call WriteBootSector
      For i = 0 To 4607: IOdados(i) = 0: Next i
      For i = 34 To 2880
        If SectorInfo(i) <> IObad Then
          SectorInfo(i) = IOempty
          SectorVal(i) = 0
        End If
        If (SectorInfo(i) = IObad) And (Light = False) Then
          SectorInfo(i) = IOempty
          SectorVal(i) = 0
        End If
      Next i
      Call WriteDiskDATA(False)
    Else
      For i = 0 To 3583: IOdados(i) = 0: Next i
      Call DiskIO(IOWriteDisk, IOFloppyA, 1, 0, 1, 1)
      Call DiskIO(IOWriteDisk, IOFloppyA, 7, 0, 1, 2)
      Call DiskIO(IOWriteDisk, IOFloppyA, 7, 0, 1, 9)
    End If
  End If
  FormatTrack = res
End Function

'-----------------------------------------------ReadDiskData
Public Sub ReadDiskDATA()
  Dim i As Long
  Dim Sector As Integer
  Dim val1 As Long
  Dim val2 As Long
  Dim FatPos As Integer
  
  Call InitializeDiskIO
  Call UltimateReadFAT
  'transfer data
  FatPos = 3
  Sector = 34
  Do While Sector <= 2880
    val1 = ((IOdados(FatPos + 1) And 15) * 256) + IOdados(FatPos)
    val2 = (IOdados(FatPos + 2) * 16) + ((IOdados(FatPos + 1) And 240) \ 16)
    SectorVal(Sector) = val1
    SectorInfo(Sector) = IOdata
    If val1 = 0 Then SectorInfo(Sector) = IOempty
    If (val1 >= &HFF0) And (val1 <= &HFF7) Then SectorInfo(Sector) = IObad
    If Sector < 2880 Then
      SectorVal(Sector + 1) = val2
      SectorInfo(Sector + 1) = IOdata
      If val2 = 0 Then SectorInfo(Sector + 1) = IOempty
      If (val2 >= &HFF0) And (val2 <= &HFF7) Then SectorInfo(Sector + 1) = IObad
    End If
    FatPos = FatPos + 3
    Sector = Sector + 2
  Loop
  SectorInfo(1) = IOboot
  For i = 2 To 10: SectorInfo(i) = IOfat1: Next i
  For i = 11 To 19: SectorInfo(i) = IOfat2: Next i
  For i = 20 To 33: SectorInfo(i) = IOdir: Next i
  For i = 1 To 2880: SectorStat(i) = statNormal: Next i
  SectorVal(1) = IOboot
  For i = 2 To 10: SectorVal(i) = IOfat1: Next i
  For i = 11 To 19: SectorVal(i) = IOfat2: Next i
  For i = 20 To 33: SectorVal(i) = IOdir: Next i
  Call CloseDiskIO
End Sub

'----------------------------------------------ClearDiskData
Public Sub ClearDiskData()
  Dim i As Long
  
  For i = 1 To 2880
    SectorInfo(i) = IOempty
    SectorStat(i) = statNormal
    SectorVal(i) = 0
  Next i
  SectorInfo(1) = IOboot
  For i = 2 To 10: SectorInfo(i) = IOfat1: Next i
  For i = 11 To 19: SectorInfo(i) = IOfat2: Next i
  For i = 20 To 33: SectorInfo(i) = IOdir: Next i
  SectorVal(1) = IOboot
  For i = 2 To 10: SectorVal(i) = IOfat1: Next i
  For i = 11 To 19: SectorVal(i) = IOfat2: Next i
  For i = 20 To 33: SectorVal(i) = IOdir: Next i
End Sub

'-------------------------------------------------GetSecType
Public Function GetSecType(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte) As SectorType
  Dim aux As Long
  Dim num As Long
  
  num = SectorNumber(Track, Side, Sector)
  aux = SectorVal(num)
  Select Case aux
    Case 0: GetSecType = IOempty
    Case 1: If num = 1 Then GetSecType = IOboot Else GetSecType = IOdata
    Case 2: If (num > 10) And (num < 20) Then GetSecType = IOfat2 Else GetSecType = IOdata
    Case 5: If (num > 1) And (num < 11) Then GetSecType = IOfat1 Else GetSecType = IOdata
    Case 6: If (num > 19) And (num < 34) Then GetSecType = IOdir Else GetSecType = IOdata
    Case Else:
      If (aux >= &HFF0) And (aux <= &HFF7) Then
         GetSecType = IObad
      Else
         GetSecType = IOdata
      End If
  End Select
End Function

'--------------------------------------------ResetDiskSystem
Public Sub DiskSystemReset()
  Call InitializeDiskIO
  Call DiskIO(IOResetSystem, IOFloppyA, 0, 0, 0, 0)
  Call CloseDiskIO
End Sub

'---------------------------------------------TestDiskChange
Public Function TestDiskChange() As Boolean
  Dim IOResult As Long
  
  Call InitializeDiskIO
  IOResult = DiskIO(IOReadDisk, IOFloppyA, 1, 0, 0, 1)
  If IOResult = &H6 Then
    TestDiskChange = True
  Else
    TestDiskChange = False
  End If
  Call CloseDiskIO
End Function

'----------------------------------------------TestDiskReady
Public Function TestDiskReady() As Boolean
  Dim IOResult As Long
  Dim Prep As Long
  
PrepDisk:
  Call InitializeDiskIO
  IOResult = DiskIO(IOReadDisk, IOFloppyA, 1, 0, 0, 1)
  If (IOResult <> 0) And (IOResult <> 6) Then
    TestDiskReady = False
    Prep = MsgBox("Error Reading Disk", vbExclamation Or vbAbortRetryIgnore, "Error")
    Select Case Prep
      Case 3: 'abort
         TestDiskReady = False
      Case 4: 'retry
         GoTo PrepDisk
      Case 5: 'Ignore
         TestDiskReady = True
    End Select
  Else
    TestDiskReady = True
  End If
  Call CloseDiskIO
End Function

'-------------------------------------------SetDiskSystemFAT
Public Sub SetDiskSystemFAT()
  Dim pFat As Long, Sec As Long

  Sec = 34: pFat = 3
  Do
    IOdados(pFat) = (SectorVal(Sec) And 255)
    IOdados(pFat + 1) = (SectorVal(Sec) \ 256)
    If Sec < 2880 Then IOdados(pFat + 1) = IOdados(pFat + 1) + ((SectorVal(Sec + 1) And 15) * 16)
    If Sec < 2880 Then IOdados(pFat + 2) = ((SectorVal(Sec + 1) And 4080) \ 16)
    pFat = pFat + 3
    Sec = Sec + 2
  Loop Until Sec = 2882
  For pFat = 4323 To 4607
    IOdados(pFat) = 0
  Next pFat
  IOdados(0) = &HF0
  IOdados(1) = &HFF
  IOdados(2) = &HFF
End Sub

'----------------------------------------------WriteDiskDATA
Public Sub WriteDiskDATA(Optional Side1 As Boolean = True)
  Dim HasBad As Long
  Dim res As Long
  Dim i As Long

  Call SetDiskSystemFAT
  HasBad = 0
  res = DiskIO(IOWriteDisk, IOFloppyA, 9, 0, 0, 2)
  If res <> 0 Then HasBad = res
  res = DiskIO(IOWriteDisk, IOFloppyA, 8, 0, 0, 11)
  If res <> 0 Then HasBad = res
  If Side1 Then
    For i = 1 To 512
      IOdados(i - 1) = IOdados(i - 1 + 4096)
    Next i
    res = DiskIO(IOWriteDisk, IOFloppyA, 1, 0, 1, 1)
    If res <> 0 Then HasBad = res
  End If
  If HasBad <> 0 Then
    MsgBox "Can't write disk File Allocation Table", vbExclamation Or vbOKOnly, "Error"
  End If
End Sub

'-----------------------------------------------CreateIdFile
Public Function CreateIdFile(ByVal FileName As String, ByVal Id As String, ByVal IdLen As Long) As Long
  Dim i As Long
  Dim cval As Byte
  
  'open for output
  On Error GoTo cfError
  FileNumber = FreeFile()
  Open FileName For Binary Access Write Lock Read Write As #FileNumber
  Do While Len(Id) < IdLen
    Id = Id & " "
  Loop
  Seek #FileNumber, 1
  For i = 1 To IdLen
    cval = Asc(Mid(Id, i, 1))
    Put #FileNumber, , cval
  Next i
  cval = 26
  Put #FileNumber, , cval
  CreateIdFile = 0
  FileChunk = IdLen + 1
  Exit Function
cfError:
  CreateIdFile = -1
End Function

'-------------------------------------------------OpenIdFile
Public Function OpenIdFile(ByVal FileName As String, ByVal Id As String, ByVal IdLen As Long) As Long
  Dim i As Long
  Dim cval As Byte
  Dim fId As String
  
  'open for output
  On Error GoTo cfError
  FileNumber = FreeFile()
  Open FileName For Binary Access Read Lock Write As #FileNumber
  Do While Len(Id) < IdLen
    Id = Id & " "
  Loop
  If LOF(FileNumber) = 0 Then
    OpenIdFile = -1
    Exit Function
  End If
  Seek #FileNumber, 1
  For i = 1 To IdLen
    Get #FileNumber, i, cval
    If cval <> Asc(Mid(Id, i, 1)) Then
      OpenIdFile = -2
      Exit Function
    End If
  Next i
  OpenIdFile = 0
  FileChunk = IdLen + 1
  Exit Function
cfError:
  OpenIdFile = -1   'read error
End Function

'------------------------------------------------CloseIdFile
Public Sub CloseIdFile()
  On Error Resume Next
  Close #FileNumber
End Sub

'------------------------------------------------WriteIOData
Public Sub WriteIOData(ByVal nSect As Byte)
  Dim i As Long
  
  For i = 1 To 512 * nSect
    Select Case nSect
      Case 1:  auxDTA1(i) = IOdados(i - 1)
      Case 3:  auxDTA3(i) = IOdados(i - 1)
      Case 9:  auxDTA9(i) = IOdados(i - 1)
      Case 18: auxDTA18(i) = IOdados(i - 1)
    End Select
  Next i
  On Error GoTo cfError
  Select Case nSect
    Case 1:  Put #FileNumber, , auxDTA1
    Case 3:  Put #FileNumber, , auxDTA3
    Case 9:  Put #FileNumber, , auxDTA9
    Case 18: Put #FileNumber, , auxDTA18
  End Select
cfError:
End Sub

'-------------------------------------------------ReadIOData
Public Sub ReadIOData(ByVal nSect As Byte)
  Dim i As Long
  
  On Error GoTo cfError
  For i = 1 To 512 * nSect
    Get #FileNumber, , IOdados(i - 1)
  Next i
cfError:
End Sub

'--------------------------------------------UltimateReadFAT
Private Sub UltimateReadFAT()
  Dim IOResult As Long
  Dim IOsecFAT(0 To 4607) As Byte
  Dim i As Long, j As Long
  Dim CancelAction As Boolean
    
  IOResult = DiskIO(IOReadDisk, IOFloppyA, 9, 0, 0, 2)
  'ask for cancel
  If IOResult = 0 Then Exit Sub
  i = MsgBox("Errors found in FAT area." & Chr(13) & Chr(10) & "Try to read good sectors in FAT2 ?", vbExclamation Or vbYesNo, "Error")
  CancelAction = True
  If i = vbYes Then
    CancelAction = False
  End If
  If (IOResult <> 0) And (CancelAction = False) Then
    'read FAT one sector at a time
    For j = 1 To 9
      IOResult = DiskIO(IOReadDisk, IOFloppyA, 1, 0, 0, 1 + j)
      If IOResult <> 0 Then
         If j = 9 Then IOResult = DiskIO(IOReadDisk, IOFloppyA, 1, 0, 1, 1)
         If j < 9 Then IOResult = DiskIO(IOReadDisk, IOFloppyA, 1, 0, 0, 10 + j)
      End If
      'pass data
      For i = 0 To 511
        IOsecFAT(i + (j - 1) * 512) = IOdados(i)
      Next i
    Next j
    For i = 0 To 4607
      IOdados(i) = IOsecFAT(i)
    Next i
  End If
End Sub

'---------------------------------------------isExpectedSize
Public Function isExpectedSize(ByVal Head As Long, ByVal Chunk As Long, ByVal Tam As Long) As Boolean
  Dim fsize As Long
  Dim fsing As Single
  
  On Error Resume Next
  fsize = LOF(FileNumber)
  If (fsize > Tam) And (Tam > 0) Then
    isExpectedSize = False
    Exit Function
  End If
  fsize = fsize - Head
  If Chunk > 0 Then
    If (fsize Mod Chunk) <> 0 Then
      isExpectedSize = False
      Exit Function
    End If
  End If
  isExpectedSize = True
End Function

'------------------------------------------------GetImageFAT
Public Function GetImageFAT() As Long()
  Dim auxFAT(1 To 2880) As Long
  Dim Sector As Integer
  Dim FatPos As Integer
  Dim i As Long
  
  Seek #FileNumber, FileChunk + 1 + 512
  Call ReadIOData(9)
  'transfer data
  FatPos = 3
  Sector = 34
  Do While Sector <= 2880
    auxFAT(Sector) = ((IOdados(FatPos + 1) And 15) * 256) + IOdados(FatPos)
    If Sector < 2880 Then auxFAT(Sector + 1) = (IOdados(FatPos + 2) * 16) + ((IOdados(FatPos + 1) And 240) \ 16)
    FatPos = FatPos + 3
    Sector = Sector + 2
  Loop
  auxFAT(1) = IOboot
  For i = 2 To 10: auxFAT(i) = IOfat1: Next i
  For i = 11 To 19: auxFAT(i) = IOfat2: Next i
  For i = 20 To 33: auxFAT(i) = IOdir: Next i
  Seek #FileNumber, FileChunk + 1
  GetImageFAT = auxFAT
End Function

'----------------------------------------------GetFloppyBoot
Private Function GetFloppyBoot() As Byte()
  Const OEMid = "        "   'will be replaced by windows
  Const SysID = "FAT12   "
  Dim BootS As String
  Dim Boot(1 To 512) As Byte
  Dim i As Long
  Dim tick As Long
  Dim valB As Long
     
  i = Timer
  'Jump Code -------------------------- 3 bytes
  Boot(1) = &HEB: Boot(2) = &H3C        'JMUP +3C
  Boot(3) = &H90                        'NOP
  'OEM Id ----------------------------- 8 bytes
  For i = 1 To 8: Boot(3 + i) = CByte(Asc(Mid(OEMid, i, 1))): Next i
  'Bios Parameter Block --------------- 25 bytes
  Boot(12) = 0: Boot(13) = 2            'bytes per sector=512
  Boot(14) = 1                          'sectors per cluster=1
  Boot(15) = 1: Boot(16) = 0            'Reserved sectors=1 (boot)
  Boot(17) = 2                          'Number of FATs=2
  Boot(18) = 224: Boot(19) = 0          'Number of root entries=224 (512*14/32)
  Boot(20) = &H40: Boot(21) = &HB       'Number of sectors=2880
  Boot(22) = &HF0                       'Media Descriptor=&HF0 (1.44MB)
  Boot(23) = 9: Boot(24) = 0            'Sectors per FAT=9
  Boot(25) = 18: Boot(26) = 0           'Sectors per Track=18
  Boot(27) = 2: Boot(28) = 0            'Number of Heads=2
  For i = 29 To 36: Boot(i) = 0: Next i 'Number of (Hidden,Large) sectors = (0,0)
  'Extended Bios Parameter Block ------ 25 bytes
  Boot(37) = 0                          'Physical drive number=0 (floppy)
  Boot(38) = 0                          'Reserved Flags=0
  Boot(39) = &H29                       'Signature=&H29
  tick = GetTickCount()
  valB = (tick And &HFF000000) \ &H1000000
  Boot(40) = CByte(valB)
  valB = (tick And &HFF0000) \ &H10000
  Boot(41) = CByte(valB)
  valB = (tick And &HFF00&) \ &H100&
  Boot(42) = CByte(valB)
  valB = tick And &HFF&
  Boot(43) = CByte(valB)                'Id Serial-Number (random)
  For i = 44 To 54: Boot(i) = 0: Next i 'old volume
  For i = 54 To 61: Boot(i) = CByte(Asc(Mid(SysID, i - 53, 1))): Next i
  'Boot Executable Code --------------- 38 Bytes
  Boot(62) = &HFA             'CLI
  Boot(63) = &HBC             'MOV SP, 7C00   #CODE AT 7C00
  Boot(64) = &H0              '
  Boot(66) = &H7C             '
  Boot(67) = &HFB             'STI
  Boot(68) = &HB2             'MOV DL, 0
  Boot(69) = &H0              '
  Boot(70) = &H33             'XOR AX, AX
  Boot(71) = &HC0             '
  Boot(72) = &HCD             'INT 13         #RESET DISK SYSTEM
  Boot(73) = &H13             '
  Boot(74) = &HE              'PUSH CS
  Boot(75) = &H1F             'POP DS         #DATA IN SAME AREA
  Boot(76) = &HFC             'CLD            #FORWARD MOVING
  Boot(77) = &HBE             'MOV SI, 7C63   #ADDRESS OF DATA
  Boot(78) = &H63             '
  Boot(79) = &H7C             '
  Boot(80) = &HAC             'LODSB          #GET BYTE AT ADDRESS
  Boot(81) = &HA              'OR AL, AL
  Boot(82) = &HC0             '
  Boot(83) = &H74             'JE +9          #JUMP IF ZERO TO POSITION 94
  Boot(84) = &H9              '
  Boot(85) = &HB4             'MOV AH, 0E
  Boot(86) = &HE              '
  Boot(87) = &HBB             'MOV BX, 7      #FOREGROUND COLOR
  Boot(88) = &H7              '
  Boot(89) = &H0              '
  Boot(90) = &HCD             'INT 10         #WRITE CHAR
  Boot(91) = &H10             '
  Boot(92) = &HEB             'JUMP -14       #JUMP TO POSITION 80
  Boot(93) = &HF2             '
  Boot(94) = &H33             'XOR AX, AX
  Boot(95) = &HC0             '
  Boot(96) = &HCD             'INT 16         #WAIT FOR KEYSTROKE
  Boot(97) = &H16             '
  Boot(98) = &HCD             'INT 19         #BOOTSTRAP LOADER (warm boot)
  Boot(99) = &H19             '
  'Boot error text -------------------- 70 bytes
  BootS = "Not a system disk or disk error." & Chr(13) & Chr(10) & "Replace or remove and press any key."
  For i = 100 To 169: Boot(i) = CByte(Asc(Mid(BootS, i - 99, 1))): Next i
  'Empty area ------------------------- 341 bytes
  For i = 170 To 510: Boot(i) = 0: Next i
  'Boot End Code ---------------------- 2 bytes
  Boot(511) = &H55: Boot(512) = &HAA
  'return
  GetFloppyBoot = Boot
End Function

'--------------------------------------------WriteBootSector
Public Sub WriteBootSector()
  Dim BootAux() As Byte
  Dim i As Long
  
  BootAux = GetFloppyBoot()
  For i = 1 To 512
    IOdados(i - 1) = BootAux(i)
  Next i
  Call DiskIO(IOWriteDisk, IOFloppyA, 1, 0, 0, 1)
End Sub

'----------------------------------------SetDeviceParameters
Public Sub SetDeviceParameters(ByVal IOdrive As FloppyNumber)
  Dim fResult As Long
  Dim BytesReturned As Long
  Dim Reg As DIOC_REGISTERS
  Dim res As Long
  Dim i As Long
  
  Reg.EAX = 8 * 256
  Reg.EBX = 0
  Reg.ECX = 0
  Reg.EDX = IOdrive                  ' Drive
  Reg.Flags = 0
  fResult = DeviceIoControl(FileHandle, VWIN32_DIOC_DOS_INT13, _
      Reg, Len(Reg), Reg, Len(Reg), BytesReturned, 0)
  Call CopyMemory(ByVal VarPtr(IOdados(0)), ByVal Reg.EDI, 30)
  'set media type for format
  Reg.EAX = &H18 * 256               ' Set Media Type
  Reg.EBX = 0
  Reg.ECX = 79 * 256 + 18            ' Tracks + Sectors/Track
  Reg.EDX = IOdrive                  ' Drive
  Reg.Flags = 0
  fResult = DeviceIoControl(FileHandle, VWIN32_DIOC_DOS_INT13, _
      Reg, Len(Reg), Reg, Len(Reg), BytesReturned, 0)
  Call CopyMemory(ByVal VarPtr(IOdados(0)), ByVal Reg.EDI, 30)
  'Set Parameters
  IOdados(0) = 0                     ' Function
  IOdados(1) = 7                     ' Device Type
  IOdados(2) = 1: IOdados(3) = 0     ' Device Attribute
  IOdados(4) = 80: IOdados(5) = 0    ' Tracks
  IOdados(6) = 0                     ' Media Type
  IOdados(7) = 0: IOdados(8) = 2     ' Bytes per Sector = 512
  IOdados(9) = 1                     ' sectors per Cluster
  IOdados(10) = 1: IOdados(11) = 0   ' Reserved Sectors
  IOdados(12) = 2                    ' Number of FATs
  IOdados(13) = 224: IOdados(14) = 0 ' Max Root Entries
  IOdados(15) = &H40: IOdados(16) = &HB ' Number of Sectors=2880
  IOdados(17) = &HF0                 ' Media Descriptor
  IOdados(18) = 9: IOdados(19) = 0   ' Sector in FAT
  IOdados(20) = 18: IOdados(21) = 0  ' Sectors per Track
  IOdados(22) = 2: IOdados(23) = 0   ' Number of Heads
  For i = 24 To 37: IOdados(i) = 0: Next i 'Hidden/Long/Reserved
  IOdados(38) = 18: IOdados(39) = 0  ' Number of Sectors
  For i = 1 To 18
    IOdados(40 + (i - 1) * 4) = i
    IOdados(41 + (i - 1) * 4) = 0    ' Sector Number
    IOdados(42 + (i - 1) * 4) = 0
    IOdados(43 + (i - 1) * 4) = 2    ' Sector Size=512
  Next i
  Reg.EAX = &H440D                   ' INT 21 IOCTL
  Reg.EBX = IOdrive + 1              ' Drive
  Reg.ECX = &H840                    ' Disk Drive Set Device Parameters
  Reg.EDX = VarPtr(IOdados(0))       ' Parameter Block Buffer
  Reg.Flags = 0
  fResult = DeviceIoControl(FileHandle, VWIN32_DIOC_DOS_IOCTL, _
      Reg, Len(Reg), Reg, Len(Reg), BytesReturned, 0)
End Sub

'-----------------------------------------------CountSectors
Public Sub CountSectors(ByRef Bad As Long, ByRef Good As Long, ByRef Avail As Long, ByRef Percent As Long)
  Dim i As Long
  
  Bad = 0
  Good = 0
  Avail = 0
  For i = 1 To 2880
    If SectorInfo(i) = IObad Then Bad = Bad + 1
    If SectorInfo(i) = IOempty Then Avail = Avail + 1
    If (SectorInfo(i) = IOempty) Or (SectorInfo(i) = IOdata) Then
      Good = Good + 1
    End If
  Next i
  Percent = (Avail * 100) \ 2847
  If Percent > 100 Then Percent = 100
  Avail = Avail * 512
End Sub

'------------------------------------SetDiskSystemSectorData
Public Sub SetDiskSystemSectorData(ByVal Track As Byte, ByVal Side As Byte, ByVal Sector As Byte, ByVal nSectors As Byte, ByVal Light As Boolean)
  Dim i As Long
  Dim nSec As Long
  Dim Tam As Long
  Dim curSec As Byte
  Dim BootAux() As Byte
  
  For i = 1 To 9216: IOdados(i - 1) = &HF6: Next i
  nSec = SectorNumber(Track, Side, Sector)
  'clear data
  If ((nSec > 1) And (nSec < 34)) Or ((nSec + nSectors - 1 > 1) And (nSec + nSectors - 1 < 34)) Then
    For i = 34 To 2880
      If SectorInfo(i) <> IObad Then
        SectorInfo(i) = IOempty
        SectorVal(i) = 0
      End If
      If (SectorInfo(i) = IObad) And (Light = False) Then
        SectorInfo(i) = IOempty
        SectorVal(i) = 0
      End If
    Next i
    Call WriteDiskDATA
    For i = 1 To 9216: IOdados(i - 1) = 0: Next i
  End If
  'boot sector
  If nSec = 1 Then
    BootAux = GetFloppyBoot()
    For i = 1 To 512
      IOdados(i - 1) = BootAux(i)
    Next i
  End If
End Sub
