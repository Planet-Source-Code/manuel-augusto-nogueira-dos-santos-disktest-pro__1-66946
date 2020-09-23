Attribute VB_Name = "modGeral"
Option Explicit

'------------------------------------------------Windows API
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As Point, ByVal nCount As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias _
       "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long


Const ALTERNATE = 1

'--------------------------------------------Types and Enums
Public Type Point
  X As Long
  Y As Long
End Type

Public Enum AMPMmode
  H24 = 1
  H12 = 2
End Enum


Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128 '  Maintenance string for PSS usage
End Type

'dwPlatforID Constants
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2


'-------------------------------------------------------GetWindowsVersion
Public Function GetWindowsVersion() As Long  '0-don't know  1-Win9x/Me  2-WinNT/2000/XP
  Dim tOSVer As OSVERSIONINFO
    
  tOSVer.dwOSVersionInfoSize = Len(tOSVer)
  GetVersionEx tOSVer
  GetWindowsVersion = 0
  Select Case tOSVer.dwPlatformId
    Case VER_PLATFORM_WIN32_NT
      If tOSVer.dwMajorVersion >= 5 Then
        GetWindowsVersion = 2  ' Windows 2000
      Else
        GetWindowsVersion = 2  ' Windows NT
      End If
    Case Else
      If tOSVer.dwMajorVersion >= 5 Then
        GetWindowsVersion = 1  ' Windows ME
      ElseIf tOSVer.dwMajorVersion = 4 And tOSVer.dwMinorVersion > 0 Then
        GetWindowsVersion = 1  ' Windows 98
      Else
        GetWindowsVersion = 1  ' Windows 95
      End If
  End Select
End Function

'-------------------------------------------------------Hi and Low Word
Public Function LoWord(ByVal LongVal As Long) As Long
  LoWord = LongVal And &HFFFF&
End Function

Public Function HiWord(ByVal LongVal As Long) As Long
  If LongVal = 0 Then
    HiWord = 0
    Exit Function
  End If
  HiWord = LongVal \ &H10000 And &HFFFF&
End Function

'-------------------------------------------------------Str3
Public Function Str03(ByVal Value As Long) As String
  Dim res As String
  
  res = Trim(Str(Value))
  Do While Len(res) < 3
    res = "0" & res
  Loop
  Str03 = res
End Function

'-------------------------------------------------------Str0N
Public Function Str0N(ByVal Value As Long, ByVal Tam As Byte) As String
  Dim res As String
  
  res = Trim(Str(Value))
  Do While Len(res) < Tam
    res = "0" & res
  Loop
  Str0N = res
End Function

'-------------------------------------------------------StrN
Public Function StrN(ByVal Tam As Long, ByVal Text As String) As String
  Do While Len(Text) < Tam
    Text = Text & " "
  Loop
  StrN = Text
End Function

'---------------------------------------------------StrClock
Public Function StrClock(ByVal Value As Long) As String
  Dim Horas As Long, aux As Long, Minu As Long, Secs As Long
  Dim res As String
  
  Horas = Value \ 3600
  aux = Value - (Horas * 3600)
  Minu = aux \ 60
  Secs = aux - (Minu * 60)
  res = Str0N(Horas, 2) & ":" & Str0N(Minu, 2) & ":" & Str0N(Secs, 2)
  StrClock = res
End Function

'-------------------------------------------GetFormCursorPos
Public Sub GetFormCursorPos(ByRef Mouse As Point, ByVal fX As Long, ByVal fY As Long, ByRef FormPos As Point)
  Dim formX As Long
  Dim formY As Long
  
  formX = fX / Screen.TwipsPerPixelX
  formY = fY / Screen.TwipsPerPixelY
  FormPos.X = Mouse.X - formX
  FormPos.Y = Mouse.Y - formY
End Sub

'----------------------------------------------IsInsideImage
Public Function IsInsideImage(ByRef pos As Point, ByRef Pic As Image) As Boolean
  Dim resp As Boolean
  
  resp = False
  If pos.X >= Pic.Left And pos.X < Pic.Left + Pic.Width And _
     pos.Y >= Pic.Top And pos.Y < Pic.Top + Pic.Height Then
    resp = True
  End If
  IsInsideImage = resp
End Function

'------------------------------------------------IsInsideBox
Public Function IsInsideBox(ByVal X As Long, ByVal Y As Long, _
  ByVal X1 As Long, ByVal Y1 As Long, ByVal W As Long, ByVal H As Long) As Boolean
  Dim resp As Boolean
  
  resp = False
  If X >= X1 And X < X1 + W And Y >= Y1 And Y < Y1 + H Then
    resp = True
  End If
  IsInsideBox = resp
End Function

'--------------------------------------------------DrawBox3D
Public Sub DrawBox3D(ByRef DrawForm As Form, ByVal modo As Byte, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long)
  Select Case modo
    Case 1: 'Lowered 2 lines
      DrawForm.ForeColor = RGB(68, 70, 68)
      DrawForm.Line (X, Y)-(X + W - 1, Y)
      DrawForm.Line (X, Y)-(X, Y + H - 1)
      DrawForm.ForeColor = RGB(148, 146, 148)
      DrawForm.Line (X + 1, Y + 1)-(X + W - 1, Y + 1)
      DrawForm.Line (X + 1, Y + 1)-(X + 1, Y + H - 1)
      DrawForm.Line (X, Y + H - 1)-(X, Y + H)
      DrawForm.ForeColor = RGB(204, 206, 204)
      DrawForm.Line (X + W - 2, Y + 2)-(X + W - 2, Y + H - 1)
      DrawForm.Line (X + 2, Y + H - 2)-(X + W - 1, Y + H - 2)
      DrawForm.ForeColor = RGB(244, 246, 244)
      DrawForm.Line (X + 1, Y + H - 1)-(X + W, Y + H - 1)
      DrawForm.Line (X + W - 1, Y)-(X + W - 1, Y + H)
    Case 2: 'Raised 2 lines
      DrawForm.ForeColor = RGB(244, 246, 244)
      DrawForm.Line (X, Y)-(X + W - 1, Y)
      DrawForm.Line (X, Y)-(X, Y + H - 1)
      DrawForm.ForeColor = RGB(204, 206, 204)
      DrawForm.Line (X + 1, Y + 1)-(X + W - 2, Y + 1)
      DrawForm.Line (X + 1, Y + 1)-(X + 1, Y + H - 2)
      DrawForm.ForeColor = RGB(148, 146, 148)
      DrawForm.Line (X + W - 2, Y + 1)-(X + W - 2, Y + H - 1)
      DrawForm.Line (X + 1, Y + H - 2)-(X + W - 1, Y + H - 2)
      DrawForm.Line (X, Y + H - 1)-(X, Y + H)
      DrawForm.ForeColor = RGB(68, 70, 68)
      DrawForm.Line (X + 1, Y + H - 1)-(X + W, Y + H - 1)
      DrawForm.Line (X + W - 1, Y)-(X + W - 1, Y + H)
    Case 3: 'Lowered 1 line
      DrawForm.ForeColor = RGB(68, 70, 68)
      DrawForm.FillColor = RGB(68, 70, 68)
      DrawForm.Line (X, Y)-(X + W - 1, Y), , BF
      DrawForm.Line (X, Y)-(X, Y + H - 2), , BF
      DrawForm.ForeColor = RGB(244, 246, 244)
      DrawForm.FillColor = RGB(244, 246, 244)
      DrawForm.Line (X, Y + H - 1)-(X + W - 1, Y + H - 1), , BF
      DrawForm.Line (X + W - 1, Y + 1)-(X + W - 1, Y + H - 1), , BF
    Case 4: 'Raised 1 line
      DrawForm.ForeColor = RGB(244, 246, 244)
      DrawForm.FillColor = RGB(244, 246, 244)
      DrawForm.Line (X, Y)-(X + W - 1, Y), , BF
      DrawForm.Line (X, Y)-(X, Y + H - 2), , BF
      DrawForm.ForeColor = RGB(68, 70, 68)
      DrawForm.FillColor = RGB(68, 70, 68)
      DrawForm.Line (X, Y + H - 1)-(X + W - 1, Y + H - 1), , BF
      DrawForm.Line (X + W - 1, Y + 1)-(X + W - 1, Y + H - 1), , BF
  End Select
End Sub

'---------------------------------------Ponteiros do Relogio
Public Sub Ponteiros(ByRef DrawForm As Form, ByVal t As Long, ByVal modo As Long, ByVal relX As Integer, ByVal relY As Integer, ByVal relL As Integer)
  Dim Horas As Long, aux As Long, Minu As Long, Secs As Long
  Dim ho As Single, mi As Single, se As Single
  Dim X1 As Integer, Y1 As Integer
  Dim X2 As Integer, Y2 As Integer

  Horas = t \ 3600
  aux = t - (Horas * 3600)
  Minu = aux \ 60
  Secs = aux - (Minu * 60)
  If (modo = 1) Or (modo = 3) Then
    DrawForm.ForeColor = RGB(84, 250, 164)
  Else
    DrawForm.ForeColor = RGB(11, 35, 34)
  End If
  'afixar horas
  ho = Horas * 0.52359877 - 1.570796327
  If modo < 2 Then ho = ho + Minu * 0.008726646
  X2 = Round(relL * 0.6 * Cos(ho))
  Y2 = Round(relL * 0.6 * Sin(ho))
  If (Horas > 0) Or (modo < 2) Then
    DrawForm.Line (relX, relY)-(relX + X2, relY + Y2)
  End If
   'afixar minutos
  mi = Minu * 0.10471955 - 1.570796327
  X2 = Round(relL * Cos(mi))
  Y2 = Round(relL * Sin(mi))
  If (Minu > 0) Or (modo < 2) Then
    DrawForm.Line (relX, relY)-(relX + X2, relY + Y2)
  End If
  'afixar segundos
  se = Secs * 0.10471955 - 1.570796327
  X1 = Round(relL * 0.65 * Cos(se))
  Y1 = Round(relL * 0.65 * Sin(se))
  X2 = Round(relL * Cos(se))
  Y2 = Round(relL * Sin(se))
  DrawForm.Line (relX + X1, relY + Y1)-(relX + X2, relY + Y2)
End Sub

'----------------------------------------------DigitalNumber
Public Sub DigitalNumber(ByRef DrawForm As Form, ByVal X As Long, ByVal Y As Long, ByVal Numero As Byte, Tam As Byte)
  Dim isON As Boolean
  Dim Segment As Byte
  
  For Segment = 1 To 7
    'determinar se o segmento está iluminado
    isON = False
    Select Case Segment
      Case 1:
        Select Case Numero
          Case 0, 1, 3, 4, 5, 6, 7, 8, 9: isON = True
        End Select
      Case 2:
        Select Case Numero
          Case 0, 2, 3, 5, 6, 8: isON = True
        End Select
      Case 3:
        Select Case Numero
          Case 0, 2, 6, 8: isON = True
        End Select
      Case 4:
        Select Case Numero
          Case 2, 3, 4, 5, 6, 8, 9: isON = True
        End Select
      Case 5:
        Select Case Numero
          Case 0, 1, 2, 3, 4, 7, 8, 9: isON = True
        End Select
      Case 6:
        Select Case Numero
          Case 0, 2, 3, 5, 7, 8, 9: isON = True
        End Select
      Case 7:
        Select Case Numero
          Case 0, 4, 5, 6, 8, 9: isON = True
        End Select
    End Select
    'cores
    If isON Then
      DrawForm.FillColor = RGB(84, 250, 164)
      DrawForm.ForeColor = RGB(84, 250, 164)
    Else
      DrawForm.FillColor = RGB(11, 35, 34)
      DrawForm.ForeColor = RGB(11, 35, 34)
    End If
    'desenha linha do segmento
    If Tam = 2 Then
      Select Case Segment
        Case 7: DrawForm.Line (X, Y + 1)-(X, Y + 3), , BF
        Case 6: DrawForm.Line (X + 1, Y)-(X + 3, Y), , BF
        Case 5: DrawForm.Line (X + 4, Y + 1)-(X + 4, Y + 3), , BF
        Case 4: DrawForm.Line (X + 1, Y + 4)-(X + 3, Y + 4), , BF
        Case 3: DrawForm.Line (X, Y + 5)-(X, Y + 7), , BF
        Case 2: DrawForm.Line (X + 1, Y + 8)-(X + 3, Y + 8), , BF
        Case 1: DrawForm.Line (X + 4, Y + 5)-(X + 4, Y + 7), , BF
      End Select
    End If
    If Tam = 1 Then
      Select Case Segment
        Case 7: DrawForm.Line (X, Y + 1)-(X, Y + 2), , BF
        Case 6: DrawForm.Line (X + 1, Y)-(X + 2, Y), , BF
        Case 5: DrawForm.Line (X + 3, Y + 1)-(X + 3, Y + 2), , BF
        Case 4: DrawForm.Line (X + 1, Y + 3)-(X + 2, Y + 3), , BF
        Case 3: DrawForm.Line (X, Y + 4)-(X, Y + 5), , BF
        Case 2: DrawForm.Line (X + 1, Y + 6)-(X + 2, Y + 6), , BF
        Case 1: DrawForm.Line (X + 3, Y + 4)-(X + 3, Y + 5), , BF
      End Select
    End If
  Next Segment
End Sub

'-------------------------------------------------DigitalINT
Public Sub DigitalINT(ByRef DrawForm As Form, ByVal X As Long, ByVal Y As Long, ByVal Numero As Long, ByVal Tam As Byte, ByVal Digits As Byte)
  Dim ds As String
  Dim DC As Byte

  ds = Mid(Str(Numero), 2)
  Do While Len(ds) < Digits
    ds = "0" & ds
  Loop
  For DC = 1 To Len(ds)
    Call DigitalNumber(DrawForm, X, Y, Val(Mid(ds, DC, 1)), Tam)
    Select Case Tam
      Case 1: X = X + 5
      Case 2: X = X + 6
    End Select
  Next DC
End Sub

'---------------------------------------------CalcNowSeconds
Public Function CalcNowSeconds(ByVal modo As AMPMmode) As Long
  Dim H As Long
  Dim m As Long
  Dim s As Long
  Dim Agora As Date
  
  Agora = Now
  H = Hour(Agora): m = Minute(Agora): s = Second(Agora)
  If modo = H12 Then
    If H > 12 Then H = H - 12
  End If
  CalcNowSeconds = s + m * 60 + H * 3600
End Function

'---------------------------------------------FilledTriangle
Public Sub FilledTriangle(ByVal hDC As Long, ByVal Cor As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long)
  Dim Poly(1 To 3) As Point
  Dim NumPoint As Long
  Dim hBrush As Long
  Dim hRgn As Long
  
  NumPoint = 3
  Poly(1).X = X1:  Poly(1).Y = Y1
  Poly(2).X = X2:  Poly(2).Y = Y2
  Poly(3).X = X3:  Poly(3).Y = Y3
  Polygon hDC, Poly(1), NumPoint
  hBrush = CreateSolidBrush(Cor)
  hRgn = CreatePolygonRgn(Poly(1), NumPoint, ALTERNATE)
  If hRgn Then
    FillRgn hDC, hRgn, hBrush
    DeleteObject hRgn
  End If
  DeleteObject hBrush
End Sub

'---------------------------------------------DigitalBitText
Public Sub DigitalBitText(ByRef mForm As Form, ByVal pX As Long, ByVal pY As Long, ByVal Text As String, ByVal Cor As OLE_COLOR, ByVal Back As OLE_COLOR, ByVal modo As Byte)
  Dim posChar As Long
  Dim StrTam As Long
  Dim CharVal As Long
  Dim CharX As Long, CharY As Long
  Dim CharBit As Long
  Dim PosX As Long
  Dim Lines As Byte
  
  StrTam = Len(Text)
  PosX = pX
  Lines = 8
  If modo = 2 Then Lines = 5
  For posChar = 1 To StrTam
    CharVal = Asc(Mid(Text, posChar, 1))
    If (CharVal < 32) Or (CharVal > 90) Then CharVal = 45
    mForm.ForeColor = Back
    mForm.FillColor = Back
    mForm.Line (PosX, pY)-(PosX + 6, pY + Lines), , BF
    For CharX = 1 To 5
      CharBit = 1
      For CharY = 1 To Lines
        If modo = 1 Then
          If (GetDigitalCharBits(CharVal, CharX) And CharBit) = CharBit Then
            mForm.ForeColor = Cor
            mForm.FillColor = Cor
            mForm.Line (PosX, pY + CharY - 1)-(PosX, pY + CharY - 1), , BF
          End If
        End If
        If modo = 2 Then
          If (GetDigitalSmallCharBits(CharVal, CharX) And CharBit) = CharBit Then
            mForm.ForeColor = Cor
            mForm.FillColor = Cor
            mForm.Line (PosX, pY + CharY - 1)-(PosX, pY + CharY - 1), , BF
          End If
        End If
        CharBit = CharBit * 2
      Next CharY
      PosX = PosX + 1
    Next CharX
  Next posChar
End Sub

Private Function GetDigitalCharBits(ByVal Char As Long, ByVal col As Long) As Long
  Dim DB(1 To 5) As Byte
  
  Select Case Char
    Case 32: DB(1) = 0:   DB(2) = 0:   DB(3) = 0:   DB(4) = 0:   DB(5) = 0   'SPACE
    Case 33: DB(1) = 0:   DB(2) = 0:   DB(3) = 95:  DB(4) = 95:  DB(5) = 0   '!
    Case 34: DB(1) = 0:   DB(2) = 7:   DB(3) = 0:   DB(4) = 7:   DB(5) = 0   '"
    Case 35: DB(1) = 40:  DB(2) = 124: DB(3) = 40:  DB(4) = 124: DB(5) = 40  '#
    Case 36: DB(1) = 38:  DB(2) = 127: DB(3) = 42:  DB(4) = 127: DB(5) = 50  '$
    Case 37: DB(1) = 76:  DB(2) = 44:  DB(3) = 16:  DB(4) = 104: DB(5) = 100 '%
    Case 38: DB(1) = 52:  DB(2) = 74:  DB(3) = 90:  DB(4) = 36:  DB(5) = 80  '&
    Case 39: DB(1) = 0:   DB(2) = 0:   DB(3) = 4:   DB(4) = 3:   DB(5) = 0   ''
    Case 40: DB(1) = 0:   DB(2) = 28:  DB(3) = 34:  DB(4) = 65:  DB(5) = 0   '(
    Case 41: DB(1) = 0:   DB(2) = 65:  DB(3) = 34:  DB(4) = 28:  DB(5) = 0   ')
    Case 42: DB(1) = 84:  DB(2) = 56:  DB(3) = 124: DB(4) = 56:  DB(5) = 84  '*
    Case 43: DB(1) = 16:  DB(2) = 16:  DB(3) = 124: DB(4) = 16:  DB(5) = 16  '+
    Case 44: DB(1) = 0:   DB(2) = 128: DB(3) = 224: DB(4) = 96:  DB(5) = 0   ',
    Case 45: DB(1) = 8:   DB(2) = 8:   DB(3) = 8:   DB(4) = 8:   DB(5) = 0   '-
    Case 46: DB(1) = 0:   DB(2) = 96:  DB(3) = 96:  DB(4) = 0:   DB(5) = 0   '.
    Case 47: DB(1) = 64:  DB(2) = 32:  DB(3) = 16:  DB(4) = 8:   DB(5) = 4   '/
    Case 48: DB(1) = 127: DB(2) = 65:  DB(3) = 65:  DB(4) = 127: DB(5) = 0   '0
    Case 49: DB(1) = 64:  DB(2) = 68:  DB(3) = 127: DB(4) = 64:  DB(5) = 0   '1
    Case 50: DB(1) = 121: DB(2) = 73:  DB(3) = 73:  DB(4) = 79:  DB(5) = 0   '2
    Case 51: DB(1) = 65:  DB(2) = 73:  DB(3) = 73:  DB(4) = 127: DB(5) = 0   '3
    Case 52: DB(1) = 15:  DB(2) = 8:   DB(3) = 8:   DB(4) = 127: DB(5) = 0   '4
    Case 53: DB(1) = 79:  DB(2) = 73:  DB(3) = 73:  DB(4) = 121: DB(5) = 0   '5
    Case 54: DB(1) = 127: DB(2) = 72:  DB(3) = 72:  DB(4) = 120: DB(5) = 0   '6
    Case 55: DB(1) = 1:   DB(2) = 1:   DB(3) = 1:   DB(4) = 127: DB(5) = 0   '7
    Case 56: DB(1) = 127: DB(2) = 73:  DB(3) = 73:  DB(4) = 127: DB(5) = 0   '8
    Case 57: DB(1) = 79:  DB(2) = 73:  DB(3) = 73:  DB(4) = 127: DB(5) = 0   '9
    Case 58: DB(1) = 0:   DB(2) = 108: DB(3) = 108: DB(4) = 0:   DB(5) = 0   ':
    Case 59: DB(1) = 0:   DB(2) = 128: DB(3) = 236: DB(4) = 108: DB(5) = 0   ';
    Case 60: DB(1) = 8:   DB(2) = 20:  DB(3) = 34:  DB(4) = 65:  DB(5) = 0   '<
    Case 61: DB(1) = 36:  DB(2) = 36:  DB(3) = 36:  DB(4) = 36:  DB(5) = 0   '=
    Case 62: DB(1) = 65:  DB(2) = 34:  DB(3) = 20:  DB(4) = 8:   DB(5) = 0   '>
    Case 63: DB(1) = 3:   DB(2) = 81:  DB(3) = 91:  DB(4) = 14:  DB(5) = 0   '?
    Case 64: DB(1) = 62:  DB(2) = 65:  DB(3) = 89:  DB(4) = 94:  DB(5) = 0   '@
    Case 65: DB(1) = 127: DB(2) = 9:   DB(3) = 9:   DB(4) = 127: DB(5) = 0   'A
    Case 66: DB(1) = 127: DB(2) = 73:  DB(3) = 73:  DB(4) = 54:  DB(5) = 0   'B
    Case 67: DB(1) = 127: DB(2) = 65:  DB(3) = 65:  DB(4) = 99:  DB(5) = 0   'C
    Case 68: DB(1) = 127: DB(2) = 65:  DB(3) = 65:  DB(4) = 62:  DB(5) = 0   'D
    Case 69: DB(1) = 127: DB(2) = 73:  DB(3) = 73:  DB(4) = 65:  DB(5) = 0   'E
    Case 70: DB(1) = 127: DB(2) = 9:   DB(3) = 9:   DB(4) = 1:   DB(5) = 0   'F
    Case 71: DB(1) = 127: DB(2) = 65:  DB(3) = 73:  DB(4) = 121: DB(5) = 0   'G
    Case 72: DB(1) = 127: DB(2) = 8:   DB(3) = 8:   DB(4) = 127: DB(5) = 0   'H
    Case 73: DB(1) = 65:  DB(2) = 127: DB(3) = 65:  DB(4) = 65:  DB(5) = 0   'I
    Case 74: DB(1) = 65:  DB(2) = 65:  DB(3) = 127: DB(4) = 1:   DB(5) = 0   'J
    Case 75: DB(1) = 127: DB(2) = 8:   DB(3) = 20:  DB(4) = 99:  DB(5) = 0   'K
    Case 76: DB(1) = 127: DB(2) = 64:  DB(3) = 64:  DB(4) = 64:  DB(5) = 0   'L
    Case 77: DB(1) = 127: DB(2) = 2:   DB(3) = 2:   DB(4) = 127: DB(5) = 0   'M
    Case 78: DB(1) = 127: DB(2) = 14:  DB(3) = 56:  DB(4) = 127: DB(5) = 0   'N
    Case 79: DB(1) = 127: DB(2) = 65:  DB(3) = 65:  DB(4) = 127: DB(5) = 0   'O
    Case 80: DB(1) = 127: DB(2) = 9:   DB(3) = 9:   DB(4) = 15:  DB(5) = 0   'P
    Case 81: DB(1) = 127: DB(2) = 65:  DB(3) = 97:  DB(4) = 255: DB(5) = 0   'Q
    Case 82: DB(1) = 127: DB(2) = 9:   DB(3) = 9:   DB(4) = 119: DB(5) = 0   'R
    Case 83: DB(1) = 79:  DB(2) = 73:  DB(3) = 73:  DB(4) = 121: DB(5) = 0   'S
    Case 84: DB(1) = 1:   DB(2) = 127: DB(3) = 1:   DB(4) = 1:   DB(5) = 0   'T
    Case 85: DB(1) = 127: DB(2) = 64:  DB(3) = 64:  DB(4) = 127: DB(5) = 0   'U
    Case 86: DB(1) = 63:  DB(2) = 64:  DB(3) = 48:  DB(4) = 15:  DB(5) = 0   'V
    Case 87: DB(1) = 127: DB(2) = 32:  DB(3) = 32:  DB(4) = 127: DB(5) = 0   'W
    Case 88: DB(1) = 115: DB(2) = 12:  DB(3) = 24:  DB(4) = 103: DB(5) = 0   'X
    Case 89: DB(1) = 79:  DB(2) = 72:  DB(3) = 72:  DB(4) = 127: DB(5) = 0   'Y
    Case 90: DB(1) = 97:  DB(2) = 89:  DB(3) = 77:  DB(4) = 67:  DB(5) = 0   'Z
  End Select
  GetDigitalCharBits = DB(col)
End Function

Private Function GetDigitalSmallCharBits(ByVal Char As Long, ByVal col As Long) As Long
  Dim DB(1 To 5) As Byte
  
  Select Case Char
    Case 32: DB(1) = 0:  DB(2) = 0:  DB(3) = 0:  DB(4) = 0:  DB(5) = 0   'SPACE
    Case 33: DB(1) = 0:  DB(2) = 23: DB(3) = 23: DB(4) = 0:  DB(5) = 0   '!
    Case 34: DB(1) = 0:  DB(2) = 3:  DB(3) = 0:  DB(4) = 3:  DB(5) = 0   '"
    Case 35: DB(1) = 10: DB(2) = 31: DB(3) = 10: DB(4) = 31: DB(5) = 10  '#
    Case 36: DB(1) = 0:  DB(2) = 10: DB(3) = 31: DB(4) = 10: DB(5) = 0   '$
    Case 37: DB(1) = 19: DB(2) = 11: DB(3) = 4:  DB(4) = 26: DB(5) = 25  '%
    Case 38: DB(1) = 7:  DB(2) = 29: DB(3) = 31: DB(4) = 20: DB(5) = 0   '&
    Case 39: DB(1) = 0:  DB(2) = 0:  DB(3) = 3:  DB(4) = 0:  DB(5) = 0   ''
    Case 40: DB(1) = 0:  DB(2) = 14: DB(3) = 17: DB(4) = 0:  DB(5) = 0   '(
    Case 41: DB(1) = 0:  DB(2) = 17: DB(3) = 14: DB(4) = 0:  DB(5) = 0   ')
    Case 42: DB(1) = 0:  DB(2) = 10: DB(3) = 4:  DB(4) = 10: DB(5) = 0   '*
    Case 43: DB(1) = 0:  DB(2) = 4:  DB(3) = 14: DB(4) = 4:  DB(5) = 0   '+
    Case 44: DB(1) = 0:  DB(2) = 16: DB(3) = 8:  DB(4) = 0:  DB(5) = 0   ',
    Case 45: DB(1) = 0:  DB(2) = 4:  DB(3) = 4:  DB(4) = 4:  DB(5) = 0   '-
    Case 46: DB(1) = 0:  DB(2) = 16: DB(3) = 16: DB(4) = 0:  DB(5) = 0   '.
    Case 47: DB(1) = 16: DB(2) = 8:  DB(3) = 4:  DB(4) = 2:  DB(5) = 1   '/
    Case 48: DB(1) = 14: DB(2) = 17: DB(3) = 17: DB(4) = 14: DB(5) = 0   '0
    Case 49: DB(1) = 0:  DB(2) = 18: DB(3) = 31: DB(4) = 16: DB(5) = 0   '1
    Case 50: DB(1) = 29: DB(2) = 21: DB(3) = 21: DB(4) = 23: DB(5) = 0   '2
    Case 51: DB(1) = 17: DB(2) = 21: DB(3) = 21: DB(4) = 31: DB(5) = 0   '3
    Case 52: DB(1) = 7:  DB(2) = 4:  DB(3) = 4:  DB(4) = 31: DB(5) = 0   '4
    Case 53: DB(1) = 23: DB(2) = 21: DB(3) = 21: DB(4) = 29: DB(5) = 0   '5
    Case 54: DB(1) = 31: DB(2) = 21: DB(3) = 21: DB(4) = 29: DB(5) = 0   '6
    Case 55: DB(1) = 1:  DB(2) = 1:  DB(3) = 1:  DB(4) = 31: DB(5) = 0   '7
    Case 56: DB(1) = 31: DB(2) = 21: DB(3) = 21: DB(4) = 31: DB(5) = 0   '8
    Case 57: DB(1) = 23: DB(2) = 21: DB(3) = 21: DB(4) = 31: DB(5) = 0   '9
    Case 58: DB(1) = 0:  DB(2) = 0:  DB(3) = 10: DB(4) = 0:  DB(5) = 0   ':
    Case 59: DB(1) = 0:  DB(2) = 16: DB(3) = 10: DB(4) = 0:  DB(5) = 0   ';
    Case 60: DB(1) = 0:  DB(2) = 4:  DB(3) = 10: DB(4) = 17: DB(5) = 0   '<
    Case 61: DB(1) = 0:  DB(2) = 10: DB(3) = 10: DB(4) = 10: DB(5) = 0   '=
    Case 62: DB(1) = 0:  DB(2) = 17: DB(3) = 10: DB(4) = 4:  DB(5) = 0   '>
    Case 63: DB(1) = 0:  DB(2) = 1:  DB(3) = 21: DB(4) = 2:  DB(5) = 0   '?
    Case 64: DB(1) = 14: DB(2) = 17: DB(3) = 21: DB(4) = 22: DB(5) = 0   '@
    Case 65: DB(1) = 30: DB(2) = 5:  DB(3) = 5:  DB(4) = 30: DB(5) = 0   'A
    Case 66: DB(1) = 31: DB(2) = 21: DB(3) = 21: DB(4) = 10: DB(5) = 0   'B
    Case 67: DB(1) = 31: DB(2) = 17: DB(3) = 17: DB(4) = 17: DB(5) = 0   'C
    Case 68: DB(1) = 31: DB(2) = 17: DB(3) = 17: DB(4) = 14: DB(5) = 0   'D
    Case 69: DB(1) = 31: DB(2) = 21: DB(3) = 21: DB(4) = 17: DB(5) = 0   'E
    Case 70: DB(1) = 31: DB(2) = 5:  DB(3) = 5:  DB(4) = 1:  DB(5) = 0   'F
    Case 71: DB(1) = 31: DB(2) = 17: DB(3) = 21: DB(4) = 29: DB(5) = 0   'G
    Case 72: DB(1) = 31: DB(2) = 4:  DB(3) = 4:  DB(4) = 31: DB(5) = 0   'H
    Case 73: DB(1) = 0:  DB(2) = 17: DB(3) = 31: DB(4) = 17: DB(5) = 0   'I
    Case 74: DB(1) = 24: DB(2) = 17: DB(3) = 31: DB(4) = 1:  DB(5) = 0   'J
    Case 75: DB(1) = 31: DB(2) = 4:  DB(3) = 10: DB(4) = 17: DB(5) = 0   'K
    Case 76: DB(1) = 31: DB(2) = 16: DB(3) = 16: DB(4) = 16: DB(5) = 0   'L
    Case 77: DB(1) = 31: DB(2) = 2:  DB(3) = 2:  DB(4) = 31: DB(5) = 0   'M
    Case 78: DB(1) = 31: DB(2) = 2:  DB(3) = 12: DB(4) = 31: DB(5) = 0   'N
    Case 79: DB(1) = 14: DB(2) = 17: DB(3) = 17: DB(4) = 14: DB(5) = 0   'O
    Case 80: DB(1) = 31: DB(2) = 5:  DB(3) = 5:  DB(4) = 7:  DB(5) = 0   'P
    Case 81: DB(1) = 31: DB(2) = 17: DB(3) = 25: DB(4) = 31: DB(5) = 0   'Q
    Case 82: DB(1) = 31: DB(2) = 5:  DB(3) = 5:  DB(4) = 27: DB(5) = 0   'R
    Case 83: DB(1) = 23: DB(2) = 21: DB(3) = 21: DB(4) = 29: DB(5) = 0   'S
    Case 84: DB(1) = 1:  DB(2) = 1:  DB(3) = 31: DB(4) = 1:  DB(5) = 0   'T
    Case 85: DB(1) = 31: DB(2) = 16: DB(3) = 16: DB(4) = 31: DB(5) = 0   'U
    Case 86: DB(1) = 15: DB(2) = 16: DB(3) = 12: DB(4) = 3:  DB(5) = 0   'V
    Case 87: DB(1) = 31: DB(2) = 8:  DB(3) = 8:  DB(4) = 31: DB(5) = 0   'W
    Case 88: DB(1) = 27: DB(2) = 4:  DB(3) = 4:  DB(4) = 27: DB(5) = 0   'X
    Case 89: DB(1) = 1:  DB(2) = 18: DB(3) = 12: DB(4) = 7:  DB(5) = 0   'Y
    Case 90: DB(1) = 25: DB(2) = 21: DB(3) = 21: DB(4) = 19: DB(5) = 0   'Z
  End Select
  GetDigitalSmallCharBits = DB(col)
End Function

