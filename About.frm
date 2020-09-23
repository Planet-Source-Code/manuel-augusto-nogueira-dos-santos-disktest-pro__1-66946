VERSION 5.00
Begin VB.Form About 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3510
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   328
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicFocus 
      Height          =   195
      Left            =   300
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   75000
      Width           =   195
   End
   Begin VB.PictureBox PicScale 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00C0FFFF&
      Height          =   2730
      Left            =   0
      ScaleHeight     =   182
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   328
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   75000
      Width           =   4920
   End
   Begin VB.PictureBox PicText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00C0FFFF&
      Height          =   2730
      Left            =   0
      ScaleHeight     =   182
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   324
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   75000
      Width           =   4860
   End
   Begin VB.PictureBox PicField 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   0
      ScaleHeight     =   182
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   328
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   4920
   End
   Begin VB.PictureBox PicScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   660
      Picture         =   "About.frx":000C
      ScaleHeight     =   182
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   328
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75000
      Width           =   4920
   End
   Begin VB.PictureBox PicMask2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3510
      Left            =   480
      Picture         =   "About.frx":02B7
      ScaleHeight     =   234
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   328
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   75000
      Width           =   4920
   End
   Begin VB.PictureBox PicStar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   0
      ScaleHeight     =   182
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   328
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   75000
      Width           =   4920
   End
   Begin VB.PictureBox PicMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3510
      Left            =   1260
      Picture         =   "About.frx":09D0
      ScaleHeight     =   3510
      ScaleWidth      =   4920
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   75000
      Width           =   4920
   End
   Begin VB.PictureBox PicAbout 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3510
      Left            =   300
      Picture         =   "About.frx":0C7B
      ScaleHeight     =   3510
      ScaleWidth      =   4920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2.25000e5
      Width           =   4920
   End
   Begin VB.Image AboutOK 
      Height          =   300
      Left            =   3420
      Top             =   3120
      Width           =   1305
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------Windows API Functions
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseWinCapture Lib "user32" Alias "ReleaseCapture" () As Long
Private Declare Function SendWinMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'------------------------------------------------------Const
Private Const WM_MOVE = &HF012
Private Const WM_SYSCOMMAND = &H112
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const RGN_OR = 2
Private Const SRCCOPY = &HCC0020
Private Const SRCPAINT = &HEE0086
Private Const SRCAND = &H8800C6

'------------------------------------------------------Stars
Private Const MaxStars = 30
Private Const ZFactor = 200

Private Type StarRec
  X As Long
  Y As Long
  Z As Long
  Speed As Long
End Type

Private Stars(0 To MaxStars) As StarRec
Private StopStar As Boolean

'-------------------------------------------------------Text
Private Const MaxText = 7
Private Const DT_CENTER = &H1

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private TextMov As Long
Private TextDelay As Long
Private TextNum As Long

'-----------------------------------------------------Events
Private Sub AboutOK_Click()
  StopStar = True
  Me.Hide
End Sub

Private Sub PicFocus_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    StopStar = True
    Me.Hide
  End If
End Sub

Private Sub Form_Load()
  Dim wndReg As Long
  Dim wndRet As Long
  Dim i As Long
  
  'Set Shaped form
  wndReg = RgnFromMask(PicMask, RGB(255, 255, 255))
  wndRet = SetWindowRgn(Me.hWnd, wndReg, True)
  'init stars
  For i = 0 To MaxStars: Call NewStar(i): Next i
  StopStar = False
  'init text
  TextMov = 0
  TextDelay = 0
  TextNum = 0
  'paint background
  BitBlt Me.hDC, 0, 0, 328, 234, PicAbout.hDC, 0, 0, SRCCOPY
  'set positions
  PicField.Top = 0: PicField.Left = 0
  PicField.Height = 182: PicField.Width = 328
  AboutOK.Left = 228: AboutOK.Top = 208
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim res As Long
  
  res = ReleaseWinCapture()
  'res = SendWinMessage(Me.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)  'does not work on NT
  res = SendWinMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  PicFocus.SetFocus
End Sub

Private Sub Form_Resize()
  PicFocus.SetFocus
  Do
    Call DoStars
    Call BitBlt(PicScreen.hDC, 0, 0, 328, 182, PicStar.hDC, 0, 0, SRCCOPY)
    Call DoText
    Call BitBlt(PicScreen.hDC, 0, 0, 328, 182, PicText.hDC, 0, 0, SRCPAINT)
    Call BitBlt(PicScreen.hDC, 0, 0, 328, 182, PicMask2.hDC, 0, 0, SRCAND)
    Call BitBlt(PicScreen.hDC, 0, 0, 328, 182, PicAbout.hDC, 0, 0, SRCPAINT)
    Call BitBlt(PicField.hDC, 0, 0, 328, 182, PicScreen.hDC, 0, 0, SRCCOPY)
    PicField.Refresh
    DoEvents
  Loop Until StopStar = True
End Sub

'---------------------------------------------RegionFromMask
Private Function RgnFromMask(PicMask As PictureBox, Optional lngTransColor As Long = -1) As Long
  Dim wndRgn As Long, wndRgnTmp As Long, wndRgnAux As Long
  Dim pX As Long, pY As Long
  Dim tX As Long, tY As Long
  Dim pixVal As Long
  Dim rX1 As Long, rX2 As Long
  
  If lngTransColor = -1 Then lngTransColor = RGB(255, 255, 255)
  wndRgn = 0
  tY = 90
  tX = PicMask.Width
  'accelerate
  wndRgn = CreateRectRgn(1, 90, tX + 1, 235)
  'get mask pixels
  For pY = 1 To tY
    pX = 1
    Do While pX <= tX
      Do While (GetPixel(PicMask.hDC, pX - 1, pY - 1) = lngTransColor) And (pX <= tX)
        pX = pX + 1
      Loop
      If pX <= tX Then
        rX1 = pX
        Do While (GetPixel(PicMask.hDC, pX - 1, pY - 1) <> lngTransColor) And (pX <= tX)
          pX = pX + 1
        Loop
        rX2 = pX - 1
        wndRgnTmp = CreateRectRgn(rX1, pY, rX2 + 1, pY + 1)
        wndRgnAux = CombineRgn(wndRgn, wndRgn, wndRgnTmp, RGN_OR)
        Call DeleteObject(wndRgnTmp)
      End If
    Loop
  Next pY
  RgnFromMask = wndRgn
End Function

'-----------------------------------------------------Stars
Private Sub NewStar(ByVal num As Long)
 Stars(num).X = Rnd * 100 - 50
 Stars(num).Y = Rnd * 100 - 50
 Stars(num).Z = Rnd * 100 + 200
 Stars(num).Speed = 1
End Sub

Private Function StarColor(ByVal Z As Long) As Long
  Dim Value As Long
  
  Value = 5 + (Z / 150)
  If Value > 100 Then Value = 100
  Value = Value + 150
  StarColor = RGB(Value, Value, Value)
End Function

Private Sub DoStars()
  Dim X As Long, Y As Long
  Dim i As Long
  
  For i = 0 To MaxStars
    'old star pos : OFF
    X = 164 + Round(Stars(i).X * Stars(i).Z / ZFactor)
    Y = 95 + Round(Stars(i).Y * Stars(i).Z / ZFactor)
    PicStar.PSet (X, Y), 0
    'calculate new pos
    Stars(i).Z = Stars(i).Z + Stars(i).Speed
    If Stars(i).Z > 20000 Then NewStar i
    Stars(i).Speed = (Stars(i).Z / 32) * (5 - (Abs(Stars(i).X * Stars(i).Y) / 500))
    If Stars(i).Speed = 0 Then Stars(i).Speed = 1
    If (X < 0) Or (X > 328) Or (Y < 0) Or (Y > 190) Then NewStar i
    'new star pos : ON
    X = 164 + Round(Stars(i).X * (Stars(i).Z + Stars(i).Speed) / ZFactor)
    Y = 95 + Round(Stars(i).Y * (Stars(i).Z + Stars(i).Speed) / ZFactor)
    PicStar.PSet (X, Y), StarColor(Stars(i).Z)
  Next i
End Sub

'------------------------------------------------------Text
Private Sub DoText()
  Select Case TextMov
    Case 0: 'change text
      TextNum = TextNum + 1
      If TextNum > MaxText Then TextNum = 1
      TextDelay = 0
      PicText.Line (0, 0)-(328, 182), 0, BF
      TextMov = 1
    Case 1: 'wait for apearance
      TextDelay = TextDelay + 1
      If TextDelay = 50 Then
        TextMov = 2
        TextDelay = 0
      End If
    Case 2: 'text appears, moving closer
      TextDelay = TextDelay + 4
      Call DrawPicText
      If TextDelay = 100 Then
        TextMov = 3
        TextDelay = 0
      End If
    Case 3: 'Text 100%, wait for leaving
      TextDelay = TextDelay + 1
      If TextDelay = 50 Then
        TextMov = 4
        TextDelay = 0
      End If
    Case 4: 'text leaving
      TextDelay = TextDelay + 1
      Call DrawPicText
      If TextDelay = 50 Then
        TextMov = 0
        TextDelay = 0
      End If
  End Select
End Sub

Private Sub DrawPicText()
  Dim strTXT As String
  Dim R As RECT
  Dim pX As Long, pY As Long
  Dim tX As Long, tY As Long
  
  PicText.Line (0, 0)-(328, 182), 0, BF
  PicScale.Line (0, 0)-(328, 182), 0, BF
  'draw text in normal size
  Select Case TextNum
    Case 1: strTXT = "DESIGN"
    Case 2: strTXT = "MANUEL AUGUSTO SANTOS"
    Case 3: strTXT = "PROGRAMMING"
    Case 4: strTXT = "MANUEL AUGUSTO SANTOS"
    Case 5: strTXT = "(c) 2000-2003"
    Case 6: strTXT = "TEALTECH"
    Case 7: strTXT = "THE ONLY WAY IS UP"
  End Select
  Call SetRect(R, 0, 84, 328, 182)
  Call DrawText(PicScale.hDC, strTXT, Len(strTXT), R, DT_CENTER)
  PicScale.Refresh
  'shrink
  If TextMov = 2 Then
    tX = (TextDelay * 328) / 200
    tY = (TextDelay * 182) / 200
    Call StretchBlt(PicText.hDC, 164 - tX, 96 - tY, tX * 2, tY * 2, PicScale.hDC, 0, 0, 328, 182, SRCCOPY)
  End If
  'expand
  If TextMov = 4 Then
    Select Case TextNum
      Case 2, 4, 6:
        tX = ((100 + 50 * TextDelay) * 328) / 200
        tY = ((100 + 50 * TextDelay) * 182) / 200
      Case 1, 3, 5:
        tX = ((100 - 2 * TextDelay) * 328) / 200
        tY = ((100 - 2 * TextDelay) * 182) / 200
    End Select
    Call StretchBlt(PicText.hDC, 164 - tX, 96 - tY, tX * 2, tY * 2, PicScale.hDC, 0, 0, 328, 182, SRCCOPY)
  End If
  PicText.Refresh
End Sub

