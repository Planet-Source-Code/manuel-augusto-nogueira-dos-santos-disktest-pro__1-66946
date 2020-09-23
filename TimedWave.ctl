VERSION 5.00
Begin VB.UserControl TimedWave 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ForeColor       =   &H00C0FFFF&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox DC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   0
      Top             =   1.50000e5
      Width           =   2055
   End
End
Attribute VB_Name = "TimedWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------Windows API Functions
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'---------------------------------------------------Private
Private mList As New Collection
Private vMin As Long, pMin As Long
Private vMax As Long, pMax As Long
Private CurPos As Long
Private PreviousTick As Long

Private Const SRCCOPY = &HCC0020
'Default Property Values:
Private Const m_def_FromZero = False
'Property Variables:
Private mFromZero As Boolean

'----------------------------------------------------Events
Public Event Added()

'------------------------------------------------Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
  DrawWidth = UserControl.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
  UserControl.DrawWidth() = New_DrawWidth
  PropertyChanged "DrawWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  UserControl.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FromZero() As Boolean
  FromZero = mFromZero
End Property

Public Property Let FromZero(ByVal New_FromZero As Boolean)
  mFromZero = New_FromZero
  PropertyChanged "FromZero"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000C)
  UserControl.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
  mFromZero = PropBag.ReadProperty("FromZero", m_def_FromZero)
End Sub

Private Sub UserControl_Resize()
  DC.Width = UserControl.ScaleWidth
  DC.Height = UserControl.ScaleHeight
  Do While mList.Count > UserControl.ScaleWidth
    mList.Remove 1
  Loop
  Call DisplayTimedWave
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000C)
  Call PropBag.WriteProperty("DrawWidth", UserControl.DrawWidth, 1)
  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
  Call PropBag.WriteProperty("FromZero", mFromZero, m_def_FromZero)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  mFromZero = m_def_FromZero
End Sub

'---------------------------------------------------Methods
Public Sub Clear()
  UserControl.Cls
  Do While mList.Count > 0
    mList.Remove 1
  Loop
  vMin = 0: pMin = 0
  vMax = 0: pMax = 0
  CurPos = 0
  PreviousTick = GetTickCount()
End Sub

Public Sub Add()
  Dim CurTick As Long
  Dim Value As Long
  
  'get time
  CurTick = GetTickCount()
  'if window full, remove left
  If mList.Count = UserControl.ScaleWidth Then
    mList.Remove 1
    pMin = pMin - 1
    pMax = pMax - 1
    If (pMin = 0) Or (pMax = 0) Then
      Call CalcMinMax
    End If
  End If
  'get part value
  Value = CurTick - PreviousTick
  'add value
  mList.Add Value
  'check min & max
  If (Value < vMin) Or (pMin = 0) Then
    vMin = Value
    pMin = mList.Count
  End If
  If (Value > vMax) Or (pMax = 0) Then
    vMax = Value
    pMax = mList.Count
  End If
  'display wave
  Call DisplayTimedWave
  PreviousTick = CurTick
End Sub

'-------------------------------------------Private Methods
Private Sub CalcMinMax()
  Dim Value As Long
  Dim i As Long
  
  pMin = 0
  pMax = 0
  For i = 1 To mList.Count
    Value = mList(i)
    If (Value < vMin) Or (pMin = 0) Then
      vMin = Value
      pMin = i
    End If
    If (Value > vMax) Or (pMax = 0) Then
      vMax = Value
      pMax = i
    End If
  Next i
End Sub

Private Sub DisplayTimedWave()
  Dim i As Long
  Dim Value As Long
  Dim PosY As Long
  Dim oY As Long
  Dim Elems As Long
  
  DC.BackColor = Me.BackColor
  DC.ForeColor = Me.ForeColor
  DC.DrawWidth = Me.DrawWidth
  DC.Cls
  Elems = mList.Count
  PosY = 0
  For i = 1 To Elems
    Value = mList(i)
    If mFromZero = False Then
      If vMax <> 0 Then PosY = UserControl.ScaleHeight - (Value * (UserControl.ScaleHeight - 1)) \ vMax
    Else
      If vMax <> 0 Then PosY = UserControl.ScaleHeight - ((Value - vMin) * (UserControl.ScaleHeight - 1)) \ vMax
    End If
    If i = 1 Then oY = PosY
    DC.Line (i - 1, oY)-(i, PosY)
    oY = PosY
  Next i
  
  'copy bitmap
  BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, DC.hDC, 0, 0, SRCCOPY
  UserControl.Refresh
  DoEvents
  RaiseEvent Added
End Sub



