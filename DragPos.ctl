VERSION 5.00
Begin VB.UserControl DragPos 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00595900&
   BackStyle       =   0  'Transparent
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   FillStyle       =   0  'Solid
   MaskColor       =   &H00000000&
   ScaleHeight     =   19
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ToolboxBitmap   =   "DragPos.ctx":0000
   Begin VB.Timer DragTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   900
      Top             =   -60
   End
End
Attribute VB_Name = "DragPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------Windows API functions
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoints As Any, ByVal nCount As Long) As Long
Private Declare Sub GetCursorPos Lib "user32" (lpPoint As Point)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Point)

'-----------------------------------------------Declarations
'private
Private m_Start As Boolean
Private DragCursor As Byte
Private mForm As Form
Private UserControlLeft As Long
Private UserControlTop As Long
'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_Positions = 2
Const m_def_StartPosition = 1
Const m_def_EndPosition = 2
'Property Variables:
Dim m_ForeColor As OLE_COLOR
Dim m_Picture As Picture
Dim m_Positions As Long
Dim m_StartPosition As Long
Dim m_EndPosition As Long
'Event Declarations:
Event Position()
Attribute Position.VB_Description = "Indicates that start position or end position has changed."

'---------------------------------------------Control Events
Private Sub DragTimer_Timer()
  Dim Mpos As Point
  Dim Wpos As Point
  Dim PosX As Long
  Dim Valor As Long
  Dim Bloco As Long
  
  If DragCursor = 0 Then Exit Sub
  'determina posição
  Call GetCursorPos(Mpos)
  Call ClientToScreen(UserControl.hWnd, Wpos)
  PosX = Mpos.X - Wpos.X
  'calcular tamanho e posicao
  Bloco = 1 + UserControl.ScaleWidth \ (m_Positions + 1)
  Valor = 1 + (PosX - Bloco \ 2) \ Bloco  'não optimizar divisão decimal
  'posicionar e desenhar cursores
  If DragCursor = 1 Then
    If m_StartPosition <> Valor Then
      Me.StartPosition = Valor
      RaiseEvent Position
    End If
  Else
    If m_EndPosition <> Valor Then
      Me.EndPosition = Valor
      RaiseEvent Position
    End If
  End If
End Sub

Private Sub UserControl_Initialize()
  m_Start = False
  DragCursor = 0
End Sub

Private Sub UserControl_Resize()
  Call DrawCursor
End Sub

'-------------------------------------------------Properties
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,2
Public Property Get Positions() As Long
Attribute Positions.VB_Description = "How many positions for the cursor."
    Positions = m_Positions
End Property

Public Property Let Positions(ByVal New_Positions As Long)
  If New_Positions > 1 Then
    m_Positions = New_Positions
    If m_StartPosition > m_Positions Then
      m_StartPosition = m_Positions
      PropertyChanged "StartPosition"
    End If
    If m_EndPosition > m_Positions Then
      m_EndPosition = m_Positions
      PropertyChanged "EndPosition"
    End If
    PropertyChanged "Positions"
    Call DrawCursor
  End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get StartPosition() As Long
Attribute StartPosition.VB_Description = "Start position of the cursor."
    StartPosition = m_StartPosition
End Property

Public Property Let StartPosition(ByVal New_StartPosition As Long)
  If (New_StartPosition > 0) And _
     (((New_StartPosition <= m_EndPosition) And (m_EndPosition <> 0)) Or (m_EndPosition = 0)) Then
    m_StartPosition = New_StartPosition
    PropertyChanged "StartPosition"
    Call DrawCursor
  End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,2
Public Property Get EndPosition() As Long
Attribute EndPosition.VB_Description = "Ending position. Set 0 if you don't want an ending position."
    EndPosition = m_EndPosition
End Property

Public Property Let EndPosition(ByVal New_EndPosition As Long)
  If ((New_EndPosition <= m_Positions) And (New_EndPosition >= m_StartPosition)) Or _
     (m_EndPosition = 0) Then
    m_EndPosition = New_EndPosition
    PropertyChanged "EndPosition"
    Call DrawCursor
  End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'--------------------------------------------Private Methods
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Positions = m_def_Positions
    m_StartPosition = m_def_StartPosition
    m_EndPosition = m_def_EndPosition
    Set m_Picture = LoadPicture("")
    m_ForeColor = m_def_ForeColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Positions = PropBag.ReadProperty("Positions", m_def_Positions)
    m_StartPosition = PropBag.ReadProperty("StartPosition", m_def_StartPosition)
    m_EndPosition = PropBag.ReadProperty("EndPosition", m_def_EndPosition)
    m_Start = True
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Positions", m_Positions, m_def_Positions)
    Call PropBag.WriteProperty("StartPosition", m_StartPosition, m_def_StartPosition)
    Call PropBag.WriteProperty("EndPosition", m_EndPosition, m_def_EndPosition)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
End Sub

'---------------------------------------------Public Methods
Public Sub DrawCursor()
  Dim Bloco As Long
  Dim Parte As Long
  Dim PolP(0 To 2) As Point
  
  If m_Start = False Then Exit Sub
  If mForm Is Nothing Then Exit Sub
  mForm.PaintPicture m_Picture, UserControlLeft, UserControlTop, , , 0, 0
  mForm.ForeColor = m_ForeColor
  mForm.FillColor = m_ForeColor
  mForm.FillStyle = 0
  'calcular tamanho
  Bloco = 1 + UserControl.ScaleWidth \ (m_Positions + 1)
  Parte = Bloco \ 2
  'posicionar e desenhar cursores
  mForm.StartCursor.Height = UserControl.ScaleHeight
  mForm.EndCursor.Height = UserControl.ScaleHeight
  mForm.StartCursor.Width = Bloco + Parte
  mForm.EndCursor.Width = Bloco + Parte
  mForm.StartCursor.Left = m_StartPosition * Bloco - Parte
  mForm.EndCursor.Left = m_EndPosition * Bloco - Parte
  PolP(0).X = UserControlLeft + mForm.StartCursor.Left - 2
  PolP(0).Y = UserControlTop
  PolP(1).X = UserControlLeft + mForm.StartCursor.Left + Bloco - 2
  PolP(1).Y = UserControlTop
  PolP(2).X = UserControlLeft + mForm.StartCursor.Left + Parte - 2
  PolP(2).Y = UserControlTop + mForm.StartCursor.Height - 1
  Call Polygon(mForm.hDC, PolP(0), 3)
  PolP(0).X = UserControlLeft + mForm.EndCursor.Left
  PolP(0).Y = UserControlTop
  PolP(1).X = UserControlLeft + mForm.EndCursor.Left + Bloco
  PolP(1).Y = UserControlTop
  PolP(2).X = UserControlLeft + mForm.EndCursor.Left + Parte
  PolP(2).Y = UserControlTop + mForm.EndCursor.Height - 1
  Call Polygon(mForm.hDC, PolP(0), 3)
  mForm.Refresh
End Sub

Public Sub SetForm(ByRef myForm As Form, ByVal Left As Long, ByVal Top As Long)
  Set mForm = myForm
  UserControlLeft = Left
  UserControlTop = Top
End Sub

Public Sub MouseOp(ByVal modo As Long)
  Select Case modo
    Case 1:
      If m_Start = False Then Exit Sub
      DragCursor = 1
      DragTimer.Enabled = True
    Case 2:
      If m_Start = False Then Exit Sub
      If m_StartPosition = m_EndPosition Then
        DragCursor = 1
      Else
        DragCursor = 2
      End If
      DragTimer.Enabled = True
    Case 0:
      DragCursor = 0
      DragTimer.Enabled = False
  End Select
End Sub
