VERSION 5.00
Begin VB.UserControl ShapedButton 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   ScaleHeight     =   2505
   ScaleWidth      =   2625
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   360
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer tmrMousePos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   960
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "ShapedButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'0101010101010101010101010101010101010101010101010101010101010
'1                   In the name of God                      1
'0                 Author: Mostafa Asgari                    0
'1                     Country: Iran                         1
'1101010101010101010101010101010101010101010101010101010101010
'-------------------------------------------------------------
'You have a royalty free right to use, reproduce, modify, _
 publish and mess with this code but PLEASE VOTE!
Option Explicit
'---------------------
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
  ByVal X As Long, ByVal Y As Long) As Long
'---------------------
Private Declare Function GetCursorPos Lib "user32" _
  (lpPoint As POINTAPI) As Long
'--------------------
Private Declare Function SetRect Lib "user32" (lpRect As RECT, _
  ByVal X1 As Long, ByVal Y1 As Long, _
  ByVal X2 As Long, ByVal Y2 As Long) As Long
'--------------------
Private Declare Function GetClientRect Lib "user32" _
  (ByVal Hwnd As Long, lpRect As RECT) As Long
'--------------------
Private Declare Function ClientToScreen Lib "user32" _
  (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
'-------------------
Private Type POINTAPI
  X As Long
  Y As Long
End Type
'------------------
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
'-----------------
'Picture on mouse leave
Private m_PictureOnML  As StdPicture
'Picture on mouse move
Private m_PictureOnMM  As StdPicture
'Picture on mouse down
Private m_PictureOnMD  As StdPicture
Private m_bShowPicture As Boolean
Private m_bMouseDown   As Boolean
'-----------------
Public Event DblClick()
Public Event Click()
Public Event MouseLeave()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private Sub ShowCaptionInCenter()
  '
  With lblCaption
    '
    .Move (ScaleWidth / 2) - .Width / 2, _
          (ScaleHeight / 2) - .Height / 2
    '
  End With
  '
End Sub
'=============================================================
Private Function GetClientRectOnScreen(ByVal Hwnd As Long) As RECT
  '
  Dim udtRect    As RECT
  Dim UpperLeft  As POINTAPI
  Dim LowerRight As POINTAPI
  '
  GetClientRect Hwnd, udtRect
  '
  With udtRect
    '
    UpperLeft.X = udtRect.Left
    UpperLeft.Y = udtRect.Top
    LowerRight.X = udtRect.Right
    LowerRight.Y = udtRect.Bottom
    '
  End With
  '
  ClientToScreen Hwnd, UpperLeft
  ClientToScreen Hwnd, LowerRight
  '
  SetRect udtRect, UpperLeft.X, UpperLeft.Y, LowerRight.X, LowerRight.Y
  '
  GetClientRectOnScreen = udtRect
  '
End Function
'=============================================================
Private Sub SetMask(ByVal SrcPicture As StdPicture, ByVal ColorKey As OLE_COLOR)
  '
  Set MaskPicture = SrcPicture
  MaskColor = ColorKey
  Refresh
  '
End Sub
'=============================================================
Private Function StretchPicture(ByVal SrcPicture As StdPicture, _
                                ByVal sngWidth As Single, _
                                ByVal sngHeight As Single) As StdPicture
  '-------------------
  '
  Set picTemp.Picture = SrcPicture
  '
  With picImage
    '
    .Cls
    .Move 0, 0, sngWidth, sngHeight
    .PaintPicture SrcPicture, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight
    Set StretchPicture = .Image
    '
  End With
  '
End Function
'=============================================================
Public Property Get ShowPicture() As Boolean
  '
  ShowPicture = m_bShowPicture
  '
End Property
'=============================================================
Public Property Let ShowPicture(ByVal bShowPicture As Boolean)
  '
  m_bShowPicture = bShowPicture
  '
  If (m_PictureOnML.Handle <> 0) And (m_bShowPicture = True) And (Picture.Handle = 0) Then
    '
    Set Picture = StretchPicture(m_PictureOnML, Width, Height)
    '
  Else
    '
    Picture = LoadPicture("")
    '
  End If
  '
  PropertyChanged "ShowPicture"
  '
End Property
'=============================================================
Public Property Get PictureOnMouseLeave() As StdPicture
  '
  Set PictureOnMouseLeave = m_PictureOnML
  '
End Property
'=============================================================
Public Property Set PictureOnMouseLeave(ByVal NewPicture As IPictureDisp)
  '
  Dim StretchPic As StdPicture
  '
  If NewPicture Is Nothing Then
    '
    Set m_PictureOnML = LoadPicture("")
    '
    SetMask m_PictureOnML, UserControl.MaskColor
    '
    Picture = LoadPicture("")
    '
  End If
  '
  'If it is a picture then
  If Not (NewPicture Is Nothing) Then
    '
    Set m_PictureOnML = NewPicture
    '
    Set StretchPic = StretchPicture(m_PictureOnML, Width, Height)
    '
    SetMask StretchPic, UserControl.MaskColor
    '
    If m_bShowPicture = True Then Set Picture = StretchPic
    '
  End If
  '
  Set StretchPic = Nothing
  '
  PropertyChanged "PictureOnMouseLeave"
  '
End Property
'=============================================================
Private Sub lblCaption_Click()
  '
  RaiseEvent Click
  '
End Sub
'=============================================================
Private Sub lblCaption_DblClick()
  '
  RaiseEvent DblClick
  '
End Sub
'=============================================================
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '
  RaiseEvent MouseDown(Button, Shift, X, Y)
  '
End Sub
'=============================================================
Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '
  RaiseEvent MouseMove(Button, Shift, X, Y)
  '
End Sub
'=============================================================
Private Sub tmrMousePos_Timer()
  '
  Dim MousePos As POINTAPI
  Dim udtRect  As RECT
  '
  GetCursorPos MousePos
  '
  udtRect = GetClientRectOnScreen(UserControl.Hwnd)
  '
  If PtInRect(udtRect, MousePos.X, MousePos.Y) <> 0 Then
    '
    If (m_bMouseDown = True) And (m_PictureOnMD.Handle <> 0) Then
      '
      SetMask StretchPicture(m_PictureOnMD, Width, Height), UserControl.MaskColor
      '
      If m_bShowPicture = True Then Set Picture = StretchPicture(m_PictureOnMD, Width, Height)
      '
    ElseIf (m_PictureOnMM.Handle <> 0) Then
      '
      SetMask StretchPicture(m_PictureOnMM, Width, Height), UserControl.MaskColor
      '
      If m_bShowPicture = True Then Set Picture = StretchPicture(m_PictureOnMM, Width, Height)
      '
    End If
    '
  Else
    '
    SetMask StretchPicture(m_PictureOnML, Width, Height), UserControl.MaskColor
    '
    If m_bShowPicture = True Then Set Picture = StretchPicture(m_PictureOnML, Width, Height)
    '
    tmrMousePos.Enabled = False
    '
    RaiseEvent MouseLeave
    '
  End If
  '
End Sub
'=============================================================
Private Sub UserControl_Click()
  '
  RaiseEvent Click
  '
End Sub
'=============================================================
Private Sub UserControl_DblClick()
  '
  RaiseEvent DblClick
  '
End Sub
'=============================================================
Private Sub UserControl_Initialize()
  '
  Set m_PictureOnML = New StdPicture
  Set m_PictureOnMM = New StdPicture
  Set m_PictureOnMD = New StdPicture
  lblCaption.Caption = ""
  '
End Sub
'=============================================================
Private Sub UserControl_InitProperties()
  '
  ShowPicture = True
  MaskColor = vbBlack
  BackColor = vbWhite
  ForeColor = vbBlack
  '
End Sub
'=============================================================
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '
  m_bMouseDown = True
  '
  RaiseEvent MouseDown(Button, Shift, X, Y)
  '
End Sub
'=============================================================
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '
  If tmrMousePos.Enabled = False Then tmrMousePos.Enabled = True
  '
  RaiseEvent MouseMove(Button, Shift, X, Y)
  '
End Sub
'=============================================================
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '
  m_bMouseDown = False
  '
  RaiseEvent MouseUp(Button, Shift, X, Y)
  '
End Sub
'=============================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  '
  Dim StretchPic As StdPicture
  '
  With PropBag
    '
    Set m_PictureOnML = .ReadProperty("PictureOnMouseLeave", LoadPicture(""))
    Set m_PictureOnMM = .ReadProperty("PictureOnMouseMove", LoadPicture(""))
    Set m_PictureOnMD = .ReadProperty("PictureOnMouseDown", LoadPicture(""))
    m_bShowPicture = .ReadProperty("ShowPicture", True)
    UserControl.MaskColor = .ReadProperty("MaskColor", vbBlack)
    UserControl.BackColor = .ReadProperty("BackColor", vbWhite)
    lblCaption.Caption = .ReadProperty("Caption", "")
    Set lblCaption.Font = .ReadProperty("Font", "MS Sans Serif")
    lblCaption.ForeColor = .ReadProperty("ForeColor", vbBlack)
    UserControl.Enabled = .ReadProperty("Enable", True)
    '
  End With
  '
  If (m_PictureOnML.Handle <> 0) Then
    '
    Set StretchPic = StretchPicture(m_PictureOnML, Width, Height)
    '
    SetMask StretchPic, UserControl.MaskColor
    '
    If m_bShowPicture = True Then Set Picture = StretchPic
    '
  End If
  '
  ShowCaptionInCenter
  '
  Set StretchPic = Nothing
  '
End Sub
'=============================================================
Private Sub UserControl_Resize()
  '
  Dim StretchPic As StdPicture
  '
  ShowCaptionInCenter
  '
  If m_PictureOnML.Handle = 0 Then Exit Sub
  '
  Set StretchPic = StretchPicture(m_PictureOnML, Width, Height)
  '
  SetMask StretchPic, UserControl.MaskColor
  '
  If m_bShowPicture = True Then Set Picture = StretchPic
  '
  Set StretchPic = Nothing
  '
  Refresh
  '
End Sub
'=============================================================
Private Sub UserControl_Terminate()
  '
  Set m_PictureOnML = Nothing
  Set m_PictureOnMM = Nothing
  Set m_PictureOnMD = Nothing
  '
End Sub
'=============================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  '
  With PropBag
    '
    .WriteProperty "PictureOnMouseLeave", m_PictureOnML, LoadPicture("")
    .WriteProperty "ShowPicture", m_bShowPicture, True
    .WriteProperty "MaskColor", UserControl.MaskColor, vbBlack
    .WriteProperty "BackColor", UserControl.BackColor, vbWhite
    .WriteProperty "PictureOnMouseMove", m_PictureOnMM, LoadPicture("")
    .WriteProperty "PictureOnMouseDown", m_PictureOnMD, LoadPicture("")
    .WriteProperty "Caption", lblCaption.Caption, ""
    .WriteProperty "Font", lblCaption.Font, "MS Sans Serif"
    .WriteProperty "ForeColor", lblCaption.ForeColor, vbBlack
    .WriteProperty "Enable", UserControl.Enabled, True
    '
  End With
  '
End Sub
'=============================================================
Public Property Get MaskColor() As OLE_COLOR
  '
  MaskColor = UserControl.MaskColor
  '
End Property
'=============================================================
Public Property Let MaskColor(ByVal NewColor As OLE_COLOR)
  '
  UserControl.MaskColor = NewColor
  '
  Refresh
  '
  PropertyChanged "MaskColor"
  '
End Property
'=============================================================
Public Property Get BackColor() As OLE_COLOR
  '
  BackColor = UserControl.BackColor
  '
End Property
'=============================================================
Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
  '
  UserControl.BackColor = NewColor
  '
  PropertyChanged "BackColor"
  '
End Property
'=============================================================
Public Property Get PictureOnMouseMove() As StdPicture
  '
  Set PictureOnMouseMove = m_PictureOnMM
  '
End Property
'=============================================================
Public Property Set PictureOnMouseMove(ByVal NewPicture As IPictureDisp)
  '
  If NewPicture Is Nothing Then Set m_PictureOnMM = LoadPicture("")
  '
  If Not (NewPicture Is Nothing) Then
    '
    Set m_PictureOnMM = NewPicture
    '
  End If
  '
  PropertyChanged "PictureOnMouseMove"
  '
End Property
'=============================================================
Public Property Get PictureOnMouseDown() As StdPicture
  '
  Set PictureOnMouseDown = m_PictureOnMD
  '
End Property
'=============================================================
Public Property Set PictureOnMouseDown(ByVal NewPicture As IPictureDisp)
  '
  If NewPicture Is Nothing Then Set m_PictureOnMD = LoadPicture("")
  '
  If Not (NewPicture Is Nothing) Then
    '
    Set m_PictureOnMD = NewPicture
    '
  End If
  '
  PropertyChanged "PictureOnMouseDown"
  '
End Property
'=============================================================
Public Property Get Caption() As String
  '
  Caption = lblCaption.Caption
  '
End Property
'=============================================================
Public Property Let Caption(ByVal NewCaption As String)
  '
  lblCaption.Caption = NewCaption
  '
  UserControl_Resize
  '
  PropertyChanged "Caption"
  '
End Property
'=============================================================
Public Property Get Font() As StdFont
  '
  Set Font = lblCaption.Font
  '
End Property
'=============================================================
Public Property Set Font(ByVal NewFont As StdFont)
  '
  Set lblCaption.Font = NewFont
  '
  PropertyChanged "Font"
  '
End Property
'=============================================================
Public Property Get ForeColor() As OLE_COLOR
  '
  ForeColor = lblCaption.ForeColor
  '
End Property
'=============================================================
Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
  '
  lblCaption.ForeColor = NewColor
  '
  PropertyChanged "ForeColor"
  '
End Property
'=============================================================
Public Property Get Enable() As Boolean
  '
  Enable = UserControl.Enabled
  '
End Property
'==========================================================
Public Property Let Enable(ByVal NewValue As Boolean)
  '
  UserControl.Enabled = NewValue
  '
End Property
'==========================================================
