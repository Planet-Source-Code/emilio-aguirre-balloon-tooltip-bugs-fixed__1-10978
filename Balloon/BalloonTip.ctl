VERSION 5.00
Begin VB.UserControl BalloonTip 
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   285
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   Picture         =   "BalloonTip.ctx":0000
   ScaleHeight     =   19
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   19
   ToolboxBitmap   =   "BalloonTip.ctx":0102
   Begin VB.PictureBox Balloonform 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Timer tmrBalloon 
      Interval        =   50
      Left            =   960
      Top             =   0
   End
End
Attribute VB_Name = "BalloonTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------
' ________  Copyright EAguirre (c)1999
'(        ) eaguirre@comtrade.com.mx
'(  ______)
' \/
' BalloonToolTip
'----------------------------------------
Option Explicit
'User Defined Enumerators
Enum WordBoolValue
    No = 0
    yes = 1
End Enum

Enum TextAlignValue
    To_Left = 0
    To_Center = 2
    To_Right = 1
End Enum

Enum StyleValue
    Rectangle = 0
    Balloon = 1
    Round_Rectangle = 2
End Enum

Enum OrientationValues
    North = 0
    NE = 1
    East = 2
    SE = 3
    South = 4
    Sw = 5
    West = 6
    NW = 7
End Enum

'Type Declarations
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type
Const GW_CHILD = 5
Const SW_SHOWNOACTIVATE = 4
'Drawing Text
Const DT_CALCRECT = &H400
Const DT_CENTER = &H1
Const DT_LEFT = &H0
Const DT_RIGHT = &H2
Const DT_WORDBREAK = &H10
'Region
Const RGN_OR = 2
'Window Styles
Const GWL_STYLE = -16
Const GWL_EXSTYLE = -20
'Window Constants
Const WS_BORDER = &H800000
Const WS_CAPTION = &HC00000
Const WS_THICKFRAME = &H40000

'Functions Declares
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, _
                        ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
                        ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
                        ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
                        ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, _
                        ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, _
                        ByVal nCombineMode As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, _
                        ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, _
                        ByVal bRedraw As Boolean) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
                        ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, _
                        ByVal wFormat As Long) As Long

Dim BalloonCtrl As Control          'Control under the mouse
Dim BalloonBox As RECT              'Balloon Box coordinates
Dim m_oldhWnd As Long

'Default Property Values:
Const m_def_AutoSize = yes
Const m_def_TextAlign = To_Left
Const m_def_WordBreak = yes
Const m_def_Orientation = NE
Const m_def_BackColor = &HFFFF&
Const m_def_ForeColor = 0
Const m_def_Text = " "
Const m_def_Style = Balloon

'Property Variables:
Dim m_AutoSize As WordBoolValue
Dim m_TextAlign As TextAlignValue
Dim m_WordBreak As WordBoolValue
Dim m_Orientation As OrientationValues
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Font As Font
Dim m_Text As String
Dim m_Style As Variant
Dim m_init As Boolean

'Properties
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
'
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    UserControl.ForeColor = m_ForeColor
End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
    Set UserControl.Font = m_Font
End Property

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    PropertyChanged "Text"
End Property

Public Property Get Style() As StyleValue
    Style = m_Style
End Property

Public Property Let Style(ByVal new_Style As StyleValue)
    m_Style = new_Style
    PropertyChanged "Style"
End Property

Public Property Get Orientation() As OrientationValues
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As OrientationValues)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
End Property

Public Property Get TextAlign() As TextAlignValue
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal New_TextAlign As TextAlignValue)
    m_TextAlign = New_TextAlign
    PropertyChanged "TextAlign"
End Property

Public Property Get WordBreak() As WordBoolValue
    WordBreak = m_WordBreak
End Property

Public Property Let WordBreak(ByVal New_WordBreak As WordBoolValue)
    m_WordBreak = New_WordBreak
    PropertyChanged "WordBreak"
End Property

Public Property Get AutoSize() As WordBoolValue
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As WordBoolValue)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
End Property

Private Sub tmrBalloon_Timer()
Search_Wnd
End Sub

Private Sub UserControl_Initialize()
    InitProc
End Sub

Private Sub UserControl_Resize()
'Keep short
Width = 240
Height = 240
End Sub

Private Sub UserControl_Terminate()
    TerminateProc
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_Text = m_def_Text
    m_Style = m_def_Style
    m_Orientation = m_def_Orientation
    m_TextAlign = m_def_TextAlign
    m_WordBreak = m_def_WordBreak
    m_AutoSize = m_def_AutoSize
    m_init = False
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    m_WordBreak = PropBag.ReadProperty("WordBreak", m_def_WordBreak)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, m_def_TextAlign)
    Call PropBag.WriteProperty("WordBreak", m_WordBreak, m_def_WordBreak)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
End Sub

Private Sub InitProc()
    Dim lngOldStyle As Long
    lngOldStyle = GetWindowLong(Balloonform.hWnd, GWL_STYLE)
    SetWindowLong Balloonform.hWnd, lngOldStyle Xor 0, 0
    SetParent Balloonform.hWnd, 0
    Balloonform.Visible = False
End Sub

Sub TerminateProc()
    Dim lngOldStyle As Long
    SetParent Balloonform.hWnd, UserControl.hWnd
    lngOldStyle = GetWindowLong(Balloonform.hWnd, GWL_STYLE)
    SetWindowLong Balloonform.hWnd, lngOldStyle, 0
End Sub

Public Sub Search_Wnd()
  Dim curhWnd As Long
  Dim p As POINTAPI
  Static oldhWnd As Long
  Dim blnFound As Boolean
  Dim ctrl As Control
  
  On Error Resume Next
  
  If GetActiveWindow() = UserControl.Parent.hWnd Then
    Call GetCursorPos(p)
    curhWnd = WindowFromPoint(p.x, p.y)
      
    If (m_oldhWnd <> curhWnd) Then
         blnFound = False
         HideTip
         'Check if mouse its over the form
         If curhWnd = UserControl.Parent.hWnd Then
              m_oldhWnd = curhWnd
              Exit Sub
         End If
         'Search the control under cursor
         For Each ctrl In UserControl.Parent.Controls
           If (Not (TypeOf ctrl Is BalloonTip)) And (Not (TypeOf ctrl Is Label)) Then
                If (ctrl.hWnd <> curhWnd) Then
                'hWnd property not supported or not found yet
                  If TypeOf ctrl Is ComboBox Then
                     If GetWindow(ctrl.hWnd, GW_CHILD) = curhWnd Then
                        If Len(ctrl.ToolTipText) > 0 Then
                             m_oldhWnd = ctrl.hWnd
                             Set BalloonCtrl = ctrl
                             With ctrl
                                 Me.Text = .ToolTipText
                                 .ToolTipText = ""
                                 blnFound = True
                                 curhWnd = ctrl.hWnd
                             End With
                         End If
                         Exit For
                     End If
                  End If
                Else
                  If Len(ctrl.ToolTipText) > 0 Then
                    m_oldhWnd = ctrl.hWnd
                    Set BalloonCtrl = ctrl
                    With ctrl
                      Me.Text = .ToolTipText
                      .ToolTipText = ""
                      blnFound = True
                    End With
                  End If
                  Exit For
                End If
            End If
         Next
         If blnFound Then
           DisplayBalloon
         Else
           HideTip
         End If
     End If
  End If
End Sub

Public Sub HideTip()
If Not (BalloonCtrl Is Nothing) Then
     SetParent Balloonform.hWnd, UserControl.hWnd
     Balloonform.Visible = False
    'Restore Values of the Control
     BalloonCtrl.ToolTipText = Me.Text
     Me.Text = ""
     Set BalloonCtrl = Nothing
End If
End Sub

Private Sub ChangeStyle()
Dim Reg(2) As Long
Dim p(3) As POINTAPI
Dim Box As RECT
Dim w As Single, h As Single
'Copy values to variables for optimization
w = Balloonform.ScaleWidth: h = Balloonform.ScaleHeight
'Establish the form of the balloon depending the Orientation
'Property.
Select Case Me.Orientation
    Case North, South
       p(0).x = (w / 2) - (w * 0.15): p(0).y = h / 2
       p(1).x = (w / 2) + (w * 0.15): p(1).y = h / 2
       p(2).x = w / 2
       Box.Left = 0: Box.Right = w
       If Me.Orientation = North Then
         Box.Top = 0:   Box.Bottom = h - (h * 0.1)
         p(2).y = h
       Else
         Box.Top = h * 0.1: Box.Bottom = h
         p(2).y = 0
       End If
    Case NE, Sw
       p(0).x = (w / 2) - (w * 0.15): p(0).y = (h / 2) - (h * 0.15)
       p(1).x = (w / 2) + (w * 0.15): p(1).y = (h / 2) + (h * 0.15)
       Box.Left = 0: Box.Right = w
       If Me.Orientation = NE Then
         Box.Top = 0: Box.Bottom = h - (h * 0.1)
         p(2).x = 0: p(2).y = h
       Else
         Box.Top = h * 0.1: Box.Bottom = h
         p(0).x = (w / 2) - (w * 0.15): p(0).y = (h / 2) - (h * 0.15)
         p(1).x = (w / 2) + (w * 0.15): p(1).y = (h / 2) + (h * 0.15)
         p(2).x = w: p(2).y = 0
       End If
    Case East, West
       p(0).x = (w / 2): p(0).y = (h / 2) + (h * 0.15)
       p(1).x = (w / 2): p(1).y = (h / 2) - (h * 0.15)
       p(2).y = h / 2
       Box.Top = 0: Box.Bottom = h
       If Me.Orientation = East Then
         Box.Left = w * 0.1: Box.Right = w
         p(2).x = 0
       Else
         Box.Left = 0: Box.Right = w - (w * 0.1)
         p(2).x = w
       End If
    Case NW, SE
       p(0).x = (w / 2) - (w * 0.15): p(0).y = (h / 2) + (h * 0.15)
       p(1).x = (w / 2) + (w * 0.15): p(1).y = (h / 2) - (h * 0.15)
       Box.Left = 0: Box.Right = w
       If Me.Orientation = NW Then
         Box.Top = 0: Box.Bottom = h - (h * 0.1)
         p(2).x = w: p(2).y = h
       Else
         Box.Top = h * 0.1: Box.Bottom = h
         p(2).x = 0: p(2).y = 0
       End If
End Select
'Create Region 1: Balloon Body
Select Case Me.Style
    Case Rectangle
      Reg(0) = CreateRectRgn(Box.Left, Box.Top, Box.Right, Box.Bottom)
    Case Balloon
      Reg(0) = CreateEllipticRgn(Box.Left, Box.Top, Box.Right, Box.Bottom)
    Case Round_Rectangle
      Reg(0) = CreateRoundRectRgn(Box.Left, Box.Top, Box.Right, Box.Bottom, w * 0.2, h * 0.2)
End Select
'Create Region 2: Tail of the balloon
Reg(1) = CreatePolygonRgn(p(0), 3, 1)
'Combine regions for balloon shape
CombineRgn Reg(1), Reg(1), Reg(0), RGN_OR
'Change the Balloonform shape
SetWindowRgn Balloonform.hWnd, Reg(1), True
'Adjust de box for fitting the label text
'in the case of elliptic regions
If Me.Style = Balloon Then
    BalloonBox.Bottom = Box.Bottom - h * 0.15
    BalloonBox.Left = Box.Left + w * 0.15
    BalloonBox.Right = Box.Right - w * 0.15
    BalloonBox.Top = Box.Top + h * 0.15
Else
    BalloonBox.Bottom = Box.Bottom
    BalloonBox.Left = Box.Left
    BalloonBox.Right = Box.Right
    BalloonBox.Top = Box.Top
End If
End Sub

Private Sub DrawLabel()
Dim lngFormat As Long
Dim new_box As RECT
Dim sngArea As Single
Dim oldArea As Single
Dim lngHeight As Long, lngWidth As Long

'Clear control's device context and change display properties
Balloonform.BackColor = Me.BackColor
Balloonform.ForeColor = Me.ForeColor

'Calculate the rectangle
DrawText Balloonform.hdc, Me.Text, Len(Me.Text), new_box, DT_CALCRECT
'Recalculate the balloon size for ensuring that all text will be displayed
sngArea = (new_box.Bottom - new_box.Top) * (new_box.Right - new_box.Left)
sngArea = sngArea * 1.15 'Leave extra space because the wordbreak
oldArea = (BalloonBox.Bottom - BalloonBox.Top) * (BalloonBox.Right - BalloonBox.Left)
If ((sngArea > oldArea) Or (sngArea < (oldArea * 0.65))) And (Me.AutoSize = yes) Then
   If Me.WordBreak = yes Then
    'New balloon width has to be twice the height
    lngHeight = CLng(Sqr(sngArea / 3) + 0.5) * 1.5
    lngWidth = 3.75 * CLng(Sqr(sngArea / 3) + 0.5)
  Else
    lngHeight = (new_box.Bottom - new_box.Top) * 1.2
    lngWidth = (new_box.Right - new_box.Left) * 1.2
  End If
  'Add space for the balloon tail
  Select Case Me.Orientation
    Case North, South, NE, Sw
       lngHeight = lngHeight + (lngHeight * 0.25)
    Case East, West, NW, SE
       lngWidth = lngWidth + (lngWidth * 0.25)
  End Select
  'Add more space in the case of elliptic shape
  If Me.Style = Balloon Then
   lngHeight = lngHeight + (lngHeight * 0.35)
   lngWidth = lngWidth + (lngWidth * 0.1)
  End If
  'Apply the new values to the Balloon
  'Remember: All calculations are made in pixels so
  'we have to convert it to Twips
  Balloonform.Width = lngWidth * Screen.TwipsPerPixelX
  Balloonform.Height = lngHeight * Screen.TwipsPerPixelY
  'Change the style of the Balloon
  ChangeStyle
End If


End Sub


Public Sub DisplayBalloon()
Dim iL As Integer, iT As Integer, iW As Integer, iH As Integer
Dim mCount As Integer
Dim ret As Integer

'Avoid Errors
'On Error Resume Next

SetParent Balloonform.hWnd, 0
'Copy data for optimization
With BalloonCtrl
    iL = .Left
    iT = .Top
    iW = .Width
    iH = .Height
End With
'Add the Caption Height if necessary
If UserControl.Parent.BorderStyle <> 0 Then iT = iT + 300
'Calculate AutoSize
DrawLabel
With Balloonform
  'Place the balloon tip behind the control in the position
  'indicated by the Orientation property
  Select Case Me.Orientation
    Case East, West
      .Top = UserControl.Parent.Top + iT + (iH / 2) - (Balloonform.Height / 2)
      If Me.Orientation = East Then
        .Left = UserControl.Parent.Left + iL + iW
      Else
        .Left = UserControl.Parent.Left + iL - Balloonform.Width
      End If
    Case Else
      If (Me.Orientation = South) Or (Me.Orientation = SE) Or (Me.Orientation = Sw) Then
        .Top = UserControl.Parent.Top + iT + iH
      Else
       .Top = UserControl.Parent.Top + iT - Balloonform.Height
      End If
      If (Me.Orientation = South) Or (Me.Orientation = North) Then
        .Left = UserControl.Parent.Left + iL + (iW / 2) - (Balloonform.Width / 2)
      ElseIf (Me.Orientation = SE) Or (Me.Orientation = NE) Then
        .Left = UserControl.Parent.Left + iL + iW
      Else
        .Left = UserControl.Parent.Left + iL - Balloonform.Width
      End If
  End Select
  'Make sure form is on top
  .ZOrder
  'Display and Draw
  .Visible = True
  DrawText Balloonform.hdc, Me.Text, Len(Me.Text), BalloonBox, 0
  End With
End Sub

Private Sub Balloonform_Paint()
Dim format As Long

format = 0
If WordBreak = yes Then format = format Or DT_WORDBREAK
Select Case TextAlign
  Case To_Center
     format = format Or DT_CENTER
  Case To_Right
     format = format Or DT_RIGHT
  Case To_Left
     format = format Or DT_LEFT
End Select
DrawText Balloonform.hdc, Me.Text, Len(Me.Text), BalloonBox, format
End Sub

