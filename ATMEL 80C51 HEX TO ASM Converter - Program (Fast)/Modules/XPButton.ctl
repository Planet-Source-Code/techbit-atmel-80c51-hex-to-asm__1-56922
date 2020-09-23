VERSION 5.00
Begin VB.UserControl XPButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   KeyPreview      =   -1  'True
   ScaleHeight     =   1485
   ScaleWidth      =   3435
End
Attribute VB_Name = "XPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ButtonType
    [Normal] = 1
    [DropDown] = 2
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNFACE = 15

Private COLOR_USER_OUTLINE As Long
Private COLOR_USER_MARKER As Long
Private COLOR_USER_MARKER_SEL As Long
Private COLOR_USER_MARKER_DOWN As Long

Private m_Caption As String
Private m_Font As IFontDisp
Private m_Enabled As Boolean
Private m_ButtonType As ButtonType

Private blnOutOfRange As Boolean
Private blnHasFocus As Boolean
Private blnIsDown As Boolean

Event Click()

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Let Caption(Str As String)
    m_Caption = Str

    UpdateCaption

    PropertyChanged "Caption"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let ButtonType(intType As ButtonType)
    m_ButtonType = intType
    
    UpdateCaption
    PropertyChanged "ButtonType"
End Property

Public Property Get ButtonType() As ButtonType
    ButtonType = m_ButtonType
End Property

Public Property Let Enabled(State As Boolean)
    m_Enabled = State

    UserControl.Extender.TabStop = m_Enabled
    UserControl.ForeColor = IIf(m_Enabled, 0, GetSysColor(COLOR_BTNSHADOW))

    UpdateCaption
    PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Set Font(Font As IFontDisp)
    Set m_Font = Font
    Set UserControl.Font = m_Font

    UpdateCaption
    PropertyChanged "Font"
End Property

Public Property Get Font() As IFontDisp
    Set Font = m_Font
End Property

Private Sub UserControl_GotFocus()
    blnHasFocus = True
    UpdateCaption
End Sub

Private Sub UserControl_Initialize()
    COLOR_USER_OUTLINE = RGB(10, 36, 106)
    COLOR_USER_MARKER = RGB(212, 213, 216)
    COLOR_USER_MARKER_SEL = RGB(182, 189, 210)
    COLOR_USER_MARKER_DOWN = RGB(133, 146, 181)
End Sub

Private Sub UserControl_InitProperties()
    Set m_Font = Ambient.Font
    Set UserControl.Font = m_Font
    m_ButtonType = Normal
    m_Caption = Ambient.DisplayName
    m_Enabled = True

    UpdateCaption
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then Call UserControl_MouseDown(vbLeftButton, Shift, 0, 0)
    If KeyCode = vbKeyReturn Then RaiseEvent Click
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then Call UserControl_MouseUp(vbLeftButton, Shift, 0, 0)
End Sub

Private Sub UserControl_LostFocus()
    blnHasFocus = False
    UpdateCaption
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_Enabled And Button = vbLeftButton Then
        blnIsDown = True
        UpdateCaption
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnOutOfRange = True
    If x >= 0 And x <= UserControl.Width Then
        If y >= 0 And y <= UserControl.Height Then
            blnOutOfRange = False
        End If
    End If

    If blnOutOfRange And GetCapture() = UserControl.hwnd Then
        Call ReleaseCapture
        blnIsDown = False

        UpdateCaption
    ElseIf Not blnOutOfRange And GetCapture() <> UserControl.hwnd Then
        Call SetCapture(UserControl.hwnd)
    End If
    
    If m_Enabled Then UpdateCaption IIf(Not blnOutOfRange, COLOR_USER_MARKER_SEL, COLOR_USER_MARKER)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_Enabled Then
        If Not blnOutOfRange And m_ButtonType = DropDown Then RaiseEvent Click

        blnIsDown = False
        UpdateCaption

        If Not blnOutOfRange And m_ButtonType <> DropDown Then RaiseEvent Click
    End If
End Sub

Private Sub UserControl_Resize()
    If UserControl.Height < 210 Then UserControl.Height = 210
    
    UpdateCaption
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    m_ButtonType = PropBag.ReadProperty("ButtonType", 1)
    
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = m_Font

    UserControl.Extender.TabStop = m_Enabled
    UserControl.ForeColor = IIf(m_Enabled, 0, GetSysColor(COLOR_BTNSHADOW))

    UpdateCaption
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", m_Caption, Ambient.DisplayName
    PropBag.WriteProperty "Enabled", m_Enabled, True
    PropBag.WriteProperty "ButtonType", m_ButtonType, 1
    PropBag.WriteProperty "Font", m_Font, Ambient.Font
End Sub

Private Sub UpdateCaption(Optional FillColor As Long)
    Dim sizSize As Size
    Dim lngY As Long
    Dim lngX As Long
    Dim rctOutline As RECT
    Dim hBrush As Long

    Call GetTextExtentPoint32(UserControl.hdc, m_Caption, Len(m_Caption), sizSize)

    lngY = (ScaleY(UserControl.Height, vbTwips, vbPixels) / 2) - sizSize.cy / 2
    lngX = ((ScaleX(UserControl.Width, vbTwips, vbPixels) - IIf(m_ButtonType = DropDown, 20, 0)) / 2) - sizSize.cx / 2
    
    If lngY < 0 Then lngY = 0
    If lngX < 0 Then lngX = 0

    If IsMissing(FillColor) Or FillColor = 0 Then FillColor = COLOR_USER_MARKER
    If Not m_Enabled Then FillColor = GetSysColor(COLOR_BTNFACE)
    If blnIsDown Then FillColor = COLOR_USER_MARKER_DOWN

    hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))

    Call SetRect(rctOutline, 0, 0, ScaleX(UserControl.Width, vbTwips, vbPixels), ScaleY(UserControl.Height, vbTwips, vbPixels))
    Call FillRect(UserControl.hdc, rctOutline, hBrush)
    Call DeleteObject(hBrush)

    hBrush = CreateSolidBrush(FillColor)

    Call FillRect(UserControl.hdc, rctOutline, hBrush)
    Call DeleteObject(hBrush)

    hBrush = CreateSolidBrush(IIf(m_Enabled, COLOR_USER_OUTLINE, GetSysColor(COLOR_BTNSHADOW)))

    Call FrameRect(UserControl.hdc, rctOutline, hBrush)
    If m_ButtonType = DropDown Then
        Call SetRect(rctOutline, 0, 0, ScaleX(UserControl.Width, vbTwips, vbPixels) - 16, ScaleY(UserControl.Height, vbTwips, vbPixels))
        Call FrameRect(UserControl.hdc, rctOutline, hBrush)
    End If
    
    Call DeleteObject(hBrush)

    Call TextOut(UserControl.hdc, lngX, lngY, m_Caption, Len(m_Caption))
    Call DeleteObject(hBrush)

    If blnHasFocus And m_Enabled Then
        Dim rctFocus As RECT

        Call SetRect(rctFocus, 4, 4, ScaleX(UserControl.Width, vbTwips, vbPixels) - IIf(m_ButtonType = DropDown, 20, 4), ScaleY(UserControl.Height, vbTwips, vbPixels) - 4)
        Call DrawFocusRect(UserControl.hdc, rctFocus)
    End If

    If m_ButtonType = DropDown Then
        Dim ptOldPos As POINTAPI

        lngY = ScaleY((UserControl.Height / 2) - 4 / 2, vbTwips, vbPixels) - 1
        lngX = ScaleX(UserControl.Width, vbTwips, vbPixels) - 12

        Call MoveToEx(UserControl.hdc, lngX, lngY, ptOldPos)
        Call LineTo(UserControl.hdc, lngX + 7, lngY)

        Call MoveToEx(UserControl.hdc, lngX + 1, lngY + 1, ptOldPos)
        Call LineTo(UserControl.hdc, lngX + 6, lngY + 1)

        Call MoveToEx(UserControl.hdc, lngX + 2, lngY + 2, ptOldPos)
        Call LineTo(UserControl.hdc, lngX + 5, lngY + 2)

        Call SetPixel(UserControl.hdc, lngX + 3, lngY + 3, UserControl.ForeColor)
    End If

    UserControl.Refresh
End Sub
