VERSION 5.00
Begin VB.UserControl AxGLine 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   ToolboxBitmap   =   "AxGLine.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Timer tmrMOUSEOVER 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "AxGLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxGLine
'Version  : 2.07.6
'Editor   : David Rojas [AxioUK]
'Date     : 19/05/2022
'------------------------------------
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTL) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTL) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal token As Long)
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long

Private Declare Function GdipCreateLineBrushI Lib "gdiplus" (point1 As POINTL, point2 As POINTL, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As Long, lineGradient As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipSetPenBrushFill Lib "gdiplus" (ByVal pen As Long, ByVal brush As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As Long, ByVal dStyle As DashStyle) As Long
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal pen As Long, ByVal startCap As LineCap) As Long
Private Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal pen As Long, ByVal endCap As LineCap) As Long

Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTL
    X As Long
    Y As Long
End Type

Public Enum LineCap
   LineCapFlat = 0
   LineCapSquare = 1
   LineCapRound = 2
   LineCapTriangle = 3
   
   LineCapNoAnchor = &H10         ' corresponds to flat cap
   LineCapSquareAnchor = &H11     ' corresponds to square cap
   LineCapRoundAnchor = &H12      ' corresponds to round cap
   LineCapDiamondAnchor = &H13    ' corresponds to triangle cap
   LineCapArrowAnchor = &H14      ' no correspondence

   'LineCapCustom = &HFF           ' custom cap

   'LineCapAnchorMask = &HF0        ' mask to check for anchor or not.
End Enum

Public Enum DashStyle
   DashStyleSolid          ' 0
   DashStyleDash           ' 1
   DashStyleDot            ' 2
   DashStyleDashDot        ' 3
   DashStyleDashDotDot     ' 4
   'DashStyleCustom         ' 5
End Enum

Private Const WM_MOUSEMOVE As Long = &H200
Private Const WrapModeTileFlipXY = &H3
Private Const SmoothingModeHighQuality As Long = &H2
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const LOGPIXELSX As Long = 88
Private Const TLS_MINIMUM_AVAILABLE As Long = 64
Private Const UnitPixel As Long = &H2&


Private gdipToken As Long
Private nScale As Double
Private uhDC As Long
Private c_lhWnd As Long

Private m_X1 As Long
Private m_Y1 As Long
Private m_X2 As Long
Private m_Y2 As Long
Private m_GradientColor1 As OLE_COLOR
Private m_GradientColor2 As OLE_COLOR
Private m_LineWidth As Long
Private m_LineOpacity1 As Long
Private m_LineOpacity2 As Long
Private m_LineDashStyle As DashStyle
Private m_LineCapStart As LineCap
Private m_LineCapEnd As LineCap
Private m_RectLine As Boolean

Private m_MouseToParent As Boolean
Private m_PT As POINTL
Private bIntercept As Boolean
Private m_Enter As Boolean
Private m_Over As Boolean
Private m_Left As Long
Private m_Top As Long



Public Sub tmrMOUSEOVER_Timer()
    Dim PT As POINTL
    Dim Left As Long, Top As Long
    Dim Rect As Rect
  
    GetCursorPos PT
    ScreenToClient c_lhWnd, PT
    
    Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
    Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)

    With Rect
        .Left = m_PT.X - (m_Left - Left)
        .Top = m_PT.Y - (m_Top - Top)
        .Right = .Left + UserControl.ScaleWidth
        .Bottom = .Top + UserControl.ScaleHeight
    End With
    
    bIntercept = False
    SendMessage c_lhWnd, WM_MOUSEMOVE, 0&, ByVal PT.X Or PT.Y * &H10000
    
    If bIntercept = False Then
        If m_Over = True Then
            m_Over = False
        End If
    End If
    
    If PtInRect(Rect, PT.X, PT.Y) = 0 Then
        m_Enter = False
        tmrMOUSEOVER.Interval = 0
    End If
    
End Sub

Private Function RGBA(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
  If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
  RGBA = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
  Opacity = CByte((Abs(Opacity) / 100) * 255)
  If Opacity < 128 Then
      If Opacity < 0& Then Opacity = 0&
      RGBA = RGBA Or Opacity * &H1000000
  Else
      If Opacity > 255& Then Opacity = 255&
      RGBA = RGBA Or (Opacity - 128&) * &H1000000 Or &H80000000
  End If
End Function


Private Sub DrawLine()
Dim hGraphics As Long
Dim hPen As Long
Dim hBrush As Long
Dim lP1 As POINTL
Dim lP2 As POINTL
  
  GdipCreateFromHDC uhDC, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
  
  lP1.X = IIf(m_X1 = 0, m_LineWidth, m_X1): lP1.Y = IIf(m_Y1 = 0, m_LineWidth, m_Y1)
  lP2.X = IIf(m_X2 = 0, m_LineWidth, m_X2): lP2.Y = IIf(m_Y2 = 0, m_LineWidth, m_Y2)

  GdipCreatePen1 RGBA(m_GradientColor1, m_LineOpacity1), m_LineWidth * nScale, UnitPixel, hPen
  GdipCreateLineBrushI lP1, lP2, RGBA(m_GradientColor1, m_LineOpacity1), RGBA(m_GradientColor2, m_LineOpacity2), WrapModeTileFlipXY, hBrush
  
  GdipSetPenBrushFill hPen, hBrush
  GdipSetPenDashStyle hPen, m_LineDashStyle
  GdipSetPenStartCap hPen, m_LineCapStart
  GdipSetPenEndCap hPen, m_LineCapEnd
  
  GdipDrawLineI hGraphics, hPen, lP1.X, lP1.Y, lP2.X, lP2.Y
  
  Call GdipDeleteBrush(hBrush)
  Call GdipDeletePen(hPen)
  Call GdipDeleteGraphics(hGraphics)
    
End Sub

Private Function GetWindowsDPI() As Double
    Dim hDC As Long, lPx  As Double
    hDC = GetDC(0)
    lPx = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    ReleaseDC 0, hDC

    If (lPx = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = lPx / 96#
    End If
End Function

'Inicia GDI+
Private Sub InitGDI()
    Dim gdipStartupInput As GdiplusStartupInput
    gdipStartupInput.GdiplusVersion = 1&
    Call GdiplusStartup(gdipToken, gdipStartupInput, ByVal 0)
End Sub

'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(gdipToken)
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
On Error Resume Next

    If UserControl.Enabled Then
        If Not MouseToParent Then
            HitResult = vbHitResultHit
        ElseIf Not Ambient.UserMode Then
            HitResult = vbHitResultHit
        End If
        If Ambient.UserMode Then
            Dim PT As POINTL
            Dim lHwnd As Long
            GetCursorPos PT
            lHwnd = WindowFromPoint(PT.X, PT.Y)
            
            If m_Enter = False Then

                ScreenToClient c_lhWnd, PT
                m_PT.X = PT.X - X
                m_PT.Y = PT.Y - Y
    
                m_Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
                m_Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)
 
                m_Enter = True
                tmrMOUSEOVER.Interval = 1
            End If
        
            bIntercept = True
            
            If lHwnd = c_lhWnd Then
                If m_Over = False Then
                    m_Over = True
                End If
            Else
                If m_Over = True Then
                    m_Over = False
                End If
            End If
        End If
    ElseIf Not Ambient.UserMode Then
        HitResult = vbHitResultHit
    End If
End Sub

Private Sub UserControl_Initialize()
InitGDI
nScale = GetWindowsDPI
End Sub


Private Sub UserControl_InitProperties()
  m_X1 = 0
  m_Y1 = 0
  m_X2 = 20
  m_Y2 = 20
  m_GradientColor1 = vbRed
  m_GradientColor2 = vbBlue
  m_LineWidth = 1
  m_LineOpacity1 = 100
  m_LineOpacity2 = 100
  m_LineDashStyle = DashStyleSolid
  m_MouseToParent = False
  m_LineCapStart = LineCapFlat
  m_LineCapEnd = LineCapFlat
  m_RectLine = True
  
End Sub

Private Sub UserControl_Paint()
uhDC = UserControl.hDC
DrawLine
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  m_X1 = .ReadProperty("X1", 0)
  m_Y1 = .ReadProperty("Y1", 0)
  m_X2 = .ReadProperty("X2", 20)
  m_Y2 = .ReadProperty("Y2", 20)
  m_GradientColor1 = .ReadProperty("GradientColor1", vbRed)
  m_GradientColor2 = .ReadProperty("GradientColor2", vbBlue)
  m_LineWidth = .ReadProperty("LineWidth", 1)
  m_LineOpacity1 = .ReadProperty("LineOpacity1", 100)
  m_LineOpacity2 = .ReadProperty("LineOpacity2", 100)
  m_LineDashStyle = .ReadProperty("LineDashStyle", DashStyleSolid)
  m_MouseToParent = .ReadProperty("MouseToParent", False)
  m_LineCapStart = .ReadProperty("LineCapStart", 0)
  m_LineCapEnd = .ReadProperty("LineCapEnd", 0)
  m_RectLine = .ReadProperty("RectLine", m_RectLine)
  
End With
End Sub

Public Sub Refresh()
UserControl.Refresh
End Sub

Private Sub UserControl_Resize()
With UserControl
  If m_RectLine Then
    m_X1 = IIf(.Height > .Width, (.ScaleWidth / 2), 0)
    m_Y1 = IIf(.Height > .Width, 0, (.ScaleHeight / 2))
    m_X2 = IIf(.Height > .Width, (.ScaleWidth / 2), .ScaleWidth - m_LineWidth)
    m_Y2 = IIf(.Height > .Width, .ScaleHeight - m_LineWidth, (.ScaleHeight / 2))
  End If
End With

Refresh
End Sub

Private Sub UserControl_Show()
Refresh
End Sub

Private Sub UserControl_Terminate()
TerminateGDI
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("X1", m_X1)
  Call .WriteProperty("Y1", m_Y1)
  Call .WriteProperty("X2", m_X2)
  Call .WriteProperty("Y2", m_Y2)
  Call .WriteProperty("GradientColor1", m_GradientColor1)
  Call .WriteProperty("GradientColor2", m_GradientColor2)
  Call .WriteProperty("LineWidth", m_LineWidth)
  Call .WriteProperty("LineOpacity1", m_LineOpacity1)
  Call .WriteProperty("LineOpacity2", m_LineOpacity2)
  Call .WriteProperty("LineDashStyle", m_LineDashStyle)
  Call .WriteProperty("MouseToParent", m_MouseToParent)
  Call .WriteProperty("LineCapStart", m_LineCapStart)
  Call .WriteProperty("LineCapEnd", m_LineCapEnd)
  Call .WriteProperty("RectLine", m_RectLine)
  
End With
End Sub


Public Property Get GradientColor1() As OLE_COLOR
  GradientColor1 = m_GradientColor1
End Property

Public Property Let GradientColor1(ByVal NewGradientColor1 As OLE_COLOR)
  m_GradientColor1 = NewGradientColor1
  PropertyChanged "GradientColor1"
  Refresh
End Property

Public Property Get GradientColor2() As OLE_COLOR
  GradientColor2 = m_GradientColor2
End Property

Public Property Let GradientColor2(ByVal NewGradientColor2 As OLE_COLOR)
  m_GradientColor2 = NewGradientColor2
  PropertyChanged "GradientColor2"
  Refresh
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get LineDashStyle() As DashStyle
  LineDashStyle = m_LineDashStyle
End Property

Public Property Let LineDashStyle(ByVal NewLineDashStyle As DashStyle)
  m_LineDashStyle = NewLineDashStyle
  PropertyChanged "LineDashStyle"
  Refresh
End Property

Public Property Get LineOpacity1() As Long
  LineOpacity1 = m_LineOpacity1
End Property

Public Property Let LineOpacity1(ByVal NewLineOpacity As Long)
  m_LineOpacity1 = NewLineOpacity
  PropertyChanged "LineOpacity1"
  Refresh
End Property

Public Property Get LineOpacity2() As Long
  LineOpacity2 = m_LineOpacity2
End Property

Public Property Let LineOpacity2(ByVal NewLineOpacity As Long)
  m_LineOpacity2 = NewLineOpacity
  PropertyChanged "LineOpacity2"
  Refresh
End Property

Public Property Get LineWidth() As Long
  LineWidth = m_LineWidth
End Property

Public Property Let LineWidth(ByVal NewLineWidth As Long)
  m_LineWidth = NewLineWidth
  PropertyChanged "LineWidth"
  Refresh
End Property

Public Property Get X1() As Long
  X1 = m_X1
End Property

Public Property Let X1(ByVal NewX1 As Long)
  m_X1 = NewX1
  PropertyChanged "X1"
  Refresh
End Property

Public Property Get x2() As Long
  x2 = m_X2
End Property

Public Property Let x2(ByVal NewX2 As Long)
  m_X2 = NewX2
  PropertyChanged "X2"
  Refresh
End Property

Public Property Get y1() As Long
  y1 = m_Y1
End Property

Public Property Let y1(ByVal NewY1 As Long)
  m_Y1 = NewY1
  PropertyChanged "Y1"
  Refresh
End Property

Public Property Get y2() As Long
  y2 = m_Y2
End Property

Public Property Let y2(ByVal NewY2 As Long)
  m_Y2 = NewY2
  PropertyChanged "Y2"
  Refresh
End Property

Public Property Get MouseToParent() As Boolean
    MouseToParent = m_MouseToParent
End Property

Public Property Let MouseToParent(ByVal New_Value As Boolean)
    m_MouseToParent = New_Value
    PropertyChanged "MouseToParent"
End Property

Public Property Get LineCapStart() As LineCap
  LineCapStart = m_LineCapStart
End Property

Public Property Let LineCapStart(ByVal NewLineCapStart As LineCap)
  m_LineCapStart = NewLineCapStart
  PropertyChanged "LineCapStart"
  Refresh
End Property

Public Property Get LineCapEnd() As LineCap
  LineCapEnd = m_LineCapEnd
End Property

Public Property Let LineCapEnd(ByVal NewLineCapEnd As LineCap)
  m_LineCapEnd = NewLineCapEnd
  PropertyChanged "LineCapEnd"
  Refresh
End Property

Public Property Get RectLine() As Boolean
    RectLine = m_RectLine
End Property

Public Property Let RectLine(ByVal New_Value As Boolean)
    m_RectLine = New_Value
    PropertyChanged "RectLine"
    Refresh
End Property

