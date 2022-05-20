VERSION 5.00
Begin VB.UserControl AxGInfoPanel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   75
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   282
   ToolboxBitmap   =   "AxGInfoPanel.ctx":0000
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   255
      Top             =   240
   End
End
Attribute VB_Name = "AxGInfoPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxGInfoPanel
'Version  : 2.07.6
'Editor   : David Rojas [AxioUK]
'Date     : 19/05/2022
'------------------------------------
Option Explicit

Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'-
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal token As Long)
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef brush As Long) As Long
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTS, ByVal mFormat As Long) As Long
Private Declare Function GdipMeasureString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByRef mBoundingBox As RECTS, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As eStringTrimming) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As eStringFormatFlags) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As eStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As eStringAlignment) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal count As Long) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByRef pPoints As Any, ByVal count As Long, ByVal FillMode As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long) As Long

Private Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, lpStr As Long, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTL) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

Private Const LF_FACESIZE = 32
Private Const SYSTEM_FONT = 13
Private Const OBJ_FONT As Long = 6&

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Enum HotkeyPrefix
    HotkeyPrefixNone = &H0
    HotkeyPrefixShow = &H1
    HotkeyPrefixHide = &H2
End Enum

Private Type POINTL
    X As Long
    Y As Long
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type RECTS
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type PicBmp
  Size As Long
  type As Long
  hBmp As Long
  hPal As Long
  Reserved As Long
End Type

Private Type lPadding
  padY As Integer
  padX As Integer
End Type

Private Enum PenAlignment
    PenAlignmentCenter = &H0
    PenAlignmentInset = &H1
End Enum

'EVENTS------------------------------------
Public Event CrossClick()
Public Event PinClick()
Public Event Click()
Public Event DrawString()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
''-----------------------------------------

Private Const CombineModeExclude As Long = &H4
Private Const WrapModeTileFlipXY = &H3
Private Const SmoothingModeHighQuality As Long = &H2
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const TLS_MINIMUM_AVAILABLE As Long = 64
Private Const CLR_INVALID = -1
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const IDC_HAND As Long = 32649
Private Const UnitPixel As Long = &H2&

'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_Appearance = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Color1 = &HD59B5B
Const m_def_Color2 = &H6A5444
Const m_def_Angulo = 0

'Property Variables:
Dim gdipToken As Long
Dim nScale    As Single
Dim hCur      As Long
Dim hFontCollection As Long
Dim hGraphics As Long
Dim hPen      As Long
Dim hBrush    As Long

Dim m_BorderColor   As OLE_COLOR
Dim m_ActiveColor    As OLE_COLOR
Dim m_Color1        As OLE_COLOR
Dim m_Color2        As OLE_COLOR
Dim m_BorderWidth   As Long
Dim m_Enabled       As Boolean
Dim m_Angulo        As Single
Dim m_CornerCurve   As Long
Dim YCrossPos       As Long
Dim XCrossPos       As Long
Dim YPinPos         As Long
Dim XPinPos         As Long
Dim m_CrossPosition As CrossPos
Dim m_CrossVisible  As Boolean
Dim m_PinPosition   As CheckPos
Dim m_PinVisible    As Boolean
Dim m_Moveable      As Boolean
Dim m_LineOrientation As LineOr
Dim m_Line1       As Boolean
Dim m_Line2       As Boolean
Dim m_Line1Pos    As Integer
Dim m_Line2Pos    As Integer
Dim m_Roll        As Integer
Dim m_FullWidth   As Integer
Dim m_FullHeight  As Integer
Dim m_Rolled     As Boolean
Dim m_Top         As Long
Dim m_Left        As Long
Dim OldX          As Single
Dim OldY          As Single
Dim m_Opacity     As Long
Dim cl_hWnd       As Long

Dim m_OldBorderColor As OLE_COLOR
Dim m_BorderColorOnFocus As OLE_COLOR
Dim m_ChangeBorderOnFocus As Boolean
Dim m_OnFocus As Boolean

Private m_IconFont      As StdFont
Private m_IconCharCode  As Long
Private m_IconForeColor As Long
Private m_PadY          As Long
Private m_PadX          As Long
Private m_IconCharCode2  As Long
Private m_IconForeColor2 As Long
Private m_PadY2          As Long
Private m_PadX2          As Long

Private m_MouseOver     As Boolean

Private m_RollCaption   As String

Private m_CaptionAngle  As Single
Private m_CaptionX      As Long
Private m_CaptionY      As Long
Private m_Caption       As String
Private m_CaptionEnabled As Boolean
Private m_CaptionAlignV  As eTextAlignV
Private m_CaptionAlignH  As eTextAlignH
Private m_CaptionOpacity As Integer
Private m_CaptionFont    As StdFont
Private m_CaptionColor   As OLE_COLOR

Private m_Caption2Angle  As Single
Private m_Caption2X      As Long
Private m_Caption2Y      As Long
Private m_Caption2       As String
Private m_Caption2Enabled As Boolean
Private m_Caption2AlignV  As eTextAlignV
Private m_Caption2AlignH  As eTextAlignH
Private m_Caption2Opacity As Integer
Private m_Caption2Font    As StdFont
Private m_Caption2Color   As OLE_COLOR

Private m_StringPosX  As Long
Private m_StringPosY  As Long
Private m_EffectFade  As FadeEffect
Private m_InitialOpacity As Long

Public Function AddString(ByVal hDC As Long, sString As String, X As Long, Y As Long, Width As Long, Height As Long, mAngle As Single, oFont As StdFont, ForeColor As OLE_COLOR, ColorOpacity As Integer, HAlign As eTextAlignH, VAlign As eTextAlignV, bWordWrap As Boolean) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RECTS
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim hGraphics As Long
    
    GdipCreateFromHDC hDC, hGraphics
  
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        GdipSetStringFormatAlign hFormat, HAlign
        GdipSetStringFormatLineAlign hFormat, VAlign
    End If
        
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hDC, LOGPIXELSY), 72)

    layoutRect.Left = X * nScale: layoutRect.Top = Y * nScale
    layoutRect.Width = Width * nScale: layoutRect.Height = Height * nScale

    GdipCreateSolidFill argb(ForeColor, ColorOpacity), hBrush
            
    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    
    If mAngle <> 0 Then
        Dim newH As Long, newW As Long
        newH = (layoutRect.Height / 2)
        newW = (layoutRect.Width / 2)
        Call GdipTranslateWorldTransform(hGraphics, newW, newH, 0)
        Call GdipRotateWorldTransform(hGraphics, mAngle, 0)
        Call GdipTranslateWorldTransform(hGraphics, -newW, -newH, 0)
    End If
    
    GdipDrawString hGraphics, StrPtr(sString), -1, hFont, layoutRect, hFormat, hBrush
    
    Dim BB As RECTS, CF As Long, LF As Long

    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    
    
    GdipMeasureString hGraphics, StrPtr(sString), -1, hFont, layoutRect, hFormat, BB, CF, LF

    If bWordWrap Then
        AddString = BB.Height / nScale
    Else
        AddString = BB.Width / nScale
    End If
    
    GdipDeleteFont hFont
    GdipDeleteBrush hBrush
    GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily
    GdipDeleteGraphics hGraphics
    
End Function

Public Sub CopyAmbient()
Dim OPic As StdPicture

On Error GoTo Err

    With UserControl
        Set .Picture = Nothing
        Set OPic = Extender.Container.Image
        .BackColor = Extender.Container.BackColor
        UserControl.PaintPicture OPic, 0, 0, , , Extender.Left, Extender.Top
        Set .Picture = .Image
    End With
Err:
End Sub

Public Function DrawLine(ByVal hGraphics As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, Optional ByVal oColor As OLE_COLOR = vbBlack, Optional ByVal Opacity As Integer = 90, Optional ByVal PenWidth As Integer = 1) As Boolean
    Dim hPen As Long
    
    GdipCreatePen1 argb(oColor, Opacity), PenWidth * nScale, UnitPixel, hPen
    DrawLine = GdipDrawLineI(hGraphics, hPen, X1 * nScale, y1 * nScale, x2 * nScale, y2 * nScale) = 0
    GdipDeletePen hPen

End Function

Public Sub Refresh()
  UserControl.Cls
  CopyAmbient
  Draw
  RaiseEvent DrawString
End Sub

Private Function AddIconChar(CharCode As Long, Pad_X, Pad_Y, ForeColor As OLE_COLOR, ActiveColor As OLE_COLOR)
Dim Rct As Rect
Dim pFont       As IFont
Dim lFontOld    As Long
    
On Error GoTo ErrF
With UserControl
  .AutoRedraw = True
  Set pFont = IconFont
  lFontOld = SelectObject(.hDC, pFont.hFont)

  If m_MouseOver Then
    .ForeColor = ActiveColor
  Else
    .ForeColor = ForeColor
  End If
  
  Rct.Left = (IconFont.Size / 2) + Pad_X
  Rct.Top = (IconFont.Size / 2) + Pad_Y
  Rct.Right = UserControl.ScaleWidth
  Rct.Bottom = UserControl.ScaleHeight
  
  DrawTextW .hDC, CharCode, 1, Rct, 0
  Call SelectObject(.hDC, lFontOld)
  
ErrF:
  Set pFont = Nothing
End With
End Function

Private Function argb(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
  If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
  argb = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
  Opacity = CByte((Abs(Opacity) / 100) * 255)
  If Opacity < 128 Then
      If Opacity < 0& Then Opacity = 0&
      argb = argb Or Opacity * &H1000000
  Else
      If Opacity > 255& Then Opacity = 255&
      argb = argb Or (Opacity - 128&) * &H1000000 Or &H80000000
  End If
End Function

Private Sub Draw()
Dim REC As RECTL
Dim stREC As RECTS
Dim stREC2 As RECTS
Dim Padding As lPadding
Dim m_PosX1 As Integer, m_PosY1 As Integer
Dim m_PosX2 As Integer, m_PosY2 As Integer
Dim m_PosXA As Integer, m_PosYA As Integer
Dim m_PosXB As Integer, m_PosYB As Integer

With REC
    .Left = m_BorderWidth * nScale
    .Top = m_BorderWidth * nScale
  If m_BorderWidth <> 0 Then
    .Width = (UserControl.ScaleWidth) - (m_BorderWidth * nScale) * 2
    .Height = (UserControl.ScaleHeight) - (m_BorderWidth * nScale) * 2
   Else
    .Width = UserControl.ScaleWidth
    .Height = UserControl.ScaleHeight
   End If
End With

  SafeRange m_Opacity, 0, 100

With UserControl
  GdipCreateFromHDC .hDC, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeHighQuality  'SmoothingModeAntiAlias

  If m_OnFocus Then
    gRoundRect hGraphics, REC, argb(m_Color1, 100), argb(m_Color2, 100), m_Angulo, argb(m_BorderColorOnFocus, 100), m_CornerCurve
  Else
    Select Case m_EffectFade
      Case BorderAndControls, BorderOnly
        gRoundRect hGraphics, REC, argb(m_Color1, 100), argb(m_Color2, 100), m_Angulo, argb(m_BorderColor, m_Opacity), m_CornerCurve
      Case AllPanel
        gRoundRect hGraphics, REC, argb(m_Color1, m_Opacity), argb(m_Color2, m_Opacity), m_Angulo, argb(m_BorderColor, m_Opacity), m_CornerCurve
      Case Else
        gRoundRect hGraphics, REC, argb(m_Color1, 100), argb(m_Color2, 100), m_Angulo, argb(m_BorderColor, 100), m_CornerCurve
    End Select
  End If
  
  '-DRAW CAPTION--------------
  Select Case m_LineOrientation
    Case loVertical
        m_PosX1 = m_Line1Pos
        m_PosY1 = 5
        m_PosX2 = m_Line1Pos
        m_PosY2 = .ScaleHeight - 8
        m_PosXA = m_Line2Pos
        m_PosYA = 5
        m_PosXB = m_Line2Pos
        m_PosYB = .ScaleHeight - 8
        
    Case loHorizontal
        m_PosX1 = 5
        m_PosY1 = m_Line1Pos
        m_PosX2 = .ScaleWidth - 8
        m_PosY2 = m_Line1Pos
        m_PosXA = 5
        m_PosYA = m_Line2Pos
        m_PosXB = .ScaleWidth - 8
        m_PosYB = m_Line2Pos
            
  End Select

  m_Roll = (Font.Size + 15)
    
  If m_Rolled Then
    Dim mRollAngle As Single, mV As eTextAlignV
      
    Select Case m_PinPosition
      Case cLeft, cRight
        stREC.Left = 0: stREC.Top = IIf(m_PinPosition = cLeft, 5, -5)
        stREC.Width = .ScaleHeight: stREC.Height = .ScaleHeight
        mRollAngle = IIf(m_PinPosition = cLeft, 270, 90):   mV = IIf(m_PinPosition = cLeft, eTop, eBottom)
      Case cTop, cBottom
        stREC.Left = 0: stREC.Top = 0
        stREC.Width = .ScaleWidth: stREC.Height = .ScaleHeight
        mRollAngle = 0:   mV = eMiddle
    End Select
    Padding.padX = 0: Padding.padY = 0
    DrawCaption hGraphics, m_RollCaption, stREC, m_CaptionColor, m_CaptionOpacity, mRollAngle, eCenter, mV, Padding, True

  Else
  
    If m_CaptionEnabled Then
      stREC.Left = REC.Left:  stREC.Top = REC.Top
      stREC.Width = REC.Width: stREC.Height = REC.Height
      Padding.padX = m_CaptionX:  Padding.padY = m_CaptionY
      DrawCaption hGraphics, m_Caption, stREC, m_CaptionColor, m_CaptionOpacity, m_CaptionAngle, m_CaptionAlignH, m_CaptionAlignV, Padding, True
    End If
    
     If m_Caption2Enabled Then
      stREC2.Left = REC.Left:  stREC2.Top = REC.Top
      stREC2.Width = REC.Width: stREC2.Height = REC.Height
      Padding.padX = m_Caption2X:  Padding.padY = m_Caption2Y
      DrawCaption hGraphics, m_Caption2, stREC2, m_Caption2Color, m_Caption2Opacity, m_Caption2Angle, m_Caption2AlignH, m_Caption2AlignV, Padding, True
    End If
    
    If m_IconCharCode <> "&H0" Then AddIconChar m_IconCharCode, m_PadX, m_PadY, m_IconForeColor, m_CaptionColor
    If m_IconCharCode2 <> "&H0" Then AddIconChar m_IconCharCode2, m_PadX2, m_PadY2, m_IconForeColor2, m_CaptionColor
    
    If m_Line1 Then
      If m_OnFocus Then
        If m_Line1 Then DrawLine hGraphics, m_PosX1, m_PosY1, m_PosX2, m_PosY2, m_BorderColorOnFocus
      Else
        If m_Line1 Then DrawLine hGraphics, m_PosX1, m_PosY1, m_PosX2, m_PosY2, m_BorderColor
      End If
    End If
    
    If m_Line2 Then
      If m_OnFocus Then
        DrawLine hGraphics, m_PosXA, m_PosYA, m_PosXB, m_PosYB, m_BorderColorOnFocus
      Else
        DrawLine hGraphics, m_PosXA, m_PosYA, m_PosXB, m_PosYB, m_BorderColor
      End If
    End If
  End If
  '---------------
  Select Case m_EffectFade
    Case ControlsOnly, BorderAndControls, AllPanel
        GdipCreateSolidFill argb(m_BorderColor, m_Opacity), hBrush
        GdipCreatePen1 argb(m_BorderColor, m_Opacity), 1, UnitPixel, hPen
    Case Else
        GdipCreateSolidFill argb(m_BorderColor, 100), hBrush
        GdipCreatePen1 argb(m_BorderColor, 100), 1, UnitPixel, hPen
  End Select
      
  If m_CrossVisible Then DrawCross hGraphics
  If m_PinVisible Then DrawPin hGraphics
   '---------------
  Call GdipDeleteBrush(hBrush)
  Call GdipDeletePen(hPen)
  
  GdipDeleteGraphics hGraphics
  '---------------
  .BackStyle = 0
  .MaskColor = .BackColor
  Set .MaskPicture = .Image
End With

End Sub

'*-
Private Function DrawCaption(ByVal hGraphics As Long, sString As String, layoutRect As RECTS, TextColor As OLE_COLOR, ColorOpacity As Integer, mAngle As Single, HAlign As eTextAlignH, VAlign As eTextAlignV, lPadd As lPadding, Optional bWordWrap As Boolean = True) As Long
    Dim hPath As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long

    If GdipCreatePath(&H0, hPath) = 0 Then

        If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
            If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
            GdipSetStringFormatTrimming hFormat, StringTrimmingEllipsisWord
            GdipSetStringFormatAlign hFormat, HAlign
            GdipSetStringFormatLineAlign hFormat, VAlign
        End If

        GetFontStyleAndSize m_CaptionFont, lFontStyle, lFontSize

        If GdipCreateFontFamilyFromName(StrPtr(m_CaptionFont.Name), 0, hFontFamily) Then
            If hFontCollection Then
                If GdipCreateFontFamilyFromName(StrPtr(m_CaptionFont.Name), hFontCollection, hFontFamily) Then
                    If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
                End If
            Else
                If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
            End If
        End If

        With layoutRect
            .Left = IIf(HAlign = eRight, .Left - lPadd.padX * nScale, .Left + lPadd.padX * nScale)
            .Top = IIf(VAlign = eBottom, .Top - lPadd.padY * nScale, .Top - lPadd.padY * nScale)
            .Width = IIf(HAlign = eRight, .Width - (lPadd.padX * nScale), .Width + (lPadd.padX * nScale))
            .Height = IIf(VAlign = eBottom, .Height - (lPadd.padY * nScale), .Height + (lPadd.padY * nScale))
        
            If mAngle <> 0 Then
                Dim newH As Long, newW As Long
                newH = (.Height / 2)  '(UserControl.ScaleHeight / 2)
                newW = (.Width / 2)   '(UserControl.ScaleWidth / 2)
                Call GdipTranslateWorldTransform(hGraphics, newW, newH, 0)
                Call GdipRotateWorldTransform(hGraphics, mAngle, 0)
                Call GdipTranslateWorldTransform(hGraphics, -newW, -newH, 0)
            End If
        End With
        GdipAddPathString hPath, StrPtr(sString), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
        GdipDeleteStringFormat hFormat

        GdipCreateSolidFill argb(TextColor, ColorOpacity), hBrush

        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush

        If mAngle <> 0 Then GdipResetWorldTransform hGraphics

        GdipDeleteFontFamily hFontFamily

        GdipDeletePath hPath
    End If

End Function

Private Sub DrawControl(ByVal iGraphics As Long, mShape As Integer, X As Long, Y As Long, W As Long, H As Long) ', oColor As OLE_COLOR, Opacity As Long)
Dim iPts() As POINTL
'Dim hPen As Long
'Dim hBrush As Long

  If mShape = 0 Then
      ReDim iPts(11)
            
      iPts(0).X = X + (W / 4)
      iPts(0).Y = Y + 0
      iPts(1).X = X + (W / 2)
      iPts(1).Y = Y + (H / 4)
      iPts(2).X = X + (W / 4) * 3
      iPts(2).Y = Y + 0
      iPts(3).X = X + W
      iPts(3).Y = Y + (H / 4)
      iPts(4).X = X + (W / 4) * 3
      iPts(4).Y = Y + (H / 2)
      iPts(5).X = X + W
      iPts(5).Y = Y + (H / 4) * 3
      iPts(6).X = X + (W / 4) * 3
      iPts(6).Y = Y + H
      iPts(7).X = X + (W / 2)
      iPts(7).Y = Y + (H / 4) * 3
      iPts(8).X = X + (W / 4)
      iPts(8).Y = Y + H
      iPts(9).X = X + 0
      iPts(9).Y = Y + (H / 4) * 3
      iPts(10).X = X + (W / 4)
      iPts(10).Y = Y + (H / 2)
      iPts(11).X = X + 0
      iPts(11).Y = Y + H / 4
      
  '---------
  ElseIf mShape = 1 Then
      ReDim iPts(2)
    Select Case m_PinPosition
      Case cTop, cBottom
          iPts(0).X = X + 0
          iPts(0).Y = Y + 0
          iPts(1).X = X + W
          iPts(1).Y = Y + 0
          iPts(2).X = X + (W / 2)
          iPts(2).Y = Y + (H / 4) * 3
      Case cLeft
          iPts(0).X = X + 0
          iPts(0).Y = Y + 0
          iPts(1).X = X + (W / 4) * 3
          iPts(1).Y = Y + (H / 2)
          iPts(2).X = X + 0
          iPts(2).Y = Y + H
      Case cRight
          iPts(0).X = X + 0
          iPts(0).Y = Y + H / 2
          iPts(1).X = X + W
          iPts(1).Y = Y + 0
          iPts(2).X = X + W
          iPts(2).Y = Y + H
    End Select
  End If
          
 ' GdipCreateSolidFill aRGB(oColor, Opacity), hBrush
  GdipFillPolygonI iGraphics, hBrush, iPts(0), UBound(iPts) + 1, &H0
  
 ' GdipCreatePen1 aRGB(oColor, Opacity), 1, UnitPixel, hPen
  GdipDrawPolygonI iGraphics, hPen, iPts(0), UBound(iPts) + 1
          
'  Call GdipDeleteBrush(hBrush)
'  Call GdipDeletePen(hPen)
End Sub

Private Sub DrawCross(ByVal hGraph As Long) ', oColor As OLE_COLOR, Opacity As Long)
Dim vBorder As Long, vCurve As Long
Dim CrossWidth As Long, vSpace As Long

vBorder = IIf(m_BorderWidth < 2, 2, m_BorderWidth)
vCurve = IIf(m_CornerCurve < 2, 2, m_CornerCurve)
CrossWidth = 6
vSpace = IIf((vBorder + (vCurve / 3) + CrossWidth) <= 18, 18, (vBorder + (vCurve / 3) + CrossWidth))

'Cross
Select Case m_CrossPosition
  Case Is = cTopRight
        'YCrossPos = vBorder + CrossWidth
        YCrossPos = (m_Roll / 2) - 3
        XCrossPos = UserControl.ScaleWidth - vSpace
        
  Case Is = cBottomRight
        YCrossPos = UserControl.ScaleHeight - (m_Roll / 2) - 5
        XCrossPos = UserControl.ScaleWidth - vSpace + 1
      
  Case Is = cTopLeft
        YCrossPos = vSpace - CrossWidth
        XCrossPos = (m_Roll / 2) - 4  'vBorder + CrossWidth
                    
  Case Is = cBottomLeft
        YCrossPos = UserControl.ScaleHeight - (BorderWidth + 14)
        XCrossPos = (m_Roll / 2) - 4  'vSpace - CrossWidth
        
End Select

  DrawControl hGraph, 0, XCrossPos, YCrossPos, 7, 7
  
End Sub

Private Sub DrawPin(ByVal hGraph As Long) ', oColor As OLE_COLOR, Opacity As Long)
Dim vBorder As Long, vCurve As Long
Dim PinWidth As Long, vSpace As Long

vBorder = IIf(m_BorderWidth < 2, 2, m_BorderWidth)
vCurve = IIf(m_CornerCurve < 2, 2, m_CornerCurve)
PinWidth = 6

vSpace = IIf((vBorder + (vCurve / 3) + PinWidth) <= 18, 18, (vBorder + (vCurve / 3) + PinWidth))

'Pin
Select Case m_PinPosition
  Case Is = cTop
        YPinPos = (m_Roll / 2) - 1
        XPinPos = UserControl.ScaleWidth - (vSpace + 15)
        DrawControl hGraph, 1, XPinPos, YPinPos, 8, 6 ', oColor, Opacity
                    
  Case Is = cBottom
        YPinPos = UserControl.ScaleHeight - (m_Roll / 2) - 2
        XPinPos = UserControl.ScaleWidth - (vSpace + 15)
        DrawControl hGraph, 1, XPinPos, YPinPos, 8, 6 ', oColor, Opacity
        
  Case Is = cLeft
        YPinPos = vSpace + PinWidth
        XPinPos = (m_Roll / 2) - 2
        DrawControl hGraph, 1, XPinPos, YPinPos, 6, 8 ', oColor, Opacity
                    
  Case Is = cRight
        YPinPos = UserControl.ScaleHeight - (vSpace + 16)
        XPinPos = UserControl.ScaleWidth - (m_Roll / 2) - 5
        DrawControl hGraph, 1, XPinPos, YPinPos, 6, 8 ', oColor, Opacity
            
End Select
  
End Sub

Private Function GetFontStyleAndSize(oFont As StdFont, lFontStyle As Long, lFontSize As Long)
On Error GoTo ErrO
    Dim hDC As Long
    lFontStyle = 0
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
    
    hDC = GetDC(0&)
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hDC, LOGPIXELSY), 72)
    ReleaseDC 0&, hDC
ErrO:
End Function

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function

Private Function GetSystemHandCursor() As Picture
  Dim Pic As PicBmp
  Dim IPic As IPicture
  Dim GUID(0 To 3) As Long
  
  If hCur Then DestroyCursor hCur: hCur = 0
  
  hCur = LoadCursor(ByVal 0&, IDC_HAND)
   
  GUID(0) = &H7BF80980
  GUID(1) = &H101ABF32
  GUID(2) = &HAA00BB8B
  GUID(3) = &HAB0C3000
  
  With Pic
    .Size = Len(Pic)
    .type = vbPicTypeIcon
    .hBmp = hCur
    .hPal = 0
  End With
  
  Call OleCreatePictureIndirect(Pic, GUID(0), 1, IPic)
  
  Set GetSystemHandCursor = IPic
End Function

Private Function GetWindowsDPI() As Double
    Dim hDC As Long, lPx  As Double, LPY As Double
    hDC = GetDC(0)
    lPx = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    LPY = CDbl(GetDeviceCaps(hDC, LOGPIXELSY))
    ReleaseDC 0, hDC

    If (lPx = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = lPx / 96#
    End If
End Function

Private Function gRoundRect(ByVal hGraphics As Long, Rect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal Angulo As Single, ByVal BorderColor As Long, Round As Long) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mpath As Long
    Dim mRound As Long
    
    If m_BorderWidth <> 0 Then
      GdipCreatePen1 BorderColor, m_BorderWidth * nScale, &H2, hPen  '&H1 * nScale, &H2, hPen
    End If
    GdipCreateLineBrushFromRectWithAngleI Rect, color1, color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mpath   '&H0
    
    With Rect
        mRound = GetSafeRound((Round * nScale), .Width, .Height)
        If mRound = 0 Then mRound = 1
        '    GdipDrawRectangleI hGraphics, hPen, .Left, .Top, .Width, .Height
        '    GdipAddPathLineI mPath, .Left, .Top, .Width, .Top       'Line-Top
        '    GdipAddPathLineI mPath, .Width, .Top, .Width, .Height   'Line-Left
        '    GdipAddPathLineI mPath, .Width, .Height, .Left, .Height 'Line-Bottom
        '    GdipAddPathLineI mPath, .Left, .Height, .Left, .Top     'Line-Right
        'Else
            GdipAddPathArcI mpath, .Left, .Top, mRound, mRound, 180, 90
            GdipAddPathArcI mpath, (.Left + .Width) - mRound, .Top, mRound, mRound, 270, 90
            GdipAddPathArcI mpath, (.Left + .Width) - mRound, (.Top + .Height) - mRound, mRound, mRound, 0, 90
            GdipAddPathArcI mpath, .Left, (.Top + .Height) - mRound, mRound, mRound, 90, 90
        'End If
    End With
    
    GdipClosePathFigures mpath
    GdipFillPath hGraphics, hBrush, mpath
    GdipDrawPath hGraphics, hPen, mpath
    
    Call GdipDeletePath(mpath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)

    gRoundRect = mpath
End Function

'Inicia GDI+
Private Sub InitGDI()
    Dim gdipStartupInput As GdiplusStartupInput
    gdipStartupInput.GdiplusVersion = 1&
    Call GdiplusStartup(gdipToken, gdipStartupInput, ByVal 0)
End Sub

Private Function IsMouseOver(hWnd As Long) As Boolean
    Dim PT As POINTL
    GetCursorPos PT
    IsMouseOver = (WindowFromPoint(PT.X, PT.Y) = hWnd)
End Function

Private Function MousePointerHands(ByVal NewValue As Boolean)
  If NewValue Then
    If Ambient.UserMode Then
      UserControl.MousePointer = vbCustom
      UserControl.MouseIcon = GetSystemHandCursor
    End If
  Else
    If hCur Then DestroyCursor hCur: hCur = 0
    UserControl.MousePointer = vbDefault
    UserControl.MouseIcon = Nothing
  End If

End Function

Private Function ReadValue(ByVal lProp As Long, Optional Default As Long) As Long
    Dim I       As Long
    For I = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(I) = lProp Then
            ReadValue = TlsGetValue(I + 1)
            Exit Function
        End If
    Next
    ReadValue = Default
End Function

Private Sub RollExtender(ByVal isRoll As Boolean)

With UserControl
    Select Case m_PinPosition
        Case cLeft
          If Not isRoll Then
            .Width = m_FullWidth
          Else
            .Width = m_Roll * Screen.TwipsPerPixelX
          End If
        
        Case cRight
          If Not isRoll Then
            .Width = m_FullWidth
            Extender.Left = m_Left + 11
          Else
            .Width = m_Roll * Screen.TwipsPerPixelX
            Extender.Left = m_Left + (m_FullWidth - .Width) / Screen.TwipsPerPixelX
          End If
        
        Case cTop
          If Not isRoll Then
            UserControl.Height = m_FullHeight
          Else
            UserControl.Height = m_Roll * Screen.TwipsPerPixelY
          End If
          
        Case cBottom
          If Not isRoll Then
            UserControl.Height = m_FullHeight
            Extender.Top = m_Top + 11
          Else
            UserControl.Height = m_Roll * Screen.TwipsPerPixelY
            Extender.Top = m_Top + (m_FullHeight - .Height) / Screen.TwipsPerPixelX
          End If
    
    End Select
End With
End Sub

Private Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub

'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(gdipToken)
End Sub

Private Sub tmrEffect_Timer()
If IsMouseOver(UserControl.hWnd) Then
  If m_Opacity < 100 Then
    m_Opacity = m_Opacity + 2
    Refresh
  Else
    Exit Sub
  End If
Else
  m_Opacity = m_InitialOpacity
  Refresh
  tmrEffect.Enabled = False
End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  Refresh
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    InitGDI
    nScale = GetWindowsDPI
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
hFontCollection = ReadValue(&HFC)
cl_hWnd = UserControl.ContainerHwnd

 m_CaptionAngle = 0
 m_CaptionX = 0
 m_CaptionY = 0
 m_Caption = UserControl.Ambient.DisplayName
 m_CaptionEnabled = True
 m_CaptionAlignV = eTop
 m_CaptionAlignH = eLeft
 m_CaptionOpacity = 100
 Set m_CaptionFont = UserControl.Ambient.Font
 m_CaptionColor = vbBlack

 m_Caption2Angle = 0
 m_Caption2X = 0
 m_Caption2Y = 0
 m_Caption2 = UserControl.Ambient.DisplayName
 m_Caption2Enabled = True
 m_Caption2AlignV = eTop
 m_Caption2AlignH = eLeft
 m_Caption2Opacity = 100
 Set m_Caption2Font = UserControl.Ambient.Font
 m_Caption2Color = vbBlack

  m_BorderColor = &HC0&
  m_OldBorderColor = &HC0&
  m_ActiveColor = &HFFFFFF
  m_Enabled = True
  m_Color1 = m_def_Color1
  m_Color2 = m_def_Color2
  m_Angulo = m_def_Angulo
  m_BorderWidth = 1
  m_CornerCurve = 3
  m_LineOrientation = loHorizontal
  m_Line1 = False
  m_Line2 = False
  m_Line2Pos = UserControl.ScaleHeight / 2
  m_Line1Pos = UserControl.ScaleHeight / 3
  m_Rolled = False
  m_Opacity = 10
  m_IconCharCode = "&H0"
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_OnFocus = True
  
  OldX = X
  OldY = Y
              
  If X > XPinPos And X < (XPinPos + 6) And Y > YPinPos And Y < (YPinPos + 6) Then
      If m_PinVisible Then
          m_Rolled = Not m_Rolled
          If m_Rolled Then
            m_FullWidth = UserControl.Width
            m_FullHeight = UserControl.Height
            m_Top = Extender.Top
            m_Left = Extender.Left
          End If
          RaiseEvent PinClick
          RollExtender m_Rolled
      End If
  ElseIf X > XCrossPos And X < XCrossPos + 6 And Y > (YCrossPos) And Y < (YCrossPos) + 6 Then
    If m_CrossVisible Then
      RaiseEvent CrossClick
      Extender.Visible = False
    End If
  End If
  
Refresh
  
RaiseEvent MouseDown(Button, Shift, X, Y)

If m_Rolled Then Debug.Print "---UserControl_MouseDown---" & vbLf & _
                              "m_Left:" & m_Left & vbLf & _
                              "m_Top:" & m_Top & vbLf & _
                              "---UserControl_MouseDown---" & vbLf & _
                              "RollExtender " & m_Rolled
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton And m_Moveable = True Then
  Dim res As Long
  MousePointerHands True
  Call ReleaseCapture
  res = SendMessage(UserControl.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  
  If Y > OldY Then m_Top = m_Top + Y
  If Y < OldY Then m_Top = m_Top - Y
  
  If X > OldX Then m_Left = m_Left + X
  If X < OldX Then m_Left = m_Left - X
  
  Debug.Print "---UserControl_MouseMove---" & vbLf & "m_Top:" & m_Top & " - m_Left:" & m_Left
End If

  If X > XCrossPos And X < (XCrossPos + 6) And Y > YCrossPos And Y < (YCrossPos + 6) Then
      If m_CrossVisible Then
        MousePointerHands True
      End If
  ElseIf X > XPinPos And X < (XPinPos + 8) And Y > YPinPos And Y < (YPinPos + 6) Then
      If m_PinVisible Then
        MousePointerHands True
      End If
  Else
      MousePointerHands False
  End If

tmrEffect.Enabled = True
RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
m_OnFocus = False
Refresh
MousePointerHands False
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
Refresh
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

With PropBag
  m_Enabled = .ReadProperty("Enabled", True)
  m_Color1 = .ReadProperty("BackColor1", m_def_Color1)
  m_Color2 = .ReadProperty("BackColor2", m_def_Color2)
  m_ActiveColor = .ReadProperty("ActiveColor", m_def_ForeColor)
  m_Angulo = .ReadProperty("Angulo", m_def_Angulo)
  m_BorderColor = .ReadProperty("BorderColor", &HC0&)
  m_BorderWidth = .ReadProperty("BorderWidth", 1)
  m_CornerCurve = .ReadProperty("CornerCurve", 0)
  
  m_CrossPosition = .ReadProperty("CrossPosition", cTopRight)
  m_CrossVisible = .ReadProperty("CrossVisible", False)
  m_PinVisible = .ReadProperty("PinVisible", False)
  m_PinPosition = .ReadProperty("PinPosition", cTop)
  
  m_Moveable = .ReadProperty("Moveable", False)
  
  m_LineOrientation = .ReadProperty("LineOrientation", loHorizontal)
  m_Line1 = .ReadProperty("Line1", False)
  m_Line2 = .ReadProperty("Line2", False)
  m_Line1Pos = .ReadProperty("Line1Pos", UserControl.ScaleHeight / 3)
  m_Line2Pos = .ReadProperty("Line2Pos", UserControl.ScaleHeight / 2)
  
  m_RollCaption = .ReadProperty("RollCaption", Ambient.DisplayName)
  
  Set m_CaptionFont = .ReadProperty("Caption1Font", UserControl.Ambient.Font)
  m_CaptionColor = .ReadProperty("Caption1Color", m_def_ForeColor)
  m_Caption = .ReadProperty("Caption1", Ambient.DisplayName)
  m_CaptionEnabled = .ReadProperty("Caption1Enabled", False)
  m_CaptionAngle = .ReadProperty("Caption1Agle", 0)
  m_CaptionX = .ReadProperty("Caption1X", 0)
  m_CaptionY = .ReadProperty("Caption1Y", 0)
  m_CaptionAlignV = .ReadProperty("Caption1AlignV", 0)
  m_CaptionAlignH = .ReadProperty("Caption1AlignH", 0)
  m_CaptionOpacity = .ReadProperty("Caption1Opacity", 0)

  Set m_Caption2Font = .ReadProperty("Caption2Font", UserControl.Ambient.Font)
  m_Caption2Color = .ReadProperty("Caption2Color", m_def_ForeColor)
  m_Caption2 = .ReadProperty("Caption2", Ambient.DisplayName)
  m_Caption2Enabled = .ReadProperty("Caption2Enabled", False)
  m_Caption2Angle = .ReadProperty("Caption2Angle", 0)
  m_Caption2X = .ReadProperty("Caption2X", 0)
  m_Caption2Y = .ReadProperty("Caption2Y", 0)
  m_Caption2AlignV = .ReadProperty("Caption2AlignV", 0)
  m_Caption2AlignH = .ReadProperty("Caption2AlignH", 0)
  m_Caption2Opacity = .ReadProperty("Caption2Opacity", 0)

  m_BorderColorOnFocus = .ReadProperty("BorderColorOnFocus", vbWhite)
  m_ChangeBorderOnFocus = .ReadProperty("ChangeBorderOnFocus", False)
  m_EffectFade = .ReadProperty("EffectFading", None)
  m_InitialOpacity = .ReadProperty("InitialOpacity", 30)
  
  Set m_IconFont = .ReadProperty("IconFont", Ambient.Font)
  m_IconCharCode = .ReadProperty("Icon1CharCode", "&H0")
  m_IconForeColor = .ReadProperty("Icon1ForeColor", &H404040)
  m_PadX = .ReadProperty("Icon1PaddingX", 0)
  m_PadY = .ReadProperty("Icon1PaddingY", 0)

  m_IconCharCode2 = .ReadProperty("Icon2CharCode", "&H0")
  m_IconForeColor2 = .ReadProperty("Icon2ForeColor", &H404040)
  m_PadX2 = .ReadProperty("Icon2PaddingX", 0)
  m_PadY2 = .ReadProperty("Icon2PaddingY", 0)

End With
  
  m_Opacity = m_InitialOpacity

End Sub

Private Sub UserControl_Resize()
Refresh
End Sub

Private Sub UserControl_Show()
Refresh
End Sub

Private Sub UserControl_Terminate()
TerminateGDI
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("Enabled", m_Enabled)
  Call .WriteProperty("BackColor1", m_Color1, m_def_Color1)
  Call .WriteProperty("BackColor2", m_Color2, m_def_Color2)
  Call .WriteProperty("ActiveColor", m_ActiveColor, m_def_ForeColor)
  Call .WriteProperty("Angulo", m_Angulo, m_def_Angulo)
  Call .WriteProperty("BorderColor", m_BorderColor, &HC0&)
  Call .WriteProperty("BorderWidth", m_BorderWidth, 1)
  Call .WriteProperty("CornerCurve", m_CornerCurve, 0)
  
  Call .WriteProperty("CrossPosition", m_CrossPosition, cTopRight)
  Call .WriteProperty("CrossVisible", m_CrossVisible, False)
  Call .WriteProperty("PinPosition", m_PinPosition, cTop)
  Call .WriteProperty("PinVisible", m_PinVisible, False)
  
  Call .WriteProperty("Moveable", m_Moveable, False)
  
  Call .WriteProperty("LineOrientation", m_LineOrientation, loHorizontal)
  Call .WriteProperty("Line1", m_Line1, False)
  Call .WriteProperty("Line2", m_Line2, False)
  Call .WriteProperty("Line1Pos", m_Line1Pos, UserControl.ScaleHeight / 3)
  Call .WriteProperty("Line2Pos", m_Line2Pos, UserControl.ScaleHeight / 2)
  
  Call .WriteProperty("RollCaption", m_RollCaption)
  
  Call .WriteProperty("Caption1Font", m_CaptionFont, UserControl.Ambient.Font)
  Call .WriteProperty("Caption1Color", m_CaptionColor, m_def_ForeColor)
  Call .WriteProperty("Caption1", m_Caption, Ambient.DisplayName)
  Call .WriteProperty("Caption1Enabled", m_CaptionEnabled, False)
  Call .WriteProperty("Caption1Agle", m_CaptionAngle)
  Call .WriteProperty("Caption1X", m_CaptionX)
  Call .WriteProperty("Caption1Y", m_CaptionY)
  Call .WriteProperty("Caption1AlignV", m_CaptionAlignV)
  Call .WriteProperty("Caption1AlignH", m_CaptionAlignH)
  Call .WriteProperty("Caption1Opacity", m_CaptionOpacity)
  
  Call .WriteProperty("Caption2Font", m_Caption2Font, UserControl.Ambient.Font)
  Call .WriteProperty("Caption2Color", m_Caption2Color, m_def_ForeColor)
  Call .WriteProperty("Caption2", m_Caption2, Ambient.DisplayName)
  Call .WriteProperty("Caption2Enabled", m_Caption2Enabled, False)
  Call .WriteProperty("Caption2Angle", m_Caption2Angle)
  Call .WriteProperty("Caption2X", m_Caption2X)
  Call .WriteProperty("Caption2Y", m_Caption2Y)
  Call .WriteProperty("Caption2AlignV", m_Caption2AlignV)
  Call .WriteProperty("Caption2AlignH", m_Caption2AlignH)
  Call .WriteProperty("Caption2Opacity", m_Caption2Opacity)
  
  Call .WriteProperty("BorderColorOnFocus", m_BorderColorOnFocus, vbWhite)
  Call .WriteProperty("ChangeBorderOnFocus", m_ChangeBorderOnFocus, False)
  Call .WriteProperty("EffectFading", m_EffectFade, None)
  Call .WriteProperty("InitialOpacity", m_InitialOpacity, 30)
  
  Call .WriteProperty("IconFont", m_IconFont)
  Call .WriteProperty("Icon1CharCode", m_IconCharCode, 0)
  Call .WriteProperty("Icon1ForeColor", m_IconForeColor, vbButtonText)
  Call .WriteProperty("Icon1PaddingX", m_PadX)
  Call .WriteProperty("Icon1PaddingY", m_PadY)

  Call .WriteProperty("Icon2CharCode", m_IconCharCode2, 0)
  Call .WriteProperty("Icon2ForeColor", m_IconForeColor2, vbButtonText)
  Call .WriteProperty("Icon2PaddingX", m_PadX2)
  Call .WriteProperty("Icon2PaddingY", m_PadY2)

End With
  
End Sub

Public Property Get ActiveColor() As OLE_COLOR
  ActiveColor = m_ActiveColor
End Property

Public Property Let ActiveColor(ByVal NewForeColor As OLE_COLOR)
  m_ActiveColor = NewForeColor
  PropertyChanged "ActiveColor"
  Refresh
End Property

Public Property Get Angulo() As Single
  Angulo = m_Angulo
End Property

Public Property Let Angulo(ByVal New_Angulo As Single)
  m_Angulo = New_Angulo
  PropertyChanged "Angulo"
  Refresh
End Property

Public Property Get BackColor1() As OLE_COLOR
  BackColor1 = m_Color1
End Property

Public Property Let BackColor1(ByVal New_Color1 As OLE_COLOR)
  m_Color1 = New_Color1
  PropertyChanged "BackColor1"
  Refresh
End Property

Public Property Get BackColor2() As OLE_COLOR
  BackColor2 = m_Color2
End Property

Public Property Let BackColor2(ByVal New_Color2 As OLE_COLOR)
  m_Color2 = New_Color2
  PropertyChanged "BackColor2"
  Refresh
End Property

'Properties-------------------
Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
  m_BorderColor = NewBorderColor
  'm_OldBorderColor = m_BorderColor
  PropertyChanged "BorderColor"
  Refresh
End Property

Public Property Get BorderColorOnFocus() As OLE_COLOR
  BorderColorOnFocus = m_BorderColorOnFocus
End Property

Public Property Let BorderColorOnFocus(ByVal NewBorderColorOnFocus As OLE_COLOR)
  m_BorderColorOnFocus = NewBorderColorOnFocus
  PropertyChanged "BorderColorOnFocus"
End Property

Public Property Get BorderWidth() As Long
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  Refresh
End Property
'--CAPTION1------------
Public Property Get Caption1AlignH() As eTextAlignH
  Caption1AlignH = m_CaptionAlignH
End Property

Public Property Let Caption1AlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_CaptionAlignH = NewCaptionAlignH
  PropertyChanged "Caption1AlignH"
  Refresh
End Property

Public Property Get Caption1AlignV() As eTextAlignV
  Caption1AlignV = m_CaptionAlignV
End Property

Public Property Let Caption1AlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_CaptionAlignV = NewCaptionAlignV
  PropertyChanged "Caption1AlignV"
  Refresh
End Property

Public Property Get Caption1Angle() As Single
    Caption1Angle = m_CaptionAngle
End Property

Public Property Let Caption1Angle(ByVal New_Angle As Single)
    m_CaptionAngle = New_Angle
    PropertyChanged "Caption1Angle"
    Refresh
End Property

Public Property Get Caption1Color() As OLE_COLOR
  Caption1Color = m_CaptionColor
End Property

Public Property Let Caption1Color(ByVal NewForeColor As OLE_COLOR)
  m_CaptionColor = NewForeColor
  PropertyChanged "Caption1Color"
  Refresh
End Property

Public Property Get Caption1Enabled() As Boolean
  Caption1Enabled = m_CaptionEnabled
End Property

Public Property Let Caption1Enabled(ByVal NewCaptionEnabled As Boolean)
  m_CaptionEnabled = NewCaptionEnabled
  PropertyChanged "Caption1Enabled"
  Refresh
End Property

Public Property Get Caption1Font() As StdFont
  Set Caption1Font = m_CaptionFont
End Property

Public Property Set Caption1Font(ByVal New_Font As StdFont)
  Set m_CaptionFont = New_Font
  PropertyChanged "Caption1Font"
  Refresh
End Property

Public Property Get Caption1() As String
  Caption1 = m_Caption
End Property

Public Property Let Caption1(ByVal NewCaption As String)
  m_Caption = NewCaption
  PropertyChanged "Caption1"
  Refresh
End Property

Public Property Get Caption1Opacity() As Integer
  Caption1Opacity = m_CaptionOpacity
End Property

Public Property Let Caption1Opacity(ByVal NewOpa As Integer)
  m_CaptionOpacity = NewOpa
  PropertyChanged "Caption1Opacity"
  Refresh
End Property

Public Property Get Caption1X() As Integer
  Caption1X = m_CaptionX
End Property

Public Property Let Caption1X(ByVal NewPosX As Integer)
  m_CaptionX = NewPosX
  PropertyChanged "Caption1X"
  Refresh
End Property

Public Property Get Caption1Y() As Integer
  Caption1Y = m_CaptionY
End Property

Public Property Let Caption1Y(ByVal NewPosY As Integer)
  m_CaptionY = NewPosY
  PropertyChanged "Caption1Y"
  Refresh
End Property
'--CAPTION2-------------
Public Property Get Caption2AlignH() As eTextAlignH
  Caption2AlignH = m_Caption2AlignH
End Property

Public Property Let Caption2AlignH(ByVal NewCaption2AlignH As eTextAlignH)
  m_Caption2AlignH = NewCaption2AlignH
  PropertyChanged "Caption2AlignH"
  Refresh
End Property

Public Property Get Caption2AlignV() As eTextAlignV
  Caption2AlignV = m_Caption2AlignV
End Property

Public Property Let Caption2AlignV(ByVal NewCaption2AlignV As eTextAlignV)
  m_Caption2AlignV = NewCaption2AlignV
  PropertyChanged "Caption2AlignV"
  Refresh
End Property

Public Property Get Caption2Angle() As Single
    Caption2Angle = m_Caption2Angle
End Property

Public Property Let Caption2Angle(ByVal New_Angle As Single)
    m_Caption2Angle = New_Angle
    PropertyChanged "Caption2Angle"
    Refresh
End Property

Public Property Get Caption2Color() As OLE_COLOR
  Caption2Color = m_Caption2Color
End Property

Public Property Let Caption2Color(ByVal NewForeColor As OLE_COLOR)
  m_Caption2Color = NewForeColor
  PropertyChanged "Caption2Color"
  Refresh
End Property

Public Property Get Caption2Enabled() As Boolean
  Caption2Enabled = m_Caption2Enabled
End Property

Public Property Let Caption2Enabled(ByVal NewCaption2Enabled As Boolean)
  m_Caption2Enabled = NewCaption2Enabled
  PropertyChanged "Caption2Enabled"
  Refresh
End Property

Public Property Get Caption2Font() As StdFont
  Set Caption2Font = m_Caption2Font
End Property

Public Property Set Caption2Font(ByVal New_Font As StdFont)
  Set m_Caption2Font = New_Font
  PropertyChanged "Caption2Font"
  Refresh
End Property

Public Property Get Caption2() As String
  Caption2 = m_Caption2
End Property

Public Property Let Caption2(ByVal NewCaption2 As String)
  m_Caption2 = NewCaption2
  PropertyChanged "Caption2"
  Refresh
End Property

Public Property Get Caption2Opacity() As Integer
  Caption2Opacity = m_Caption2Opacity
End Property

Public Property Let Caption2Opacity(ByVal NewOpa As Integer)
  m_Caption2Opacity = NewOpa
  PropertyChanged "Caption2Opacity"
  Refresh
End Property

Public Property Get Caption2X() As Integer
  Caption2X = m_Caption2X
End Property

Public Property Let Caption2X(ByVal NewPosX As Integer)
  m_Caption2X = NewPosX
  PropertyChanged "Caption2X"
  Refresh
End Property

Public Property Get Caption2Y() As Integer
  Caption2Y = m_Caption2Y
End Property

Public Property Let Caption2Y(ByVal NewPosY As Integer)
  m_Caption2Y = NewPosY
  PropertyChanged "Caption2Y"
  Refresh
End Property
'----------------------
Public Property Get ChangeBorderOnFocus() As Boolean
  ChangeBorderOnFocus = m_ChangeBorderOnFocus
End Property

Public Property Let ChangeBorderOnFocus(ByVal NewChangeBorderOnFocus As Boolean)
  m_ChangeBorderOnFocus = NewChangeBorderOnFocus
  m_OldBorderColor = m_BorderColor
  PropertyChanged "ChangeBorderOnFocus"
End Property

Public Property Get CornerCurve() As Long
  CornerCurve = m_CornerCurve
End Property

Public Property Let CornerCurve(ByVal NewCornerCurve As Long)
  m_CornerCurve = NewCornerCurve
  PropertyChanged "CornerCurve"
  Refresh
End Property

'm_CrossPosition
Public Property Get CrossPosition() As CrossPos
    CrossPosition = m_CrossPosition
End Property

Public Property Let CrossPosition(ByVal New_Value As CrossPos)
    m_CrossPosition = New_Value
    PropertyChanged "CrossPosition"
    Refresh
End Property

'm_CrossVisible
Public Property Get CrossVisible() As Boolean
    CrossVisible = m_CrossVisible
End Property

Public Property Let CrossVisible(ByVal New_Value As Boolean)
    m_CrossVisible = New_Value
    PropertyChanged "CrossVisible"
    Refresh
End Property

Public Property Get EffectFading() As FadeEffect
EffectFading = m_EffectFade
End Property

Public Property Let EffectFading(ByVal vNewValue As FadeEffect)
m_EffectFade = vNewValue
PropertyChanged "EffectFading"
Refresh
End Property

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  UserControl.Enabled = m_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Icon1CharCode() As String
    Icon1CharCode = "&H" & Hex(m_IconCharCode)
End Property

Public Property Let Icon1CharCode(ByVal New_IconCharCode As String)
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not VBA.Left$(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        m_IconCharCode = "&H" & New_IconCharCode
    Else
        m_IconCharCode = New_IconCharCode
    End If
    PropertyChanged "Icon1CharCode"
    Refresh
End Property

Public Property Get Icon2CharCode() As String
    Icon2CharCode = "&H" & Hex(m_IconCharCode2)
End Property

Public Property Let Icon2CharCode(ByVal New_IconCharCode As String)
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not VBA.Left$(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        m_IconCharCode2 = "&H" & New_IconCharCode
    Else
        m_IconCharCode2 = New_IconCharCode
    End If
    PropertyChanged "Icon2CharCode"
    Refresh
End Property

Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property

Public Property Set IconFont(New_Font As StdFont)
    With m_IconFont
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Strikethrough = New_Font.Strikethrough
        .Underline = New_Font.Underline
        .Weight = New_Font.Weight
    End With
    PropertyChanged "IconFont"
  Refresh
End Property

Public Property Get Icon1ForeColor() As OLE_COLOR
    Icon1ForeColor = m_IconForeColor
End Property

Public Property Let Icon1ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor = New_ForeColor
    PropertyChanged "Icon1ForeColor"
    Refresh
End Property

Public Property Get Icon2ForeColor() As OLE_COLOR
    Icon2ForeColor = m_IconForeColor2
End Property

Public Property Let Icon2ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor2 = New_ForeColor
    PropertyChanged "Icon2ForeColor"
    Refresh
End Property

Public Property Get Icon1PaddingX() As Long
Icon1PaddingX = m_PadX
End Property

Public Property Let Icon1PaddingX(ByVal XpadVal As Long)
m_PadX = XpadVal
PropertyChanged "Icon1PaddingX"
Refresh
End Property

Public Property Get Icon1PaddingY() As Long
Icon1PaddingY = m_PadY
End Property

Public Property Let Icon1PaddingY(ByVal YpadVal As Long)
m_PadY = YpadVal
PropertyChanged "Icon1PaddingY"
Refresh
End Property

Public Property Get Icon2PaddingX() As Long
Icon2PaddingX = m_PadX2
End Property

Public Property Let Icon2PaddingX(ByVal XpadVal As Long)
m_PadX2 = XpadVal
PropertyChanged "Icon2PaddingX"
Refresh
End Property

Public Property Get Icon2PaddingY() As Long
Icon2PaddingY = m_PadY2
End Property

Public Property Let Icon2PaddingY(ByVal YpadVal As Long)
m_PadY2 = YpadVal
PropertyChanged "Icon2PaddingY"
Refresh
End Property

Public Property Get InitialOpacity() As Long
  InitialOpacity = m_InitialOpacity
End Property

Public Property Let InitialOpacity(ByVal NewInitialOpacity As Long)
  m_InitialOpacity = NewInitialOpacity
  PropertyChanged "InitialOpacity"
  m_Opacity = m_InitialOpacity
  Refresh
End Property

Public Property Get Line1() As Boolean
  Line1 = m_Line1
End Property

Public Property Let Line1(ByVal bLine1 As Boolean)
  m_Line1 = bLine1
  PropertyChanged "Line1"
  Refresh
End Property

Public Property Get Line1Pos() As Integer
  Line1Pos = (m_Line1Pos)
End Property

Public Property Let Line1Pos(ByVal bLine1pos As Integer)
  m_Line1Pos = (bLine1pos)
  PropertyChanged "Line1Pos"
  Refresh
End Property

Public Property Get Line2() As Boolean
  Line2 = m_Line2
End Property

Public Property Let Line2(ByVal bLine2 As Boolean)
  m_Line2 = bLine2
  PropertyChanged "Line2"
  Refresh
End Property

Public Property Get Line2Pos() As Integer
  Line2Pos = (m_Line2Pos)
End Property

Public Property Let Line2Pos(ByVal bLine2pos As Integer)
  m_Line2Pos = (bLine2pos)
  PropertyChanged "Line2Pos"
  Refresh
End Property

Public Property Get LineOrientation() As LineOr
  LineOrientation = m_LineOrientation
End Property

Public Property Let LineOrientation(ByVal NewOri As LineOr)
  m_LineOrientation = NewOri
  PropertyChanged "LineOrientation"
  Refresh
End Property

Public Property Get Moveable() As Boolean
    Moveable = m_Moveable
End Property

Public Property Let Moveable(ByVal New_Moveable As Boolean)
    m_Moveable = New_Moveable
    PropertyChanged "Moveable"
End Property

Public Property Get PinPosition() As CheckPos
    PinPosition = m_PinPosition
End Property

Public Property Let PinPosition(ByVal New_Value As CheckPos)
    m_PinPosition = New_Value
    PropertyChanged "PinPosition"
    Refresh
End Property

Public Property Get PinVisible() As Boolean
    PinVisible = m_PinVisible
End Property

Public Property Let PinVisible(ByVal New_Value As Boolean)
    m_PinVisible = New_Value
    PropertyChanged "PinVisible"
    Refresh
End Property

Public Property Get RollCaption() As String
  RollCaption = m_RollCaption
End Property

Public Property Let RollCaption(ByVal NewCaption As String)
  m_RollCaption = NewCaption
  PropertyChanged "RollCaption"
  Refresh
End Property

Public Property Get Version() As String
Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Property Get Visible() As Boolean
  Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal NewVisible As Boolean)
  Extender.Visible = NewVisible
End Property

