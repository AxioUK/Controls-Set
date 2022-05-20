VERSION 5.00
Begin VB.UserControl AxGButtonLabel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H008D4214&
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
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
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   75
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   282
   ToolboxBitmap   =   "AxGButtonLabel.ctx":0000
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   75
      Top             =   75
   End
End
Attribute VB_Name = "AxGButtonLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxGButtonLabel
'Version  : 2.07.6
'Editor   : David Rojas [AxioUK]
'Date     : 19/05/2022
'------------------------------------
Option Explicit

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_SYSKEYDOWN As Long = &H104 'WM_SYSKEYDOWN = &H104&
Private Const WM_SYSKEYUP As Long = &H105
Private Const WM_ACTIVATEAPP As Long = &H1C
'---
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'-
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

'Private Declare Function ReleaseCapture Lib "User32" () As Long
'Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal token As Long)
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
'Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mpath As Long) As Long
'Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef brush As Long) As Long
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTS, ByVal mFormat As Long) As Long
'Private Declare Function GdipMeasureString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByRef mBoundingBox As RECTS, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
'Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
'Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
'Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As eStringTrimming) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatHotkeyPrefix Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mHotkeyPrefix As HotkeyPrefix) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As eStringFormatFlags) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As eStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As eStringAlignment) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
'Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal count As Long) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByRef pPoints As Any, ByVal count As Long, ByVal FillMode As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long) As Long
Private Declare Function GdipSetClipRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipResetPath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long

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

'EVENTS------------------------------------
Public Event Click()
Public Event ChangeValue(ByVal Value As Boolean)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
''-----------------------------------------

Private Const CombineModeExclude As Long = &H4
Private Const WrapModeTileFlipXY = &H3
'Private Const SmoothingModeHighQuality As Long = &H2
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const TLS_MINIMUM_AVAILABLE As Long = 64
'Private Const CLR_INVALID = -1
'Private Const WM_NCLBUTTONDOWN = &HA1
'Private Const HTCAPTION = 2
Private Const IDC_HAND As Long = 32649
Private Const UnitPixel As Long = &H2&

'Default Property Values:
'Const m_def_ForeColor = 0
'Const m_def_Appearance = 0
'Const m_def_BackStyle = 0
'Const m_def_BorderStyle = 0
'Const m_def_Color1 = &HD59B5B
'Const m_def_Color2 = &H6A5444
'Const m_def_Angulo = 0

'Property Variables:
'Private cSC As c_SubClass
'Private ucchWnd As Long

Dim gdipToken As Long
Dim nScale    As Single
Dim hCur      As Long
Dim hFontCollection As Long
Dim hGraphics As Long
Dim hPen      As Long
Dim hBrush    As Long

Dim m_BackGroundColor As OLE_COLOR
Dim m_BorderColor   As OLE_COLOR
Dim m_ForeColor     As OLE_COLOR
Dim m_ForeColor2    As OLE_COLOR
Dim m_Color1        As OLE_COLOR
Dim m_Color2        As OLE_COLOR
Dim m_ColorActive1  As OLE_COLOR
Dim m_ColorActive2  As OLE_COLOR
Dim m_BorderWidth   As Long
Dim m_Enabled       As Boolean
Dim m_Angulo        As Single
Dim m_CornerCurve   As Long
Dim m_CaptionAngle  As Single
Dim m_CaptionX      As Long
Dim m_CaptionY      As Long
Dim m_Caption       As String
Dim m_Value         As Boolean
Dim m_Clicked       As Boolean
Dim m_Clickable     As Boolean
Dim m_OptionButton  As Boolean
Dim m_Filled        As Boolean
Dim m_ButtonEffect  As eButtonEffect
Dim m_PressEffect   As ePressEffect
Dim m_FlatEffect    As eFlatEffect
Dim m_Opacity       As Long
Dim m_OldBorderColor      As OLE_COLOR
Dim m_BorderColorOnFocus  As OLE_COLOR
Dim m_ChangeBorderOnFocus As Boolean
Dim m_ChangeBorderMouseOver As Boolean

Private m_HotLine As Boolean
Private m_HotLineWidth As Long
Private m_HotLinePosition As HotLinePos
Private m_HotLineColor As OLE_COLOR
Private m_HotLineActiveColor As OLE_COLOR

Private m_Font          As StdFont
Private m_IconFont      As StdFont
Private m_IconCharCode  As Long
Private m_IconForeColor As Long
Private m_PadY          As Long
Private m_PadX          As Long
Private m_MouseOver     As Boolean
Private m_CaptionAlignV As eTextAlignV
Private m_CaptionAlignH As eTextAlignH

Private cHKey         As String

Private m_EffectFade  As Boolean
Private m_InitialOpacity As Long
Private m_Transparent As Boolean



Private Function DefineHotKey(sString As String) As Boolean
  Dim I As Integer, Pos As Integer
  Pos = InStr(1, m_Caption, "&")
  If Pos <> 0 Then
    cHKey = Mid$(m_Caption, Pos + 1, 1)
    DefineHotKey = True
  Else
    cHKey = "-1"
    DefineHotKey = False
  End If
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

Public Sub Refresh()
  UserControl.Cls
  If m_Transparent Then CopyAmbient
  Draw
  AddIconChar
End Sub

Private Function AddIconChar()
Dim Rct As Rect
Dim pFont       As IFont
Dim lFontOld    As Long
    
On Error GoTo ErrF
With UserControl
  .AutoRedraw = True
  Set pFont = IconFont
  lFontOld = SelectObject(.hDC, pFont.hFont)
  
  Rct.Left = (IconFont.Size / 2) + m_PadX
  Rct.Top = (IconFont.Size / 2) + m_PadY
  Rct.Right = UserControl.ScaleWidth
  Rct.Bottom = UserControl.ScaleHeight

  If m_Clicked Then
    .ForeColor = m_ForeColor2
    If m_ButtonEffect = bPressEffect Then
       If m_PressEffect = pSlanted Then
         Rct.Left = (IconFont.Size / 2) + m_PadX + 3
         Rct.Top = (IconFont.Size / 2) + m_PadY + 3
       Else
         Rct.Top = (IconFont.Size / 2) + m_PadY + 3
       End If
    End If
  Else
    .ForeColor = m_IconForeColor
  End If
 
  DrawTextW .hDC, IconCharCode, 1, Rct, 0
  
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
Dim lBorder As Long, mBorder As Long
Dim mpath As Long
Dim TextOpacity As Long, TextColor As OLE_COLOR

With UserControl
  .BackStyle = 1
  .BackColor = m_BackGroundColor
  
  GdipCreateFromHDC .hDC, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

  lBorder = m_BorderWidth * 2
  mBorder = lBorder / 2
  
  stREC.Left = mBorder * nScale
  stREC.Top = mBorder * nScale
  stREC.Width = .ScaleWidth - lBorder * nScale
  stREC.Height = .ScaleHeight - lBorder * nScale
    
  If m_ButtonEffect = bPressEffect Then
    Dim REC2 As RECTL
      
    REC.Left = 0:     REC.Top = 0
    REC.Width = IIf(m_PressEffect = pSlanted, .ScaleWidth - 4, .ScaleWidth - 2) * nScale
    REC.Height = .ScaleHeight - 4 * nScale
    
    REC2.Left = IIf(m_PressEffect = pSlanted, 2, 0) * nScale
    REC2.Top = 2 * nScale
    REC2.Width = IIf(m_PressEffect = pSlanted, .ScaleWidth - 3, .ScaleWidth - 2) * nScale
    REC2.Height = .ScaleHeight - 3 * nScale
    
    stREC.Width = IIf(m_PressEffect = pSlanted, .ScaleWidth - lBorder - 3, .ScaleWidth - lBorder) * nScale
    stREC.Height = .ScaleHeight - lBorder - 2 * nScale

  Else
    REC.Left = 1 * nScale:    REC.Top = 1 * nScale
    REC.Width = .ScaleWidth - 2 * nScale
    REC.Height = .ScaleHeight - 2 * nScale
  End If
  
  SafeRange m_Opacity, 0, 100

  If m_Clickable Then
    If m_OptionButton Then
      If m_Value Then
        GoTo Clicked
      Else
        GoTo NoClicked
      End If
    Else
      If m_Clicked Then
        GoTo Clicked
      Else
        GoTo NoClicked
      End If
    End If
  End If

  '-DRAW BUTTON--------------
NoClicked:
  If m_EffectFade Then
    GoTo Effect
  Else
    mpath = gRoundRect(hGraphics, REC, argb(m_Color2, 100), argb(m_Color1, 100), m_Angulo, m_BorderWidth, argb(m_BorderColor, 100), m_CornerCurve, m_Filled)
    
    If m_ButtonEffect = bPressEffect Then
      gRoundRect hGraphics, REC2, argb(m_BorderColor, 60), argb(m_BorderColor, 60), m_Angulo, 1, argb(m_BorderColor, 60), m_CornerCurve, m_Filled
      mpath = gRoundRect(hGraphics, REC, argb(m_Color2, 100), argb(m_Color1, 100), m_Angulo, m_BorderWidth, argb(m_BorderColor, 100), m_CornerCurve, m_Filled)
      'DrawCaption hGraphics, m_Caption, stREC, m_ForeColor, 100, m_CaptionAngle, m_CaptionAlignH, m_CaptionAlignV, True
    Else
      mpath = gRoundRect(hGraphics, REC, argb(m_Color2, 100), argb(m_Color1, 100), m_Angulo, m_BorderWidth, IIf(m_MouseOver = True, argb(m_BorderColorOnFocus, 100), argb(m_BorderColor, m_Opacity)), m_CornerCurve, m_Filled)
      'DrawCaption hGraphics, m_Caption, stREC, m_ForeColor, 100, m_CaptionAngle, m_CaptionAlignH, m_CaptionAlignV, True
    End If
    TextColor = m_ForeColor
    TextOpacity = 100
  End If
  GoTo Continuar
  
Clicked:
  If m_ButtonEffect = bPressEffect Then
    stREC.Top = stREC.Top + 2
    If m_PressEffect = pSlanted Then stREC.Left = stREC.Left + 1
    If m_PressEffect = pSlanted Then stREC.Width = .ScaleWidth - lBorder - 1 * nScale
    stREC.Height = .ScaleHeight - lBorder - 1 * nScale
    mpath = gRoundRect(hGraphics, REC2, argb(m_Color2, 100), argb(m_Color1, 100), m_Angulo, m_BorderWidth, argb(m_BorderColorOnFocus, 100), m_CornerCurve, m_Filled)
    'DrawCaption hGraphics, m_Caption, stREC, m_ForeColor2, 100, m_CaptionAngle, m_CaptionAlignH, m_CaptionAlignV, True
  Else
    If m_FlatEffect = pInvertGradient Then
      mpath = gRoundRect(hGraphics, REC, argb(m_Color1, 100), argb(m_Color2, 100), m_Angulo, m_BorderWidth, argb(m_BorderColorOnFocus, 100), m_CornerCurve, m_Filled)
    Else
      mpath = gRoundRect(hGraphics, REC, argb(m_ColorActive1, 100), argb(m_ColorActive2, 100), m_Angulo, m_BorderWidth, argb(m_BorderColorOnFocus, 100), m_CornerCurve, m_Filled)
    End If
    'DrawCaption hGraphics, m_Caption, stREC, m_ForeColor2, 100, m_CaptionAngle, m_CaptionAlignH, m_CaptionAlignV, True
  End If
  TextColor = m_ForeColor2
  TextOpacity = 100
  GoTo Continuar
 
Effect:
  If m_ButtonEffect = bPressEffect Then
    gRoundRect hGraphics, REC2, argb(m_BorderColor, m_Opacity), argb(m_BorderColor, m_Opacity), m_Angulo, 1, argb(m_BorderColor, m_Opacity), m_CornerCurve, m_Filled
    mpath = gRoundRect(hGraphics, REC, argb(m_Color1, m_Opacity), argb(m_Color2, m_Opacity), m_Angulo, m_BorderWidth, argb(m_BorderColor, m_Opacity), m_CornerCurve, m_Filled)
    'DrawCaption hGraphics, m_Caption, stREC, m_ForeColor, m_Opacity, m_CaptionAngle, m_CaptionAlignH, m_CaptionAlignV, True
  Else
    mpath = gRoundRect(hGraphics, REC, argb(m_Color1, m_Opacity), argb(m_Color2, m_Opacity), m_Angulo, m_BorderWidth, argb(m_BorderColor, m_Opacity), m_CornerCurve, m_Filled)
    'DrawCaption hGraphics, m_Caption, stREC, m_ForeColor, m_Opacity, m_CaptionAngle, m_CaptionAlignH, m_CaptionAlignV, True
  End If
  TextColor = m_ForeColor
  TextOpacity = m_Opacity

Continuar:
  'Draw HotLine & Caption
  If m_HotLine Then DrawHotLine hGraphics, mpath, IIf(m_MouseOver = True, m_HotLineActiveColor, m_HotLineColor)
  DrawCaption hGraphics, m_Caption, stREC, TextColor, TextOpacity, m_CaptionAngle, m_CaptionAlignH, m_CaptionAlignV, True
' '---------------
' '---------------
  GdipDeletePath mpath
  GdipDeleteGraphics hGraphics
  '---------------
  If m_Transparent Then
    .BackStyle = 0
    .MaskColor = .BackColor
    Set .MaskPicture = .Image
  End If
  '---------------
End With

End Sub

Private Function DrawCaption(ByVal hGraphics As Long, sString As String, layoutRect As RECTS, TextColor As OLE_COLOR, _
                             ColorOpacity As Long, mAngle As Single, HAlign As eTextAlignH, VAlign As eTextAlignV, _
                             Optional bWordWrap As Boolean = True) As Long
    Dim hPath As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE

    If GdipCreatePath(&H0, hPath) = 0 Then

        If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
            If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
            GdipSetStringFormatHotkeyPrefix hFormat, HotkeyPrefixShow
            GdipSetStringFormatTrimming hFormat, StringTrimmingEllipsisWord
            GdipSetStringFormatAlign hFormat, HAlign
            GdipSetStringFormatLineAlign hFormat, VAlign
        End If

        GetFontStyleAndSize m_Font, lFontStyle, lFontSize

        If GdipCreateFontFamilyFromName(StrPtr(m_Font.Name), 0, hFontFamily) Then
            If hFontCollection Then
                If GdipCreateFontFamilyFromName(StrPtr(m_Font.Name), hFontCollection, hFontFamily) Then
                    If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
                End If
            Else
                If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
            End If
        End If
'------------------------------------------------------------------------
        If mAngle <> 0 Then
            Dim newY As Long, newX As Long
            newY = (layoutRect.Height / 2)
            newX = (layoutRect.Width / 2)
            Call GdipTranslateWorldTransform(hGraphics, newX, newY, 0)
            Call GdipRotateWorldTransform(hGraphics, mAngle, 0)
            Call GdipTranslateWorldTransform(hGraphics, -newX, -newY, 0)
        End If
'------------------------------------------------------------------------
         layoutRect.Left = layoutRect.Left + m_CaptionX
         layoutRect.Top = layoutRect.Top + m_CaptionY
'------------------------------------------------------------------------
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

Private Function DrawHotLine(hGraphics As Long, hPath As Long, hlColor As OLE_COLOR) ', ByVal PosX As Long, ByVal PosY As Long)
    Dim hBrush As Long  ', hPath As Long
    Dim X As Long, Y As Long
    Dim WW As Long, HH As Long
    Dim bW As Long
    Dim lW As Long
    Dim CL As Long
    Dim SS As Long
        
    bW = m_BorderWidth * nScale
    
    X = bW
    Y = bW
    WW = (UserControl.ScaleWidth - bW * 2) * nScale + 1
    HH = (UserControl.ScaleHeight - bW * 2) * nScale + 1
    lW = m_HotLineWidth * nScale
    
'    GdipCreatePath &H0, hPath
    
    Select Case m_HotLinePosition
        Case hlLeft
            X = X + lW
        Case hlTop
            Y = Y + lW
        Case hlRight
            WW = WW - lW
        Case hlBottom
            HH = HH - lW
    End Select
            
    GdipSetClipRectI hGraphics, X, Y, WW, HH, CombineModeExclude
    
    GdipCreateSolidFill argb(hlColor, m_Opacity), hBrush
    
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipResetClip hGraphics
        
End Function

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

Public Function GetWindowsDPI() As Double
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

Private Function gRoundRect(ByVal hGraphics As Long, Rect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal Angulo As Single, ByVal BorderW As Long, ByVal BorderColor As Long, ByVal Round As Long, Filled As Boolean) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mpath As Long
    Dim mRound As Long
    
    If m_BorderWidth <> 0 Then GdipCreatePen1 BorderColor, BorderW * nScale, &H2, hPen   '&H1 * nScale, &H2, hPen
    If Filled Then GdipCreateLineBrushFromRectWithAngleI Rect, color1, color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mpath   '&H0
    
    With Rect
        mRound = GetSafeRound((Round * nScale), .Width * 2, .Height * 2)
        If mRound = 0 Then mRound = 1
            GdipAddPathArcI mpath, .Left, .Top, mRound, mRound, 180, 90
            GdipAddPathArcI mpath, (.Left + .Width) - mRound, .Top, mRound, mRound, 270, 90
            GdipAddPathArcI mpath, (.Left + .Width) - mRound, (.Top + .Height) - mRound, mRound, mRound, 0, 90
            GdipAddPathArcI mpath, .Left, (.Top + .Height) - mRound, mRound, mRound, 90, 90
    End With
    
    GdipClosePathFigures mpath
    GdipFillPath hGraphics, hBrush, mpath
    GdipDrawPath hGraphics, hPen, mpath
    
    gRoundRect = mpath
    'Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)

End Function

'Inicia GDI+
Private Sub InitGDI()
    Dim gdipStartupInput As GdiplusStartupInput
    gdipStartupInput.GdiplusVersion = 1&
    Call GdiplusStartup(gdipToken, gdipStartupInput, ByVal 0)
End Sub

Public Function IsMouseOver() As Boolean
    Dim PT As POINTL
    GetCursorPos PT
    IsMouseOver = (WindowFromPoint(PT.X, PT.Y) = UserControl.hWnd)
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


Private Sub OptBehavior()
Dim Frm As Form
    Set Frm = Extender.Parent
    
Dim lHwnd As Long
    lHwnd = UserControl.ContainerHwnd

    Dim Ctrl As Control
    For Each Ctrl In Frm.Controls
        With Ctrl
           If TypeOf Ctrl Is AxGButtonLabel Then
              If .OptionButton = True Then
                 If (.Container.hWnd = lHwnd) And ObjPtr(Ctrl) <> ObjPtr(Extender) Then
                  If .Value Then .Value = False
                 End If
              End If
           End If
        End With
    Next
End Sub

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

Private Sub SafeRange(ByVal Value As Long, ByVal Min As Long, ByVal Max As Long)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub

'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(gdipToken)
End Sub

Private Sub tmrEffect_Timer()
If IsMouseOver Then

  m_MouseOver = True
  If m_ChangeBorderMouseOver Then Refresh
  
  If m_EffectFade Then
    If m_Opacity < 100 Then
      m_Opacity = m_Opacity + 2
      Refresh
    Else
      m_MouseOver = False
      Exit Sub
    End If
  End If
  
Else

  m_MouseOver = False
  m_Opacity = m_InitialOpacity
  Refresh
  tmrEffect.Enabled = False
  
End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  Refresh
End Sub

Private Sub UserControl_Click()
m_MouseOver = True

  If m_OptionButton = True Then
    Value = True
    OptBehavior
    Refresh
  End If
  
  RaiseEvent Click

End Sub

Private Sub UserControl_EnterFocus()
Debug.Print "EnterFocus"
End Sub

Private Sub UserControl_Initialize()
'---GDI+--------------
    InitGDI
    nScale = GetWindowsDPI
'---------------------
'Set cSC = New c_SubClass
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
  
  hFontCollection = ReadValue(&HFC)

  m_Clickable = True
  m_Clicked = False
  
  'm_BackGroundColor = Extender.Container.BackColor
  Set m_Font = UserControl.Ambient.Font
  m_CaptionAlignV = eMiddle
  m_CaptionAlignH = eCenter
  m_Caption = Ambient.DisplayName
  m_CaptionX = 0
  m_CaptionY = 0

  m_Filled = True
  m_Transparent = True
  m_ButtonEffect = bPressEffect
  m_PressEffect = pDownStraight
  m_FlatEffect = pInvertGradient
  m_BorderColor = &HDDB784
  m_OldBorderColor = &HDDB784
  m_BorderColorOnFocus = &HDDB784
  m_ChangeBorderMouseOver = True
  m_ForeColor = &HFFFFFF
  m_ForeColor2 = &HFFFFFF
  m_Enabled = True
  m_Color1 = &H8C562F
  m_Color2 = &H714626
  m_ColorActive1 = &H714626
  m_ColorActive2 = &H714626
  m_Angulo = 180
  m_BorderWidth = 0
  m_CornerCurve = 10
  m_Transparent = False
  m_Opacity = 50
  m_InitialOpacity = m_Opacity
  m_IconCharCode = "&H0"
  m_HotLine = False
  m_HotLineWidth = 10
  m_HotLinePosition = 0
  m_HotLineColor = vbRed
  m_HotLineActiveColor = vbBlue
  
  DefineHotKey m_Caption
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
m_Clicked = True
m_MouseOver = True
Refresh
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_Clickable Then MousePointerHands True
tmrEffect.Enabled = True
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
m_Clicked = False
m_MouseOver = True
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
  
  m_BackGroundColor = .ReadProperty("BackGroundColor", UserControl.Ambient.BackColor)
  m_Color1 = .ReadProperty("BackColor1", &H8C562F)
  m_Color2 = .ReadProperty("BackColor2", &H714626)
  m_ColorActive1 = .ReadProperty("BackColorActive1", &H714626)
  m_ColorActive2 = .ReadProperty("BackColorActive2", &H714626)
  m_ForeColor = .ReadProperty("ForeColor", vbWhite)
  m_ForeColor2 = .ReadProperty("ForeColor2", vbWhite)
  m_Angulo = .ReadProperty("BackAngle", 180)
  m_BorderColor = .ReadProperty("BorderColor", &HDDB784)
  m_BorderWidth = .ReadProperty("BorderWidth", 0)
  m_CornerCurve = .ReadProperty("CornerCurve", 15)
  
  m_Filled = .ReadProperty("Filled", True)
  m_Transparent = .ReadProperty("Transparent", True)
  m_BorderColorOnFocus = .ReadProperty("ColorOnFocus", &HDDB784)
  m_ChangeBorderOnFocus = .ReadProperty("ChangeColorOnFocus", False)
  m_ChangeBorderMouseOver = .ReadProperty("ChangeBorderMouseOver", True)
  m_EffectFade = .ReadProperty("EffectFading", False)
  m_InitialOpacity = .ReadProperty("InitialOpacity", 50)
  
  m_ButtonEffect = .ReadProperty("ButtonEffect", bPressEffect)
  m_PressEffect = .ReadProperty("ButtonPressEffect", pDownStraight)
  m_FlatEffect = .ReadProperty("ButtonFlatEffect", pInvertGradient)
  
  Set m_Font = .ReadProperty("Font", UserControl.Ambient.Font)
  
  m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
  m_CaptionAlignV = .ReadProperty("CaptionAlignV", 1)
  m_CaptionAlignH = .ReadProperty("CaptionAlignH", 1)
  m_CaptionAngle = .ReadProperty("CaptionAngle", 0)
  m_CaptionX = .ReadProperty("CaptionX", 0)
  m_CaptionY = .ReadProperty("CaptionY", 0)
        
  Set m_IconFont = .ReadProperty("IconFont", UserControl.Ambient.Font)
  m_IconCharCode = .ReadProperty("IconCharCode", "&H0")
  m_IconForeColor = .ReadProperty("IconForeColor", &H404040)
  
  m_PadX = .ReadProperty("IcoPaddingX", 0)
  m_PadY = .ReadProperty("IcoPaddingY", 0)

  m_Value = .ReadProperty("Value", False)
  m_OptionButton = .ReadProperty("OptionButton", False)
  m_Clickable = .ReadProperty("Clickable", True)
  
  m_HotLine = .ReadProperty("HotLine", False)
  m_HotLineWidth = .ReadProperty("HotLineWidth", 10)
  m_HotLinePosition = .ReadProperty("HotLinePosition", 0)
  m_HotLineColor = .ReadProperty("HotLineColor", vbRed)
  m_HotLineActiveColor = .ReadProperty("HotLineActiveColor", vbBlue)
End With
  
  m_Opacity = m_InitialOpacity
  DefineHotKey m_Caption
  
'  With cSC
'      ucchWnd = UserControl.ContainerHwnd
'      If .Subclass(ucchWnd, , , Me) Then
'          .AddMsg ucchWnd, WM_SYSKEYDOWN, MSG_AFTER
'          .AddMsg ucchWnd, WM_KEYDOWN, MSG_AFTER
'      End If
'  End With
  
End Sub

Private Sub UserControl_Resize()
Refresh
End Sub

Private Sub UserControl_Show()
Refresh
End Sub

Private Sub UserControl_Terminate()
TerminateGDI
'Set cSC = Nothing

End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("Enabled", m_Enabled)
  Call .WriteProperty("BackGroundColor", m_BackGroundColor)
  Call .WriteProperty("BackColor1", m_Color1)
  Call .WriteProperty("BackColor2", m_Color2)
  Call .WriteProperty("BackColorActive1", m_ColorActive1)
  Call .WriteProperty("BackColorActive2", m_ColorActive2)
  Call .WriteProperty("ForeColor", m_ForeColor)
  Call .WriteProperty("ForeColor2", m_ForeColor2)
  Call .WriteProperty("BackAngle", m_Angulo)
  Call .WriteProperty("BorderColor", m_BorderColor)
  Call .WriteProperty("BorderWidth", m_BorderWidth)
  Call .WriteProperty("CornerCurve", m_CornerCurve)
  
  Call .WriteProperty("Filled", m_Filled)
  Call .WriteProperty("ButtonEffect", m_ButtonEffect)
  Call .WriteProperty("ButtonPressEffect", m_PressEffect)
  Call .WriteProperty("ButtonFlatEffect", m_FlatEffect)
  
  Call .WriteProperty("Font", m_Font)
  Call .WriteProperty("CaptionAngle", m_CaptionAngle)
  Call .WriteProperty("Caption", m_Caption)
  Call .WriteProperty("CaptionAlignV", m_CaptionAlignV)
  Call .WriteProperty("CaptionAlignH", m_CaptionAlignH)
  Call .WriteProperty("CaptionX", m_CaptionX)
  Call .WriteProperty("CaptionY", m_CaptionY)
  
  Call .WriteProperty("Transparent", m_Transparent)
    
  Call .WriteProperty("ColorOnFocus", m_BorderColorOnFocus)
  Call .WriteProperty("ChangeColorOnFocus", m_ChangeBorderOnFocus)
  Call .WriteProperty("ChangeBorderMouseOver", m_ChangeBorderMouseOver)
  Call .WriteProperty("EffectFading", m_EffectFade)
  Call .WriteProperty("InitialOpacity", m_InitialOpacity)
  
  Call .WriteProperty("IconFont", m_IconFont)
  Call .WriteProperty("IconCharCode", m_IconCharCode, 0)
  Call .WriteProperty("IconForeColor", m_IconForeColor)
  
  Call .WriteProperty("IcoPaddingX", m_PadX)
  Call .WriteProperty("IcoPaddingY", m_PadY)

  Call .WriteProperty("Value", m_Value)
  Call .WriteProperty("OptionButton", m_OptionButton)
  Call .WriteProperty("Clickable", m_Clickable)
  
  Call .WriteProperty("HotLine", m_HotLine, False)
  Call .WriteProperty("HotLineWidth", m_HotLineWidth, 10)
  Call .WriteProperty("HotLinePosition", m_HotLinePosition, 0)
  Call .WriteProperty("HotLineColor", m_HotLineColor)
  Call .WriteProperty("HotLineActiveColor", m_HotLineActiveColor)
End With
  
End Sub

Public Property Get BackAngle() As Single
  BackAngle = m_Angulo
End Property

Public Property Let BackAngle(ByVal New_Angulo As Single)
  m_Angulo = New_Angulo
  PropertyChanged "BackAngle"
  Refresh
End Property

Public Property Get BackGroundColor() As OLE_COLOR
  BackGroundColor = m_BackGroundColor
End Property

Public Property Let BackGroundColor(ByVal NewColor As OLE_COLOR)
  m_BackGroundColor = NewColor
  PropertyChanged "BackGroundColor"
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

Public Property Get BackColorActive1() As OLE_COLOR
  BackColorActive1 = m_ColorActive1
End Property

Public Property Let BackColorActive1(ByVal New_Color1 As OLE_COLOR)
  m_ColorActive1 = New_Color1
  PropertyChanged "BackColorActive1"
  Refresh
End Property

Public Property Get BackColorActive2() As OLE_COLOR
  BackColorActive2 = m_ColorActive2
End Property

Public Property Let BackColorActive2(ByVal New_Color2 As OLE_COLOR)
  m_ColorActive2 = New_Color2
  PropertyChanged "BackColorActive2"
  Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
  m_BorderColor = NewBorderColor
  PropertyChanged "BorderColor"
  Refresh
End Property

Public Property Get BorderColorOnFocus() As OLE_COLOR
  BorderColorOnFocus = m_BorderColorOnFocus
End Property

Public Property Let BorderColorOnFocus(ByVal NewBorderColorOnFocus As OLE_COLOR)
  m_BorderColorOnFocus = NewBorderColorOnFocus
  PropertyChanged "BorderColorOnFocus"
  Refresh
End Property

Public Property Get HotLineColor() As OLE_COLOR
  HotLineColor = m_HotLineColor
End Property

Public Property Let HotLineColor(ByVal NewHotLineColor As OLE_COLOR)
  m_HotLineColor = NewHotLineColor
  PropertyChanged "HotLineColor"
  Refresh
End Property

Public Property Get BorderWidth() As Long
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  Refresh
End Property

Public Property Get ButtonEffect() As eButtonEffect
ButtonEffect = m_ButtonEffect
End Property

Public Property Let ButtonEffect(ByVal nEffect As eButtonEffect)
  m_ButtonEffect = nEffect
  PropertyChanged "ButtonEffect"
  Refresh
End Property

Public Property Get CaptionAlignH() As eTextAlignH
  CaptionAlignH = m_CaptionAlignH
End Property

Public Property Let CaptionAlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_CaptionAlignH = NewCaptionAlignH
  PropertyChanged "CaptionAlignH"
  Refresh
End Property

Public Property Get CaptionAlignV() As eTextAlignV
  CaptionAlignV = m_CaptionAlignV
End Property

Public Property Let CaptionAlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_CaptionAlignV = NewCaptionAlignV
  PropertyChanged "CaptionAlignV"
  Refresh
End Property

Public Property Get CaptionAngle() As Single
    CaptionAngle = m_CaptionAngle
End Property

Public Property Let CaptionAngle(ByVal New_Angle As Single)
    m_CaptionAngle = New_Angle
    PropertyChanged "CaptionAngle"
    Refresh
End Property

Public Property Get Caption() As String
  Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
  m_Caption = NewCaption
  PropertyChanged "Caption"
  Refresh
  'Define HotKey
  DefineHotKey m_Caption
End Property

Public Property Get CaptionX() As Long
  CaptionX = m_CaptionX
End Property

Public Property Let CaptionX(ByVal NewPosX As Long)
  m_CaptionX = NewPosX
  PropertyChanged "CaptionX"
  Refresh
End Property

Public Property Get CaptionY() As Long
  CaptionY = m_CaptionY
End Property

Public Property Let CaptionY(ByVal NewPosY As Long)
  m_CaptionY = NewPosY
  PropertyChanged "CaptionY"
  Refresh
End Property

Public Property Get ChangeBorderOnFocus() As Boolean
  ChangeBorderOnFocus = m_ChangeBorderOnFocus
End Property

Public Property Let ChangeBorderOnFocus(ByVal NewChangeBorderOnFocus As Boolean)
  m_ChangeBorderOnFocus = NewChangeBorderOnFocus
  m_OldBorderColor = m_BorderColor
  PropertyChanged "ChangeBorderOnFocus"
End Property

Public Property Get ChangeBorderMouseOver() As Boolean
  ChangeBorderMouseOver = m_ChangeBorderMouseOver
End Property

Public Property Let ChangeBorderMouseOver(ByVal NewChangeBorderMouseOver As Boolean)
  m_ChangeBorderMouseOver = NewChangeBorderMouseOver
  m_OldBorderColor = m_BorderColor
  PropertyChanged "ChangeBorderMouseOver"
End Property

Public Property Get Clickable() As Boolean
   Clickable = m_Clickable
End Property

Public Property Let Clickable(ByVal bClickable As Boolean)
   m_Clickable = bClickable
   PropertyChanged "Clickable"
End Property

Public Property Get CornerCurve() As Long
  CornerCurve = m_CornerCurve
End Property

Public Property Let CornerCurve(ByVal NewCornerCurve As Long)
  m_CornerCurve = NewCornerCurve
  PropertyChanged "CornerCurve"
  Refresh
End Property

Public Property Get EffectFading() As Boolean
EffectFading = m_EffectFade
End Property

Public Property Let EffectFading(ByVal vNewValue As Boolean)
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
  'Refresh
End Property

Public Property Get Filled() As Boolean
  Filled = m_Filled
End Property

Public Property Let Filled(ByVal New_Filled As Boolean)
  m_Filled = New_Filled
  PropertyChanged "Filled"
  Refresh
End Property

Public Property Get Font() As StdFont
  Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
  Set m_Font = New_Font
  PropertyChanged "Font"
  Refresh
End Property

Public Property Get ForeColor2() As OLE_COLOR
  ForeColor2 = m_ForeColor2
End Property

Public Property Let ForeColor2(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor2 = NewForeColor
  PropertyChanged "ForeColor2"
  Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor = NewForeColor
  PropertyChanged "ForeColor"
  Refresh
End Property

Public Property Get ButtonFlatEffect() As eFlatEffect
  ButtonFlatEffect = m_FlatEffect
End Property

Public Property Let ButtonFlatEffect(ByVal nPress As eFlatEffect)
  m_FlatEffect = nPress
  PropertyChanged "ButtonFlatEffect"
  Refresh
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get IconCharCode() As String
    IconCharCode = "&H" & Hex(m_IconCharCode)
End Property

Public Property Let IconCharCode(ByVal New_IconCharCode As String)
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not VBA.Left$(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        m_IconCharCode = "&H" & New_IconCharCode
    Else
        m_IconCharCode = New_IconCharCode
    End If
    PropertyChanged "IconCharCode"
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

Public Property Get IconForeColor() As OLE_COLOR
    IconForeColor = m_IconForeColor
End Property

Public Property Let IconForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor = New_ForeColor
    PropertyChanged "IconForeColor"
    Refresh
End Property

Public Property Get IcoPaddingX() As Long
IcoPaddingX = m_PadX
End Property

Public Property Let IcoPaddingX(ByVal XpadVal As Long)
m_PadX = XpadVal
PropertyChanged "IcoPaddingX"
Refresh
End Property

Public Property Get IcoPaddingY() As Long
IcoPaddingY = m_PadY
End Property

Public Property Let IcoPaddingY(ByVal YpadVal As Long)
m_PadY = YpadVal
PropertyChanged "IcoPaddingY"
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

Public Property Get OptionButton() As Boolean
   OptionButton = m_OptionButton
End Property

Public Property Let OptionButton(ByVal bOptionButton As Boolean)
   m_OptionButton = bOptionButton
   PropertyChanged "OptionButton"
End Property

Public Property Get ButtonPressEffect() As ePressEffect
  ButtonPressEffect = m_PressEffect
End Property

Public Property Let ButtonPressEffect(ByVal nPress As ePressEffect)
  m_PressEffect = nPress
  PropertyChanged "ButtonPressEffect"
  Refresh
End Property

Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal NewValue As Boolean)
    m_Transparent = NewValue
    PropertyChanged "Transparent"
    UserControl.BackStyle = 1
    Refresh
End Property

Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Boolean)
  If m_OptionButton And NewValue Then OptBehavior
    m_Value = NewValue
    m_Clicked = NewValue
    Refresh
    PropertyChanged "Value"
    RaiseEvent ChangeValue(m_Value)
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

Public Property Get HotLine() As Boolean
  HotLine = m_HotLine
End Property

Public Property Let HotLine(ByVal NewHotLine As Boolean)
  m_HotLine = NewHotLine
  PropertyChanged "HotLine"
  Refresh
End Property

Public Property Get HotLineWidth() As Long
  HotLineWidth = m_HotLineWidth
End Property

Public Property Let HotLineWidth(ByVal NewHotLineWidth As Long)
  m_HotLineWidth = NewHotLineWidth
  PropertyChanged "HotLineWidth"
  Refresh
End Property

Public Property Get HotLinePosition() As HotLinePos
  HotLinePosition = m_HotLinePosition
End Property

Public Property Let HotLinePosition(ByVal NewHotLinePosition As HotLinePos)
  m_HotLinePosition = NewHotLinePosition
  PropertyChanged "HotLinePosition"
  Refresh
End Property

Public Property Get HotLineActiveColor() As OLE_COLOR
  HotLineActiveColor = m_HotLineActiveColor
End Property

Public Property Let HotLineActiveColor(ByVal NewHotLineActiveColor As OLE_COLOR)
  m_HotLineActiveColor = NewHotLineActiveColor
  PropertyChanged "HotLineActiveColor"
  Refresh
End Property


