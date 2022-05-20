VERSION 5.00
Begin VB.UserControl AxGOption 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
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
   ToolboxBitmap   =   "AxGOption2.ctx":0000
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   255
      Top             =   240
   End
End
Attribute VB_Name = "AxGOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxGOptionButton
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

Private Enum PenAlignment
    PenAlignmentCenter = &H0
    PenAlignmentInset = &H1
End Enum

Public Enum OpStyle
    stOption = 0
    stSwitch = 1
End Enum

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
Dim m_ForeColor     As OLE_COLOR
Dim m_ForeColor2    As OLE_COLOR
Dim m_Color1        As OLE_COLOR
Dim m_Color2        As OLE_COLOR
Dim m_ActiveColor   As OLE_COLOR
Dim m_CheckColor    As OLE_COLOR
Dim m_BorderColorFocus As OLE_COLOR

'Dim m_FillColor     As OLE_COLOR
'Dim m_FillEnable    As Boolean

Dim m_BorderWidth   As Long
Dim m_Enabled       As Boolean
Dim m_Angulo        As Single
Dim m_CornerCurve   As Long
Dim YCheckPos       As Long
Dim XCheckPos       As Long
Dim m_BoxPosition As CheckPos
Dim m_CheckVisible  As Boolean
Dim m_Moveable      As Boolean
Dim m_CaptionAngle  As Single
Dim m_CaptionX      As Long
Dim m_CaptionY      As Long
Dim m_Caption       As String
Dim m_CaptionEnabled  As Boolean
Dim m_Line1       As Boolean
Dim m_FullWidth   As Integer
Dim m_FullHeight  As Integer
Dim m_Top         As Long
Dim m_Left        As Long
Dim OldX          As Single
Dim OldY          As Single
Dim m_Opacity     As Long
Dim cl_hWnd       As Long

Dim m_ChangeBorderOnFocus As Boolean
Dim m_OnFocus As Boolean

Private m_Font          As StdFont
Private m_IconFont      As StdFont
Private m_IconCharCodeOn   As Long
Private m_IconForeColorOn  As Long
Private m_IconCharCodeOff  As Long
Private m_IconForeColorOff As Long
Private m_PadY          As Long
Private m_PadX          As Long
Private m_MouseOver     As Boolean
Private m_CaptionAlignV As eTextAlignV
Private m_CaptionAlignH As eTextAlignH
'Private m_CaptionShadow As Boolean

Private m_StringPosX  As Long
Private m_StringPosY  As Long
Private m_EffectFade  As Boolean
Private m_InitialOpacity As Long
Private m_Transparent As Boolean

Private m_Value           As Boolean
Private m_OptionBehavior  As Boolean
Private m_Clicked         As Boolean
Private m_CheckStyle      As StyleDraw
Private m_Style           As OpStyle

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
  AddIconChar
  UserControl.Refresh
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

  If m_MouseOver Then
    .ForeColor = m_ForeColor
  Else
    .ForeColor = IIf(m_Value = True, m_IconForeColorOn, m_IconForeColorOff)
  End If
  
  Rct.Left = (IconFont.Size / 2) + m_PadX
  Rct.Top = (IconFont.Size / 2) + m_PadY
  Rct.Right = UserControl.ScaleWidth
  Rct.Bottom = UserControl.ScaleHeight
  
  DrawTextW .hDC, IIf(m_Value = True, IconCharCodeOn, IconCharCodeOff), 1, Rct, 0
  
  Call SelectObject(.hDC, lFontOld)
  
ErrF:
  Set pFont = Nothing
End With
End Function


Private Function AddString(ByVal hGraphics As Long, sString As String, layoutRect As RECTS, ForeColor As OLE_COLOR, ColorOpacity As Integer, HAlign As eTextAlignH, VAlign As eTextAlignV, bWordWrap As Boolean) As Long
'x As Long, y As Long, Width As Long, Height As Long ---> Replaced by layoutRect
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
      
    If GdipCreateFontFamilyFromName(StrPtr(m_Font.Name), 0, hFontFamily) Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        GdipSetStringFormatAlign hFormat, HAlign
        GdipSetStringFormatLineAlign hFormat, VAlign
    End If
        
    If m_Font.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If m_Font.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If m_Font.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If m_Font.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        
    lFontSize = MulDiv(m_Font.Size, GetDeviceCaps(hDC, LOGPIXELSY), 72)

    GdipCreateSolidFill argb(ForeColor, ColorOpacity), hBrush
            
    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
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
Dim REC2 As RECTL
Dim TmpR As RECTL
Dim stREC As RECTS
Dim lBorder As Long, mBorder As Long
Dim m_PosX1 As Integer, m_PosY1 As Integer
Dim m_PosX2 As Integer, m_PosY2 As Integer
Dim BColor As OLE_COLOR

With UserControl
    
  GdipCreateFromHDC .hDC, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

  lBorder = (m_BorderWidth * 2) * nScale
  mBorder = lBorder / 2
    
  Select Case m_Style
      Case Is = stOption
          TmpR.Left = mBorder * nScale
          TmpR.Top = mBorder * nScale
          TmpR.Width = .ScaleWidth - lBorder * nScale
          TmpR.Height = .ScaleHeight - lBorder * nScale
          
          stREC.Width = IIf(.ScaleHeight > .ScaleWidth, .ScaleHeight - ScaleWidth, .ScaleWidth - ScaleHeight)
          stREC.Height = IIf(.ScaleHeight > .ScaleWidth, ScaleWidth, .ScaleHeight)
      
      Case Is = stSwitch
          TmpR.Left = mBorder * nScale
          TmpR.Top = mBorder * nScale
          TmpR.Height = .ScaleHeight - lBorder * nScale
          TmpR.Width = (.ScaleHeight * 2) - (lBorder * 2) * nScale
          
          stREC.Width = IIf(.ScaleHeight > .ScaleWidth, .ScaleHeight - ScaleWidth, .ScaleWidth - TmpR.Width)
          stREC.Height = IIf(.ScaleHeight > .ScaleWidth, ScaleWidth, .ScaleHeight)
  End Select
    
  '---------------
  'If m_FillEnable Then
  '  .BackColor = m_FillColor
  'End If
  '---------------

  'Box Rect --------------------------------------------------
  Select Case m_BoxPosition
    Case cLeft
        REC.Left = TmpR.Left:      REC.Top = TmpR.Top
        REC.Width = IIf(m_Style = stOption, TmpR.Height, TmpR.Width): REC.Height = TmpR.Height
  'String Rect --------------------------------------------------
        stREC.Left = REC.Width + lBorder:       stREC.Top = 2
        
    Case cRight
        REC.Left = .ScaleWidth - TmpR.Height - mBorder: REC.Top = TmpR.Top
        REC.Width = IIf(m_Style = stOption, TmpR.Height, TmpR.Width): REC.Height = TmpR.Height
  'String Rect --------------------------------------------------
        stREC.Left = 2:            stREC.Top = 2
    
    Case cTop
        REC.Left = TmpR.Left:      REC.Top = TmpR.Top
        REC.Width = TmpR.Width:    REC.Height = IIf(m_Style = stOption, TmpR.Width, TmpR.Width / 2)
  'String Rect --------------------------------------------------
        stREC.Left = .ScaleWidth - (stREC.Width / 2) - (mBorder * 3): stREC.Top = (.ScaleHeight / 2)
        
    Case cBottom
        REC.Left = TmpR.Left:      REC.Top = .ScaleHeight - TmpR.Width - mBorder
        REC.Width = TmpR.Width:    REC.Height = IIf(m_Style = stOption, TmpR.Width, TmpR.Width / 2)
  'String Rect --------------------------------------------------
        stREC.Left = .ScaleWidth - (stREC.Width / 2) - (mBorder * 3):    stREC.Top = (.ScaleHeight / 2) - (REC.Height + mBorder)
  
  End Select
      
  Select Case m_Style
      Case Is = stOption
          REC2.Left = REC.Left + (mBorder)
          REC2.Top = REC.Top + (mBorder)
          REC2.Width = REC.Width - mBorder * 2
          REC2.Height = REC.Height - mBorder * 2
          
      Case Is = stSwitch
          REC2.Top = REC.Top + (mBorder)
          REC2.Width = REC.Height - mBorder * 2
          REC2.Height = REC.Height - mBorder * 2
          REC2.Left = IIf(m_Value = False, (REC.Left + mBorder), (REC.Width - mBorder * 2) - (REC.Height - lBorder * 2))
  End Select

  SafeRange m_Opacity, 0, 100

  If m_EffectFade Then
        GoTo Effect
  Else
        GoTo NoEffect
  End If

    
NoEffect:
    If m_Value Then
      DrawRoundRect hGraphics, REC, argb(m_Color1, 100), argb(m_Color2, 100), m_Angulo, m_BorderWidth, argb(m_BorderColor, 100), m_CornerCurve
      DrawRoundRect hGraphics, REC2, argb(m_ActiveColor, 100), argb(m_ActiveColor, 100), m_Angulo, mBorder, argb(m_BorderColor, 50), m_CornerCurve
    Else
      DrawRoundRect hGraphics, REC, argb(m_Color1, 100), argb(m_Color2, 100), m_Angulo, m_BorderWidth, argb(m_BorderColor, 100), m_CornerCurve
      If m_Style = stSwitch Then DrawRoundRect hGraphics, REC2, argb(m_BorderColor, 100), argb(m_BorderColor, 100), m_Angulo, mBorder, argb(m_BorderColor, 50), m_CornerCurve
    End If
    GoTo Continuar
      
Effect:
  DrawRoundRect hGraphics, REC, argb(m_Color1, m_Opacity), argb(m_Color2, m_Opacity), m_Angulo, m_BorderWidth, argb(m_BorderColor, m_Opacity), m_CornerCurve
  If m_Style = stSwitch Then DrawRoundRect hGraphics, REC2, argb(m_ActiveColor, m_Opacity), argb(m_ActiveColor, m_Opacity), m_Angulo, mBorder, argb(m_BorderColor, 50), m_CornerCurve

Continuar:
  '-DRAW CAPTION--------------
  Dim lPan As Single
  lPan = IIf(m_BoxPosition = cTop Or m_BoxPosition = cLeft, CSng(REC.Width + mBorder), 0)
  
  If m_CaptionEnabled Then
      DrawCaption hGraphics, m_Caption, stREC, m_ForeColor, 100, m_CaptionAngle, m_CaptionAlignH, m_CaptionAlignV, lPan, True
  End If
 '---------------
  If m_EffectFade Then
        GdipCreateSolidFill argb(m_CheckColor, m_Opacity), hBrush
        GdipCreatePen1 argb(m_CheckColor, m_Opacity), 1, UnitPixel, hPen
  Else
        GdipCreateSolidFill argb(m_CheckColor, 100), hBrush
        GdipCreatePen1 argb(m_CheckColor, 100), 1, UnitPixel, hPen
  End If
  '---------------
  
'  Select Case m_CheckStyle
'    Case Is = stDrawing
'    Case Is = stIconFont
'  End Select
  
  If m_Value Then
    If m_Style = stOption Then
      If m_CheckVisible Then
        Dim CheckWidth As Long
        Select Case m_BoxPosition
          Case cLeft, cRight
                CheckWidth = REC.Width
          Case cTop, cBottom
                CheckWidth = REC.Width
        End Select
        
        Select Case m_BoxPosition
          Case cTop
                YCheckPos = mBorder
                XCheckPos = mBorder
          Case cBottom
                YCheckPos = UserControl.ScaleHeight - (mBorder + CheckWidth)
                XCheckPos = mBorder
          Case cLeft
                YCheckPos = mBorder
                XCheckPos = mBorder
          Case cRight
                YCheckPos = UserControl.ScaleHeight - (mBorder + CheckWidth)
                XCheckPos = UserControl.ScaleWidth - (mBorder + CheckWidth)
        End Select
          DrawShape hGraphics, 0, XCheckPos, YCheckPos, CheckWidth, CheckWidth
      Else
        DrawRoundRect hGraphics, REC, argb(m_Color1, 50), argb(m_Color2, 50), m_Angulo, m_BorderWidth, argb(m_BorderColor, m_Opacity), m_CornerCurve
        DrawRoundRect hGraphics, REC2, argb(m_ActiveColor, m_Opacity), argb(m_ActiveColor, m_Opacity), m_Angulo, (m_BorderWidth / 2), argb(vbWhite, m_Opacity), m_CornerCurve
      End If
    End If
  End If
    
  Call GdipDeleteBrush(hBrush)
  Call GdipDeletePen(hPen)
  '---------------
  GdipDeleteGraphics hGraphics
  '---------------
  If m_Transparent Then
    .BackStyle = 0
    .MaskColor = .BackColor
    Set .MaskPicture = .Image
  Else
    .BackStyle = 1
    UserControl.Refresh
  End If
  '---------------
End With

End Sub

Private Function DrawCaption(ByVal hGraphics As Long, sString As String, layoutRect As RECTS, TextColor As OLE_COLOR, _
                             ColorOpacity As Integer, mAngle As Single, HAlign As eTextAlignH, VAlign As eTextAlignV, _
                             lPan As Single, Optional bWordWrap As Boolean = True) As Long
    Dim hPath As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim hDC As Long

    If GdipCreatePath(&H0, hPath) = 0 Then

        If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
            If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
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
            Select Case m_BoxPosition
              Case cTop
                  newY = (UserControl.ScaleHeight / 2) + (layoutRect.Height / 2)
                  newX = (UserControl.ScaleWidth / 2)
              Case cBottom
                  newY = (UserControl.ScaleHeight / 2) - (lPan - 1) - (layoutRect.Height / 2)
                  newX = (UserControl.ScaleWidth / 2) + 4
              Case cLeft
                  newY = (layoutRect.Height / 2)
                  newX = (layoutRect.Width / 2) + lPan
              Case cRight
                  newY = (layoutRect.Height / 2)
                  newX = (layoutRect.Width / 2)
            End Select
            
            Call GdipTranslateWorldTransform(hGraphics, newX, newY, 0)  '(layoutRect.Width / 2), (layoutRect.Height / 2), 0)
            Call GdipRotateWorldTransform(hGraphics, mAngle, 0)
            Call GdipTranslateWorldTransform(hGraphics, -newX, -newY, 0)
        End If
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

Private Sub DrawShape(ByVal iGraphics As Long, mShape As Integer, X As Long, Y As Long, W As Long, H As Long) ', oColor As OLE_COLOR, Opacity As Long)
Dim iPts() As POINTL
'Dim hPen As Long
'Dim hBrush As Long

  If mShape = 0 Then
      ReDim iPts(5)
            
      iPts(0).X = X + 0
      iPts(0).Y = Y + (W / 3) * 2
      iPts(1).X = X + (W / 3)
      iPts(1).Y = Y + (H / 2)
      iPts(2).X = X + (W / 2)
      iPts(2).Y = Y + (H / 3) * 2
      iPts(3).X = X + W
      iPts(3).Y = Y + 0
      iPts(4).X = X + W
      iPts(4).Y = Y + (H / 6)
      iPts(5).X = X + (W / 2)
      iPts(5).Y = Y + H
      
  '---------
  ElseIf mShape = 1 Then
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

  End If
          
 ' GdipCreateSolidFill ARGB(oColor, Opacity), hBrush
  GdipFillPolygonI iGraphics, hBrush, iPts(0), UBound(iPts) + 1, &H0
  
 ' GdipCreatePen1 ARGB(oColor, Opacity), 1, UnitPixel, hPen
  GdipDrawPolygonI iGraphics, hPen, iPts(0), UBound(iPts) + 1
          
'  Call GdipDeleteBrush(hBrush)
'  Call GdipDeletePen(hPen)
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

Private Function DrawRoundRect(ByVal hGraphics As Long, Rect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal Angulo As Single, ByVal BorderWidth As Long, ByVal BorderColor As Long, ByVal Round As Long) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mpath As Long
    Dim mRound As Long
    
    If m_BorderWidth <> 0 Then
      GdipCreatePen1 BorderColor, BorderWidth * nScale, &H2, hPen  '&H1 * nScale, &H2, hPen
    End If
    GdipCreateLineBrushFromRectWithAngleI Rect, color1, color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mpath   '&H0
    
    With Rect
        mRound = GetSafeRound((Round * nScale), .Width * 2, .Height * 2)
        If mRound = 0 Then mRound = 1
            'GdipDrawRectangleI hGraphics, hPen, .Left, .Top, .Width, .Height
            'GdipAddPathLineI mPath, .Left, .Top, .Width, .Top       'Line-Top
            'GdipAddPathLineI mPath, .Width, .Top, .Width, .Height   'Line-Left
            'GdipAddPathLineI mPath, .Width, .Height, .Left, .Height 'Line-Bottom
            'GdipAddPathLineI mPath, .Left, .Height, .Left, .Top     'Line-Right
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

    DrawRoundRect = mpath
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



'*-

Private Sub OptBehavior()
Dim Frm As Form
    Set Frm = Extender.Parent
    
Dim lHwnd As Long
    lHwnd = UserControl.ContainerHwnd

    Dim Ctrl As Control
    For Each Ctrl In Frm.Controls
        With Ctrl
           If TypeOf Ctrl Is AxGOption Then
              If .OptionBehavior = True Then
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
    m_OnFocus = True
    Refresh
  Else
    Exit Sub
  End If
Else
  m_Opacity = m_InitialOpacity
  m_OnFocus = False
  Refresh
  tmrEffect.Enabled = False
End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  Refresh
End Sub

Private Sub UserControl_Click()
  If m_OptionBehavior = True Then
    m_Value = True
    OptBehavior
  Else
    m_Value = Not m_Value
  End If

  PropertyChanged "Value"
  RaiseEvent Click
  RaiseEvent ChangeValue(m_Value)
  Refresh
End Sub

Private Sub UserControl_Initialize()
    InitGDI
    nScale = GetWindowsDPI
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
hFontCollection = ReadValue(&HFC)
cl_hWnd = UserControl.ContainerHwnd

m_Clicked = False
m_Value = False
Set m_Font = UserControl.Ambient.Font

  m_BorderColor = &HC0&
  'm_BorderColorFocus = &HC0&
  m_ForeColor = m_def_ForeColor
  m_ForeColor2 = &HFFFFFF
  m_Enabled = True
  m_Color1 = m_def_Color1
  m_Color2 = m_def_Color2
  m_Angulo = m_def_Angulo
  m_BorderWidth = 4
  m_CornerCurve = 30
  m_Caption = Ambient.DisplayName
  m_CaptionEnabled = True
  m_Transparent = True
  m_Clicked = False
  m_Opacity = 50
  m_InitialOpacity = m_Opacity
  m_IconCharCodeOn = "&H0"
  m_IconCharCodeOff = "&H0"
  m_CheckStyle = stDrawing
  m_Style = 0
  
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
Refresh
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointerHands True
tmrEffect.Enabled = True
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
m_Clicked = False
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
  m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
  m_ForeColor2 = .ReadProperty("ForeColor2", m_def_ForeColor)
  m_Angulo = .ReadProperty("BackAngle", m_def_Angulo)
  m_BorderColor = .ReadProperty("BorderColor", &HC0&)
  'm_BorderColorFocus = .ReadProperty("BorderColorFocus", &HC0&)
  m_BorderWidth = .ReadProperty("BorderWidth", 1)
  m_CornerCurve = .ReadProperty("CornerCurve", 0)
  
'  m_FillColor = .ReadProperty("FillColor", vbBlack)
'  m_FillEnable = .ReadProperty("FillEnable", False)
  
  m_BoxPosition = .ReadProperty("BoxPosition", cLeft)
  m_CheckVisible = .ReadProperty("CheckVisible", False)
  
  Set m_Font = .ReadProperty("Font", UserControl.Ambient.Font)
  
  m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
  m_CaptionEnabled = .ReadProperty("CaptionEnabled", False)
  m_CaptionAlignV = .ReadProperty("CaptionAlignV", 1)
  m_CaptionAlignH = .ReadProperty("CaptionAlignH", 1)
  m_CaptionAngle = .ReadProperty("CaptionAngle", 0)
  
  m_Transparent = .ReadProperty("Transparent", True)
  
  m_Line1 = .ReadProperty("Line1", False)
  
  m_ActiveColor = .ReadProperty("ActiveColor", vbWhite)
  m_CheckColor = .ReadProperty("CheckColor", vbBlue)
  
  'm_ChangeBorderOnFocus = .ReadProperty("ChangeColorOnFocus", False)
  m_EffectFade = .ReadProperty("EffectFading", False)
  m_InitialOpacity = .ReadProperty("InitialOpacity", 50)
  
  Set m_IconFont = .ReadProperty("IconFont", UserControl.Ambient.Font)
  m_IconCharCodeOn = .ReadProperty("IconCharCodeOn", "&H0")
  m_IconForeColorOn = .ReadProperty("IconForeColorOn", &H404040)
  m_IconCharCodeOff = .ReadProperty("IconCharCodeOff", "&H0")
  m_IconForeColorOff = .ReadProperty("IconForeColorOff", &H404040)
  
  m_PadX = .ReadProperty("IcoPaddingX", 0)
  m_PadY = .ReadProperty("IcoPaddingY", 0)

  m_Value = .ReadProperty("Value", False)
  m_OptionBehavior = .ReadProperty("OptionBehavior", False)
  'm_CheckStyle = .ReadProperty("CheckStyle", 0)
  m_Style = .ReadProperty("Style", 0)
  m_Opacity = m_InitialOpacity

End With
  
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
  Call .WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
  Call .WriteProperty("ForeColor2", m_ForeColor2, m_def_ForeColor)
  Call .WriteProperty("BackAngle", m_Angulo, m_def_Angulo)
  Call .WriteProperty("BorderColor", m_BorderColor, &HC0&)
  'Call .WriteProperty("BorderColorFocus", m_BorderColorFocus, &HC0&)
  Call .WriteProperty("BorderWidth", m_BorderWidth, 1)
  
  Call .WriteProperty("ActiveColor", m_ActiveColor, vbWhite)
  Call .WriteProperty("CheckColor", m_CheckColor)
  
'  Call .WriteProperty("FillColor", m_FillColor)
'  Call .WriteProperty("FillEnable", m_FillEnable)
  
  Call .WriteProperty("CornerCurve", m_CornerCurve, 0)
  Call .WriteProperty("BoxPosition", m_BoxPosition, cLeft)
  Call .WriteProperty("CheckVisible", m_CheckVisible, False)
  
  Call .WriteProperty("Font", m_Font, UserControl.Ambient.Font)
  Call .WriteProperty("CaptionAngle", m_CaptionAngle, 0)
  Call .WriteProperty("Caption", m_Caption, Ambient.DisplayName)
  Call .WriteProperty("CaptionEnabled", m_CaptionEnabled, False)
  Call .WriteProperty("CaptionAlignV", m_CaptionAlignV, 1)
  Call .WriteProperty("CaptionAlignH", m_CaptionAlignH, 1)
  
  Call .WriteProperty("Transparent", m_Transparent, True)
  
  Call .WriteProperty("Line1", m_Line1, False)
  
  'Call .WriteProperty("ChangeColorOnFocus", m_ChangeBorderOnFocus, False)
  Call .WriteProperty("EffectFading", m_EffectFade, False)
  Call .WriteProperty("InitialOpacity", m_InitialOpacity, 50)
  
  Call .WriteProperty("IconFont", m_IconFont)
  Call .WriteProperty("IconCharCodeOn", m_IconCharCodeOn, 0)
  Call .WriteProperty("IconCharCodeOff", m_IconCharCodeOff, 0)
  Call .WriteProperty("IconForeColorOn", m_IconForeColorOn, vbButtonText)
  Call .WriteProperty("IconForeColorOff", m_IconForeColorOff, vbButtonText)
  
  Call .WriteProperty("IcoPaddingX", m_PadX)
  Call .WriteProperty("IcoPaddingY", m_PadY)
  'Call .WriteProperty("CheckStyle", m_CheckStyle)
  Call .WriteProperty("Style", m_Style)
  Call .WriteProperty("Value", m_Value)
  Call .WriteProperty("OptionBehavior", m_OptionBehavior)

End With
  
End Sub

Public Property Get ActiveColor() As OLE_COLOR
  ActiveColor = m_ActiveColor
End Property

Public Property Let ActiveColor(ByVal NewActiveColor As OLE_COLOR)
  m_ActiveColor = NewActiveColor
  PropertyChanged "ActiveColor"
End Property

Public Property Get BackAngle() As Single
  BackAngle = m_Angulo
End Property

Public Property Let BackAngle(ByVal New_Angulo As Single)
  m_Angulo = New_Angulo
  PropertyChanged "BackAngle"
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
  PropertyChanged "BorderColor"
  Refresh
End Property

'Public Property Get BorderColorFocus() As OLE_COLOR
'  BorderColorFocus = m_BorderColorFocus
'End Property
'
'Public Property Let BorderColorFocus(ByVal NewBorderColorF As OLE_COLOR)
'  m_BorderColorFocus = NewBorderColorF
'  PropertyChanged "BorderColorFocus"
'  Refresh
'End Property

Public Property Get BorderWidth() As Long
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  Refresh
End Property

Public Property Get BoxPosition() As CheckPos
    BoxPosition = m_BoxPosition
End Property

Public Property Let BoxPosition(ByVal New_Value As CheckPos)
    m_BoxPosition = New_Value
    PropertyChanged "BoxPosition"
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

Public Property Get CaptionEnabled() As Boolean
  CaptionEnabled = m_CaptionEnabled
End Property

Public Property Let CaptionEnabled(ByVal NewCaptionEnabled As Boolean)
  m_CaptionEnabled = NewCaptionEnabled
  PropertyChanged "CaptionEnabled"
  Refresh
End Property

Public Property Get Caption() As String
  Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
  m_Caption = NewCaption
  PropertyChanged "Caption"
  Refresh
End Property

'Public Property Get CheckStyle() As StyleDraw
'  CheckStyle = m_CheckStyle
'End Property
'
'Public Property Let CheckStyle(ByVal NewCheckStyle As StyleDraw)
'  m_CheckStyle = NewCheckStyle
'  PropertyChanged "CheckStyle"
'  Refresh
'End Property

'Public Property Get ChangeBorderOnFocus() As Boolean
'  ChangeBorderOnFocus = m_ChangeBorderOnFocus
'End Property
'
'Public Property Let ChangeBorderOnFocus(ByVal NewChangeBorderOnFocus As Boolean)
'  m_ChangeBorderOnFocus = NewChangeBorderOnFocus
'  PropertyChanged "ChangeBorderOnFocus"
'End Property

Public Property Get CheckColor() As OLE_COLOR
  CheckColor = m_CheckColor
End Property

Public Property Let CheckColor(ByVal NCheckColor As OLE_COLOR)
  m_CheckColor = NCheckColor
  PropertyChanged "CheckColor"
  Refresh
End Property

Public Property Get CheckVisible() As Boolean
    CheckVisible = m_CheckVisible
End Property

Public Property Let CheckVisible(ByVal New_Value As Boolean)
    m_CheckVisible = New_Value
    PropertyChanged "CheckVisible"
    Refresh
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
End Property

'Public Property Get FillColor() As OLE_COLOR
'  FillColor = m_FillColor
'End Property
'
'Public Property Let FillColor(ByVal New_Color As OLE_COLOR)
'  m_FillColor = New_Color
'  PropertyChanged "FillColor"
'  Refresh
'End Property
'
'Public Property Get FillEnable() As Boolean
'  FillEnable = m_FillEnable
'End Property
'
'Public Property Let FillEnable(ByVal bEnable As Boolean)
'  m_FillEnable = bEnable
'  PropertyChanged "FillEnable"
'  Refresh
'End Property

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

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get IconCharCodeOff() As String
    IconCharCodeOff = "&H" & Hex(m_IconCharCodeOff)
End Property

Public Property Let IconCharCodeOff(ByVal New_IconCharCode As String)
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not VBA.Left$(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        m_IconCharCodeOff = "&H" & New_IconCharCode
    Else
        m_IconCharCodeOff = New_IconCharCode
    End If
    PropertyChanged "IconCharCodeOff"
    Refresh
End Property

Public Property Get IconCharCodeOn() As String
    IconCharCodeOn = "&H" & Hex(m_IconCharCodeOn)
End Property

Public Property Let IconCharCodeOn(ByVal New_IconCharCode As String)
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not VBA.Left$(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        m_IconCharCodeOn = "&H" & New_IconCharCode
    Else
        m_IconCharCodeOn = New_IconCharCode
    End If
    PropertyChanged "IconCharCodeOn"
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

Public Property Get IconForeColorOff() As OLE_COLOR
    IconForeColorOff = m_IconForeColorOff
End Property

Public Property Let IconForeColorOff(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColorOff = New_ForeColor
    PropertyChanged "IconForeColorOff"
    Refresh
End Property

Public Property Get IconForeColorOn() As OLE_COLOR
    IconForeColorOn = m_IconForeColorOn
End Property

Public Property Let IconForeColorOn(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColorOn = New_ForeColor
    PropertyChanged "IconForeColorOn"
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

Public Property Get Line1() As Boolean
  Line1 = m_Line1
End Property

Public Property Let Line1(ByVal bLine1 As Boolean)
  m_Line1 = bLine1
  PropertyChanged "Line1"
  Refresh
End Property

Public Property Get OptionBehavior() As Boolean
   OptionBehavior = m_OptionBehavior
End Property

Public Property Let OptionBehavior(ByVal bOptionBehavior As Boolean)
   m_OptionBehavior = bOptionBehavior
   PropertyChanged "OptionBehavior"
End Property

Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal NewValue As Boolean)
    m_Transparent = NewValue
    PropertyChanged "Transparent"
    Refresh
End Property

Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Boolean)
  If m_OptionBehavior And NewValue Then OptBehavior
    m_Value = NewValue
    m_Clicked = NewValue
    PropertyChanged "Value"
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

Public Property Get Style() As OpStyle
  Style = m_Style
End Property

Public Property Let Style(ByVal NewStyle As OpStyle)
  m_Style = NewStyle
  PropertyChanged "Style"
  Refresh
End Property
