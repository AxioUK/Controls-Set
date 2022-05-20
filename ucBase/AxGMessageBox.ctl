VERSION 5.00
Begin VB.UserControl AxGMessageBox 
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
   ToolboxBitmap   =   "AxGMessageBox.ctx":0000
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   75
      Top             =   75
   End
End
Attribute VB_Name = "AxGMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxGMessageBox
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

Private Type POINTAPI
    X As Long
    Y As Long
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

'EVENTS------------------------------------
Public Event Click()
Public Event ButtonClick(ButtonPress As ButtonResult)
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
Dim m_ModalColor    As OLE_COLOR
Dim m_OldBorderColor As OLE_COLOR
Dim m_BorderColorOnFocus As OLE_COLOR
Dim m_BorderWidth   As Long
Dim m_Enabled       As Boolean
Dim m_Visible       As Boolean
Dim m_Angulo        As Single
Dim m_CornerCurve   As Long
Dim m_Top         As Long
Dim m_Left        As Long
Dim OldX          As Single
Dim OldY          As Single
Dim m_Opacity     As Long
Dim m_ModalOpacity  As Long
Dim cl_hWnd       As Long

Dim m_Message       As String
Dim m_MessageAlignV As eTextAlignV
Dim m_MessageAlignH As eTextAlignH

Dim m_ChangeBorderOnFocus As Boolean
Dim m_OnFocus As Boolean
Dim m_Modal   As Boolean
Dim m_Moveable      As Boolean
Dim m_MouseOver     As Boolean

Private m_Font          As StdFont
Private m_IconFont      As StdFont
Private m_IconCharCode  As Long
Private m_IconForeColor As Long
Private m_PadY          As Long
Private m_PadX          As Long

Private m_CaptionAngle  As Single
Private m_CaptionX      As Long
Private m_CaptionY      As Long
Private m_Caption       As String
Private m_CaptionEnabled  As Boolean
Private m_CaptionAlignV As eTextAlignV
Private m_CaptionAlignH As eTextAlignH

Private m_StringPosX  As Long
Private m_StringPosY  As Long
Private m_EffectFade  As Boolean
Private m_InitialOpacity As Long
Private m_Transparent As Boolean

Private m_Clicked      As Boolean
Private m_Filled       As Boolean

Private RECb    As RECTL
Private Button1 As RECTL
Private Button2 As RECTL

Private m_Parent As Object


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

Public Sub Show(Optional fParent As Object)
Set m_Parent = fParent
If m_Modal Then
  If m_Parent Is Nothing Then
    Debug.Print "ERROR:Debe indicar fParent para usar Modal"
  Else
    m_Parent.AutoRedraw = False
    DrawShaded True, Me, m_Parent, m_ModalColor
  End If
End If
Visible = True
DrawMessage
End Sub

Private Sub DrawMessage()
  UserControl.Cls
  CopyAmbient
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

  If m_MouseOver Then
    .ForeColor = m_ForeColor
  Else
    .ForeColor = m_IconForeColor
  End If
  
  Rct.Left = (IconFont.Size / 2) + m_PadX
  Rct.Top = (IconFont.Size / 2) + m_PadY
  Rct.Right = UserControl.ScaleWidth
  Rct.Bottom = UserControl.ScaleHeight
  
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
Dim fRec As RECTL
Dim stCap As RECTS
Dim stMsg As RECTS
Dim stBt1 As RECTS
Dim stBt2 As RECTS
Dim lBorder As Long, mBorder As Long
Dim m_PosX1 As Integer, m_PosY1 As Integer
Dim m_PosX2 As Integer, m_PosY2 As Integer

With UserControl
    
  GdipCreateFromHDC .hDC, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
  
  SafeRange m_Opacity, 0, 100

  lBorder = m_BorderWidth * 2
  mBorder = lBorder / 2
      
  'Body
  REC.Left = 1 * nScale
  REC.Top = 1 * nScale
  REC.Width = .ScaleWidth - 2 * nScale
  REC.Height = .ScaleHeight - 2 * nScale
  fRoundRect hGraphics, REC, argb(m_Color2, 100), argb(m_Color1, 100), m_Angulo, argb(m_BorderColor, 100), m_CornerCurve, m_Filled
  
  'TopBar
  RECb.Left = 1 * nScale
  RECb.Top = 1 * nScale
  RECb.Width = .ScaleWidth - 2 * nScale
  RECb.Height = (Font.Size + 12) * nScale
  fRoundCut hGraphics, RECb, argb(m_Color2, 100), argb(m_Color1, 100), m_Angulo, argb(m_BorderColor, 100), m_CornerCurve, m_Filled, rBottom
  
  'Caption
  stCap.Left = mBorder * nScale
  stCap.Top = mBorder * nScale
  stCap.Width = .ScaleWidth - lBorder * nScale
  stCap.Height = (Font.Size + 10) * nScale
  DrawString hGraphics, m_Caption, stCap, m_ForeColor, 100, 0, m_CaptionAlignH, m_CaptionAlignV, True

  'Button1
  Button1.Left = .ScaleWidth - (.TextWidth("Aceptar") + 25) * nScale
  Button1.Top = .ScaleHeight - (.TextHeight("Áj") + 20) * nScale
  Button1.Width = .TextWidth("Aceptar") + 10 * nScale
  Button1.Height = .TextHeight("Áj") + 5 * nScale
  stBt1.Left = Button1.Left: stBt1.Top = Button1.Top: stBt1.Width = Button1.Width: stBt1.Height = Button1.Height
  fRoundRect hGraphics, Button1, argb(m_Color2, 100), argb(m_Color1, 100), m_Angulo, argb(m_BorderColor, 100), m_CornerCurve, m_Filled
  DrawString hGraphics, "Aceptar", stBt1, m_ForeColor, 100, 0, eCenter, eMiddle, True

  'Button2
  Button2.Left = .ScaleWidth - (.TextWidth("Cancelar") + Button1.Width + 40) * nScale
  Button2.Top = .ScaleHeight - (.TextHeight("Áj") + 20) * nScale
  Button2.Width = .TextWidth("Cancelar") + 10 * nScale
  Button2.Height = .TextHeight("Áj") + 5 * nScale
  stBt2.Left = Button2.Left: stBt2.Top = Button2.Top: stBt2.Width = Button2.Width: stBt2.Height = Button2.Height
  fRoundRect hGraphics, Button2, argb(m_Color2, 100), argb(m_Color1, 100), m_Angulo, argb(m_BorderColor, 100), m_CornerCurve, m_Filled
  DrawString hGraphics, "Cancelar", stBt2, m_ForeColor, 100, 0, eCenter, eMiddle, True

  'Message
  stMsg.Left = 20: stMsg.Top = .ScaleHeight / 5
  stMsg.Width = .ScaleWidth - 40: stMsg.Height = .ScaleHeight - (.ScaleHeight / 2)
  DrawString hGraphics, m_Message, stMsg, m_ForeColor, 100, 0, m_MessageAlignH, m_MessageAlignV, True

' '---------------
'  If m_EffectFade Then
'        GdipCreateSolidFill ARGB(m_BorderColorOnFocus, m_Opacity), hBrush
'        GdipCreatePen1 ARGB(m_BorderColorOnFocus, m_Opacity), 1, UnitPixel, hPen
'  Else
'        GdipCreateSolidFill ARGB(m_BorderColorOnFocus, 100), hBrush
'        GdipCreatePen1 ARGB(m_BorderColorOnFocus, 100), 1, UnitPixel, hPen
'  End If
'  '---------------
'  '---------------
'  Call GdipDeleteBrush(hBrush)
'  Call GdipDeletePen(hPen)
'  '---------------
  GdipDeleteGraphics hGraphics
  '---------------
  .BackStyle = 0
  .MaskColor = .BackColor
  Set .MaskPicture = .Image
  '---------------
End With

End Sub

Private Function DrawString(ByVal hGraphics As Long, sString As String, layoutRect As RECTS, TextColor As OLE_COLOR, _
                             ColorOpacity As Integer, mAngle As Single, HAlign As eTextAlignH, VAlign As eTextAlignV, _
                             Optional bWordWrap As Boolean = True) As Long
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

Private Sub DrawShape(ByVal iGraphics As Long, mShape As Integer, X As Long, Y As Long, W As Long, H As Long) ', oColor As OLE_COLOR, Opacity As Long)
Dim iPts() As POINTL
'Dim hPen As Long
'Dim hBrush As Long

  If mShape = 0 Then
      ReDim iPts(5)
            
      iPts(0).X = X + 0
      iPts(0).Y = Y + (H / 3) * 2
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
          
  GdipFillPolygonI iGraphics, hBrush, iPts(0), UBound(iPts) + 1, &H0
  GdipDrawPolygonI iGraphics, hPen, iPts(0), UBound(iPts) + 1
          
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

Private Function fRoundCut(ByVal hGraphics As Long, Rect As RECTL, ByVal color1 As Long, ByVal color2 As Long, _
                           ByVal Angulo As Single, ByVal BorderColor As Long, ByVal Round As Long, _
                           Filled As Boolean, SideCut As eFlatSide) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mpath As Long
    Dim mRound As Long
    
    GdipCreateSolidFill BackColor, hBrush
    If BorderWidth <> 0 Then GdipCreatePen1 BorderColor, m_BorderWidth * nScale, &H2, hPen
    If Filled Then GdipCreateLineBrushFromRectWithAngleI Rect, color1, color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mpath
    

    With Rect
        mRound = GetSafeRound((Round * nScale), .Width, .Height)
        If mRound = 0 Then mRound = 1
        Select Case SideCut
          Case rUp
              GdipAddPathArcI mpath, .Left, .Top, 1, 1, 180, 90
              GdipAddPathArcI mpath, (.Left + .Width) - 1, .Top, 1, 1, 270, 90
              GdipAddPathArcI mpath, (.Left + .Width) - mRound, (.Top + .Height) - mRound, mRound, mRound, 0, 90
              GdipAddPathArcI mpath, .Left, (.Top + .Height) - mRound, mRound, mRound, 90, 90
          Case rBottom
              GdipAddPathArcI mpath, .Left, .Top, mRound, mRound, 180, 90
              GdipAddPathArcI mpath, (.Left + .Width) - mRound, .Top, mRound, mRound, 270, 90
              GdipAddPathArcI mpath, (.Left + .Width) - 1, (.Top + .Height) - 1, 1, 1, 0, 90
              GdipAddPathArcI mpath, .Left, (.Top + .Height) - 1, 1, 1, 90, 90
          Case rLeft
              GdipAddPathArcI mpath, .Left, .Top, 1, 1, 180, 90
              GdipAddPathArcI mpath, (.Left + .Width) - mRound, .Top, mRound, mRound, 270, 90
              GdipAddPathArcI mpath, (.Left + .Width) - mRound, (.Top + .Height) - mRound, mRound, mRound, 0, 90
              GdipAddPathArcI mpath, .Left, (.Top + .Height) - 1, 1, 1, 90, 90
          Case rRight
              GdipAddPathArcI mpath, .Left, .Top, mRound, mRound, 180, 90
              GdipAddPathArcI mpath, (.Left + .Width) - 1, .Top, 1, 1, 270, 90
              GdipAddPathArcI mpath, (.Left + .Width) - 1, (.Top + .Height) - 1, 1, 1, 0, 90
              GdipAddPathArcI mpath, .Left, (.Top + .Height) - mRound, mRound, mRound, 90, 90
        End Select
    End With
    
    GdipClosePathFigures mpath
    GdipFillPath hGraphics, hBrush, mpath
    GdipDrawPath hGraphics, hPen, mpath
    
    Call GdipDeletePath(mpath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)

    fRoundCut = mpath
End Function

Private Function fRoundRect(ByVal hGraphics As Long, Rect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal Angulo As Single, ByVal BorderColor As Long, ByVal Round As Long, Filled As Boolean) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mpath As Long
    Dim mRound As Long
    
    If m_BorderWidth <> 0 Then GdipCreatePen1 BorderColor, m_BorderWidth * nScale, &H2, hPen   '&H1 * nScale, &H2, hPen
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
    
    Call GdipDeletePath(mpath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)

    fRoundRect = mpath
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
    DrawMessage
  Else
    Exit Sub
  End If
Else
  m_Opacity = m_InitialOpacity
  DrawMessage
  tmrEffect.Enabled = False
End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  DrawMessage
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

  m_Clicked = False
  Set m_Font = UserControl.Ambient.Font
  m_CaptionAlignV = eMiddle
  m_CaptionAlignH = eCenter
  m_Caption = Ambient.DisplayName

  m_MessageAlignV = eMiddle
  m_MessageAlignH = eCenter
  m_Message = Ambient.DisplayName

'  m_CaptionX = 0
'  m_CaptionY = 0

  m_Filled = True

  m_BorderColor = &H404040
  m_OldBorderColor = &HC0&
  m_ForeColor = m_def_ForeColor
  m_ForeColor2 = &HFFFFFF
  m_Enabled = True
  m_Color1 = m_def_Color1
  m_Color2 = m_def_Color2
  m_Angulo = m_def_Angulo
  m_BorderWidth = 2
  m_CornerCurve = 10
  m_Transparent = True
  m_Opacity = 50
  m_InitialOpacity = m_Opacity
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

If X > Button1.Left And X < (Button1.Left + Button1.Width) And Y > Button1.Top And Y < (Button1.Top + Button1.Height) Then
    RaiseEvent ButtonClick(vrOK)
    GoTo HideM
ElseIf X > Button2.Left And X < (Button2.Left + Button2.Width) And Y > Button2.Top And Y < (Button2.Top + Button2.Height) Then
    RaiseEvent ButtonClick(vrCancel)
    GoTo HideM
Else
    GoTo RefreshM
End If

HideM:
If m_Modal Then
  If m_Parent Is Nothing Then
    Debug.Print "ERROR:Debe indicar fParent para usar Modal"
  Else
    DrawShaded False, Me, m_Parent, m_ModalColor
  End If
End If

Visible = False

RefreshM:
m_Clicked = True
DrawMessage
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
If X > RECb.Left And X < (RECb.Left + RECb.Width) And Y > RECb.Top And Y < (RECb.Top + RECb.Height) Then
  If m_Moveable Then
    MousePointerHands True
    If Button = vbLeftButton Then
      Dim res As Long
      Call ReleaseCapture
      res = SendMessage(UserControl.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
  End If
ElseIf X > Button1.Left And X < (Button1.Left + Button1.Width) And Y > Button1.Top And Y < (Button1.Top + Button1.Height) Then
    MousePointerHands True
ElseIf X > Button2.Left And X < (Button2.Left + Button2.Width) And Y > Button2.Top And Y < (Button2.Top + Button2.Height) Then
    MousePointerHands True
Else
    MousePointerHands False
End If

m_MouseOver = True
tmrEffect.Enabled = True
RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
m_Clicked = False
DrawMessage

MousePointerHands False
RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Paint()
DrawMessage
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
  m_BorderColor = .ReadProperty("BorderColor", &H404040)
  m_BorderWidth = .ReadProperty("BorderWidth", 1)
  m_CornerCurve = .ReadProperty("CornerCurve", 0)
  m_Filled = .ReadProperty("Filled", True)
  
  m_ModalColor = .ReadProperty("ModalColor", vbBlack)
  m_Modal = .ReadProperty("Modal", False)
  m_ModalOpacity = .ReadProperty("ModalOpacity", 30)
  m_Moveable = .ReadProperty("Moveable", False)
  
  Set m_Font = .ReadProperty("Font", UserControl.Ambient.Font)
  
  m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
  m_CaptionAlignV = .ReadProperty("CaptionAlignV", 1)
  m_CaptionAlignH = .ReadProperty("CaptionAlignH", 1)
  
  m_Message = .ReadProperty("Message", Ambient.DisplayName)
  m_MessageAlignV = .ReadProperty("MessageAlignV", 1)
  m_MessageAlignH = .ReadProperty("MessageAlignH", 1)
  
'  m_CaptionX = .ReadProperty("CaptionX", 0)
'  m_CaptionY = .ReadProperty("CaptionY", 0)
      
  m_BorderColorOnFocus = .ReadProperty("ColorOnFocus", vbWhite)
  m_ChangeBorderOnFocus = .ReadProperty("ChangeColorOnFocus", False)
  m_EffectFade = .ReadProperty("EffectFading", False)
  m_InitialOpacity = .ReadProperty("InitialOpacity", 50)
  
  Set m_IconFont = .ReadProperty("IconFont", Ambient.Font)
  m_IconCharCode = .ReadProperty("IconCharCode", "&H0")
  m_IconForeColor = .ReadProperty("IconForeColor", &H404040)
  
  m_PadX = .ReadProperty("IcoPaddingX", 0)
  m_PadY = .ReadProperty("IcoPaddingY", 0)
  
End With
  
  m_Opacity = m_InitialOpacity
  Extender.Visible = False
End Sub

Private Sub UserControl_Resize()
DrawMessage
End Sub

Private Sub UserControl_Show()
If Ambient.UserMode = True Then
  Dim Frm As Form
  Set Frm = Extender.Parent
  If Frm.MDIChild = True Then MsgBox "Para su correcto funcionamiento no se recomienda" & vbLf & "el uso de este control en formularios MDIChild", vbOKOnly, "AxGMessageBox"
End If

DrawMessage
End Sub

Private Sub UserControl_Terminate()
TerminateGDI
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("Enabled", m_Enabled)
  'Call .WriteProperty("Visible", m_Visible)
  
  Call .WriteProperty("BackColor1", m_Color1, m_def_Color1)
  Call .WriteProperty("BackColor2", m_Color2, m_def_Color2)
  Call .WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
  Call .WriteProperty("ForeColor2", m_ForeColor2, m_def_ForeColor)
  Call .WriteProperty("BackAngle", m_Angulo, m_def_Angulo)
  Call .WriteProperty("BorderColor", m_BorderColor, &HC0&)
  Call .WriteProperty("BorderWidth", m_BorderWidth, 1)
  Call .WriteProperty("CornerCurve", m_CornerCurve, 0)
  Call .WriteProperty("Filled", m_Filled)
  
  Call .WriteProperty("ModalColor", m_ModalColor, vbBlack)
  Call .WriteProperty("Modal", m_Modal, False)
  Call .WriteProperty("ModalOpacity", m_ModalOpacity, 30)
  Call .WriteProperty("Moveable", m_Moveable, False)
  
  Call .WriteProperty("Font", m_Font, UserControl.Ambient.Font)
  Call .WriteProperty("Caption", m_Caption, Ambient.DisplayName)
  Call .WriteProperty("CaptionAlignV", m_CaptionAlignV, 1)
  Call .WriteProperty("CaptionAlignH", m_CaptionAlignH, 1)
  
  Call .WriteProperty("Message", m_Message, Ambient.DisplayName)
  Call .WriteProperty("MessageAlignV", m_MessageAlignV, 1)
  Call .WriteProperty("MessageAlignH", m_MessageAlignH, 1)
  
  'Call .WriteProperty("CaptionX", m_CaptionX)
  'Call .WriteProperty("CaptionY", m_CaptionY)
  
  Call .WriteProperty("ColorOnFocus", m_BorderColorOnFocus, vbWhite)
  Call .WriteProperty("ChangeColorOnFocus", m_ChangeBorderOnFocus, False)
  Call .WriteProperty("EffectFading", m_EffectFade, False)
  Call .WriteProperty("InitialOpacity", m_InitialOpacity, 50)
  
  Call .WriteProperty("IconFont", m_IconFont)
  Call .WriteProperty("IconCharCode", m_IconCharCode, 0)
  Call .WriteProperty("IconForeColor", m_IconForeColor, vbButtonText)
  
  Call .WriteProperty("IcoPaddingX", m_PadX)
  Call .WriteProperty("IcoPaddingY", m_PadY)

End With
  
End Sub

Public Property Get BackAngle() As Single
  BackAngle = m_Angulo
End Property

Public Property Let BackAngle(ByVal New_Angulo As Single)
  m_Angulo = New_Angulo
  PropertyChanged "BackAngle"
  DrawMessage
End Property

Public Property Get BackColor1() As OLE_COLOR
  BackColor1 = m_Color1
End Property

Public Property Let BackColor1(ByVal New_Color1 As OLE_COLOR)
  m_Color1 = New_Color1
  PropertyChanged "BackColor1"
  DrawMessage
End Property

Public Property Get BackColor2() As OLE_COLOR
  BackColor2 = m_Color2
End Property

Public Property Let BackColor2(ByVal New_Color2 As OLE_COLOR)
  m_Color2 = New_Color2
  PropertyChanged "BackColor2"
  DrawMessage
End Property

'Properties-------------------
Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
  m_BorderColor = NewBorderColor
  'm_OldBorderColor = m_BorderColor
  PropertyChanged "BorderColor"
  DrawMessage
End Property

Public Property Get BorderColorOnFocus() As OLE_COLOR
  BorderColorOnFocus = m_BorderColorOnFocus
End Property

Public Property Let BorderColorOnFocus(ByVal NewBorderColorOnFocus As OLE_COLOR)
  m_BorderColorOnFocus = NewBorderColorOnFocus
  PropertyChanged "BorderColorOnFocus"
  DrawMessage
End Property

Public Property Get BorderWidth() As Long
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  DrawMessage
End Property

Public Property Get CaptionAlignH() As eTextAlignH
  CaptionAlignH = m_CaptionAlignH
End Property

Public Property Let CaptionAlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_CaptionAlignH = NewCaptionAlignH
  PropertyChanged "CaptionAlignH"
  DrawMessage
End Property

Public Property Get CaptionAlignV() As eTextAlignV
  CaptionAlignV = m_CaptionAlignV
End Property

Public Property Let CaptionAlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_CaptionAlignV = NewCaptionAlignV
  PropertyChanged "CaptionAlignV"
  DrawMessage
End Property

Public Property Get Caption() As String
  Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
  m_Caption = NewCaption
  PropertyChanged "Caption"
  DrawMessage
End Property

'Public Property Get CaptionX() As Long
'  CaptionX = m_CaptionX
'End Property
'
'Public Property Let CaptionX(ByVal NewPosX As Long)
'  m_CaptionX = NewPosX
'  PropertyChanged "CaptionX"
'  DrawMessage
'End Property
'
'Public Property Get CaptionY() As Long
'  CaptionY = m_CaptionY
'End Property
'
'Public Property Let CaptionY(ByVal NewPosY As Long)
'  m_CaptionY = NewPosY
'  PropertyChanged "CaptionY"
'  DrawMessage
'End Property

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
  DrawMessage
End Property

Public Property Get EffectFading() As Boolean
EffectFading = m_EffectFade
End Property

Public Property Let EffectFading(ByVal vNewValue As Boolean)
m_EffectFade = vNewValue
PropertyChanged "EffectFading"
DrawMessage
End Property

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  UserControl.Enabled = m_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get Filled() As Boolean
  Filled = m_Filled
End Property

Public Property Let Filled(ByVal New_Filled As Boolean)
  m_Filled = New_Filled
  PropertyChanged "Filled"
  DrawMessage
End Property

Public Property Get Font() As StdFont
  Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
  Set m_Font = New_Font
  PropertyChanged "Font"
  DrawMessage
End Property

Public Property Get ForeColor2() As OLE_COLOR
  ForeColor2 = m_ForeColor2
End Property

Public Property Let ForeColor2(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor2 = NewForeColor
  PropertyChanged "ForeColor2"
  DrawMessage
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor = NewForeColor
  PropertyChanged "ForeColor"
  DrawMessage
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
    DrawMessage
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
  DrawMessage
End Property

Public Property Get IconForeColor() As OLE_COLOR
    IconForeColor = m_IconForeColor
End Property

Public Property Let IconForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor = New_ForeColor
    PropertyChanged "IconForeColor"
    DrawMessage
End Property

Public Property Get IcoPaddingX() As Long
IcoPaddingX = m_PadX
End Property

Public Property Let IcoPaddingX(ByVal XpadVal As Long)
m_PadX = XpadVal
PropertyChanged "IcoPaddingX"
DrawMessage
End Property

Public Property Get IcoPaddingY() As Long
IcoPaddingY = m_PadY
End Property

Public Property Let IcoPaddingY(ByVal YpadVal As Long)
m_PadY = YpadVal
PropertyChanged "IcoPaddingY"
DrawMessage
End Property

Public Property Get InitialOpacity() As Long
  InitialOpacity = m_InitialOpacity
End Property

Public Property Let InitialOpacity(ByVal NewInitialOpacity As Long)
  m_InitialOpacity = NewInitialOpacity
  PropertyChanged "InitialOpacity"
  m_Opacity = m_InitialOpacity
  DrawMessage
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

Public Property Get Modal() As Boolean
Modal = m_Modal
End Property

Public Property Let Modal(ByVal vNewValue As Boolean)
m_Modal = vNewValue
PropertyChanged "Modal"
End Property

Public Property Get ModalColor() As OLE_COLOR
ModalColor = m_ModalColor
End Property

Public Property Let ModalColor(ByVal vNewValue As OLE_COLOR)
m_ModalColor = vNewValue
PropertyChanged "ModalColor"
End Property

Public Property Get ModalOpacity() As Long
ModalOpacity = m_ModalOpacity
End Property

Public Property Let ModalOpacity(ByVal vNewValue As Long)
m_ModalOpacity = vNewValue
PropertyChanged "ModalOpacity"
End Property

Public Property Get Moveable() As Boolean
    Moveable = m_Moveable
End Property

Public Property Let Moveable(ByVal New_Moveable As Boolean)
    m_Moveable = New_Moveable
    PropertyChanged "Moveable"
End Property

Public Property Get MessageAlignH() As eTextAlignH
  MessageAlignH = m_MessageAlignH
End Property

Public Property Let MessageAlignH(ByVal NewMessageAlignH As eTextAlignH)
  m_MessageAlignH = NewMessageAlignH
  PropertyChanged "MessageAlignH"
  DrawMessage
End Property

Public Property Get MessageAlignV() As eTextAlignV
  MessageAlignV = m_MessageAlignV
End Property

Public Property Let MessageAlignV(ByVal NewMessageAlignV As eTextAlignV)
  m_MessageAlignV = NewMessageAlignV
  PropertyChanged "MessageAlignV"
  DrawMessage
End Property

Public Property Get Message() As String
  Message = m_Message
End Property

Public Property Let Message(ByVal NewMessage As String)
  m_Message = NewMessage
  PropertyChanged "Message"
  DrawMessage
End Property
