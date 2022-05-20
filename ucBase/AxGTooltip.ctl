VERSION 5.00
Begin VB.UserControl AxGToolTip 
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
   ToolboxBitmap   =   "AxGTooltip.ctx":0000
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   75
      Top             =   75
   End
End
Attribute VB_Name = "AxGToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxGToolTip
'Version  : 2.07.6
'Editor   : David Rojas [AxioUK]
'Date     : 19/05/2022
'------------------------------------
Option Explicit

Private Const VERS As String = "2.07.6"


'Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SetRECL Lib "user32" Alias "SetRect" (lpRect As RECTL, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal token As Long)
'Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
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
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal RGBA As Long, ByRef brush As Long) As Long
'Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
'---
Private Declare Function GdipCreateRegionPath Lib "GdiPlus.dll" (ByVal mpath As Long, Region As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
'---
'Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
'Private Declare Function TextOutW Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function DrawTextA Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
'Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
'Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long

Private Const LF_FACESIZE = 32
'Private Const SYSTEM_FONT = 13
'Private Const OBJ_FONT As Long = 6&

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE - 1) As Byte
End Type

'---
'Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTL) As Long
'Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
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

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

'Constants
'Private Const CombineModeExclude As Long = &H4
'Private Const WrapModeTileFlipXY = &H3
'Private Const SmoothingModeHighQuality As Long = &H2
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const LOGPIXELSX As Long = 88
'Private Const LOGPIXELSY As Long = 90
'Private Const TLS_MINIMUM_AVAILABLE As Long = 64
'Private Const IDC_HAND As Long = 32649
Private Const UnitPixel As Long = &H2&
'Private Const DT_CENTER As Long = &H1
'Private Const DT_LEFT As Long = &H0
'Private Const DT_RIGHT As Long = &H2


'Define EVENTS-------------------
Public Event Click()
'Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Property Variables:
'Private hFontCollection As Long
Private gdipToken As Long
Private nScale    As Single
Private hGraphics As Long

Private m_Enabled       As Boolean
Private m_BackColor     As OLE_COLOR
Private m_BackColorParent As OLE_COLOR
Private m_BorderColor   As OLE_COLOR
Private m_ForeColor1    As OLE_COLOR
Private m_ForeColor2    As OLE_COLOR
Private m_BorderWidth   As Long
Private m_CornerCurve   As Long
Private m_Transparent   As Boolean

Private m_Caption1 As String
Private m_Caption2 As String

Private m_Font1 As StdFont
Private m_Font2 As StdFont

Private m_Opacity As Long
Private m_CallOutLen As Long
Private m_CallOutWidth As Long
Private m_CallOutPos As CallOutPosition
Private m_CaptionAlign As DTAlign

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
  CopyAmbient
  Draw
  UserControl.Refresh
End Sub

Private Function AddString(mCaption As String, Rct As Rect, mFont As Font, TextColor As OLE_COLOR)
Dim pFont       As IFont
Dim lFontOld    As Long
    
On Error GoTo ErrF
With UserControl
  Set pFont = mFont
  lFontOld = SelectObject(.hDC, pFont.hFont)
  
  .ForeColor = TextColor
    
  DrawTextA .hDC, mCaption, -1, Rct, m_CaptionAlign
  
  Call SelectObject(.hDC, lFontOld)
  
ErrF:
  Set pFont = Nothing
End With
End Function

Private Sub Draw()
Dim REC As RECTL
Dim cpBar As Rect
Dim stBar As Rect

With UserControl

  .BackColor = m_BackColorParent
  
  GdipCreateFromHDC .hDC, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
  'draw Background Buble
  SetRECL REC, 1, 1, (.ScaleWidth - 2), (.ScaleHeight - 2)
  DrawBubble hGraphics, REC, RGBA(m_BorderColor, m_Opacity), 1, RGBA(m_BackColor, m_Opacity), m_CornerCurve, m_CallOutWidth, m_CallOutLen, m_CallOutPos

  'Draw Captions
  Select Case m_CallOutPos
    Case Is = coTop
        SetREC2 cpBar, 5, m_CallOutLen + 3, (.ScaleWidth - 9), (m_CallOutLen + 3) + .TextHeight(m_Caption1)
        SetREC2 stBar, 5, (m_CallOutLen + 4) + .TextHeight(m_Caption1), (.ScaleWidth - 9), (.ScaleHeight - 5)
    Case Is = coBottom
        SetREC2 cpBar, 6, 3, (.ScaleWidth - 9), .TextHeight(m_Caption1) + 3
        SetREC2 stBar, 6, .TextHeight(m_Caption1) + 4, (.ScaleWidth - 9), (.ScaleHeight - (5 + m_CallOutLen))
    Case Is = coLeft
        SetREC2 cpBar, m_CallOutLen + 5, 3, (.ScaleWidth - 5), .TextHeight(m_Caption1) + 3
        SetREC2 stBar, m_CallOutLen + 5, .TextHeight(m_Caption1) + 4, (.ScaleWidth - 5), (.ScaleHeight - 5)
    Case Is = coRight
        SetREC2 cpBar, 5, 3, (.ScaleWidth - (m_CallOutLen + 10)), .TextHeight(m_Caption1) + 3
        SetREC2 stBar, 5, .TextHeight(m_Caption1) + 4, (.ScaleWidth - (m_CallOutLen + 10)), (.ScaleHeight - 5)
  End Select
  
  AddString m_Caption1, cpBar, m_Font1, m_ForeColor1
  AddString m_Caption2, stBar, m_Font2, m_ForeColor2
  
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

Private Function DrawBubble(ByVal hGraphics As Long, Rct As RECTL, lBorderColor As Long, lBorderWidth As Long, lBackColor As Long, lCornerCurve As Long, coWidth As Long, coLen As Long, COPos As CallOutPosition) As Long
Dim mpath As Long, hPen As Long
Dim hBrush As Long, mRound As Long
Dim Xx As Long, Yy As Long
Dim lMax As Long, coAngle  As Long
Dim mRegion As Long

With Rct
        
    coAngle = coWidth / 2
    mRound = GetSafeRound(lCornerCurve * nScale, .Width, .Height)
    
    Select Case COPos
        Case coLeft
            .Left = .Left + coLen
            .Width = .Width - coLen
            lMax = .Height - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coTop
            .Top = .Top + coLen
            .Height = .Height - coLen
            lMax = .Width - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coRight
            .Width = .Width - coLen
            lMax = .Height - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coBottom
            .Height = .Height - coLen
            lMax = .Width - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
    End Select

    GdipCreatePen1 lBorderColor, lBorderWidth, UnitPixel, hPen
    GdipCreateSolidFill lBackColor, hBrush
    Call GdipCreatePath(&H0, mpath)
                    
    GdipAddPathArcI mpath, .Left, .Top, mRound * 2, mRound * 2, 180, 90

    If COPos = coTop Then
        Xx = .Left + (.Width - coWidth) / 2
        If mRound = 0 Then GdipAddPathLineI mpath, .Left, .Top, .Left, .Top
        GdipAddPathLineI mpath, Xx, .Top, Xx + coAngle, .Top - coLen
        GdipAddPathLineI mpath, Xx + coAngle, .Top - coLen, Xx + coWidth, .Top
    End If

    GdipAddPathArcI mpath, .Left + .Width - mRound * 2, .Top, mRound * 2, mRound * 2, 270, 90

    If COPos = coRight Then
        Yy = .Top + (.Height - coWidth) / 2
        Xx = .Left + .Width
        If mRound = 0 Then GdipAddPathLineI mpath, .Left + .Width, .Top, .Left + .Width, .Top
        GdipAddPathLineI mpath, Xx, Yy, Xx + coLen, Yy + coAngle
        GdipAddPathLineI mpath, Xx + coLen, Yy + coAngle, Xx, Yy + coWidth
    End If

    GdipAddPathArcI mpath, .Left + .Width - mRound * 2, .Top + .Height - mRound * 2, mRound * 2, mRound * 2, 0, 90

    If COPos = coBottom Then
        Xx = .Left + (.Width - coWidth) / 2
        Yy = .Top + .Height
        If mRound = 0 Then GdipAddPathLineI mpath, .Left + .Width, .Top + .Height, .Left + .Width, .Top + .Height
        GdipAddPathLineI mpath, Xx + coWidth, Yy, Xx + coAngle, Yy + coLen
        GdipAddPathLineI mpath, Xx + coAngle, Yy + coLen, Xx, Yy
    End If

    GdipAddPathArcI mpath, .Left, .Top + .Height - mRound * 2, mRound * 2, mRound * 2, 90, 90
    
    If COPos = coLeft Then
        Yy = .Top + (.Height - coWidth) / 2
        If mRound = 0 Then GdipAddPathLineI mpath, .Left, .Top + .Height, .Left, .Top + .Height
        GdipAddPathLineI mpath, .Left, Yy + coWidth, .Left - coLen, Yy + coAngle
        GdipAddPathLineI mpath, .Left - coLen, Yy + coAngle, .Left, Yy
    End If
End With
        
    GdipClosePathFigures mpath
    GdipFillPath hGraphics, hBrush, mpath
    GdipDrawPath hGraphics, hPen, mpath
        
    Call GdipDeletePath(mpath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)

End Function

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function

Private Function GetWindowsDPI() As Double
    Dim hDC As Long, lPx  As Double ', LPY As Double
    hDC = GetDC(0)
    lPx = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    'LPY = CDbl(GetDeviceCaps(hDC, LOGPIXELSY))
    ReleaseDC 0, hDC

    If (lPx = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = lPx / 96#
    End If
End Function

Private Sub InitGDI()
    Dim gdipStartupInput As GdiplusStartupInput
    gdipStartupInput.GdiplusVersion = 1&
    Call GdiplusStartup(gdipToken, gdipStartupInput, ByVal 0)
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

Private Function SetREC2(lpRect As Rect, ByVal X As Long, ByVal Y As Long, ByVal R As Long, ByVal B As Long) As Long
  lpRect.Left = X:    lpRect.Top = Y
  lpRect.Right = R:   lpRect.Bottom = B
End Function

Private Function SetRECS(lpRect As RECTS, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As Long
  lpRect.Left = X:    lpRect.Top = Y
  lpRect.Width = W:   lpRect.Height = H
End Function

Private Sub TerminateGDI()
    Call GdiplusShutdown(gdipToken)
End Sub

Private Sub UserControl_Initialize()
InitGDI
nScale = GetWindowsDPI

End Sub

Private Sub UserControl_InitProperties()
'hFontCollection = ReadValue(&HFC)

  m_Enabled = True
  m_BorderColor = vbRed
  m_BackColor = &H8D4214
  m_BackColorParent = vbWhite
  m_BorderWidth = 1
  m_CornerCurve = 10
  m_ForeColor1 = vbWhite
  m_ForeColor2 = vbWhite
  Set m_Font1 = UserControl.Font
  Set m_Font2 = UserControl.Font
  m_Opacity = 90
  m_CallOutLen = 10
  m_CallOutWidth = 5
  m_CallOutPos = 0
  m_Caption1 = UserControl.Ambient.DisplayName
  m_Caption2 = UserControl.Ambient.DisplayName & "_2"
  m_CaptionAlign = DT_CENTER
  m_Transparent = True
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

'*2
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
RaiseEvent Click
End Sub

'*3
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  UserControl.MousePointer = vbSizeWE
  Refresh
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl.MousePointer = vbDefault
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  m_Enabled = .ReadProperty("Enabled", True)
  m_BorderColor = .ReadProperty("BorderColor", &HFF8080)
  m_BackColorParent = .ReadProperty("BackColorParent", vbWhite)
  m_BackColor = .ReadProperty("BackColor", &H8000000F)
  m_BorderWidth = .ReadProperty("BorderWidth", 1)
  m_CornerCurve = .ReadProperty("CornerCurve", 10)
  m_ForeColor1 = .ReadProperty("ForeColor1", &H8D4214)
  m_ForeColor2 = .ReadProperty("ForeColor2", &HFFFFFF)
  Set m_Font1 = .ReadProperty("Font1", UserControl.Font)
  Set m_Font2 = .ReadProperty("Font2", UserControl.Font)
  
  m_Transparent = .ReadProperty("Transparent", True)
  
  m_Caption1 = .ReadProperty("Caption1", UserControl.Ambient.DisplayName)
  m_Caption2 = .ReadProperty("Caption2", UserControl.Ambient.DisplayName & "_2")
  m_CaptionAlign = .ReadProperty("CAptionAlign", DT_CENTER)
  
  m_Opacity = .ReadProperty("Opacity", 90)
  m_CallOutLen = .ReadProperty("CallOutLen", 5)
  m_CallOutWidth = .ReadProperty("CallOutWidth", 5)
  m_CallOutPos = .ReadProperty("CallOutPos", 0)
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

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("Enabled", m_Enabled)
  Call .WriteProperty("BorderColor", m_BorderColor)
  Call .WriteProperty("BackColor", m_BackColor)
  Call .WriteProperty("BackColorParent", m_BackColorParent)
  Call .WriteProperty("Transparent", m_Transparent)
  Call .WriteProperty("BorderWidth", m_BorderWidth)
  Call .WriteProperty("CornerCurve", m_CornerCurve)
  Call .WriteProperty("ForeColor1", m_ForeColor1)
  Call .WriteProperty("ForeColor2", m_ForeColor2)
  Call .WriteProperty("Caption1", m_Caption1)
  Call .WriteProperty("Caption2", m_Caption2)
  Call .WriteProperty("CAptionAlign", m_CaptionAlign)
  Call .WriteProperty("Font1", m_Font1)
  Call .WriteProperty("Font2", m_Font2)
  Call .WriteProperty("Opacity", m_Opacity)
  Call .WriteProperty("CallOutLen", m_CallOutLen)
  Call .WriteProperty("CallOutWidth", m_CallOutWidth)
  Call .WriteProperty("CallOutPos", m_CallOutPos, 0)
End With
  
End Sub

Public Property Get BackColor() As OLE_COLOR
  BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
  m_BackColor = New_Color
  PropertyChanged "BackColor"
  Refresh
End Property

Public Property Get BackColorParent() As OLE_COLOR
  BackColorParent = m_BackColorParent
End Property

Public Property Let BackColorParent(ByVal New_Color As OLE_COLOR)
  m_BackColorParent = New_Color
  PropertyChanged "BackColorParent"
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

Public Property Get BorderWidth() As Long
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  Refresh
End Property

Public Property Get CallOutLen() As Long
  CallOutLen = m_CallOutLen
End Property

Public Property Let CallOutLen(ByVal NewCallOutLen As Long)
  m_CallOutLen = NewCallOutLen
  PropertyChanged "CallOutLen"
  Refresh
End Property

Public Property Get CallOutPos() As CallOutPosition
  CallOutPos = m_CallOutPos
End Property

Public Property Let CallOutPos(ByVal NewCallOutPos As CallOutPosition)
  m_CallOutPos = NewCallOutPos
  PropertyChanged "CallOutPos"
  Refresh
End Property

Public Property Get CallOutWidth() As Long
  CallOutWidth = m_CallOutWidth
End Property

Public Property Let CallOutWidth(ByVal NewCallOutWidth As Long)
  m_CallOutWidth = NewCallOutWidth
  PropertyChanged "CallOutWidth"
  Refresh
End Property

Public Property Get Caption1() As String
  Caption1 = m_Caption1
End Property

Public Property Let Caption1(ByVal NewCaption1 As String)
  m_Caption1 = NewCaption1
  PropertyChanged "Caption1"
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

Public Property Get CaptionAlign() As DTAlign
  CaptionAlign = m_CaptionAlign
End Property

Public Property Let CaptionAlign(ByVal NewCaptionAlign As DTAlign)
  m_CaptionAlign = NewCaptionAlign
  PropertyChanged "CaptionAlign"
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

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  UserControl.Enabled = m_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get Font1() As StdFont
  Set Font1 = m_Font1
End Property

Public Property Set Font1(ByVal New_Font As StdFont)
  Set m_Font1 = New_Font
  PropertyChanged "Font1"
  Refresh
End Property

Public Property Get Font2() As StdFont
  Set Font2 = m_Font2
End Property

Public Property Set Font2(ByVal New_Font As StdFont)
  Set m_Font2 = New_Font
  PropertyChanged "Font2"
  Refresh
End Property

Public Property Get ForeColor1() As OLE_COLOR
  ForeColor1 = m_ForeColor1
End Property

Public Property Let ForeColor1(ByVal NewForeColor1 As OLE_COLOR)
  m_ForeColor1 = NewForeColor1
  PropertyChanged "ForeColor1"
  Refresh
End Property

Public Property Get ForeColor2() As OLE_COLOR
  ForeColor2 = m_ForeColor2
End Property

Public Property Let ForeColor2(ByVal NewForeColor2 As OLE_COLOR)
  m_ForeColor2 = NewForeColor2
  PropertyChanged "ForeColor2"
  Refresh
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Opacity() As Long
  Opacity = m_Opacity
End Property

Public Property Let Opacity(ByVal NewOpacity As Long)
  m_Opacity = NewOpacity
  PropertyChanged "Opacity"
  Refresh
End Property

Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal NewValue As Boolean)
    m_Transparent = NewValue
    PropertyChanged "Transparent"
    Refresh
End Property

Public Property Get Version() As String
Version = VERS  'App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Property Get Visible() As Boolean
  Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal NewVisible As Boolean)
  Extender.Visible = NewVisible
End Property
