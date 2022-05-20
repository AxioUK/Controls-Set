VERSION 5.00
Begin VB.UserControl AxGSlider 
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
   ToolboxBitmap   =   "AxGSlider.ctx":0000
End
Attribute VB_Name = "AxGSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxGSlider
'Version  : 2.07.6
'Editor   : David Rojas [AxioUK]
'Date     : 19/05/2022
'------------------------------------
Option Explicit

Private Const VERS As String = "2.07.6"

'Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SetRECL Lib "user32" Alias "SetRect" (lpRect As RECTL, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal token As Long)
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
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
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal RGBA As Long, ByRef brush As Long) As Long
'Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
'---
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function TextOutW Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function DrawTextA Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long

Private Const LF_FACESIZE = 32
'Private Const SYSTEM_FONT = 13
Private Const OBJ_FONT As Long = 6&

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

Private Type POINTS
   X As Single
   Y As Single
End Type

Private Type POINTL
   X As Long
   Y As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type pPoints
  X1 As Long
  x2 As Long
  Valor As String
End Type

Public Enum pStyle
  pVertical
  pHorizontal
End Enum

'Public Enum CallOutPosition
'  coLeft
'  coTop
'  coRight
'  coBottom
'End Enum

Public Enum coMark
  cmNothing
  cmBar
End Enum

Public Enum eTypeValue
  eDateValue
  eNumValue
  eLetterValue
End Enum

Public Enum eDateValueI
   byDay
   byMonth
   byYear
End Enum

Public Enum eStyleLine
   stInner
   stOuter
End Enum

'Constants
'Private Const CombineModeExclude As Long = &H4
Private Const WrapModeTileFlipXY = &H3
'Private Const SmoothingModeHighQuality As Long = &H2
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const LOGPIXELSX As Long = 88
'Private Const LOGPIXELSY As Long = 90
'Private Const TLS_MINIMUM_AVAILABLE As Long = 64
'Private Const IDC_HAND As Long = 32649
Private Const UnitPixel As Long = &H2&
Private Const DT_CENTER = &H1

'Define EVENTS-------------------
Public Event Click()
'Public Event DblClick()
Public Event ChangeMarks(vMark As String)
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
Private m_BorderColor   As OLE_COLOR
Private m_ForeColor1    As OLE_COLOR
Private m_ForeColor2    As OLE_COLOR
Private m_GradientColor1  As OLE_COLOR
Private m_GradientColor2  As OLE_COLOR
Private m_ValuesLineColor As OLE_COLOR
Private m_ColorSelector As OLE_COLOR
Private m_Transparent   As Boolean
Private m_BorderWidth   As Long
Private m_CornerCurve   As Long
Private m_ValuesVisible As Boolean
Private m_ValueRotation As Single
Private m_BarThickness  As Long
Private m_BarMargin     As Long

Private m_ValueType As eTypeValue
Private m_DateValueIntervalBy As eDateValueI
Private m_ValueLine As eStyleLine

Private m_Font1 As StdFont
Private m_Font2 As StdFont
Private m_Style As pStyle
Private mActive As coMark
Private iPts()  As POINTS
Private pPts()  As pPoints
Private Bar     As RECTL

Private m_MarkValue As String

Private m_Min As String
Private m_Max As String
Private m_Interval As Long
Private stMark As String
Private xPos   As Long

Private Function AddString2(sCaption As String, X As Long, Y As Long, Rotation As Single, TextColor As OLE_COLOR) As Boolean
Dim LF As LOGFONT
Dim FontToUse As Long
Dim oldFont As Long
'Dim RCT As RECT

With UserControl
  oldFont = GetCurrentObject(.hDC, OBJ_FONT)
  GetObjectAPI oldFont, Len(LF), LF
  GetObjectAPI oldFont, Len(LF), LF
  
  LF.lfEscapement = Rotation * 10
  LF.lfOrientation = LF.lfEscapement

  FontToUse = CreateFontIndirect(LF)
  oldFont = SelectObject(.hDC, FontToUse)
    
  .ForeColor = TextColor
  TextOutW .hDC, X, Y, StrPtr(sCaption), Len(sCaption)
  SelectObject .hDC, oldFont
  DeleteObject FontToUse
End With

ErrorTextRotate:
End Function

Private Function AddString(mCaption As String, Rct As Rect, mFont As Font, TextColor As OLE_COLOR)
Dim pFont       As IFont
Dim lFontOld    As Long
    
On Error GoTo ErrF
With UserControl
  Set pFont = mFont
  lFontOld = SelectObject(.hDC, pFont.hFont)
  
  .ForeColor = TextColor
    
  DrawTextA .hDC, mCaption, -1, Rct, DT_CENTER
  
  Call SelectObject(.hDC, lFontOld)
  
ErrF:
  Set pFont = Nothing
End With
End Function

'Public Sub CopyAmbient()
'Dim OPic As StdPicture
'
'On Error GoTo Err
'
'    With UserControl
'        Set .Picture = Nothing
'        Set OPic = Extender.Container.Image
'        .BackColor = Extender.Container.BackColor
'        UserControl.PaintPicture OPic, 0, 0, , , Extender.Left, Extender.Top
'        Set .Picture = .Image
'    End With
'Err:
'End Sub

Public Sub Refresh()
'If m_Transparent Then CopyAmbient
Draw
End Sub

Private Sub Draw()
Dim I As Long, tPoints As Long
Dim REC As RECTL
Dim cpBar As Rect

On Error GoTo ErrDraw
With UserControl
  .Cls
  GdipCreateFromHDC .hDC, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
  
  'Control BackColor
  .BackColor = m_BackColor
  
  'draw Background Bar
  SetRECL REC, m_BarMargin, 5, (.ScaleWidth - (m_BarMargin * 2)), m_BarThickness ' (.ScaleHeight / 3)
  DrawRoundRect hGraphics, REC, RGBA(m_GradientColor1, 100), RGBA(m_GradientColor2, 100), RGBA(m_BorderColor, 100), m_BorderWidth, m_CornerCurve
  'Get Points Values
  tPoints = fSetPoints(REC, m_Min, m_Max)

  If m_ValueLine = stOuter Then
    SetRECL REC, m_BarMargin, 5 + m_BarThickness, (.ScaleWidth - (m_BarMargin * 2)), m_BarThickness
    DrawPoints REC, m_ValuesLineColor, m_ForeColor1, tPoints, pHorizontal
  Else
    SetRECL REC, m_BarMargin, 5, (.ScaleWidth - (m_BarMargin * 2)), m_BarThickness
    DrawPoints REC, m_ValuesLineColor, m_ForeColor1, tPoints, pHorizontal
  End If
  
'*1  'Draw Slider
  Dim bW As Long
  bW = .TextWidth((pPts(UBound(pPts)).Valor) & "10")
  
  For I = 0 To UBound(pPts)
    If xPos >= pPts(I).X1 And xPos <= pPts(I).x2 Then stMark = pPts(I).Valor
  Next I
  
  SetRECL Bar, (xPos - (bW / 2)), 5, bW, m_BarThickness + 5
  DrawBubble hGraphics, Bar, RGBA(m_BorderColor, 100), 1, RGBA(m_ColorSelector, 100), 3, 5, 5, coBottom

  'Draw Marks Captions
  SetREC2 cpBar, (Bar.Left), 7, (Bar.Left + bW), m_BarThickness + 5
  
  AddString stMark, cpBar, m_Font2, m_ForeColor2
  
  RaiseEvent ChangeMarks(stMark)

 GdipDeleteGraphics hGraphics

  ''---------------
  If m_Transparent Then
    .BackStyle = 0
    .MaskColor = .BackColor
    Set .MaskPicture = .Image
  Else
    .BackStyle = 1
    UserControl.Refresh
  End If

End With

ErrDraw:
End Sub

Private Sub DrawPoints(Rct As RECTL, ColorLine As OLE_COLOR, ColorText As OLE_COLOR, _
                      vPoints As Long, lStyle As pStyle)
Dim I As Integer
Dim X As Single, Y As Single
Dim W As Single, H As Single
Dim sREC As POINTL
Dim tW As Single, stMark As String
Dim pY2 As Single
Dim pSpace As Single  ', vMark As Long
Dim sPar As Boolean, Th As Single

X = Rct.Left:  Y = Rct.Top
W = Rct.Width: H = Rct.Height

sPar = False
pSpace = (W / vPoints) * nScale

ReDim iPts(vPoints) As POINTS
          
For I = 0 To vPoints Step m_Interval

  tW = UserControl.TextWidth(pPts(I).Valor) * nScale
  Th = UserControl.TextHeight(pPts(I).Valor) * nScale

  iPts(I).X = X + (pSpace * I) * nScale
  iPts(I).Y = IIf(m_ValueLine = stOuter, Y + 1, Y + (H / 3)) * nScale
  
  If m_ValueLine = stOuter Then
    If tW * (vPoints / Interval) > (W - 100) Then
      pY2 = IIf(sPar = True, Y + Th, Y + 5) * nScale
    Else
      pY2 = (Y + 5) * nScale
    End If
  Else
      pY2 = (iPts(I).Y + (H / 3) + 2) * nScale
  End If
  
  If m_ValueLine = stOuter Then
    If m_ValuesVisible = True Then UserControl.Line (iPts(I).X, iPts(I).Y)-(iPts(I).X, pY2), ColorLine
  Else
    If I <> 0 Then
      If iPts(I).X <> (X + W) Then
        UserControl.Line (iPts(I).X, iPts(I).Y)-(iPts(I).X, pY2), ColorLine
      End If
    End If
  End If
  
  If m_ValuesVisible = True Then
    sREC.X = (iPts(I).X - (tW / 2))
    If m_ValueRotation < 340 Then sREC.X = iPts(I).X + (Th / 2)
    If m_ValueLine = stOuter Then
      If tW * (vPoints / Interval) > (W - 100) Then
        sREC.Y = IIf(sPar = True, iPts(I).Y + Th, iPts(I).Y + 5)
      Else
        sREC.Y = (iPts(I).Y + 5)
      End If
    Else
      sREC.Y = (Y + H) * nScale
    End If
      
    Select Case m_ValueType
      Case Is = eLetterValue
          stMark = Chr$(Asc(m_Min) + (I))
      Case Is = eNumValue
          stMark = CLng(m_Min) + (I)
      Case Is = eDateValue
        If m_DateValueIntervalBy = byDay Then
            stMark = DateAdd("d", CDbl(I), CDate(m_Min))
        ElseIf m_DateValueIntervalBy = byMonth Then
            stMark = Format$(DateAdd("m", CDbl(I), CDate(m_Min)), "MMMM-yyyy")
        ElseIf m_DateValueIntervalBy = byYear Then
            stMark = Year(DateAdd("yyyy", CDbl(I), CDate(m_Min)))
        End If
    End Select
    
    AddString2 stMark, sREC.X, sREC.Y, m_ValueRotation, ColorText
  End If
  sPar = Not sPar
Next I
  
zEnd:

End Sub

Private Function DrawRoundRect(ByVal hGraphics As Long, Rect As RECTL, ByVal color1 As Long, color2 As Long, _
                               ByVal BorderColor As Long, ByVal BorderWidth As Long, ByVal Round As Long) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mpath As Long
    Dim mRound As Long
    
    If m_BorderWidth <> 0 Then GdipCreatePen1 BorderColor, BorderWidth * nScale, &H2, hPen
    GdipCreateLineBrushFromRectWithAngleI Rect, color1, color2, 0, 0, WrapModeTileFlipXY, hBrush
    
    GdipCreatePath &H0, mpath   '&H0
    
    With Rect
        mRound = GetSafeRound((Round * nScale), .Width, .Height)
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
End Function

Private Function DrawBubble(ByVal hGraphics As Long, Rct As RECTL, BorderColor As Long, BorderWidth As Long, BackColor As Long, CornerCurve As Long, coWidth As Long, coLen As Long, COPos As CallOutPosition) As Long
    Dim mpath As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mRound As Long
    Dim Xx As Long, Yy As Long
    Dim lMax As Long
    Dim coAngle  As Long

With Rct
        
    coAngle = coWidth / 2

    mRound = GetSafeRound(CornerCurve * nScale, .Width, .Height)
    
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

    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, hPen
    GdipCreateSolidFill BackColor, hBrush
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

Private Function fSetPoints(Rct As RECTL, minVal As String, maxVal As String) As Long
Dim p As Integer, iSpace As Single, I As Integer

On Error GoTo ErrPoints

Select Case m_ValueType
  Case Is = eNumValue
      p = CLng(maxVal) - CLng(minVal)
  Case Is = eLetterValue
      p = Asc(maxVal) - Asc(minVal)
  Case Is = eDateValue
      If m_DateValueIntervalBy = byDay Then
          p = DateDiff("d", CDate(minVal), CDate(maxVal))
      ElseIf m_DateValueIntervalBy = byMonth Then
          p = DateDiff("m", CDate(minVal), CDate(maxVal))
      ElseIf m_DateValueIntervalBy = byYear Then
          p = DateDiff("yyyy", CDate(minVal), CDate(maxVal))
      End If
End Select

iSpace = (Rct.Width / p) * nScale

ReDim pPts(p) As pPoints

fSetPoints = p

For I = 0 To UBound(pPts)

  pPts(I).X1 = Rct.Left + (iSpace * I)
  pPts(I).x2 = pPts(I).X1 + iSpace
  
  Select Case m_ValueType
    Case Is = eLetterValue
          pPts(I).Valor = Chr$(Asc(minVal) + I)
          
    Case Is = eNumValue
          pPts(I).Valor = CLng(minVal) + I
          
    Case Is = eDateValue
      If m_DateValueIntervalBy = byDay Then
          pPts(I).Valor = DateAdd("d", CDbl(I), CDate(minVal))
      ElseIf m_DateValueIntervalBy = byMonth Then
          pPts(I).Valor = Format$(DateAdd("m", CDbl(I), CDate(minVal)), "MM-yyyy")
      ElseIf m_DateValueIntervalBy = byYear Then
          pPts(I).Valor = Format$(DateAdd("yyyy", CDbl(I), CDate(minVal)), "yyyy")
      End If
      
  End Select
  
Next I

Exit Function
ErrPoints:
  Debug.Print "Error setting points"
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

Private Sub UserControl_AmbientChanged(PropertyName As String)
  Refresh
End Sub

Private Sub UserControl_Initialize()
InitGDI
nScale = GetWindowsDPI

End Sub

Private Sub UserControl_InitProperties()
'hFontCollection = ReadValue(&HFC)

  m_Enabled = True
  'm_Style = pHorizontal
  m_BorderColor = &HFF8080
  m_BackColor = &H8000000F
  m_GradientColor1 = &HFF&
  m_GradientColor2 = &HC000&
  'm_Transparent = False
  m_BorderWidth = 1
  m_CornerCurve = 10
  m_ForeColor1 = &H8D4214
  m_ForeColor2 = &HFFFFFF
  Set m_Font1 = UserControl.Font
  Set m_Font2 = UserControl.Font
  m_ValuesLineColor = &H8D4214
  m_Min = "0"
  m_Max = "100"
  m_Interval = 10
  m_ValueType = eNumValue
  m_DateValueIntervalBy = byDay
  m_ColorSelector = &H0&
  m_ValueLine = stOuter
  m_ValuesVisible = True
  m_ValueRotation = 0
  m_BarThickness = 25
  m_BarMargin = 10
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

If Button = vbLeftButton Then
  If X > Bar.Left And X < (Bar.Left + Bar.Width) And Y > Bar.Top And Y < (Bar.Top + Bar.Height) Then
    mActive = cmBar
  End If
End If

RaiseEvent MouseDown(Button, Shift, X, Y)
RaiseEvent Click
End Sub

'*3
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If X > (Bar.Left) And X < (Bar.Left + Bar.Width) And Y > Bar.Top And Y < (Bar.Top + Bar.Height) Then
  UserControl.MousePointer = vbSizeWE
Else
  UserControl.MousePointer = vbDefault
End If

If Button = vbLeftButton Then
  If mActive = cmBar Then
      If X >= m_BarMargin And X <= (UserControl.ScaleWidth - m_BarMargin) Then xPos = X
      UserControl.MousePointer = vbSizeWE
  End If
  Refresh
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mActive = cmNothing
UserControl.MousePointer = vbDefault
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  m_Enabled = .ReadProperty("Enabled", True)
  m_Transparent = .ReadProperty("Transparent", True)
  'm_Style = .ReadProperty("Style", 0)
  m_BorderColor = .ReadProperty("BorderColor", &HFF8080)
  m_BackColor = .ReadProperty("BackColor", &H8000000F)
  m_GradientColor1 = .ReadProperty("GradientColor1", &HFF&)
  m_GradientColor2 = .ReadProperty("GradientColor2", &HC000&)
  m_BorderWidth = .ReadProperty("BorderWidth", 1)
  m_CornerCurve = .ReadProperty("CornerCurve", 10)
  m_ForeColor1 = .ReadProperty("ValuesForeColor", &H8D4214)
  m_ForeColor2 = .ReadProperty("MarksForeColor", &HFFFFFF)
  Set m_Font1 = .ReadProperty("ValuesFont", UserControl.Font)
  Set m_Font2 = .ReadProperty("MarksFont", UserControl.Font)
  m_ValuesLineColor = .ReadProperty("ValuesLineColor", &H8D4214)
  m_Min = .ReadProperty("Min", "0")
  m_Max = .ReadProperty("Max", "100")
  m_Interval = .ReadProperty("Interval", 10)
  m_ValueType = .ReadProperty("ValueType", eNumValue)
  m_DateValueIntervalBy = .ReadProperty("DateValueIntervalBy", 0)
  m_ColorSelector = .ReadProperty("ColorSelector", &H0&)
  m_ValueLine = .ReadProperty("ValueLine", 0)
  m_ValuesVisible = .ReadProperty("ValuesVisible", True)
  m_ValueRotation = .ReadProperty("ValueRotation", 0)
  m_BarThickness = .ReadProperty("BarThickness", 25)
  m_BarMargin = .ReadProperty("BarMargin", 10)
End With

End Sub

Private Sub UserControl_Resize()
  xPos = m_BarMargin
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
  Call .WriteProperty("Transparent", m_Transparent)
  'Call .WriteProperty("Style", m_Style)
  Call .WriteProperty("BorderColor", m_BorderColor)
  Call .WriteProperty("BackColor", m_BackColor)
  Call .WriteProperty("GradientColor1", m_GradientColor1)
  Call .WriteProperty("GradientColor2", m_GradientColor2)
  Call .WriteProperty("BorderWidth", m_BorderWidth)
  Call .WriteProperty("CornerCurve", m_CornerCurve)
  Call .WriteProperty("ValuesForeColor", m_ForeColor1)
  Call .WriteProperty("MarksForeColor", m_ForeColor2)
  Call .WriteProperty("ValuesFont", m_Font1)
  Call .WriteProperty("MarksFont", m_Font2)
  Call .WriteProperty("ValuesLineColor", m_ValuesLineColor)
  Call .WriteProperty("Min", m_Min)
  Call .WriteProperty("Max", m_Max)
  Call .WriteProperty("Interval", m_Interval, 10)
  Call .WriteProperty("ValueType", m_ValueType)
  Call .WriteProperty("DateValueIntervalBy", m_DateValueIntervalBy)
  Call .WriteProperty("ColorSelector", m_ColorSelector)
  Call .WriteProperty("ValueLine", m_ValueLine)
  Call .WriteProperty("ValuesVisible", m_ValuesVisible)
  Call .WriteProperty("ValueRotation", m_ValueRotation)
  Call .WriteProperty("BarThickness", m_BarThickness)
  Call .WriteProperty("BarMargin", m_BarMargin)
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

Public Property Get GradientColor1() As OLE_COLOR
  GradientColor1 = m_GradientColor1
End Property

Public Property Let GradientColor1(ByVal New_Color As OLE_COLOR)
  m_GradientColor1 = New_Color
  PropertyChanged "GradientColor1"
  Refresh
End Property

Public Property Get GradientColor2() As OLE_COLOR
  GradientColor2 = m_GradientColor2
End Property

Public Property Let GradientColor2(ByVal New_Color As OLE_COLOR)
  m_GradientColor2 = New_Color
  PropertyChanged "GradientColor2"
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

Public Property Get ValuesForeColor() As OLE_COLOR
  ValuesForeColor = m_ForeColor1
End Property

Public Property Let ValuesForeColor(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor1 = NewForeColor
  PropertyChanged "ValuesForeColor"
  Refresh
End Property

Public Property Get ValuesFont() As StdFont
  Set ValuesFont = m_Font1
End Property

Public Property Set ValuesFont(ByVal New_Font As StdFont)
  Set m_Font1 = New_Font
  PropertyChanged "ValuesFont"
  Refresh
End Property

Public Property Get MarksForeColor() As OLE_COLOR
  MarksForeColor = m_ForeColor2
End Property

Public Property Let MarksForeColor(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor2 = NewForeColor
  PropertyChanged "MarksForeColor"
  Refresh
End Property

Public Property Get MarksFont() As StdFont
  Set MarksFont = m_Font2
End Property

Public Property Set MarksFont(ByVal New_Font As StdFont)
  Set m_Font2 = New_Font
  PropertyChanged "MarksFont"
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

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get ValuesLineColor() As OLE_COLOR
  ValuesLineColor = m_ValuesLineColor
End Property

Public Property Let ValuesLineColor(ByVal NewValuesLineColor As OLE_COLOR)
  m_ValuesLineColor = NewValuesLineColor
  PropertyChanged "ValuesLineColor"
  Refresh
End Property

'Public Property Get Style() As pStyle
'  Style = m_Style
'End Property
'
'Public Property Let Style(ByVal NewStyle As pStyle)
'  m_Style = NewStyle
'  PropertyChanged "Style"
'  UserControl_Resize
'End Property

Public Property Get Version() As String
Version = VERS  'App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Property Get Visible() As Boolean
  Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal NewVisible As Boolean)
  Extender.Visible = NewVisible
End Property

Public Property Get Min() As String
  Min = m_Min
End Property

Public Property Let Min(ByVal NewMin As String)
  m_Min = NewMin
  PropertyChanged "Min"
  Refresh
End Property

Public Property Get Max() As String
  Max = m_Max
End Property

Public Property Let Max(ByVal NewMax As String)
  m_Max = NewMax
  PropertyChanged "Max"
  Refresh
End Property

Public Property Get Interval() As Long
  Interval = m_Interval
End Property

Public Property Let Interval(ByVal NewInterval As Long)
  m_Interval = NewInterval
  If m_Interval = 0 Then m_Interval = 1
  PropertyChanged "Interval"
  Refresh
End Property

Public Property Get ValueType() As eTypeValue
  ValueType = m_ValueType
End Property

Public Property Let ValueType(ByVal NewValueType As eTypeValue)
  m_ValueType = NewValueType
  PropertyChanged "ValueType"
  Select Case m_ValueType
    Case Is = eNumValue:    Min = 0:    Max = 100
    Case Is = eLetterValue: Min = "A":  Max = "Z"
    Case Is = eDateValue:   Min = "01-01-2021": Max = "31-01-2021"
  End Select
  Refresh
End Property

Public Property Get DateValueIntervalBy() As eDateValueI
  DateValueIntervalBy = m_DateValueIntervalBy
End Property

Public Property Let DateValueIntervalBy(ByVal NewDateValueIntervalBy As eDateValueI)
  m_DateValueIntervalBy = NewDateValueIntervalBy
  PropertyChanged "DateValueIntervalBy"
  Refresh
End Property

Public Property Get MarkValue() As String
  MarkValue = stMark
End Property

Public Function SetMarkValue(ByVal NewMarkValue As String)
Dim I As Integer
  
  m_MarkValue = NewMarkValue

  For I = 0 To UBound(pPts)
    If pPts(I).Valor = m_MarkValue Then
      xPos = pPts(I).X1
      mActive = cmBar
      Draw
    End If
  Next I
  
End Function

Public Property Get ColorSelector() As OLE_COLOR
  ColorSelector = m_ColorSelector
End Property

Public Property Let ColorSelector(ByVal NewColorSelector As OLE_COLOR)
  m_ColorSelector = NewColorSelector
  PropertyChanged "ColorSelector"
  Refresh
End Property

Public Property Get ValueLine() As eStyleLine
  ValueLine = m_ValueLine
End Property

Public Property Let ValueLine(ByVal NewValueLine As eStyleLine)
  m_ValueLine = NewValueLine
  PropertyChanged "ValueLine"
  Refresh
End Property

Public Property Get ValuesVisible() As Boolean
  ValuesVisible = m_ValuesVisible
End Property

Public Property Let ValuesVisible(ByVal NewValuesVisible As Boolean)
  m_ValuesVisible = NewValuesVisible
  PropertyChanged "ValuesVisible"
  Refresh
End Property

Public Property Get ValueRotation() As Single
  ValueRotation = 360 - m_ValueRotation
End Property

Public Property Let ValueRotation(ByVal NewValueRotation As Single)
  m_ValueRotation = 360 - NewValueRotation
  PropertyChanged "ValueRotation"
  Refresh
End Property

Public Property Get BarThickness() As Long
  BarThickness = m_BarThickness
End Property

Public Property Let BarThickness(ByVal NewBarThickness As Long)
  m_BarThickness = NewBarThickness
  PropertyChanged "BarThickness"
  Refresh
End Property

Public Property Get BarMargin() As Long
  BarMargin = m_BarMargin
End Property

Public Property Let BarMargin(ByVal NewBarMargin As Long)
  m_BarMargin = NewBarMargin
  PropertyChanged "BarMargin"
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

