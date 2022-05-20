VERSION 5.00
Begin VB.UserControl AxGText 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "AxGText.ctx":0000
   ScaleHeight     =   36
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "AxGText.ctx":000F
   Begin VB.TextBox mEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1950
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmrOver 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Edit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "AxGText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxGTtext
'Version  : 2.07.6
'Editor   : David Rojas [AxioUK]
'Date     : 19/05/2022
'------------------------------------

Option Explicit

Private Const VERS As String = "2.07.6"

'/DPI
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'/GDI+
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As Long, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef brush As Long) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipSetPenMode Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mPenMode As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mRect As Rect, ByVal mFlags As Long, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixelOffsetMode As Long) As Long

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function SetRect Lib "user32" (lpRect As Any, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long

Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)

Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFileName As String, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'+-<-----------------------------------
Private Type Rect
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
End Type

Private Type RECTF
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type BitmapData
    Width   As Long
    Height  As Long
    stride  As Long
    PixelFormat As Long
    Scan0Ptr    As Long
    ReservedPtr As Long
End Type

Private Type GUID
    Data1       As Long
    Data2       As Integer
    Data3       As Integer
    Data4(0 To 7) As Byte
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

'''-----------------------------------------------------------
'MousePointerHand
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Type PicBmp
  Size As Long
  type As Long
  hBmp As Long
  hPal As Long
  Reserved As Long
End Type

'Constantes
Private Const IDC_HAND As Long = 32649

'Variables
Dim hCur      As Long

'''-----------------------------------------------------------

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

'New
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef Image As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long

Private Const PixelFormat32bppPARGB As Long = &HE200B
Private Const PASSWORD_CHAR As String = "•"

Private Const EM_SETRECT = &HB3
Private Const EM_GETRECT = &HB2
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINELENGTH As Long = &HC1


Public Enum jImageAlignmentH
    ImageInLeft
    ImageInRight
End Enum

Public Enum jImageAlignmentV
    ImgTop
    ImgMiddle
    ImgBottom
End Enum

Public Enum jImagePosition
    ImageInBox
    ImageOutBox
End Enum

'ENUMS----------------------
Public Enum RegionalConstant
  LOCALE_SCURRENCY = &H14
  LOCALE_SCOUNTRY = &H6
  LOCALE_SDATE = &H1D
  LOCALE_SDECIMAL = &HE
  LOCALE_SLANGUAGE = &H2
  LOCALE_SLONGDATE = &H20
  LOCALE_SMONDECIMALSEP = &H16
  LOCALE_SMONGROUPING = &H18
  LOCALE_SMONTHOUSANDSEP = &H17
  LOCALE_SNATIVECTRYNAME = &H8
  LOCALE_SNATIVECURRNAME = &H1008
  LOCALE_SNATIVEDIGITS = &H13
  LOCALE_SNEGATIVESIGN = &H51
  LOCALE_SSHORTDATE = &H1F
  LOCALE_STIME = &H1E
  LOCALE_STIMEFORMAT = &H1003
End Enum

Public Enum CharacterType
    AllChars
    LettersOnly
    NumbersOnly
    LettersAndNumbers
    Money
    Percent
    Fraction
    Decimals
    Dates
    ChileanRUT
    IPAddress
End Enum

Public Enum CaseType
    Normal
    UpperCase
    LowerCase
End Enum

Public Enum eAlignConst
   [Left Justify] = 0
   [Right Justify] = 1
   Center = 2
End Enum

Public Enum eEnterKeyBehavior
    [eNone] = 0
    [eKeyTab] = 1
    [Validate] = 2
End Enum

'EVENTOS---------------------------------------------------
Public Event Change()
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event EnterKeyPress()
Public Event IconClick()
Public Event MouseEnter()
Public Event MouseLeave()

'VARIABLES------------------------------------------------
'You have to have MSScripting Runtime referenced : WshShell.SendKeys "{Tab}"
Dim WshShell   As Object
Dim FlechasTab As Boolean     'Usar las Flechas del cursor como Tabulador
Dim EnterTab   As Boolean     'Usar Enter como Tabulador

Private m_bvData()  As Byte

Private mEnabled        As Boolean
Private m_token         As Long
Private m_BmpS          As Long
Private m_Bmp           As Long
Private m_BmpRct        As Rect
Private m_BmpSrcW       As Single
Private m_BmpSrcH       As Single
Private m_Image         As Byte

Private m_BorderWidth       As Single
Private m_BackColor         As OLE_COLOR
Private m_BackColorFocus    As OLE_COLOR
Private m_BorderColor       As OLE_COLOR
Private m_BorderColorFocus  As OLE_COLOR
Private m_ShadowColor       As OLE_COLOR
Private m_ShadowColorFocus  As OLE_COLOR
Private m_AutoSel           As Boolean
Private m_Pwd               As Boolean
Private m_Round             As Long
Private m_Shadow            As Long
Private m_BmpAlignH         As jImageAlignmentH
Private m_BmpAlignV         As jImageAlignmentV
Private m_BmpSize           As String
Private m_ImagePosition     As jImagePosition
Private shwRect             As Rect
Private m_MultiLine         As Boolean
Private m_Alignment         As AlignmentConstants
Private m_Text              As String
Private m_Locked            As Boolean
Private m_ForeColor         As OLE_COLOR
Private m_Opacity           As Long

Private m_FormatToString As CharacterType
Private m_CaseText       As CaseType
Private m_KeyBehavior    As eEnterKeyBehavior
Private sDecimal         As String
Private m_ParteDecimal   As Long    'Yaco
Private sThousand        As String
Private sDateDiv         As String
Private sMoney           As String
Private iCount           As Integer
Private bCancel          As Boolean
Private txtBaseString    As String
Private SetSize          As Boolean
Private m_UseCue         As Boolean
Private m_Transp         As Boolean

Private bMouseOver   As Boolean
Private bInFocus     As Boolean
Private dpiScale     As Double
Private Th           As Long
Private Px           As Long
Private bS           As Long

Dim lW      As Long
Dim lh      As Long
Dim lX      As Long
Dim lY      As Long



Public Function fCleanValue(sValor As String) As String
Dim sValue As Variant, I As Integer

   For I = 1 To Len(sValor)
      If IsNumeric(Mid(sValor, I, 1)) Or Mid(sValor, I, 1) = sDecimal Then
         sValue = sValue & Mid(sValor, I, 1)
      End If
   Next I

fCleanValue = Trim(sValue)
End Function

Private Function fGetLocaleInfo(Valor As RegionalConstant) As String
   Dim Simbolo As String
   Dim r1 As Long
   Dim r2 As Long
   Dim p As Integer
   Dim Locale As Long
     
   Locale = GetUserDefaultLCID()
   r1 = GetLocaleInfo(Locale, Valor, vbNullString, 0)
   'buffer
   Simbolo = String$(r1, 0)
   'En esta llamada devuelve el símbolo en el Buffer
   r2 = GetLocaleInfo(Locale, Valor, Simbolo, r1)
   'Localiza el espacio nulo de la cadena para eliminarla
   p = InStr(Simbolo, Chr$(0))
     
   If p > 0 Then
      'Elimina los nulos
      fGetLocaleInfo = Left$(Simbolo, p - 1)
   End If
     
End Function

Private Function InsertStr(ByVal InsertTo As String, ByVal Str As String, ByVal Position As Integer) As String
    Dim Str1 As String
    Dim Str2 As String
    
    Str1 = Mid$(InsertTo, 1, Position - 1)
    Str2 = Mid$(InsertTo, Position, Len(InsertTo) - Len(Str1))
    
    InsertStr = Str1 & Str & Str2
End Function

Public Function GetWindowsDPI() As Double
Dim hDC As Long, lPx  As Double
Const LOGPIXELSX    As Long = 88
    hDC = GetDC(0)
    lPx = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    ReleaseDC 0, hDC
    
    If (lPx = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = lPx / 96#
    End If
End Function

Public Sub ppCopyAmbient()
On Error GoTo e
Dim OPic As StdPicture
    With UserControl
        Set .Picture = Nothing
        Set OPic = Extender.Container.Image
        .BackColor = Extender.Container.BackColor
        UserControl.PaintPicture OPic, 0, 0, , , Extender.Left, Extender.Top ', Extender.Width * 15, Extender.Height * 15
        Set .Picture = .Image
    End With
    Exit Sub
e:
End Sub

Public Function LoadPictureFromFile(ByVal FileName As String) As Boolean
Dim imgW As Long, imgH As Long
Dim BmpW As Single, BmpH As Single
Dim Bmp  As Long, Grph As Long
    
If m_Bmp Then
    Call GdipDisposeImage(m_Bmp)
    m_Bmp = 0
End If

imgW = m_BmpRct.Width
imgH = m_BmpRct.Height

Call GdipLoadImageFromFile(StrPtr(FileName), m_Bmp)
If m_Bmp <> 0 Then
    GdipGetImageDimension m_Bmp, BmpW, BmpH
    '-->
    If GdipCreateBitmapFromScan0(imgW, imgH, 0&, &HE200B, ByVal 0&, Bmp) = 0 Then
      If GdipGetImageGraphicsContext(Bmp, Grph) = 0 Then
      
          If imgW > BmpW Or imgH > BmpH Then
              Call GdipSetInterpolationMode(Grph, 5&)  '// IterpolationModeNearestNeighbor
          Else
              Call GdipSetInterpolationMode(Grph, 7&)  '//InterpolationModeHighQualityBicubic
              Call GdipSetPixelOffsetMode(Grph, 4&)
          End If
          
          Call GdipDrawImageRectRectI(Grph, m_Bmp, 0, 0, imgW, imgH, 0, 0, BmpW, BmpH, &H2)
          GdipDeleteGraphics Grph
          
          Call GdipDisposeImage(m_Bmp)
          m_Bmp = Bmp
          '...Draw Image on Control
          ppDraw
          LoadPictureFromFile = True
      End If
    End If
Else
    LoadPictureFromFile = False
End If
End Function

Private Function ppCreateBitmapFromStream(bvData() As Byte, lBitmap As Long) As Long
On Error GoTo e
Dim IStream     As IUnknown
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then Call GdipLoadImageFromStream(IStream, lBitmap)
e:
    Set IStream = Nothing
End Function

Friend Function ppgGetStream() As Byte()
 ppgGetStream = m_bvData
 End Function
 
Friend Function ppgSetStream(bvData() As Byte)
    m_bvData = bvData
    Call SetPictureStream(bvData)
    PropertyChanged "ImageStream"
End Function

Public Function SetPictureStream(bvData() As Byte)
On Error GoTo e
Dim hGrph   As Long
Dim hBmp    As Long
Dim picW      As Long
Dim picH      As Long

    If m_Bmp Then
        GdipDisposeImage m_Bmp
        m_Bmp = 0
        m_BmpSrcW = 0
        m_BmpSrcH = 0
    End If
    
    ppCreateBitmapFromStream bvData, hBmp
    If hBmp = 0 Then GoTo e
    GdipGetImageDimension hBmp, m_BmpSrcW, m_BmpSrcH
    m_Bmp = hBmp
    
    picW = Split(m_BmpSize, "x")(0)
    picH = Split(m_BmpSize, "x")(1)
    m_BmpRct.Width = IIf(picW > 0, picW * dpiScale, m_BmpSrcW)
    m_BmpRct.Height = IIf(picW > 0, picW * dpiScale, m_BmpSrcH)
    
e:
    Call UserControl_Resize
End Function

Private Function ConvertColor(ByVal Color As Long, ByVal Opacity As Long) As Long
Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function

Private Function ConvertOLE(ByVal Color As Long) As Long
Dim BGRA(0 To 2) As Byte
  
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    
    ConvertOLE = RGB(BGRA(2), BGRA(1), BGRA(0))
End Function

Private Function RGBtoARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long 'By LaVople
    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
End Function

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
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

Private Function CreateBlurShadow(ByVal hImage As Long, ByVal Color As Long, blurDepth As Long, _
                                        Optional ByVal Left As Long, Optional ByVal Top As Long, _
                                        Optional ByVal Width As Long, Optional ByVal Height As Long) As Long
    Dim REC As Rect
    Dim X As Long, Y As Long
    Dim hImgShadow As Long
    Dim bmpData1 As BitmapData
    Dim bmpData2 As BitmapData
    Dim t2xBlur As Long
    Dim R As Long, G As Long, B As Long
    Dim dBytes() As Byte
    Dim srcBytes() As Byte
    Dim vTally() As Long
    Dim tAlpha As Long, tColumn As Long, tAvg As Long
    Dim initY As Long, initYstop As Long, initYstart As Long
  
    
    If hImage = 0& Then Exit Function
 
    If Width = 0& Then Call GdipGetImageWidth(hImage, Width)
    If Height = 0& Then Call GdipGetImageHeight(hImage, Height)
 
    t2xBlur = blurDepth * 2
 
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    B = (Color \ &H10000) And &HFF
 
    SetRect REC, Left, Top, Width, Height
 
    ReDim srcBytes(REC.Width * 4 - 1&, REC.Height - 1&)
  
    With bmpData1
        .Scan0Ptr = VarPtr(srcBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
   
    'Call GdipBitmapLockBits(hImage, REC, ImageLockModeUserInputBuf Or ImageLockModeRead, PixelFormat32bppPARGB, bmpData1)
    Call GdipBitmapLockBits(hImage, REC, &H4 Or &H1, PixelFormat32bppPARGB, bmpData1)
 
    SetRect REC, Left, Top, Width + t2xBlur, Height + t2xBlur
    
    Call GdipCreateBitmapFromScan0(REC.Width, REC.Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImgShadow)

    ReDim dBytes(REC.Width * 4 - 1&, REC.Height - 1&)
    
    With bmpData2
        .Scan0Ptr = VarPtr(dBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
    
    'Call GdipBitmapLockBits(hImgShadow, REC, ImageLockModeUserInputBuf Or ImageLockModeRead Or ImageLockModeWrite, PixelFormat32bppPARGB, bmpData2)
    Call GdipBitmapLockBits(hImage, REC, &H4 Or &H1, PixelFormat32bppPARGB, bmpData2)
    
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    B = (Color \ &H10000) And &HFF
    
    tAvg = (t2xBlur + 1) * (t2xBlur + 1)    ' how many pixels are being blurred
    
    ReDim vTally(0 To t2xBlur)              ' number of blur columns per pixel
    
    For Y = 0 To Height + t2xBlur - 1     ' loop thru shadow dib
    
        FillMemory vTally(0), (t2xBlur + 1) * 4, 0  ' reset column totals
        
        If Y < t2xBlur Then         ' y does not exist in source
            initYstart = 0          ' use 1st row
        Else
            initYstart = Y - t2xBlur ' start n blur rows above y
        End If
        ' how may source rows can we use for blurring?
        If Y < Height Then initYstop = Y Else initYstop = Height - 1
        
        tAlpha = 0  ' reset alpha sum
        tColumn = 0    ' reset column counter
        
        ' the first n columns will all be zero
        ' only the far right blur column has values; tally them
        For initY = initYstart To initYstop
            tAlpha = tAlpha + srcBytes(3, initY)
        Next
        ' assign the right column value
        vTally(t2xBlur) = tAlpha
        
        For X = 3 To (Width - 2) * 4 - 1 Step 4
            ' loop thru each source pixel's alpha
            
            ' set shadow alpha using blur average
            dBytes(X, Y) = tAlpha \ tAvg
            ' and set shadow color
            Select Case dBytes(X, Y)
            Case 255
                dBytes(X - 1, Y) = R
                dBytes(X - 2, Y) = G
                dBytes(X - 3, Y) = B
            Case 0
            Case Else
                dBytes(X - 1, Y) = R * dBytes(X, Y) \ 255
                dBytes(X - 2, Y) = G * dBytes(X, Y) \ 255
                dBytes(X - 3, Y) = B * dBytes(X, Y) \ 255
            End Select
            ' remove the furthest left column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' count the next column of alphas
            vTally(tColumn) = 0&
            For initY = initYstart To initYstop
                vTally(tColumn) = vTally(tColumn) + srcBytes(X + 4, initY)
            Next
            ' add the new column's sum to the overall sum
            tAlpha = tAlpha + vTally(tColumn)
            ' set the next column to be recalculated
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
        
        ' now to finish blurring from right edge of source
        For X = X To (Width + t2xBlur - 1) * 4 - 1 Step 4
            dBytes(X, Y) = tAlpha \ tAvg
            Select Case dBytes(X, Y)
            Case 255
                dBytes(X - 1, Y) = R
                dBytes(X - 2, Y) = G
                dBytes(X - 3, Y) = B
            Case 0
            Case Else
                dBytes(X - 1, Y) = R * dBytes(X, Y) \ 255
                dBytes(X - 2, Y) = G * dBytes(X, Y) \ 255
                dBytes(X - 3, Y) = B * dBytes(X, Y) \ 255
            End Select
            ' remove this column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' set next column to be removed
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
    Next
 
    Call GdipBitmapUnlockBits(hImage, bmpData1)
    Call GdipBitmapUnlockBits(hImgShadow, bmpData2)
    
    CreateBlurShadow = hImgShadow
End Function


Private Function ppBlur(hImage As Long, Color As Long, blurDepth As Long, _
                                  Optional ByVal Left As Long, Optional ByVal Top As Long, _
                                  Optional ByVal Width As Long, Optional ByVal Height As Long) As Long
'On Error Resume Next
Dim REC As Rect
Dim X As Long, Y As Long
Dim hImgShadow As Long
Dim bmpData1 As BitmapData
Dim bmpData2 As BitmapData
Dim t2xBlur As Long
Dim R As Long, G As Long, B As Long
Dim Alpha As Byte
Dim lSrcAlpha As Long, lDestAlpha As Long
Dim dBytes() As Byte
Dim srcBytes() As Byte
Dim vTally() As Long
Dim tAlpha As Long, tColumn As Long, tAvg As Long
Dim initY As Long, initYstop As Long, initYstart As Long
Dim initX As Long, initXstop As Long
    
    If hImage = 0& Then Exit Function
 
    If Width = 0& Then Call GdipGetImageWidth(hImage, Width)
    If Height = 0& Then Call GdipGetImageHeight(hImage, Height)
 
    t2xBlur = blurDepth * 2
 
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    B = (Color \ &H10000) And &HFF
 
    SetRect REC, 0, 0, Width, Height
    'SetRect REC, Left, Top, Width, Height   '<---AxioUK TEST
    
    ReDim srcBytes(REC.Width * 4 - 1&, REC.Height - 1&)
    With bmpData1
        .Scan0Ptr = VarPtr(srcBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
   
    Call GdipBitmapLockBits(hImage, REC, &H4 Or &H1, PixelFormat32bppPARGB, bmpData1)
 
    SetRect REC, 0, 0, Width + t2xBlur, Height + t2xBlur
    Call GdipCreateBitmapFromScan0(REC.Width, REC.Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImgShadow)

    ReDim dBytes(REC.Width * 4 - 1&, REC.Height - 1&)
    
    With bmpData2
        .Scan0Ptr = VarPtr(dBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
    
    Call GdipBitmapLockBits(hImgShadow, REC, &H4 Or &H1 Or &H2, PixelFormat32bppPARGB, bmpData2)
 
    'SetRect REC, Left, Top, Width + t2xBlur, Height + t2xBlur '<---AxioUK TEST
    
    tAvg = (t2xBlur + 1) * (t2xBlur + 1)    ' how many pixels are being blurred
    
    ReDim vTally(0 To t2xBlur)              ' number of blur columns per pixel
    
    For Y = 0 To Height + t2xBlur - 1     ' loop thru shadow dib
    
        FillMemory vTally(0), (t2xBlur + 1) * 4, 0  ' reset column totals
        
        If Y < t2xBlur Then         ' y does not exist in source
            initYstart = 0          ' use 1st row
        Else
            initYstart = Y - t2xBlur ' start n blur rows above y
        End If
        ' how may source rows can we use for blurring?
        If Y < Height Then initYstop = Y Else initYstop = Height - 1
        
        tAlpha = 0  ' reset alpha sum
        tColumn = 0    ' reset column counter
        
        ' the first n columns will all be zero
        ' only the far right blur column has values; tally them
        For initY = initYstart To initYstop
            tAlpha = tAlpha + srcBytes(3, initY)
        Next
        ' assign the right column value
        vTally(t2xBlur) = tAlpha
        
        For X = 3 To (Width - 2) * 4 - 1 Step 4
            ' loop thru each source pixel's alpha
            ' set shadow alpha using blur average
            dBytes(X, Y) = tAlpha \ tAvg
            ' and set shadow color
            Select Case dBytes(X, Y)
            Case 255
                dBytes(X - 1, Y) = R
                dBytes(X - 2, Y) = G
                dBytes(X - 3, Y) = B
            Case 0
            Case Else
                dBytes(X - 1, Y) = R * dBytes(X, Y) \ 255
                dBytes(X - 2, Y) = G * dBytes(X, Y) \ 255
                dBytes(X - 3, Y) = B * dBytes(X, Y) \ 255
            End Select
            ' remove the furthest left column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' count the next column of alphas
            vTally(tColumn) = 0&
            For initY = initYstart To initYstop
                vTally(tColumn) = vTally(tColumn) + srcBytes(X + 4, initY)
            Next
            ' add the new column's sum to the overall sum
            tAlpha = tAlpha + vTally(tColumn)
            ' set the next column to be recalculated
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
        
        ' now to finish blurring from right edge of source
        For X = X To (Width + t2xBlur - 1) * 4 - 1 Step 4
            dBytes(X, Y) = tAlpha \ tAvg
            Select Case dBytes(X, Y)
            Case 255
                dBytes(X - 1, Y) = R
                dBytes(X - 2, Y) = G
                dBytes(X - 3, Y) = B
            Case 0
            Case Else
                dBytes(X - 1, Y) = R * dBytes(X, Y) \ 255
                dBytes(X - 2, Y) = G * dBytes(X, Y) \ 255
                dBytes(X - 3, Y) = B * dBytes(X, Y) \ 255
            End Select
            ' remove this column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' set next column to be removed
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
    Next
 
    Call GdipBitmapUnlockBits(hImage, bmpData1)
    Call GdipBitmapUnlockBits(hImgShadow, bmpData2)
    
    ppBlur = hImgShadow
End Function

Private Sub ppCreateShadow(oleShadowColor As OLE_COLOR)
Dim hGrph   As Long
Dim hBmp    As Long
Dim hPath   As Long
Dim hPen    As Long
Dim hBrush  As Long

Dim eW      As Long
Dim eH      As Long
Dim lpz     As Long

    If m_BmpS Then GdipDisposeImage m_BmpS: m_BmpS = 0
    If m_Shadow = 0 Then Exit Sub
    
    lpz = m_Shadow * dpiScale
    
    eW = shwRect.Width - (lpz * 2) 'UserControl.ScaleWidth - (lpz * 2)
    eH = UserControl.ScaleHeight - (lpz * 2)
    
    GdipCreateBitmapFromScan0 eW, eH, 0&, &HE200B, ByVal 0&, hBmp
    GdipGetImageGraphicsContext hBmp, hGrph
    GdipSetSmoothingMode hGrph, 4& '->SmoothingModeAntiAlias
    
    hPath = ppRound(shwRect.Left, 0, eW, eH)
    GdipCreateSolidFill ConvertColor(m_BorderColorFocus, m_Opacity), hBrush
    GdipFillPath hGrph, hBrush, hPath
    GdipDeleteBrush hBrush
    
    With shwRect
      m_BmpS = ppBlur(hBmp, oleShadowColor, m_Shadow)
    End With
    
    GdipDeletePath hPath
    GdipDeleteGraphics hGrph
    GdipDisposeImage hBmp
    
End Sub

'*2
Private Sub ppDraw()
Dim hGrph   As Long
Dim hPath   As Long
Dim hPen    As Long
Dim hBrush  As Long

    Edit.Alignment = m_Alignment
    mEdit.Alignment = m_Alignment
    Edit.Locked = m_Locked
    mEdit.Locked = m_Locked
    Edit.ForeColor = m_ForeColor
    mEdit.ForeColor = m_ForeColor

    With UserControl
        lY = m_Shadow * dpiScale
        lh = (.ScaleHeight) - (lY * 2)
        
        If m_ImagePosition = ImageInBox Then
          lX = (m_Shadow / 2) * dpiScale
        Else
          If m_BmpAlignH = ImageInLeft Then
            lX = (m_Shadow * dpiScale) + Px '+ (bS / 2)
          Else
            lX = (m_Shadow / 2) * dpiScale
          End If
        End If
        
        If m_ImagePosition = ImageInBox Then
          lW = (.ScaleWidth) - (lX * 2)
        Else
          If m_BmpAlignH = ImageInLeft Then
            lW = (.ScaleWidth) - (lX)
          Else
            lW = .ScaleWidth - ((m_Shadow * dpiScale) + Px + (bS / 2))
          End If
        End If
        
        .Cls
        .BackColor = m_BackColor
        
        If bMouseOver Then
          If bInFocus Then
            Edit.BackColor = m_BackColorFocus
            mEdit.BackColor = m_BackColorFocus
          Else
            Edit.BackColor = m_BackColorFocus
            mEdit.BackColor = m_BackColorFocus
          End If
        Else
            Edit.BackColor = m_BackColor
            mEdit.BackColor = m_BackColor
        End If
        
        If GdipCreateFromHDC(.hDC, hGrph) <> 0 Then Exit Sub
        GdipSetSmoothingMode hGrph, 4 '-> SmoothingModeAntiAlias
        
        If m_BmpS And bMouseOver Then
          If m_ImagePosition = ImageOutBox And m_BmpAlignH = ImageInRight Then
            GdipDrawImageRectI hGrph, m_BmpS, 0, 0, lW, UserControl.ScaleHeight
          Else
            GdipDrawImageRectI hGrph, m_BmpS, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
          End If
        End If
        
'AxioUK 'Dimensiones sombra?
        SetRect shwRect, lX, lY, lW, lh

        hPath = ppRound(lX, lY, lW, lh)
        
'AxioUK 'Shadow
        Call ppCreateShadow(IIf(bMouseOver Or bInFocus, m_ShadowColorFocus, m_ShadowColor))
        
        '\Back
        Edit.BackColor = IIf(bMouseOver Or bInFocus, m_BackColorFocus, m_BackColor)
        GdipCreateSolidFill ConvertColor(IIf(bMouseOver Or bInFocus, m_BackColorFocus, m_BackColor), 100), hBrush
        GdipFillPath hGrph, hBrush, hPath
        GdipDeleteBrush hBrush

        '\Border
        'GdipCreatePen1 ConvertColor(IIf(bMouseOver, m_BorderColorFocus, m_BorderColor), 100), (m_BorderWidth * dpiScale), &H2&, hPen
        GdipCreatePen1 ConvertColor(IIf(bMouseOver Or bInFocus, m_BorderColorFocus, m_BorderColor), m_Opacity), (m_BorderWidth * dpiScale), &H2&, hPen
        GdipDrawPath hGrph, hPen, hPath
        GdipDeletePen hPen
        
        If m_Bmp Then
            Call GdipSetInterpolationMode(hGrph, 7&)  'HIGH_QUALYTY_BICUBIC
            Call GdipSetPixelOffsetMode(hGrph, 4&)
            GdipDrawImageRectI hGrph, m_Bmp, m_BmpRct.Left, m_BmpRct.Top, m_BmpRct.Width, m_BmpRct.Height
        End If
        
        Call GdipDeletePath(hPath)
        Call GdipDeleteGraphics(hGrph)
        
       'Transparent
       If m_Transp Then
          .BackStyle = 0
          .MaskColor = .BackColor
          Set .MaskPicture = .Image
       End If
    End With
        
End Sub

'?GDIP
Private Sub ppGdipStart(ByVal Startup As Boolean)
    If Startup Then
        If m_token = 0& Then
            Dim gdipSI(3) As Long
            gdipSI(0) = 1&
            Call GdiplusStartup(m_token, gdipSI(0), ByVal 0)
        End If
    Else
        If m_token <> 0 Then Call GdiplusShutdown(m_token): m_token = 0
    End If
End Sub

Private Function ppRound(X As Long, Y As Long, ByVal W As Long, ByVal H As Long) As Long
Dim ePath   As Long
Dim BCLT    As Integer
Dim BCRT    As Integer
Dim BCBR    As Integer
Dim BCBL    As Integer

    W = W - 1 'Antialias pixel
    H = H - 1 'Antialias pixel
    
    BCLT = GetSafeRound(m_Round * dpiScale, W, H)
    BCRT = GetSafeRound(m_Round * dpiScale, W, H)
    BCBR = GetSafeRound(m_Round * dpiScale, W, H)
    BCBL = GetSafeRound(m_Round * dpiScale, W, H)
    
    Call GdipCreatePath(&H0, ePath)
    If BCLT Then GdipAddPathArcI ePath, X, Y, BCLT * 2, BCLT * 2, 180, 90
    If BCLT = 0 Then GdipAddPathLineI ePath, X, Y, X + W - BCRT, Y
        
    If BCRT Then GdipAddPathArcI ePath, X + W - BCRT * 2, Y, BCRT * 2, BCRT * 2, 270, 90
    If BCRT = 0 Then GdipAddPathLineI ePath, X + W, Y, X + W, Y + H - BCBR
        
    If BCBR Then GdipAddPathArcI ePath, X + W - BCBR * 2, Y + H - BCBR * 2, BCBR * 2, BCBR * 2, 0, 90
    If BCBR = 0 Then GdipAddPathLineI ePath, X + W, Y + H, X + BCBL, Y + H
    
    If BCBL Then GdipAddPathArcI ePath, X, Y + H - BCBL * 2, BCBL * 2, BCBL * 2, 90, 90
    If BCBL = 0 Then GdipAddPathLineI ePath, X, Y + H, X, Y + BCLT
    
    GdipClosePathFigures ePath
    ppRound = ePath
    
End Function


Private Function ppSaveBmp(ByVal FileName As String, Bmp As Long) As Boolean
Dim eGuid   As GUID
    If Bmp = 0& Then Exit Function
    CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), eGuid
    ppSaveBmp = GdipSaveImageToFile(Bmp, StrConv(FileName, vbUnicode), eGuid, ByVal 0&) = 0&
End Function

Private Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub

Private Sub Edit_Change()
m_Text = Edit.Text
If Not m_MultiLine Then RaiseEvent Change
End Sub

Private Sub Edit_GotFocus()
bInFocus = True
If m_AutoSel Then
    Edit.SelStart = 0
    Edit.SelLength = Len(Edit.Text)
End If
tmrOver.Enabled = True
End Sub

Private Sub Edit_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    
    Select Case KeyCode
        Case 13 '39, 40, 13  Next Control: right arrow, down arrow and Enter
            If m_KeyBehavior = eKeyTab Then
                WshShell.SendKeys "{Tab}"
            ElseIf m_KeyBehavior = Validate Then
                Call ValidateString(Edit)
            End If
        Case 37, 38 'Previous Control: left and up arrows
            WshShell.SendKeys "+{Tab}"
    End Select
End Sub

Private Sub Edit_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    SubKeyPress KeyAscii, Edit
End Sub

Private Sub SubKeyPress(KeyAscii As Integer, oCtrl As TextBox)
    Dim lCurPos As Long
    Dim lLineLength As Long
    Dim I As Integer
    
    Dim MoneyDolB As Boolean
    Dim MoneyDotB As Boolean
    Dim MoneyDot As String
    Dim MoneyDolLoc As Long
    Dim MoneyDotLoc As Long
    
    Dim PercentDotB As Boolean
    Dim PercentPerB As Boolean
    Dim PercentNum As String
    Dim PercentDot As String
    Dim PercentLoc As Long
    Dim PercentDotLoc As Long
    
    Dim DecimalDotB As Boolean
    
    Dim Space As Boolean
    Dim FractionSlash As Boolean
    Dim SpaceLoc As Long
    Dim FractionLoc As Long
    
    Dim ipPoint As Integer
    
    oCtrl.ForeColor = m_ForeColor
    
    Select Case FormatToString
    Case LettersOnly
        If Not (KeyAscii > 64 And KeyAscii < 91) And Not (KeyAscii > 96 And KeyAscii < 123) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
    
    Case NumbersOnly
        If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
    
    Case LettersAndNumbers
        If IsNumeric(Chr$(KeyAscii)) = False And (Not (KeyAscii > 64 And KeyAscii < 91) And Not (KeyAscii > 96 And KeyAscii < 123)) And KeyAscii <> 8 And KeyAscii <> 32 Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
    
    Case Money
        If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 36 And KeyAscii <> 44 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            If oCtrl.SelLength <> 0 Then
                Exit Sub
            End If
            
            ' Determine cursor position
            If oCtrl.SelLength = 0 Then
                lCurPos = oCtrl.SelStart
            Else
                lCurPos = oCtrl.SelStart + oCtrl.SelLength
            End If
            
            ' Determine textbox length
            lLineLength = SendMessage(oCtrl.hWnd, EM_LINELENGTH, lCurPos, 0)
            
            ' Determine location/existance of "$" and ","
            For I = 1 To lLineLength
                If Mid$(oCtrl.Text, I, 1) = "$" Then
                    MoneyDolB = True
                    MoneyDolLoc = I
                    Exit For
                End If
            Next I
            For I = 1 To lLineLength
                If Mid$(oCtrl.Text, I, 1) = "," Then
                    MoneyDotB = True
                    MoneyDotLoc = I
                    Exit For
                End If
            Next I
                        
            ' Make sure number only goes to 2 decimal places
            If MoneyDotB = True Then
                'MoneyDot = Mid$(oCtrl.Text, InStr(1, oCtrl.Text, ",") + 1, Len(oCtrl.Text) + InStr(1, oCtrl.Text, ",") + 1)
                MoneyDot = Mid$(oCtrl.Text, InStr(1, oCtrl.Text, ",") + 1, Len(oCtrl.Text) + 1)
                     
                If Len(MoneyDot) = 2 And lCurPos = MoneyDotLoc + 1 And KeyAscii <> 8 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If Len(MoneyDot) = 2 And lCurPos = MoneyDotLoc And KeyAscii <> 8 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If lCurPos = MoneyDotLoc + 2 And KeyAscii <> 8 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
            End If
                
            ' Make sure "," and "$" is only typed once
            If KeyAscii = 36 And MoneyDolB = False Then
                MoneyDolB = True
            ElseIf KeyAscii = 36 And MoneyDolB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            If lCurPos <> 0 And MoneyDolB <> False And KeyAscii = 36 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If

            If KeyAscii = 44 And MoneyDotB = False Then
                MoneyDotB = True
            ElseIf KeyAscii = 44 And MoneyDotB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
        End If
    
    Case Percent
        If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 37 And KeyAscii <> Asc(sDecimal) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            If oCtrl.SelLength <> 0 Then
                Exit Sub
            End If
            
            ' Determine cursor position
            If oCtrl.SelLength = 0 Then
                lCurPos = oCtrl.SelStart
            Else
                lCurPos = oCtrl.SelStart + oCtrl.SelLength
            End If
            
            ' Determine textbox length
            lLineLength = SendMessage(oCtrl.hWnd, EM_LINELENGTH, lCurPos, 0)
            
            ' Determine location of "%" and ","
            For I = 1 To lLineLength
                If Mid$(oCtrl.Text, I, 1) = "%" Then
                    PercentPerB = True
                    PercentLoc = I
                    Exit For
                End If
            Next I
            For I = 1 To lLineLength
                If Mid$(oCtrl.Text, I, 1) = sDecimal Then
                    PercentDotB = True
                    PercentDotLoc = I
                    Exit For
                End If
            Next I

            ' Make sure number only goes to 2 decimal places
            If PercentDotB = True Then
                PercentDot = Mid$(oCtrl.Text, InStr(1, oCtrl.Text, sDecimal) + 1, Len(oCtrl.Text) + InStr(1, oCtrl.Text, sDecimal) + 1)
        
                If InStr(1, PercentDot, "%") <> 0 Then
                    PercentDot = Mid$(PercentDot, 1, Len(PercentDot) - 1)
                End If
        
                If Len(PercentDot) = 2 And lCurPos = PercentDotLoc + 1 And KeyAscii <> 8 And KeyAscii <> 37 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If Len(PercentDot) = 2 And lCurPos = PercentDotLoc And KeyAscii <> 8 And KeyAscii <> 37 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If lCurPos = PercentDotLoc + 2 And KeyAscii <> 8 And KeyAscii <> 37 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
            End If

            ' Make sure "%" and "," is only typed once
            If KeyAscii = 37 And PercentPerB = False Then
                PercentPerB = True
            ElseIf KeyAscii = 37 And PercentPerB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            If lCurPos <> Len(oCtrl.Text) And PercentPerB <> False And KeyAscii = 37 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            If KeyAscii = Asc(sDecimal) And PercentDotB = False Then
                MoneyDotB = True
            ElseIf KeyAscii = Asc(sDecimal) And PercentDotB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            ' Make sure numbers are not written after the "%"
            If KeyAscii <> 37 And KeyAscii <> 8 And PercentPerB = True And lCurPos = PercentLoc Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            ' Determine if the percentage is >100
            If IsNumeric(Chr$(KeyAscii)) = True Then
                PercentNum = oCtrl.Text
                PercentNum = InsertStr(PercentNum, Chr$(KeyAscii), lCurPos + 1)
                If InStr(1, PercentNum, "%") <> 0 Then
                    If Val(Mid$(PercentNum, 1, Len(PercentNum) - 1)) > 100 Then
                        KeyAscii = 0
                        Beep
                        Exit Sub
                    End If
                Else
                    'If Val(PercentNum) > 100 Then
                    '    KeyAscii = 0
                    '    Beep
                    '    Exit Sub
                    'End If
                End If
            End If
        End If
    
    Case Fraction
        If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 47 And KeyAscii <> 32 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            If oCtrl.SelLength <> 0 Then
                Exit Sub
            End If
            ' Determine cursor position
            If oCtrl.SelLength = 0 Then
                lCurPos = oCtrl.SelStart
            Else
                lCurPos = oCtrl.SelStart + oCtrl.SelLength
            End If
            
            ' Determine textbox length
            lLineLength = SendMessage(oCtrl.hWnd, EM_LINELENGTH, lCurPos, 0)
            
            ' Determine location of " " and "/"
            For I = 1 To lLineLength
                If Mid$(oCtrl.Text, I, 1) = "/" Then
                    FractionLoc = I
                    Exit For
                End If
            Next I
    
            For I = 1 To lLineLength
                If Mid$(oCtrl.Text, I, 1) = " " Then
                    SpaceLoc = I
                    Exit For
                End If
            Next I
            
            If FractionLoc <> 0 Then
                FractionSlash = True
            End If
            If SpaceLoc <> 0 Then
                Space = True
            End If
            
            ' Don't allow more then 1 space in the field
            If (Space = True Or Fraction = True) And KeyAscii = 32 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            If Space = False And KeyAscii = 32 Then
                Space = True
            End If
            
            ' Check if " " is being used correctly
            If lCurPos = 0 And KeyAscii = 32 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            ' Don't allow more then 1 "/" in the field
            If FractionSlash = True And KeyAscii = 47 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            If FractionSlash = False And KeyAscii = 47 Then
                FractionSlash = True
            End If
            
            ' Check if "/" is being used correctly
            If lLineLength >= 1 Then
                If lCurPos > 0 Then
                    If KeyAscii = 47 And IsNumeric(Mid$(oCtrl.Text, lCurPos, 1)) = False Then
                        KeyAscii = 0
                        Beep
                        Exit Sub
                    End If
                End If
            ElseIf KeyAscii = 47 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
        End If
    
    Case Decimals
        If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 44 And KeyAscii <> 46 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            ' Determine textbox length
            lLineLength = SendMessage(oCtrl.hWnd, EM_LINELENGTH, lCurPos, 0)
        
            ' Determine existance of ","
            For I = 1 To lLineLength
                If Mid$(oCtrl.Text, I, 1) = sDecimal Then
                    DecimalDotB = True
                    Exit For
                End If
            Next I
                        
            ' Make sure Decimal separator is only typed once
            If sDecimal = Chr$(44) Then
              If KeyAscii = 44 And DecimalDotB = False Then
                  DecimalDotB = True
              ElseIf KeyAscii = 44 And DecimalDotB = True Then
                  KeyAscii = 0
                  Beep
                  Exit Sub
              End If
            ElseIf sDecimal = Chr$(46) Then
              If KeyAscii = 46 And DecimalDotB = False Then
                  DecimalDotB = True
              ElseIf KeyAscii = 46 And DecimalDotB = True Then
                  KeyAscii = 0
                  Beep
                  Exit Sub
              End If
            End If
        End If
    
    Case IPAddress
          ipPoint = 0
          If iCount >= 15 Then
            KeyAscii = 0
            Exit Sub
          End If
          
          If Len(oCtrl.Text) = 0 Then iCount = 0
          
          For I = 1 To Len(oCtrl.Text)
            If Mid$(oCtrl.Text, I, 1) = "." Then ipPoint = ipPoint + 1
          Next I
          
          Select Case KeyAscii
            Case 8 'Borrar
              iCount = iCount - 1
              
            Case 48 To 57
              
            Case 46
              If Len(oCtrl.Text) = 0 Then KeyAscii = 0
              If ipPoint = 3 Then KeyAscii = 0
              
            Case Else
              KeyAscii = 0
          End Select
    
    End Select
    
    Select Case CaseText
    Case UpperCase
        KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    Case LowerCase
        KeyAscii = Asc(LCase(Chr$(KeyAscii)))
    End Select
    
    If LenB(Trim$(m_Text)) <> 0& Then
      m_Text = oCtrl.Text
    Else
      m_Text = vbNullString
    End If
    
End Sub

Private Sub Edit_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Edit_LostFocus()
bInFocus = False
End Sub

Private Sub Edit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  bMouseOver = True
  tmrOver.Enabled = True
End Sub

Private Sub ValidateString(oCtrl As TextBox)
Dim sValor As String, I As Integer, sCar As String

  Select Case FormatToString
      Case Is = Money
        oCtrl.Text = fCleanValue(oCtrl.Text)
        oCtrl.Text = sMoney & " " & Format$(oCtrl.Text, "Standard")
        
    Case Is = Percent
        oCtrl.Text = fCleanValue(oCtrl.Text)

        If oCtrl.Text = "" Then
            oCtrl.Text = "0 %"
        End If
        If InStr(1, oCtrl.Text, "%") = 0 Then
            oCtrl.Text = oCtrl.Text & " %"
        End If
        If InStr(1, oCtrl.Text, "%") <> 0 Then
            If Len(oCtrl.Text) = 1 Then
                oCtrl.Text = "0 %"
            End If
        End If
        If InStr(1, oCtrl.Text, sDecimal) <> 0 Then
            If Mid$(oCtrl.Text, 1, Len(oCtrl.Text) - 1) = sDecimal Then
                oCtrl.Text = Mid$(oCtrl.Text, 1, Len(oCtrl.Text) - 2) & " %"
            End If
            If Mid$(oCtrl.Text, 1, 1) = sDecimal Then
                oCtrl.Text = "0" & Mid$(oCtrl.Text, 1, Len(oCtrl.Text))
            End If
        End If
        
    Case Is = NumbersOnly
      sValor = ""
      If oCtrl.Text = "" Then oCtrl.Text = "0"
      For I = 1 To Len(oCtrl.Text)
           sCar = Mid$(oCtrl.Text, I, 1)
           If IsNumeric(sCar) Then
              sValor = sValor & sCar
           ElseIf sCar = sDecimal Then
              Exit For
           End If
      Next I
      
      oCtrl.Text = sValor
      
    Case Is = LettersAndNumbers
      sValor = ""
      If oCtrl.Text = "" Then oCtrl.Text = "0"
      For I = 1 To Len(oCtrl.Text)
           sCar = Mid$(oCtrl.Text, I, 1)
          If IsNumeric(sCar) Then
              sValor = sValor & sCar
          ElseIf (Asc(sCar) >= Asc("a") And Asc(sCar) <= Asc("z")) Or (Asc(sCar) >= Asc("A") And Asc(sCar) <= Asc("Z")) Then
              sValor = sValor & sCar
          End If
      Next I
      
      oCtrl.Text = sValor
    
    Case Is = Fraction
        If oCtrl.Text = "" Then
            oCtrl.Text = "0"
        End If
        
        ' if the user inputs a fractional number
        If InStr(1, oCtrl.Text, "/") <> 0 Then
            ' if / is the first character in the text box then set to 0
            If InStr(1, oCtrl.Text, "/") = 1 Then
                oCtrl.Text = "0"
            ' make sure there are numbers before and after the /
            ElseIf (IsNumeric(Mid$(oCtrl.Text, InStr(1, oCtrl.Text, "/") - 1, 1)) = False) Or (IsNumeric(Mid$(oCtrl.Text, InStr(1, oCtrl.Text, "/") + 1, 1)) = False) Then
                oCtrl.Text = "0"
            End If
        End If
        oCtrl.Text = Trim(oCtrl.Text)
        
    Case Is = Decimals
        oCtrl.Text = fCleanValue(oCtrl.Text)
        If Trim$(oCtrl.Text) = "" Then oCtrl.Text = "0"
        'oCtrl.Text = FormatNumber(oCtrl.Text, 2, vbTrue)   'Yaco: Antes tenia en duro el numero 2, le agregue m_ParteDecimal
        oCtrl.Text = FormatNumber(CDbl(oCtrl.Text), m_ParteDecimal, vbTrue)

          
    Case Is = Dates
      If oCtrl.Text = "" Or oCtrl.Text = "00/00/0000" Then Exit Sub
      If Not IsDate(oCtrl.Text) Then
          For I = 1 To Len(oCtrl.Text)
            If I = 3 Or I = 5 Then
              sValor = sValor & "/" & Mid$(oCtrl.Text, I, 1)
            Else
              sValor = sValor & Mid$(oCtrl.Text, I, 1)
            End If
          Next I
      Else
          oCtrl.Text = Format(oCtrl, "Short Date")
          oCtrl.ForeColor = vbBlack
      End If
      
      oCtrl.Text = sValor
      
    'Case Is = ChileanRUT
    '    oCtrl.Text = FormatoRUT(oCtrl)
    '    If EsRut(oCtrl.Text) = False Then MsgBox "RUT no Válido...!", vbInformation + vbOKOnly, "Error!"
    
    Case Is = IPAddress
        Dim arIP() As String
        
        arIP = Split(oCtrl.Text, ".")
        If UBound(arIP) <> 3 Then
          oCtrl.ForeColor = vbRed
        Else
          For I = 0 To 3
           If (CInt(arIP(I)) > 255) Or (CInt(arIP(I)) < 0) Then
                  oCtrl.ForeColor = vbRed
                  Exit For
           Else
              oCtrl.Text = CInt(arIP(0)) & "." & CInt(arIP(1)) & "." & CInt(arIP(2)) & "." & CInt(arIP(3))
           End If
          Next I
        End If
  End Select
  
End Sub


Private Sub mEdit_Change()
m_Text = mEdit.Text
If m_MultiLine Then RaiseEvent Change
End Sub

Private Sub mEdit_GotFocus()
bInFocus = True
End Sub

Private Sub mEdit_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)

  If m_KeyBehavior = eKeyTab Then
      WshShell.SendKeys "{Tab}"
  ElseIf m_KeyBehavior = Validate Then
      Call ValidateString(mEdit)
  End If

End Sub

Private Sub mEdit_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
SubKeyPress KeyAscii, mEdit
End Sub

Private Sub mEdit_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub mEdit_LostFocus()
bInFocus = False
End Sub

Private Sub tmrOver_Timer()
If Not IsMouseOver(UserControl.hWnd) And Not IsMouseOver(Edit.hWnd) Then
  bMouseOver = False
  RaiseEvent MouseLeave
  tmrOver.Enabled = False
ElseIf IsMouseOver(UserControl.hWnd) Or IsMouseOver(Edit.hWnd) Then
  bMouseOver = True
  RaiseEvent MouseEnter
End If

ppDraw
End Sub

'*1
Private Function IsMouseOver(hWnd As Long) As Boolean
    Dim PT As POINTAPI
    GetCursorPos PT
    IsMouseOver = (WindowFromPoint(PT.X, PT.Y) = hWnd)
End Function

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Call ppCopyAmbient
    Call ppDraw
    'If m_ShadowVisible = True Then Call ppCreateShadow(m_ShadowColor)
End Sub

Private Sub UserControl_EnterFocus()
  bMouseOver = True
  If m_AutoSel Then
      Edit.SelStart = 0
      Edit.SelLength = Len(Edit.Text)
  End If
  tmrOver.Enabled = True
End Sub

Private Sub UserControl_ExitFocus()
    bMouseOver = False
    tmrOver.Enabled = True
End Sub

Private Sub UserControl_Initialize()
  ppGdipStart True
  dpiScale = GetWindowsDPI
  Set WshShell = CreateObject("WScript.Shell")

End Sub

Private Sub UserControl_InitProperties()
    m_BackColor = vbWhite
    m_BackColorFocus = vbWhite
    m_BorderColor = &HB2ACA5
    m_BorderColorFocus = &HE8A859   'RGB(148, 199, 240)
    m_ShadowColor = &HE8A859
    m_ShadowColorFocus = &HE8A859
    m_AutoSel = True
    m_Shadow = 3
    m_BmpSize = "0x0"
    m_ParteDecimal = 2        'Yaco
    m_BorderWidth = 1
    m_Round = 1
    m_ShadowColor = &HE8A859
    m_ShadowColorFocus = &HE8A859
    m_Pwd = False
    m_BmpAlignH = 0
    m_BmpAlignV = 1
    m_ImagePosition = 0
    m_MultiLine = False
    m_KeyBehavior = eNone
    m_FormatToString = AllChars
    m_CaseText = Normal
    Set Edit.Font() = Edit.Font
    Set mEdit.Font() = Edit.Font
    m_Alignment = 0
    m_Text = ""
    m_Locked = False
    m_ForeColor = vbBlack
    m_Transp = False
    m_Opacity = 100
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With m_BmpRct
  If X > .Left And X < (.Left + .Width) And Y > .Top And Y < (.Top + .Height) Then
    RaiseEvent IconClick
  End If
End With
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  bMouseOver = True
  tmrOver.Enabled = True
  
With m_BmpRct
  If X > .Left And X < (.Left + .Width) And Y > .Top And Y < (.Top + .Height) Then
    MousePointerHands True
  Else
    MousePointerHands False
  End If
End With
  
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        m_BackColor = .ReadProperty("BackColor", vbWhite)
        m_BackColorFocus = .ReadProperty("BackColorFocus", vbWhite)
        m_BorderColor = .ReadProperty("BorderColor", &HB2ACA5)
        m_BorderColorFocus = .ReadProperty("BorderColorFocus", &HE8A859)
        m_BorderWidth = .ReadProperty("BorderWidth", 1)
        m_AutoSel = .ReadProperty("AutoSel", True)
        m_Round = .ReadProperty("CornerRound", 1)
        m_Shadow = .ReadProperty("ShadowSize", 0)
        m_ShadowColor = .ReadProperty("ShadowColor", &HE8A859)
        m_ShadowColorFocus = .ReadProperty("ShadowColorFocus", &HE8A859)
        m_Pwd = .ReadProperty("Pwd", False)
        m_BmpSize = .ReadProperty("ImageResize", "0x0")
        m_BmpAlignH = .ReadProperty("ImageAlignH", 0)
        m_BmpAlignV = .ReadProperty("ImageAlignV", 1)
        m_ImagePosition = .ReadProperty("ImagePosition", 0)
        m_MultiLine = .ReadProperty("MultiLine", False)
        
        m_KeyBehavior = .ReadProperty("EnterKeyBehavior", eNone)
        m_FormatToString = .ReadProperty("FormatToString", AllChars)
        m_CaseText = .ReadProperty("CaseText", Normal)
        m_ParteDecimal = .ReadProperty("ParteDecimal", 2)      'yaco
        Set Edit.Font() = .ReadProperty("Font", Edit.Font)
        Set mEdit.Font() = .ReadProperty("Font", Edit.Font)

        m_Alignment = .ReadProperty("Alignment", 0)
        m_Text = .ReadProperty("Text", "")
        m_Locked = .ReadProperty("Locked", False)
        m_ForeColor = .ReadProperty("ForeColor", vbBlack)
        m_Transp = .ReadProperty("Transparent", False)
        
        m_Opacity = .ReadProperty("Opacity", 100)
        m_bvData() = .ReadProperty("ImageStream", "")
        Call SetPictureStream(m_bvData())
        If Ambient.UserMode Then Erase m_bvData
        
    End With
    
    sDecimal = fGetLocaleInfo(LOCALE_SDECIMAL)
    sThousand = fGetLocaleInfo(LOCALE_SMONTHOUSANDSEP)
    sDateDiv = fGetLocaleInfo(LOCALE_SDATE)
    sMoney = fGetLocaleInfo(LOCALE_SCURRENCY)

    If m_MultiLine Then
      mEdit.Text = m_Text
    Else
      Edit.Text = m_Text
    End If
    
    Call UserControl_Resize
End Sub

'*3
Private Sub UserControl_Resize()
Dim bs2 As Long

If m_MultiLine = True Then
  mEdit.Visible = True
  Edit.Visible = False
Else
  mEdit.Visible = False
  Edit.Visible = True
End If

  With UserControl
      bS = (2 * dpiScale) + (m_Shadow * dpiScale) + (m_Round * dpiScale)
      bs2 = (m_Shadow * dpiScale) + (2 * dpiScale)
      If m_Round = 0 Then bS = (2 * dpiScale) + (m_Shadow * dpiScale)
      
      If m_Bmp Then
          Px = m_BmpRct.Width + (2 * dpiScale)
          m_BmpRct.Left = IIf(m_ImagePosition = ImageInBox, bS - 1, 1)
          
          If m_BmpAlignH = 1 Then _
            m_BmpRct.Left = IIf(m_ImagePosition = ImageInBox, .ScaleWidth - ((bS / 2) + Px), .ScaleWidth - Px)
      
          Select Case m_BmpAlignV
            Case Is = ImgTop
                m_BmpRct.Top = bs2
            Case Is = ImgMiddle
                m_BmpRct.Top = (.ScaleHeight - m_BmpRct.Height) \ 2
            Case Is = ImgBottom
                m_BmpRct.Top = .ScaleHeight - (m_BmpRct.Height + bs2)
          End Select
      End If
      
    Th = .TextHeight("Ájq\") * dpiScale
    
    Edit.Height = Th + bS
    Edit.Top = (.ScaleHeight / 2) - (Edit.Height / 2)

    Select Case m_ImagePosition
      Case Is = ImageInBox
        Select Case m_BmpAlignH
            Case 0  '=ImageInLeft
              'Edit.Move (bS + Px + 1), bs2, .ScaleWidth - ((bS / 2) + bS + 1) - Px, .ScaleHeight - (bs2 * 2)
              Edit.Width = .ScaleWidth - ((bS / 2) + bS + 1) - Px
              Edit.Left = (bS + Px + 1)
              mEdit.Move (bS + Px + 1), bs2, .ScaleWidth - ((bS / 2) + bS + 1) - Px, .ScaleHeight - (bs2 * 2)
              
            Case 1  '=ImageInRight
              'Edit.Move (bS / 1.5), bs2, .ScaleWidth - ((bS / 1.5) + bS) - Px, .ScaleHeight - (bs2 * 2)
              Edit.Left = (bS / 1.5)
              Edit.Width = .ScaleWidth - ((bS / 1.5) + bS) - Px
              mEdit.Move (bS / 1.5), bs2, .ScaleWidth - ((bS / 1.5) + bS) - Px, .ScaleHeight - (bs2 * 2)
        End Select
        
      Case Is = ImageOutBox
        Select Case m_BmpAlignH
            Case 0  '=ImageInLeft
              'Edit.Move (bS + Px + 2), bs2, .ScaleWidth - ((bS / 2) + bS + 1) - Px, .ScaleHeight - (bs2 * 2)
              Edit.Left = (bS + Px + 2)
              Edit.Width = .ScaleWidth - ((bS / 2) + bS + 1) - Px
              mEdit.Move (bS + Px + 2), bs2, .ScaleWidth - ((bS / 2) + bS + 1) - Px, .ScaleHeight - (bs2 * 2)
              
            Case 1  '=ImageInRight
              'Edit.Move (bS / 1.5), bs2, .ScaleWidth - ((bS / 1.5) + bS) - Px, .ScaleHeight - (bs2 * 2)
              Edit.Left = (bS / 1.5)
              Edit.Width = .ScaleWidth - ((bS / 1.5) + bS) - Px
              mEdit.Move (bS / 1.5), bs2, .ScaleWidth - ((bS / 1.5) + bS) - Px, .ScaleHeight - (bs2 * 2)
        End Select
    End Select
  End With
  
  ppCopyAmbient
  ppDraw
    
End Sub

Private Sub UserControl_Show()
  ppCopyAmbient
  ppDraw
    'If m_ShadowVisible = True Then Call ppCreateShadow(m_ShadowColor)
End Sub

Private Sub UserControl_Terminate()
    tmrOver.Enabled = False
    If m_BmpS Then GdipDisposeImage m_BmpS
    If m_Bmp Then GdipDisposeImage m_Bmp
    ppGdipStart False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackColor", m_BackColor
        .WriteProperty "BackColorFocus", m_BackColorFocus
        .WriteProperty "BorderColor", m_BorderColor
        .WriteProperty "BorderColorFocus", m_BorderColorFocus
        .WriteProperty "BorderWidth", m_BorderWidth, 1
        .WriteProperty "ForeColor", m_ForeColor
        .WriteProperty "AutoSel", m_AutoSel
        .WriteProperty "CornerRound", m_Round
        .WriteProperty "ShadowSize", m_Shadow
        .WriteProperty "Font", Edit.Font
        .WriteProperty "Text", m_Text
        .WriteProperty "Locked", m_Locked
        .WriteProperty "Alignment", m_Alignment
        .WriteProperty "ShadowColor", m_ShadowColor
        .WriteProperty "ShadowColorFocus", m_ShadowColorFocus
        .WriteProperty "Pwd", m_Pwd
        .WriteProperty "ImageResize", m_BmpSize
        .WriteProperty "ImageAlignH", m_BmpAlignH
        .WriteProperty "ImageAlignV", m_BmpAlignV
        .WriteProperty "ImageStream", m_bvData
        .WriteProperty "ImagePosition", m_ImagePosition
        .WriteProperty "MultiLine", m_MultiLine
        
        .WriteProperty "EnterKeyBehavior", m_KeyBehavior, eNone
        .WriteProperty "FormatToString", m_FormatToString, AllChars
        .WriteProperty "CaseText", m_CaseText, Normal
        .WriteProperty "ParteDecimal", m_ParteDecimal    'yaco
        .WriteProperty "Transparent", m_Transp, False
        
        .WriteProperty "Opacity", m_Opacity, 100
    End With
End Sub

Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute Alignment.VB_UserMemId = -520
 Alignment = m_Alignment
 End Property
 
Property Let Alignment(ByVal Value As AlignmentConstants)
    m_Alignment = Value
    PropertyChanged "Alignment"
End Property

Property Get BackColorFocus() As OLE_COLOR
 BackColorFocus = m_BackColorFocus
 End Property
 
Property Let BackColorFocus(ByVal Value As OLE_COLOR)
    m_BackColorFocus = Value
    PropertyChanged "BackColorFocus"
End Property

Property Get BackColor() As OLE_COLOR
 BackColor = m_BackColor
 End Property
 
Property Let BackColor(ByVal Value As OLE_COLOR)
    m_BackColor = Value
    PropertyChanged "BackColor"
    ppDraw
End Property

Property Get BorderColorFocus() As OLE_COLOR
 BorderColorFocus = m_BorderColorFocus
 End Property
 
Property Let BorderColorFocus(ByVal Value As OLE_COLOR)
    m_BorderColorFocus = Value
    PropertyChanged "BorderColorFocus"
End Property

Property Get BorderColor() As OLE_COLOR
 BorderColor = m_BorderColor
 End Property
 
Property Let BorderColor(ByVal Value As OLE_COLOR)
    m_BorderColor = Value
    PropertyChanged "BorderColor"
    ppDraw
End Property

Public Property Get BorderWidth() As Single
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Single)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  ppDraw
End Property

Public Property Get CaseText() As CaseType
    CaseText = m_CaseText
End Property

Public Property Let CaseText(ByVal New_CaseText As CaseType)
    m_CaseText = New_CaseText
    PropertyChanged "CaseText"
End Property

Property Get CornerRound() As Long
 CornerRound = m_Round
End Property
 
Property Let CornerRound(ByVal nValue As Long)
    m_Round = nValue
    PropertyChanged "CornerRound"
    UserControl_Resize
End Property

Property Get DPI() As Double
DPI = dpiScale
End Property

Property Get ParteDecimal() As Long
'Agregado por YAco
 ParteDecimal = m_ParteDecimal
End Property
 
 Property Let ParteDecimal(ByVal Value As Long)
'Agregado por YAco
    m_ParteDecimal = Value
    PropertyChanged "ParteDecimal"
End Property

Public Property Get EnterKeyBehavior() As eEnterKeyBehavior
EnterKeyBehavior = m_KeyBehavior
End Property

Public Property Let EnterKeyBehavior(ByVal NewBehavior As eEnterKeyBehavior)
m_KeyBehavior = NewBehavior
PropertyChanged "EnterKeyBehavior"
End Property

Property Get Font() As StdFont
 Set Font = Edit.Font
 End Property
 
Property Set Font(ByVal Value As Font)
    Set Edit.Font = Value
    Set mEdit.Font = Value
    PropertyChanged "Font"
    UserControl_Resize
End Property

Property Get ForeColor() As OLE_COLOR
 ForeColor = m_ForeColor
 End Property
 
Property Let ForeColor(ByVal Value As OLE_COLOR)
    m_ForeColor = Value
    PropertyChanged "ForeColor"
    Call UserControl_Resize
End Property

Public Property Get FormatToString() As CharacterType
    FormatToString = m_FormatToString
End Property

Public Property Let FormatToString(ByVal New_FormatToString As CharacterType)
m_FormatToString = New_FormatToString
PropertyChanged "FormatToString"

If m_MultiLine Then
  Call ValidateString(mEdit)
Else
  Call ValidateString(Edit)
End If
End Property

Public Property Get Image() As Byte
Attribute Image.VB_ProcData.VB_Invoke_Property = "JTextppg"
  Image = m_Image
End Property

Public Property Let Image(ByVal NewImage As Byte)
  m_Image = NewImage
  PropertyChanged "Image"
End Property

Property Get ImageAlignH() As jImageAlignmentH
 ImageAlignH = m_BmpAlignH
 End Property
 
Property Let ImageAlignH(ByVal Value As jImageAlignmentH)
    m_BmpAlignH = Value
    Call UserControl_Resize
    PropertyChanged "ImageAlignH"
End Property

Property Get ImageAlignV() As jImageAlignmentV
 ImageAlignV = m_BmpAlignV
 End Property
 
Property Let ImageAlignV(ByVal Value As jImageAlignmentV)
    m_BmpAlignV = Value
    Call UserControl_Resize
    PropertyChanged "ImageAlignV"
End Property

Public Property Get ImagePosition() As jImagePosition
  ImagePosition = m_ImagePosition
End Property

Public Property Let ImagePosition(ByVal NewImagePosition As jImagePosition)
  m_ImagePosition = NewImagePosition
  Call UserControl_Resize
  PropertyChanged "ImagePosition"
End Property
 
Property Get ImageResize() As String
 ImageResize = m_BmpSize
 End Property
 
Property Let ImageResize(ByVal Value As String)
On Error Resume Next
Dim lSep As String
Dim imgW   As Long
Dim imgH   As Long

    If InStr(Value, "*") Then lSep = "*"
    If InStr(LCase(Value), "x") Then lSep = "x"
    imgW = Val(Split(Value, lSep)(0))
    imgH = Val(Split(Value, lSep)(1))
    
    m_BmpSize = imgW & "x" & imgH
    m_BmpRct.Width = IIf(imgW > 0, imgW * dpiScale, m_BmpSrcW)
    m_BmpRct.Height = IIf(imgW > 0, imgW * dpiScale, m_BmpSrcH)
    
    Call UserControl_Resize
    PropertyChanged "ImageResize"
End Property

Property Get ImageSize() As String
 ImageSize = m_BmpSrcW & "x" & m_BmpSrcH
 End Property
 
Public Property Get hWnd()
  hWnd = UserControl.hWnd
End Property

Property Get Opacity() As Long
Opacity = m_Opacity
End Property
 
Property Let Opacity(ByVal Value As Long)
m_Opacity = Value
PropertyChanged "Opacity"
ppDraw
End Property
 
Property Get Password() As Boolean
 Password = m_Pwd
 End Property
 
Property Get Multiline() As Boolean
  Multiline = m_MultiLine
End Property

Property Let Multiline(ByVal isMultiLine As Boolean)
  m_MultiLine = isMultiLine
    
  If m_MultiLine Then
    mEdit.Text = m_Text
  Else
    Edit.Text = m_Text
  End If
  
  PropertyChanged "MultiLine"
  UserControl_Resize
End Property

Property Let Password(ByVal Value As Boolean)
    m_Pwd = Value
    Edit.PasswordChar = IIf(m_Pwd, PASSWORD_CHAR, "")
    PropertyChanged "Pwd"
End Property

Property Get Locked() As Boolean
 Locked = m_Locked
 End Property
 
Property Let Locked(ByVal Value As Boolean)
    m_Locked = Value
    PropertyChanged "Locked"
End Property

Property Get SelFocus() As Boolean
 SelFocus = m_AutoSel
 End Property
 
Property Let SelFocus(ByVal Value As Boolean)
    m_AutoSel = Value
    PropertyChanged "AutoSel"
End Property

Property Get ShadowColorFocus() As OLE_COLOR
 ShadowColorFocus = m_ShadowColorFocus
End Property

Property Let ShadowColorFocus(ByVal Value As OLE_COLOR)
    m_ShadowColorFocus = Value
    'If bMouseOver Then ppDraw
    PropertyChanged "ShadowColorFocus"
End Property

Property Get ShadowColor() As OLE_COLOR
 ShadowColor = m_ShadowColor
 End Property
 
Property Let ShadowColor(ByVal Value As OLE_COLOR)
    m_ShadowColor = Value
    'If bMouseOver Then ppDraw
    PropertyChanged "ShadowColor"
End Property

Property Get ShadowSize() As Long
 ShadowSize = m_Shadow
 End Property
 
Property Let ShadowSize(ByVal Value As Long)
    If Value < 0 Then Value = 0
    m_Shadow = Value
    If m_BmpS Then
        GdipDisposeImage m_BmpS
        m_BmpS = 0
    End If
    UserControl_Resize
End Property

Property Get Text() As String
 Text = m_Text
End Property
 
Public Property Let Text(ByVal sValue As String)
   m_Text = sValue
   PropertyChanged "Text"
   
  If m_MultiLine Then
    mEdit.Text = m_Text
  Else
    Edit.Text = m_Text
  End If
   
End Property

Public Property Get Transparent() As Boolean
Transparent = m_Transp
End Property

Public Property Let Transparent(ByVal vNewValue As Boolean)
m_Transp = vNewValue
PropertyChanged "Transparent"

If vNewValue = False Then ppCopyAmbient
ppDraw
End Property

Property Get Version() As String
Version = VERS  'App.Major & "." & App.Minor & ".R" & App.Revision
End Property

Public Property Get Enabled() As Boolean
Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
mEnabled = vNewValue
UserControl.Enabled = mEnabled
Edit.Enabled = mEnabled
mEdit.Enabled = mEnabled
PropertyChanged "Enabled"
End Property


