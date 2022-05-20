VERSION 5.00
Begin VB.UserControl axGTabControl 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HitBehavior     =   2  'Use Paint
   KeyPreview      =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "axTabControl.ctx":0000
   Begin VB.Timer tmrSlide 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   490
      Top             =   225
   End
   Begin VB.PictureBox picSliderDsabled 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      Picture         =   "axTabControl.ctx":0312
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSliderNormal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      Picture         =   "axTabControl.ctx":05E8
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSliderDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      Picture         =   "axTabControl.ctx":08BE
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSliderHover 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      Picture         =   "axTabControl.ctx":0B94
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Timer TimerCheckMouseOut 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   15
      Top             =   225
   End
End
Attribute VB_Name = "axGTabControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-----------------------------------------------------------------------------------------------------
'// Title:    TabControl  (original)
'// Author:   Joshy Francis
'// Version:  1.0
'// Copyright: All rights reserved
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'// Mod:       AxGTabControl
'// Editor:    David Rojas (AxioUK)
'// Version:   1.3b
'// Copyright: All rights reserved
'-----------------------------------------------------------------------------------------------------
'--- for MST SubClassing (1)
#Const ImplNoIdeProtection = True ' (MST_NO_IDE_PROTECTION <> 0)
#Const ImplSelfContained = True

#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If
#If ImplSelfContained Then
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
#End If

Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1
Private Const SIGN_BIT                      As Long = &H80000000
Private Const EBMODE_DESIGN                 As Long = 0
Private m_pSubclass         As IUnknown
Private mUserMode As Boolean
Private mContainerHwnd As Long
Dim mUserControlHwnd As Long
Dim mUserControlHdc As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcAddressByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'--- End for MST SubClassing (1)

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
'Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
'Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function apiTranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal Color As Long, ByVal hpal As Long, ByRef RGBResult As Long) As Long
''-GDI+--------------------------------------------
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
'
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal token As Long)
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, ByRef brush As Long) As Long
'Private Declare Function GdipMeasureString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByRef mBoundingBox As RECTS, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
'Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
'Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipSetPenMode Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mPenMode As PenAlignment) As Long
'Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
'Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
'Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTS, ByVal mFormat As Long) As Long
'Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
'Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
'Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
'Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As eStringFormatFlags) As Long
'Private Declare Function GdipSetStringFormatHotkeyPrefix Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mHotkeyPrefix As HotkeyPrefix) As Long
'Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As eStringTrimming) As Long
'Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As eStringAlignment) As Long
'Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As eStringAlignment) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipSetClipRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipResetPath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal count As Long) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal Graphics As Long, ByVal brush As Long, ByRef pPoints As Any, ByVal count As Long, ByVal FillMode As Long) As Long

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Enum PenAlignment
    PenAlignmentCenter = &H0
    PenAlignmentInset = &H1
End Enum

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

'Public Enum eFlatSide
'  rUp
'  rBottom
'  rLeft
'  rRight
'End Enum

'Public Enum HLPosition
'    hlLeft
'    hlTop
'    hlRight
'    hlBottom
'End Enum

Private Const SmoothingModeAntiAlias As Long = 4
Private Const WrapModeTileFlipXY = &H3
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const UnitPixel As Long = &H2&
'-------------------------------------------------
Private Const DT_BOTTOM = &H8&
Private Const DT_CENTER = &H1&
Private Const DT_LEFT = &H0&
Private Const DT_CALCRECT = &H400&
Private Const DT_WORDBREAK = &H10&
Private Const DT_VCENTER = &H4&
Private Const DT_TOP = &H0&
Private Const DT_TABSTOP = &H80&
Private Const DT_SINGLELINE = &H20&
Private Const DT_RIGHT = &H2&
Private Const DT_NOCLIP = &H100&
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000&
Private Const DT_EXTERNALLEADING = &H200&
Private Const DT_EXPANDTABS = &H40&
Private Const DT_CHARSTREAM = 4&
Private Const DT_WORD_ELLIPSIS = &H40000
Private Const DT_END_ELLIPSIS = &H8000&
 
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const HALFTONE = 4
Private Const NULL_BRUSH = 5
Private Const NULL_PEN = 8
Private Const Transparent = 1
Private Const OPAQUE = 2
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type tTabInfo
    sCaption As String
    sTag  As String
    sKey As String
    ItemData As Long
    Left As Long
    Width As Long
    Image As Long
    'Alignment As TabCaptionAlignment
    Enabled As Boolean
    Visible As Boolean
    Active As Boolean
    Top As Long
    Height As Long
    Right As Long
    Bottom As Long
End Type

Private Type tControlInfo
    ctlName As String
    iTabIndex As Long
End Type

'Public Enum TabCaptionAlignment
'   TabCaptionAlignLeftTop = DT_LEFT Or DT_TOP
'   TabCaptionAlignLeftCenter = DT_LEFT Or DT_VCENTER
'   TabCaptionAlignLeftBottom = DT_LEFT Or DT_BOTTOM
'   TabCaptionAlignCenterTop = DT_CENTER Or DT_TOP
'   TabCaptionAlignCenterCenter = DT_CENTER Or DT_VCENTER
'   TabCaptionAlignCenterBottom = DT_CENTER Or DT_BOTTOM
'   TabCaptionAlignRightTop = DT_RIGHT Or DT_TOP
'   TabCaptionAlignRightCenter = DT_RIGHT Or DT_VCENTER
'   TabCaptionAlignRightBottom = DT_RIGHT Or DT_BOTTOM
'   TabCaptionSingleLine = DT_SINGLELINE
'   TabCaptionWordWrap = DT_WORDBREAK
'   TabCaptionEllipsis = DT_WORD_ELLIPSIS
'   TabCaptionNoPrefix = DT_NOPREFIX
'End Enum

Private Enum ssSliderStatus
    ssDisabled
    ssNormal
    ssDown
    ssHover
End Enum

'---------------------------------
Public Event TabClick(ByVal TabIndex As Long)
Public Event BeforeTabChange(ByVal LastTab As Long, ByRef NewTab As Long)
Public Event TabOrderChanged(ByVal LastTab As Long, ByRef NewTab As Long)
Public Event MouseDown(TabIndex As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(TabIndex As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(TabIndex As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(TabIndex As Long, KeyCode As Integer, Shift As Integer)
Public Event KeyPress(TabIndex As Long, KeyAscii As Integer)
Public Event KeyUp(TabIndex As Long, KeyCode As Integer, Shift As Integer)


Private GdipToken As Long
Private nScale    As Single
Private hGraphics As Long

Private m_Tabs() As tTabInfo
Private m_TabCount As Long
Private m_TabOrder() As Long
Private m_Ctls() As tControlInfo
Private m_ctlCount As Long

Private mhDC As Long, hBmp As Long, hBmpOld As Long
Private lTabIndex As Long
Private cx As Long, cy As Long
Private lTabDragging As Boolean
Private lTabDragged As Long
Private lTabHover As Long
Private mFont As IFont
Private selIndex As Long
Private SliderBox As RECT
Private bSliderShown As Boolean
Private lScrollX As Long, lScrollWidth As Long
Private LeftSliderStatus As ssSliderStatus
Private RightSliderStatus As ssSliderStatus
Private bInFocus As Boolean
Private m_AllowReorder As Boolean
Private m_FocusRect As Boolean
'Default Property Values:
Const m_def_BorderColor = &H8D4214
Const m_def_ColorDisabled = &H69BDFF
Const m_def_ColorActive = &H8D4214
Const m_def_BackColor = &H8D4214
Const m_def_ForeColor = &H808080
'Property Variables:
Dim m_BorderColor As OLE_COLOR
Dim m_ColorDisabled As OLE_COLOR
Dim m_ColorActive As OLE_COLOR
Dim m_BackColor1 As OLE_COLOR
Dim m_BackColor2 As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_ForeColorActive As OLE_COLOR
Dim m_Angle  As Single
Dim m_TabWidth As Long
Dim m_Enabled As Boolean

'*-
'Autor: wqweto http://www.vbforums.com/showthread.php?872819
' The Modern Subclassing Thunk (MST)
'=========================================================================
Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As axGTabControl
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgepwEB4BV1aLdCQUg8YIgz4AdC+L+oHHABIeAYvCBQgRHgGri8IFRBEeAauLwgVUER4Bq4vCBXwRHgGruQkAAADzpYHCABIeAVJqGP9SEFqL+IvCq7gBAAAAqzPAq4tEJAyri3QkFKWlg+8YagBX/3IM/3cM/1IYi0QkGIk4Xl+4NBIeAS1wEB4BwhAAZpCLRCQIgzgAdSqDeAQAdSSBeAjAAAAAdRuBeAwAAABGdRKLVCQE/0IEi0QkDIkQM8DCDAC4AkAAgMIMAJCLVCQE/0IEi0IEwgQADx8Ai1QkBP9KBItCBHUYiwpS/3EM/3IM/1Eci1QkBIsKUv9RFDPAwgQAkFWL7ItVGIsKi0EshcB0OFL/0FqJQgiD+AF3VIP4AHUJgX0MAwIAAHRGiwpS/1EwWoXAdTuLClJq8P9xJP9RKFqpAAAACHUoUjPAUFCNRCQEUI1EJARQ/3UU/3UQ/3UM/3UI/3IQ/1IUWVhahcl1EYsK/3UU/3UQ/3UM/3UI/1EgXcIYAA==" ' 1.4.2019 11:41:46
    Const THUNK_SIZE    As Long = 452
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitSubclassingThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 410)      '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 412)      '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 413)      '--- 413 = DefSubclassProc ordinal
        '--- for IDE protection
        Debug.Assert pvGetIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitSubclassingThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

#If ImplSelfContained Then
Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String

    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & App.hInstance & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & App.hInstance & "_" & sKey, lValue)
End Property
#End If
 
Private Property Get ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long) As Long
    Dim lPtr As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ThunkPrivateData, ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, 4)
    End If
End Property

Private Function pvGetIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvGetIdeOwner = True
End Function

Private Sub pvSubclass()
        pvUnsubclass
    If mUserControlHwnd <> 0 Then
        Set m_pSubclass = InitSubclassingThunk(mUserControlHwnd, Me, InitAddressOfMethod(Me, 5).SubclassProc(0, 0, 0, 0, 0))
    End If
End Sub

Private Sub pvUnsubclass()
    Set m_pSubclass = Nothing
End Sub

Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
    
    If mUserMode = True Then
        Handled = True
        SubclassProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    End If

Dim lTab As Long
Dim X As Single
Dim Y As Single
    Select Case wMsg
         Case WM_LBUTTONDOWN ' UserControl message, only in design mode (Not Ambient.UserMode), to provide change of selected tab by clicking at design time
                    
                    lTab = selIndex
                        
                        X = (lParam And &HFFFF&)
                        Y = (lParam \ &H10000 And &HFFFF&)
                    
                    Call UserControl_MouseDown(vbLeftButton, -1, X, Y)
                    If selIndex <> lTab Then
                        Handled = True
                        SubclassProc = 0
                        Exit Function
                    End If
    End Select
        Handled = True '<<<< changed here
        SubclassProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)

End Function

'--- End for MST subclassing (2)

Public Function AddTab(Optional ByVal sKey As String, Optional ByVal sCaption As String, Optional ByVal lWidth As Long = 64, Optional ByVal ItemData As Long, Optional ByVal lImage As Long) As Long
ReDim Preserve m_Tabs(m_TabCount)
    With m_Tabs(m_TabCount)
      .sCaption = sCaption
      .sKey = sKey
      .Width = lWidth
      .ItemData = ItemData
      .Image = lImage
      '.Alignment = DT_CENTER Or DT_VCENTER 'DT_VCENTER Or DT_NOPREFIX Or DT_END_ELLIPSIS
      .Visible = True
      .Enabled = True
    End With
    
    ReDim Preserve m_TabOrder(m_TabCount)
    m_TabOrder(m_TabCount) = m_TabCount
    m_TabCount = m_TabCount + 1
    AddTab = m_TabCount
            
    Refresh
End Function

'*3
Sub Draw()
Dim hFont As Long, hFontOld As Long
Dim hBrush As Long, hBrushOld As Long
Dim hPen As Long, hPenOld As Long
Dim i As Long, c As Long, X As Long
Dim rcCalc As RECT, rc As RECT, lWidth As Long, lHeight As Long
Dim W As Long, H As Long, OldTextColor As Long

  If mFont Is Nothing Then
      Set mFont = UserControl.Font
  End If
  
  mhDC = UserControl.hDC
  hFontOld = SelectObject(mhDC, mFont.hFont)
  lWidth = ScaleX(ScaleWidth, ScaleMode, 3)
  lHeight = ScaleY(ScaleHeight, ScaleMode, 3)
  
  GdipCreateFromHDC hDC, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
     
  X = 4 + -lScrollX
  
For i = 0 To m_TabCount - 1
        c = m_TabOrder(i)
    With m_Tabs(c)
        If .Visible = True Then
              rcCalc.Left = 0
              rcCalc.Right = lWidth
              rcCalc.Top = 0
              rcCalc.Bottom = lHeight
            SetTextColor mhDC, m_ForeColor
            DrawText mhDC, .sCaption, Len(.sCaption), rcCalc, DT_CALCRECT Or DT_CENTER '.Alignment
              .Left = X
              .Right = .Left + (rcCalc.Right + 36) ' 48)
              .Width = IIf(m_TabWidth <> 0, m_TabWidth, .Right - .Left)
              .Height = .Bottom - .Top
              .Top = 1
              .Bottom = .Top + rcCalc.Bottom + 8
              X = X + IIf(m_TabWidth <> 0, m_TabWidth, .Right - .Left) + 2 'x + (.Right - .Left)
        End If
    End With
Next

  Dim B As Boolean
  B = bSliderShown
  bSliderShown = X + lScrollX + 32 >= lWidth
  lScrollWidth = ((X + lScrollX) - lWidth) + 48 '50
      If B = False And bSliderShown = True Then
          RightSliderStatus = ssNormal
      End If
  If lScrollWidth < 0 Then
      lScrollX = 0
  End If
        
'Draw Border/Back HERE!
Dim uRct As RECTL   '-1, -1, lWidth + 1, rcCalc.Bottom + 9
uRct.Left = 0
uRct.Top = 25
uRct.Width = lWidth - 1
uRct.Height = lHeight - 26

fRoundRect hGraphics, uRct, AColor(m_BackColor1, 100), AColor(m_BackColor2, 100), m_Angle, 1, AColor(m_BorderColor, 100), 5

For i = m_TabCount - 1 To 0 Step -1
        c = m_TabOrder(i)
   If c <> selIndex Then
        With m_Tabs(c)
            If .Left > -.Width And .Right < lWidth + .Width Then
                DrawTab hGraphics, c
            End If
        End With
    End If
Next

If m_TabCount > 0 Then
    DrawTab hGraphics, selIndex
End If

    If lTabDragging = True And lTabHover <> lTabDragged And lTabHover >= 0 And lTabHover < m_TabCount And m_TabCount > 0 Then
        c = m_TabOrder(lTabHover)
        DrawTab hGraphics, c
        SetTextColor mhDC, m_ColorActive 'OldTextColor
       With m_Tabs(c)
             rc.Left = .Left: rc.Top = .Top: rc.Right = .Right: rc.Bottom = .Bottom
                rc.Left = rc.Left + 4
                rc.Right = rc.Right - 24
                rc.Top = rc.Top + 4
                rc.Bottom = rc.Bottom - 4
        
              'Draw Dragging Rectangle HERE!
              uRct.Left = 0
              uRct.Top = 25
              uRct.Width = lWidth
              uRct.Height = lHeight
              
              'fRoundRect hGraphics, uRct, AColor(m_BackColor1, 90), AColor(vbBlue, 90), 45, 2, AColor(m_BorderColor, 90), 5
       End With
    End If

If bSliderShown Then
    Dim sldOldColor As Long
    sldOldColor = GetPixel(picSliderNormal.hDC, 1, 1)
    ReplaceColor picSliderDown, sldOldColor, m_BackColor2
    ReplaceColor picSliderDsabled, sldOldColor, m_BackColor1
    ReplaceColor picSliderNormal, sldOldColor, m_BackColor1
    ReplaceColor picSliderHover, sldOldColor, m_ColorActive

    With SliderBox
        .Top = 0
        .Bottom = rcCalc.Bottom + 8
        .Left = lWidth - IIf(.Bottom > 32, .Bottom, 32)
        .Right = lWidth
            
        ''Draw SliderBox HERE!
        'Dim slR As RECTL
        'slR.Left = .Left: slR.Top = .Top
        'slR.Height = .Bottom - .Top - 2: slR.Width = .Right - .Left
        'fRoundRect hGraphics, slR, AColor(m_BackColor1, 90), AColor(m_BackColor1, 90), 45, 1, m_BorderColor, 0
        
            W = ((.Right - .Left) - 4) / 2
            H = ((.Bottom - .Top) - 4) / 2
                If W < 14 Then
                    W = 14
                End If
                If H < 15 Then
                    H = 15
                End If
        
        Select Case LeftSliderStatus
            Case ssSliderStatus.ssDisabled
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, -W, H, picSliderDsabled.hDC, 0, 0, 14, 14, vbSrcCopy
            Case ssSliderStatus.ssNormal
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, -W, H, picSliderNormal.hDC, 0, 0, 14, 14, vbSrcCopy
            Case ssSliderStatus.ssDown
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, -W, H, picSliderDown.hDC, 0, 0, 14, 14, vbSrcCopy
            Case ssSliderStatus.ssHover
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, -W, H, picSliderHover.hDC, 0, 0, 14, 14, vbSrcCopy
        End Select
        Select Case RightSliderStatus
            Case ssSliderStatus.ssDisabled
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, W, H, picSliderDsabled.hDC, 0, 0, 14, 14, vbSrcCopy
            Case ssSliderStatus.ssNormal
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, W, H, picSliderNormal.hDC, 0, 0, 14, 14, vbSrcCopy
            Case ssSliderStatus.ssDown
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, W, H, picSliderDown.hDC, 0, 0, 14, 14, vbSrcCopy
            Case ssSliderStatus.ssHover
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, W, H, picSliderHover.hDC, 0, 0, 14, 14, vbSrcCopy
        End Select
    End With
End If

    'Replace or Keep?
    'BitBlt mUserControlHdc, 0, 0, lWidth, lHeight, mhDC, 0, 0, vbSrcCopy
    
    'GDI+---
    GdipDeleteGraphics hGraphics
    
   'Useless?---------------
    UserControl.BackStyle = 0
    UserControl.MaskColor = UserControl.BackColor
    Set UserControl.MaskPicture = UserControl.Image
    '------------------------
End Sub

Public Function FindTabByKey(ByVal sKey As String) As Long
Dim c As Long
    FindTabByKey = -1
        For c = 0 To m_TabCount - 1
            If m_Tabs(c).sKey = sKey Then
                FindTabByKey = c
                Exit For
            End If
        Next
End Function

Sub GetRGB(Color As Long, Red As Integer, Green As Integer, Blue As Integer)
Blue = (Color And &HFF0000) / (2 ^ 16)
Green = (Color And &HFF00&) / (2 ^ 8)
Red = (Color And &HFF&)
End Sub

Private Function HasIndex(ByVal ctl As Control) As Boolean
    'determine if it's a control array
    HasIndex = Not ctl.Parent.Controls(ctl.Name) Is ctl
End Function

Public Function IsDir(ByVal sDir As String) As Boolean
    IsDir = ((GetFileAttributes(sDir) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
End Function

Public Function IsFile(ByVal sFile As String) As Boolean
    IsFile = GetFileAttributes(sFile) <> -1
End Function

Sub LoadTabOrder(ByVal sFile As String)
Dim ctlCount As Long, nIndex As Long, tempOrder() As Long, m_tabIndex() As Long
Dim j As Long, tabCount As Long, k As Long
Dim m_tabLeft() As Single, m_tabVisible() As Boolean, ctl As Control
    If IsFile(sFile) = False Then
        Exit Sub
    End If
    
Open sFile For Binary As 1
    Get #1, , nIndex
    Get #1, , ctlCount
    If ctlCount > 0 Then
        Get #1, , tabCount
            ReDim tempOrder(tabCount)
        If tabCount > 0 Then
            ReDim tempOrder(tabCount - 1)
        End If
            ReDim m_tabIndex(ctlCount - 1)
            ReDim m_tabLeft(ctlCount - 1)
            ReDim m_tabVisible(ctlCount - 1)
            Get #1, , tempOrder
            Get #1, , m_tabIndex
            Get #1, , m_tabLeft
            Get #1, , m_tabVisible
    End If
Close 1
    
    If ctlCount > 0 Then
        If tabCount > 0 Then
            For j = 0 To IIf(m_TabCount > tabCount, tabCount, m_TabCount) - 1
                 m_TabOrder(j) = tempOrder(j)
            Next
        End If
            For j = 0 To IIf(m_ctlCount > ctlCount, ctlCount, m_ctlCount) - 1
                m_Ctls(j).iTabIndex = m_tabIndex(j)
                    For Each ctl In UserControl.ContainedControls
                        If pGetControlId(ctl) = m_Ctls(j).ctlName Then
                             ctl.Left = m_tabLeft(j)
                             ctl.Visible = m_tabVisible(j)
                             Exit For
                        End If
                    Next
            Next
            SelectedItem = nIndex
    End If
End Sub

Public Sub Refresh()
  UserControl.Cls
  Draw
  UserControl.Refresh
End Sub

Public Sub RemoveTab(ByVal nIndex As Long)
Dim j As Long
Dim itemOrder As Long
If nIndex < m_TabCount Then
        itemOrder = m_TabOrder(nIndex)
   '// Reset m_Tabs
   For j = m_TabOrder(nIndex) To m_TabCount - 2
      m_Tabs(j) = m_Tabs(j + 1)
   Next
   '// Adjust m_TabOrder
   For j = nIndex To m_TabCount - 2
      m_TabOrder(j) = m_TabOrder(j + 1)
   Next
   '// Validate Indexes for Items after deleted Item
   For j = 0 To m_TabCount - 1
      If m_TabOrder(j) > itemOrder Then
         m_TabOrder(j) = m_TabOrder(j) - 1
      End If
   Next

m_TabCount = m_TabCount - 1
    ReDim Preserve m_Tabs(m_TabCount)
    ReDim Preserve m_TabOrder(m_TabCount)
        If selIndex > m_TabCount - 1 Then
            SelectedItem = m_TabCount - 1
        Else
            Refresh
        End If
End If
End Sub

Sub SaveTabOrder(ByVal sFile As String)
    
    If IsFile(sFile) = True Then
        DeleteFile sFile
    End If
Dim m_tabIndex() As Long, j As Long
Dim m_tabLeft() As Single, m_tabVisible() As Boolean, ctl As Control

        ReDim m_tabIndex(m_ctlCount)
        ReDim m_tabLeft(m_ctlCount)
        ReDim m_tabVisible(m_ctlCount)
    If m_ctlCount > 0 Then
        ReDim m_tabIndex(m_ctlCount - 1)
        ReDim m_tabLeft(m_ctlCount - 1)
        ReDim m_tabVisible(m_ctlCount - 1)
    End If
            For j = 0 To m_ctlCount - 1
                m_tabIndex(j) = m_Ctls(j).iTabIndex
                    For Each ctl In UserControl.ContainedControls
                        If pGetControlId(ctl) = m_Ctls(j).ctlName Then
                             m_tabLeft(j) = ctl.Left
                             m_tabVisible(j) = ctl.Visible
                             Exit For
                        End If
                    Next
            Next
'Open sFile For Binary As 1

'Close 1
    
End Sub

Public Sub SetTabCaption(ByVal nIndex As Long, ByVal sCaption As String)
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    m_Tabs(nIndex).sCaption = sCaption
    Draw
End If
End Sub

Public Sub SetTabEnabled(ByVal nIndex As Long, ByVal Newval As Boolean)
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    m_Tabs(m_TabOrder(nIndex)).Enabled = Newval
    Refresh
End If
End Sub

Public Function tabCount() As Long
    tabCount = m_TabCount
End Function

Private Sub AddTabControls(ByVal lIndex As Long, ByVal ctlName As String)
    ReDim Preserve m_Ctls(m_ctlCount)
        m_Ctls(m_ctlCount).ctlName = ctlName
        m_Ctls(m_ctlCount).iTabIndex = lIndex
    m_ctlCount = m_ctlCount + 1
End Sub

'*2
Private Sub DrawTab(ByVal hGraphics As Long, ByVal i As Long)
Dim OldTextColor As Long
Dim rc As RECT, tRct As RECTL

Dim ColX As Long, ColY As Long

If i > -1 And i < m_TabCount And m_TabCount > 0 Then
    With m_Tabs(i)
        rc.Left = .Left:     rc.Top = .Top
        rc.Right = IIf(m_TabWidth = 0, .Right, .Left + m_TabWidth - 10)
        rc.Bottom = .Bottom

        tRct.Left = .Left:    tRct.Top = .Top
        tRct.Width = .Width:  tRct.Height = IIf(selIndex = i, .Height, .Height - 4)
        
        If selIndex = i Then
            SetTextColor mhDC, m_ForeColorActive
            fRoundCut hGraphics, tRct, AColor(m_BackColor1, 100), AColor(m_ColorActive, 100), 0, 2, AColor(m_BorderColor, 100), 6, rBottom
            DrawLine hGraphics, rc.Left, rc.Bottom - 2, rc.Right + 10, rc.Bottom - 2, m_ColorActive, 90, 5
        Else
            SetTextColor mhDC, m_ForeColor
            fRoundRect hGraphics, tRct, AColor(m_BackColor1, 100), AColor(m_BackColor2, 100), 0, 1, AColor(m_BorderColor, 100), 6
        End If
                    
        If .Enabled = False Then
            OldTextColor = SetTextColor(mhDC, m_ColorDisabled)     'RGB(128, 128, 128))
        End If
        DrawText mhDC, .sCaption, Len(.sCaption), rc, DT_CENTER Or DT_VCENTER '.Alignment
        If .Enabled = False Then
            SetTextColor mhDC, OldTextColor
        End If
            
        If selIndex = i Then
            If bInFocus = True Then
                DrawText mhDC, .sCaption, Len(.sCaption), rc, DT_CENTER Or DT_VCENTER '.Alignment
                rc.Left = rc.Left
                rc.Right = rc.Right
                If m_FocusRect Then DrawFocusRect mhDC, rc
            End If
        End If
    End With
End If
End Sub

Private Sub FocusTab(ByVal lTab As Long)
Dim lWidth As Long, lHeight As Long
If lTab < m_TabCount And m_TabCount > 0 Then
    If bSliderShown = True Then
        lWidth = ScaleX(ScaleWidth, ScaleMode, 3)
        lHeight = ScaleY(ScaleHeight, ScaleMode, 3)
        
        If m_Tabs(lTab).Right + IIf(bSliderShown, (SliderBox.Right - SliderBox.Left) + 16, 0) > (lWidth) Then
            lScrollX = lScrollX + m_Tabs(lTab).Right + (SliderBox.Right - SliderBox.Left) + 16 - lWidth
                LeftSliderStatus = ssNormal
                RightSliderStatus = ssNormal
            If lScrollX > lScrollWidth Then
                lScrollX = lScrollWidth
                RightSliderStatus = ssDisabled
            End If
        ElseIf m_Tabs(lTab).Left < (-4) Then
            lScrollX = 4 + m_Tabs(lTab).Left
                LeftSliderStatus = ssNormal
                RightSliderStatus = ssNormal
            If lScrollX < 0 Then
                lScrollX = 0
                LeftSliderStatus = ssDisabled
            End If
        End If
        Dim lTabOrder As Long
                lTabOrder = getTabOrder(lTab)
            If lTabOrder = 0 Then
                LeftSliderStatus = ssDisabled
            ElseIf lTabOrder = m_TabCount - 1 Then
                RightSliderStatus = ssDisabled
            End If
    End If
End If
End Sub

Private Function getTabOrder(ByVal Index As Long) As Long
Dim c As Long
    For c = 0 To m_TabCount - 1
        If m_TabOrder(c) = Index Then
            getTabOrder = c
            Exit For
        End If
    Next
End Function
 
Private Sub handleControls(ByVal LastIndex As Long, ByVal nIndex As Long)
Dim i As Long, z As Long, mCTL As Control, ctlName As String
If m_TabCount > 0 Then
    LastIndex = getTabOrder(LastIndex)
    nIndex = getTabOrder(nIndex)
End If
            
    For Each mCTL In UserControl.ContainedControls
                ctlName = pGetControlId(mCTL)
        If IsInTabControls(ctlName, nIndex) <> -1 Then
            If mCTL.Left < -35000 Then
                mCTL.Visible = True
                mCTL.Left = mCTL.Left + 70000
            End If
        Else
            If mCTL.Left > -35000 Then
                If IsInTabControls(ctlName, LastIndex) = -1 Then
                    AddTabControls LastIndex, ctlName
                End If
                    mCTL.Left = mCTL.Left - 70000
                    mCTL.Visible = False
            End If
        End If
    Next
Dim j As Long
    For j = 0 To m_ctlCount - 1
        If m_Ctls(j).iTabIndex > (m_TabCount - 1) Then
            For Each mCTL In UserControl.ContainedControls
                    ctlName = pGetControlId(mCTL)
                If Trim$(m_Ctls(j).ctlName) = ctlName Then
                    If mCTL.Left > -35000 Then
                        mCTL.Left = mCTL.Left - 70000
                        mCTL.Visible = False
                    End If
                        Exit For
                End If
            Next
        End If
    Next
End Sub

'*1
Private Function HitTest(ByVal lX As Long, ByVal lY As Long) As Long
Dim i As Long, c As Long, rc As RECT, lTab As Long
        lTab = -1
    For i = 0 To m_TabCount - 1
            c = m_TabOrder(i)
        With m_Tabs(c)
            If .Enabled = True And .Visible = True Then
                    rc.Left = .Left
                    rc.Top = .Top
                    rc.Right = IIf(m_TabWidth = 0, .Right + 10, .Left + m_TabWidth - 10) '.Right
                    rc.Bottom = .Bottom
                If lX >= rc.Left And lX <= rc.Right And lY >= rc.Top And lY <= rc.Bottom Then
                    lTab = i ' C
                    Exit For
                End If
            End If
        End With
    Next
        HitTest = lTab
End Function

Private Function IsInTabControls(ByVal ctlName As String, ByVal lIndex As Long) As Long
Dim j As Long
        IsInTabControls = -1
    For j = 0 To m_ctlCount - 1
        If Trim$(m_Ctls(j).ctlName) = ctlName And m_Ctls(j).iTabIndex = lIndex Then
            IsInTabControls = j
            Exit For
        End If
    Next
End Function

Private Sub MoveTab(ByVal nTab As Long, ByVal toTab As Long)
Dim c As Long, j As Long, tempIndex As Long
Dim nInfo As tTabInfo, nIndex As Long, toIndex As Long
'****** Sorting Controls **************
Dim NoMoreSwaps As Boolean, NumberOfItems As Long, Temp As tControlInfo, bDirection As Boolean
    bDirection = True
        NumberOfItems = UBound(m_Ctls)
    Do Until NoMoreSwaps = True
            NoMoreSwaps = True
         For c = 0 To (NumberOfItems - 1)
            If bDirection = True Then 'Ascending
                 If m_Ctls(c).iTabIndex > m_Ctls(c + 1).iTabIndex Then
                     NoMoreSwaps = False
                     Temp = m_Ctls(c)
                     m_Ctls(c) = m_Ctls(c + 1)
                     m_Ctls(c + 1) = Temp
                 End If
            Else
                 If m_Ctls(c).iTabIndex < m_Ctls(c + 1).iTabIndex Then
                     NoMoreSwaps = False
                     Temp = m_Ctls(c)
                     m_Ctls(c) = m_Ctls(c + 1)
                     m_Ctls(c + 1) = Temp
                 End If
            End If
         Next
            NumberOfItems = NumberOfItems - 1
    Loop
'********* End Sorting *************************************
            nIndex = m_TabOrder(nTab)
            toIndex = m_TabOrder(toTab)
If toTab > nTab Then
   For c = nTab To toTab - 1
      m_TabOrder(c) = m_TabOrder(c + 1)
   Next
        For j = 0 To m_ctlCount - 1
            If m_Ctls(j).iTabIndex > nTab And m_Ctls(j).iTabIndex <= toTab Then
                m_Ctls(j).iTabIndex = m_Ctls(j).iTabIndex - 1
            ElseIf m_Ctls(j).iTabIndex = nTab Then
                m_Ctls(j).iTabIndex = toTab
            End If
        Next

Else
   For c = nTab To toTab + 1 Step -1
      m_TabOrder(c) = m_TabOrder(c - 1)
   Next
        For j = 0 To m_ctlCount - 1
            If m_Ctls(j).iTabIndex >= toTab And m_Ctls(j).iTabIndex < nTab Then
                m_Ctls(j).iTabIndex = m_Ctls(j).iTabIndex + 1
            ElseIf m_Ctls(j).iTabIndex = nTab Then
                m_Ctls(j).iTabIndex = toTab
            End If
        Next
End If
    
    m_TabOrder(toTab) = nIndex

End Sub

Private Function pGetControlId(ByVal oCtl As Control) As String
Dim sCtlName As String
Dim iCtlIndex As Integer
        iCtlIndex = -1
    sCtlName = oCtl.Name
On Error Resume Next
If HasIndex(oCtl) Then
    iCtlIndex = oCtl.Index
End If
    pGetControlId = sCtlName & IIf(iCtlIndex <> -1, "(" & iCtlIndex & ")", "")
End Function

Private Sub TransBlt(ByVal hdcScreen As Long, ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, _
            ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, _
            ByVal xSrc As Long, ByVal ySrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal clrMask As OLE_COLOR)
'one check to see if GdiTransparentBlt is supported
'better way to check if function is suported is using LoadLibrary and GetProcAdress
'than using GetVersion or GetVersionEx
'=====================================================
    Dim Lib As Long
    Dim ProcAdress As Long
    Dim lMaskColor As Long
    lMaskColor = TranslateColor(clrMask)
    Lib = LoadLibrary("gdi32.dll")
    '--------------------->make sure to specify corect name for function
    ProcAdress = GetProcAddress(Lib, "GdiTransparentBlt")
    FreeLibrary Lib
    If ProcAdress <> 0 Then
        'works on XP
        GdiTransparentBlt hdcDest, xDest, yDest, nWidth, nHeight, hdcSrc, xSrc, ySrc, nWidthSrc, nHeightSrc, lMaskColor
        Exit Sub 'make it short
    End If
'=====================================================
    Const DSna              As Long = &H220326
    Dim hdcMask             As Long
    Dim hdcColor            As Long
    Dim hbmMask             As Long
    Dim hbmColor            As Long
    Dim hbmColorOld         As Long
    Dim hbmMaskOld          As Long
    Dim hdcScnBuffer        As Long
    Dim hbmScnBuffer        As Long
    Dim hbmScnBufferOld     As Long
    
   lMaskColor = TranslateColor(clrMask)
   hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
   hdcScnBuffer = CreateCompatibleDC(hdcScreen)
   hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)

   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcDest, xDest, yDest, vbSrcCopy

   hbmColor = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
   hbmMask = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)

   hdcColor = CreateCompatibleDC(hdcScreen)
   hbmColorOld = SelectObject(hdcColor, hbmColor)
    
   Call SetBkColor(hdcColor, GetBkColor(hdcSrc))
   Call SetTextColor(hdcColor, GetTextColor(hdcSrc))
   Call StretchBlt(hdcColor, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, nWidthSrc, nHeightSrc, vbSrcCopy)

   hdcMask = CreateCompatibleDC(hdcScreen)
   hbmMaskOld = SelectObject(hdcMask, hbmMask)

   SetBkColor hdcColor, lMaskColor
   SetTextColor hdcColor, vbWhite
   BitBlt hdcMask, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcCopy
 
   SetTextColor hdcColor, vbBlack
   SetBkColor hdcColor, vbWhite
   BitBlt hdcColor, 0, 0, nWidth, nHeight, hdcMask, 0, 0, DSna
   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcMask, 0, 0, vbSrcAnd
   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcPaint
   BitBlt hdcDest, xDest, yDest, nWidth, nHeight, hdcScnBuffer, 0, 0, vbSrcCopy
     
  'Clear
   DeleteObject SelectObject(hdcColor, hbmColorOld)
   DeleteDC hdcColor
   DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
   DeleteDC hdcScnBuffer
   DeleteObject SelectObject(hdcMask, hbmMaskOld)
   
   DeleteDC hdcMask
   'ReleaseDC 0, hdcScreen
End Sub

Private Function TranslateColor(ByVal OLE_COLOR As Long) As Long
        apiTranslateColor OLE_COLOR, 0, TranslateColor
End Function

Private Sub TimerCheckMouseOut_Timer()
    Dim Pos As POINTAPI
    Dim WFP As Long
    
    GetCursorPos Pos
    WFP = WindowFromPoint(Pos.X, Pos.Y)
    
    If WFP <> Me.hWnd Then
        UserControl_MouseMove -1, 0, -1, -1
        TimerCheckMouseOut.Enabled = False 'kill that timer at once
    End If
End Sub

Private Sub tmrSlide_Timer()
If tmrSlide.Enabled = False Then
    Exit Sub
End If
    If RightSliderStatus = ssDown Then
        lScrollX = lScrollX + 10
            If lScrollX > lScrollWidth Then
                    lScrollX = lScrollWidth
                RightSliderStatus = ssDisabled
                        tmrSlide.Enabled = False
                    Refresh
                    Exit Sub
            End If
                LeftSliderStatus = ssNormal
    Else
        lScrollX = lScrollX - 10
            If lScrollX < 0 Then
                lScrollX = 0
                LeftSliderStatus = ssDisabled
                    tmrSlide.Enabled = False
                Refresh
                Exit Sub
            End If
                RightSliderStatus = ssNormal
    End If

Refresh
    If tmrSlide.Interval > 10 Then tmrSlide.Interval = tmrSlide.Interval - 2
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "UserMode" Then mUserMode = Ambient.UserMode
End Sub

Private Sub UserControl_EnterFocus()
    bInFocus = True
    Refresh
End Sub

Private Sub UserControl_ExitFocus()
    bInFocus = False
    Refresh
End Sub

Private Sub UserControl_Initialize()
    InitGDI
    nScale = GetWindowsDPI
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lTab As Long, lCount As Long
     lTab = getTabOrder(selIndex)
NextTab:
    lCount = lCount + 1
    
    Select Case KeyCode
        Case vbKeyLeft
            If lCount > 2 Then
                lTab = m_TabCount ' + 1
            End If
            If lTab > 0 Then
                lTab = lTab - 1
            End If
    Case vbKeyRight
            If lCount > 2 Then
                lTab = -1 '0
            End If
            If lTab < m_TabCount - 1 Then
                lTab = lTab + 1
            End If
    End Select

    If lTab >= 0 And lTab < m_TabCount And lCount < m_TabCount Then
        If m_Tabs(m_TabOrder(lTab)).Enabled = False Or m_Tabs(m_TabOrder(lTab)).Visible = False Then
            GoTo NextTab:
        End If
    End If
    lTab = m_TabOrder(lTab)
    
    If selIndex <> lTab Then
        RaiseEvent BeforeTabChange(selIndex, lTab)
        RaiseEvent TabClick(lTab)
        SelectedItem = lTab
    End If
    
    RaiseEvent KeyDown(lTab, KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(selIndex, KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(selIndex, KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift <> -1 Then
      cx = ScaleX(X, UserControl.ScaleMode, 3)
      cy = ScaleY(Y, UserControl.ScaleMode, 3)
  Else
      cx = X
      cy = Y
  End If
    
  If tmrSlide.Enabled = True Then
      Exit Sub
  End If
  
Dim bDraw As Boolean
Dim rc As RECT, p As POINTAPI
Dim L As Long, lX As Long, lY As Long, W As Long, H As Long
Dim lTab As Long

    If Shift <> -1 Then
        lX = ScaleX(X, UserControl.ScaleMode, 3)
        lY = ScaleY(Y, UserControl.ScaleMode, 3)
    Else
        lX = X
        lY = Y
    End If
        
    If bSliderShown = True And Button = 1 Then
        Dim lss As ssSliderStatus, rss As ssSliderStatus
            W = (SliderBox.Right - SliderBox.Left) / 2
                lss = ssDisabled
                rss = ssDisabled
        If lX >= SliderBox.Left And lX <= SliderBox.Right And lY >= SliderBox.Top And lY <= SliderBox.Bottom Then
            If lX >= (SliderBox.Right - W) Then
                rss = ssDown
            Else
                lss = ssDown
            End If
        End If
            If lScrollX < lScrollWidth And rss <> ssDisabled And RightSliderStatus <> rss Then
                RightSliderStatus = rss
                bDraw = True
            End If
            If lScrollX > 0 And lss <> ssDisabled And LeftSliderStatus <> lss Then
                LeftSliderStatus = lss
                bDraw = True
            End If
    End If
    
    If bDraw Then
        tmrSlide.Interval = 50
        tmrSlide.Enabled = True
        Refresh
        Exit Sub
    End If
    lTab = HitTest(lX, lY)
    lTabIndex = lTab
    
    If lTab <> -1 Then
            lTab = m_TabOrder(lTab)
        If selIndex <> lTab Then
            RaiseEvent BeforeTabChange(selIndex, lTab)
            RaiseEvent TabClick(lTab)
            SelectedItem = lTab
        End If
    End If
    lTabDragging = False
    RaiseEvent MouseDown(lTab, Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bDraw As Boolean
Dim rc As RECT, p As POINTAPI
Dim c As Long, L As Long, lX As Long, lY As Long, W As Long, H As Long
Dim lTab As Long
        
        If Shift <> -1 Then
            lX = ScaleX(X, UserControl.ScaleMode, 3)
            lY = ScaleY(Y, UserControl.ScaleMode, 3)
        Else
            lX = X
            lY = Y
        End If
        If m_AllowReorder = True And lTabDragging = False And Button = 1 And lTabIndex <> -1 And Abs(cx - lX) > 4 Then
            lTabDragged = lTabIndex
            lTabDragging = True
        End If
    If tmrSlide.Enabled = True And lTabDragging = False Then
        Exit Sub
    End If
        If Button = 0 And lTabDragging = False Then
           TimerCheckMouseOut.Enabled = True
        End If
    lTab = -1
If lTabDragging = True Then
    lTab = HitTest(lX, lY)
End If
    If bSliderShown = True And lTabDragging = False Then
        Dim lss As ssSliderStatus, rss As ssSliderStatus
            W = (SliderBox.Right - SliderBox.Left) / 2
                rss = ssDisabled
                lss = ssDisabled
        If lX >= SliderBox.Left And lX <= SliderBox.Right And lY >= SliderBox.Top And lY <= SliderBox.Bottom Then
            If lX >= (SliderBox.Right - W) Then
                rss = ssHover
            Else
                lss = ssHover
            End If
        End If
            If RightSliderStatus = ssHover And rss = ssDisabled Then
                RightSliderStatus = ssNormal
                bDraw = True
            End If
            If LeftSliderStatus = ssHover And lss = ssDisabled Then
                LeftSliderStatus = ssNormal
                bDraw = True
            End If
            If rss <> ssDisabled And RightSliderStatus = ssNormal Then
                RightSliderStatus = rss
                bDraw = True
            End If
            If lss <> ssDisabled And LeftSliderStatus = ssNormal Then
                LeftSliderStatus = lss
                bDraw = True
            End If
    End If
    
    If bDraw Then
        Refresh
        Exit Sub
    End If

    If lTabDragging = False And Screen.MousePointer <> 0 Then
        Screen.MousePointer = 0
    End If
    
    lTabHover = lTab
    
    If lTabDragging = True Then
        Screen.MousePointer = 5
        If lTab <> -1 Then
            FocusTab m_TabOrder(lTab)
        End If
        If lTabHover <> lTabIndex Then
            lTabIndex = lTab
        End If
            Refresh
    End If
    
    RaiseEvent MouseMove(lTab, Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If lTabDragging = True Then
      Screen.MousePointer = 0
      lTabHover = -1
      lTabDragging = False
      If lTabIndex <> -1 And lTabIndex <> lTabDragged Then
          lTabIndex = m_TabOrder(lTabIndex)
          RaiseEvent TabOrderChanged(m_TabOrder(lTabDragged), lTabIndex)
          lTabIndex = getTabOrder(lTabIndex)
          MoveTab lTabDragged, lTabIndex
          SelectedItem = m_TabOrder(lTabIndex)
      Else
          Refresh
      End If
      lTabDragged = -1
      lTabIndex = -1
      Exit Sub
  End If
  
  If RightSliderStatus = ssDown Or LeftSliderStatus = ssDown Then
      tmrSlide.Enabled = False
      RightSliderStatus = IIf(lScrollX < lScrollWidth, ssNormal, ssDisabled)
      LeftSliderStatus = IIf(lScrollX > 0, ssNormal, ssDisabled)
      Refresh
  End If
  
  RaiseEvent MouseUp(lTabIndex, Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
'mUserControlHdc = hDC
End Sub

Private Sub UserControl_InitProperties()
Dim c As Long
  On Error GoTo 0
  mUserMode = Ambient.UserMode
  mUserControlHwnd = UserControl.hWnd
  mContainerHwnd = UserControl.ContainerHwnd
  mUserControlHdc = UserControl.hDC
  
  pvSubclass
  
  For c = 1 To 4
      AddTab , "Tab " & c
  Next
  m_BackColor1 = m_def_BackColor
  m_BackColor2 = &H8D4214
  m_ForeColor = m_def_ForeColor
  m_ForeColorActive = vbWhite
  m_ColorActive = &H8D4214
  m_ColorDisabled = m_def_ColorDisabled
  m_BorderColor = m_def_BorderColor
  m_FocusRect = False
  m_Angle = 45
  m_Enabled = True
  'Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set mFont = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = mFont
    m_AllowReorder = PropBag.ReadProperty("AllowReorder", False)
    selIndex = PropBag.ReadProperty("SelectedItem", 0)
    m_TabCount = PropBag.ReadProperty("ItemCount", 4)
   
    ReDim m_Tabs(m_TabCount - 1)
    ReDim m_TabOrder(m_TabCount - 1)

    ReDim ctlLst(m_TabCount - 1)
        m_ctlCount = 0
    ReDim m_Ctls(m_ctlCount)
        
    Dim i As Long, z As Long
    Dim mCCount As Long, ctlName As String, ItemMax As Long
    For i = 0 To m_TabCount - 1
        m_TabOrder(i) = PropBag.ReadProperty("TabOrder" & i, i)
        m_Tabs(i).Image = PropBag.ReadProperty("TabIcon" & i, 0)
        m_Tabs(i).Enabled = PropBag.ReadProperty("TabEnabled" & i, True)
        m_Tabs(i).sKey = PropBag.ReadProperty("Key" & i, "")
        m_Tabs(i).sTag = PropBag.ReadProperty("TabTag" & i, "")
        m_Tabs(i).Visible = PropBag.ReadProperty("TabVisible" & i, True)
        
        m_Tabs(i).sCaption = PropBag.ReadProperty("Item(" & i & ").Caption", "Tab " & i + 1)
        mCCount = PropBag.ReadProperty("Item(" & i & ").ControlCount", 0)
                
        For z = 0 To mCCount - 1
                ctlName = PropBag.ReadProperty("Item(" & i & ").Control(" & z & ")", "")
            If ctlName <> "" Then
                ReDim Preserve m_Ctls(m_ctlCount)
                   m_Ctls(m_ctlCount).ctlName = ctlName
                   m_Ctls(m_ctlCount).iTabIndex = i
                m_ctlCount = m_ctlCount + 1
            End If
        Next z
        
    Next i
        
        ItemMax = PropBag.ReadProperty("ItemMax", 0)
    For i = m_TabCount To ItemMax
        mCCount = PropBag.ReadProperty("Item(" & i & ").ControlCount", 0)
        For z = 0 To mCCount - 1
                ctlName = PropBag.ReadProperty("Item(" & i & ").Control(" & z & ")", "")
            If ctlName <> "" Then
                ReDim Preserve m_Ctls(m_ctlCount)
                   m_Ctls(m_ctlCount).ctlName = ctlName
                   m_Ctls(m_ctlCount).iTabIndex = i
                m_ctlCount = m_ctlCount + 1
            End If
        Next z
    Next
        If selIndex > m_TabCount - 1 Then
            selIndex = m_TabCount - 1
        End If
     SelectedItem = selIndex
        
    On Error GoTo 0
  mUserMode = Ambient.UserMode
  mUserControlHwnd = UserControl.hWnd
  mContainerHwnd = UserControl.ContainerHwnd
  mUserControlHdc = UserControl.hDC
    
  m_BackColor1 = PropBag.ReadProperty("BackColor1", m_def_BackColor)
  m_BackColor2 = PropBag.ReadProperty("BackColor2", &H8D4214)
  m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
  m_ForeColorActive = PropBag.ReadProperty("ForeColorActive", vbWhite)
  m_ColorActive = PropBag.ReadProperty("ColorActive", &H8D4214)
  m_ColorDisabled = PropBag.ReadProperty("ColorDisabled", m_def_ColorDisabled)
  m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
  m_FocusRect = PropBag.ReadProperty("FocusRect", False)
  m_TabWidth = PropBag.ReadProperty("ButtonTabWidth", 120)
  m_Angle = PropBag.ReadProperty("AngleGradient", 45)
  m_Enabled = PropBag.ReadProperty("Enabled", True)
  
  pvSubclass
  Refresh

End Sub

Private Sub UserControl_Resize()
'On Error Resume Next
  mUserControlHdc = hDC
    
  Refresh
End Sub

Private Sub UserControl_Terminate()
    TerminateGDI
    pvUnsubclass
Erase m_Tabs
Erase m_TabOrder
    m_TabCount = 0
Erase m_Ctls
    m_ctlCount = 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Font", mFont, Ambient.Font
    PropBag.WriteProperty "AllowReorder", m_AllowReorder, False
    PropBag.WriteProperty "SelectedItem", selIndex, 0
    
    PropBag.WriteProperty "ItemCount", m_TabCount, 4
    
    Dim i As Long, z As Long, c As Long, MaxIndex As Long
    For i = 0 To m_TabCount - 1
        PropBag.WriteProperty "TabOrder" & i, m_TabOrder(i), i
        PropBag.WriteProperty "TabIcon" & i, m_Tabs(i).Image, 0
        PropBag.WriteProperty "TabEnabled" & i, m_Tabs(i).Enabled, True
        PropBag.WriteProperty "TabTag" & i, m_Tabs(i).sTag, ""
        PropBag.WriteProperty "TabKey" & i, m_Tabs(i).sKey, ""
        PropBag.WriteProperty "TabVisible" & i, m_Tabs(i).Visible, True

        PropBag.WriteProperty "Item(" & i & ").Caption", m_Tabs(i).sCaption, "Tab " & i + 1
        c = 0
        For z = 0 To m_ctlCount - 1
            If m_Ctls(z).iTabIndex = i Then
                PropBag.WriteProperty "Item(" & i & ").Control(" & c & ")", m_Ctls(z).ctlName, ""
                c = c + 1
            End If
            If MaxIndex < m_Ctls(z).iTabIndex Then
                MaxIndex = m_Ctls(z).iTabIndex
            End If
        Next z
        PropBag.WriteProperty "Item(" & i & ").ControlCount", c, 0
    Next i
        PropBag.WriteProperty "ItemMax", MaxIndex, 0
    For i = m_TabCount To MaxIndex
            c = 0
        For z = 0 To m_ctlCount - 1
            If m_Ctls(z).iTabIndex = i Then
                PropBag.WriteProperty "Item(" & i & ").Control(" & c & ")", m_Ctls(z).ctlName, ""
                c = c + 1
            End If
        Next z
        PropBag.WriteProperty "Item(" & i & ").ControlCount", c, 0
    Next
  Call PropBag.WriteProperty("BackColor1", m_BackColor1)
  Call PropBag.WriteProperty("BackColor2", m_BackColor2)
  Call PropBag.WriteProperty("ForeColor", m_ForeColor)
  Call PropBag.WriteProperty("ForeColorActive", m_ForeColorActive)
  Call PropBag.WriteProperty("ColorActive", m_ColorActive)
  Call PropBag.WriteProperty("ColorDisabled", m_ColorDisabled)
  Call PropBag.WriteProperty("BorderColor", m_BorderColor)
  Call PropBag.WriteProperty("FocusRect", m_FocusRect)
  Call PropBag.WriteProperty("ButtonTabWidth", m_TabWidth, 0)
  Call PropBag.WriteProperty("AngleGradient", m_Angle)
  Call PropBag.WriteProperty("Enabled", m_Enabled)
End Sub

Public Property Get AngleGradient() As Single
AngleGradient = m_Angle
End Property

Public Property Let AngleGradient(ByVal NewAngle As Single)
m_Angle = NewAngle
PropertyChanged "AngleGradient"
Refresh
End Property

Public Property Get AllowReorder() As Boolean
    AllowReorder = m_AllowReorder
End Property

Public Property Let AllowReorder(ByVal nV As Boolean)
  m_AllowReorder = nV
  PropertyChanged "AllowReorder"
End Property

Public Property Get BackColor1() As OLE_COLOR
  BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor As OLE_COLOR)
  m_BackColor1 = New_BackColor
  PropertyChanged "BackColor1"
  Refresh
End Property

Public Property Get BackColor2() As OLE_COLOR
  BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor As OLE_COLOR)
  m_BackColor2 = New_BackColor
  PropertyChanged "BackColor2"
  Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
  m_BorderColor = New_BorderColor
  PropertyChanged "BorderColor"
  Refresh
End Property

Public Property Get ButtonTabWidth() As Long
  ButtonTabWidth = m_TabWidth
End Property

Public Property Let ButtonTabWidth(ByVal vNewValue As Long)
  m_TabWidth = vNewValue
  PropertyChanged "ButtonTabWidth"
  Refresh
End Property

Public Property Get ColorActive() As OLE_COLOR
  ColorActive = m_ColorActive
End Property

Public Property Let ColorActive(ByVal New_ColorActive As OLE_COLOR)
  m_ColorActive = New_ColorActive
  PropertyChanged "ColorActive"
  Refresh
End Property

Public Property Get ColorDisabled() As OLE_COLOR
  ColorDisabled = m_ColorDisabled
End Property

Public Property Let ColorDisabled(ByVal New_ColorDisabled As OLE_COLOR)
  m_ColorDisabled = New_ColorDisabled
  PropertyChanged "ColorDisabled"
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

Public Property Get FocusRect() As Boolean
    FocusRect = m_FocusRect
End Property

Public Property Let FocusRect(ByVal fR As Boolean)
  m_FocusRect = fR
  PropertyChanged "FocusRect"
End Property

Public Property Get Font() As StdFont
    Set Font = mFont
End Property

Public Property Set Font(ByVal nV As StdFont)
    Set mFont = nV
    PropertyChanged "Font"
    Refresh
End Property

Public Property Get ForeColorActive() As OLE_COLOR
  ForeColorActive = m_ForeColorActive
End Property

Public Property Let ForeColorActive(ByVal New_ForeColorA As OLE_COLOR)
  m_ForeColorActive = New_ForeColorA
  PropertyChanged "ForeColorActive"
  Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  m_ForeColor = New_ForeColor
  PropertyChanged "ForeColor"
  Refresh
End Property

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get ItemCount() As Long
    ItemCount = m_TabCount
End Property

Public Property Let ItemCount(ByVal nV As Long)
Dim i As Long
If nV > m_TabCount Then
    For i = m_TabCount + 1 To nV
        AddTab , "Tab " & (i)
    Next
    Refresh
Else
    If nV > 0 Then
        Dim T As Long
            T = m_TabCount - 1
        For i = T To nV Step -1
            RemoveTab i
        Next
    End If
End If
    PropertyChanged "ItemCount"
End Property

Public Property Get MoveItem() As Long
    MoveItem = getTabOrder(selIndex)
End Property

Public Property Let MoveItem(ByVal mTabIndex As Long)
    If mTabIndex < 0 Or mTabIndex > m_TabCount - 1 Then
        MsgBox "invalid property value", vbCritical
        Exit Property
    End If

    MoveTab getTabOrder(selIndex), mTabIndex
    SelectedItem = m_TabOrder(mTabIndex)
End Property

Public Property Get SelectedItem() As Long
    SelectedItem = selIndex
End Property

Public Property Let SelectedItem(ByVal mTabIndex As Long)
    If mTabIndex < 0 Or mTabIndex > m_TabCount Then
        MsgBox "invalid property value", vbCritical
        Exit Property
    End If
    If mTabIndex > m_TabCount - 1 Then
        mTabIndex = m_TabCount - 1
    End If

    handleControls selIndex, mTabIndex
    selIndex = mTabIndex
    PropertyChanged "SelectedItem"
    Draw
    FocusTab selIndex
    Refresh
End Property

'Public Property Get TabAlignment(ByVal nIndex As Long) As TabCaptionAlignment
'If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
'    TabAlignment = m_Tabs(nIndex).Alignment
'End If
'End Property
'
'Public Property Let TabAlignment(ByVal nIndex As Long, ByVal Value As TabCaptionAlignment)
'If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
'    m_Tabs(nIndex).Alignment = Value
'    Draw
'End If
'End Property

Public Property Get TabCaption() As String
Dim nIndex As Long
    nIndex = selIndex
If nIndex > -1 And nIndex < m_TabCount Then
    TabCaption = m_Tabs(nIndex).sCaption
End If
End Property

Public Property Let TabCaption(ByVal sCaption As String)
Dim nIndex As Long
     nIndex = selIndex
If nIndex > -1 And nIndex < m_TabCount Then
    m_Tabs(nIndex).sCaption = sCaption
    PropertyChanged "TabCaption"
    Refresh
End If
End Property

Public Property Get TabEnabled() As Boolean
Dim nIndex As Long
       nIndex = selIndex
If nIndex > -1 And nIndex < m_TabCount Then
    TabEnabled = m_Tabs(nIndex).Enabled
End If
End Property

Public Property Let TabEnabled(ByVal Newval As Boolean)
Dim nIndex As Long
       nIndex = selIndex
If nIndex > -1 And nIndex < m_TabCount Then
    m_Tabs(nIndex).Enabled = Newval
    PropertyChanged "TabEnabled"
    Refresh
End If
End Property

Public Property Get TabImage(ByVal nIndex As Long) As Long
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    TabImage = m_Tabs(nIndex).Image
End If
End Property

Public Property Let TabImage(ByVal nIndex As Long, ByVal lImage As Long)
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
     m_Tabs(nIndex).Image = lImage
    Draw
End If
End Property

Public Property Get TabKey(ByVal nIndex As Long) As String
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    TabKey = m_Tabs(nIndex).sKey
End If
End Property

Public Property Get TabTag(ByVal nIndex As Long) As String
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    TabTag = m_Tabs(nIndex).sTag
End If
End Property

Public Property Let TabTag(ByVal nIndex As Long, ByVal sTag As String)
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
     m_Tabs(nIndex).sTag = sTag
End If
End Property

Public Property Get TabWidth(ByVal nIndex As Long) As Long
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    TabWidth = m_Tabs(nIndex).Width
End If
End Property

Public Property Let TabWidth(ByVal nIndex As Long, ByVal lWidth As Long)
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
     m_Tabs(nIndex).Width = lWidth
End If
End Property

Public Sub ReplaceColor(ByVal PictureBox As Object, ByVal FromColor As Long, ByVal ToColor As Long)
  If PictureBox.Picture Is Nothing Then Exit Sub
  If PictureBox.Picture.Handle = 0 Then Exit Sub
  Dim WinFromColor As Long, WinToColor As Long, MemAutoRedraw As Boolean
  WinFromColor = WinColor(FromColor)
  WinToColor = WinColor(ToColor)
  With PictureBox
    MemAutoRedraw = .AutoRedraw
    .AutoRedraw = True
    Dim X As Long, Y As Long
    For X = 0 To CInt(.ScaleX(.Picture.Width, vbHimetric, vbPixels))
        For Y = 0 To CInt(.ScaleY(.Picture.Height, vbHimetric, vbPixels))
            If GetPixel(.hDC, X, Y) = WinFromColor Then SetPixel .hDC, X, Y, WinToColor
        Next Y
    Next X
    .Refresh
    .Picture = .Image
    .AutoRedraw = MemAutoRedraw
  End With
End Sub

Public Function WinColor(ByVal Color As Long, Optional ByVal hpal As Long) As Long
If OleTranslateColor(Color, hpal, WinColor) <> 0 Then WinColor = -1
End Function

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

Private Function fRoundCut(ByVal hGraphics As Long, RECT As RECTL, ByVal color1 As Long, ByVal color2 As Long, _
                           ByVal Angulo As Single, ByVal BorderWidth As Long, ByVal BorderColor As Long, _
                           ByVal Round As Long, SideCut As eFlatSide) As Long  '(ByVal hGraphics As Long, Rect As RECTL, ByVal Color1 As Long, ByVal Color2 As Long, ByVal BorderWidth As Long, ByVal BorderColor As Long, Round As Long, Side As sFlatSide) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mPath As Long
    
    
    'GdipCreateSolidFill BackColor, hBrush
    'GdipCreatePen1 BorderColor, &H2 * nScale, &H2, hPen
    If BorderWidth <> 0 Then
      GdipCreatePen1 BorderColor, BorderWidth * nScale, &H2, hPen
    End If
    GdipCreateLineBrushFromRectWithAngleI RECT, color1, color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mPath   '&H0
    
    With RECT
        Select Case SideCut
          Case rUp
              GdipAddPathArcI mPath, .Left, .Top, 1, 1, 180, 90
              GdipAddPathArcI mPath, (.Left + .Width) - 1, .Top, 1, 1, 270, 90
              GdipAddPathArcI mPath, (.Left + .Width) - Round, (.Top + .Height) - Round, Round, Round, 0, 90
              GdipAddPathArcI mPath, .Left, (.Top + .Height) - Round, Round, Round, 90, 90
          Case rBottom
              GdipAddPathArcI mPath, .Left, .Top, Round, Round, 180, 90
              GdipAddPathArcI mPath, (.Left + .Width) - Round, .Top, Round, Round, 270, 90
              GdipAddPathArcI mPath, (.Left + .Width) - 1, (.Top + .Height) - 1, 1, 1, 0, 90
              GdipAddPathArcI mPath, .Left, (.Top + .Height) - 1, 1, 1, 90, 90
          Case rLeft
              GdipAddPathArcI mPath, .Left, .Top, 1, 1, 180, 90
              GdipAddPathArcI mPath, (.Left + .Width) - Round, .Top, Round, Round, 270, 90
              GdipAddPathArcI mPath, (.Left + .Width) - Round, (.Top + .Height) - Round, Round, Round, 0, 90
              GdipAddPathArcI mPath, .Left, (.Top + .Height) - 1, 1, 1, 90, 90
          Case rRight
              GdipAddPathArcI mPath, .Left, .Top, Round, Round, 180, 90
              GdipAddPathArcI mPath, (.Left + .Width) - 1, .Top, 1, 1, 270, 90
              GdipAddPathArcI mPath, (.Left + .Width) - 1, (.Top + .Height) - 1, 1, 1, 0, 90
              GdipAddPathArcI mPath, .Left, (.Top + .Height) - Round, Round, Round, 90, 90
        End Select
    End With
    
    GdipClosePathFigures mPath
    GdipFillPath hGraphics, hBrush, mPath
    GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)

    fRoundCut = mPath
End Function

Private Function fRoundRect(ByVal hGraphics As Long, RECT As RECTL, ByVal color1 As Long, ByVal color2 As Long, _
                            ByVal Angulo As Single, ByVal BorderWidth As Long, ByVal BorderColor As Long, _
                            ByVal Round As Long) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mPath As Long
    
    If BorderWidth <> 0 Then
      GdipCreatePen1 BorderColor, BorderWidth * nScale, &H2, hPen  '&H1 * nScale, &H2, hPen
    End If
    GdipCreateLineBrushFromRectWithAngleI RECT, color1, color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mPath   '&H0
    
    With RECT
        If Round = 0 Then
            GdipDrawRectangleI hGraphics, hPen, .Left, .Top, .Width, .Height
            GdipAddPathLineI mPath, .Left, .Top, .Width, .Top       'Line-Top
            GdipAddPathLineI mPath, .Width, .Top, .Width, .Height   'Line-Left
            GdipAddPathLineI mPath, .Width, .Height, .Left, .Height 'Line-Bottom
            GdipAddPathLineI mPath, .Left, .Height, .Left, .Top     'Line-Right
        Else
            GdipAddPathArcI mPath, .Left, .Top, Round, Round, 180, 90
            GdipAddPathArcI mPath, (.Left + .Width) - Round, .Top, Round, Round, 270, 90
            GdipAddPathArcI mPath, (.Left + .Width) - Round, (.Top + .Height) - Round, Round, Round, 0, 90
            GdipAddPathArcI mPath, .Left, (.Top + .Height) - Round, Round, Round, 90, 90
        End If
    End With
    
    GdipClosePathFigures mPath
    GdipFillPath hGraphics, hBrush, mPath
    GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)

    fRoundRect = mPath
End Function

'Inicia GDI+
Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub

'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub

Private Function AColor(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
  If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
  AColor = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
  Opacity = CByte((Abs(Opacity) / 100) * 255)
  If Opacity < 128 Then
      If Opacity < 0& Then Opacity = 0&
      AColor = AColor Or Opacity * &H1000000
  Else
      If Opacity > 255& Then Opacity = 255&
      AColor = AColor Or (Opacity - 128&) * &H1000000 Or &H80000000
  End If
End Function

Public Function DrawLine(ByVal hGraphics As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, Optional ByVal oColor As OLE_COLOR = vbWhite, Optional ByVal Opacity As Integer = 100, Optional ByVal PenWidth As Integer = 1) As Boolean
    Dim hPen As Long
    
    GdipCreatePen1 AColor(oColor, Opacity), PenWidth * nScale, UnitPixel, hPen
    DrawLine = GdipDrawLineI(hGraphics, hPen, X1 * nScale, y1 * nScale, x2 * nScale, y2 * nScale) = 0
    GdipDeletePen hPen
End Function

