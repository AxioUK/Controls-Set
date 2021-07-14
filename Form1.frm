VERSION 5.00
Object = "*\AAxFramework.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   836
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   11745
      TabIndex        =   48
      Top             =   255
      Width           =   720
   End
   Begin AxFramework.AxGMessageBox AxGMessageBox1 
      Height          =   2340
      Left            =   9435
      TabIndex        =   36
      Top             =   5940
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4128
      Enabled         =   -1  'True
      BackColor1      =   9257492
      BackColor2      =   9257492
      ForeColor       =   16777215
      ForeColor2      =   16777215
      BorderColor     =   14737632
      CornerCurve     =   10
      Filled          =   -1  'True
      ModalOpacity    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorOnFocus    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61294
      IconForeColor   =   12648384
      IcoPaddingX     =   20
      IcoPaddingY     =   35
   End
   Begin AxFramework.AxGButtonLabel cmdMessage1 
      Height          =   420
      Left            =   10725
      TabIndex        =   34
      Top             =   4920
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   741
      Enabled         =   -1  'True
      ForeColor2      =   16777215
      BorderColor     =   4210752
      BorderWidth     =   2
      CornerCurve     =   10
      Filled          =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Msg Modal"
      CaptionX        =   0
      CaptionY        =   0
      ColorOnFocus    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Value           =   0   'False
      OptionButton    =   0   'False
      Clickable       =   -1  'True
   End
   Begin AxFramework.AxGButtonLabel Label1 
      Height          =   270
      Left            =   225
      TabIndex        =   23
      Top             =   4665
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   476
      Enabled         =   -1  'True
      BackColor1      =   9257492
      ForeColor       =   9257492
      ForeColor2      =   16777215
      BorderWidth     =   0
      CornerCurve     =   10
      Filled          =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "AxGButtonLabel1.Value=False"
      CaptionAlignH   =   0
      CaptionX        =   0
      CaptionY        =   0
      ColorOnFocus    =   0
      InitialOpacity  =   100
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Value           =   0   'False
      OptionButton    =   0   'False
      Clickable       =   0   'False
   End
   Begin AxFramework.AxGButtonLabel AxGButtonLabel4 
      Height          =   1395
      Left            =   2865
      TabIndex        =   12
      Top             =   4920
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   2461
      Enabled         =   -1  'True
      ForeColor2      =   16777215
      BorderColor     =   32768
      BorderWidth     =   2
      CornerCurve     =   10
      Filled          =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAngle    =   45
      Caption         =   "AxGButtonLabel4      Not Clickable            Not Filled"
      CaptionX        =   0
      CaptionY        =   0
      ColorOnFocus    =   65280
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Value           =   0   'False
      OptionButton    =   0   'False
      Clickable       =   0   'False
   End
   Begin VB.PictureBox pBack 
      Height          =   450
      Left            =   8460
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   390
      ScaleWidth      =   510
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   135
      Width           =   570
   End
   Begin AxFramework.axGTabControl axGTabControl1 
      Height          =   3525
      Left            =   105
      TabIndex        =   6
      Top             =   210
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   6218
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Item(0).Caption =   "Options"
      Item(0).Control(0)=   "AxGOption2"
      Item(0).Control(1)=   "AxGOption1"
      Item(0).Control(2)=   "AxGOption4"
      Item(0).Control(3)=   "AxGOption3"
      Item(0).Control(4)=   "AxGButtonLabel10"
      Item(0).Control(5)=   "AxGButtonLabel9"
      Item(0).ControlCount=   6
      Item(1).Caption =   "ButtonLabels"
      Item(1).Control(0)=   "AxGButtonLabel2"
      Item(1).Control(1)=   "AxGButtonLabel1"
      Item(1).Control(2)=   "Label1"
      Item(1).Control(3)=   "Check1"
      Item(1).Control(4)=   "Check2"
      Item(1).Control(5)=   "AxGButtonLabel6"
      Item(1).Control(6)=   "AxGButtonLabel5"
      Item(1).ControlCount=   7
      Item(2).Caption =   "ProgressBar"
      Item(2).Control(0)=   "AxGProgBar4"
      Item(2).Control(1)=   "AxGProgBar3"
      Item(2).Control(2)=   "AxGButtonLabel7"
      Item(2).Control(3)=   "AxGButtonLabel8"
      Item(2).ControlCount=   4
      ItemMax         =   2
      BackColor1      =   9257492
      BackColor2      =   9257492
      ForeColor       =   8421504
      ForeColorActive =   16777215
      ColorActive     =   9257492
      ColorDisabled   =   6929919
      BorderColor     =   9257492
      FocusRect       =   0   'False
      ButtonTabWidth  =   120
      AngleGradient   =   45
      Enabled         =   -1  'True
      Begin AxFramework.AxGOption AxGOption7 
         Height          =   390
         Left            =   3060
         TabIndex        =   33
         Top             =   2820
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   688
         Enabled         =   -1  'True
         ForeColor2      =   16777215
         BorderWidth     =   4
         CheckColor      =   16711680
         FillColor       =   0
         FillEnable      =   0   'False
         CornerCurve     =   30
         CheckVisible    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Option3"
         CaptionEnabled  =   -1  'True
         CaptionAlignH   =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionBehavior  =   -1  'True
      End
      Begin AxFramework.AxGOption AxGOption6 
         Height          =   390
         Left            =   1680
         TabIndex        =   32
         Top             =   2820
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   688
         Enabled         =   -1  'True
         ForeColor2      =   16777215
         BorderWidth     =   4
         CheckColor      =   16711680
         FillColor       =   0
         FillEnable      =   0   'False
         CornerCurve     =   30
         CheckVisible    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Option2"
         CaptionEnabled  =   -1  'True
         CaptionAlignH   =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionBehavior  =   -1  'True
      End
      Begin AxFramework.AxGOption AxGOption5 
         Height          =   390
         Left            =   300
         TabIndex        =   31
         Top             =   2820
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   688
         Enabled         =   -1  'True
         ForeColor2      =   16777215
         BorderWidth     =   4
         CheckColor      =   16711680
         FillColor       =   0
         FillEnable      =   0   'False
         CornerCurve     =   30
         CheckVisible    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Option1"
         CaptionEnabled  =   -1  'True
         CaptionAlignH   =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionBehavior  =   -1  'True
      End
      Begin AxFramework.AxGProgBar AxGProgBar4 
         Height          =   1860
         Left            =   -68890
         TabIndex        =   18
         Top             =   780
         Visible         =   0   'False
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   3281
         Enabled         =   -1  'True
         BarColor1       =   128
         BarColor2       =   12632319
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BarAngle        =   45
         BorderColor     =   128
         CornerCurve     =   10
         Filled          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PreCaption      =   ""
         PostCaption     =   "%"
         CaptionPos      =   1
         ColorOnFocus    =   128
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   73
         Orientation     =   0
      End
      Begin AxFramework.AxGProgBar AxGProgBar3 
         Height          =   630
         Left            =   -67105
         TabIndex        =   17
         Top             =   1890
         Visible         =   0   'False
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   1111
         Enabled         =   -1  'True
         BarColor1       =   8421504
         BarColor2       =   14737632
         ForeColor2      =   16777215
         BarAngle        =   45
         BorderColor     =   4210752
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PreCaption      =   ""
         PostCaption     =   "V"
         ColorOnFocus    =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   65
      End
      Begin AxFramework.AxGOption AxGOption4 
         Height          =   390
         Left            =   3405
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   688
         Enabled         =   -1  'True
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BorderColor     =   8421504
         BorderWidth     =   4
         ActiveColor     =   32768
         CheckColor      =   16711680
         FillColor       =   9257492
         FillEnable      =   -1  'True
         CornerCurve     =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CheckButton"
         CaptionEnabled  =   -1  'True
         CaptionAlignH   =   0
         Transparent     =   0   'False
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionBehavior  =   0   'False
      End
      Begin AxFramework.AxGOption AxGOption3 
         Height          =   390
         Left            =   3420
         TabIndex        =   15
         Top             =   2025
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   688
         Enabled         =   -1  'True
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BorderColor     =   8421504
         BorderWidth     =   4
         ActiveColor     =   32768
         CheckColor      =   16711680
         FillColor       =   9257492
         FillEnable      =   -1  'True
         CornerCurve     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CheckButton"
         CaptionEnabled  =   -1  'True
         CaptionAlignH   =   0
         Transparent     =   0   'False
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionBehavior  =   0   'False
      End
      Begin AxFramework.AxGButtonLabel AxGButtonLabel6 
         Height          =   1395
         Left            =   -67495
         TabIndex        =   14
         Top             =   810
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2461
         Enabled         =   -1  'True
         ForeColor2      =   16777215
         BorderColor     =   32768
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAngle    =   45
         Caption         =   "AxGButtonLabel4      Clickable & Filled"
         CaptionX        =   0
         CaptionY        =   0
         ColorOnFocus    =   65280
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionButton    =   0   'False
         Clickable       =   -1  'True
      End
      Begin AxFramework.AxGButtonLabel AxGButtonLabel5 
         Height          =   1395
         Left            =   -69535
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2461
         Enabled         =   -1  'True
         ForeColor2      =   16777215
         BorderColor     =   32768
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAngle    =   45
         Caption         =   "AxGButtonLabel4      Not Clickable            Not Filled"
         CaptionX        =   0
         CaptionY        =   0
         ColorOnFocus    =   65280
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionButton    =   0   'False
         Clickable       =   0   'False
      End
      Begin AxFramework.AxGOption AxGOption2 
         Height          =   390
         Left            =   510
         TabIndex        =   8
         Top             =   1995
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   688
         Enabled         =   -1  'True
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BorderColor     =   8421504
         BorderWidth     =   4
         ActiveColor     =   32768
         CheckColor      =   16711680
         FillColor       =   0
         FillEnable      =   0   'False
         CornerCurve     =   10
         CheckVisible    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CheckButton"
         CaptionEnabled  =   -1  'True
         CaptionAlignH   =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionBehavior  =   0   'False
      End
      Begin AxFramework.AxGOption AxGOption1 
         Height          =   390
         Left            =   510
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   688
         Enabled         =   -1  'True
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BorderColor     =   8421504
         BorderWidth     =   4
         ActiveColor     =   32768
         CheckColor      =   16711680
         FillColor       =   0
         FillEnable      =   0   'False
         CornerCurve     =   30
         CheckVisible    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CheckButton"
         CaptionEnabled  =   -1  'True
         CaptionAlignH   =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionBehavior  =   0   'False
      End
      Begin AxFramework.AxGButtonLabel AxGButtonLabel7 
         Height          =   2085
         Left            =   -69340
         TabIndex        =   19
         Top             =   675
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   3678
         Enabled         =   -1  'True
         ForeColor       =   12648447
         ForeColor2      =   16777215
         BorderColor     =   14737632
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAngle    =   270
         Caption         =   "Vertical Bar"
         CaptionX        =   0
         CaptionY        =   -28
         ColorOnFocus    =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionButton    =   0   'False
         Clickable       =   0   'False
      End
      Begin AxFramework.AxGButtonLabel AxGButtonLabel8 
         Height          =   1050
         Left            =   -67240
         TabIndex        =   20
         Top             =   1605
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1852
         Enabled         =   -1  'True
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BorderColor     =   14737632
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Horizontal Bar"
         CaptionAlignV   =   0
         CaptionAlignH   =   0
         CaptionX        =   5
         CaptionY        =   0
         ColorOnFocus    =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionButton    =   0   'False
         Clickable       =   0   'False
      End
      Begin AxFramework.AxGButtonLabel AxGButtonLabel10 
         Height          =   2070
         Left            =   2940
         TabIndex        =   22
         Top             =   555
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   3651
         Enabled         =   -1  'True
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BorderColor     =   14737632
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Transparent=False  All within     the Control are Clickable, No Check Visible"
         CaptionAlignV   =   0
         CaptionAlignH   =   0
         CaptionX        =   5
         CaptionY        =   0
         ColorOnFocus    =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionButton    =   0   'False
         Clickable       =   0   'False
      End
      Begin AxFramework.AxGButtonLabel AxGButtonLabel9 
         Height          =   2070
         Left            =   150
         TabIndex        =   21
         Top             =   555
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   3651
         Enabled         =   -1  'True
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BorderColor     =   14737632
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Transparent=True  Only Text and Option are Clickable, with Check Visible"
         CaptionAlignV   =   0
         CaptionAlignH   =   0
         CaptionX        =   5
         CaptionY        =   0
         ColorOnFocus    =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionButton    =   0   'False
         Clickable       =   0   'False
      End
   End
   Begin AxFramework.AxGFrame AxGFrame1 
      Height          =   2595
      Left            =   6105
      TabIndex        =   5
      Top             =   1275
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   4577
      Enabled         =   -1  'True
      BackColor1      =   9257492
      BackColor2      =   9257492
      ForeColor       =   16777215
      BorderColor     =   8421504
      CornerCurve     =   10
      Filled          =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionX        =   0
      CaptionY        =   0
      CaptionBoxLeft  =   -15
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   4210752
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Begin AxFramework.AxGButtonLabel AxGButtonLabel14 
         Height          =   1200
         Left            =   720
         TabIndex        =   45
         Top             =   645
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2117
         Enabled         =   -1  'True
         ForeColor2      =   16777215
         BorderColor     =   32768
         BorderWidth     =   0
         CornerCurve     =   10
         Filled          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAngle    =   35
         Caption         =   "No usar     Filled=False y Transparent=True     a la vez"
         CaptionX        =   0
         CaptionY        =   0
         ColorOnFocus    =   65280
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionButton    =   0   'False
         Clickable       =   0   'False
      End
      Begin AxFramework.AxGButtonLabel AxGButtonLabel11 
         Height          =   945
         Index           =   1
         Left            =   2145
         TabIndex        =   29
         Top             =   795
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   1667
         Enabled         =   -1  'True
         BackColor1      =   9257492
         BackColor2      =   9257492
         ForeColor       =   14737632
         ForeColor2      =   16777215
         BorderColor     =   4194304
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAngle    =   90
         Caption         =   "Right"
         CaptionAlignH   =   2
         CaptionX        =   8
         CaptionY        =   -2
         ColorOnFocus    =   14737632
         EffectFading    =   -1  'True
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IcoFont"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconCharCode    =   60013
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   -2
         Value           =   0   'False
         OptionButton    =   -1  'True
         Clickable       =   -1  'True
      End
      Begin AxFramework.AxGButtonLabel AxGButtonLabel11 
         Height          =   945
         Index           =   0
         Left            =   345
         TabIndex        =   28
         Top             =   795
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   1667
         Enabled         =   -1  'True
         BackColor1      =   9257492
         BackColor2      =   9257492
         ForeColor       =   14737632
         ForeColor2      =   16777215
         BorderColor     =   4194304
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAngle    =   90
         Caption         =   "Left"
         CaptionAlignH   =   2
         CaptionX        =   5
         CaptionY        =   -2
         ColorOnFocus    =   14737632
         EffectFading    =   -1  'True
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IcoFont"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconCharCode    =   60012
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   -2
         Value           =   0   'False
         OptionButton    =   -1  'True
         Clickable       =   -1  'True
      End
      Begin AxFramework.AxGButtonLabel AxGButtonLabel11 
         Height          =   435
         Index           =   3
         Left            =   1005
         TabIndex        =   27
         Top             =   1800
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   767
         Enabled         =   -1  'True
         BackColor1      =   9257492
         BackColor2      =   9257492
         ForeColor       =   14737632
         ForeColor2      =   16777215
         BorderColor     =   4194304
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Bottom"
         CaptionAlignH   =   2
         CaptionX        =   -2
         CaptionY        =   0
         ColorOnFocus    =   14737632
         EffectFading    =   -1  'True
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IcoFont"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconCharCode    =   60011
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   -2
         Value           =   0   'False
         OptionButton    =   -1  'True
         Clickable       =   -1  'True
      End
      Begin AxFramework.AxGButtonLabel AxGButtonLabel11 
         Height          =   435
         Index           =   2
         Left            =   1005
         TabIndex        =   26
         Top             =   315
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   767
         Enabled         =   -1  'True
         BackColor1      =   9257492
         BackColor2      =   9257492
         ForeColor       =   14737632
         ForeColor2      =   16777215
         BorderColor     =   4194304
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Top"
         CaptionAlignH   =   2
         CaptionX        =   -10
         CaptionY        =   0
         ColorOnFocus    =   14737632
         EffectFading    =   -1  'True
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IcoFont"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconCharCode    =   60014
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   -2
         Value           =   -1  'True
         OptionButton    =   -1  'True
         Clickable       =   -1  'True
      End
   End
   Begin AxFramework.AxGInfoPanel AxGInfoPanel1 
      Height          =   3150
      Left            =   9285
      TabIndex        =   4
      Top             =   1215
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   5556
      Enabled         =   -1  'True
      BackColor1      =   9257492
      BackColor2      =   9257492
      ActiveColor     =   16777215
      BorderColor     =   8421504
      CornerCurve     =   20
      CrossVisible    =   -1  'True
      PinVisible      =   -1  'True
      Moveable        =   -1  'True
      LineOrientation =   0
      Line1           =   -1  'True
      Line2           =   -1  'True
      Line1Pos        =   25
      Line2Pos        =   38
      RollCaption     =   "Test Roll"
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1Color   =   16777215
      Caption1        =   "AxGInfoPanel Caption1"
      Caption1Enabled =   -1  'True
      Caption1Agle    =   270
      Caption1X       =   10
      Caption1Y       =   0
      Caption1AlignV  =   0
      Caption1AlignH  =   1
      Caption1Opacity =   100
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2Color   =   16777215
      Caption2        =   "AxGInfoPanel Caption2"
      Caption2Enabled =   -1  'True
      Caption2Angle   =   0
      Caption2X       =   0
      Caption2Y       =   -30
      Caption2AlignV  =   0
      Caption2AlignH  =   1
      Caption2Opacity =   100
      BorderColorOnFocus=   0
      EffectFading    =   4
      InitialOpacity  =   85
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon1CharCode   =   61389
      Icon1ForeColor  =   12632256
      Icon1PaddingX   =   30
      Icon1PaddingY   =   0
      Icon2CharCode   =   61390
      Icon2ForeColor  =   12632256
      Icon2PaddingX   =   130
      Icon2PaddingY   =   40
      Begin AxFramework.AxGButtonLabel AxGButtonLabel12 
         Height          =   1590
         Left            =   1020
         TabIndex        =   30
         Top             =   1350
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   2805
         Enabled         =   -1  'True
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BorderColor     =   16777215
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAngle    =   30
         Caption         =   ""
         CaptionX        =   0
         CaptionY        =   -5
         ColorOnFocus    =   65280
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionButton    =   0   'False
         Clickable       =   0   'False
      End
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   11070
      TabIndex        =   3
      Top             =   255
      Width           =   660
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset Back"
      Height          =   300
      Left            =   9180
      TabIndex        =   2
      Top             =   660
      Width           =   1740
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Scalemode 3-Twip"
      Height          =   300
      Left            =   9180
      TabIndex        =   1
      Top             =   360
      Width           =   1740
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Scalemode 3-Pixel"
      Height          =   300
      Left            =   9180
      TabIndex        =   0
      Top             =   45
      Width           =   1740
   End
   Begin AxFramework.AxGButtonLabel AxGButtonLabel2 
      Height          =   600
      Left            =   285
      TabIndex        =   10
      Top             =   5700
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1058
      Enabled         =   -1  'True
      BackColor1      =   9257492
      BackColor2      =   9257492
      ForeColor2      =   16777215
      BorderColor     =   0
      BorderWidth     =   2
      CornerCurve     =   10
      Filled          =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAngle    =   350
      Caption         =   "AxGButtonLabel1"
      CaptionX        =   0
      CaptionY        =   0
      ColorOnFocus    =   16711680
      EffectFading    =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61094
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Value           =   0   'False
      OptionButton    =   0   'False
      Clickable       =   0   'False
   End
   Begin AxFramework.AxGButtonLabel AxGButtonLabel1 
      Height          =   600
      Left            =   285
      TabIndex        =   11
      Top             =   4995
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1058
      Enabled         =   -1  'True
      BackColor1      =   9257492
      BackColor2      =   9257492
      ForeColor       =   14737632
      ForeColor2      =   16777215
      BorderColor     =   4194304
      BorderWidth     =   2
      CornerCurve     =   10
      Filled          =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionX        =   0
      CaptionY        =   0
      ColorOnFocus    =   16711680
      EffectFading    =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Value           =   0   'False
      OptionButton    =   0   'False
      Clickable       =   0   'False
   End
   Begin AxFramework.AxGOption Check1 
      Height          =   300
      Left            =   315
      TabIndex        =   24
      Top             =   3990
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      Enabled         =   -1  'True
      BackColor1      =   16777215
      BackColor2      =   16777215
      ForeColor       =   9257492
      BorderColor     =   8421504
      BorderWidth     =   4
      CheckColor      =   16711680
      FillColor       =   0
      FillEnable      =   0   'False
      CornerCurve     =   30
      CheckVisible    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Clickable ?"
      CaptionEnabled  =   -1  'True
      CaptionAlignH   =   0
      Transparent     =   0   'False
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Value           =   0   'False
      OptionBehavior  =   0   'False
   End
   Begin AxFramework.AxGOption Check2 
      Height          =   300
      Left            =   315
      TabIndex        =   25
      Top             =   4290
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      Enabled         =   -1  'True
      BackColor1      =   16777215
      BackColor2      =   16777215
      ForeColor       =   9257492
      BorderColor     =   8421504
      BorderWidth     =   4
      CheckColor      =   16711680
      FillColor       =   0
      FillEnable      =   0   'False
      CornerCurve     =   30
      CheckVisible    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OptionButton ?"
      CaptionEnabled  =   -1  'True
      CaptionAlignH   =   0
      Transparent     =   0   'False
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Value           =   0   'False
      OptionBehavior  =   0   'False
   End
   Begin AxFramework.AxGButtonLabel cmdMessage2 
      Height          =   420
      Left            =   10725
      TabIndex        =   35
      Top             =   5415
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   741
      Enabled         =   -1  'True
      ForeColor2      =   16777215
      BorderColor     =   4210752
      BorderWidth     =   2
      CornerCurve     =   10
      Filled          =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Msg No Modal"
      CaptionX        =   0
      CaptionY        =   0
      ColorOnFocus    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Value           =   0   'False
      OptionButton    =   0   'False
      Clickable       =   -1  'True
   End
   Begin AxFramework.AxGFrame AxGFrame2 
      Height          =   2415
      Left            =   5025
      TabIndex        =   37
      Top             =   4005
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   4260
      Enabled         =   -1  'True
      BackColor1      =   9257492
      BackColor2      =   9257492
      ForeColor       =   16777215
      ForeColor2      =   16777215
      BorderColor     =   9257492
      CornerCurve     =   10
      Filled          =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ProgressBar"
      CaptionX        =   0
      CaptionY        =   0
      CaptionBoxLeft  =   150
      CaptionBoxWidth =   70
      ColorOnFocus    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Begin AxFramework.AxGButtonLabel AxGButtonLabel13 
         Height          =   480
         Left            =   1230
         TabIndex        =   44
         Top             =   975
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   847
         Enabled         =   -1  'True
         BackColor1      =   9257492
         BackColor2      =   9257492
         ForeColor       =   16777215
         ForeColor2      =   65280
         BackAngle       =   45
         BorderColor     =   12632256
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Enabled?"
         CaptionX        =   0
         CaptionY        =   0
         ColorOnFocus    =   16776960
         EffectFading    =   -1  'True
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionButton    =   0   'False
         Clickable       =   -1  'True
      End
      Begin AxFramework.AxGOption axgPos 
         Height          =   390
         Index           =   2
         Left            =   2430
         TabIndex        =   43
         Top             =   1155
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   688
         Enabled         =   -1  'True
         BackColor2      =   9257492
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BorderColor     =   12632256
         BorderWidth     =   4
         CheckColor      =   16777215
         FillColor       =   9257492
         FillEnable      =   0   'False
         CornerCurve     =   30
         CheckVisible    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Center"
         CaptionEnabled  =   -1  'True
         CaptionAlignH   =   0
         EffectFading    =   -1  'True
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionBehavior  =   -1  'True
      End
      Begin AxFramework.AxGOption axgPos 
         Height          =   390
         Index           =   1
         Left            =   2430
         TabIndex        =   42
         Top             =   795
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   688
         Enabled         =   -1  'True
         BackColor2      =   9257492
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BorderColor     =   12632256
         BorderWidth     =   4
         CheckColor      =   16777215
         FillColor       =   9257492
         FillEnable      =   0   'False
         CornerCurve     =   30
         CheckVisible    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TopValue"
         CaptionEnabled  =   -1  'True
         CaptionAlignH   =   0
         EffectFading    =   -1  'True
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   -1  'True
         OptionBehavior  =   -1  'True
      End
      Begin AxFramework.AxGOption axgPos 
         Height          =   390
         Index           =   0
         Left            =   2430
         TabIndex        =   41
         Top             =   435
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   688
         Enabled         =   -1  'True
         BackColor2      =   9257492
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BorderColor     =   12632256
         BorderWidth     =   4
         CheckColor      =   16777215
         FillColor       =   9257492
         FillEnable      =   0   'False
         CornerCurve     =   30
         CheckVisible    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Start"
         CaptionEnabled  =   -1  'True
         CaptionAlignH   =   0
         EffectFading    =   -1  'True
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionBehavior  =   -1  'True
      End
      Begin AxFramework.AxGProgBar AxGProgBar1 
         Height          =   1860
         Left            =   435
         TabIndex        =   40
         Top             =   390
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   3281
         Enabled         =   -1  'True
         BarColor1       =   9257492
         BarColor2       =   9257492
         ForeColor       =   16777215
         ForeColor2      =   16777215
         BarAngle        =   45
         BorderColor     =   12632256
         CornerCurve     =   10
         Filled          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PreCaption      =   ""
         PostCaption     =   "%"
         ColorOnFocus    =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   99
         Orientation     =   0
      End
      Begin VB.Timer TimerBar 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1785
         Top             =   1005
      End
      Begin AxFramework.AxGProgBar AxGProgBar2 
         Height          =   570
         Left            =   1245
         TabIndex        =   39
         Top             =   1680
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1005
         Enabled         =   -1  'True
         BarColor1       =   9257492
         BarColor2       =   9257492
         ForeColor       =   16776960
         ForeColor2      =   16777215
         BarAngle        =   45
         BorderColor     =   12632256
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PreCaption      =   "Value"
         PostCaption     =   "%"
         ColorOnFocus    =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   45
      End
      Begin AxFramework.AxGButtonLabel AxGButtonLabel3 
         Height          =   480
         Left            =   1230
         TabIndex        =   38
         Top             =   405
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   847
         Enabled         =   -1  'True
         BackColor1      =   9257492
         BackColor2      =   9257492
         ForeColor       =   16777215
         ForeColor2      =   65280
         BackAngle       =   45
         BorderColor     =   12632256
         BorderWidth     =   2
         CornerCurve     =   10
         Filled          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Start"
         CaptionX        =   0
         CaptionY        =   0
         ColorOnFocus    =   16776960
         EffectFading    =   -1  'True
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IcoPaddingX     =   0
         IcoPaddingY     =   0
         Value           =   0   'False
         OptionButton    =   0   'False
         Clickable       =   -1  'True
      End
   End
   Begin AxFramework.AxGButtonLabel FrameOp1 
      Height          =   420
      Left            =   6600
      TabIndex        =   46
      Top             =   105
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   741
      Enabled         =   -1  'True
      ForeColor       =   16777215
      ForeColor2      =   16777215
      BorderColor     =   4210752
      BorderWidth     =   2
      CornerCurve     =   10
      Filled          =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Filled?"
      CaptionX        =   0
      CaptionY        =   0
      ColorOnFocus    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Value           =   0   'False
      OptionButton    =   0   'False
      Clickable       =   -1  'True
   End
   Begin AxFramework.AxGButtonLabel FrameOp2 
      Height          =   420
      Left            =   6600
      TabIndex        =   47
      Top             =   600
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   741
      Enabled         =   -1  'True
      ForeColor       =   16777215
      ForeColor2      =   16777215
      BorderColor     =   4210752
      BorderWidth     =   2
      CornerCurve     =   10
      Filled          =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Transparent?"
      CaptionX        =   0
      CaptionY        =   0
      ColorOnFocus    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Value           =   0   'False
      OptionButton    =   0   'False
      Clickable       =   -1  'True
   End
   Begin AxFramework.AxGButtonLabel AxGButtonLabel15 
      Height          =   510
      Left            =   10245
      TabIndex        =   50
      Top             =   2415
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   900
      Enabled         =   -1  'True
      ForeColor       =   16777215
      ForeColor2      =   16777215
      BorderColor     =   4210752
      BorderWidth     =   2
      CornerCurve     =   10
      Filled          =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "InfoPanel Visible"
      CaptionX        =   0
      CaptionY        =   0
      ColorOnFocus    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IcoPaddingX     =   0
      IcoPaddingY     =   0
      Value           =   0   'False
      OptionButton    =   0   'False
      Clickable       =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PIN       CROSS"
      Height          =   195
      Left            =   11250
      TabIndex        =   49
      Top             =   75
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim V As Integer

Private Sub AxGButtonLabel11_Click(Index As Integer)
AxGFrame1.CaptionPos = Index
End Sub

Private Sub AxGButtonLabel13_Click()
AxGProgBar1.Enabled = Not AxGProgBar1.Enabled
End Sub

Private Sub AxGButtonLabel1_ChangeValue(ByVal Value As Boolean)
Label1.Caption = "AxGButtonLabel1.Value=" & AxGButtonLabel1.Value
End Sub

Private Sub AxGButtonLabel15_Click()
AxGInfoPanel1.Visible = True
 AxGButtonLabel15.Visible = False
End Sub

Private Sub AxGButtonLabel3_Click()
TimerBar.Enabled = Not TimerBar.Enabled
End Sub

Private Sub AxGInfoPanel1_CrossClick()
AxGButtonLabel15.Visible = True
End Sub

Private Sub AxGInfoPanel1_DrawString()
With AxGInfoPanel1
  .AddString .hDC, "Test AddString w/Angle", 60, 50, 150, 20, 45, Me.Font, vbWhite, 60, eCenter, eMiddle, False
End With
End Sub

Private Sub AxGMessageBox1_ButtonClick(ButtonPress As AxFramework.ButtonResult)
  If ButtonPress = vrOK Then MsgBox "Aceptar presionado"
  If ButtonPress = vrCancel Then MsgBox "Cancelar presionado"
  
End Sub

Private Sub axgPos_Click(Index As Integer)
AxGProgBar1.CaptionPos = Index
AxGProgBar2.CaptionPos = Index
End Sub

Private Sub AxGProgBar2_ChangeProgress(ByVal Value As Long)
AxGProgBar3.Value = Value
End Sub

Private Sub Check1_Click()
AxGButtonLabel1.Clickable = Check1.Value
AxGButtonLabel2.Clickable = Check1.Value
End Sub

Private Sub Check2_Click()
AxGButtonLabel1.OptionButton = Check2.Value
AxGButtonLabel2.OptionButton = Check2.Value
End Sub

Private Sub cmdMessage1_Click()
With AxGMessageBox1
  .Top = 100  'Pixels
  .Left = 100  'Pixels
  .Modal = True
  .Show Me
End With
End Sub

Private Sub cmdMessage2_Click()
With AxGMessageBox1
  .Top = ((Me.Height - .Height) / 2) / Screen.TwipsPerPixelY
  .Left = ((Me.Width - .Width) / 2) / Screen.TwipsPerPixelX
  .Modal = False
  .Show
End With
End Sub

Private Sub Command2_Click()
Me.ScaleMode = 3
End Sub

Private Sub Command3_Click()
Me.ScaleMode = 1
End Sub

Private Sub Command4_Click()
Set Me.Picture = pBack.Picture
End Sub

Private Sub Form_Load()
V = 0
With List1
    .AddItem "Left", 0
    .AddItem "Right", 1
    .AddItem "Top", 2
    .AddItem "Bottom", 3
End With
With List2
    .AddItem "TopRight", 0
    .AddItem "BottomRight", 1
    .AddItem "TopLeft", 2
    .AddItem "cBottomLeft", 3
End With

AxGButtonLabel12.Caption = "InfoPanel Container," & vbLf & "Double Caption," & vbLf & "Double IconChar," & vbLf & "RolledCaption"
End Sub

Private Sub FrameOp1_Click()
AxGFrame1.Filled = Not AxGFrame1.Filled
End Sub

Private Sub FrameOp2_Click()
AxGFrame1.Transparent = Not AxGFrame1.Transparent
End Sub

Private Sub List1_Click()
AxGInfoPanel1.PinPosition = List1.ListIndex
End Sub

Private Sub List2_Click()
AxGInfoPanel1.CrossPosition = List2.ListIndex
End Sub

Private Sub TimerBar_Timer()
If V = 100 Then
  V = 0
Else
  V = V + 1
End If
AxGProgBar2.Value = V

If AxGProgBar1.Value = 100 Then
  AxGProgBar1.Value = 0
Else
  AxGProgBar1.Value = AxGProgBar1.Value + 1
End If

End Sub
