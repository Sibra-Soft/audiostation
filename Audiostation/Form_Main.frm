VERSION 5.00
Object = "{0A362340-2E5E-11D3-85BF-00105AC8B715}#1.0#0"; "isDigitalLibrary.ocx"
Object = "{C5412DA5-2E2F-11D3-85BF-00105AC8B715}#1.0#0"; "isAnalogLibrary.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiostation"
   ClientHeight    =   9105
   ClientLeft      =   4560
   ClientTop       =   1500
   ClientWidth     =   12750
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9105
   ScaleWidth      =   12750
   Begin VB.PictureBox ElementOff 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   780
      Index           =   5
      Left            =   120
      Picture         =   "Form_Main.frx":088B
      ScaleHeight     =   780
      ScaleWidth      =   9615
      TabIndex        =   118
      Top             =   8400
      Width           =   9615
   End
   Begin VB.PictureBox Element 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   6
      Left            =   120
      Picture         =   "Form_Main.frx":824D
      ScaleHeight     =   855
      ScaleWidth      =   9615
      TabIndex        =   143
      Top             =   8400
      Visible         =   0   'False
      Width           =   9615
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   2880
         ScaleHeight     =   435
         ScaleWidth      =   6555
         TabIndex        =   144
         Top             =   55
         Width           =   6615
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright � 2009 - 2021 Sibra-Soft Software Production"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   45
            TabIndex        =   146
            Top             =   210
            Width           =   4875
         End
         Begin VB.Label lbl_version 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   45
            TabIndex        =   145
            Top             =   20
            Width           =   180
         End
      End
      Begin isDigitalLibrary.iSevenSegmentClockX Digit_Clock 
         Height          =   495
         Left            =   80
         TabIndex        =   147
         Top             =   55
         Width           =   2775
         Time            =   0
         ShowSeconds     =   -1  'True
         ShowHours       =   -1  'True
         HourStyle       =   0
         AutoSize        =   -1  'True
         DigitSpacing    =   6
         SegmentMargin   =   5
         SegmentColor    =   16777215
         SegmentSeperation=   1
         SegmentSize     =   1
         ShowOffSegments =   -1  'True
         PowerOff        =   0   'False
         BackGroundColor =   0
         BorderStyle     =   2
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         Hours           =   0
         Minutes         =   0
         Seconds         =   0
         CountDirection  =   0
         CountTimerEnabled=   0   'False
         SegmentOffColor =   8421504
         AutoSegmentOffColor=   -1  'True
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   185
         Object.Height          =   33
         OPCItemCount    =   0
      End
   End
   Begin VB.PictureBox ElementOff 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   2
      Left            =   120
      Picture         =   "Form_Main.frx":1F14B
      ScaleHeight     =   1500
      ScaleWidth      =   9615
      TabIndex        =   115
      Top             =   2340
      Width           =   9615
   End
   Begin VB.PictureBox ElementOff 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   1
      Left            =   120
      Picture         =   "Form_Main.frx":2EF8D
      ScaleHeight     =   1500
      ScaleWidth      =   9615
      TabIndex        =   114
      Top             =   840
      Width           =   9615
   End
   Begin VB.PictureBox ElementOff 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3060
      Index           =   4
      Left            =   120
      Picture         =   "Form_Main.frx":3EDCF
      ScaleHeight     =   3060
      ScaleWidth      =   9615
      TabIndex        =   117
      Top             =   5330
      Width           =   9615
   End
   Begin VB.PictureBox ElementOff 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   3
      Left            =   120
      Picture         =   "Form_Main.frx":5F291
      ScaleHeight     =   1500
      ScaleWidth      =   9615
      TabIndex        =   116
      Top             =   3840
      Width           =   9615
   End
   Begin VB.Timer Trm_Lights 
      Interval        =   500
      Left            =   10920
      Top             =   4920
   End
   Begin VB.Timer Trm_Animation 
      Enabled         =   0   'False
      Interval        =   110
      Left            =   11400
      Tag             =   "1"
      Top             =   2520
   End
   Begin VB.Timer Trm_Main 
      Interval        =   50
      Left            =   10920
      Tag             =   "1"
      Top             =   4440
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   10200
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   53
      ImageHeight     =   42
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":6EE53
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":708E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":72377
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":73E09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":7589B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   10200
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":7732D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":7F777
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":8F7BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":9F7FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":AF843
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":BF887
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":CF8CB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Trm_Midi_Play 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10920
      Tag             =   "0"
      Top             =   3960
   End
   Begin VB.Timer Trm_Floppy_Drive_Light 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10920
      Top             =   3480
   End
   Begin VB.Timer Trm_Lights_Midi 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10920
      Top             =   3000
   End
   Begin VB.Timer Trm_CD_Animation 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   10920
      Tag             =   "1"
      Top             =   2520
   End
   Begin VB.Timer Trm_VU 
      Interval        =   25
      Left            =   10920
      Top             =   600
   End
   Begin VB.Timer Trm_Check_File 
      Interval        =   10
      Left            =   10920
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   10200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":DF90F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Element 
      BackColor       =   &H00000000&
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
      Height          =   1575
      Index           =   4
      Left            =   120
      Picture         =   "Form_Main.frx":DFEA9
      ScaleHeight     =   1575
      ScaleWidth      =   9735
      TabIndex        =   1
      Top             =   5330
      Visible         =   0   'False
      Width           =   9735
      Begin isAnalogLibrary.iLabelX ILaMaster 
         Height          =   195
         Left            =   120
         TabIndex        =   131
         Top             =   1192
         Width           =   1215
         AutoSize        =   0   'False
         Alignment       =   0
         BorderStyle     =   0
         Caption         =   "Master"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterMarginLeft =   0
         OuterMarginTop  =   0
         OuterMarginRight=   0
         OuterMarginBottom=   0
         ShadowShow      =   -1  'True
         ShadowXOffset   =   -1
         ShadowYOffset   =   -1
         ShadowColor     =   16777215
         BackGroundColor =   12632256
         UpdateFrameRate =   60
         Object.Visible         =   -1  'True
         FontColor       =   -16777208
         Transparent     =   -1  'True
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Enabled         =   -1  'True
         Object.Width           =   81
         Object.Height          =   13
         WordWrap        =   0   'False
         OPCItemCount    =   0
      End
      Begin isAnalogLibrary.iLabelX iLabelX2 
         Height          =   210
         Left            =   3210
         TabIndex        =   110
         Top             =   1185
         Width           =   900
         AutoSize        =   0   'False
         Alignment       =   0
         BorderStyle     =   0
         Caption         =   "REC"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterMarginLeft =   0
         OuterMarginTop  =   0
         OuterMarginRight=   0
         OuterMarginBottom=   0
         ShadowShow      =   -1  'True
         ShadowXOffset   =   -1
         ShadowYOffset   =   -1
         ShadowColor     =   16777215
         BackGroundColor =   -16777201
         UpdateFrameRate =   60
         Object.Visible         =   -1  'True
         FontColor       =   -16777208
         Transparent     =   -1  'True
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Enabled         =   -1  'True
         Object.Width           =   60
         Object.Height          =   14
         WordWrap        =   0   'False
         OPCItemCount    =   0
      End
      Begin isAnalogLibrary.iLabelX iLabelX5 
         Height          =   210
         Left            =   6480
         TabIndex        =   113
         Top             =   1192
         Width           =   855
         AutoSize        =   0   'False
         Alignment       =   0
         BorderStyle     =   0
         Caption         =   "MIDI"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterMarginLeft =   0
         OuterMarginTop  =   0
         OuterMarginRight=   0
         OuterMarginBottom=   0
         ShadowShow      =   -1  'True
         ShadowXOffset   =   -1
         ShadowYOffset   =   -1
         ShadowColor     =   16777215
         BackGroundColor =   -16777201
         UpdateFrameRate =   60
         Object.Visible         =   -1  'True
         FontColor       =   -16777208
         Transparent     =   -1  'True
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Enabled         =   -1  'True
         Object.Width           =   57
         Object.Height          =   14
         WordWrap        =   0   'False
         OPCItemCount    =   0
      End
      Begin isAnalogLibrary.iLabelX iLabelX4 
         Height          =   210
         Left            =   5400
         TabIndex        =   112
         Top             =   1192
         Width           =   855
         AutoSize        =   0   'False
         Alignment       =   0
         BorderStyle     =   0
         Caption         =   "DAT"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterMarginLeft =   0
         OuterMarginTop  =   0
         OuterMarginRight=   0
         OuterMarginBottom=   0
         ShadowShow      =   -1  'True
         ShadowXOffset   =   -1
         ShadowYOffset   =   -1
         ShadowColor     =   16777215
         BackGroundColor =   -16777201
         UpdateFrameRate =   60
         Object.Visible         =   -1  'True
         FontColor       =   -16777208
         Transparent     =   -1  'True
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Enabled         =   -1  'True
         Object.Width           =   57
         Object.Height          =   14
         WordWrap        =   0   'False
         OPCItemCount    =   0
      End
      Begin isAnalogLibrary.iLabelX iLabelX3 
         Height          =   210
         Left            =   4320
         TabIndex        =   111
         Top             =   1192
         Width           =   855
         AutoSize        =   0   'False
         Alignment       =   0
         BorderStyle     =   0
         Caption         =   "CD"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterMarginLeft =   0
         OuterMarginTop  =   0
         OuterMarginRight=   0
         OuterMarginBottom=   0
         ShadowShow      =   -1  'True
         ShadowXOffset   =   -1
         ShadowYOffset   =   -1
         ShadowColor     =   16777215
         BackGroundColor =   -16777201
         UpdateFrameRate =   60
         Object.Visible         =   -1  'True
         FontColor       =   -16777208
         Transparent     =   -1  'True
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Enabled         =   -1  'True
         Object.Width           =   57
         Object.Height          =   14
         WordWrap        =   0   'False
         OPCItemCount    =   0
      End
      Begin Audiostation.MixSlider Slider_Master_Left 
         Height          =   1335
         Left            =   150
         TabIndex        =   129
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
      End
      Begin VB.PictureBox Picture1 
         Height          =   1215
         Left            =   9160
         ScaleHeight     =   1155
         ScaleWidth      =   195
         TabIndex        =   105
         Top             =   120
         Width           =   255
         Begin isAnalogLibrary.iLedBarX VU_Master_Peak 
            Height          =   1215
            Left            =   0
            TabIndex        =   106
            Top             =   0
            Width           =   255
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   2
            SegmentSpacing  =   0
            SegmentStyle    =   0
            BackGroundColor =   4210752
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   3
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   17
            Object.Height          =   81
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
      End
      Begin isDigitalLibrary.iSwitchLedX Switch_Master 
         Height          =   270
         Left            =   7560
         TabIndex        =   68
         Top             =   300
         Width           =   975
         Active          =   -1  'True
         ActiveColor     =   16776960
         AutoLedSize     =   -1  'True
         Caption         =   "MASTER"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionMargin   =   5
         IndicatorAlignment=   3
         IndicatorHeight =   4
         IndicatorMargin =   5
         IndicatorWidth  =   10
         ShowFocusRect   =   0   'False
         Enabled         =   -1  'True
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         CaptionFontColor=   -16777208
         CaptionAlignment=   1
         UpdateFrameRate =   60
         WordWrap        =   0   'False
         Glyph           =   "Form_Main.frx":111843
         BorderSize      =   2
         BorderHighlightColor=   -16777196
         BorderShadowColor=   8421504
         BackGroundColor =   12632256
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   65
         Object.Height          =   18
         MomentaryStyle  =   0
         CaptionFontName =   "Tahoma"
         CaptionFontSize =   8
         CaptionFontBold =   0   'False
         CaptionFontItalic=   0   'False
         CaptionFontUnderline=   0   'False
         CaptionFontStrikeOut=   0   'False
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSwitchLedX Switch_Dat 
         Height          =   270
         Left            =   7560
         TabIndex        =   103
         Top             =   600
         Width           =   975
         Active          =   -1  'True
         ActiveColor     =   65280
         AutoLedSize     =   -1  'True
         Caption         =   "DAT"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionMargin   =   5
         IndicatorAlignment=   3
         IndicatorHeight =   4
         IndicatorMargin =   5
         IndicatorWidth  =   10
         ShowFocusRect   =   0   'False
         Enabled         =   -1  'True
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         CaptionFontColor=   -16777208
         CaptionAlignment=   1
         UpdateFrameRate =   60
         WordWrap        =   0   'False
         Glyph           =   "Form_Main.frx":111899
         BorderSize      =   2
         BorderHighlightColor=   -16777196
         BorderShadowColor=   8421504
         BackGroundColor =   12632256
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   65
         Object.Height          =   18
         MomentaryStyle  =   0
         CaptionFontName =   "Tahoma"
         CaptionFontSize =   8
         CaptionFontBold =   0   'False
         CaptionFontItalic=   0   'False
         CaptionFontUnderline=   0   'False
         CaptionFontStrikeOut=   0   'False
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSwitchLedX Switch_Midi 
         Height          =   270
         Left            =   7560
         TabIndex        =   104
         Top             =   900
         Width           =   975
         Active          =   -1  'True
         ActiveColor     =   65280
         AutoLedSize     =   -1  'True
         Caption         =   "MIDI"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionMargin   =   5
         IndicatorAlignment=   3
         IndicatorHeight =   4
         IndicatorMargin =   5
         IndicatorWidth  =   10
         ShowFocusRect   =   0   'False
         Enabled         =   -1  'True
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         CaptionFontColor=   -16777208
         CaptionAlignment=   1
         UpdateFrameRate =   60
         WordWrap        =   0   'False
         Glyph           =   "Form_Main.frx":1118EF
         BorderSize      =   2
         BorderHighlightColor=   -16777196
         BorderShadowColor=   8421504
         BackGroundColor =   12632256
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   65
         Object.Height          =   18
         MomentaryStyle  =   0
         CaptionFontName =   "Tahoma"
         CaptionFontSize =   8
         CaptionFontBold =   0   'False
         CaptionFontItalic=   0   'False
         CaptionFontUnderline=   0   'False
         CaptionFontStrikeOut=   0   'False
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSwitchLedX Switch_Rec 
         Height          =   270
         Left            =   1680
         TabIndex        =   107
         Top             =   300
         Width           =   975
         Active          =   -1  'True
         ActiveColor     =   65280
         AutoLedSize     =   -1  'True
         Caption         =   "REC"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionMargin   =   5
         IndicatorAlignment=   3
         IndicatorHeight =   4
         IndicatorMargin =   5
         IndicatorWidth  =   10
         ShowFocusRect   =   0   'False
         Enabled         =   -1  'True
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         CaptionFontColor=   -16777208
         CaptionAlignment=   1
         UpdateFrameRate =   60
         WordWrap        =   0   'False
         Glyph           =   "Form_Main.frx":111945
         BorderSize      =   2
         BorderHighlightColor=   -16777196
         BorderShadowColor=   8421504
         BackGroundColor =   12632256
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   65
         Object.Height          =   18
         MomentaryStyle  =   0
         CaptionFontName =   "Tahoma"
         CaptionFontSize =   8
         CaptionFontBold =   0   'False
         CaptionFontItalic=   0   'False
         CaptionFontUnderline=   0   'False
         CaptionFontStrikeOut=   0   'False
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSwitchLedX Switch_Line 
         Height          =   270
         Left            =   1680
         TabIndex        =   108
         Top             =   600
         Width           =   975
         Active          =   0   'False
         ActiveColor     =   65280
         AutoLedSize     =   -1  'True
         Caption         =   "LINE"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionMargin   =   5
         IndicatorAlignment=   3
         IndicatorHeight =   4
         IndicatorMargin =   5
         IndicatorWidth  =   10
         ShowFocusRect   =   0   'False
         Enabled         =   -1  'True
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         CaptionFontColor=   -16777208
         CaptionAlignment=   1
         UpdateFrameRate =   60
         WordWrap        =   0   'False
         Glyph           =   "Form_Main.frx":11199B
         BorderSize      =   2
         BorderHighlightColor=   -16777196
         BorderShadowColor=   8421504
         BackGroundColor =   12632256
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   65
         Object.Height          =   18
         MomentaryStyle  =   0
         CaptionFontName =   "Tahoma"
         CaptionFontSize =   8
         CaptionFontBold =   0   'False
         CaptionFontItalic=   0   'False
         CaptionFontUnderline=   0   'False
         CaptionFontStrikeOut=   0   'False
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSwitchLedX Switch_CD 
         Height          =   270
         Left            =   1680
         TabIndex        =   109
         Top             =   900
         Width           =   975
         Active          =   -1  'True
         ActiveColor     =   65280
         AutoLedSize     =   -1  'True
         Caption         =   "CD"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionMargin   =   5
         IndicatorAlignment=   3
         IndicatorHeight =   4
         IndicatorMargin =   5
         IndicatorWidth  =   10
         ShowFocusRect   =   0   'False
         Enabled         =   -1  'True
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         CaptionFontColor=   -16777208
         CaptionAlignment=   1
         UpdateFrameRate =   60
         WordWrap        =   0   'False
         Glyph           =   "Form_Main.frx":1119F1
         BorderSize      =   2
         BorderHighlightColor=   -16777196
         BorderShadowColor=   8421504
         BackGroundColor =   12632256
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   65
         Object.Height          =   18
         MomentaryStyle  =   0
         CaptionFontName =   "Tahoma"
         CaptionFontSize =   8
         CaptionFontBold =   0   'False
         CaptionFontItalic=   0   'False
         CaptionFontUnderline=   0   'False
         CaptionFontStrikeOut=   0   'False
         OPCItemCount    =   0
      End
      Begin Audiostation.MixSlider Slider_Master_Right 
         Height          =   1335
         Left            =   592
         TabIndex        =   130
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
      End
      Begin Audiostation.MixSlider Slider_Record_Left 
         Height          =   1335
         Left            =   3120
         TabIndex        =   132
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
      End
      Begin Audiostation.MixSlider Slider_Record_Right 
         Height          =   1215
         Left            =   3555
         TabIndex        =   133
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2143
      End
      Begin Audiostation.MixSlider Slider_CD_Left 
         Height          =   1335
         Left            =   4200
         TabIndex        =   134
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
         Value           =   50
         Max             =   1000
      End
      Begin Audiostation.MixSlider Slider_CD_Right 
         Height          =   1215
         Left            =   4635
         TabIndex        =   135
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2143
         Value           =   50
         Max             =   1000
      End
      Begin Audiostation.MixSlider Slider_Dat_Left 
         Height          =   1335
         Left            =   5280
         TabIndex        =   136
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
         Max             =   1000
      End
      Begin Audiostation.MixSlider Slider_Dat_Right 
         Height          =   1215
         Left            =   5715
         TabIndex        =   137
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2143
         Max             =   1000
      End
      Begin Audiostation.MixSlider Slider_Midi_Left 
         Height          =   1335
         Left            =   6360
         TabIndex        =   138
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
      End
      Begin Audiostation.MixSlider Slider_Midi_Right 
         Height          =   1215
         Left            =   6795
         TabIndex        =   139
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2143
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   8640
         Picture         =   "Form_Main.frx":111A47
         Top             =   240
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   1005
         Left            =   2750
         Picture         =   "Form_Main.frx":112F89
         Top             =   240
         Width           =   345
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   104
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":1142A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":114493
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":114683
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":114872
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":114A5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":114C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":114E41
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":115031
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Element 
      BackColor       =   &H00000000&
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
      Height          =   1500
      Index           =   3
      Left            =   120
      Picture         =   "Form_Main.frx":11521D
      ScaleHeight     =   1500
      ScaleWidth      =   9615
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   9615
      Begin VB.PictureBox AniCD 
         BackColor       =   &H00C0C0C0&
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
         Height          =   1215
         Left            =   1030
         Picture         =   "Form_Main.frx":14464B
         ScaleHeight     =   1215
         ScaleWidth      =   5220
         TabIndex        =   69
         Top             =   50
         Width           =   5220
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00C0C0C0&
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
         Height          =   495
         Left            =   1500
         ScaleHeight     =   495
         ScaleWidth      =   2415
         TabIndex        =   66
         Top             =   840
         Width           =   2415
         Begin Audiostation.ButtonBig cmdCDRandom 
            Height          =   390
            Left            =   40
            TabIndex        =   141
            Top             =   50
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   688
            Caption         =   "Random"
            ShowLed         =   -1  'True
         End
         Begin Audiostation.ButtonBig cmdCDLoop 
            Height          =   390
            Left            =   1200
            TabIndex        =   142
            Top             =   50
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   688
            Caption         =   "Loop"
            ShowLed         =   -1  'True
         End
      End
      Begin VB.PictureBox Light_Panel_CD 
         BackColor       =   &H00000000&
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
         Height          =   615
         Left            =   8540
         ScaleHeight     =   615
         ScaleWidth      =   825
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   155
         Width           =   830
      End
      Begin VB.PictureBox Picture24 
         BackColor       =   &H00C0C0C0&
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
         Height          =   435
         Left            =   4420
         ScaleHeight     =   435
         ScaleWidth      =   5055
         TabIndex        =   4
         Top             =   840
         Width           =   5055
         Begin VB.CommandButton Command20 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            Picture         =   "Form_Main.frx":14CA85
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            Picture         =   "Form_Main.frx":14D04F
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command21 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Picture         =   "Form_Main.frx":14D5D9
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command22 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Picture         =   "Form_Main.frx":14DB63
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command23 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            Picture         =   "Form_Main.frx":14E0ED
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command24 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1495
            Picture         =   "Form_Main.frx":14E6B7
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            Picture         =   "Form_Main.frx":14EC81
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   50
            Width           =   495
         End
      End
      Begin isDigitalLibrary.iSevenSegmentHexadecimalX Digit_Track_CD 
         Height          =   420
         Left            =   5800
         TabIndex        =   12
         Top             =   180
         Width           =   705
         Value           =   "0"
         DigitCount      =   2
         LeadingStyle    =   1
         AutoSize        =   0   'False
         DigitSpacing    =   3
         SegmentMargin   =   3
         SegmentColor    =   65280
         SegmentSeperation=   1
         SegmentSize     =   1
         ShowOffSegments =   -1  'True
         PowerOff        =   0   'False
         BackGroundColor =   0
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         SegmentOffColor =   8421504
         AutoSegmentOffColor=   -1  'True
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   47
         Object.Height          =   28
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSevenSegmentClockX Digit_Time_CD 
         Height          =   420
         Left            =   7190
         TabIndex        =   13
         Top             =   180
         Width           =   1380
         Time            =   0
         ShowSeconds     =   -1  'True
         ShowHours       =   0   'False
         HourStyle       =   0
         AutoSize        =   0   'False
         DigitSpacing    =   3
         SegmentMargin   =   3
         SegmentColor    =   65280
         SegmentSeperation=   1
         SegmentSize     =   1
         ShowOffSegments =   -1  'True
         PowerOff        =   0   'False
         BackGroundColor =   0
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         Hours           =   0
         Minutes         =   0
         Seconds         =   0
         CountDirection  =   0
         CountTimerEnabled=   0   'False
         SegmentOffColor =   65280
         AutoSegmentOffColor=   -1  'True
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   92
         Object.Height          =   28
         OPCItemCount    =   0
      End
      Begin VB.PictureBox Picture15 
         BackColor       =   &H00C0C0C0&
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
         Height          =   1215
         Left            =   80
         ScaleHeight     =   1215
         ScaleWidth      =   1410
         TabIndex        =   3
         Top             =   100
         Width           =   1410
         Begin Audiostation.ButtonBig cmdOpenCD 
            Height          =   390
            Left            =   50
            TabIndex        =   128
            Top             =   50
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "Eject CD"
         End
      End
      Begin VB.Image Light_CD_Play_On 
         Height          =   135
         Left            =   6900
         Picture         =   "Form_Main.frx":14F24B
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Light_CD_Pause_On 
         Height          =   150
         Left            =   6895
         Picture         =   "Form_Main.frx":14F725
         Top             =   550
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10200
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Element 
      BackColor       =   &H00000000&
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
      Height          =   1575
      Index           =   5
      Left            =   120
      Picture         =   "Form_Main.frx":14FC0F
      ScaleHeight     =   1575
      ScaleWidth      =   9735
      TabIndex        =   71
      Top             =   6900
      Visible         =   0   'False
      Width           =   9735
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00000000&
         Height          =   1095
         Left            =   4080
         ScaleHeight     =   1035
         ScaleWidth      =   5295
         TabIndex        =   148
         Top             =   150
         Width           =   5350
         Begin isAnalogLibrary.iLedBarX VU_Right_Output 
            Height          =   135
            Left            =   2670
            TabIndex        =   150
            Top             =   840
            Width           =   2630
            SegmentDirection=   2
            SegmentMargin   =   2
            SegmentSize     =   5
            SegmentSpacing  =   2
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   175
            Object.Height          =   9
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isDigitalLibrary.iSwitchLedX Switch_SpectrumAnalyzer 
            Height          =   270
            Left            =   120
            TabIndex        =   149
            Top             =   120
            Width           =   1815
            Active          =   -1  'True
            ActiveColor     =   65280
            AutoLedSize     =   -1  'True
            Caption         =   "Spectrum Analyzer"
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CaptionMargin   =   5
            IndicatorAlignment=   3
            IndicatorHeight =   4
            IndicatorMargin =   5
            IndicatorWidth  =   10
            ShowFocusRect   =   0   'False
            Enabled         =   -1  'True
            BorderStyle     =   0
            Object.Visible         =   -1  'True
            CaptionFontColor=   -16777208
            CaptionAlignment=   1
            UpdateFrameRate =   60
            WordWrap        =   0   'False
            Glyph           =   "Form_Main.frx":17F03D
            BorderSize      =   2
            BorderHighlightColor=   -16777196
            BorderShadowColor=   8421504
            BackGroundColor =   12632256
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   121
            Object.Height          =   18
            MomentaryStyle  =   0
            CaptionFontName =   "Tahoma"
            CaptionFontSize =   8
            CaptionFontBold =   0   'False
            CaptionFontItalic=   0   'False
            CaptionFontUnderline=   0   'False
            CaptionFontStrikeOut=   0   'False
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Left_Output 
            Height          =   135
            Left            =   20
            TabIndex        =   151
            Top             =   840
            Width           =   2630
            SegmentDirection=   3
            SegmentMargin   =   2
            SegmentSize     =   5
            SegmentSpacing  =   2
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   176
            Object.Height          =   9
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Audio output level"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   120
            TabIndex        =   152
            Top             =   650
            Width           =   5115
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   130
         ScaleHeight     =   1020
         ScaleWidth      =   3780
         TabIndex        =   73
         Top             =   170
         Width           =   3840
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   1
            Left            =   240
            TabIndex        =   75
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   2
            Left            =   360
            TabIndex        =   76
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   3
            Left            =   480
            TabIndex        =   77
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   4
            Left            =   600
            TabIndex        =   78
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   5
            Left            =   720
            TabIndex        =   79
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   6
            Left            =   840
            TabIndex        =   80
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   7
            Left            =   960
            TabIndex        =   81
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   8
            Left            =   1080
            TabIndex        =   82
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   9
            Left            =   1200
            TabIndex        =   83
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   10
            Left            =   1320
            TabIndex        =   84
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   11
            Left            =   1440
            TabIndex        =   85
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   12
            Left            =   1560
            TabIndex        =   86
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   13
            Left            =   1680
            TabIndex        =   87
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   14
            Left            =   1800
            TabIndex        =   88
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   15
            Left            =   1920
            TabIndex        =   89
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   16
            Left            =   2040
            TabIndex        =   90
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   17
            Left            =   2160
            TabIndex        =   91
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   18
            Left            =   2280
            TabIndex        =   92
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   19
            Left            =   2400
            TabIndex        =   93
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   20
            Left            =   2520
            TabIndex        =   94
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   21
            Left            =   2640
            TabIndex        =   95
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   22
            Left            =   2760
            TabIndex        =   96
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   23
            Left            =   2880
            TabIndex        =   97
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   24
            Left            =   3000
            TabIndex        =   98
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   25
            Left            =   3120
            TabIndex        =   99
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   26
            Left            =   3240
            TabIndex        =   100
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   27
            Left            =   3360
            TabIndex        =   101
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Spectrum 
            Height          =   795
            Index           =   28
            Left            =   3480
            TabIndex        =   102
            Top             =   100
            Width           =   150
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   255
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   10
            Object.Height          =   53
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00C0C0C0&
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
         Height          =   975
         Left            =   1350
         ScaleHeight     =   975
         ScaleWidth      =   1095
         TabIndex        =   72
         Top             =   50
         Width           =   1095
      End
   End
   Begin VB.PictureBox Element 
      BackColor       =   &H00000000&
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
      Height          =   1500
      Index           =   2
      Left            =   120
      Picture         =   "Form_Main.frx":17F093
      ScaleHeight     =   1500
      ScaleWidth      =   9735
      TabIndex        =   14
      Top             =   2340
      Visible         =   0   'False
      Width           =   9735
      Begin VB.PictureBox Picture33 
         BackColor       =   &H00000000&
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
         Height          =   550
         Left            =   6510
         ScaleHeight     =   555
         ScaleWidth      =   255
         TabIndex        =   65
         Top             =   200
         Visible         =   0   'False
         Width           =   255
         Begin VB.Image Image2 
            Height          =   240
            Left            =   0
            Picture         =   "Form_Main.frx":1AE4C1
            Top             =   0
            Width           =   240
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   0
            Picture         =   "Form_Main.frx":1AEA4B
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.PictureBox Picture17 
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
         Height          =   1300
         Left            =   1480
         Picture         =   "Form_Main.frx":1AEFD5
         ScaleHeight     =   1305
         ScaleWidth      =   2055
         TabIndex        =   59
         Top             =   40
         Width           =   2055
         Begin VB.PictureBox Picture32 
            BackColor       =   &H00C0C0C0&
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
            Height          =   375
            Left            =   960
            ScaleHeight     =   375
            ScaleWidth      =   1035
            TabIndex        =   63
            Top             =   830
            Width           =   1035
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Digital Audio Transport"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   375
               Left            =   0
               TabIndex        =   64
               Top             =   50
               Width           =   1005
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H00000000&
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
            Height          =   530
            Left            =   210
            ScaleHeight     =   525
            ScaleWidth      =   1695
            TabIndex        =   60
            Top             =   240
            Width           =   1695
            Begin VB.PictureBox Picture11 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
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
               Height          =   360
               Left            =   50
               Picture         =   "Form_Main.frx":1B802F
               ScaleHeight     =   360
               ScaleWidth      =   1560
               TabIndex        =   61
               Top             =   160
               Width           =   1560
            End
            Begin VB.Label lbl_Filename 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Onbekend"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   90
               TabIndex        =   62
               Tag             =   "1013"
               Top             =   0
               UseMnemonic     =   0   'False
               Width           =   1485
            End
         End
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   2160
         Top             =   720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Main.frx":1B820F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Main.frx":1B87A9
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H00C0C0C0&
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
         Height          =   495
         Left            =   3705
         ScaleHeight     =   495
         ScaleWidth      =   5775
         TabIndex        =   16
         Top             =   840
         Width           =   5775
         Begin VB.CommandButton cmdAudioPlayer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            Picture         =   "Form_Main.frx":1B8D43
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton cmdAudioPlayer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   5220
            Picture         =   "Form_Main.frx":1B92CD
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton cmdAudioPlayer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   4680
            Picture         =   "Form_Main.frx":1B9897
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton cmdAudioPlayer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   4200
            Picture         =   "Form_Main.frx":1B9E61
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton cmdAudioPlayer 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   3720
            Picture         =   "Form_Main.frx":1BA3EB
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton cmdAudioPlayer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   3240
            Picture         =   "Form_Main.frx":1BA975
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton cmdAudioPlayer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2220
            Picture         =   "Form_Main.frx":1BAEFF
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton cmdAudioPlayer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   2760
            Picture         =   "Form_Main.frx":1BB4C9
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton cmdAudioPlayer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1080
            Picture         =   "Form_Main.frx":1BBA93
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   50
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture16 
         BackColor       =   &H00C0C0C0&
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
         Height          =   1215
         Left            =   80
         ScaleHeight     =   1215
         ScaleWidth      =   1455
         TabIndex        =   15
         Top             =   100
         Width           =   1455
         Begin Audiostation.ButtonBig cmdPlaylistDat 
            Height          =   390
            Left            =   50
            TabIndex        =   126
            Top             =   10
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "Playlist"
         End
         Begin Audiostation.ButtonBig cmdSettingsDat 
            Height          =   390
            Left            =   50
            TabIndex        =   127
            Top             =   810
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "Settings"
         End
      End
      Begin isAnalogLibrary.iLedBarX VU_Left 
         Height          =   135
         Left            =   7310
         TabIndex        =   24
         Top             =   240
         Width           =   1935
         SegmentDirection=   2
         SegmentMargin   =   1
         SegmentSize     =   5
         SegmentSpacing  =   2
         SegmentStyle    =   0
         BackGroundColor =   0
         BorderStyle     =   0
         SectionColor1   =   65280
         SectionColor2   =   65535
         SectionColor3   =   255
         SectionEnd1     =   50
         SectionEnd2     =   75
         SectionCount    =   3
         ShowOffSegments =   -1  'True
         CurrentMax      =   0
         CurrentMin      =   30
         PositionPercent =   0
         Position        =   0
         PositionMax     =   100
         PositionMin     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         MinMaxFixed     =   0   'False
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   129
         Object.Height          =   9
         FillReferenceValue=   0
         FillReferenceEnabled=   0   'False
         SectionColor4   =   65535
         SectionColor5   =   65535
         SectionEnd3     =   0
         SectionEnd4     =   0
         OPCItemCount    =   0
      End
      Begin isAnalogLibrary.iLedBarX VU_Right 
         Height          =   135
         Left            =   7310
         TabIndex        =   25
         Top             =   420
         Width           =   1935
         SegmentDirection=   2
         SegmentMargin   =   1
         SegmentSize     =   5
         SegmentSpacing  =   2
         SegmentStyle    =   0
         BackGroundColor =   0
         BorderStyle     =   0
         SectionColor1   =   65280
         SectionColor2   =   65535
         SectionColor3   =   255
         SectionEnd1     =   50
         SectionEnd2     =   75
         SectionCount    =   3
         ShowOffSegments =   -1  'True
         CurrentMax      =   0
         CurrentMin      =   0
         PositionPercent =   0
         Position        =   0
         PositionMax     =   100
         PositionMin     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         MinMaxFixed     =   0   'False
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   129
         Object.Height          =   9
         FillReferenceValue=   0
         FillReferenceEnabled=   0   'False
         SectionColor4   =   65535
         SectionColor5   =   65535
         SectionEnd3     =   0
         SectionEnd4     =   0
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSevenSegmentHexadecimalX Digit_Track_Dat 
         Height          =   420
         Left            =   3780
         TabIndex        =   26
         Top             =   180
         Width           =   705
         Value           =   "0"
         DigitCount      =   2
         LeadingStyle    =   1
         AutoSize        =   0   'False
         DigitSpacing    =   3
         SegmentMargin   =   3
         SegmentColor    =   65280
         SegmentSeperation=   1
         SegmentSize     =   1
         ShowOffSegments =   -1  'True
         PowerOff        =   0   'False
         BackGroundColor =   0
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         SegmentOffColor =   8421504
         AutoSegmentOffColor=   -1  'True
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   47
         Object.Height          =   28
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSevenSegmentClockX Digit_Time_Dat 
         Height          =   420
         Left            =   4800
         TabIndex        =   27
         Top             =   180
         Width           =   1380
         Time            =   0
         ShowSeconds     =   -1  'True
         ShowHours       =   0   'False
         HourStyle       =   0
         AutoSize        =   0   'False
         DigitSpacing    =   3
         SegmentMargin   =   3
         SegmentColor    =   65280
         SegmentSeperation=   1
         SegmentSize     =   1
         ShowOffSegments =   -1  'True
         PowerOff        =   0   'False
         BackGroundColor =   0
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         Hours           =   0
         Minutes         =   0
         Seconds         =   0
         CountDirection  =   0
         CountTimerEnabled=   0   'False
         SegmentOffColor =   65280
         AutoSegmentOffColor=   -1  'True
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   92
         Object.Height          =   28
         OPCItemCount    =   0
      End
      Begin VB.Image Light_Dat_Play_On 
         Height          =   135
         Left            =   6590
         Picture         =   "Form_Main.frx":1BC01D
         Top             =   255
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Light_Dat_Pause_On 
         Height          =   150
         Left            =   6580
         Picture         =   "Form_Main.frx":1BC4F7
         Top             =   550
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.PictureBox Element 
      BackColor       =   &H00000000&
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
      Height          =   1500
      Index           =   1
      Left            =   120
      Picture         =   "Form_Main.frx":1BC9E1
      ScaleHeight     =   1500
      ScaleWidth      =   9735
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   9735
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00000000&
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
         Height          =   400
         Left            =   6310
         ScaleHeight     =   405
         ScaleWidth      =   1935
         TabIndex        =   38
         Top             =   200
         Width           =   1935
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   400
            Index           =   0
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   100
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   2
            Left            =   240
            TabIndex        =   41
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   3
            Left            =   360
            TabIndex        =   42
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   4
            Left            =   480
            TabIndex        =   43
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   5
            Left            =   600
            TabIndex        =   44
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   6
            Left            =   720
            TabIndex        =   45
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   7
            Left            =   840
            TabIndex        =   46
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   8
            Left            =   960
            TabIndex        =   47
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   9
            Left            =   1080
            TabIndex        =   48
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   10
            Left            =   1200
            TabIndex        =   49
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   11
            Left            =   1320
            TabIndex        =   50
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   12
            Left            =   1440
            TabIndex        =   51
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   13
            Left            =   1560
            TabIndex        =   52
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   14
            Left            =   1680
            TabIndex        =   53
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   405
            Index           =   15
            Left            =   1800
            TabIndex        =   54
            Top             =   0
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   2
            SegmentSize     =   2
            SegmentSpacing  =   1
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   65280
            SectionColor2   =   65535
            SectionColor3   =   255
            SectionEnd1     =   50
            SectionEnd2     =   75
            SectionCount    =   1
            ShowOffSegments =   -1  'True
            CurrentMax      =   0
            CurrentMin      =   0
            PositionPercent =   0
            Position        =   0
            PositionMax     =   100
            PositionMin     =   0
            Object.Visible         =   -1  'True
            Enabled         =   -1  'True
            MinMaxFixed     =   0   'False
            Transparent     =   0   'False
            UpdateFrameRate =   60
            OptionSaveAllProperties=   0   'False
            AutoFrameRate   =   0   'False
            Object.Width           =   7
            Object.Height          =   27
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H00C0C0C0&
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
         Height          =   495
         Left            =   4420
         ScaleHeight     =   495
         ScaleWidth      =   5055
         TabIndex        =   30
         Top             =   840
         Width           =   5055
         Begin VB.CommandButton Command12 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            Picture         =   "Form_Main.frx":1EC58B
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            Picture         =   "Form_Main.frx":1ECB55
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Picture         =   "Form_Main.frx":1ED0DF
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Picture         =   "Form_Main.frx":1ED669
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            Picture         =   "Form_Main.frx":1EDBF3
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1495
            Picture         =   "Form_Main.frx":1EE1BD
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4500
            Picture         =   "Form_Main.frx":1EE787
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   50
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
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
         Height          =   1215
         Left            =   80
         ScaleHeight     =   1215
         ScaleWidth      =   1335
         TabIndex        =   29
         Top             =   100
         Width           =   1335
         Begin Audiostation.ButtonBig cmdPlaylistMidi 
            Height          =   390
            Left            =   50
            TabIndex        =   124
            Top             =   20
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "Playlist"
         End
         Begin Audiostation.ButtonBig cmdSettingsMidi 
            Height          =   390
            Left            =   50
            TabIndex        =   125
            Top             =   800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "Settings"
         End
      End
      Begin isDigitalLibrary.iSevenSegmentHexadecimalX Digit_Track_Midi 
         Height          =   420
         Left            =   5040
         TabIndex        =   56
         Top             =   180
         Width           =   705
         Value           =   "0"
         DigitCount      =   2
         LeadingStyle    =   1
         AutoSize        =   0   'False
         DigitSpacing    =   3
         SegmentMargin   =   3
         SegmentColor    =   65280
         SegmentSeperation=   1
         SegmentSize     =   1
         ShowOffSegments =   -1  'True
         PowerOff        =   0   'False
         BackGroundColor =   0
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         SegmentOffColor =   8421504
         AutoSegmentOffColor=   -1  'True
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   47
         Object.Height          =   28
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSevenSegmentClockX Digit_Time_Midi 
         Height          =   420
         Left            =   8250
         TabIndex        =   57
         Top             =   180
         Width           =   1140
         Time            =   0
         ShowSeconds     =   -1  'True
         ShowHours       =   0   'False
         HourStyle       =   0
         AutoSize        =   0   'False
         DigitSpacing    =   3
         SegmentMargin   =   3
         SegmentColor    =   65280
         SegmentSeperation=   1
         SegmentSize     =   1
         ShowOffSegments =   -1  'True
         PowerOff        =   0   'False
         BackGroundColor =   0
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         Hours           =   0
         Minutes         =   0
         Seconds         =   0
         CountDirection  =   0
         CountTimerEnabled=   0   'False
         SegmentOffColor =   65280
         AutoSegmentOffColor=   -1  'True
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   76
         Object.Height          =   28
         OPCItemCount    =   0
      End
      Begin VB.Image Light_Midi_Floppy_Drive 
         Height          =   240
         Left            =   1800
         Picture         =   "Form_Main.frx":1EED51
         Top             =   555
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image6 
         Height          =   165
         Left            =   1920
         Picture         =   "Form_Main.frx":1EF3D3
         Top             =   590
         Width           =   255
      End
      Begin VB.Image FloppyIn 
         Height          =   735
         Left            =   1480
         Picture         =   "Form_Main.frx":1EF651
         Top             =   120
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.Image FloppyOut 
         Height          =   735
         Left            =   1480
         Picture         =   "Form_Main.frx":1F23EB
         Top             =   120
         Width           =   3225
      End
      Begin VB.Label lbl_Midi_Filename 
         BackStyle       =   0  'Transparent
         Caption         =   "Onbekend"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   2210
         TabIndex        =   55
         Tag             =   "1013"
         Top             =   1013
         Width           =   2055
      End
      Begin VB.Image Light_Midi_Play_On 
         Height          =   135
         Left            =   6090
         Picture         =   "Form_Main.frx":1F5185
         Top             =   250
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Light_Midi_Pause_On 
         Height          =   150
         Left            =   6085
         Picture         =   "Form_Main.frx":1F565F
         Top             =   550
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.PictureBox Element 
      BackColor       =   &H00000000&
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
      Height          =   735
      Index           =   0
      Left            =   120
      Picture         =   "Form_Main.frx":1F5B49
      ScaleHeight     =   735
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin Audiostation.Button cmdButton 
         Height          =   420
         Index           =   1
         Left            =   7680
         TabIndex        =   119
         Top             =   100
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " MIDI"
         ShowLed         =   -1  'True
         Active          =   -1  'True
      End
      Begin Audiostation.Button cmdButton 
         Height          =   420
         Index           =   2
         Left            =   6820
         TabIndex        =   120
         Top             =   100
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " DAT"
         ShowLed         =   -1  'True
         Active          =   -1  'True
      End
      Begin Audiostation.Button cmdButton 
         Height          =   420
         Index           =   3
         Left            =   5970
         TabIndex        =   121
         Top             =   100
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " CD"
         ShowLed         =   -1  'True
         Active          =   -1  'True
      End
      Begin Audiostation.Button cmdButton 
         Height          =   420
         Index           =   4
         Left            =   5120
         TabIndex        =   122
         Top             =   100
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Mixer"
         ShowLed         =   -1  'True
         Active          =   -1  'True
      End
      Begin Audiostation.Button cmdButton 
         Height          =   420
         Index           =   5
         Left            =   4260
         TabIndex        =   123
         Top             =   100
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " Clock"
         ShowLed         =   -1  'True
         Active          =   -1  'True
      End
      Begin Audiostation.Button cmdButton 
         Height          =   420
         Index           =   0
         Left            =   8610
         TabIndex        =   140
         Top             =   100
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "   Power  "
         Alignment       =   2
      End
      Begin VB.Image OptionsMenuButton 
         Height          =   405
         Left            =   120
         Picture         =   "Form_Main.frx":20CA47
         Top             =   100
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   450
         Left            =   720
         Picture         =   "Form_Main.frx":20D28D
         Top             =   105
         Width           =   3165
      End
   End
   Begin VB.Timer Trm_CD_Play 
      Interval        =   1000
      Left            =   10920
      Tag             =   "0"
      Top             =   1560
   End
   Begin Audiostation.ShellPipe ShellPipe 
      Left            =   11400
      Top             =   120
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.Menu mnupopup_player 
      Caption         =   "- POPUP (1) -"
      Begin VB.Menu mnushuffle_popup 
         Caption         =   "&Shuffle"
         HelpContextID   =   1056
      End
      Begin VB.Menu mnuautonext_popup 
         Caption         =   "&Auto next"
         HelpContextID   =   1057
      End
      Begin VB.Menu mnuplaytrack_popup 
         Caption         =   "&Play one track"
         Checked         =   -1  'True
         HelpContextID   =   1058
      End
      Begin VB.Menu space02 
         Caption         =   "-"
      End
      Begin VB.Menu mnurepeattrack_popup 
         Caption         =   "&Repeat track"
         HelpContextID   =   1059
      End
      Begin VB.Menu mnurepeatplaylist_popup 
         Caption         =   "&Repeat playlist"
         Checked         =   -1  'True
         HelpContextID   =   1060
      End
   End
   Begin VB.Menu mnupopup_app 
      Caption         =   "- POPUP (2) -"
      Begin VB.Menu mnumidiplayersettings_popup 
         Caption         =   "&Midi player settings"
         HelpContextID   =   1061
      End
      Begin VB.Menu mnuaudioplayersettings_popup 
         Caption         =   "&Audio player settings"
         HelpContextID   =   1062
      End
      Begin VB.Menu space03 
         Caption         =   "-"
      End
      Begin VB.Menu mnucheck_for_updates 
         Caption         =   "&Check for updates"
      End
      Begin VB.Menu mnuabout_popup 
         Caption         =   "&About Audiostation"
         HelpContextID   =   1064
      End
      Begin VB.Menu mnuclose_popup 
         Caption         =   "&Close Audiostation"
         HelpContextID   =   1065
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AudioPeakMeter As New AudioMeter
Private AudioMaster As New AudioVolume

Dim CDRomDriveFound As Boolean
Dim DoorClose As Boolean
Dim LoopbackInit As Boolean
Dim AppInit As Boolean
Private Sub AniCD_Click()
Trm_CD_Animation.Enabled = True
End Sub
Private Sub cmdButton_Click(Index As Integer)
Select Case Index
    Case 0
        CloseApplication = True
        Unload Me
    
    Case Else
        If cmdButton(Index).Active Then
            'Turn off element
            cmdButton(Index).Active = False
            Element(Index).Tag = "DELETE"
            
            Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "Element-" & Index, "OFF")
        Else
            'Turn on element
            cmdButton(Index).Active = True
            Element(Index).Tag = "ADD"
            
            Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "Element-" & Index, "ON")
        End If
        
        If Index = 3 And cmdButton(Index).Active = True Then
            AniCD.Picture = ImageList5.ListImages(1).Picture
            AniCD.Visible = True
            
            Trm_CD_Animation.Tag = 1
            Trm_CD_Animation.Enabled = True
        End If
End Select
End Sub
Private Sub cmdAudioPlayer_Click(Index As Integer)
Select Case Index
    Case 0: PopupMenu mnupopup_player
    
    Case 1
        If AudiostationRecorder.RecordActive = True Then
            AudiostationRecorder.StopRecorder
            Trm_VU.Enabled = False
        Else
            If Settings.ReadSetting("Sibra-Soft", "Audiostation", "RecordingFileType", 0) = 0 Then
                AudiostationRecorder.RecordFilename = Environ("UserProfile") & "\output.wav"
            Else
                AudiostationRecorder.RecordFilename = Environ("UserProfile") & "\output.mp3"
            End If
            
            AudiostationRecorder.StartRecorder
            Trm_VU.Enabled = True
        End If
    
    Case 2: AudiostationMP3Player.PreviousTrack
    Case 3: AudiostationMP3Player.Forward
    Case 4: AudiostationMP3Player.StopPlay
    Case 5
        If Mp3Playlist.StorageContainer.Count = 0 Then
            If MsgBox(GetLanguage(1023), vbQuestion + vbYesNo, "Playlist") = vbYes Then
                Form_Playlist.CurrentFormType = Mp3Player
                Form_Playlist.Show vbModal
            End If
        Else
            AudiostationMP3Player.StartPlay
        End If
    
    Case 6: AudiostationMP3Player.Pause
    Case 7: AudiostationMP3Player.Rewind
    Case 8: AudiostationMP3Player.NextTrack 0, True
End Select
End Sub

Private Sub cmdCDLoop_Click()
If cmdCDLoop.Active = False Then
    cmdCDLoop.Active = True
    Light_Panel_CD.Picture = ImageList3.ListImages(3).Picture
    cmdCDRandom.Active = False
Else
    cmdCDLoop.Active = False
    Light_Panel_CD.Picture = ImageList3.ListImages(2).Picture
End If
End Sub

Private Sub cmdCDRandom_Click()
If cmdCDRandom.Active = False Then
    cmdCDRandom.Active = True
    Light_Panel_CD.Picture = ImageList3.ListImages(4).Picture
    cmdCDLoop.Active = False
Else
    cmdCDRandom.Active = False
    Light_Panel_CD.Picture = ImageList3.ListImages(2).Picture
End If
End Sub

Private Sub cmdOpenCD_Click()
Trm_CD_Animation.Enabled = True
MediaPlayerForCD.setDoorOpen
End Sub

Private Sub cmdPlaylistDat_Click()
Form_Playlist.CurrentFormType = Mp3Player
Form_Playlist.Show , Me
End Sub

Private Sub CmdPlaylistMidi_Click()
Form_Playlist.CurrentFormType = MidiPlayer
Form_Playlist.Show , Me
End Sub

Private Sub cmdSettingsDat_Click()
Call Extensions.ShellAndWait("ffmpeg.exe", "-list_devices true -f dshow -i dummy > devices.txt 2>&1")

If Extensions.FileExists(App.path & "\devices.txt") Then: Form_Settings_Record.Show vbModal
End Sub

Private Sub cmdSettingsMidi_Click()
If Form_Settings_Midi.OutputDevCombo.ListCount = 0 Then
    MsgBox GetLanguage(1051), vbExclamation
Else
    Form_Settings_Midi.Show
End If
End Sub
Private Sub Command10_Click()
Form_Settings_Midi.MIDIOutput1.Pause
Trm_Lights_Midi.Tag = 3
MidiPause = True
End Sub

Private Sub Command12_Click()
Form_Settings_Midi.MIDIOutput1.PlaybackRate = Form_Settings_Midi.MIDIOutput1.PlaybackRate + 10
End Sub

Private Sub Command13_Click()
If MidiPlaylist.StorageContainer.Count = 0 Then
    If MsgBox(GetLanguage(1023), vbQuestion + vbYesNo, "Playlist") = vbYes Then
        Form_Playlist.CurrentFormType = MidiPlayer
        Form_Playlist.Show vbModal
    End If
Else
    AudiostationMidiPlayer.StartMidiPlayback
End If
End Sub

Private Sub Command14_Click()
AudiostationMidiPlayer.StopMidiPlayBack
End Sub

Private Sub Command15_Click()
Form_Settings_Midi.MIDIOutput1.PlaybackRate = Form_Settings_Midi.MIDIOutput1.PlaybackRate - 10
End Sub

Private Sub Command16_Click()
If CurrentMidiTrackNumber = 1 Or MidiPlaylist.StorageContainer.Count = 0 Then: Exit Sub

AudiostationMidiPlayer.CurrentMidiTrackNumber = CurrentMidiTrackNumber - 1
AudiostationMidiPlayer.PreviousMidiTrack
End Sub

Private Sub Command17_Click()
MediaPlayerForCD.NextTrack
End Sub

Private Sub Command18_Click()
MediaPlayerForCD.pauseCD
End Sub

Private Sub Command20_Click()
MediaPlayerForCD.fastForward 2
End Sub

Private Sub Command21_Click()
MediaPlayerForCD.playCD
End Sub

Private Sub Command22_Click()
MediaPlayerForCD.stopCD
End Sub

Private Sub Command23_Click()
MediaPlayerForCD.fastRewind 2
End Sub

Private Sub Command24_Click()
MediaPlayerForCD.prevTrack
End Sub
Private Sub Command9_Click()
If CurrentMidiTrackNumber = MidiPlaylist.StorageContainer.Count Then: Exit Sub

AudiostationMidiPlayer.CurrentMidiTrackNumber = CurrentMidiTrackNumber + 1
AudiostationMidiPlayer.NextMidiTrack
End Sub

Private Sub Digit_Time_Dat_OnClick()
If AudiostationMP3Player.ShowElapsedTime = True Then
    AudiostationMP3Player.ShowElapsedTime = False
Else
    AudiostationMP3Player.ShowElapsedTime = True
End If
End Sub

Private Sub Form_Load()
Dim First As Boolean

Width = 9900
Height = 9855 - 310

Slider_Dat_Left.Value = 1000
Slider_Dat_Right.Value = 1000
Slider_CD_Left.Value = 1000
Slider_CD_Right.Value = 1000

ChDrive App.path
ChDir App.path

Call AudiostationMP3Player.Init
If Not Extensions.GetCurrentWindowsVersion = "Windows XP" And Not IsDebuggig Then Call initLoopBack

For I = 0 To Element.Count - 1
    Element(I).Visible = True
Next

mnupopup_player.Visible = False
mnupopup_app.Visible = False

' Get the application settings
Switch_SpectrumAnalyzer.Active = Settings.ReadSetting("Sibra-Soft", "Audiostation", "UseSpectrumAnalyzer", True)
If VarType(Settings.ReadSetting("Sibra-Soft", "Audiostation", "CD", "False")) = vbBoolean Then
    CDRomDriveFound = Settings.ReadSetting("Sibra-Soft", "Audiostation", "CD", "False")
Else
    CDRomDriveFound = True
End If

'Display program version
lbl_version.Caption = "Version: " & App.Major & "." & App.Minor & " Build: " & App.Revision

'Set application default
Form_Settings_Midi.MIDIOutput1.VolumeLeft = 65535
Form_Settings_Midi.MIDIOutput1.VolumeRight = 65535
    
Call SetLanguage(Me)

Light_Panel_CD.Picture = ImageList3.ListImages(1).Picture
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True

Call ApplicationDestructor
Call Shell(App.path & "/close.exe", vbHide)

Cancel = False
End Sub

Private Sub lbl_Midi_Filename_Change()
'Caption = "Audiostation (" & lbl_Midi_Filename.Caption & ")"
lbl_Midi_Filename.ToolTipText = lbl_Midi_Filename.Caption
End Sub

Private Sub mnuabout_popup_Click()
Form_About.Show vbModal
End Sub

Private Sub mnuaudioplayersettings_popup_Click()
cmdSettingsDat_Click
End Sub

Private Sub mnuautonext_popup_Click()
mnushuffle_popup.Checked = False
mnuautonext_popup.Checked = True
mnuplaytrack_popup.Checked = False

AudiostationMP3Player.PlaySingleTrack = False
AudiostationMP3Player.Shuffle = False
AudiostationMP3Player.AutoNext = True
End Sub

Private Sub mnucheck_for_updates_Click()
Dim iRet As Integer

Dim appMinor As Integer
Dim appMajor As Integer
Dim appRevision As Integer

Dim webMinor As Integer
Dim webMajor As Integer
Dim webRevision As Integer

Dim webResponse As String

webResponse = WebRequest.WebRequest("https://www.audiostation.org/app-deploy/audiostation/version.txt")

webMinor = Extensions.Explode(webResponse, ".", "0")
webMajor = Extensions.Explode(webResponse, ".", "1")
webRevision = Extensions.Explode(webResponse, ".", "2")

appMinor = App.Minor
appMajor = App.Major
appRevision = App.Revision

If webMinor > appMinor Or webMajor = appMajor Or webRevision > appRevision Then
    iRet = MsgBox("A new version of Audiostation was found, do you want to download this new version?", vbYesNo + vbQuestion, "New version")
    
    If iRet = vbYes Then
        Shell "explorer.exe https://www.audiostation.org", vbNormalFocus
    End If
Else
    MsgBox "You have the newest version of Audiostation", vbInformation
End If
End Sub

Private Sub mnuclose_popup_Click()
End
End Sub

Private Sub mnumidiplayersettings_popup_Click()
cmdSettingsMidi_Click
End Sub

Private Sub mnuplaytrack_popup_Click()
mnushuffle_popup.Checked = False
mnuautonext_popup.Checked = False
mnuplaytrack_popup.Checked = True

AudiostationMP3Player.PlaySingleTrack = True
AudiostationMP3Player.AutoNext = False
AudiostationMP3Player.Shuffle = False
End Sub

Private Sub mnurepeatplaylist_popup_Click()
mnurepeatplaylist_popup.Checked = True
mnurepeattrack_popup.Checked = False

AudiostationMP3Player.RepeatTrack = False
AudiostationMP3Player.RepeatPlaylist = True
End Sub

Private Sub mnurepeattrack_popup_Click()
mnurepeatplaylist_popup.Checked = False
mnurepeattrack_popup.Checked = True

AudiostationMP3Player.RepeatPlaylist = False
AudiostationMP3Player.RepeatTrack = True
End Sub

Private Sub mnushuffle_popup_Click()
mnushuffle_popup.Checked = True
mnuautonext_popup.Checked = False
mnuplaytrack_popup.Checked = False

AudiostationMP3Player.PlaySingleTrack = False
AudiostationMP3Player.AutoNext = False
AudiostationMP3Player.Shuffle = True
End Sub

Private Sub OptionsMenuButton_Click()
PopupMenu mnupopup_app
End Sub

Private Sub ShellPipe_ChildFinished()
Dim ReturnCode As Long

ReturnCode = ShellPipe.FinishChild(0)
Debug.Print "Program complete. Return code: " & CStr(ReturnCode)
End Sub

Private Sub Slider_CD_Left_OnPositionChange()
Call MediaPlayerForCD.SetLeftVolume(Slider_CD_Left.Value)
End Sub

Private Sub Slider_CD_Right_OnPositionChange()
Call MediaPlayerForCD.SetRightVolume(Slider_CD_Right.Value)
End Sub

Private Sub Slider_Dat_Left_OnPositionChange()
Call MediaPlayer.SetLeftVolume(Slider_Dat_Left.Value)
End Sub

Private Sub Slider_Dat_Right_OnPositionChange()
Call MediaPlayer.SetRightVolume(Slider_Dat_Right.Value)
End Sub

Private Sub Slider_Master_Left_OnPositionChange()
Call AudioMaster.SetChannelVolumeLevelScalar(0, Slider_Master_Left.Value / 100)
End Sub

Private Sub Slider_Master_Right_OnPositionChange()
Call AudioMaster.SetChannelVolumeLevelScalar(1, Slider_Master_Right.Value / 100)
End Sub

Private Sub Switch_CD_OnChange()
If Switch_CD.Active Then
    MediaPlayerForCD.setAudioOn
Else
    MediaPlayerForCD.setAudioOff
End If
End Sub

Private Sub Switch_Dat_OnChange()
If Switch_Dat.Active Then
    MediaPlayer.setAudioOn
Else
    MediaPlayer.setAudioOff
End If
End Sub

Private Sub Switch_Master_OnChange()
If Switch_Master.Active Then
    AudioMaster.SetMute (0)
Else
    AudioMaster.SetMute (1)
End If
End Sub

Private Sub Switch_SpectrumAnalyzer_OnChange()
Dim I As Integer

If Not Switch_SpectrumAnalyzer.Active Then
    For I = 0 To VU_Spectrum.Count - 1
        VU_Spectrum(I).Position = 0
    Next
End If

Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "UseSpectrumAnalyzer", Switch_SpectrumAnalyzer.Active)
End Sub

Private Sub Trm_Animation_Timer()
Dim J As Integer

J = Trm_Animation.Tag
Picture11.Picture = ImageList1.ListImages.Item(J).Picture

If Trm_Animation.Tag = ImageList1.ListImages.Count Then
    Trm_Animation.Tag = 1
Else
    Trm_Animation.Tag = Trm_Animation.Tag + 1
End If
End Sub

Private Sub Trm_AudioPlayer_Timer()

End Sub

Private Sub Trm_CD_Animation_Timer()
Dim ImgIndex As Integer
    
ImgIndex = Trm_CD_Animation.Tag

AniCD.Picture = ImageList5.ListImages(ImgIndex).Picture
AniCD.Visible = True

If DoorClose = True Then
    If Trm_CD_Animation.Tag = 1 Then
        DoorClose = False
        Trm_CD_Animation.Enabled = False
        AniCD.Visible = False
    Else
        Trm_CD_Animation.Tag = Trm_CD_Animation.Tag - 1
    End If
Else
    If Trm_CD_Animation.Tag = 7 Then
        DoorClose = True
        Trm_CD_Animation.Enabled = False
    Else
        Trm_CD_Animation.Tag = Trm_CD_Animation.Tag + 1
    End If
End If
End Sub

Private Sub Trm_CD_Play_Timer()
Dim Time_Seconds As String
Dim Time_Minutes As String
Dim split_value

'Start the CD-Rom drive
If Not Settings.ReadSetting("Sibra-Soft", "Audiostation", "CD") = "false" Then
    MediaPlayerForCD.startCD (Settings.ReadSetting("Sibra-Soft", "Audiostation", "CD"))
End If

On Error Resume Next
split_value = Split(MediaPlayerForCD.getPositionTMSF, ":")

Digit_Track_CD.Value = Left(MediaPlayerForCD.getPositionTMSF, 2)
Digit_Time_CD.seconds = split_value(2)
Digit_Time_CD.Minutes = split_value(1)
End Sub

Private Sub Trm_Check_File_Timer()
Dim MediaFile As String
Dim MediaIndex As String
Dim MediaDuration As String
Dim MediaTagManager As New Mp3Info

Begin:
If Settings.ReadSetting("Sibra-Soft", "Audiostation", "CheckFile") = "" Then
    'Check if there is a file to open
Else
    MediaFile = Settings.ReadSetting("Sibra-Soft", "Audiostation", "CheckFile")
    
    If Not Extensions.FileExists(MediaFile) Then
        Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CheckFile", vbNullString)
        Exit Sub
    End If
    
    Select Case LCase(Right(MediaFile, 3))
        Case "mp3", "wav", "mp2", "aac", "snd", "au", "rmi", "cda", "wma", "m4a"
            MediaDuration = 0
            AudiostationMidiPlayer.StopMidiPlayBack
            
            If Mp3Playlist.IsExistingItem(MediaFile) > 0 Then
                AudiostationMP3Player.CurrentTrackNumber = Mp3Playlist.IsExistingItem(MediaFile)
                AudiostationMP3Player.StartPlay
            Else
                MediaIndex = format(Mp3Playlist.StorageContainer.Count + 1, "00")
                
                ' Only get the duration when it's a mp3 file
                If LCase(Right(MediaFile, 3)) = "mp3" Then
                    MediaTagManager.FileName = MediaFile
                    MediaDuration = Extensions.TimeString(MediaTagManager.SongLength)
                End If
                
                If MediaDuration = "0" Then: MediaDuration = "-"
                
                Mp3Playlist.AddToStorage MediaFile, MediaIndex & ";" & MediaFile & ";" & MediaDuration
                
                AudiostationMP3Player.CurrentTrackNumber = Mp3Playlist.StorageContainer.Count
                AudiostationMP3Player.StartPlay
            End If
                        
        Case "mid", "kar", "mus", "sid"
            AudiostationMP3Player.StopPlay
            
            CurrentIndex = format(MidiPlaylist.StorageContainer.Count + 1, "00")
            CurrentMediaDuration = "-"
    
            MidiPlaylist.AddToStorage MediaFile, CurrentIndex & ";" & MediaFile & ";" & CurrentMediaDuration
            
            AudiostationMidiPlayer.CurrentMidiTrackNumber = MidiPlaylist.StorageContainer.Count
            AudiostationMidiPlayer.StartMidiPlayback
        
    Case "apl", "wpl", "m3u", "pls" 'Playlist files
        If Not (Dir(MediaFile, vbDirectory) = vbNullString) Then
            Screen.MousePointer = vbHourglass
            
            Select Case LCase(Right(file, 3))
                Case "apl": Call ModPlaylist.OpenAplPlaylist(MediaFile)
                Case "m3u": Call ModPlaylist.OpenM3uPlaylist(MediaFile)
                Case "pls": Call ModPlaylist.OpenPlsPlaylist(MediaFile)
                Case "wpl": Call ModPlaylist.OpenWplPlaylist(MediaFile)
            End Select
            
            Form_Playlist.CurrentFormType = Mp3Player
            Form_Playlist.Show , Form_Main
        Else
            Debug.Print "Playlist file could not be found"
        End If
            
    Case Else
        'Check if it's a file that needs to be converted
        Select Case LCase(Right(MediaFile, 3))
            Case "act": Call ModConvert.Convert(MediaFile, [Voice File Format], MP3): GoTo Begin
            Case "caf": Call ModConvert.Convert(MediaFile, [Apple Core Format], MP3): GoTo Begin
            Case "ogg": Call ModConvert.Convert(MediaFile, [OGG], MP3): GoTo Begin
            Case "omo": Call ModConvert.Convert(MediaFile, [Sony OpenMG Audio], MP3): GoTo Begin
            Case "s64": Call ModConvert.Convert(MediaFile, [Sony Wave64], MP3): GoTo Begin
            Case "voc": Call ModConvert.Convert(MediaFile, [Voice File Format], MP3): GoTo Begin
        End Select
        
        'Check if it's a file that needs to be converted
        Select Case LCase(Right(MediaFile, 2))
            Case "ra": Call ModConvert.Convert(MediaFile, [Real Audio], MP3): GoTo Begin
            Case "rm": Call ModConvert.Convert(MediaFile, [Real Media], MP3): GoTo Begin
           
            Case Else: MsgBox GetLanguage(1020), vbCritical
        End Select
    End Select
    
    'Delete check file setting
    Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CheckFile", vbNullString)
End If
End Sub

Private Sub Trm_Floppy_Drive_Light_Timer()
If Light_Midi_Floppy_Drive.Visible = True Then
    Light_Midi_Floppy_Drive.Visible = False
Else
    Light_Midi_Floppy_Drive.Visible = True
End If
End Sub

Private Sub Trm_Lights_Midi_Timer()
If Trm_Lights_Midi.Tag = "1" Then
    Trm_Floppy_Drive_Light.Enabled = True
    FloppyIn.Visible = True
    
    If Light_Midi_Play_On.Visible = True Then
        Light_Midi_Play_On.Visible = False
    Else
        Light_Midi_Play_On.Visible = True
    End If
End If

If Trm_Lights_Midi.Tag = "2" Then
    Trm_Floppy_Drive_Light.Enabled = False
    FloppyIn.Visible = False
    
    Light_Midi_Floppy_Drive.Visible = False
    Light_Midi_Play_On.Visible = False
    Light_Midi_Pause_On.Visible = False
    Light_Midi_Play_On.Visible = False
End If

If Trm_Lights_Midi.Tag = "3" Then
    Trm_Floppy_Drive_Light.Enabled = False
    Light_Midi_Floppy_Drive.Visible = False
    FloppyIn.Visible = False
    
    If Light_Midi_Pause_On.Visible = True Then
        Light_Midi_Pause_On.Visible = False
    Else
        Light_Midi_Pause_On.Visible = True
    End If
End If
End Sub

Private Sub Trm_Lights_Timer()
If AudiostationRecorder.RecordActive = True Then
    Picture33.Visible = True
    
    If Picture33.Visible = True Then
        If Image2.Visible = True Then
            Image2.Visible = False
        Else
            Image2.Visible = True
        End If
    End If
Else
    Picture33.Visible = False
End If

' Media player
If AudiostationMP3Player.PlayState = Playing Then
    Trm_Animation.Enabled = True
    Trm_VU.Enabled = True
    Light_Dat_Pause_On.Visible = False
    
    If Light_Dat_Play_On.Visible = True Then
        Light_Dat_Play_On.Visible = False
    Else
        Light_Dat_Play_On.Visible = True
    End If
Else
    If AudiostationMP3Player.PlayState = Paused Then
        Trm_Animation.Enabled = False
        Trm_VU.Enabled = False
        Light_Dat_Play_On.Visible = True
    
        VU_Left.Position = 0
        VU_Right.Position = 0
    
        If Light_Dat_Pause_On.Visible = True Then
            Light_Dat_Pause_On.Visible = False
        Else
            Light_Dat_Pause_On.Visible = True
        End If
    Else
        Trm_VU.Enabled = False
        Light_Dat_Pause_On.Visible = False
        Light_Dat_Play_On.Visible = False
        Trm_Animation.Enabled = False
        
        VU_Left.Position = 0
        VU_Right.Position = 0
    End If
End If

' CD Player
If MediaPlayerForCD.isPlaying Then
    If Light_CD_Play_On.Visible = True Then
        Light_CD_Play_On.Visible = False
    Else
        Light_CD_Play_On.Visible = True
    End If
Else
    Light_CD_Play_On.Visible = False
End If
End Sub
Private Sub Trm_Main_Timer()
Digit_Clock.Hours = format(Now, "hh")
Digit_Clock.Minutes = format(Now, "nn")
Digit_Clock.seconds = format(Now, "ss")

Digit_Track_Dat.Value = AudiostationMP3Player.CurrentTrackNumber
Digit_Track_Midi.Value = AudiostationMidiPlayer.CurrentMidiTrackNumber

' Enable the activated rack
Dim I As Integer
For I = 1 To cmdButton.Count - 1
    Dim mustBeOff As String
    
    mustBeOff = Settings.ReadSetting("Sibra-Soft", "Audiostation", "Element-" & I, "OFF")
        
    If Not mustBeOff = "OFF" Then
        ElementOff(I).Visible = False
        cmdButton(I).Active = True
    Else
        ElementOff(I).Visible = True
        cmdButton(I).Active = False
    End If
Next

' Startup loop
If Trm_Main.Tag = 6 Then
    Trm_Main.Interval = 1
Else
    cmdButton(Trm_Main.Tag).Active = True
    Trm_Main.Tag = Trm_Main.Tag + 1
End If

' Show the elapsed or leftover time
If AudiostationMP3Player.PlayState = Playing Then
    lbl_Filename.Caption = Extensions.GetFileNameFromFilePath(AudiostationMP3Player.CurrentMediaFilename, False)
    lbl_Filename.ToolTipText = Extensions.GetFileNameFromFilePath(AudiostationMP3Player.CurrentMediaFilename, False)

    If AudiostationMP3Player.ShowElapsedTime Then
        Dim TimeSerial As String
        TimeSerial = Extensions.MilliSecondsToTimeSerial(MediaPlayer.GetPositioninMS, SmallTimeSerial)
        
        Digit_Time_Dat.Minutes = Extensions.Explode(TimeSerial, ":", 0)
        Digit_Time_Dat.seconds = Extensions.Explode(TimeSerial, ":", 1)
    Else
        Dim SecondsLeft As Long
        SecondsLeft = MediaPlayer.GetDurationInSec - MediaPlayer.GetPositioninSec
    
        Digit_Time_Dat.Minutes = Extensions.Explode(Extensions.TimeString(SecondsLeft), ":", 0)
        Digit_Time_Dat.seconds = Extensions.Explode(Extensions.TimeString(SecondsLeft), ":", 1)
    End If
End If

' Check for stream ending
If AudiostationMP3Player.PlayState = Playing And MediaPlayer.GetDurationInMS = MediaPlayer.GetPositioninMS Then
    AudiostationMP3Player.PlayState = MediaEnded
End If

If AudiostationMP3Player.PlayState = MediaEnded Then AudiostationMP3Player.NextTrack
End Sub

Private Sub Trm_Midi_Play_Timer()
Dim I As Integer

Digit_Time_Midi.Minutes = Trm_Midi_Play.Tag \ 60
Digit_Time_Midi.seconds = format(Int(Trm_Midi_Play.Tag Mod 60), "00")

Trm_Midi_Play.Tag = Trm_Midi_Play.Tag + 1
End Sub
Private Sub initLoopBack()
Dim retVal As Long
Dim level As Single
Dim C As Integer, di As BASS_WASAPI_DEVICEINFO

outdev = -1

' Check the correct BASS was loaded
If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
    Call MsgBox("An incorrect version of BASS.DLL was loaded", vbCritical)
    End
End If

Call BASS_Free

' Get list of audio devices
C = 0
While BASS_WASAPI_GetDeviceInfo(C, di)
    If ((di.flags And BASS_DEVICE_LOOPBACK) = BASS_DEVICE_LOOPBACK And (di.flags And BASS_DEVICE_ENABLED) = BASS_DEVICE_ENABLED) Then ' it's an enabled input device
        Debug.Print "Current audio device: " & VBStrFromAnsiPtr(di.Name) & "  (" + str(C) + " )"
        If (Not First) Then
        outdev = C
        First = True
        End If
    End If
    C = C + 1
Wend

If outdev = -1 Then Debug.Print "No audio device could be found"
Call BASS_SetConfig(BASS_CONFIG_UPDATETHREADS, 0)

retVal = BASS_Init(0, 44100, BASS_DEVICE_DEFAULT, 0, 0)

Debug.Print "Bass_Init:"; retVal
If retVal = 0 Then
    Call MsgBox("Bass Initialisation failed", vbCritical)
    End
End If

retVal = BASS_WASAPI_Init(-3, 0, 0, BASS_WASAPI_BUFFER, 0.4, 0.05, AddressOf OutWasapiProc, 0)
Debug.Print "Bass_Wasapi_Init:"; retVal

If retVal = 0 Then
    Call MsgBox("WASAPI Initialisation failed", vbCritical)
    BASS_Free
    End
End If
    
BASS_WASAPI_Start

outdev = BASS_WASAPI_GetDevice()
chan = outdev   ' spectrum uses chan

LoopbackInit = True
End Sub
Private Sub Trm_VU_Timer()
VU_Master_Peak.Position = AudioPeakMeter.GetPeak * 100

VU_Left.Position = AudioPeakMeter.GetChannelPeak(0) * 100
VU_Right.Position = AudioPeakMeter.GetChannelPeak(1) * 100

VU_Left_Output.Position = AudioPeakMeter.GetChannelPeak(0) * 100
VU_Right_Output.Position = AudioPeakMeter.GetChannelPeak(1) * 100

If LoopbackInit And Switch_SpectrumAnalyzer.Active Then UpdateSpectrum
End Sub
