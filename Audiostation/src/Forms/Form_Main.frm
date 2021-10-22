VERSION 5.00
Object = "{0A362340-2E5E-11D3-85BF-00105AC8B715}#1.0#0"; "isDigitalLibrary.ocx"
Object = "{C5412DA5-2E2F-11D3-85BF-00105AC8B715}#1.0#0"; "isAnalogLibrary.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form_Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiostation"
   ClientHeight    =   9105
   ClientLeft      =   4695
   ClientTop       =   1275
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
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ElementOff 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   6
      Left            =   120
      Picture         =   "Form_Main.frx":088B
      ScaleHeight     =   1455
      ScaleWidth      =   10095
      TabIndex        =   163
      Top             =   6900
      Width           =   10095
   End
   Begin VB.Timer Timer_Stream 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   10920
      Tag             =   "0"
      Top             =   2040
   End
   Begin VB.PictureBox ElementOff 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   780
      Index           =   5
      Left            =   120
      Picture         =   "Form_Main.frx":2F53D
      ScaleHeight     =   780
      ScaleWidth      =   9615
      TabIndex        =   117
      Top             =   8400
      Width           =   9615
   End
   Begin VB.PictureBox Element 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   6
      Left            =   120
      Picture         =   "Form_Main.frx":45543
      ScaleHeight     =   855
      ScaleWidth      =   9615
      TabIndex        =   135
      Top             =   8400
      Visible         =   0   'False
      Width           =   9615
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   2880
         ScaleHeight     =   435
         ScaleWidth      =   6555
         TabIndex        =   136
         Top             =   55
         Width           =   6615
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright © 2009 - 2021 Sibra-Soft Software Production"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   45
            TabIndex        =   138
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
            TabIndex        =   137
            Top             =   20
            Width           =   180
         End
      End
      Begin isDigitalLibrary.iSevenSegmentClockX Digit_Clock 
         Height          =   495
         Left            =   80
         TabIndex        =   139
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
      Picture         =   "Form_Main.frx":5C441
      ScaleHeight     =   1500
      ScaleWidth      =   9615
      TabIndex        =   114
      Top             =   2340
      Width           =   9615
   End
   Begin VB.PictureBox ElementOff 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   1
      Left            =   120
      Picture         =   "Form_Main.frx":8B0F3
      ScaleHeight     =   1500
      ScaleWidth      =   9615
      TabIndex        =   113
      Top             =   840
      Width           =   9615
   End
   Begin VB.PictureBox ElementOff 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1590
      Index           =   4
      Left            =   120
      Picture         =   "Form_Main.frx":9AF35
      ScaleHeight     =   1590
      ScaleWidth      =   9615
      TabIndex        =   116
      Top             =   5330
      Width           =   9615
   End
   Begin VB.PictureBox ElementOff 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   3
      Left            =   120
      Picture         =   "Form_Main.frx":CB9D7
      ScaleHeight     =   1500
      ScaleWidth      =   9615
      TabIndex        =   115
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
            Picture         =   "Form_Main.frx":DB599
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":DD02B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":DEABD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":E054F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":E1FE1
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
            Picture         =   "Form_Main.frx":E3A73
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":EBEBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":FBF01
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":10BF45
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":11BF89
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":12BFCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":13C011
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Trm_Floppy_Drive_Light 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10920
      Top             =   3480
   End
   Begin VB.Timer Trm_Lights_Midi 
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
            Picture         =   "Form_Main.frx":14C055
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
      Picture         =   "Form_Main.frx":14C5EF
      ScaleHeight     =   1575
      ScaleWidth      =   9735
      TabIndex        =   1
      Top             =   5330
      Visible         =   0   'False
      Width           =   9735
      Begin isAnalogLibrary.iLabelX ILaMaster 
         Height          =   195
         Left            =   120
         TabIndex        =   124
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
      Begin Audiostation.MixSlider Slider_Master 
         Height          =   1335
         Left            =   360
         TabIndex        =   123
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
         Glyph           =   "Form_Main.frx":17DF89
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
         Glyph           =   "Form_Main.frx":17DFDF
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
         Glyph           =   "Form_Main.frx":17E035
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
         Glyph           =   "Form_Main.frx":17E08B
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
         Glyph           =   "Form_Main.frx":17E0E1
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
         Glyph           =   "Form_Main.frx":17E137
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
      Begin Audiostation.MixSlider Slider_Record_Left 
         Height          =   1335
         Left            =   3120
         TabIndex        =   125
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
      End
      Begin Audiostation.MixSlider Slider_Record_Right 
         Height          =   1215
         Left            =   3555
         TabIndex        =   126
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2143
      End
      Begin Audiostation.MixSlider Slider_CD_Left 
         Height          =   1335
         Left            =   4200
         TabIndex        =   127
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
         TabIndex        =   128
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
         TabIndex        =   129
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
         Max             =   1000
      End
      Begin Audiostation.MixSlider Slider_Dat_Right 
         Height          =   1215
         Left            =   5715
         TabIndex        =   130
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2143
         Max             =   1000
      End
      Begin Audiostation.MixSlider Slider_Midi_Left 
         Height          =   1335
         Left            =   6360
         TabIndex        =   131
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
      End
      Begin Audiostation.MixSlider Slider_Midi_Right 
         Height          =   1215
         Left            =   6795
         TabIndex        =   132
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2143
      End
      Begin isAnalogLibrary.iLabelX iLabelX5 
         Height          =   210
         Left            =   6480
         TabIndex        =   155
         Top             =   1200
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
      Begin VB.Image Image1 
         Height          =   960
         Left            =   8640
         Picture         =   "Form_Main.frx":17E18D
         Top             =   240
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   1005
         Left            =   2750
         Picture         =   "Form_Main.frx":17F6CF
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
            Picture         =   "Form_Main.frx":1809E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":180BD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":180DC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":180FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":1811A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":181394
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":181587
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":181777
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
      Picture         =   "Form_Main.frx":181963
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
         Picture         =   "Form_Main.frx":1B0D91
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
         Begin Audiostation.ButtonBig Button_CDRandom 
            Height          =   390
            Left            =   40
            TabIndex        =   133
            Top             =   50
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   688
            Caption         =   "Random"
            ShowLed         =   -1  'True
            TextAlignment   =   0
         End
         Begin Audiostation.ButtonBig Button_CDLoop 
            Height          =   390
            Left            =   1200
            TabIndex        =   134
            Top             =   50
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   688
            Caption         =   "Loop"
            ShowLed         =   -1  'True
            TextAlignment   =   0
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
         Begin VB.CommandButton Button_CDPlayer 
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
            Left            =   3960
            Picture         =   "Form_Main.frx":1B91CB
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_CDPlayer 
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
            Left            =   3480
            Picture         =   "Form_Main.frx":1B9795
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_CDPlayer 
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
            Left            =   3000
            Picture         =   "Form_Main.frx":1B9D1F
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_CDPlayer 
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
            Left            =   2520
            Picture         =   "Form_Main.frx":1BA2A9
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_CDPlayer 
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
            Left            =   2040
            Picture         =   "Form_Main.frx":1BA833
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_CDPlayer 
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
            Left            =   1495
            Picture         =   "Form_Main.frx":1BADFD
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_CDPlayer 
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
            Left            =   4560
            Picture         =   "Form_Main.frx":1BB3C7
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
         Begin Audiostation.ButtonBig Button_CDOpen 
            Height          =   390
            Left            =   50
            TabIndex        =   122
            Top             =   50
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "Eject CD"
            TextAlignment   =   0
         End
      End
      Begin VB.Image Light_CD_Play_On 
         Height          =   135
         Left            =   6900
         Picture         =   "Form_Main.frx":1BB991
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Light_CD_Pause_On 
         Height          =   150
         Left            =   6895
         Picture         =   "Form_Main.frx":1BBE6B
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
      Picture         =   "Form_Main.frx":1BC355
      ScaleHeight     =   1575
      ScaleWidth      =   9735
      TabIndex        =   71
      Top             =   6900
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CommandButton Button_OpenStream 
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
         Left            =   5520
         Picture         =   "Form_Main.frx":1EB783
         Style           =   1  'Graphical
         TabIndex        =   145
         ToolTipText     =   "Open stream"
         Top             =   870
         Width           =   495
      End
      Begin isAnalogLibrary.iLabelX Label_StreamTitle 
         Height          =   210
         Left            =   6240
         TabIndex        =   156
         Top             =   120
         Width           =   3135
         AutoSize        =   0   'False
         Alignment       =   2
         BorderStyle     =   0
         Caption         =   "Nothing playing"
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
         Object.Width           =   209
         Object.Height          =   14
         WordWrap        =   0   'False
         OPCItemCount    =   0
      End
      Begin isAnalogLibrary.iLabelX iLabelX1 
         Height          =   210
         Left            =   4320
         TabIndex        =   146
         Top             =   120
         Width           =   1935
         AutoSize        =   0   'False
         Alignment       =   1
         BorderStyle     =   0
         Caption         =   "Radio Tuner Memory"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Object.Width           =   129
         Object.Height          =   14
         WordWrap        =   0   'False
         OPCItemCount    =   0
      End
      Begin VB.CommandButton Button_StopStream 
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
         Left            =   4800
         Picture         =   "Form_Main.frx":1EB94D
         Style           =   1  'Graphical
         TabIndex        =   144
         ToolTipText     =   "Stop"
         Top             =   870
         Width           =   495
      End
      Begin VB.CommandButton Button_PlayStream 
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
         Left            =   4320
         Picture         =   "Form_Main.frx":1EBED7
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "Play"
         Top             =   870
         Width           =   495
      End
      Begin Audiostation.ButtonBig Button_TunerMemory 
         Height          =   390
         Index           =   1
         Left            =   4967
         TabIndex        =   148
         Top             =   370
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   688
         Caption         =   "2"
         ShowLed         =   -1  'True
         TextAlignment   =   1
      End
      Begin Audiostation.ButtonBig Button_TunerMemory 
         Height          =   390
         Index           =   0
         Left            =   4320
         TabIndex        =   147
         Top             =   370
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   688
         Caption         =   "1"
         ShowLed         =   -1  'True
         TextAlignment   =   1
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   6720
         ScaleHeight     =   195
         ScaleWidth      =   2595
         TabIndex        =   141
         Top             =   930
         Width           =   2655
         Begin VB.Label label_StreamStatus 
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
            Left            =   30
            TabIndex        =   142
            Tag             =   "1013"
            Top             =   0
            Width           =   2655
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
      Begin Audiostation.ButtonBig Button_TunerMemory 
         Height          =   390
         Index           =   2
         Left            =   5614
         TabIndex        =   149
         Top             =   370
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   688
         Caption         =   "3"
         ShowLed         =   -1  'True
         TextAlignment   =   1
      End
      Begin Audiostation.ButtonBig Button_TunerMemory 
         Height          =   390
         Index           =   3
         Left            =   6261
         TabIndex        =   150
         Top             =   370
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   688
         Caption         =   "4"
         ShowLed         =   -1  'True
         TextAlignment   =   1
      End
      Begin Audiostation.ButtonBig Button_TunerMemory 
         Height          =   390
         Index           =   4
         Left            =   6908
         TabIndex        =   151
         Top             =   370
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   688
         Caption         =   "5"
         ShowLed         =   -1  'True
         TextAlignment   =   1
      End
      Begin Audiostation.ButtonBig Button_TunerMemory 
         Height          =   390
         Index           =   5
         Left            =   7555
         TabIndex        =   152
         Top             =   370
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   688
         Caption         =   "6"
         ShowLed         =   -1  'True
         TextAlignment   =   1
      End
      Begin Audiostation.ButtonBig Button_TunerMemory 
         Height          =   390
         Index           =   6
         Left            =   8202
         TabIndex        =   153
         Top             =   370
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   688
         Caption         =   "7"
         ShowLed         =   -1  'True
         TextAlignment   =   1
      End
      Begin Audiostation.ButtonBig Button_TunerMemory 
         Height          =   390
         Index           =   7
         Left            =   8850
         TabIndex        =   154
         Top             =   370
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   688
         Caption         =   "8"
         ShowLed         =   -1  'True
         TextAlignment   =   1
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
      Picture         =   "Form_Main.frx":1EC461
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
            Picture         =   "Form_Main.frx":21B88F
            Top             =   0
            Width           =   240
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   0
            Picture         =   "Form_Main.frx":21BE19
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
         Picture         =   "Form_Main.frx":21C3A3
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
               Picture         =   "Form_Main.frx":2253FD
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
               Picture         =   "Form_Main.frx":2255DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Main.frx":225B77
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
            Picture         =   "Form_Main.frx":226111
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
            Picture         =   "Form_Main.frx":22669B
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
            Picture         =   "Form_Main.frx":226C65
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
            Picture         =   "Form_Main.frx":22722F
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
            Picture         =   "Form_Main.frx":2277B9
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
            Picture         =   "Form_Main.frx":227D43
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
            Picture         =   "Form_Main.frx":2282CD
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
            Picture         =   "Form_Main.frx":228897
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
            Picture         =   "Form_Main.frx":228E61
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
         Begin Audiostation.ButtonBig Button_EditDatTrack 
            Height          =   390
            Left            =   50
            TabIndex        =   140
            Tag             =   "1003"
            Top             =   410
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "Bewerken"
            TextAlignment   =   0
         End
         Begin Audiostation.ButtonBig cmdPlaylistDat 
            Height          =   390
            Left            =   50
            TabIndex        =   120
            Top             =   5
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "Playlist"
            TextAlignment   =   0
         End
         Begin Audiostation.ButtonBig cmdSettingsDat 
            Height          =   390
            Left            =   50
            TabIndex        =   121
            Top             =   810
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "Settings"
            TextAlignment   =   0
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
         Picture         =   "Form_Main.frx":2293EB
         Top             =   255
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Light_Dat_Pause_On 
         Height          =   150
         Left            =   6580
         Picture         =   "Form_Main.frx":2298C5
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
      Picture         =   "Form_Main.frx":229DAF
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
         Begin VB.CommandButton Button_MidiPlayer 
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
            Left            =   3960
            Picture         =   "Form_Main.frx":259959
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_MidiPlayer 
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
            Left            =   3480
            Picture         =   "Form_Main.frx":259F23
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_MidiPlayer 
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
            Left            =   1495
            Picture         =   "Form_Main.frx":25A4AD
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_MidiPlayer 
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
            Left            =   4500
            Picture         =   "Form_Main.frx":25AA77
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_MidiPlayer 
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
            Left            =   3000
            Picture         =   "Form_Main.frx":25B041
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_MidiPlayer 
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
            Left            =   2520
            Picture         =   "Form_Main.frx":25B5CB
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   50
            Width           =   495
         End
         Begin VB.CommandButton Button_MidiPlayer 
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
            Left            =   2040
            Picture         =   "Form_Main.frx":25BB55
            Style           =   1  'Graphical
            TabIndex        =   37
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
            TabIndex        =   118
            Top             =   20
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "Playlist"
            TextAlignment   =   0
         End
         Begin Audiostation.ButtonBig cmdSettingsMidi 
            Height          =   390
            Left            =   50
            TabIndex        =   119
            Top             =   800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "Settings"
            TextAlignment   =   0
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
         Picture         =   "Form_Main.frx":25C11F
         Top             =   555
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image6 
         Height          =   165
         Left            =   1920
         Picture         =   "Form_Main.frx":25C7A1
         Top             =   590
         Width           =   255
      End
      Begin VB.Image FloppyIn 
         Height          =   735
         Left            =   1480
         Picture         =   "Form_Main.frx":25CA1F
         Top             =   120
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.Image FloppyOut 
         Height          =   735
         Left            =   1480
         Picture         =   "Form_Main.frx":25F7B9
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
         Picture         =   "Form_Main.frx":262553
         Top             =   250
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Light_Midi_Pause_On 
         Height          =   150
         Left            =   6085
         Picture         =   "Form_Main.frx":262A2D
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
      Picture         =   "Form_Main.frx":262F17
      ScaleHeight     =   735
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   5
         Left            =   4320
         TabIndex        =   157
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
         Caption         =   "Clock"
         ShowLed         =   -1  'True
         Active          =   -1  'True
         TextAlignment   =   1
      End
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   4
         Left            =   5160
         TabIndex        =   158
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
         Caption         =   "Mixer"
         ShowLed         =   -1  'True
         Active          =   -1  'True
         TextAlignment   =   1
      End
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   3
         Left            =   6000
         TabIndex        =   159
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
         Caption         =   "CD"
         ShowLed         =   -1  'True
         Active          =   -1  'True
         TextAlignment   =   1
      End
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   1
         Left            =   7680
         TabIndex        =   160
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
         Caption         =   "MIDI"
         ShowLed         =   -1  'True
         Active          =   -1  'True
         TextAlignment   =   1
      End
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   2
         Left            =   6840
         TabIndex        =   161
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
         Caption         =   "DAT"
         ShowLed         =   -1  'True
         Active          =   -1  'True
         TextAlignment   =   1
      End
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   0
         Left            =   8640
         TabIndex        =   162
         Top             =   120
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   688
         Caption         =   "Power"
         TextAlignment   =   0
      End
      Begin VB.Image OptionsMenuButton 
         Height          =   405
         Left            =   120
         Picture         =   "Form_Main.frx":279E15
         Top             =   100
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   450
         Left            =   720
         Picture         =   "Form_Main.frx":27A65B
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
      Begin VB.Menu space05 
         Caption         =   "-"
      End
      Begin VB.Menu MenuItem_Popup_Properties 
         Caption         =   "&Properties"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnupopup_app 
      Caption         =   "- POPUP (2) -"
      Begin VB.Menu mnuplugins 
         Caption         =   "&Loaded plugins"
      End
      Begin VB.Menu mnuaudioplayersettings_popup 
         Caption         =   "&Audiostation Settings"
         HelpContextID   =   1062
      End
      Begin VB.Menu space04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuspectrumanalyzer 
         Caption         =   "&Show spectrum analyzer"
         Checked         =   -1  'True
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
Private AudioMaster As New AudioVolume

Dim DoorClose As Boolean
Dim AppInit As Boolean
Public Sub ResetMidiVU()
Dim i As Integer

For i = 0 To VU_Midi.count - 1
    VU_Midi(i).Position = 0
Next
End Sub
Private Sub SetPlaylistModeBasedOnSelected()
If mnuplaytrack_popup.Checked Then AudiostationMP3Player.PlayMode = PlaySingleTrack
If mnuautonext_popup.Checked Then AudiostationMP3Player.PlayMode = AutoNextTrack
If mnushuffle_popup.Checked Then AudiostationMP3Player.PlayMode = Shuffle

If mnurepeatplaylist_popup.Checked Then AudiostationMP3Player.PlaylistMode = RepeatPlaylist
If mnurepeattrack_popup.Checked Then AudiostationMP3Player.PlaylistMode = RepeatSingleTrack
End Sub
Private Sub AssignToMemorySlotAndSave(MemorySlot As Integer, Url As String, Optional Name As String = "")
If Url <> vbNullString Then
    Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "TunerMemory-" & MemorySlot, Url & "~" & Name)
    Button_TunerMemory(MemorySlot).Tag = Url & "~" & Name
End If
End Sub
Private Sub AniCD_Click()
Trm_CD_Animation.Enabled = True
End Sub

Private Sub Button_CDLoop_Click()
If Button_CDLoop.Active = False Then
    Button_CDLoop.Active = True
    Button_CDRandom.Active = False
    
    Light_Panel_CD.Picture = ImageList3.ListImages(3).Picture
Else
    Button_CDLoop.Active = False
    
    Light_Panel_CD.Picture = ImageList3.ListImages(2).Picture
End If
End Sub

Private Sub Button_CDOpen_Click()
Trm_CD_Animation.Enabled = True

If BASS_CD_DoorIsOpen(curdrive) Then
    Call BASS_CD_Door(curdrive, BASS_CD_DOOR_CLOSE)
Else
    Call BASS_CD_Door(curdrive, BASS_CD_DOOR_OPEN)
End If
End Sub

Private Sub Button_CDPlayer_Click(index As Integer)
Select Case index
    Case 0: Call AudiostationCDPlayer.PreviousTrack
    Case 1: Call AudiostationCDPlayer.Rewind
    Case 2: Call AudiostationCDPlayer.StopPlay
    Case 3: Call AudiostationCDPlayer.Play
    Case 4: Call AudiostationCDPlayer.Pause
    Case 5: Call AudiostationCDPlayer.Forward
End Select
End Sub

Private Sub Button_CDRandom_Click()
If Button_CDRandom.Active = False Then
    Button_CDRandom.Active = True
    Button_CDLoop.Active = False
    
    Light_Panel_CD.Picture = ImageList3.ListImages(4).Picture
Else
    Button_CDRandom.Active = False
    
    Light_Panel_CD.Picture = ImageList3.ListImages(2).Picture
End If
End Sub

Private Sub Button_EditDatTrack_Click()
Call Shell(App.path & "\windat.exe " & Chr(34) & AudiostationMP3Player.CurrentMediaFilename & Chr(34), vbNormalFocus)
End Sub

Private Sub Button_MidiPlayer_Click(index As Integer)
Select Case index
    Case 0: AudiostationMIDIPlayer.PreviousMidiTrack
    Case 1: AudiostationMIDIPlayer.RewindMidi10Seconds
    Case 2: AudiostationMIDIPlayer.StopMidiPlayback
    
    Case 3
        If AudiostationMIDIPlayer.MidiPlaylist.StorageContainer.count = 0 Then
            If MsgBox(GetLanguage(1023), vbQuestion + vbYesNo, "Playlist") = vbYes Then
                Form_Playlist.CurrentFormType = MidiPlayer
                Form_Playlist.Show vbModal
            End If
        Else
            AudiostationMIDIPlayer.StartMidiPlayback
        End If
    
    Case 4: AudiostationMIDIPlayer.PauseMidiPlayback
    Case 5: AudiostationMIDIPlayer.ForwardMidi10Seconds
    Case 6: AudiostationMIDIPlayer.NextMidiTrack
End Select
End Sub

Private Sub Button_OpenStream_Click()
Form_OpenStream.Show vbModal

If AudioStaStreamer.Url <> vbNullString Then
    Timer_Stream.Enabled = False
    If AudioStaStreamer.OpenStream(AudioStaStreamer.Url, AudioStaStreamer.Name) Then Timer_Stream.Enabled = True
    
    Dim i As Integer
    For i = 0 To Button_TunerMemory.count - 1
        Button_TunerMemory(i).Active = False
    Next
End If
End Sub

Private Sub Button_PlayStream_Click()
If AudioStaStreamer.Url = vbNullString Then
    If MsgBox("No stream is active" & vbNewLine & "Do you want to open a stream", vbQuestion + vbYesNo, "No stream") = vbYes Then
        Call Button_OpenStream_Click
    End If
Else
    Timer_Stream.Enabled = False
    If AudioStaStreamer.OpenStream(AudioStaStreamer.Url, AudioStaStreamer.Name) Then Timer_Stream.Enabled = True
End If
End Sub

Private Sub Button_Power_Click(index As Integer)
Select Case index
    Case 0
        CloseApplication = True
        Unload Me
    
    Case Else
        If Button_Power(index).Active Then
            'Turn off element
            Button_Power(index).Active = False
            Element(index).Tag = "DELETE"
            
            Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "Element-" & index, "OFF")
        Else
            'Turn on element
            Button_Power(index).Active = True
            Element(index).Tag = "ADD"
            
            Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "Element-" & index, "ON")
        End If
        
        If index = 3 And Button_Power(index).Active = True Then
            AniCD.Picture = ImageList5.ListImages(1).Picture
            AniCD.Visible = True
            
            Trm_CD_Animation.Tag = 1
            Trm_CD_Animation.Enabled = True
        End If
End Select
End Sub

Private Sub Button_StopStream_Click()
Call BASS_ChannelPause(chan)
End Sub

Private Sub Button_TunerMemory_Click(index As Integer)
Dim i As Integer

For i = 0 To Button_TunerMemory.count - 1
    Button_TunerMemory(i).Active = False
Next

If Button_TunerMemory(index).Tag = vbNullString Then
    If MsgBox("No stream has been assigned to the selected memory slot" & vbNewLine & "Do you wanto to assign the current stream?", vbQuestion + vbYesNo, "No stream") = vbYes Then
        Call AssignToMemorySlotAndSave(index, AudioStaStreamer.Url, AudioStaStreamer.Name)
    End If
Else
    Button_TunerMemory(index).Active = True
    
    Dim StreamUrl As String
    Dim StreamName As String
    
    StreamUrl = Extensions.Explode(Button_TunerMemory(index).Tag, "~", 0)
    StreamName = Extensions.Explode(Button_TunerMemory(index).Tag, "~", 1)
    
    Timer_Stream.Enabled = False
    Call AudioStaStreamer.OpenStream(StreamUrl, StreamName)
    Timer_Stream.Enabled = True
End If
End Sub

Private Sub cmdAudioPlayer_Click(index As Integer)
Select Case index
    Case 0: PopupMenu mnupopup_player
    
    Case 1
        ' Recording
    
    Case 2: AudiostationMP3Player.PreviousTrack
    Case 3: AudiostationMP3Player.Forward
    Case 4: AudiostationMP3Player.StopPlay
    Case 5
        If Mp3Playlist.StorageContainer.count = 0 Then
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

Private Sub cmdPlaylistDat_Click()
Form_Playlist.CurrentFormType = Mp3Player
Form_Playlist.Show , Me
End Sub

Private Sub CmdPlaylistMidi_Click()
Form_Playlist.CurrentFormType = MidiPlayer
Form_Playlist.Show , Me
End Sub

Private Sub cmdSettingsDat_Click()
Form_Settings.Show vbModal, Me
End Sub

Private Sub cmdSettingsMidi_Click()
Form_Settings.Show vbModal, Me
End Sub
Private Sub Digit_Time_Dat_OnClick()
If AudiostationMP3Player.ShowElapsedTime = True Then
    AudiostationMP3Player.ShowElapsedTime = False
Else
    AudiostationMP3Player.ShowElapsedTime = True
End If
End Sub

Private Sub Form_Load()
Width = 9900
Height = 9855 - 310

Slider_Dat_Left.value = 1000
Slider_Dat_Right.value = 1000
Slider_CD_Left.value = 1000
Slider_CD_Right.value = 1000

ChDrive App.path
ChDir App.path

' Check the correct BASS was loaded
If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
    MsgBox "An incorrect version of BASS.DLL was loaded", vbCritical
    End
End If

' Initialize BASS
If (BASS_Init(-1, 44100, 0, Me.hwnd, 0) = 0) Then
    MsgBox es & vbCrLf & vbCrLf & "error code: " & BASS_ErrorGetCode, vbExclamation, "Error"
    End
End If

Call BASS_SetConfig(BASS_CONFIG_NET_PLAYLIST, 1) ' enable playlist processing
Call BASS_SetConfig(BASS_CONFIG_NET_PREBUF, 0) ' minimize automatic pre-buffering, so we can do it (and display it) instead

Call AudiostationMP3Player.Init

Dim i As Integer
For i = 0 To Button_TunerMemory.count - 1
    Button_TunerMemory(i).Tag = Settings.ReadSetting("Sibra-Soft", "Audiostation", "TunerMemory-" & i, vbNullString)
Next

For i = 0 To Element.count - 1
    Element(i).Visible = True
Next

mnupopup_player.Visible = False
mnupopup_app.Visible = False

' Get the application settings
mnuspectrumanalyzer.Checked = Settings.ReadSetting("Sibra-Soft", "Audiostation", "UseSpectrumAnalyzer", True)

'Display program version
lbl_version.Caption = "Version: " & App.Major & "." & App.Minor & " Build: " & App.Revision

' Check if the CD-Rom drive exists
If Not CheckIfCDRomDriveExists Then
    Button_Power(3).Enabled = False
    
    Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "Element-3", "OFF")
    Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CD", 0)
End If

Light_Panel_CD.Picture = ImageList3.ListImages(1).Picture
cthread = 0

Call SetLanguage(Me)
Call SetPlaylistModeBasedOnSelected
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True

AudiostationMIDIPlayer.StopMidiPlayback
AudiostationMP3Player.StopPlay

Call ApplicationDestructor
Call Shell(App.path & "/close.exe", vbHide)

Cancel = False
End Sub

Private Sub lbl_Midi_Filename_Change()
lbl_Midi_Filename.ToolTipText = lbl_Midi_Filename.Caption
End Sub

Private Sub MenuItem_Popup_Properties_Click()
Form_Track_Properties.Show vbModal
End Sub

Private Sub mnuabout_popup_Click()
Form_About.Show vbModal
End Sub

Private Sub mnuaudioplayersettings_popup_Click()
Form_Settings.Show vbModal, Me
End Sub

Private Sub mnuautonext_popup_Click()
mnushuffle_popup.Checked = False
mnuautonext_popup.Checked = True
mnuplaytrack_popup.Checked = False

Call SetPlaylistModeBasedOnSelected
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

If Extensions.IsWebConnected Then
    webResponse = WebRequest.WebRequest("https://www.audiostation.org/app-deploy/audiostation/version.txt")
    
    webMinor = Extensions.Explode(webResponse, ".", "0")
    webMajor = Extensions.Explode(webResponse, ".", "1")
    webRevision = Extensions.Explode(webResponse, ".", "2")
    
    appMinor = App.Minor
    appMajor = App.Major
    appRevision = App.Revision
    
    If webMinor > appMinor Or webMajor = appMajor Or webRevision > appRevision Then
        iRet = MsgBox("A new version of Audiostation was found, do you want to download this new version?", vbYesNo + vbQuestion, "New version")
        
        If iRet = vbYes Then Shell "explorer.exe https://www.audiostation.org", vbNormalFocus
    Else
        MsgBox "You have the newest version of Audiostation", vbInformation
    End If
Else
    MsgBox "You are not connected to the internet, you must have a working internet connection to check for updates", vbExclamation, "Connection error"
End If
End Sub

Private Sub mnuclose_popup_Click()
End
End Sub

Private Sub mnuplaytrack_popup_Click()
mnushuffle_popup.Checked = False
mnuautonext_popup.Checked = False
mnuplaytrack_popup.Checked = True

Call SetPlaylistModeBasedOnSelected
End Sub

Private Sub mnuplugins_Click()
Form_Plugins.Show vbModal
End Sub

Private Sub mnurepeatplaylist_popup_Click()
mnurepeatplaylist_popup.Checked = True
mnurepeattrack_popup.Checked = False

Call SetPlaylistModeBasedOnSelected
End Sub

Private Sub mnurepeattrack_popup_Click()
mnurepeatplaylist_popup.Checked = False
mnurepeattrack_popup.Checked = True

Call SetPlaylistModeBasedOnSelected
End Sub

Private Sub mnushuffle_popup_Click()
mnushuffle_popup.Checked = True
mnuautonext_popup.Checked = False
mnuplaytrack_popup.Checked = False

Call SetPlaylistModeBasedOnSelected
End Sub

Private Sub mnuspectrumanalyzer_Click()
Dim i As Integer

If mnuspectrumanalyzer.Checked Then
    mnuspectrumanalyzer.Checked = False

    For i = 0 To VU_Spectrum.count - 1
        VU_Spectrum(i).Position = 0
    Next
Else
    mnuspectrumanalyzer.Checked = True
End If

Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "UseSpectrumAnalyzer", mnuspectrumanalyzer.Checked)
End Sub

Private Sub OptionsMenuButton_Click()
PopupMenu mnupopup_app
End Sub

Private Sub Slider_CD_Left_OnPositionChange()
'Call MediaPlayerForCD.SetLeftVolume(Slider_CD_Left.value)
End Sub

Private Sub Slider_CD_Right_OnPositionChange()
'Call MediaPlayerForCD.SetRightVolume(Slider_CD_Right.value)
End Sub

Private Sub Slider_Dat_Left_OnPositionChange()
Call BASS_ChannelSetAttribute(chan, BASS_ATTRIB_PAN, 0)
End Sub

Private Sub Slider_Dat_Right_OnPositionChange()
Debug.Print Slider_Dat_Right.value * 10
Call BASS_ChannelSetAttribute(chan, BASS_ATTRIB_PAN, -0.2)
End Sub

Private Sub Slider_Master_OnPositionChange()
Call BASS_SetVolume(Slider_Master.value / 100)
End Sub

Private Sub Slider_Midi_Left_OnPositionChange()
Form_Midi.MIDIOutput1.VolumeLeft = Slider_Midi_Left.value
End Sub

Private Sub Slider_Midi_Right_OnPositionChange()
Form_Midi.MIDIOutput1.VolumeRight = Slider_Midi_Right.value
End Sub

Private Sub Switch_CD_OnChange()
If Switch_CD.Active Then
    'MediaPlayerForCD.setAudioOn
Else
    'MediaPlayerForCD.setAudioOff
End If
End Sub

Private Sub Switch_Dat_OnChange()
If Switch_Dat.Active Then
    Call BASS_ChannelSetAttribute(chan, BASS_ATTRIB_VOL, 1)
Else
    Call BASS_ChannelSetAttribute(chan, BASS_ATTRIB_VOL, 0)
End If
End Sub

Private Sub Switch_Master_OnChange()
If Switch_Master.Active Then
    Call AudioMaster.SetMute(0)
Else
    Call AudioMaster.SetMute(1)
End If
End Sub

Private Sub Timer_Stream_Timer()
Dim progress As Long

progress = BASS_StreamGetFilePosition(chan, BASS_FILEPOS_BUFFER) * 100 / BASS_StreamGetFilePosition(chan, BASS_FILEPOS_END)    ' percentage of buffer filled

If (progress > 75 Or BASS_StreamGetFilePosition(chan, BASS_FILEPOS_CONNECTED) = 0) Then  ' over 75% full (or end of download)
    Timer_Stream.Enabled = False
    
    Call DoMeta
    
    Call BASS_ChannelSetSync(chan, BASS_SYNC_META, 0, AddressOf MetaSync, 0)
    Call BASS_ChannelSetSync(chan, BASS_SYNC_END, 0, AddressOf EndSync, 0)
    Call BASS_ChannelPlay(chan, True)
    
    If AudioStaStreamer.Error Then
        label_StreamStatus.Caption = "Error: Can't play stream"
    Else
        If AudioStaStreamer.Name <> vbNullString Then
            label_StreamStatus.Caption = "Playing: " & AudioStaStreamer.Name
        Else
            label_StreamStatus.Caption = "Playing"
        End If
    End If
Else
    Label_StreamTitle.Caption = "Nothing playing"
    label_StreamStatus.Caption = "Buffering... " & progress & "%"
End If
End Sub

Private Sub Trm_Animation_Timer()
Dim J As Integer

J = Trm_Animation.Tag
Picture11.Picture = ImageList1.ListImages.Item(J).Picture

If Trm_Animation.Tag = ImageList1.ListImages.count Then
    Trm_Animation.Tag = 1
Else
    Trm_Animation.Tag = Trm_Animation.Tag + 1
End If
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

Private Sub Trm_Check_File_Timer()
Dim MediaFile As String
Dim MediaIndex As String
Dim MediaDuration As String

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
            
            AudiostationMIDIPlayer.StopMidiPlayback
            AudiostationCDPlayer.StopPlay
            
            If Mp3Playlist.IsExistingItem(MediaFile) > 0 Then
                AudiostationMP3Player.CurrentTrackNumber = Mp3Playlist.IsExistingItem(MediaFile)
                AudiostationMP3Player.StartPlay
            Else
                MediaIndex = Format(Mp3Playlist.StorageContainer.count + 1, "00")
                
                ' Only get the duration when it's a mp3 file
                If LCase(Right(MediaFile, 3)) = "mp3" Then
                    Mp3Info.Filename = MediaFile
                    MediaDuration = Extensions.TimeString(Mp3Info.SongLength)
                End If
                
                If MediaDuration = "0" Then: MediaDuration = "-"
                
                Mp3Playlist.AddToStorage MediaFile, MediaIndex & ";" & MediaFile & ";" & MediaDuration
                
                AudiostationMP3Player.CurrentTrackNumber = Mp3Playlist.StorageContainer.count
                AudiostationMP3Player.StartPlay
            End If
                        
        Case "mid", "kar", "mus", "sid"
            AudiostationMP3Player.StopPlay
            AudiostationCDPlayer.StopPlay
            
            CurrentIndex = Format(MidiPlaylist.StorageContainer.count + 1, "00")
            CurrentMediaDuration = "-"
    
            MidiPlaylist.AddToStorage MediaFile, CurrentIndex & ";" & MediaFile & ";" & CurrentMediaDuration
            
            AudiostationMIDIPlayer.MidiTrackNr = MidiPlaylist.StorageContainer.count
            AudiostationMIDIPlayer.StartMidiPlayback
        
    Case "apl", "wpl", "m3u", "pls" 'Playlist files
        If Not (Dir(MediaFile, vbDirectory) = vbNullString) Then
            Screen.MousePointer = vbHourglass
            
            Select Case LCase(Right(MediaFile, 3))
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
If PlayStateMediaMode = MidiMediaMode Then
    If Form_Midi.HScrollPlayerTime.value = Form_Midi.HScrollPlayerTime.max Then PlayState = MediaEnded

    If PlayState = Playing Then
        Trm_Floppy_Drive_Light.Enabled = True
        FloppyIn.Visible = True
        
        If Light_Midi_Play_On.Visible = True Then
            Light_Midi_Play_On.Visible = False
        Else
            Light_Midi_Play_On.Visible = True
        End If
        
    ElseIf PlayState = Stopped Or MediaEnded Then
        Trm_Floppy_Drive_Light.Enabled = False
        FloppyIn.Visible = False
        
        Light_Midi_Floppy_Drive.Visible = False
        Light_Midi_Play_On.Visible = False
        Light_Midi_Pause_On.Visible = False
        Light_Midi_Play_On.Visible = False
    
    ElseIf PlayState = Paused Then
        Trm_Floppy_Drive_Light.Enabled = False
        Light_Midi_Floppy_Drive.Visible = False
        FloppyIn.Visible = False
        
        If Light_Midi_Pause_On.Visible = True Then
            Light_Midi_Pause_On.Visible = False
        Else
            Light_Midi_Pause_On.Visible = True
        End If
    
    End If
End If
End Sub

Private Sub Trm_Lights_Timer()
' Media Player
If PlayStateMediaMode = MP3MediaMode Then
    If PlayState = Playing Then
        Trm_Animation.Enabled = True
        Light_Dat_Pause_On.Visible = False
        
        If Light_Dat_Play_On.Visible = True Then
            Light_Dat_Play_On.Visible = False
        Else
            Light_Dat_Play_On.Visible = True
        End If
    Else
        If PlayState = Paused Then
            Trm_Animation.Enabled = False
            Light_Dat_Play_On.Visible = True
        
            VU_Left.Position = 0
            VU_Right.Position = 0
        
            If Light_Dat_Pause_On.Visible = True Then
                Light_Dat_Pause_On.Visible = False
            Else
                Light_Dat_Pause_On.Visible = True
            End If
        Else
            Light_Dat_Pause_On.Visible = False
            Light_Dat_Play_On.Visible = False
            Trm_Animation.Enabled = False
        End If
    End If
End If

' CD Player
If PlayStateMediaMode = CDMediaMode Then
    If PlayState = Playing Then
        If Light_CD_Play_On.Visible = True Then
            Light_CD_Play_On.Visible = False
        Else
            Light_CD_Play_On.Visible = True
        End If
    Else
        Light_CD_Play_On.Visible = False
    End If
End If
End Sub
Private Sub Trm_Main_Timer()
Dim length, pos As Long
Dim Totaltime, Elapsedtime, Remainingtime  As Double
Dim MidiPos As Long

Digit_Clock.Hours = Format(Now, "hh")
Digit_Clock.Minutes = Format(Now, "nn")
Digit_Clock.seconds = Format(Now, "ss")

Digit_Track_Dat.value = AudiostationMP3Player.CurrentTrackNumber
Digit_Track_Midi.value = AudiostationMIDIPlayer.MidiTrackNr

If Form_Midi.LabelQueueTime.Caption = "(wait)" Then
    
Else
    If Left(Form_Midi.LabelQueueTime.Caption, 1) = "." Then
        MidiPos = 0
    Else
        MidiPos = Extensions.Explode(Form_Midi.LabelQueueTime.Caption, ".", 0)
    End If
End If

Digit_Time_Midi.seconds = Extensions.Explode(Extensions.TimeString(MidiPos), ":", 1)
Digit_Time_Midi.Minutes = Extensions.Explode(Extensions.TimeString(MidiPos), ":", 0)

If AudiostationMIDIPlayer.MidiFilename = vbNullString Then
    lbl_Midi_Filename.Caption = "Unknown"
Else
    lbl_Midi_Filename.Caption = AudiostationMIDIPlayer.MidiFilename
End If

' Enable the activated rack
Dim i As Integer
For i = 1 To Button_Power.count - 1
    Dim mustBeOff As String
    
    mustBeOff = Settings.ReadSetting("Sibra-Soft", "Audiostation", "Element-" & i, "OFF")
        
    If Not mustBeOff = "OFF" Then
        If i = 4 Then: ElementOff(6).Visible = False
        
        ElementOff(i).Visible = False
        Button_Power(i).Active = True
    Else
        If i = 4 Then: ElementOff(6).Visible = True
        
        ElementOff(i).Visible = True
        Button_Power(i).Active = False
    End If
Next

' Startup loop
If Trm_Main.Tag = 6 Then
    Trm_Main.Interval = 1
Else
    Button_Power(Trm_Main.Tag).Active = True
    Trm_Main.Tag = Trm_Main.Tag + 1
End If

' Show the elapsed or leftover time
If PlayStateMediaMode = MP3MediaMode Then
    If PlayState = Playing Then
        lbl_Filename.Caption = Extensions.GetFileNameFromFilePath(AudiostationMP3Player.CurrentMediaFilename, False)
        lbl_Filename.ToolTipText = Extensions.GetFileNameFromFilePath(AudiostationMP3Player.CurrentMediaFilename, False)
        
        Dim TimeSerial As String
        
        length = BASS_ChannelGetLength(chan, BASS_POS_BYTE)
        pos = BASS_ChannelGetPosition(chan, BASS_POS_BYTE)
        Totaltime = BASS_ChannelBytes2Seconds(chan, length)
        Elapsedtime = BASS_ChannelBytes2Seconds(chan, pos)
        Remainingtime = Totaltime - Elapsedtime
            
        If AudiostationMP3Player.ShowElapsedTime Then
            TimeSerial = Extensions.SecondsToTimeSerial(Elapsedtime, SmallTimeSerial)
    
            Digit_Time_Dat.Minutes = Extensions.Explode(TimeSerial, ":", 0)
            Digit_Time_Dat.seconds = Extensions.Explode(TimeSerial, ":", 1)
        Else
            TimeSerial = Extensions.SecondsToTimeSerial(Remainingtime, SmallTimeSerial)
    
            Digit_Time_Dat.Minutes = Extensions.Explode(TimeSerial, ":", 0)
            Digit_Time_Dat.seconds = Extensions.Explode(TimeSerial, ":", 1)
        End If
    End If
    
    If CurrentMediaFilename <> vbNullString Then: MenuItem_Popup_Properties.Enabled = True
    
    If AudiostationMP3Player.PlayState = Playing And Remainingtime = 0 Then AudiostationMP3Player.PlayState = MediaEnded
    If AudiostationMP3Player.PlayState = MediaEnded Then AudiostationMP3Player.NextTrack
End If
End Sub
Private Sub Trm_VU_Timer()
If PlayStateMediaMode = MP3MediaMode Then
    Dim level As Long
    Dim leftVU, rightVU As Long
    
    level = BASS_ChannelGetLevel(chan)
    
    leftVU = LoWord(level) * 2
    rightVU = HiWord(level) * 2
    
    If PlayState = Playing Then
        VU_Left.Position = leftVU / 32768 * 100
        VU_Right.Position = rightVU / 32768 * 100
        
        VU_Master_Peak.Position = leftVU / 32768 * 100
    Else
        VU_Master_Peak.Position = 0
        VU_Left.Position = 0
        VU_Right.Position = 0
    End If
    
    If mnuspectrumanalyzer.Checked Then: UpdateSpectrum
Else
    VU_Master_Peak.Position = 0
    VU_Left.Position = 0
    VU_Right.Position = 0
End If

On Error Resume Next
If PlayStateMediaMode = MidiMediaMode And Form_Midi.VIndicator1.count > 1 Then
    For i = 0 To VU_Midi.count - 1
        VU_Midi(i).Position = Form_Midi.VIndicator1(i).value
    Next
End If
End Sub
