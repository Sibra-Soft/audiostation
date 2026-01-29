VERSION 5.00
Object = "{0A362340-2E5E-11D3-85BF-00105AC8B715}#1.0#0"; "ISDIGI~1.OCX"
Object = "{C5412DA5-2E2F-11D3-85BF-00105AC8B715}#1.0#0"; "ISANAL~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5F5C69A3-5434-4A28-B392-38259F02830A}#1.0#0"; "DataInter.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{40F6D89D-D6BF-4EAD-B885-E1869BDF4E31}#41.0#0"; "AdioLibrary.ocx"
Object = "{966CF34C-191F-4FB6-BF33-C8DB07C6A40D}#1.0#0"; "DigitBox.ocx"
Begin VB.Form Form_Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiostation"
   ClientHeight    =   10005
   ClientLeft      =   4695
   ClientTop       =   1275
   ClientWidth     =   12960
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
   ScaleHeight     =   10005
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_MidiVu 
      Interval        =   30
      Left            =   10440
      Top             =   1080
   End
   Begin AdioLibrary.AdioCore AdioCore 
      Left            =   120
      Top             =   9240
      _ExtentX        =   10186
      _ExtentY        =   873
      Begin AdioLibrary.AdioMidiPlayer AdioMidiPlayer 
         Left            =   3000
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioMediaPlayer AdioMediaPlayer 
         Left            =   2400
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioRecorder AdioRecorder 
         Left            =   1800
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioTagging AdioTagging 
         Left            =   3840
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioCDPlayer AdioCDPlayer 
         Left            =   1200
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioAudioPeak AdioAudioPeak 
         Left            =   600
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioPlaylist AdioMidiPlaylist 
         Left            =   5280
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         AllowDuplicateItems=   0   'False
      End
      Begin AdioLibrary.AdioPlaylist AdioMediaPlaylist 
         Left            =   4680
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         AllowDuplicateItems=   0   'False
      End
   End
   Begin VB.PictureBox PictureBox_Disabled 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Index           =   0
      Left            =   12000
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   62
      Top             =   2520
      Width           =   495
   End
   Begin DataInter.uDataInter DataInter 
      Left            =   12000
      Top             =   1920
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.PictureBox Element 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   6
      Left            =   120
      Picture         =   "Form_Main.frx":088B
      ScaleHeight     =   855
      ScaleWidth      =   9615
      TabIndex        =   82
      Tag             =   "OFF"
      Top             =   8400
      Width           =   9615
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   80
         ScaleHeight     =   435
         ScaleWidth      =   9345
         TabIndex        =   83
         Top             =   55
         Width           =   9400
         Begin DigitBox.DigitBoxControl DigitBox_Clock 
            Height          =   405
            Left            =   15
            TabIndex        =   117
            Top             =   15
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   714
            DigitDisplay    =   "00:00:00"
            DigitSize       =   1
            DigitColor      =   0
            DigitOutLine    =   0   'False
            DigitJustify    =   0
            DigitFormat     =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright © 2009 - 2025 Sibra-Soft"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1920
            TabIndex        =   85
            Top             =   195
            Width           =   3090
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
            Left            =   1920
            TabIndex        =   84
            Top             =   15
            Width           =   180
         End
      End
   End
   Begin VB.Timer Timer_Lights 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9840
      Top             =   120
   End
   Begin VB.Timer Timer_MediaAnimation 
      Enabled         =   0   'False
      Interval        =   110
      Left            =   10440
      Tag             =   "1"
      Top             =   2760
   End
   Begin VB.Timer Timer_Main 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10320
      Tag             =   "1"
      Top             =   120
   End
   Begin MSComctlLib.ImageList Imagelist_CDDisplay 
      Left            =   12000
      Top             =   120
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
            Picture         =   "Form_Main.frx":17789
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":1921B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":1ACAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":1C73F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":1E1D1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Imagelist_CDAnimation 
      Left            =   12000
      Top             =   720
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
            Picture         =   "Form_Main.frx":1FC63
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":280AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":380F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":48135
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":58179
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":681BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":78201
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Trm_CD_Animation 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   10440
      Tag             =   "1"
      Top             =   4320
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
      Picture         =   "Form_Main.frx":88245
      ScaleHeight     =   1575
      ScaleWidth      =   9615
      TabIndex        =   1
      Tag             =   "OFF"
      Top             =   5340
      Width           =   9615
      Begin isAnalogLibrary.iLabelX iLabelX5 
         Height          =   210
         Left            =   6600
         TabIndex        =   116
         Top             =   1185
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
      Begin isAnalogLibrary.iLabelX iLabelX3 
         Height          =   210
         Left            =   4440
         TabIndex        =   152
         Top             =   1185
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
      Begin isAnalogLibrary.iLabelX iLabelX4 
         Height          =   210
         Left            =   5520
         TabIndex        =   108
         Top             =   1185
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
      Begin isAnalogLibrary.iLabelX ILaMaster 
         Height          =   195
         Left            =   120
         TabIndex        =   73
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
         TabIndex        =   67
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
      Begin Audiostation.MixSlider Slider_Master 
         Height          =   1335
         Left            =   360
         TabIndex        =   72
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
         Max             =   100
      End
      Begin VB.PictureBox Picture1 
         Height          =   1215
         Left            =   9160
         ScaleHeight     =   1155
         ScaleWidth      =   195
         TabIndex        =   65
         Top             =   120
         Width           =   255
         Begin isAnalogLibrary.iLedBarX LedBar_Master 
            Height          =   1215
            Left            =   0
            TabIndex        =   66
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
      Begin Audiostation.MixSlider Slider_Record 
         Height          =   1335
         Left            =   3360
         TabIndex        =   74
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
         Max             =   100
      End
      Begin Audiostation.MixSlider Slider_CD_Volume 
         Height          =   1335
         Left            =   4320
         TabIndex        =   75
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
         Max             =   100
      End
      Begin Audiostation.MixSlider Slider_CD_Balance 
         Height          =   1215
         Left            =   4755
         TabIndex        =   76
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2143
         Value           =   0
         Min             =   -1000
         Max             =   1000
      End
      Begin Audiostation.MixSlider Slider_Dat_Volume 
         Height          =   1335
         Left            =   5400
         TabIndex        =   77
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
         Max             =   100
      End
      Begin Audiostation.MixSlider Slider_Dat_Balance 
         Height          =   1215
         Left            =   5835
         TabIndex        =   78
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2143
         Value           =   0
         Min             =   -1000
         Max             =   1000
      End
      Begin Audiostation.MixSlider Slider_Midi_Volume 
         Height          =   1335
         Left            =   6720
         TabIndex        =   79
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2355
         Max             =   100
      End
      Begin isDigitalLibrary.iSwitchLedX Switch_Master 
         Height          =   270
         Left            =   7560
         TabIndex        =   118
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
         Glyph           =   "Form_Main.frx":B9BDF
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
         TabIndex        =   119
         Top             =   400
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
         Glyph           =   "Form_Main.frx":B9C35
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
         TabIndex        =   120
         Top             =   720
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
         Glyph           =   "Form_Main.frx":B9C8B
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
         TabIndex        =   150
         Top             =   610
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
         Glyph           =   "Form_Main.frx":B9CE1
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
         TabIndex        =   151
         Top             =   930
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
         Glyph           =   "Form_Main.frx":B9D37
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
      Begin VB.Image Image1 
         Height          =   960
         Left            =   8640
         Picture         =   "Form_Main.frx":B9D8D
         Top             =   240
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   1005
         Left            =   2750
         Picture         =   "Form_Main.frx":BB2CF
         Top             =   240
         Width           =   345
      End
   End
   Begin MSComctlLib.ImageList Imagelist_MediaPlayerAnimation 
      Left            =   12000
      Top             =   1320
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
            Picture         =   "Form_Main.frx":BC5E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":BC7D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":BC9C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":BCBB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":BCDA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":BCF94
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":BD187
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Main.frx":BD377
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
      Picture         =   "Form_Main.frx":BD563
      ScaleHeight     =   1500
      ScaleWidth      =   9615
      TabIndex        =   2
      Tag             =   "OFF"
      Top             =   3840
      Width           =   9615
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
         TabIndex        =   60
         Top             =   840
         Width           =   2415
         Begin Audiostation.ButtonBig Button_CDRandom 
            Height          =   390
            Left            =   40
            TabIndex        =   80
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
            TabIndex        =   81
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
         TabIndex        =   52
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
         Left            =   4410
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
            Index           =   6
            Left            =   4500
            Picture         =   "Form_Main.frx":EC991
            Style           =   1  'Graphical
            TabIndex        =   6
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
            Index           =   7
            Left            =   930
            Picture         =   "Form_Main.frx":ECF5B
            Style           =   1  'Graphical
            TabIndex        =   114
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
            Index           =   5
            Left            =   3960
            Picture         =   "Form_Main.frx":ED4E5
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
            Picture         =   "Form_Main.frx":EDAAF
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
            Picture         =   "Form_Main.frx":EE039
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
            Picture         =   "Form_Main.frx":EE5C3
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
            Picture         =   "Form_Main.frx":EEB4D
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
            Picture         =   "Form_Main.frx":EF117
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   50
            Width           =   495
         End
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
      End
      Begin DigitBox.DigitBoxControl DigitBox_CDDuration 
         Height          =   405
         Left            =   7300
         TabIndex        =   111
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         DigitDisplay    =   "00:00"
         DigitSize       =   1
         DigitColor      =   1
         DigitOutLine    =   0   'False
         DigitJustify    =   0
         DigitFormat     =   1
      End
      Begin DigitBox.DigitBoxControl DigitBox_CDTrack 
         Height          =   405
         Left            =   5880
         TabIndex        =   113
         Top             =   180
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   714
         DigitDisplay    =   "00"
         DigitPlaceHolders=   2
         DigitSize       =   1
         DigitColor      =   1
         DigitOutLine    =   0   'False
         DigitJustify    =   0
      End
      Begin VB.Image Light_CD_Play_On 
         Height          =   135
         Left            =   6900
         Picture         =   "Form_Main.frx":EF6E1
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Light_CD_Pause_On 
         Height          =   150
         Left            =   6895
         Picture         =   "Form_Main.frx":EFBBB
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
      Height          =   1575
      Index           =   5
      Left            =   120
      Picture         =   "Form_Main.frx":F00A5
      ScaleHeight     =   1575
      ScaleWidth      =   9615
      TabIndex        =   64
      Tag             =   "OFF"
      Top             =   6900
      Width           =   9615
      Begin isAnalogLibrary.iLabelX iLabelX1 
         Height          =   210
         Left            =   4200
         TabIndex        =   91
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
      Begin isAnalogLibrary.iLabelX Label_StreamTitle 
         Height          =   210
         Left            =   6120
         TabIndex        =   100
         Top             =   120
         Width           =   3135
         AutoSize        =   0   'False
         Alignment       =   2
         BorderStyle     =   0
         Caption         =   "T(1054)"
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
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   1050
         Left            =   120
         ScaleHeight     =   990
         ScaleWidth      =   3720
         TabIndex        =   107
         Top             =   200
         Width           =   3780
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   0
            Left            =   120
            TabIndex        =   121
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   1
            Left            =   240
            TabIndex        =   122
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   2
            Left            =   360
            TabIndex        =   123
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   3
            Left            =   480
            TabIndex        =   124
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   4
            Left            =   600
            TabIndex        =   125
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   5
            Left            =   720
            TabIndex        =   126
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   6
            Left            =   840
            TabIndex        =   127
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   7
            Left            =   960
            TabIndex        =   128
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   8
            Left            =   1080
            TabIndex        =   129
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   9
            Left            =   1200
            TabIndex        =   130
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   10
            Left            =   1320
            TabIndex        =   131
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   11
            Left            =   1440
            TabIndex        =   132
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   12
            Left            =   1560
            TabIndex        =   133
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   13
            Left            =   1680
            TabIndex        =   134
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   14
            Left            =   1800
            TabIndex        =   135
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   15
            Left            =   1920
            TabIndex        =   136
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   16
            Left            =   2040
            TabIndex        =   137
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   17
            Left            =   2160
            TabIndex        =   138
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   18
            Left            =   2280
            TabIndex        =   139
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   19
            Left            =   2400
            TabIndex        =   140
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   20
            Left            =   2520
            TabIndex        =   141
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   21
            Left            =   2640
            TabIndex        =   142
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   22
            Left            =   2760
            TabIndex        =   143
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   23
            Left            =   2880
            TabIndex        =   144
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   24
            Left            =   3000
            TabIndex        =   145
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   25
            Left            =   3120
            TabIndex        =   146
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   26
            Left            =   3240
            TabIndex        =   147
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   27
            Left            =   3360
            TabIndex        =   148
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
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
            Object.Width           =   7
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
         Begin isAnalogLibrary.iLedBarX LedBar_Spectrum 
            Height          =   975
            Index           =   28
            Left            =   3480
            TabIndex        =   149
            Top             =   15
            Width           =   105
            SegmentDirection=   0
            SegmentMargin   =   0
            SegmentSize     =   3
            SegmentSpacing  =   2
            SegmentStyle    =   0
            BackGroundColor =   0
            BorderStyle     =   0
            SectionColor1   =   33023
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
            Object.Height          =   65
            FillReferenceValue=   0
            FillReferenceEnabled=   0   'False
            SectionColor4   =   65535
            SectionColor5   =   65535
            SectionEnd3     =   0
            SectionEnd4     =   0
            OPCItemCount    =   0
         End
      End
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
         Left            =   5400
         Picture         =   "Form_Main.frx":11F4D3
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Open stream"
         Top             =   870
         Width           =   495
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
         Left            =   4680
         Picture         =   "Form_Main.frx":11F69D
         Style           =   1  'Graphical
         TabIndex        =   89
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
         Left            =   4200
         Picture         =   "Form_Main.frx":11FC27
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Play"
         Top             =   870
         Width           =   495
      End
      Begin Audiostation.ButtonBig Button_TunerMemory 
         Height          =   390
         Index           =   1
         Left            =   4845
         TabIndex        =   93
         Top             =   375
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
         Left            =   4200
         TabIndex        =   92
         Top             =   375
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
         Left            =   6600
         ScaleHeight     =   195
         ScaleWidth      =   2595
         TabIndex        =   86
         Top             =   930
         Width           =   2655
         Begin VB.Label label_StreamStatus 
            BackStyle       =   0  'Transparent
            Caption         =   "T(1019)"
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
            TabIndex        =   87
            Tag             =   "1013"
            Top             =   0
            Width           =   2655
         End
      End
      Begin Audiostation.ButtonBig Button_TunerMemory 
         Height          =   390
         Index           =   2
         Left            =   5490
         TabIndex        =   94
         Top             =   375
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
         Left            =   6135
         TabIndex        =   95
         Top             =   375
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
         Left            =   6795
         TabIndex        =   96
         Top             =   375
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
         Left            =   7440
         TabIndex        =   97
         Top             =   375
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
         Left            =   8085
         TabIndex        =   98
         Top             =   375
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
         Left            =   8730
         TabIndex        =   99
         Top             =   375
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
      Picture         =   "Form_Main.frx":1201B1
      ScaleHeight     =   1500
      ScaleWidth      =   9615
      TabIndex        =   12
      Tag             =   "OFF"
      Top             =   2340
      Width           =   9615
      Begin DigitBox.DigitBoxControl DigitBox_DatDuration 
         Height          =   405
         Left            =   4920
         TabIndex        =   109
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         DigitDisplay    =   "00:00"
         DigitSize       =   1
         DigitColor      =   1
         DigitOutLine    =   0   'False
         DigitJustify    =   0
         DigitFormat     =   1
      End
      Begin VB.PictureBox Picturebox_Recording 
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
         TabIndex        =   59
         Top             =   200
         Visible         =   0   'False
         Width           =   255
         Begin VB.Image Image_Recording 
            Height          =   240
            Left            =   0
            Picture         =   "Form_Main.frx":14F5DF
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   0
            Picture         =   "Form_Main.frx":14FB69
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
         Left            =   1440
         Picture         =   "Form_Main.frx":1500F3
         ScaleHeight     =   1305
         ScaleWidth      =   2055
         TabIndex        =   53
         Top             =   50
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
            TabIndex        =   57
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
               TabIndex        =   58
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
            TabIndex        =   54
            Top             =   240
            Width           =   1695
            Begin VB.PictureBox Picture_MediaPlayerAnimation 
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
               ScaleHeight     =   360
               ScaleWidth      =   1560
               TabIndex        =   55
               Top             =   160
               Width           =   1560
            End
            Begin VB.Label Label_FilenameDat 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "T(1019)"
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
               TabIndex        =   56
               Tag             =   "1013"
               Top             =   0
               UseMnemonic     =   0   'False
               Width           =   1485
            End
         End
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
         TabIndex        =   14
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
            Picture         =   "Form_Main.frx":15914D
            Style           =   1  'Graphical
            TabIndex        =   63
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
            Picture         =   "Form_Main.frx":1596D7
            Style           =   1  'Graphical
            TabIndex        =   16
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
            Picture         =   "Form_Main.frx":159CA1
            Style           =   1  'Graphical
            TabIndex        =   15
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
            Picture         =   "Form_Main.frx":15A26B
            Style           =   1  'Graphical
            TabIndex        =   17
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
            Picture         =   "Form_Main.frx":15A7F5
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
            Index           =   4
            Left            =   3240
            Picture         =   "Form_Main.frx":15AD7F
            Style           =   1  'Graphical
            TabIndex        =   19
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
            Picture         =   "Form_Main.frx":15B309
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
            Index           =   3
            Left            =   2760
            Picture         =   "Form_Main.frx":15B8D3
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
            Index           =   1
            Left            =   1080
            Picture         =   "Form_Main.frx":15BE9D
            Style           =   1  'Graphical
            TabIndex        =   61
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
         TabIndex        =   13
         Top             =   100
         Width           =   1455
         Begin Audiostation.ButtonBig cmdPlaylistDat 
            Height          =   390
            Left            =   50
            TabIndex        =   70
            Top             =   5
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "T(1001)"
            TextAlignment   =   0
         End
         Begin Audiostation.ButtonBig cmdSettingsDat 
            Height          =   390
            Left            =   50
            TabIndex        =   71
            Top             =   810
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "T(1002)"
            TextAlignment   =   0
         End
      End
      Begin isAnalogLibrary.iLedBarX LedBar_DatLeft 
         Height          =   135
         Left            =   7310
         TabIndex        =   22
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
      Begin isAnalogLibrary.iLedBarX LedBar_DatRight 
         Height          =   135
         Left            =   7310
         TabIndex        =   23
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
      Begin DigitBox.DigitBoxControl DigitBox_DatTrack 
         Height          =   405
         Left            =   3860
         TabIndex        =   110
         Top             =   180
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   714
         DigitDisplay    =   "00"
         DigitPlaceHolders=   2
         DigitSize       =   1
         DigitColor      =   1
         DigitOutLine    =   0   'False
      End
      Begin VB.Image Light_Dat_Play_On 
         Height          =   135
         Left            =   6590
         Picture         =   "Form_Main.frx":15C427
         Top             =   255
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Light_Dat_Pause_On 
         Height          =   150
         Left            =   6580
         Picture         =   "Form_Main.frx":15C901
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
      Picture         =   "Form_Main.frx":15CDEB
      ScaleHeight     =   1500
      ScaleWidth      =   9615
      TabIndex        =   24
      Tag             =   "OFF"
      Top             =   840
      Width           =   9615
      Begin DigitBox.DigitBoxControl DigitBox_MidiDuration 
         Height          =   315
         Left            =   8420
         TabIndex        =   115
         Top             =   230
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         DigitDisplay    =   "00:00"
         DigitColor      =   1
         DigitOutLine    =   0   'False
         DigitFormat     =   1
      End
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
         TabIndex        =   34
         Top             =   200
         Width           =   1935
         Begin isAnalogLibrary.iLedBarX VU_Midi 
            Height          =   400
            Index           =   0
            Left            =   0
            TabIndex        =   35
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
            TabIndex        =   36
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
            TabIndex        =   37
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
            TabIndex        =   38
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
            TabIndex        =   39
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
            Index           =   6
            Left            =   720
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
            Index           =   7
            Left            =   840
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
            Index           =   8
            Left            =   960
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
            Index           =   9
            Left            =   1080
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
            Index           =   10
            Left            =   1200
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
            Index           =   11
            Left            =   1320
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
            Index           =   12
            Left            =   1440
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
            Index           =   13
            Left            =   1560
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
            Index           =   14
            Left            =   1680
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
            Index           =   15
            Left            =   1800
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
         Left            =   4410
         ScaleHeight     =   495
         ScaleWidth      =   5055
         TabIndex        =   26
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
            Picture         =   "Form_Main.frx":18C995
            Style           =   1  'Graphical
            TabIndex        =   27
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
            Picture         =   "Form_Main.frx":18CF5F
            Style           =   1  'Graphical
            TabIndex        =   29
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
            Picture         =   "Form_Main.frx":18D4E9
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
            Index           =   6
            Left            =   4500
            Picture         =   "Form_Main.frx":18DAB3
            Style           =   1  'Graphical
            TabIndex        =   28
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
            Picture         =   "Form_Main.frx":18E07D
            Style           =   1  'Graphical
            TabIndex        =   30
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
            Picture         =   "Form_Main.frx":18E607
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
            Index           =   1
            Left            =   2040
            Picture         =   "Form_Main.frx":18EB91
            Style           =   1  'Graphical
            TabIndex        =   33
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
         TabIndex        =   25
         Top             =   100
         Width           =   1335
         Begin Audiostation.ButtonBig cmdPlaylistMidi 
            Height          =   390
            Left            =   50
            TabIndex        =   68
            Top             =   20
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "T(1001)"
            TextAlignment   =   0
         End
         Begin Audiostation.ButtonBig cmdSettingsMidi 
            Height          =   390
            Left            =   50
            TabIndex        =   69
            Top             =   800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            Caption         =   "T(1002)"
            TextAlignment   =   0
         End
      End
      Begin DigitBox.DigitBoxControl DigitBox_MidiTrack 
         Height          =   405
         Left            =   5110
         TabIndex        =   112
         Top             =   180
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   714
         DigitDisplay    =   "00"
         DigitPlaceHolders=   2
         DigitSize       =   1
         DigitColor      =   1
         DigitOutLine    =   0   'False
         DigitJustify    =   0
      End
      Begin VB.Image Light_Midi_Floppy_Drive 
         Height          =   240
         Left            =   1800
         Picture         =   "Form_Main.frx":18F15B
         Top             =   555
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image6 
         Height          =   165
         Left            =   1920
         Picture         =   "Form_Main.frx":18F7DD
         Top             =   590
         Width           =   255
      End
      Begin VB.Image FloppyIn 
         Height          =   735
         Left            =   1480
         Picture         =   "Form_Main.frx":18FA5B
         Top             =   120
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.Image FloppyOut 
         Height          =   735
         Left            =   1480
         Picture         =   "Form_Main.frx":1927F5
         Top             =   120
         Width           =   3225
      End
      Begin VB.Label Label_MidiFilename 
         BackStyle       =   0  'Transparent
         Caption         =   "T(1019)"
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
         Height          =   200
         Left            =   2210
         TabIndex        =   51
         Top             =   1013
         Width           =   2055
      End
      Begin VB.Image Light_Midi_Play_On 
         Height          =   135
         Left            =   6090
         Picture         =   "Form_Main.frx":19558F
         Top             =   250
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Light_Midi_Pause_On 
         Height          =   150
         Left            =   6085
         Picture         =   "Form_Main.frx":195A69
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
      Picture         =   "Form_Main.frx":195F53
      ScaleHeight     =   735
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   5
         Left            =   4320
         TabIndex        =   101
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
         Caption         =   "Clock"
         ShowLed         =   -1  'True
         TextAlignment   =   0
      End
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   4
         Left            =   5160
         TabIndex        =   102
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
         Caption         =   "Mixer"
         ShowLed         =   -1  'True
         TextAlignment   =   0
      End
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   3
         Left            =   6000
         TabIndex        =   103
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
         Caption         =   "CD"
         ShowLed         =   -1  'True
         TextAlignment   =   0
      End
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   1
         Left            =   7680
         TabIndex        =   104
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
         Caption         =   "MIDI"
         ShowLed         =   -1  'True
         TextAlignment   =   0
      End
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   2
         Left            =   6840
         TabIndex        =   105
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
         Caption         =   "DAT"
         ShowLed         =   -1  'True
         TextAlignment   =   0
      End
      Begin Audiostation.ButtonBig Button_Power 
         Height          =   390
         Index           =   0
         Left            =   8640
         TabIndex        =   106
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
         Picture         =   "Form_Main.frx":1ACE51
         Top             =   100
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   450
         Left            =   720
         Picture         =   "Form_Main.frx":1AD697
         Top             =   105
         Width           =   3165
      End
   End
   Begin ComctlLib.ImageList Imagelist_ElementsDisabled 
      Left            =   11400
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   638
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form_Main.frx":1B2161
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form_Main.frx":1C88F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form_Main.frx":1F75B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form_Main.frx":226277
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form_Main.frx":254F39
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form_Main.frx":2B465B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form_Main.frx":313D7D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Audiostation.ShellPipe ShellPipe 
      Left            =   12000
      Top             =   3120
      _ExtentX        =   635
      _ExtentY        =   635
      PollInterval    =   300
   End
   Begin VB.Menu menu_Popup01 
      Caption         =   "- POPUP (1) -"
      Begin VB.Menu menu_Popup_AutoStop 
         Caption         =   "&Auto Stop"
         Checked         =   -1  'True
         HelpContextID   =   1003
      End
      Begin VB.Menu menu_Popup1_Space01 
         Caption         =   "-"
      End
      Begin VB.Menu menu_Popup_PlayOneTrack 
         Caption         =   "&Play Single"
         Checked         =   -1  'True
         HelpContextID   =   1005
      End
      Begin VB.Menu menu_Popup_RepeatTrack 
         Caption         =   "&Repeat Track"
         HelpContextID   =   1006
      End
      Begin VB.Menu menu_Popup1_Space02 
         Caption         =   "-"
      End
      Begin VB.Menu menu_Popup_NoRepeat 
         Caption         =   "&No Repeat"
         Checked         =   -1  'True
      End
      Begin VB.Menu menu_Popup_RepeatPlaylist 
         Caption         =   "&Repeat Playlist"
         HelpContextID   =   1007
      End
      Begin VB.Menu menu_Popup_Shuffle 
         Caption         =   "&Shuffle Playlist"
         HelpContextID   =   1008
      End
      Begin VB.Menu menu_Popup1_Space03 
         Caption         =   "-"
      End
      Begin VB.Menu menu_Popup_Properties 
         Caption         =   "&Properties"
         HelpContextID   =   1009
      End
   End
   Begin VB.Menu menu_Popup02 
      Caption         =   "- POPUP (2) -"
      Begin VB.Menu menu_Popup_Settings 
         Caption         =   "&Audiostation Settings"
         HelpContextID   =   1010
      End
      Begin VB.Menu menu_Popup_Space03 
         Caption         =   "-"
      End
      Begin VB.Menu menu_Popup_ShowSpectrum 
         Caption         =   "&Show spectrum analyzer"
         Checked         =   -1  'True
         HelpContextID   =   1011
      End
      Begin VB.Menu menu_Popup_Space04 
         Caption         =   "-"
      End
      Begin VB.Menu menu_Popup_CheckForUpdates 
         Caption         =   "&Check for updates"
         HelpContextID   =   1012
      End
      Begin VB.Menu menu_Popup_About 
         Caption         =   "&About Audiostation"
         HelpContextID   =   1013
      End
      Begin VB.Menu menu_Popup_Close 
         Caption         =   "&Close Audiostation"
         HelpContextID   =   1014
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu menu_Popup03 
      Caption         =   "- POPUP (3) -"
      Begin VB.Menu menu_Popup_PlayRecording 
         Caption         =   "&Play"
         HelpContextID   =   1015
      End
      Begin VB.Menu menu_Popup_Record 
         Caption         =   "&Record"
         HelpContextID   =   1016
      End
      Begin VB.Menu menu_Popup_SaveRecording 
         Caption         =   "&Save Recording"
         HelpContextID   =   1017
      End
      Begin VB.Menu menu_Popup_RecordingSettings 
         Caption         =   "&Settings"
         HelpContextID   =   1018
      End
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RecorderFound As Boolean
Public Recording As Boolean

Public ShowRemaining As Boolean
Public ShowRemainingForMidi As Boolean

Dim VolumeChannelId As String
Dim MidiMediaType As enumMidiMediaType
Dim InitDone As Boolean
Public Sub SettingsChanged()

End Sub
Private Sub GetElementsState()
Dim I, ButtonIndex As Integer

For I = 1 To 6
    Element(I).Tag = Extensions.INIRead("main", "Element-" & I, ConfigFile, "OFF")
    
    If (I - 1) = 0 Then
        ButtonIndex = 1
    Else
        If I = 6 Then
            ButtonIndex = 5
        Else
            ButtonIndex = I
        End If
    End If
    
    If StrExt.Contains(Element(I).Tag, "OFF") Then
        Button_Power(ButtonIndex).Active = False
    Else
        Button_Power(ButtonIndex).Active = True
    End If
Next
End Sub
Private Sub AssignToMemorySlotAndSave(MemorySlot As Integer, url As String, Optional Name As String = "")
Form_Streams.Show vbModal, Me

If url <> vbNullString Then
    Call Extensions.INIWrite("main", "TunerMemory-" & MemorySlot, url & "~" & Name, ConfigFile)
    
    Button_TunerMemory(MemorySlot).Tag = url & "~" & Name
End If
End Sub

Private Sub AdioAudioPeak_ChannelAudioLevelChange(leftValue As Integer, rightValue As Integer)
LedBar_DatLeft.Position = leftValue
LedBar_DatRight.Position = rightValue
End Sub

Private Sub AdioAudioPeak_MasterAudioPeakLevelChange(Value As Integer)
LedBar_Master.Position = Value
End Sub

Private Sub AdioAudioPeak_SpectrumLevelChange(col As Integer, Value As Integer)
If menu_Popup_ShowSpectrum.Checked Then: LedBar_Spectrum(col).Position = Value
End Sub

Private Sub AdioCDPlayer_DeviceFound(Id As Long, DriveName As String, DriveLetter As String)
Call AppLog.LogInfo("CD-Rom device found: " & DriveName & " - " & DriveLetter & ":")
End Sub

Private Sub AdioCDPlayer_NoCdRomDeviceFound()
Call AppLog.LogInfo("No CD-Rom device found on this computer")
Button_Power(3).Enabled = False
End Sub

Private Sub AdioCDPlayer_StartPlay()
AdioMediaPlayer.StopPlay
AdioMidiPlayer.StopPlay

Call AdioAudioPeak.SetChannel(AdioCDPlayer.Channel)
End Sub

Private Sub AdioCore_DeviceFound(Id As Integer, Name As String, InputDev As Boolean, OutputDev As Boolean)
Call AppLog.LogInfo("Audio device found: " & Name & " - " & Id & " - Input: " & InputDev & " - Output: " & OutputDev)
End Sub

Private Sub AdioMediaPlayer_MediaEnded()
If Not menu_Popup_AutoStop.Checked Then
    If menu_Popup_PlayOneTrack.Checked Then
        Call AdioMediaPlaylist.GetTrack(PLS_NEXT)
        Exit Sub
    Else
        Call AdioMediaPlayer.StartPlay
        Exit Sub
    End If
End If
End Sub

Private Sub AdioMediaPlayer_NewMediaFile(File As String)
Dim fso As New FileSystemObject

Label_FilenameDat.Caption = fso.GetFileName(File)

Call AppLog.LogInfo("Starting playing " & File)

menu_Popup_Properties.Enabled = True
End Sub

Private Sub AdioMediaPlayer_NewStream()
label_StreamStatus.Caption = "Loading..."
End Sub

Private Sub AdioMediaPlayer_StartPlay()
AdioCDPlayer.StopPlay
AdioMidiPlayer.StopPlay

Call AdioAudioPeak.SetChannel(AdioMediaPlayer.Channel)
End Sub

Private Sub AdioMediaPlayer_StreamBuffering(Percent As Integer)
Debug.Print Percent
End Sub

Private Sub AdioMediaPlayer_StreamTitleChange(Title As String)
Label_StreamTitle.Caption = Title
End Sub

Private Sub AdioMediaPlaylist_TrackChanged(Track As AdioLibrary.mdlAdioPlaylistItem)
AdioMediaPlayer.LoadFile Track.LocalFile
AdioMediaPlayer.StartPlay

CurrentMediaPlayerTrackNr = Track.nR
End Sub

Private Sub AdioMidiPlayer_MidiTrack(Name As String, TrackNr As Integer)
Debug.Print "Midi track: " & Name
End Sub

Private Sub AdioMidiPlayer_MidiTrackAudioLevelChange(TrackNr As Integer, Level As Integer)
If TrackNr > 15 Then: Exit Sub

VU_Midi(TrackNr).Position = Level
End Sub

Private Sub AdioMidiPlayer_NewMediaFile(File As String)
Dim fso As New FileSystemObject

Label_MidiFilename.Caption = fso.GetFileName(File)
End Sub

Private Sub AdioMidiPlayer_Ready()
Debug.Print "Midi player ready"
End Sub

Private Sub AdioMidiPlayer_StartPlay()
AdioMediaPlayer.StopPlay
AdioCDPlayer.StopPlay
End Sub

Private Sub AdioMidiPlaylist_TrackChanged(Track As AdioLibrary.mdlAdioPlaylistItem)
AdioMidiPlayer.StopPlay
AdioMidiPlayer.LoadFile Track.LocalFile
AdioMidiPlayer.StartPlay

CurrentMidiPlayerTrackNr = Track.nR
End Sub

Private Sub Button_CDLoop_Click()
If Button_CDLoop.Active = False Then
    Button_CDLoop.Active = True
    Button_CDRandom.Active = False
    
    Light_Panel_CD.Picture = Imagelist_CDDisplay.ListImages(3).Picture
Else
    Button_CDLoop.Active = False
    
    Light_Panel_CD.Picture = Imagelist_CDDisplay.ListImages(2).Picture
End If
End Sub

Private Sub Button_CDPlayer_Click(Index As Integer)
Select Case Index
    'Case 0: Call AdioCDPlayer.PreviousTrack
    Case 1: Call AdioCDPlayer.SeekBySeconds(AdioRewind)
    Case 2: Call AdioCDPlayer.StopPlay
    Case 3: Call AdioCDPlayer.StartPlay
    Case 4: Call AdioCDPlayer.PausePlay
    Case 5: Call AdioCDPlayer.SeekBySeconds(AdioForward)
    'Case 6: Call MediaPlayer_CD.NextTrack
    Case 7: Call AdioCDPlayer.OpenDoor
End Select
End Sub

Private Sub Button_CDRandom_Click()
If Button_CDRandom.Active = False Then
    Button_CDRandom.Active = True
    Button_CDLoop.Active = False
    
    Light_Panel_CD.Picture = Imagelist_CDDisplay.ListImages(4).Picture
Else
    Button_CDRandom.Active = False
    
    Light_Panel_CD.Picture = Imagelist_CDDisplay.ListImages(2).Picture
End If
End Sub

Private Sub Button_MidiPlayer_Click(Index As Integer)
Select Case Index
    Case 0: Call AdioMidiPlaylist.GetTrack(PLS_PREV)
    Case 1: Call AdioMidiPlayer.SeekBySeconds(AdioRewind)
    Case 2: Call AdioMidiPlayer.StopPlay
    Case 3
        If AdioMidiPlaylist.ListCount = 0 Then
            If MsgBox(GetTranslation(1056), vbQuestion + vbYesNo) = vbYes Then
                Form_Playlist.FormType = enumFormTypes.MidiPlayer
                Form_Playlist.Show vbModal
            End If
        Else
            Call AdioMidiPlayer.StartPlay
        End If

    Case 4: Call AdioMidiPlayer.PausePlay
    Case 5: Call AdioMidiPlayer.SeekBySeconds(AdioForward)
    Case 6: Call AdioMidiPlaylist.GetTrack(PLS_NEXT)
End Select
End Sub
Private Sub Button_OpenStream_Click()
Form_Streams.Show vbModal, Me

If Not StrExt.IsNullOrWhiteSpace(StreamUrl) Then
    AdioMediaPlayer.LoadStream StreamUrl
        
    Dim I As Integer
    For I = 0 To Button_TunerMemory.Count - 1
        Button_TunerMemory(I).Active = False
    Next
End If
End Sub

Private Sub Button_PlayStream_Click()
If StrExt.IsNullOrWhiteSpace(StreamUrl) Then
    If MsgBox(GetTranslation(1055), vbQuestion + vbYesNo) = vbYes Then
        Call Button_OpenStream_Click
    End If
Else
    Call AdioMediaPlayer.LoadStream(StreamUrl)
End If
End Sub

Private Sub Button_Power_Click(Index As Integer)
Dim ElmIndex As Integer

Select Case Index
    Case 1, 2, 3, 4, 5
        If Index = 5 Then
            ElmIndex = 6
        Else
            ElmIndex = Index
        End If
        
        If Element(ElmIndex).Tag = "ON" Then
            Button_Power(Index).Active = False
            
            If Index = 4 Then
                Element(4).Tag = "OFF"
                Element(5).Tag = "OFF"
            Else
                Element(ElmIndex).Tag = "OFF"
            End If
            
            InitDone = False
        Else
            Button_Power(Index).Active = True
            
            If Index = 4 Then
                Element(4).Tag = "ON"
                Element(5).Tag = "ON"
            Else
                Element(ElmIndex).Tag = "ON"
            End If
            
            InitDone = False
        End If
    
    Case 0: Unload Me
End Select
End Sub

Private Sub Button_StopStream_Click()
AdioMediaPlayer.StopPlay
End Sub

Private Sub Button_TunerMemory_Click(Index As Integer)
Dim I As Integer

For I = 0 To Button_TunerMemory.Count - 1
    Button_TunerMemory(I).Active = False
Next

If Button_TunerMemory(Index).Tag = vbNullString Then
    If MsgBox(GetTranslation(1060), vbQuestion + vbYesNo) = vbYes Then
        Button_OpenStream_Click
    End If
Else
    Button_TunerMemory(Index).Active = True
    
    StreamUrl = StrExt.SplitStr(Button_TunerMemory(Index).Tag, "~", 0)
    StreamName = StrExt.SplitStr(Button_TunerMemory(Index).Tag, "~", 1)
    
    Call AdioMediaPlayer.LoadStream(StreamUrl)
End If
End Sub

Private Sub cmdAudioPlayer_Click(Index As Integer)
Select Case Index
    Case 0: PopupMenu menu_Popup01
    Case 1: PopupMenu menu_Popup03
    
    Case 2: Call AdioMediaPlaylist.GetTrack(PLS_PREV)
    Case 3: Call AdioMediaPlayer.SeekBySeconds(AdioRewind)
    Case 4: Call AdioMediaPlayer.StopPlay
    Case 5
        If AdioMediaPlaylist.ListCount = 0 Then
            If MsgBox(GetTranslation(1056), vbQuestion + vbYesNo) = vbYes Then
                Form_Playlist.FormType = enumFormTypes.Mp3Player
                Form_Playlist.Show vbModal
            End If
        Else
            ' If no track is selected, select the first in the playlist
            If CurrentMediaPlayerTrackNr = 0 Then: Call AdioMediaPlaylist.GetTrack(PLS_FIRST)
            Call AdioMediaPlayer.StartPlay
        End If
    
    Case 6: Call AdioMediaPlayer.PausePlay
    Case 7: Call AdioMediaPlayer.SeekBySeconds(AdioForward)
    Case 8: Call AdioMediaPlaylist.GetTrack(PLS_NEXT)
End Select
End Sub

Private Sub cmdPlaylistDat_Click()
Form_Playlist.FormType = Mp3Player
Form_Playlist.Show , Me
End Sub

Private Sub CmdPlaylistMidi_Click()
Form_Playlist.FormType = MidiPlayer
Form_Playlist.Show , Me
End Sub

Private Sub cmdSettingsDat_Click()
Form_Settings.Show vbModal, Me
End Sub

Private Sub cmdSettingsMidi_Click()
Form_Settings.Show vbModal, Me
End Sub

Private Sub DataInter_DataReceived(Data As String)
Call AppLog.LogInfo("DataInterFileCopy" & Data)

Select Case StrExt.SplitStr(Data, "~", 0)
    Case "OpenFile" ' Open a file from a other Audiostation instance
        Call OpenFile(StrExt.SplitStr(Data, "~", 1))
        
    Case "StartRecording"
        menu_Popup_Record.Caption = "Stop Recording"
        menu_Popup_Record.Tag = "stop"
        
        menu_Popup_SaveRecording.Enabled = False
        menu_Popup_PlayRecording.Enabled = False
        
        Recording = True
        
    Case "StopRecording"
        menu_Popup_Record.Caption = "Record"
        menu_Popup_Record.Tag = "start"
        
        menu_Popup_SaveRecording.Enabled = True
        menu_Popup_PlayRecording.Enabled = True
        
        Recording = False
           
        Debug.Print "Stop recording, Length: " & StrExt.SplitStr(Data, "~", 1)
        
    Case "Done"
        DataInter.Finish
        RecorderFound = False
        
End Select
End Sub

Private Sub DataInter_ReceptorConnected()
RecorderFound = True
Call AppLog.LogInfo("Application connected to data interchange")
End Sub

Private Sub DigitBox_DatDuration_Click()
If ShowRemaining Then
    ShowRemaining = False
Else
    ShowRemaining = True
End If
End Sub

Private Sub DigitBox_MidiDuration_Click()
If ShowRemainingForMidi Then
    ShowRemainingForMidi = False
Else
    ShowRemainingForMidi = True
End If
End Sub

Private Sub Form_Load()
Dim PlaybackDeviceIndex, PlaybackDeviceId As Long
Dim MidiDeviceId As Integer

Call AppLog.LogInfo("Initialize AdioCore component")
AdioCore.Initialize

Call AppLog.LogInfo("Initialize AdioAudioPeak component")
AdioAudioPeak.Run

Call AppLog.LogInfo("Initialize AdioCDPlayer component")
AdioCDPlayer.Initialize

Call AppLog.LogInfo("Initialize AdioMidiPlayer component")
MidiDeviceId = CInt(Extensions.INIRead("main", "MidiPlaybackDeviceId", ConfigFile, 0)) - 1

AdioMidiPlayer.GetListOfMidiDevices
AdioMidiPlayer.InitComponent MidiDeviceId

Randomize
Call DataInter.Connect("15448")

Width = 9900
Height = 9855 - 310

Dim I As Integer
For I = 0 To Button_TunerMemory.Count - 1
    Button_TunerMemory(I).Tag = Extensions.INIRead("main", "TunerMemory-" & I, ConfigFile, vbNullString)
Next

' Disable menu items
menu_Popup01.Visible = False
menu_Popup02.Visible = False
menu_Popup03.Visible = False

menu_Popup_Record.Tag = "start"

menu_Popup_Properties.Enabled = False
menu_Popup_SaveRecording.Enabled = False
menu_Popup_PlayRecording.Enabled = False

' Get the application settings
menu_Popup_ShowSpectrum.Checked = Extensions.INIRead("main", "UseSpectrumAnalyzer", ConfigFile, 1)
menu_Popup_AutoStop.Checked = Extensions.INIRead("main", "AutoStop", ConfigFile, 1)

ShowRemaining = Extensions.INIRead("main", "ShowRemainingTime", ConfigFile, 0)
ShowRemainingForMidi = Extensions.INIRead("main", "ShowRemainingTimeForMidi", ConfigFile, 0)

'Display program version
lbl_version.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision

Light_Panel_CD.Picture = Imagelist_CDDisplay.ListImages(2).Picture
Picture_MediaPlayerAnimation.Picture = Imagelist_MediaPlayerAnimation.ListImages(1).Picture

DigitBox_Clock.DigitDisplay = Time

Call TranslateFormAndControls(Me)

Call GetElementsState
Call SettingsChanged

' Get the midi volume controller
'If Not modVolume.ChannelExists("VirtualMIDISynth.exe") Is Nothing Then
'    VolumeChannelId = modVolume.ChannelExists("VirtualMIDISynth.exe").Pid
'Else
'    If Not modVolume.ChannelExists("Audiostation.exe") Is Nothing Then
'        VolumeChannelId = modVolume.ChannelExists("Audiostation.exe").Pid
'    End If
'End If

Timer_Main.Enabled = True
Timer_Lights.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim VirtualMidiSyncPid As Long

Cancel = True

Call AppLog.LogInfo("Ending")

' Stop playing, recording, etc.
Call AdioMediaPlayer.StopPlay
Call AdioCDPlayer.StopPlay

' Send end to all connected applications
DataInter.SendData "end"

' Save element state
Dim I As Integer
For I = 1 To 6
    Call Extensions.INIWrite("main", "Element-" & I, Element(I).Tag, ConfigFile)
Next

Call AppLog.CloseLog

Cancel = False
End
End Sub

Private Sub Label_FilenameDat_Change()
Label_FilenameDat.ToolTipText = Label_FilenameDat.Caption
End Sub

Private Sub Label_MidiFilename_Change()
Label_MidiFilename.ToolTipText = Label_MidiFilename.Caption
End Sub
Private Sub menu_Popup_About_Click()
Form_About.Show vbModal, Me
End Sub

Private Sub menu_Popup_AutoStop_Click()
If menu_Popup_AutoStop.Checked Then
    menu_Popup_AutoStop.Checked = False
Else
    menu_Popup_AutoStop.Checked = True
End If

Call Extensions.INIWrite("main", "AutoStop", IIf(menu_Popup_AutoStop.Checked, 1, 0), ConfigFile)
Call SettingsChanged
End Sub

Private Sub menu_Popup_NoRepeat_Click()
menu_Popup_Shuffle.Checked = False
menu_Popup_RepeatPlaylist.Checked = False
menu_Popup_NoRepeat.Checked = True

AdioMediaPlaylist.RepeatMode = PLS_NO_REPEAT
End Sub

Private Sub menu_Popup_PlayOneTrack_Click()
menu_Popup_PlayOneTrack.Checked = True
menu_Popup_RepeatTrack.Checked = False

AdioMediaPlayer.RepeatMode = AdioPlayTrack
End Sub

Private Sub menu_Popup_PlayRecording_Click()
DataInter.SendData "play"
End Sub

Private Sub menu_Popup_Properties_Click()
Form_Track_Properties.File = AdioMediaPlayer.LoadedFile
Form_Track_Properties.Show vbModal, Me
End Sub

Private Sub menu_Popup_Record_Click()
If RecorderFound = False Then
    Shell App.path & "\recorder.exe", vbHide
    
    Do While RecorderFound = False
        DoEvents
    Loop
End If

If menu_Popup_Record.Tag = "start" Then
    
    Call DataInter.SendData("record")
    
ElseIf menu_Popup_Record.Tag = "stop" Then

    Call DataInter.SendData("stop")
    
End If
End Sub

Private Sub menu_Popup_RecordingSettings_Click()
Form_Settings_Recorder.Show vbModal, Me
End Sub

Private Sub menu_Popup_RepeatPlaylist_Click()
menu_Popup_Shuffle.Checked = False
menu_Popup_RepeatPlaylist.Checked = True
menu_Popup_NoRepeat.Checked = False

AdioMediaPlaylist.RepeatMode = PLS_REPEAT
End Sub

Private Sub menu_Popup_RepeatTrack_Click()
menu_Popup_PlayOneTrack.Checked = False
menu_Popup_RepeatTrack.Checked = True

AdioMediaPlayer.RepeatMode = AdioRepeatTrack
End Sub

Private Sub menu_Popup_SaveRecording_Click()
DataInter.SendData "save"
DataInter.SendData "end"
End Sub

Private Sub menu_Popup_Settings_Click()
Form_Settings.Show vbModal, Me
End Sub

Private Sub menu_Popup_ShowSpectrum_Click()
Dim Result As Integer

If menu_Popup_ShowSpectrum.Checked Then
    menu_Popup_ShowSpectrum.Checked = False
    Result = 0
    
    ' Call AdioAudioPeak.reset
Else
    menu_Popup_ShowSpectrum.Checked = True
    Result = 1
End If

Call Extensions.INIWrite("main", "UseSpectrumAnalyzer", CStr(Result), ConfigFile)
End Sub

Private Sub menu_Popup_Shuffle_Click()
menu_Popup_Shuffle.Checked = True
menu_Popup_RepeatPlaylist.Checked = False
menu_Popup_NoRepeat.Checked = False

AdioMediaPlaylist.RepeatMode = PLS_SHUFFLE
End Sub

Private Sub OptionsMenuButton_Click()
PopupMenu menu_Popup02
End Sub

Private Sub ShellPipe_DataArrival(ByVal CharsTotal As Long)
DigitBox_MidiDuration.DigitDisplay = Right(Trim(ShellPipe.GetData), 5)
End Sub

Private Sub Slider_CD_Balance_OnPositionChange()
Call AdioCDPlayer.SetBalance(Slider_CD_Balance.Value)
End Sub

Private Sub Slider_Dat_Balance_OnPositionChange()
Call AdioMediaPlayer.SetBalance(Slider_Dat_Balance.Value)
End Sub

Private Sub Slider_Dat_Volume_OnPositionChange()
Call AdioMediaPlayer.SetVolume(Slider_Dat_Volume.Value)
End Sub

Private Sub Slider_Master_OnPositionChange()
Call AdioAudioPeak.SetMasterVolume(Slider_Master.Value)
End Sub

Private Sub Slider_Midi_Volume_OnPositionChangeFinished()
Call modVolume.SetVolumeById(VolumeChannelId, Slider_Midi_Volume.Value)
End Sub

Private Sub Switch_CD_OnChange()
AdioCDPlayer.Mute
End Sub

Private Sub Switch_Dat_OnChange()
AdioMediaPlayer.MuteAudio
End Sub

Private Sub Switch_Master_OnChange()
If Switch_Master.Active Then
    Call AdioAudioPeak.MuteMasterVolume
Else
    Call AdioAudioPeak.MuteMasterVolume
End If
End Sub

Private Sub Switch_Midi_OnChange()
If Switch_Midi.Active Then
    Call modVolume.SetUnmuteById(VolumeChannelId)
Else
    Call modVolume.SetMuteById(VolumeChannelId)
End If
End Sub

Private Sub Timer_Lights_Timer()
' Recorder
If Recording Then
    Picturebox_Recording.Visible = True

    If Image_Recording.Visible = True Then
        Image_Recording.Visible = False
    Else
        Image_Recording.Visible = True
    End If
Else
    Picturebox_Recording.Visible = False
End If

' Midi
If AdioMidiPlayer.State = AdioPlaying Then
    FloppyIn.Visible = True

    Light_Midi_Pause_On.Visible = False

    If Light_Midi_Play_On.Visible = True Then
        Light_Midi_Play_On.Visible = False
        Light_Midi_Floppy_Drive.Visible = True
    Else
        Light_Midi_Play_On.Visible = True
        Light_Midi_Floppy_Drive.Visible = False
    End If

ElseIf AdioMidiPlayer.State = AdioStopped Or AdioEnded Then
    FloppyIn.Visible = False

    Light_Midi_Floppy_Drive.Visible = False
    Light_Midi_Play_On.Visible = False
    Light_Midi_Pause_On.Visible = False
    Light_Midi_Play_On.Visible = False

ElseIf AdioMidiPlayer.State = AdioPaused Then
    Light_Midi_Floppy_Drive.Visible = False
    FloppyIn.Visible = False

    If Light_Midi_Pause_On.Visible = True Then
        Light_Midi_Pause_On.Visible = False
    Else
        Light_Midi_Pause_On.Visible = True
    End If
End If

' Media
If AdioMediaPlayer.State = AdioPlaying Then
    Timer_MediaAnimation.Enabled = True
    Light_Dat_Pause_On.Visible = False

    If Light_Dat_Play_On.Visible = True Then
        Light_Dat_Play_On.Visible = False
    Else
        Light_Dat_Play_On.Visible = True
    End If
ElseIf AdioMediaPlayer.State = AdioPaused Then
    Timer_MediaAnimation.Enabled = False
    Light_Dat_Play_On.Visible = True

    LedBar_DatLeft.Position = 0
    LedBar_DatRight.Position = 0

    If Light_Dat_Pause_On.Visible = True Then
        Light_Dat_Pause_On.Visible = False
    Else
        Light_Dat_Pause_On.Visible = True
    End If
ElseIf AdioMediaPlayer.State = AdioStopped Or AdioEnded Then
    Light_Dat_Pause_On.Visible = False
    Light_Dat_Play_On.Visible = False

    LedBar_DatLeft.Position = 0
    LedBar_DatRight.Position = 0

    Timer_MediaAnimation.Enabled = False
End If
End Sub

Private Sub Timer_Main_Timer()
Dim E As Integer

DigitBox_Clock.DigitDisplay = Time

' Duration of the media player
DigitBox_DatTrack.DigitDisplay = format(CurrentMediaPlayerTrackNr, "00")
DigitBox_MidiTrack.DigitDisplay = format(CurrentMidiPlayerTrackNr, "00")

If AdioMediaPlayer.State = AdioPaused Or AdioMediaPlayer.State = AdioStopped Then
    Call AdioAudioPeak.ResetSpectrum
End If

If ShowRemaining Then
    DigitBox_DatDuration.DigitDisplay = AdioMediaPlayer.GetProperties.RemainingString
Else
    DigitBox_DatDuration.DigitDisplay = AdioMediaPlayer.GetProperties.ElapsedString
End If

If ShowRemainingForMidi Then
    DigitBox_MidiDuration.DigitDisplay = AdioMidiPlayer.GetProperties.RemainingString
Else
    DigitBox_MidiDuration.DigitDisplay = AdioMidiPlayer.GetProperties.ElapsedString
End If

' Init of the Audiostation application
If Not InitDone Then
    On Error GoTo ErrorHandler
    
    For E = 1 To PictureBox_Disabled.Count
        Unload PictureBox_Disabled(E)
    Next

AddDisabled:
    For E = 1 To Element.Count - 1
        Dim NewElementIndex, DisabledIndex As Integer
        
        NewElementIndex = PictureBox_Disabled.Count
        
        Load PictureBox_Disabled(NewElementIndex)
        
        PictureBox_Disabled(NewElementIndex).Top = Element(E).Top
        PictureBox_Disabled(NewElementIndex).Left = Element(E).Left
        PictureBox_Disabled(NewElementIndex).Height = Element(E).Height
        PictureBox_Disabled(NewElementIndex).Width = Element(E).Width
        PictureBox_Disabled(NewElementIndex).BackColor = vbBlack
        
        DisabledIndex = E + 1
        If DisabledIndex = 6 Then
            PictureBox_Disabled(4).Picture = Imagelist_ElementsDisabled.ListImages(6).Picture
            PictureBox_Disabled(4).Height = 3060
            PictureBox_Disabled(4).ZOrder 0
        Else
            PictureBox_Disabled(NewElementIndex).Picture = Imagelist_ElementsDisabled.ListImages(DisabledIndex).Picture
            PictureBox_Disabled(NewElementIndex).ZOrder 0
        End If
        
        If LCase(Element(E).Tag) = "off" And Not InStr(1, Element(E).Tag, "force") > 0 Then
            PictureBox_Disabled(NewElementIndex).Visible = True
        Else
            PictureBox_Disabled(NewElementIndex).Visible = False
        End If
    Next
End If

InitDone = True

Exit Sub

ErrorHandler:
Select Case Err.Number
    Case 0
    Case 340: GoTo AddDisabled
End Select
End Sub

Private Sub Timer_MediaAnimation_Timer()
Dim J As Integer

J = Timer_MediaAnimation.Tag
Picture_MediaPlayerAnimation.Picture = Imagelist_MediaPlayerAnimation.ListImages.Item(J).Picture

If Timer_MediaAnimation.Tag = Imagelist_MediaPlayerAnimation.ListImages.Count Then
    Timer_MediaAnimation.Tag = 1
Else
    Timer_MediaAnimation.Tag = Timer_MediaAnimation.Tag + 1
End If
End Sub

Private Sub Timer_MidiVu_Timer()
Dim I As Integer

For I = 0 To VU_Midi.Count - 1
    VU_Midi(I).Position = VU_Midi(I).Position - 5
Next
End Sub

Private Sub Trm_CD_Animation_Timer()
'Dim ImgIndex As Integer
'
'ImgIndex = Trm_CD_Animation.Tag
'
'AniCD.Picture = ImageList5.ListImages(ImgIndex).Picture
'AniCD.Visible = True
'
'If DoorClose = True Then
'    If Trm_CD_Animation.Tag = 1 Then
'        DoorClose = False
'        Trm_CD_Animation.Enabled = False
'        AniCD.Visible = False
'    Else
'        Trm_CD_Animation.Tag = Trm_CD_Animation.Tag - 1
'    End If
'Else
'    If Trm_CD_Animation.Tag = 7 Then
'        DoorClose = True
'        Trm_CD_Animation.Enabled = False
'    Else
'        Trm_CD_Animation.Tag = Trm_CD_Animation.Tag + 1
'    End If
'End If
End Sub
