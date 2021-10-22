VERSION 5.00
Object = "{C5412DA5-2E2F-11D3-85BF-00105AC8B715}#1.0#0"; "isAnalogLibrary.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ButtonBig 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "ButtonBig.ctx":0000
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   619
   Begin VB.PictureBox ButtonContent 
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
      Height          =   210
      Left            =   360
      ScaleHeight     =   210
      ScaleWidth      =   1935
      TabIndex        =   2
      Top             =   120
      Width           =   1935
      Begin isAnalogLibrary.iLabelX iLabelX1 
         Height          =   210
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1200
         AutoSize        =   0   'False
         Alignment       =   1
         BorderStyle     =   0
         Caption         =   "CommandButton"
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
         FontColor       =   0
         Transparent     =   0   'False
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   -1  'True
         Enabled         =   -1  'True
         Object.Width           =   80
         Object.Height          =   14
         WordWrap        =   0   'False
         OPCItemCount    =   0
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
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
      Height          =   3390
      Left            =   9045
      ScaleHeight     =   3390
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   0
      Width           =   240
      Begin VB.Image Image2 
         Height          =   105
         Left            =   60
         Picture         =   "ButtonBig.ctx":01E2
         Top             =   120
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Image Image1 
         Height          =   105
         Left            =   60
         Picture         =   "ButtonBig.ctx":02CC
         Top             =   120
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Image RightPixels 
         Height          =   375
         Left            =   120
         Picture         =   "ButtonBig.ctx":03D2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   120
      End
      Begin VB.Image LastPixel 
         Height          =   375
         Left            =   0
         Picture         =   "ButtonBig.ctx":0684
         Stretch         =   -1  'True
         Top             =   0
         Width           =   120
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   26
      ImageHeight     =   8
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ButtonBig.ctx":0936
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ButtonBig.ctx":0CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ButtonBig.ctx":105A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ButtonBig.ctx":13EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ButtonBig.ctx":177E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ButtonBig.ctx":1B10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
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
      Height          =   3390
      Left            =   0
      ScaleHeight     =   3390
      ScaleWidth      =   120
      TabIndex        =   1
      Top             =   0
      Width           =   120
      Begin VB.Image LeftPixels 
         Height          =   375
         Left            =   0
         Picture         =   "ButtonBig.ctx":1EA2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Image CenterPixels 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "ButtonBig.ctx":2154
      Stretch         =   -1  'True
      Top             =   0
      Width           =   120
   End
End
Attribute VB_Name = "ButtonBig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Const m_def_ShowLed = False
Const m_def_Active = False

Public Enum enumTextAlign
    TextCenter
    TextLeft
    TextRight
End Enum

Dim m_ShowLed As Boolean
Private Sub ControlChanges()
If UserControl.Enabled = False Then
    iLabelX1.ShadowXOffset = 1
    iLabelX1.ShadowYOffset = 1
    iLabelX1.FontColor = &H808080
    iLabelX1.top = 10
Else
    iLabelX1.ShadowXOffset = -1
    iLabelX1.ShadowYOffset = -1
    iLabelX1.FontColor = vbBlack
    iLabelX1.top = 0
End If
End Sub
Private Sub CenterPixels_Click(index As Integer)
RaiseEvent Click
End Sub

Private Sub iLabelX1_OnClick()
RaiseEvent Click
End Sub

Private Sub Image3_Click()
RaiseEvent Click
End Sub

Private Sub Image5_Click()
RaiseEvent Click
End Sub

Private Sub iLabelX1_OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim J As Integer

For J = 0 To CenterPixels.count - 1
    CenterPixels(J).Picture = ImageList1.ListImages(5).Picture
Next

LastPixel.Picture = ImageList1.ListImages(5).Picture
LeftPixels.Picture = ImageList1.ListImages(4).Picture
RightPixels.Picture = ImageList1.ListImages(6).Picture

ButtonContent.top = 7
ButtonContent.Left = 6

Image1.top = 140
Image1.Left = 75
Image2.top = 140
Image2.Left = 75
End Sub

Private Sub iLabelX1_OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim J As Integer

For J = 0 To CenterPixels.count - 1
    CenterPixels(J).Picture = ImageList1.ListImages(2).Picture
Next

LastPixel.Picture = ImageList1.ListImages(2).Picture
LeftPixels.Picture = ImageList1.ListImages(1).Picture
RightPixels.Picture = ImageList1.ListImages(3).Picture

ButtonContent.top = 6
ButtonContent.Left = 5

Image1.top = 120
Image1.Left = 60
Image2.top = 120
Image2.Left = 60
End Sub
Private Sub UserControl_Initialize()
Dim K As Integer

For K = 1 To CenterPixels.count - 1
    Unload CenterPixels(K)
Next

ButtonContent.top = 6
ButtonContent.Left = 5
End Sub

Private Sub UserControl_Resize()
Dim J As Integer
Dim Pixels As Integer

Pixels = UserControl.ScaleWidth / 8
ButtonContent.Width = UserControl.ScaleWidth - 16

On Error Resume Next
For J = 1 To Pixels
    Load CenterPixels(J)
    With CenterPixels(J)
        .Left = 8 * J
        .top = 0
        .Visible = True
    End With
Next

iLabelX1.Width = ButtonContent.ScaleWidth + 20
UserControl.Height = 390
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HC0C0C0)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    iLabelX1.Caption = PropBag.ReadProperty("Caption", "CommandButton")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Image1.Visible = PropBag.ReadProperty("ShowLed", m_def_ShowLed)
    Image2.Visible = PropBag.ReadProperty("Active", m_def_Active)
    iLabelX1.Alignment = PropBag.ReadProperty("TextAlignment", iLabelX1.Alignment)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HC0C0C0)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Caption", iLabelX1.Caption, "CommandButton")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ShowLed", Image1.Visible, m_def_ShowLed)
    Call PropBag.WriteProperty("Active", Image2.Visible, m_def_Active)
    Call PropBag.WriteProperty("TextAlignment", iLabelX1.Alignment)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=iLabelX1,iLabelX1,-1,Caption
Public Property Get Caption() As String
    Caption = iLabelX1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    iLabelX1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
    
    Call ControlChanges
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
    Call ControlChanges
End Property
Public Property Get ShowLed() As Boolean
    ShowLed = Image1.Visible
End Property

Public Property Let ShowLed(ByVal New_ShowLed As Boolean)
    Image1.Visible = New_ShowLed
    PropertyChanged "ShowLed"
End Property
Public Property Get Active() As Boolean
    Active = Image2.Visible
End Property

Public Property Let Active(ByVal New_Active As Boolean)
    Image2.Visible = New_Active
    PropertyChanged "Active"
End Property
Public Property Get Alignment() As enumTextAlign
    Select Case iLabelX1.Alignment
        Case iahLeft: Alignment = TextLeft
        Case iahRight: Alignment = TextRight
        Case iahCenter: Alignment = TextCenter
    End Select
End Property

Public Property Let Alignment(ByVal New_Alignment As enumTextAlign)
    Select Case New_Alignment
        Case TextLeft: iLabelX1.Alignment = iahLeft
        Case TextRight: iLabelX1.Alignment = iahRight
        Case TextCenter: iLabelX1.Alignment = iahCenter
    End Select
    
    PropertyChanged "Alignment"
End Property

