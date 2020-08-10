VERSION 5.00
Object = "{C5412DA5-2E2F-11D3-85BF-00105AC8B715}#1.0#0"; "isAnalogLibrary.ocx"
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   ScaleHeight     =   1650
   ScaleWidth      =   3945
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   200
      Left            =   25
      ScaleHeight     =   195
      ScaleWidth      =   750
      TabIndex        =   0
      Top             =   90
      Width           =   750
      Begin isAnalogLibrary.iLabelX iLabelX1 
         Height          =   210
         Left            =   30
         TabIndex        =   1
         Top             =   0
         Width           =   510
         AutoSize        =   -1  'True
         Alignment       =   0
         BorderStyle     =   0
         Caption         =   " Power"
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
         BackGroundColor =   16777215
         UpdateFrameRate =   60
         Object.Visible         =   -1  'True
         FontColor       =   -16777208
         Transparent     =   -1  'True
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Enabled         =   -1  'True
         Object.Width           =   34
         Object.Height          =   14
         WordWrap        =   0   'False
         OPCItemCount    =   0
      End
      Begin VB.Image Image4 
         Height          =   105
         Left            =   600
         Picture         =   "Button.ctx":0000
         Top             =   30
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Image Image3 
         Height          =   105
         Left            =   600
         Picture         =   "Button.ctx":0106
         Top             =   30
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   0
      Picture         =   "Button.ctx":020C
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   0
      Picture         =   "Button.ctx":1406
      Top             =   0
      Width           =   825
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
'Default Property Values:
Const m_def_Active = False
Const m_def_ShowLed = False
'Property Variables:
Dim m_Active As Boolean
Dim m_ShowLed As Boolean
Private Sub iLabelX1_OnClick()
UserControl_Click
End Sub

Private Sub iLabelX1_OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Image2.Visible = True
Picture1.top = 100
Picture1.Left = 40
End Sub

Private Sub iLabelX1_OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Image2.Visible = False
Picture1.top = 90
Picture1.Left = 25
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Picture1.top = 100
Picture1.Left = 40
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Picture1.top = 90
Picture1.Left = 25
End Sub
Private Sub Image2_Click()
UserControl_Click
End Sub

Private Sub Picture1_Click()
UserControl_Click
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Picture1.top = 100
Picture1.Left = 40
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Picture1.top = 90
Picture1.Left = 25
End Sub

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_ShowLed = m_def_ShowLed
    m_Active = m_def_Active
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HC0C0C0)
    iLabelX1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    iLabelX1.Caption = PropBag.ReadProperty("Caption", "Power")
    iLabelX1.Alignment = PropBag.ReadProperty("Alignment", 0)
    Image3.Visible = PropBag.ReadProperty("ShowLed", m_def_ShowLed)
    Image4.Visible = PropBag.ReadProperty("Active", m_def_Active)
    
    If Image3.Visible = True Then
        iLabelX1.AutoSize = True
    Else
        iLabelX1.Width = 705
        Image4.Visible = False
        PropertyChanged "ShowLed"
    End If
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HC0C0C0)
    Call PropBag.WriteProperty("Enabled", iLabelX1.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", iLabelX1.Caption, UserControl.Name)
    Call PropBag.WriteProperty("Alignment", iLabelX1.Alignment, 0)
    Call PropBag.WriteProperty("ShowLed", Image3.Visible, m_def_ShowLed)
    Call PropBag.WriteProperty("Active", Image4.Visible, m_def_Active)
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
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=iLabelX1,iLabelX1,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = iLabelX1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    iLabelX1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
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
Public Property Get Alignment() As TxiAlignmentHorizontal
    Alignment = iLabelX1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As TxiAlignmentHorizontal)
    iLabelX1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowLed() As Boolean
    ShowLed = Image3.Visible
End Property

Public Property Let ShowLed(ByVal New_ShowLed As Boolean)
    Image3.Visible = New_ShowLed
    PropertyChanged "ShowLed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get Active() As Boolean
    Active = Image4.Visible
End Property

Public Property Let Active(ByVal New_Active As Boolean)
    Image4.Visible = New_Active
    PropertyChanged "Active"
End Property

