VERSION 5.00
Object = "{852E65AD-72F8-11CF-840E-444553540000}#1.1#0"; "midiio32.ocx"
Object = "{9BAC3ED0-40B2-44B9-ADE5-52766970ED57}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form Form_Init 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audiostation"
   ClientHeight    =   9105
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   9810
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Init.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Trm_Association 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2640
      Top             =   8040
   End
   Begin VB.Timer Trm_Search_CD_Device 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   8040
   End
   Begin VB.Timer Trm_Main 
      Interval        =   1000
      Left            =   240
      Tag             =   "0"
      Top             =   8040
   End
   Begin VB.Timer Trm_Search_Record_Device 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1680
      Top             =   8040
   End
   Begin VB.Timer Trm_Search_Midi_Device 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2160
      Top             =   8040
   End
   Begin VB.Timer Trm_Check_Drivers 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Tag             =   "0"
      Top             =   8040
   End
   Begin VB.PictureBox picOptions 
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
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
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
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
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
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
   Begin MidiioLib.MIDIOutput MidiOutput 
      Left            =   120
      Top             =   120
      _Version        =   65537
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      DeviceID        =   -1
      VolumeLeft      =   65535
      VolumeRight     =   65535
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "When all the settings have been set, the program will restart automatically."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   8550
      Width           =   7560
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checking device drivers..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   9
      Top             =   4875
      Width           =   9780
   End
   Begin VB.Label lblStatusTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   8
      Top             =   4680
      Width           =   9840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while audiostation configures the program settings for the first time."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   480
      TabIndex        =   7
      Top             =   2595
      Width           =   8010
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks for choosing Audiostation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   3600
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl2 
      Height          =   1365
      Left            =   360
      Top             =   840
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   2408
      Image           =   "Form_Init.frx":000C
      Frame           =   5
      Attr            =   514
      Effects         =   "Form_Init.frx":1BBFB
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
      Height          =   9105
      Left            =   0
      Top             =   0
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   16060
      Image           =   "Form_Init.frx":1BC13
      Settings        =   20480
      Effects         =   "Form_Init.frx":24174
   End
End
Attribute VB_Name = "Form_Init"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function FindRecordingDevice(deviceName As String) As Boolean
Dim I As Integer
Dim deviceList As String
Dim Lines
Dim strPos As Integer
Dim startReadingDevices As Boolean
Dim deviceFound As Boolean

deviceList = Extensions.FileGetContents(App.path & "\devices.txt")
Lines = Split(deviceList, vbNewLine)

For I = 0 To UBound(Lines)
    If startReadingDevices Then
        strPos = InStr(1, Lines(I), "]", vbTextCompare)
        
        If strPos > 0 Then
            If Replace(Trim(Mid(Lines(I), strPos + 1)), Chr(34), vbNullString) = deviceName Then
                deviceFound = True
            End If
        End If
    End If
    
    If InStr(1, Lines(I), "DirectShow audio devices", vbTextCompare) Then
        startReadingDevices = True
    End If
Next

FindRecordingDevice = deviceFound
End Function
Private Sub UnzipAppFiles(FileToUnzip As String)
On Error GoTo vbErrorHandler
Dim oUnZip As CGUnzipFiles

lblStatus.Caption = "Unpacking application files..."

Set oUnZip = New CGUnzipFiles

With oUnZip
    .ZipFileName = FileToUnzip
    .ExtractDir = App.path
    .HonorDirectories = True
    
    If .Unzip <> 0 Then
        MsgBox .GetLastMessage
    End If
End With

Set oUnZip = Nothing

Kill App.path & "\app_files_1_9_2019.zip"

Exit Sub

vbErrorHandler:
    MsgBox err.Number & " " & err.Description
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    
lblStatus.Caption = GetLanguage(1031)
Trm_Main.Enabled = True
End Sub
Private Function rgbGetVolumeLabel(CDPath As String) As String
Dim DrvVolumeName As String
Dim pos As Integer
Dim UnusedVal1 As Long
Dim UnusedVal2 As Long
Dim UnusedVal3 As Long
Dim UnusedStr As String
  
DrvVolumeName = Space$(14)
UnusedStr = Space$(32)

If GetVolumeInformation(CDPath, _
                            DrvVolumeName, _
                            Len(DrvVolumeName), _
                            UnusedVal1, UnusedVal2, _
                            UnusedVal3, _
                            UnusedStr, Len(UnusedStr)) > 0 Then
    
    pos = InStr(DrvVolumeName, Chr$(0))
    
    If pos Then DrvVolumeName = Left$(DrvVolumeName, pos - 1)
    If Len(Trim$(DrvVolumeName)) = 0 Then DrvVolumeName = "(no label)"
    
    rgbGetVolumeLabel = DrvVolumeName
End If
End Function

Private Sub Trm_Association_Timer()
Dim AppPath As String
Dim I As Integer

' Make sure all the elements are visible when the program opens
For I = 1 To 5
    Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "Element-" & I, "ON")
Next

On Error GoTo ErrorHandler

AppPath = App.path & "\" & App.EXEName & ".exe"

lblStatus.Caption = GetLanguage(1050)

AppPath = App.path + "\" + App.EXEName + ".exe"

'--- Bestanden ---
Call ModFileCommands.AssociateFile(".mp3", AppPath, "ext_mp3", "Mp3 audio file", App.path & "\win32.dll,10")
Call ModFileCommands.AssociateFile(".wav", AppPath, "ext_wav", "Wav audio file", App.path & "\win32.dll,14")
Call ModFileCommands.AssociateFile(".mid", AppPath, "ext_mid", "Midi audio file", App.path & "\win32.dll,8")
Call ModFileCommands.AssociateFile(".wma", AppPath, "ext_wma", "Wma audio file", App.path & "\win32.dll,18")
Call ModFileCommands.AssociateFile(".kar", AppPath, "ext_kar", "Kar audio file", App.path & "\win32.dll,5")
Call ModFileCommands.AssociateFile(".mp2", AppPath, "ext_mp2", "Mp2 audio file", App.path & "\win32.dll,9")
Call ModFileCommands.AssociateFile(".aac", AppPath, "ext_aac", "Aac audio file", App.path & "\win32.dll,0")
Call ModFileCommands.AssociateFile(".snd", AppPath, "ext_snd", "Snd audio file", App.path & "\win32.dll,16")
Call ModFileCommands.AssociateFile(".au", AppPath, "ext_au", "Au audio file", App.path & "\win32.dll,4")
Call ModFileCommands.AssociateFile(".rmi", AppPath, "ext_rmi", "Au audio file", App.path & "\win32.dll,14")
Call ModFileCommands.AssociateFile(".m4a", AppPath, "ext_m4a", "M4a audio file", App.path & "\win32.dll,9")
Call ModFileCommands.AssociateFile(".cda", AppPath, "ext_cda", "Cda audio file", App.path & "\win32.dll,4")
Call ModFileCommands.AssociateFile(".ra", AppPath, "ext_ra", "Ra audio file", App.path & "\win32.dll,0")
Call ModFileCommands.AssociateFile(".mus", AppPath, "ext_mus", "Mus audio file", App.path & "\win32.dll,11")
Call ModFileCommands.AssociateFile(".sid", AppPath, "ext_sid", "Sid audio file", App.path & "\win32.dll,15")

'--- Afspeellijsten ---
Call ModFileCommands.AssociateFile(".apl", AppPath, "ext_apl", "Audiostation playlist file", App.path & "\win32.dll,1")
Call ModFileCommands.AssociateFile(".pls", AppPath, "ext_pls", "ShoutCast playlist file", App.path & "\win32.dll,12")
Call ModFileCommands.AssociateFile(".m3u", AppPath, "ext_m3u", "Common playlist file", App.path & "\win32.dll,6")
Call ModFileCommands.AssociateFile(".wpl", AppPath, "ext_wpl", "Windows media player playlist file", App.path & "\win32.dll,19")

'--- Write Defaults ---
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "ApplicationFirstRun", "0")
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "OutputMode", 1)
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "Channel", 2)
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "Samples", 2)
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "Freq", "44100")
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CheckFile", "")

Trm_Main.Tag = 6
Trm_Main.Enabled = True

ErrorHandler:
Select Case err.Number
    Case 53
        MsgBox GetLanguage(1032), vbCritical
        End

End Select
End Sub

Private Sub Trm_Check_Drivers_Timer()
lblStatus.Caption = GetLanguage(1049)

' Check if dll exists, if so register the dll
If Not Dir(App.path & "\audio_sniffer.dll", vbDirectory) = vbNullString Then Extensions.RegUnReg (App.path & "\audio_sniffer.dll")
If Not Dir(App.path & "\audio_sniffer-x64.dll", vbDirectory) = vbNullString Then Extensions.RegUnReg (App.path & "\audio_sniffer-x64.dll")

Trm_Check_Drivers.Enabled = False

Trm_Main.Tag = 2
Trm_Main.Enabled = True
End Sub

Private Sub Trm_Main_Timer()
Dim Result As Long

Select Case Trm_Main.Tag
    Case 0
        Trm_Main.Enabled = False
        Trm_Main.Tag = 1
        
        If Not Dir(App.path & "\app_files_1_9_2019.zip", vbDirectory) = vbNullString Then
            Call UnzipAppFiles(App.path & "\app_files_1_9_2019.zip")
        End If
        
        Trm_Main.Enabled = True
        
    Case 1
        Trm_Main.Enabled = False
        Trm_Check_Drivers.Enabled = True
        
    Case 2
        Trm_Main.Enabled = False
        Trm_Search_CD_Device.Enabled = True
        
    Case 3
        Trm_Main.Enabled = False
        Trm_Search_Record_Device.Enabled = True

    Case 4
        Trm_Main.Enabled = False
        Trm_Search_Midi_Device.Enabled = True
    
    Case 5
        Trm_Main.Enabled = False
        Trm_Association.Enabled = True
        
    Case 6
        Trm_Main.Enabled = False
        
        Call Reg.SetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers\", App.path & "\" & App.EXEName & ".exe", vbNullString, REG_SZ, Result)
        
        Shell (App.path & "\AudioStation.exe")
        End
    
End Select
End Sub

Private Sub Trm_Search_CD_Device_Timer()
Dim r As Long
Dim DriveType As Long
Dim allDrives As String
Dim oneDrive As String
Dim CDLabel As String
Dim pos As Integer
Dim CDfound As Boolean

lblStatus.Caption = GetLanguage(1033)

CDfound = False
allDrives = Space$(64)
r = GetLogicalDriveStrings(Len(allDrives), allDrives)

If r > 0 Then
    allDrives = Left$(allDrives, r)
    Do
        pos = InStr(allDrives, Chr$(0))
        If pos Then
            oneDrive = Left$(allDrives, pos - 1)
            allDrives = Mid$(allDrives, pos + 1)
            DriveType = GetDriveType(oneDrive)
        
            If DriveType = DRIVE_CDROM Then
                CDfound = True
                CDLabel = rgbGetVolumeLabel(oneDrive)
                Exit Do
            End If
        End If
    Loop Until (allDrives = "") Or (DriveType = DRIVE_CDROM)
End If

If CDfound = True Then
    lblStatus.Caption = GetLanguage(1034) & " " & UCase$(oneDrive)
    
    Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CD", UCase$(oneDrive))
    
    Trm_Search_CD_Device.Enabled = False
    
    Trm_Main.Tag = 3
    Trm_Main.Enabled = True
Else
    MsgBox GetLanguage(1035), vbCritical
    Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CD", "false")
    
    Trm_Search_CD_Device.Enabled = False
    
    Trm_Main.Tag = 3
    Trm_Main.Enabled = True
End If
End Sub

Private Sub Trm_Search_Midi_Device_Timer()
Dim I As Integer
Dim deviceFound As Boolean
Dim DeviceFoundIndex As Integer

deviceFound = False

'Check if the midi device can be found
On Error Resume Next
For I = -1 To MidiOutput.DeviceCount - 1
    MidiOutput.DeviceID = I
    
    If InStr(1, LCase(MidiOutput.ProductName), "virtualmidisynth") > 0 Then
        deviceFound = True
        DeviceFoundIndex = I
    End If
Next

'The device is found
If deviceFound = True Then
    MidiOutput.DeviceID = DeviceFoundIndex
    lblStatus.Caption = GetLanguage(1047) & " " & MidiOutput.ProductName
    
    Call INI.WriteINI("SoundFonts", "sf1", App.path & "\devices\midi\default.sf2", "C:\Program Files\VirtualMIDISynth\VirtualMIDISynth.conf")
    Call INI.WriteINI("SoundFonts", "sf1", App.path & "\devices\midi\default.sf2", "C:\Program Files (x86)\VirtualMIDISynth\VirtualMIDISynth.conf")
    
    Call INI.WriteINI("SoundFonts", "sf1.Enabled", "1", "C:\Program Files\VirtualMIDISynth\VirtualMIDISynth.conf")
    Call INI.WriteINI("SoundFonts", "sf1.Enabled", "1", "C:\Program Files (x86)\VirtualMIDISynth\VirtualMIDISynth.conf")
    
    Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "MidiOutputDevice", DeviceFoundIndex + 1)
Else
    MsgBox GetLanguage(1038), vbExclamation
End If

Trm_Main.Tag = 5
Trm_Main.Enabled = True
    
Trm_Search_Midi_Device.Enabled = False
End Sub

Private Sub Trm_Search_Record_Device_Timer()
lblStatus.Caption = GetLanguage(1037)

Call Extensions.ShellAndWait("ffmpeg.exe", "-list_devices true -f dshow -i dummy > devices.txt 2>&1")

' Search for the record device
If FindRecordingDevice("virtual-audio-capturer") Then
    lblStatus.Caption = GetLanguage(1038) & " virtual-audio-capturer"
    Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "RecordingDevice", "virtual-audio-capturer")
   
    Trm_Search_Record_Device.Enabled = False
    
    Trm_Main.Tag = 4
    Trm_Main.Enabled = True
Else
    Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "RecordingDevice", "0")
   
    MsgBox GetLanguage(1039), vbCritical
    
    Trm_Search_Record_Device.Enabled = False
    
    Trm_Main.Tag = 4
    Trm_Main.Enabled = True
End If
End Sub
