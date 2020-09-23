VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4410
   ClientLeft      =   17310
   ClientTop       =   6225
   ClientWidth     =   7575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   8520
      Top             =   3600
   End
   Begin VB.Timer Timer7 
      Interval        =   300
      Left            =   9000
      Top             =   3600
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   8040
      Top             =   3600
   End
   Begin VB.PictureBox Picture14 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   7050
      Picture         =   "Form1.frx":0BC2
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   270
      Width           =   225
   End
   Begin VB.PictureBox Picture12 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   6720
      Picture         =   "Form1.frx":0FA1
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   270
      Width           =   225
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Playlist"
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   7095
      Begin VB.ListBox PlayList 
         BackColor       =   &H00800000&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   1545
         IntegralHeight  =   0   'False
         ItemData        =   "Form1.frx":1358
         Left            =   120
         List            =   "Form1.frx":135A
         TabIndex        =   2
         Top             =   240
         Width           =   6855
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   360
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   635
      _Version        =   393216
      Max             =   100
      TickStyle       =   3
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9480
      Top             =   3120
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   9120
      TabIndex        =   19
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7680
      TabIndex        =   16
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Timer Timer5 
      Left            =   9000
      Top             =   3120
   End
   Begin VB.PictureBox PicBorder 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1320
      ScaleHeight     =   375
      ScaleWidth      =   4860
      TabIndex        =   12
      Top             =   840
      Width           =   4860
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7920
      TabIndex        =   11
      Top             =   6120
      Width           =   855
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   8520
      Top             =   3120
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9120
      TabIndex        =   10
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox Picture13 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   6720
      Picture         =   "Form1.frx":135C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   270
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture15 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   7050
      Picture         =   "Form1.frx":1713
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   270
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   360
      TabIndex        =   0
      Top             =   6000
      Width           =   7095
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8040
      Top             =   3120
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   8160
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open MP3"
      Filter          =   "MP3 Files [*.mp3]|*.mp3"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NTM MP3  Player"
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
      Height          =   330
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label A 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   7095
   End
   Begin MediaPlayerCtl.MediaPlayer MP3 
      Height          =   255
      Left            =   7920
      TabIndex        =   9
      Top             =   4200
      Width           =   2175
      Visible         =   0   'False
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   4455
      Left            =   120
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   4695
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7575
   End
   Begin VB.Menu mnuA 
      Caption         =   "NTM"
      Visible         =   0   'False
      Begin VB.Menu mnuaa 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuAAA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaav 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuk 
      Caption         =   "NTM2"
      Visible         =   0   'False
      Begin VB.Menu mnuR 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mhuzfd 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuP 
         Caption         =   "Pause"
      End
      Begin VB.Menu mhuazd 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnukksa 
         Caption         =   "-"
      End
      Begin VB.Menu zdfa 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COORD As POINT_TYPE, MOVEok As Boolean
Dim MOVE2ok As Boolean, MOVE3ok As Boolean
Dim SPOTx, SPOTy, MOVE4ok As Boolean
Dim PL As String
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Sub cmdOpen_Click()
Timer1.Enabled = True
On Error Resume Next
CD.FileName = MP3.FileName
CD.ShowOpen
If CD.FileName = MP3.FileName Then
GoTo Crapy
Else
On Error Resume Next
MP3.FileName = CD.FileName
Slider1.Max = MP3.Duration
A.Caption = MP3.GetMediaInfoString(mpClipAuthor) & " - " & MP3.GetMediaInfoString(mpClipTitle)
PlayList.AddItem List1.ListCount + 1 & ". " & MP3.GetMediaInfoString(mpClipAuthor) & " - " & MP3.GetMediaInfoString(mpClipTitle)
List1.AddItem MP3.FileName
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
cmdPause.Enabled = True
Crapy:
End If
End Sub
Private Sub cmdPause_Click()
If Text1.Text = "" Then
Timer3.Enabled = False
MP3.Pause
Text1.Text = "k"
Timer1.Enabled = False
cmdPlay.Enabled = True
mnuPlay.Enabled = True
Else
cmdPlay_Click
Text1.Text = ""
cmdPlay.Enabled = False
mnuPlay.Enabled = False
End If
End Sub
Private Sub cmdPlay_Click()
MP3.Play
Timer1.Enabled = True
A.Caption = MP3.GetMediaInfoString(mpClipAuthor) & " - " & MP3.GetMediaInfoString(mpClipTitle)
cmdPause.Enabled = True
mnuP.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Slider1.Max = MP3.Duration
Text1.Text = ""
End Sub
Private Sub cmdStop_Click()
MP3.Stop
cmdPause.Enabled = False
mnuP.Enabled = False
mnuPlay.Enabled = True
cmdPlay.Enabled = True
Text1.Text = ""
Text2.Text = ""
txtTime.Text = ""
Label2.Caption = ""
A.Caption = ""
Slider1.Value = 0
MP3.CurrentPosition = 0
CD.FileName = ""
Timer3.Enabled = False
Timer4.Enabled = False
End Sub
Private Sub Form_Load()
Timer1.Enabled = False
Timer7.Enabled = False
cmdPlay.Enabled = False
mnuPlay.Enabled = False
PlayList.Enabled = True
cmdPause.Enabled = False
mnuP.Enabled = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If RespondToTray(X) = 1 Then
ShowFormAgain Me
Form1.WindowState = 0 - Normal
End If
If RespondToTray(X) = 2 Then
Me.PopupMenu mnuk
End If
End Sub
Private Sub label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture12.Visible = True
Picture13.Visible = False
Picture14.Visible = True
Picture15.Visible = False
End Sub
Private Sub Label1_Click()
MsgBox "This program was " & "Made By NeoTheMaster @ Darkside Reborn" & vbCrLf & "Uptime: " & GetUptime & vbCrLf & "Du har " & List1.ListCount & " låtar i playlist", vbInformation, "NTM"
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SPOTx = X
SPOTy = Y
MOVE3ok = True
Timer5.Interval = 10
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVE3ok = False
Timer5.Interval = 0
End Sub
Private Sub mhuazd_Click()
cmdStop_Click
End Sub
Private Sub mhuzfd_Click()
cmdOpen_Click
End Sub
Private Sub mnuaa_Click()
cmdOpen_Click
End Sub
Private Sub mnuaav_Click()
Picture15_Click
End Sub
Private Sub mnuP_Click()
cmdPause_Click
End Sub
Private Sub mnuPlay_Click()
cmdPlay_Click
End Sub
Private Sub mnuR_Click()
ShowFormAgain Me
Form1.WindowState = 0 - Normal
End Sub
Private Sub MP3_EndOfStream(ByVal Result As Long)
cmdPause.Enabled = False
Text1.Text = ""
Text2.Text = ""
txtTime.Text = ""
Label2.Caption = ""
A.Caption = ""
Slider1.Value = 0
On Error GoTo error
PlayList.ListIndex = PlayList.ListIndex + 1
MP3.Open List1.ListIndex + 1
Playlist_DblClick
Exit Sub
error:
PlayList.ListIndex = 0
MP3.Open PlayList.ListIndex = 0
Playlist_DblClick
End Sub
Private Sub Picture13_Click()
Form1.WindowState = vbMinimized
End Sub
Private Sub Picture15_Click()
MP3.Stop
End
End Sub
Private Sub Picture12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture12.Visible = False
Picture13.Visible = True
Picture14.Visible = True
Picture15.Visible = False
End Sub
Private Sub Picture14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture12.Visible = True
Picture13.Visible = False
Picture14.Visible = False
Picture15.Visible = True
End Sub
Private Sub Playlist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
Me.PopupMenu mnuA
End If
End Sub
Private Sub Playlist_DblClick()
On Error GoTo B
List1.ListIndex = PlayList.ListIndex
MP3.FileName = List1.Text
PL = MP3.GetMediaInfoString(mpClipAuthor) & " - " & MP3.GetMediaInfoString(mpClipTitle)
A.Caption = PL
Slider1.Max = MP3.Duration
cmdPause.Enabled = True
Timer3.Enabled = True
B:
End Sub
Private Sub Slider1_Scroll()
MP3.CurrentPosition = Slider1.Value
End Sub
Private Sub Timer1_Timer()
If A.Caption = "" Then
cmdPlay.Enabled = True
cmdPause.Enabled = False
mnuP.Enabled = False
mnuPlay.Enabled = True
Else
cmdPlay.Enabled = False
cmdPause.Enabled = True
mnuP.Enabled = True
mnuPlay.Enabled = False
End If
End Sub
Private Sub Timer2_Timer()
If Form1.WindowState = 1 - minimized Then
AddToTray Me.Icon, "NTM MP3 Player Nice Version", Me
Timer2.Enabled = False
Timer6.Enabled = True
End If
End Sub
Private Sub Timer6_Timer()
If Form1.WindowState = 0 - Normal Then
ShowFormAgain Me
Timer2.Enabled = True
Timer6.Enabled = False
End If
End Sub
Private Sub Timer3_Timer()
Label2.Caption = txtTime.Text & " / " & Text2.Text
End Sub
Private Sub Timer4_Timer()
tinseconden = MP3.CurrentPosition
Dim min As Integer
Dim sec As Integer
min = tinseconden \ 60
sec = tinseconden - (min * 60)
If sec = "-1" Then sec = "0"
txtTime.Text = min & ":" & sec
Slider1.Value = MP3.CurrentPosition
twinsec = MP3.Duration
Dim A As Integer
Dim B As Integer
A = twinsec \ 60
B = twinsec - (A * 60)
If B = "-1" Then sec = "0"
 Text2.Text = A & ":" & B
End Sub
Private Sub Timer5_Timer()
If MOVE3ok = True Then
GetCursorPos COORD
Form1.Left = COORD.X * Screen.TwipsPerPixelX - SPOTx
Form1.Top = COORD.Y * Screen.TwipsPerPixelY - SPOTy
End If
End Sub
Private Sub Timer7_Timer()
Me.Hide
Timer7.Enabled = False
End Sub
Private Sub zdfa_Click()
Picture15_Click
End Sub
