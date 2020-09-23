VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   -9345
   ClientTop       =   3210
   ClientWidth     =   7575
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdFastB 
      Caption         =   ">"
      Height          =   195
      Left            =   4680
      TabIndex        =   31
      Top             =   2280
      Width           =   315
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "l<"
      Height          =   195
      Left            =   3720
      TabIndex        =   30
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">l"
      Height          =   195
      Left            =   5160
      TabIndex        =   29
      Top             =   2280
      Width           =   375
   End
   Begin VB.Timer Timer11 
      Interval        =   150
      Left            =   6600
      Top             =   4800
   End
   Begin VB.Timer Timer10 
      Interval        =   150
      Left            =   4920
      Top             =   4920
   End
   Begin VB.CommandButton cmdFastF 
      Caption         =   "<"
      Height          =   195
      Left            =   4200
      TabIndex        =   28
      Top             =   2280
      Width           =   315
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   2280
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   450
      _Version        =   393216
      Max             =   2500
      SelStart        =   2500
      TickStyle       =   3
      Value           =   2500
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   9600
      TabIndex        =   24
      Top             =   600
      Width           =   375
   End
   Begin VB.Timer Timer9 
      Interval        =   1
      Left            =   9480
      Top             =   3600
   End
   Begin VB.Timer Timer8 
      Interval        =   300
      Left            =   8640
      Top             =   1680
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   9000
      Top             =   3600
   End
   Begin VB.PictureBox PicBorder 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   7140
      TabIndex        =   16
      Top             =   840
      Width           =   7140
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear List"
         Height          =   375
         Left            =   6000
         TabIndex        =   23
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save List"
         Height          =   375
         Left            =   5040
         TabIndex        =   22
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load List"
         Height          =   375
         Left            =   4080
         TabIndex        =   21
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   3120
         TabIndex        =   19
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8040
      Top             =   3120
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9120
      TabIndex        =   7
      Top             =   5760
      Width           =   495
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   8520
      Top             =   3120
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   6120
      Width           =   855
   End
   Begin VB.Timer Timer5 
      Left            =   9000
      Top             =   3120
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   4800
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9480
      Top             =   3120
   End
   Begin VB.PictureBox Picture12 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   6720
      Picture         =   "Form2.frx":0BC2
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   270
      Width           =   225
   End
   Begin VB.PictureBox Picture14 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   7050
      Picture         =   "Form2.frx":0F79
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   270
      Width           =   225
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   8040
      Top             =   3600
   End
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   8520
      Top             =   3600
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   635
      _Version        =   393216
      Max             =   100
      TickStyle       =   3
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   8040
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open MP3"
      Filter          =   "MP3 Files [*.mp3]|*.mp3"
   End
   Begin VB.PictureBox Picture13 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   6720
      Picture         =   "Form2.frx":1358
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   270
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture15 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   7050
      Picture         =   "Form2.frx":170F
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   270
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Playlist"
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   7095
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000005&
         Height          =   1590
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   6855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   9240
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   8640
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2760
      TabIndex        =   27
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1920
      TabIndex        =   26
      Top             =   2280
      Width           =   855
   End
   Begin MediaPlayerCtl.MediaPlayer MP3 
      Height          =   255
      Left            =   7920
      TabIndex        =   13
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
   Begin VB.Label A 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   7095
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
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
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
      TabIndex        =   10
      Top             =   240
      Width           =   6375
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
      TabIndex        =   14
      Top             =   0
      Width           =   7575
   End
   Begin VB.Menu mnuA 
      Caption         =   "NTM"
      Visible         =   0   'False
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuRemoveA 
         Caption         =   "Remove All"
      End
      Begin VB.Menu mnuFileName 
         Caption         =   "Show Filename"
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
      Begin VB.Menu mnuSong 
         Caption         =   "Curent Song"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuP 
         Caption         =   "Pause"
      End
      Begin VB.Menu mhuazd 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuNext 
         Caption         =   "Next"
      End
      Begin VB.Menu mnuPrevious 
         Caption         =   "Previous"
      End
      Begin VB.Menu mnukksa 
         Caption         =   "-"
      End
      Begin VB.Menu zdfa 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuNTM3 
      Caption         =   "NTM3"
      Visible         =   0   'False
      Begin VB.Menu mnuFile 
         Caption         =   "File"
      End
      Begin VB.Menu mnuDir 
         Caption         =   "Directory"
      End
   End
End
Attribute VB_Name = "Form2"
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
Private Sub cmdBack_Click()
mnuPrevious_Click
End Sub
Private Sub cmdClear_Click()
List1.Clear
cmdStop_Click
Text4.Text = "ka"
cmdClear.Enabled = False
cmdSave.Enabled = False
mnuFileName.Enabled = False
End Sub
Private Sub cmdDir_Click()
Form4.Show
End Sub
Private Sub cmdFastB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer11.Enabled = True
End Sub
Private Sub cmdFastB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer11.Enabled = False
End Sub
Private Sub cmdFastF_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdPause_Click
cmdPlay.Enabled = False
Timer10.Enabled = True
End Sub
Private Sub cmdFastF_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdPause_Click
cmdPlay.Enabled = True
Timer10.Enabled = False
End Sub
Private Sub cmdLoad_Click()
Dim file As String
CommonDialog2.DialogTitle = "Load your list."
CommonDialog2.MaxFileSize = 16384
CommonDialog2.FileName = ""
CommonDialog2.Filter = "NMP Playlist Files [*.NMP]|*.NMP|NMP Playlist Files [*.NTM]|*.NTM"
CommonDialog2.ShowOpen
If CommonDialog2.FileName = "" Then Exit Sub
file = CommonDialog2.FileName
Dim A As String
On Error GoTo error
cmdClear_Click
Open file For Input As #1
Do Until EOF(1)
Input #1, A$
List1.AddItem A$
Loop
Close 1
Timer3.Enabled = True
Timer4.Enabled = True
Text4.Text = ""
cmdClear.Enabled = True
Timer7.Enabled = True
cmdSave.Enabled = True
Exit Sub
error:
MsgBox "File Not Found", vbOKOnly, "Error"
End Sub
Private Sub cmdNext_Click()
mnuNext_Click
End Sub
Private Sub cmdOpen_Click()
Form2.PopupMenu mnuNTM3
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
Private Sub cmdSave_Click()
Dim file As String
CommonDialog2.DialogTitle = "Save your list."
CommonDialog2.MaxFileSize = 16384
CommonDialog2.FileName = ""
CommonDialog2.Filter = "Playlist File [*.NMP]|*.NMP |Playlist File [*.NTM]|*.NTM"
CommonDialog2.ShowSave
If CommonDialog2.FileName = "" Then Exit Sub
file = CommonDialog2.FileName
Open file For Output As #1
Dim i%
For i = 0 To List1.ListCount - 1
Print #1, List1.List(i)
Next
Close #1
MsgBox "The playlist was saved in" & vbCrLf & file & vbCrLf & "By NTM Mp3 Player @ " & Time
End Sub
Private Sub cmdStop_Click()
MP3.Stop
cmdPause.Enabled = False
mnuP.Enabled = False
mnuPlay.Enabled = True
cmdPlay.Enabled = True
cmdSave.Enabled = True
cmdClear.Enabled = True
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
cmdNext.Enabled = False
cmdBack.Enabled = False
cmdFastF.Enabled = False
cmdFastB.Enabled = False
End Sub
Private Sub Form_Load()
On Error Resume Next
cmdSave.Enabled = False
cmdClear.Enabled = False
Timer1.Enabled = False
cmdPlay.Enabled = False
mnuPlay.Enabled = False
mhuazd.Enabled = False
mnuSong.Enabled = False
cmdStop.Enabled = False
mnuNext.Enabled = False
mnuPrevious.Enabled = False
cmdNext.Enabled = False
cmdBack.Enabled = False
cmdFastF.Enabled = False
cmdFastB.Enabled = False
mnuFileName.Enabled = False
List1.Enabled = True
cmdPause.Enabled = False
mnuP.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
List1.OLEDropMode = vbOLEDropManual
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If RespondToTray(x) = 1 Then
ShowFormAgain Me
Form2.WindowState = 0 - Normal
End If
If RespondToTray(x) = 2 Then
Me.PopupMenu mnuk
End If
End Sub
Private Sub label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture12.Visible = True
Picture13.Visible = False
Picture14.Visible = True
Picture15.Visible = False
End Sub
Private Sub Label1_Click()
MsgBox "This program was " & "Made By NeoTheMaster @ Darkside Reborn" & vbCrLf & "Uptime: " & GetUptime & vbCrLf & "Du har " & List1.ListCount & " låtar i playlist", vbInformation, "NTM"
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
SPOTx = x
SPOTy = y
MOVE3ok = True
Timer5.Interval = 10
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MOVE3ok = False
Timer5.Interval = 0
End Sub
Private Sub List1_DblClick()
On Error Resume Next
Timer3.Enabled = True
Timer4.Enabled = True
Text4.Text = ""
cmdClear.Enabled = True
cmdSave.Enabled = True
Timer1.Enabled = True
Timer2.Enabled = True
MP3.FileName = List1.Text
PL = MP3.GetMediaInfoString(mpClipAuthor) & " - " & MP3.GetMediaInfoString(mpClipTitle)
A.Caption = PL
Slider1.Max = MP3.Duration
cmdPause.Enabled = True
Timer3.Enabled = True
End Sub
Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
If Not Data.GetFormat(vbCFFiles) Then
Effect = vbDropEffectNone
Exit Sub
End If
For i = 1 To Data.Files.Count
List1.AddItem Data.Files(i)
Next
End Sub
Private Sub List1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Data.GetFormat(vbCFFiles) Then
Effect = vbDropEffectCopy
Else
Effect = vbDropEffectNone
End If
End Sub
Private Sub mhuazd_Click()
cmdStop_Click
End Sub
Private Sub mhuzfd_Click()
cmdOpen_Click
End Sub
Private Sub mnuaav_Click()
Picture15_Click
End Sub
Private Sub mnuDir_Click()
Form4.Show
End Sub
Private Sub mnuFile_Click()
CD.FileName = MP3.FileName
CD.ShowOpen
If CD.FileName = "" Then
cmdSave.Enabled = False
cmdClear.Enabled = False
Text4.Text = "ka"
Exit Sub
End If
If CD.FileName = MP3.FileName Then
GoTo Crapy
Else
MP3.FileName = CD.FileName
Slider1.Max = MP3.Duration
A.Caption = MP3.GetMediaInfoString(mpClipAuthor) & " - " & MP3.GetMediaInfoString(mpClipTitle)
List1.AddItem MP3.FileName
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Text4.Text = ""
cmdClear.Enabled = True
cmdSave.Enabled = True
Crapy:
End If
End Sub
Private Sub mnuFileName_Click()
Dim x, y, k, b
k = List1.Text
x = InStrRev(k, "\")
y = k
b = Mid(y, x + 1)
MsgBox b, vbOKOnly, "NMP"
End Sub
Private Sub mnuNext_Click()
On Error GoTo k
List1.ListIndex = List1.ListIndex + 1
List1_DblClick
Exit Sub
k:
List1.ListIndex = 0
MP3.Open List1.ListIndex = 0
List1_DblClick
End Sub
Private Sub mnuP_Click()
cmdPause_Click
End Sub
Private Sub mnuPlay_Click()
cmdPlay_Click
End Sub
Private Sub mnuPrevious_Click()
If List1.ListIndex = 0 Then
cmdStop_Click
cmdPlay_Click
Exit Sub
Else
List1.ListIndex = List1.ListIndex - 1
List1_DblClick
End If
End Sub
Private Sub mnuR_Click()
ShowFormAgain Me
Form2.WindowState = 0 - Normal
End Sub
Private Sub mnuRemove_Click()
On Error Resume Next
If List1.ListIndex = -1 Then
MsgBox "No file selected", vbExclamation, "Error"
Else
List1.RemoveItem List1.ListIndex
On Error GoTo k
List1.ListIndex = List1.ListIndex + 1
k:
On Error Resume Next
If List1.ListIndex = 0 Then
Else
mnuRemoveA_Click
Exit Sub
End If
End If
End Sub
Private Sub mnuRemoveA_Click()
List1.Clear
cmdStop_Click
Text4.Text = "ka"
cmdClear.Enabled = False
cmdSave.Enabled = False
cmdNext.Enabled = False
cmdBack.Enabled = False
cmdFastF.Enabled = False
cmdFastB.Enabled = False
mnuFileName.Enabled = False
End Sub
Private Sub mnuSong_Click()
If A.Caption = " - " Then
MsgBox "Current Song: " & MP3.FileName, vbOKOnly, "NMP"
'MsgBox "Current Song: " & MP3.FileName & vbCrLf & "Current Time: " & Label2.Caption, vbOKOnly, "NMP"
Else
MsgBox "Current Song: " & A.Caption, vbOKOnly, "NMP"
'MsgBox "Current Song: " & A.Caption & vbCrLf & "Current Time: " & Label2.Caption, vbOKOnly, "NMP"
End If
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
List1.ListIndex = List1.ListIndex + 1
MP3.Open List1.ListIndex + 1
List1_DblClick
Exit Sub
error:
List1.ListIndex = 0
MP3.Open List1.ListIndex = 0
List1_DblClick
End Sub
Private Sub Picture13_Click()
Form2.WindowState = vbMinimized
End Sub
Private Sub Picture15_Click()
MP3.Stop
End
End Sub
Private Sub Picture12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture12.Visible = False
Picture13.Visible = True
Picture14.Visible = True
Picture15.Visible = False
End Sub
Private Sub Picture14_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture12.Visible = True
Picture13.Visible = False
Picture14.Visible = False
Picture15.Visible = True
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
Me.PopupMenu mnuA
End If
End Sub
Private Sub List1_DubbelClick()
On Error GoTo b
MP3.FileName = List1.Text
PL = MP3.GetMediaInfoString(mpClipAuthor) & " - " & MP3.GetMediaInfoString(mpClipTitle)
A.Caption = PL
Text4.Text = MP3.FileName
Slider1.Max = MP3.Duration
cmdPause.Enabled = True
Timer3.Enabled = True
b:
End Sub
Private Sub Slider1_Scroll()
MP3.CurrentPosition = Slider1.Value
End Sub
Private Sub Slider2_Scroll()
Dim Pi, S
Dim F As Integer, P As Integer
If Slider2.Value = 0 Then
MP3.Mute = True
Else
MP3.Mute = False
End If
S = Slider2.Value - 2500
MP3.Volume = S
On Error GoTo H
P = Slider2.min
F = Slider2.Value
Label4.Caption = F \ 25 & " %"
H:
Exit Sub
End Sub
Private Sub Timer1_Timer()
If Text4.Text = "ka" Then
cmdPlay.Enabled = False
mnuPlay.Enabled = False
cmdStop.Enabled = False
mhuazd.Enabled = False
Exit Sub
End If
If A.Caption = "" Then
cmdPlay.Enabled = True
cmdPause.Enabled = False
mnuP.Enabled = False
mnuPlay.Enabled = True
cmdStop.Enabled = False
mhuazd.Enabled = False
mnuSong.Enabled = False
mnuNext.Enabled = False
mnuPrevious.Enabled = False
Else
cmdSave.Enabled = True
cmdClear.Enabled = True
cmdPlay.Enabled = False
cmdPause.Enabled = True
mnuP.Enabled = True
mnuPlay.Enabled = False
cmdStop.Enabled = True
mhuazd.Enabled = True
mnuSong.Enabled = True
mnuNext.Enabled = True
mnuPrevious.Enabled = True
cmdNext.Enabled = True
cmdBack.Enabled = True
cmdFastF.Enabled = True
cmdFastB.Enabled = True
End If
If A.Caption = " - " Then
A.Caption = MP3.FileName
End If
If List1.Text = "" Then
mnuFileName.Enabled = False
Else
mnuFileName.Enabled = True
End If
End Sub
Private Sub Timer10_Timer()
On Error Resume Next
MP3.CurrentPosition = MP3.CurrentPosition - 1
End Sub
Private Sub Timer11_Timer()
MP3.CurrentPosition = MP3.CurrentPosition + 1
End Sub
Private Sub Timer2_Timer()
If Form2.WindowState = 1 - minimized Then
AddToTray Me.Icon, "NTM MP3 Player Playlist Version", Me
Timer2.Enabled = False
Timer6.Enabled = True
End If
End Sub
Private Sub Timer6_Timer()
If Form2.WindowState = 0 - Normal Then
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
Dim b As Integer
A = twinsec \ 60
b = twinsec - (A * 60)
If b = "-1" Then sec = "0"
 Text2.Text = A & ":" & b
End Sub
Private Sub Timer5_Timer()
If MOVE3ok = True Then
GetCursorPos COORD
Form2.Left = COORD.x * Screen.TwipsPerPixelX - SPOTx
Form2.Top = COORD.y * Screen.TwipsPerPixelY - SPOTy
End If
End Sub
Private Sub Timer7_Timer()
List1.ListIndex = 0
MP3.Open List1.ListIndex = 0
List1_DblClick
Timer7.Enabled = False
End Sub
Private Sub Timer8_Timer()
Form2.WindowState = 1 - minimized
If Form2.WindowState = 1 - minimized Then
Form2.WindowState = 0 - Normal
End If
If Form2.WindowState = 0 - Normal Then
Me.Hide
End If
Timer8.Enabled = False
End Sub
Private Sub zdfa_Click()
Picture15_Click
End Sub

