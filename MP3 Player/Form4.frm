VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5535
   ClientLeft      =   11520
   ClientTop       =   2970
   ClientWidth     =   3045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   3045
   Begin VB.OptionButton O2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Don´t close after adding dir"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   5040
      Width           =   2295
   End
   Begin VB.OptionButton O1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close after adding dir"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   120
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   3840
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3690
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   1875
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   720
      Pattern         =   "*.mp3"
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
      Visible         =   0   'False
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   5295
      Left            =   120
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   5535
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COORD As POINT_TYPE, MOVEok As Boolean
Dim MOVE2ok As Boolean, MOVE3ok As Boolean
Dim SPOTx, SPOTy, MOVE4ok As Boolean
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
File1.Path = Dir1.Path
If File1.ListCount <> 0 Then
For NTM = 1 To File1.ListCount
File1.ListIndex = NTM - 1
If Len(Dir1.Path) > 3 Then
Form2.List1.AddItem Dir1.Path & "\" & File1.FileName
Else
Form2.List1.AddItem Dir1.Path & File1.FileName
Form2.cmdClear.Enabled = True
Form2.cmdSave.Enabled = True
Form2.cmdPause.Enabled = True
End If
Next NTM
Timer2.Enabled = True
Else
MsgBox "No files were found in specific folder", vbOKOnly, "Error"
End If
End Sub
Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub
Private Sub Form_Load()
Timer2.Enabled = False
O1.Value = True
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SPOTx = X
SPOTy = Y
MOVE3ok = True
Timer1.Interval = 10
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVE3ok = False
Timer1.Interval = 0
End Sub
Private Sub Timer1_Timer()
If MOVE3ok = True Then
GetCursorPos COORD
Form4.Left = COORD.X * Screen.TwipsPerPixelX - SPOTx
Form4.Top = COORD.Y * Screen.TwipsPerPixelY - SPOTy
End If
End Sub
Private Sub Timer2_Timer()
If O1.Value = True Then
Unload Me
End If
If O2.Value = True Then
Form4.Show
End If
Timer2.Enabled = False
End Sub
