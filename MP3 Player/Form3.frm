VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3525
   ClientLeft      =   3885
   ClientTop       =   3360
   ClientWidth     =   8310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   3525
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton O3 
      Caption         =   "Option1"
      Height          =   495
      Left            =   8400
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.OptionButton O2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Playlist Version (Advance)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   5055
   End
   Begin VB.OptionButton O1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nice Version (Simple)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   1800
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2760
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NTM MP3  Player"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   3255
      Left            =   120
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
O1.Value = False
O2.Value = False
O3.Value = True
End Sub
Private Sub Label2_Click()
If O1.Value = True Then
Form1.Show
Form3.Hide
Form2.Hide
Form2.Timer8.Enabled = True
End If
If O2.Value = True Then
Form2.Show
Form3.Hide
Form1.Hide
Form1.Timer7.Enabled = True
End If
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 1
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
End Sub

