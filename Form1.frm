VERSION 5.00
Object = "{1015925F-1CD1-11D6-895C-002078085A01}#1.0#0"; "SNDPlayer.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Channel Sound Player"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin BCSoundPlayer.SoundPlayer SoundPlayer1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop All"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "D:\nssc2\mover.wav"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "D:\mp3\Crystal Method\MoonshineMusic.mp3"
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Loop"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":008E
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the filenames of the sound files you want to play."
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If CheckFile(Text1.Text) = False Then Exit Sub
Text1.Tag = SoundPlayer1.PlaySound(Text1.Text, False)
End Sub

Private Sub Command2_Click()
If Text1.Tag = "" Then Exit Sub
SoundPlayer1.StopSound CLng(Text1.Tag)
Text1.Tag = ""
End Sub

Private Sub Command3_Click()
If Text2.Tag = "" Then Exit Sub
SoundPlayer1.StopSound CLng(Text2.Tag)
Text2.Tag = ""
End Sub

Private Sub Command4_Click()
If CheckFile(Text2.Text) = False Then Exit Sub
Text2.Tag = SoundPlayer1.PlaySound(Text2.Text, True)
End Sub

Private Sub Command5_Click()
SoundPlayer1.StopAll
Text1.Tag = ""
Text2.Tag = ""
End Sub

Private Function CheckFile(sTXT As String) As Boolean
If sTXT = "" Or Dir(sTXT) = "" Then
MsgBox "Opps..." & vbCrLf & "Please enter a file to play.", vbInformation, "User Error"
CheckFile = False
Else
CheckFile = True
End If
End Function
