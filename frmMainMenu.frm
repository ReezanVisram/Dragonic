VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dragonic"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   5625
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   3
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label lblHowToPlay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "How to Play"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   2
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label lblPlay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "Dragonic"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   3720
      TabIndex        =   0
      Top             =   -240
      Width           =   7815
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPlay.ForeColor = &H80000012
lblHowToPlay.ForeColor = &H80000012
lblCredits.ForeColor = &H80000012
End Sub


Private Sub lblCredits_Click()
frmMainMenu.Hide
frmCredits.Show
End Sub

Private Sub lblCredits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 0 And X <= 3615 And Y >= 0 And Y <= 975 Then
    lblCredits.ForeColor = &HFF0000
End If
lblPlay.ForeColor = &H80000012
lblHowToPlay.ForeColor = &H80000012
End Sub

Private Sub lblHowToPlay_Click()
frmMainMenu.Hide
frmHowToPlay.Show
End Sub

Private Sub lblHowToPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 0 And X <= 3615 And Y >= 0 And Y <= 975 Then
    lblHowToPlay.ForeColor = &HFF0000
End If
lblPlay.ForeColor = &H80000012
lblCredits.ForeColor = &H80000012
End Sub

Private Sub lblPlay_Click()
frmCutscene.Show
frmMainMenu.Hide
End Sub

Private Sub lblPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 0 And X <= 1335 And Y >= 0 And Y <= 975 Then
    lblPlay.ForeColor = &HFF0000
End If
lblHowToPlay.ForeColor = &H80000012
lblCredits.ForeColor = &H80000012
End Sub

