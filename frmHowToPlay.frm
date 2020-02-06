VERSION 5.00
Begin VB.Form frmHowToPlay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "How To Play"
   ClientHeight    =   10380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHowToPlay.frx":0000
   ScaleHeight     =   10380
   ScaleWidth      =   17670
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblBackToMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back To Main Menu"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   15600
      TabIndex        =   10
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label lblObjective 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The objective is to reach the open end of each level to progress to the next, collecting powerups and doging enemies along the way"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   0
      TabIndex        =   9
      Top             =   6960
      Width           =   6015
   End
   Begin VB.Label lblClick 
      BackStyle       =   0  'Transparent
      Caption         =   "Left Click - Shoot fire"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   6360
      Width           =   3735
   End
   Begin VB.Label lblD 
      BackStyle       =   0  'Transparent
      Caption         =   "D - Fly Right"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label lblS 
      BackStyle       =   0  'Transparent
      Caption         =   "S - Fly Down"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      Caption         =   "A - Fly Left"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label lblW 
      BackStyle       =   0  'Transparent
      Caption         =   "W - Fly up"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label lblControls 
      BackStyle       =   0  'Transparent
      Caption         =   "Controls:"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   720
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblBlue 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1215
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblRed 
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   7320
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblHowToPlay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHowToPlay.frx":C4EFA
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   16575
   End
End
Attribute VB_Name = "frmHowToPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBackToMenu.ForeColor = &HFFFFFF
End Sub

Private Sub lblBackToMenu_Click()
frmHowToPlay.Hide
frmMainMenu.Show
End Sub

Private Sub lblBackToMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 0 And X <= lblBackToMenu.Width And Y >= 0 And Y <= lblBackToMenu.Height Then
    lblBackToMenu.ForeColor = &HFF&
End If
End Sub

