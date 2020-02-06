VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCredits.frx":0000
   ScaleHeight     =   7560
   ScaleWidth      =   15300
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblBackToMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label lblContact 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contact me at: reezan.visram@rogers.com"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1560
      TabIndex        =   1
      Top             =   6960
      Width           =   12255
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCredits.frx":18937
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   12255
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBackToMenu.ForeColor = &HFFFFFF
End Sub

Private Sub lblBackToMenu_Click()
frmCredits.Hide
frmMainMenu.Show
End Sub

Private Sub lblBackToMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 0 And X <= 1455 And Y >= 0 And Y <= 1455 Then
    lblBackToMenu.ForeColor = &HFF&
End If
End Sub

Private Sub lblCredits_Click()

End Sub
