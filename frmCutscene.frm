VERSION 5.00
Begin VB.Form frmCutscene 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battle"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCutscene.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerSecondaryBattle 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   11640
      Top             =   120
   End
   Begin VB.Timer timerDragonBattle 
      Interval        =   300
      Left            =   5520
      Top             =   240
   End
   Begin VB.Image imgLightning 
      Height          =   735
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Image imgFireBreath 
      Height          =   735
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Image imgBlueDragon 
      Height          =   1815
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Image imgRedDragon 
      Height          =   495
      Left            =   480
      Stretch         =   -1  'True
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "frmCutscene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dragonCount As Integer
Dim secondaryDragonCount As Integer


Private Sub timerRedDragon_Timer()

End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub timerDragonBattle_Timer()
dragonCount = dragonCount + 1

Select Case dragonCount
Case Is = 1
    imgRedDragon.Picture = frmPictures.imgDragonForward1.Picture
    imgRedDragon.Width = imgRedDragon.Width + 50
    imgRedDragon.Height = imgRedDragon.Height + 50

Case Is = 2
    imgRedDragon.Picture = frmPictures.imgDragonForward2.Picture
    imgRedDragon.Width = imgRedDragon.Width + 50
    imgRedDragon.Height = imgRedDragon.Height + 50
    
Case Is = 3
    imgRedDragon.Picture = frmPictures.imgDragonForward3.Picture
    imgRedDragon.Width = imgRedDragon.Width + 50
    imgRedDragon.Height = imgRedDragon.Height + 50

Case Is = 4
    imgRedDragon.Picture = frmPictures.imgDragonForward1.Picture
    imgRedDragon.Width = imgRedDragon.Width + 50
    imgRedDragon.Height = imgRedDragon.Height + 50

Case Is = 5
    imgRedDragon.Picture = frmPictures.imgDragonForward2.Picture
    imgRedDragon.Width = imgRedDragon.Width + 50
    imgRedDragon.Height = imgRedDragon.Height + 50

Case Is = 6
    imgRedDragon.Picture = frmPictures.imgDragonForward3.Picture
    imgRedDragon.Width = imgRedDragon.Width + 50
    imgRedDragon.Height = imgRedDragon.Height + 50
    
Case Is = 7
    imgRedDragon.Picture = frmPictures.imgDragonForward1.Picture
    imgRedDragon.Width = imgRedDragon.Width + 75
    imgRedDragon.Height = imgRedDragon.Height + 75

Case Is = 8
    imgRedDragon.Picture = frmPictures.imgDragonForward2.Picture
    imgRedDragon.Width = imgRedDragon.Width + 75
    imgRedDragon.Height = imgRedDragon.Height + 75

Case Is = 9
    imgRedDragon.Picture = frmPictures.imgDragonForward3.Picture
    imgRedDragon.Width = imgRedDragon.Width + 75
    imgRedDragon.Height = imgRedDragon.Height + 75
    
Case Is = 10
    imgRedDragon.Picture = frmPictures.imgDragonForward1.Picture
    imgRedDragon.Width = imgRedDragon.Width + 100
    imgRedDragon.Height = imgRedDragon.Height + 100
    
Case Is = 11
    imgRedDragon.Picture = frmPictures.imgDragonForward2.Picture
    imgRedDragon.Width = imgRedDragon.Width + 100
    imgRedDragon.Height = imgRedDragon.Height + 100

Case Is = 12
    imgRedDragon.Picture = frmPictures.imgDragonForward3.Picture
    imgRedDragon.Width = imgRedDragon.Width + 100
    imgRedDragon.Height = imgRedDragon.Height + 100
    
Case Is = 13
    imgRedDragon.Picture = frmPictures.imgDragonForward1.Picture
    imgRedDragon.Width = imgRedDragon.Width + 125
    imgRedDragon.Height = imgRedDragon.Height + 125

Case Is = 14
    imgRedDragon.Picture = frmPictures.imgDragonForward2.Picture
    imgRedDragon.Width = imgRedDragon.Width + 125
    imgRedDragon.Height = imgRedDragon.Height + 125

Case Is = 15
    imgRedDragon.Picture = frmPictures.imgDragonForward3.Picture
    imgRedDragon.Width = imgRedDragon.Width + 125
    imgRedDragon.Height = imgRedDragon.Height + 125
    
Case Is = 16
    imgRedDragon.Picture = frmPictures.imgDragonForward1.Picture
    imgRedDragon.Width = imgRedDragon.Width + 150
    imgRedDragon.Height = imgRedDragon.Height + 150

Case Is = 17
    imgRedDragon.Picture = frmPictures.imgDragonForward2.Picture
    imgRedDragon.Width = imgRedDragon.Width + 150
    imgRedDragon.Height = imgRedDragon.Height + 150
    
Case Is = 18
    imgRedDragon.Picture = frmPictures.imgDragonForward3.Picture
    imgRedDragon.Width = imgRedDragon.Width + 150
    imgRedDragon.Height = imgRedDragon.Height + 150

Case Is = 19
    imgRedDragon.Picture = frmPictures.imgDragonRight1.Picture
    imgRedDragon.Width = imgRedDragon.Width + 150
    imgRedDragon.Left = imgRedDragon.Left + 150

Case Is = 20
    imgRedDragon.Picture = frmPictures.imgDragonRight2.Picture
    imgRedDragon.Left = imgRedDragon.Left + 150
    
Case Is = 21
    imgRedDragon.Picture = frmPictures.imgDragonRight3.Picture
    imgRedDragon.Left = imgRedDragon.Left + 150
    
Case Is = 22
    imgRedDragon.Picture = frmPictures.imgDragonRight1.Picture
    imgRedDragon.Left = imgRedDragon.Left + 150
    
Case Is = 23
    imgRedDragon.Picture = frmPictures.imgDragonRight2.Picture
    imgRedDragon.Left = imgRedDragon.Left + 150

Case Is = 24
    imgRedDragon.Picture = frmPictures.imgDragonRight3.Picture
    imgRedDragon.Left = imgRedDragon.Left + 150
    
Case Is = 25
    imgRedDragon.Picture = frmPictures.imgDragonRight1.Picture
    imgRedDragon.Left = imgRedDragon.Left + 150
    
Case Is = 26
    imgRedDragon.Picture = frmPictures.imgDragonRight2.Picture
    imgRedDragon.Left = imgRedDragon.Left + 150
    
Case Is = 27
    imgRedDragon.Picture = frmPictures.imgDragonRight3.Picture
    imgRedDragon.Left = imgRedDragon.Left + 150
    timerDragonBattle.Enabled = False
    timerSecondaryBattle.Enabled = True
    
End Select
End Sub

Private Sub timerSecondaryBattle_Timer()
secondaryDragonCount = secondaryDragonCount + 1

Select Case secondaryDragonCount
Case Is = 1
    imgRedDragon.Picture = frmPictures.imgDragonRight1.Picture
    imgRedDragon.Top = imgRedDragon.Top - 20
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft1.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 200
    
Case Is = 2
    imgRedDragon.Picture = frmPictures.imgDragonRight2.Picture
    imgRedDragon.Top = imgRedDragon.Top + 20
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft2.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 200
    
Case Is = 3
    imgRedDragon.Picture = frmPictures.imgDragonRight3.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft4.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 200
    
Case Is = 4
    imgRedDragon.Picture = frmPictures.imgDragonRight1.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft4.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 250
    
Case Is = 5
    imgRedDragon.Picture = frmPictures.imgDragonRight2.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft4.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 300
   
Case Is = 6
    imgRedDragon.Picture = frmPictures.imgDragonRight3.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft4.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 350
    
Case Is = 7
    imgRedDragon.Picture = frmPictures.imgDragonRight1.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft4.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 400

Case Is = 7
    imgRedDragon.Picture = frmPictures.imgDragonRight2.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft4.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 400
    
Case Is = 8
    imgRedDragon.Picture = frmPictures.imgDragonRight3.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft4.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 400

Case Is = 9
    imgRedDragon.Picture = frmPictures.imgDragonRight1.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft4.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 400
    imgBlueDragon.Top = imgBlueDragon.Top + 150
    
Case Is = 10
    imgRedDragon.Picture = frmPictures.imgDragonRight2.Picture
    
    imgFireBreath.Picture = frmPictures.imgFireBreath1.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft4.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 400
    imgBlueDragon.Top = imgBlueDragon.Top + 250
    
Case Is = 11
    imgRedDragon.Picture = frmPictures.imgDragonRight3.Picture
    
    imgFireBreath.Picture = frmPictures.imgFireBreath2.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft4.Picture
    imgBlueDragon.Left = imgBlueDragon.Left - 400
    imgBlueDragon.Top = imgBlueDragon.Top + 350

Case Is = 12
    imgRedDragon.Picture = frmPictures.imgDragonRight1.Picture
    
    imgFireBreath.Picture = frmPictures.imgFireBreath3.Picture
    imgFireBreath.Visible = False
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft1.Picture
    
Case Is = 13
    imgRedDragon.Picture = frmPictures.imgDragonRight2.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft2.Picture
    
Case Is = 14
    imgRedDragon.Picture = frmPictures.imgDragonRight3.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft3.Picture
    imgBlueDragon.Top = imgBlueDragon.Top - 150
    
Case Is = 15
    imgRedDragon.Picture = frmPictures.imgDragonRight3.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft3.Picture
    imgBlueDragon.Top = imgBlueDragon.Top - 150
    
Case Is = 16
    imgRedDragon.Picture = frmPictures.imgDragonRight1.Picture
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft3.Picture
    imgBlueDragon.Top = imgBlueDragon.Top - 150
    
Case Is = 17
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft1.Picture
    
Case Is = 18
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft1.Picture
    imgBlueDragon.Top = imgBlueDragon.Top - 450
    imgLightning.Picture = frmPictures.imgLightning1.Picture
    
Case Is = 19
    imgLightning.Picture = frmPictures.imgLightning2.Picture
    
Case Is = 20
    imgLightning.Picture = frmPictures.imgLightning3.Picture
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150
    
Case Is = 21
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150
    
    imgLightning.Visible = False
    
    imgBlueDragon.Picture = frmPictures.imgBlueDragonLeft3.Picture
Case Is = 22
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150
    
Case Is = 23
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150

Case Is = 24
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150

Case Is = 25
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150

Case Is = 26
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150
    
Case Is = 27
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150
    
Case Is = 28
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150
    
Case Is = 29
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150
    
Case Is = 30
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150
    
Case Is = 31
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150
    
Case Is = 32
    imgRedDragon.Height = imgRedDragon.Height - 150
    imgRedDragon.Width = imgRedDragon.Width - 150
    
Case Is = 33
    imgRedDragon.Visible = False
    frmTutorialLevel1.Show
    timerSecondaryBattle.Enabled = False
    Unload Me
    
    
End Select
End Sub
