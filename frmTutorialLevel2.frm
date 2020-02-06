VERSION 5.00
Begin VB.Form frmTutorialLevel2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tutorial Level 2"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTutorialLevel2.frx":0000
   ScaleHeight     =   7440
   ScaleWidth      =   15000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerRockCollision 
      Interval        =   1
      Left            =   14400
      Top             =   6960
   End
   Begin VB.Timer timerEnemy3Movement 
      Interval        =   10
      Left            =   3240
      Top             =   4920
   End
   Begin VB.Timer timerEnemy2Movement 
      Interval        =   10
      Left            =   7680
      Top             =   4800
   End
   Begin VB.Timer timerEnemy3Death 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3240
      Top             =   5640
   End
   Begin VB.Timer timerEnemy3Animation 
      Interval        =   250
      Left            =   3240
      Top             =   5280
   End
   Begin VB.Timer timerEnemy2Animation 
      Interval        =   250
      Left            =   7680
      Top             =   5160
   End
   Begin VB.Timer timerEnemy2Death 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   7680
      Top             =   5520
   End
   Begin VB.Timer timerFireBreath 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   13680
      Top             =   3000
   End
   Begin VB.Timer timerCollision 
      Interval        =   1
      Left            =   7080
      Top             =   7080
   End
   Begin VB.Timer timerLeftMovement 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timerRightMovement 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   14640
      Top             =   0
   End
   Begin VB.Timer timerEnemy1Movement 
      Interval        =   10
      Left            =   4440
      Top             =   1560
   End
   Begin VB.Timer timerEnemy1Animation 
      Interval        =   250
      Left            =   4440
      Top             =   840
   End
   Begin VB.Timer timerFire 
      Interval        =   250
      Left            =   13680
      Top             =   2400
   End
   Begin VB.Image imgRockUpDown 
      Height          =   975
      Index           =   8
      Left            =   9720
      Picture         =   "frmTutorialLevel2.frx":197DC
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image imgRockUpDown 
      Height          =   975
      Index           =   6
      Left            =   8400
      Picture         =   "frmTutorialLevel2.frx":1A97F
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image imgRockUpDown 
      Height          =   975
      Index           =   7
      Left            =   -840
      Picture         =   "frmTutorialLevel2.frx":1BB22
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image imgRockUpLeft 
      Height          =   975
      Left            =   11040
      Picture         =   "frmTutorialLevel2.frx":1CCC5
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image imgRockUpDown 
      Height          =   975
      Index           =   5
      Left            =   7080
      Picture         =   "frmTutorialLevel2.frx":1DE68
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image imgRockUpDown 
      Height          =   975
      Index           =   4
      Left            =   5760
      Picture         =   "frmTutorialLevel2.frx":1F00B
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image imgRockUpDown 
      Height          =   975
      Index           =   3
      Left            =   4440
      Picture         =   "frmTutorialLevel2.frx":201AE
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image imgRockUpDown 
      Height          =   975
      Index           =   2
      Left            =   3120
      Picture         =   "frmTutorialLevel2.frx":21351
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image imgRockUpDown 
      Height          =   975
      Index           =   1
      Left            =   1800
      Picture         =   "frmTutorialLevel2.frx":224F4
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image imgRockDownRight 
      Height          =   1095
      Left            =   8520
      Picture         =   "frmTutorialLevel2.frx":23697
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgFireBreath 
      Height          =   615
      Left            =   600
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1335
   End
   Begin VB.Image imgEnemy3 
      Height          =   975
      Left            =   3120
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgEnemy2 
      Height          =   975
      Left            =   7560
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgRockRightFurther 
      Height          =   1215
      Index           =   5
      Left            =   14760
      Picture         =   "frmTutorialLevel2.frx":2483A
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Image imgRockRightFurther 
      Height          =   1215
      Index           =   4
      Left            =   14760
      Picture         =   "frmTutorialLevel2.frx":259DD
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgRockRightFurther 
      Height          =   1215
      Index           =   3
      Left            =   14760
      Picture         =   "frmTutorialLevel2.frx":26B80
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Image imgRockRightFurther 
      Height          =   1215
      Index           =   2
      Left            =   14760
      Picture         =   "frmTutorialLevel2.frx":27D23
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image imgRockRightFurther 
      Height          =   1215
      Index           =   1
      Left            =   14760
      Picture         =   "frmTutorialLevel2.frx":28EC6
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Image imgRockRightFurther 
      Height          =   1215
      Index           =   0
      Left            =   14760
      Picture         =   "frmTutorialLevel2.frx":2A069
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image imgCharacter 
      Height          =   1215
      Left            =   480
      Picture         =   "frmTutorialLevel2.frx":2B20C
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Image imgEnemy1 
      Height          =   1140
      Left            =   4320
      Top             =   840
      Width           =   750
   End
   Begin VB.Image imgFirePowerUp 
      Height          =   1095
      Left            =   13560
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   615
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   12
      Left            =   -480
      Picture         =   "frmTutorialLevel2.frx":2BFA6
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   11
      Left            =   720
      Picture         =   "frmTutorialLevel2.frx":2D149
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   10
      Left            =   1920
      Picture         =   "frmTutorialLevel2.frx":2E2EC
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   9
      Left            =   3120
      Picture         =   "frmTutorialLevel2.frx":2F48F
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   8
      Left            =   4320
      Picture         =   "frmTutorialLevel2.frx":30632
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   7
      Left            =   5520
      Picture         =   "frmTutorialLevel2.frx":317D5
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   6
      Left            =   6600
      Picture         =   "frmTutorialLevel2.frx":32978
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   5
      Left            =   7800
      Picture         =   "frmTutorialLevel2.frx":33B1B
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   4
      Left            =   9000
      Picture         =   "frmTutorialLevel2.frx":34CBE
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   3
      Left            =   10200
      Picture         =   "frmTutorialLevel2.frx":35E61
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   2
      Left            =   11400
      Picture         =   "frmTutorialLevel2.frx":37004
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   1
      Left            =   12600
      Picture         =   "frmTutorialLevel2.frx":381A7
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1215
      Index           =   0
      Left            =   13800
      Picture         =   "frmTutorialLevel2.frx":3934A
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgRockDownLeft 
      Height          =   1095
      Left            =   10920
      Picture         =   "frmTutorialLevel2.frx":3A4ED
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1215
      Index           =   9
      Left            =   14040
      Picture         =   "frmTutorialLevel2.frx":3B690
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1215
      Index           =   8
      Left            =   12840
      Picture         =   "frmTutorialLevel2.frx":3C833
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1215
      Index           =   7
      Left            =   11760
      Picture         =   "frmTutorialLevel2.frx":3D9D6
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1215
      Index           =   6
      Left            =   10560
      Picture         =   "frmTutorialLevel2.frx":3EB79
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   1215
   End
   Begin VB.Image imgRockDownFurther 
      Height          =   1095
      Left            =   9720
      Picture         =   "frmTutorialLevel2.frx":3FD1C
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1215
      Index           =   5
      Left            =   9360
      Picture         =   "frmTutorialLevel2.frx":40EBF
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1215
      Index           =   4
      Left            =   8160
      Picture         =   "frmTutorialLevel2.frx":42062
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1215
      Index           =   3
      Left            =   6960
      Picture         =   "frmTutorialLevel2.frx":43205
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1215
      Index           =   2
      Left            =   5760
      Picture         =   "frmTutorialLevel2.frx":443A8
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1215
      Index           =   1
      Left            =   4560
      Picture         =   "frmTutorialLevel2.frx":4554B
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   1215
   End
   Begin VB.Image imgRockUpDown 
      Height          =   975
      Index           =   0
      Left            =   480
      Picture         =   "frmTutorialLevel2.frx":466EE
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image imgRockUp 
      Height          =   1215
      Index           =   0
      Left            =   3360
      Picture         =   "frmTutorialLevel2.frx":47891
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   1215
   End
   Begin VB.Image imgRockRightUp 
      Height          =   1215
      Left            =   2160
      Picture         =   "frmTutorialLevel2.frx":48A34
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   1215
   End
   Begin VB.Image imgRockLeft 
      Height          =   1380
      Index           =   3
      Left            =   -1200
      Picture         =   "frmTutorialLevel2.frx":49BD7
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Image imgRockLeft 
      Height          =   1380
      Index           =   1
      Left            =   -1200
      Picture         =   "frmTutorialLevel2.frx":4AD7A
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Image imgRockLeft 
      Height          =   1380
      Index           =   0
      Left            =   -1200
      Picture         =   "frmTutorialLevel2.frx":4BF1D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmTutorialLevel2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fireCount As Integer
Dim enemy1Count As Integer
Dim enemy2Count As Integer
Dim enemy3Count As Integer
Dim fadeCount As Integer

Dim enemy1Speed As Integer
Dim enemy2Speed As Integer
Dim enemy3Speed As Integer
Dim characterBottom As Integer
Dim fireBreathCount As Integer
Dim enemy2DeathCount As Integer
Dim enemy3DeathCount As Integer


Dim enemy1Death As Boolean
Dim enemy2Death As Boolean
Dim enemy3Death As Boolean

Public firePowerup As Boolean

Dim playerUpSpeed As Integer
Dim playerDownSpeed As Integer
Dim playerLeftSpeed As Integer
Dim playerRightSpeed As Integer '





Private Sub imgRockRight_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 119 Then
    imgCharacter.Top = imgCharacter.Top - playerUpSpeed

End If

If KeyAscii = 97 Then
    imgCharacter.Left = imgCharacter.Left - playerLeftSpeed
    imgFireBreath.Left = imgFireBreath.Left - playerLeftSpeed
    timerLeftMovement.Enabled = True
    timerRightMovement.Enabled = False

End If

If KeyAscii = 115 Then
    imgCharacter.Top = imgCharacter.Top + playerDownSpeed

End If

If KeyAscii = 100 Then
    imgCharacter.Left = imgCharacter.Left + playerRightSpeed
    imgFireBreath.Left = imgFireBreath.Left + playerRightSpeed
    timerRightMovement.Enabled = True
    timerLeftMovement.Enabled = False
End If
End Sub

Private Sub Form_Load()
fireCount = 0
enemy1Count = 0
enemy2Count = 0
enemy3Count = 0
enemy1Speed = 5
enemy2Speed = 10
enemy3Speed = 10
fireBreathCount = 0
firePowerup = False

enemy2DeathCount = 0
enemy3DeathCount = 0

enemy1Death = False
enemy2Death = False
enemy3Death = False

playerUpSpeed = 50
playerDownSpeed = 50
playerLeftSpeed = 75
playerRightSpeed = 75

End Sub

Private Sub timerEnemy_Timer()

        
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If firePowerup = True Then
    timerFireBreath.Enabled = True
    imgFireBreath.Visible = True
End If

    
End Sub

Private Sub timerCollision_Timer()

'Tracks if the Character collects the fire powerup
If imgCharacter.Top + imgCharacter.Height >= imgFirePowerUp.Top And imgCharacter.Top <= imgFirePowerUp.Top + imgFirePowerUp.Height And imgFirePowerUp.Left >= imgCharacter.Left And imgFirePowerUp.Left + imgFirePowerUp.Width <= imgCharacter.Left + imgCharacter.Width And imgFirePowerUp.Visible = True Then
    firePowerup = True
    imgFirePowerUp.Visible = False
End If

'Tracks if the fire breath hits Enemy 2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If imgFireBreath.Left <= imgEnemy2.Left + imgEnemy2.Height And imgFireBreath.Top >= imgEnemy2.Top And imgFireBreath.Top + imgFireBreath.Height <= imgEnemy2.Top + imgEnemy2.Height And imgFireBreath.Visible = True And enemy2Death = False Then
    timerEnemy2Death.Enabled = True
    timerEnemy2Animation.Enabled = False
    enemy2Death = True
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Tracks if the fire breath hits Enemy 3
If imgFireBreath.Left <= imgEnemy3.Left + imgEnemy3.Height And imgFireBreath.Top >= imgEnemy3.Top And imgFireBreath.Top + imgFireBreath.Height <= imgEnemy3.Top + imgEnemy3.Height And imgFireBreath.Visible = True And enemy3Death = False Then
    timerEnemy3Death.Enabled = True
    timerEnemy3Animation.Enabled = False
    enemy3Death = True
End If

'Tracks if the character hits Enemy 1
If imgCharacter.Left + imgCharacter.Width >= imgEnemy1.Left And imgCharacter.Left + imgCharacter.Width <= imgEnemy1.Left + imgEnemy1.Width And imgCharacter.Top + imgCharacter.Height >= imgEnemy1.Top And imgCharacter.Top <= imgEnemy1.Top + imgEnemy1.Height And imgEnemy1.Visible = True Then
    fadeCount = fadeCount + 1
    frmTutorialLevel2.Hide
    
    If fadeCount >= 50 Then
        Load Me
        frmTutorialLevel2.Show
        fadeCount = 0
        imgCharacter.Left = 120
        imgCharacter.Top = -815
        imgEnemy1.Visible = True
        imgEnemy2.Visible = True
        imgEnemy3.Visible = True
    End If
End If

'Tracks if the character hits Enemy 2
If imgCharacter.Left <= imgEnemy2.Left And imgCharacter.Left + imgCharacter.Width <= imgEnemy2.Left + imgEnemy2.Width And imgCharacter.Top + imgCharacter.Height >= imgEnemy2.Top And imgCharacter.Top <= imgEnemy2.Top + imgEnemy2.Height And imgEnemy2.Visible = True Then
    fadeCount = fadeCount + 1
    frmTutorialLevel2.Hide
    
    If fadeCount >= 50 Then
        Load Me
        frmTutorialLevel2.Show
        fadeCount = 0
        imgCharacter.Left = 120
        imgCharacter.Top = -815
        imgEnemy2.Visible = True
        imgEnemy2.Visible = True
        imgEnemy3.Visible = True
    End If
End If


'Tracks if the character hits Enemy 3
If imgCharacter.Left <= imgEnemy3.Left And imgCharacter.Left + imgCharacter.Width <= imgEnemy3.Left + imgEnemy3.Width And imgCharacter.Top + imgCharacter.Height >= imgEnemy3.Top And imgCharacter.Top <= imgEnemy3.Top + imgEnemy3.Height And imgEnemy3.Visible = True Then
    fadeCount = fadeCount + 1
    frmTutorialLevel2.Hide
    
    If fadeCount >= 50 Then
        Load Me
        frmTutorialLevel2.Show
        fadeCount = 0
        imgCharacter.Left = 120
        imgCharacter.Top = -815
        imgEnemy3.Visible = True
        imgEnemy3.Visible = True
        imgEnemy3.Visible = True
    End If
End If

If imgCharacter.Left + imgCharacter.Width <= 0 Then
    Unload Me
    frmRealLevel1.Show
End If

End Sub

Private Sub timerEnemy1Animation_Timer()
enemy1Count = enemy1Count + 1

Select Case enemy1Count
    Case Is = 1
        imgEnemy1.Picture = frmPictures.imgGhostLeft1.Picture
        
    Case Is = 2
        imgEnemy1.Picture = frmPictures.imgGhostLeft2.Picture
        
    Case Is = 3
        imgEnemy1.Picture = frmPictures.imgGhostLeft3.Picture
        
    Case Is = 4
        imgEnemy1.Picture = frmPictures.imgGhostLeft4.Picture
        enemy1Count = 1
End Select
End Sub

Private Sub timerEnemy1Movement_Timer()
imgEnemy1.Top = imgEnemy1.Top - enemy1Speed

For upRock = 0 To 11
    If imgEnemy1.Top <= imgRockUp(upRock).Top + imgRockUp(upRock).Height Then
        enemy1Speed = -enemy1Speed
    End If
Next upRock

For upDownRock = 0 To 6
    If imgEnemy1.Top + imgEnemy1.Height >= imgRockUpDown(upDownRock).Top Then
        enemy1Speed = -enemy1Speed
    End If
Next upDownRock
End Sub

Private Sub timerEnemy2Animation_Timer()
enemy2Count = enemy2Count + 1

Select Case enemy2Count
    Case Is = 1
        imgEnemy2.Picture = frmPictures.imgGhostRight1.Picture
    
    Case Is = 2
        imgEnemy2.Picture = frmPictures.imgGhostRight2.Picture
        
    Case Is = 3
        imgEnemy2.Picture = frmPictures.imgGhostRight3.Picture
    
    
    Case Is = 4
        imgEnemy2.Picture = frmPictures.imgGhostRight4.Picture
        enemy2Count = 1
End Select
End Sub

Private Sub timerEnemy2Death_Timer()
enemy2DeathCount = enemy2DeathCount + 1

Select Case enemy2DeathCount
    Case Is = 1
        imgEnemy2.Picture = frmPictures.imgGhostRightDeath1.Picture
        
    Case Is = 2
        imgEnemy2.Picture = frmPictures.imgGhostRightDeath2.Picture
        
    Case Is = 3
        imgEnemy2.Picture = frmPictures.imgGhostRightDeath3.Picture
        
    Case Is = 4
        imgEnemy2.Picture = frmPictures.imgGhostRightDeath4.Picture
    
    Case Is = 5
        imgEnemy2.Picture = frmPictures.imgGhostRightDeath5.Picture
        enemy2DeathCount = 1
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       imgEnemy2.Visible = False
       timerEnemy2Death.Enabled = False
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Select
End Sub

Private Sub timerEnemyAnimation_Timer()

End Sub

Private Sub timerEnemy2Movement_Timer()
imgEnemy2.Top = imgEnemy2.Top - enemy2Speed


For upDownRock = 0 To 6
    If imgEnemy2.Top <= imgRockUpDown(upDownRock).Top + imgRockUpDown(upDownRock).Height Then
        enemy2Speed = -enemy2Speed
    End If
Next upDownRock

For downRock = 0 To 12
    If imgEnemy2.Top + imgEnemy2.Height >= imgRockDown(downRock).Top Then
        enemy2Speed = -enemy2Speed
    End If
Next downRock

End Sub

Private Sub timerEnemy3Animation_Timer()
enemy3Count = enemy3Count + 1

Select Case enemy3Count
    Case Is = 1
        imgEnemy3.Picture = frmPictures.imgGhostRight1.Picture
        
    Case Is = 2
        imgEnemy3.Picture = frmPictures.imgGhostRight2.Picture
        
    Case Is = 3
        imgEnemy3.Picture = frmPictures.imgGhostRight3.Picture
        
    Case Is = 4
        imgEnemy3.Picture = frmPictures.imgGhostRight3.Picture
        enemy3Count = 1
End Select
End Sub

Private Sub timerEnemy3Death_Timer()
enemy3DeathCount = enemy3DeathCount + 1

Select Case enemy3DeathCount
    Case Is = 1
        imgEnemy3.Picture = frmPictures.imgGhostRightDeath1.Picture
        
    Case Is = 2
        imgEnemy3.Picture = frmPictures.imgGhostRightDeath2.Picture
        
    Case Is = 3
        imgEnemy3.Picture = frmPictures.imgGhostRightDeath3.Picture
        
    Case Is = 4
        imgEnemy3.Picture = frmPictures.imgGhostRightDeath4.Picture
    
    Case Is = 5
        imgEnemy3.Picture = frmPictures.imgGhostRightDeath5.Picture
        enemy3DeathCount = 1
        imgEnemy3.Visible = False
End Select
End Sub

Private Sub timerEnemyMovement_Timer()



'------------------------------------------------------------------------Enemy 2 & 3 Movement---------------------------------------------------------------------------

End Sub

Private Sub timerEnemy3Movement_Timer()
imgEnemy3.Top = imgEnemy3.Top + enemy3Speed


For upDownRock = 0 To 6
    If imgEnemy3.Top <= imgRockUpDown(upDownRock).Top + imgRockUpDown(upDownRock).Height Then
        enemy3Speed = -enemy3Speed
    End If
Next upDownRock

For downRock = 0 To 12
    If imgEnemy3.Top + imgEnemy3.Height >= imgRockDown(downRock).Top Then
        enemy3Speed = -enemy3Speed
    End If
Next downRock
End Sub

Private Sub timerFire_Timer()
fireCount = fireCount + 1

Select Case fireCount
    Case Is = 1
        imgFirePowerUp.Picture = frmPictures.imgFire1.Picture
        
    Case Is = 2
        imgFirePowerUp.Picture = frmPictures.imgFire2.Picture
        
    Case Is = 3
        imgFirePowerUp.Picture = frmPictures.imgFire3.Picture
        
    Case Is = 4
        imgFirePowerUp.Picture = frmPictures.imgFire4.Picture
        fireCount = 1
End Select
End Sub

Private Sub timerFireBreath_Timer()
fireBreathCount = fireBreathCount + 1

If timerRightMovement.Enabled = True Then
    imgFireBreath.Left = imgCharacter.Left + imgCharacter.Width
    imgFireBreath.Top = ((imgCharacter.Top + (imgCharacter.Top + imgCharacter.Height)) / 2) - (imgFireBreath.Height / 2)
    
    

ElseIf timerLeftMovement.Enabled = True Then
    imgFireBreath.Left = imgCharacter.Left - imgFireBreath.Width
    imgFireBreath.Top = (imgCharacter.Top + (imgCharacter.Top + imgCharacter.Height)) / 2 - (imgFireBreath.Height / 2)
    
Else
    imgFireBreath.Left = imgCharacter.Left - imgFireBreath.Width
    imgFireBreath.Top = (imgCharacter.Top + (imgCharacter.Top + imgCharacter.Height)) / 2 - (imgFireBreath.Height / 2)
End If


Select Case fireBreathCount
    Case Is = 1
        imgFireBreath.Picture = frmPictures.imgFireBreath1.Picture
        
    Case Is = 2
        imgFireBreath.Picture = frmPictures.imgFireBreath2.Picture
        
    Case Is = 3
        imgFireBreath.Picture = frmPictures.imgFireBreath3.Picture
    
    Case Is = 4
        imgFireBreath.Picture = frmPictures.imgFireBreath4.Picture
        
    Case Is = 5
        imgFireBreath.Visible = False
        fireBreathCount = 1
End Select
End Sub

Private Sub timerLeftMovement_Timer()
'Increments the animation for the sprites facing left
frmTutorialLevel1.leftSprites = frmTutorialLevel1.leftSprites + 1

'Loops through the sprites
Select Case frmTutorialLevel1.leftSprites
    Case Is = 1
        imgCharacter.Picture = frmPictures.imgDragonLeft1.Picture
        
    Case Is = 2
        imgCharacter.Picture = frmPictures.imgDragonLeft2.Picture
        
    Case Is = 3
        imgCharacter.Picture = frmPictures.imgDragonLeft3.Picture
        frmTutorialLevel1.leftSprites = 0
End Select
End Sub

Private Sub timerRightMovement_Timer()
'Increments the animation for the sprites facing right
frmTutorialLevel1.rightSprites = frmTutorialLevel1.rightSprites + 1

'Loops through the sprites
Select Case frmTutorialLevel1.rightSprites
    Case Is = 1
        imgCharacter.Picture = frmPictures.imgDragonRight1.Picture
    
    Case Is = 2
        imgCharacter.Picture = frmPictures.imgDragonRight2.Picture
        
    Case Is = 3
        imgCharacter.Picture = frmPictures.imgDragonRight3.Picture
        frmTutorialLevel1.rightSprites = 0
End Select
End Sub

Private Sub timerRockCollision_Timer()
'REQUIRED so that the character can move after moving away from collision bounds
playerUpSpeed = 50
playerDownSpeed = 50
playerLeftSpeed = 75
playerRightSpeed = 75

'Checks the collision for the rocks that the character can collide with while moving upwards
For upRock = 0 To 9
    If imgCharacter.Top <= imgRockUp(upRock).Top + imgRockUp(upRock).Height And imgCharacter.Left + imgCharacter.Width >= imgRockUp(upRock).Left And imgCharacter.Left + imgCharacter.Width <= imgRockUp(upRock).Left + imgRockUp(upRock).Width Then
        playerUpSpeed = 0
        playerDownSpeed = 50
    End If
Next upRock

'Checks the collision for the rocks that the character can collide with while moving downwards
For downRock = 0 To 11
    If imgCharacter.Top + imgCharacter.Height >= imgRockDown(downRock).Top And imgCharacter.Left + imgCharacter.Width >= imgRockDown(downRock).Left And imgCharacter.Left + imgCharacter.Width <= imgRockDown(downRock).Left + imgRockDown(downRock).Width Then
        playerDownSpeed = 0
        playerUpSpeed = 50
        timerLeftMovement.Enabled = False
        timerRightMovement.Enabled = False
    End If
Next downRock


If imgCharacter.Top + imgCharacter.Height >= imgRockDownFurther.Top And imgCharacter.Top <= imgRockDownFurther.Top + imgRockDownFurther.Height And imgCharacter.Left + imgCharacter.Width >= imgRockDownFurther.Left And imgCharacter.Left + imgCharacter.Width <= imgRockDownFurther.Left + imgRockDownFurther.Width Then
    playerDownSpeed = 0
    playerUpSpeed = 50
    timerLeftMovement.Enabled = False
    timerRightMovement.Enabled = False
End If

'Checks the collision for the rocks that the character can collide with while moving left
For leftRock = 0 To 1
    If imgCharacter.Left <= imgRockLeft(leftRock).Left + imgRockLeft(leftRock).Width And imgCharacter.Top <= imgRockLeft(leftRock).Top + imgRockLeft(leftRock).Height Then
        playerLeftSpeed = 0
        playerRightSpeed = 75
    End If
Next leftRock

'Checks the collision for the rocks that the character can collide with while moving right
For rightRockfurther = 0 To 4
    If imgCharacter.Left + imgCharacter.Width >= imgRockRightFurther(rightRockfurther).Left And imgCharacter.Top <= imgRockRightFurther(rightRockfurther).Top + imgRockRightFurther(rightRockfurther).Height And imgCharacter.Top + imgCharacter.Height >= imgRockRightFurther(rightRockfurther).Top Then
        playerRightSpeed = 0
        playerLeftSpeed = 75
    End If
Next rightRockfurther

'Checks the collision for the rocks that the character can collide with either moving down or right
If imgCharacter.Top + imgCharacter.Height >= imgRockDownRight.Top And imgCharacter.Top <= imgRockDownRight.Top + imgRockDownRight.Height And imgCharacter.Left >= imgRockDownRight.Left And imgCharacter.Left <= imgRockDownRight.Left + imgRockDownRight.Width Then
    playerDownSpeed = 0
    playerUpSpeed = 50
End If

If imgCharacter.Left + imgCharacter.Width >= imgRockDownRight.Left And imgCharacter.Left + imgCharacter.Width <= imgRockDownRight.Left + imgRockDownRight.Width And imgCharacter.Top + imgCharacter.Height <= imgRockDownRight.Top + imgRockDownRight.Height And imgCharacter.Top + imgCharacter.Height >= imgRockDownRight.Top Then
    playerRightSpeed = 0
    playerLeftSpeed = 75
End If

'Checks the collision for the rocks that the character can collide with either going up or down
For upDownRock = 0 To 8
    If imgCharacter.Top <= imgRockUpDown(upDownRock).Top + imgRockUpDown(upDownRock).Height And imgCharacter.Top >= imgRockUpDown(upDownRock).Top And imgCharacter.Left <= imgRockUpDown(upDownRock).Left Then
        playerUpSpeed = 0
        playerDownSpeed = 50
    End If
    
    If imgCharacter.Top + imgCharacter.Height >= imgRockUpDown(upDownRock).Top And imgCharacter.Top <= imgRockUpDown(upDownRock).Top + imgRockUpDown(upDownRock).Height And imgCharacter.Left + imgCharacter.Width >= imgRockUpDown(upDownRock).Left And imgCharacter.Left + imgCharacter.Width <= imgRockUpDown(upDownRock).Left + imgRockUpDown(upDownRock).Width Then
        playerDownSpeed = 0
        playerUpSpeed = 50
        timerLeftMovement.Enabled = False
        timerRightMovement.Enabled = False
    End If
Next upDownRock

'Checks the collision for the rock in the left half that the character can collide with either moving upwards or moving left
If imgCharacter.Top <= imgRockUpLeft.Top + imgRockUpLeft.Height And imgCharacter.Top >= imgRockUpLeft.Top And imgCharacter.Left <= imgRockUpLeft.Left + imgRockUpLeft.Width Then
    playerUpSpeed = 0
    playerDownSpeed = 75
End If

If imgCharacter.Left <= imgRockUpLeft.Left + imgRockUpLeft.Width And imgCharacter.Top >= imgRockUpLeft.Top And imgCharacter.Top + imgCharacter.Height <= imgRockUpLeft.Top + imgRockUpLeft.Width Then
    playerLeftSpeed = 0
    playerRightSpeed = 75
End If

'Checks the collision for the rock that the character can collide with moving down or left
If imgCharacter.Top + imgCharacter.Height >= imgRockDownLeft.Top And imgCharacter.Top <= imgRockDownLeft.Top + imgRockDownLeft.Height And imgCharacter.Left >= imgRockDownLeft.Left And imgCharacter.Left <= imgRockDownLeft.Left + imgRockDownLeft.Width Then
    playerDownSpeed = 0
    playerUpSpeed = 50
End If

If imgCharacter.Left <= imgRockDownLeft.Left + imgRockDownLeft.Width And imgCharacter.Top >= imgRockDownLeft.Top - 100 And imgCharacter.Top <= imgRockDownLeft.Top + imgRockDownLeft.Height Then
        playerLeftSpeed = 0
        playerRightSpeed = 75
End If
End Sub
