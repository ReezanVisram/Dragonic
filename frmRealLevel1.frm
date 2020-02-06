VERSION 5.00
Begin VB.Form frmRealLevel1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Real Level 1"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRealLevel1.frx":0000
   ScaleHeight     =   7440
   ScaleWidth      =   15000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerEnemy1Animation 
      Interval        =   250
      Left            =   4560
      Top             =   2280
   End
   Begin VB.Timer timerEnemy1Movement 
      Interval        =   10
      Left            =   4560
      Top             =   1320
   End
   Begin VB.Timer timerCollision 
      Interval        =   1
      Left            =   7560
      Top             =   6960
   End
   Begin VB.Timer timerRightMovement 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   14640
      Top             =   0
   End
   Begin VB.Timer timerLeftMovement 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   0
      Top             =   0
   End
   Begin VB.Image imgEnemy1 
      Height          =   1335
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image imgRockLeftFurther 
      Height          =   1215
      Index           =   5
      Left            =   -720
      Picture         =   "frmRealLevel1.frx":197DC
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Image imgRockLeftFurther 
      Height          =   1215
      Index           =   4
      Left            =   -720
      Picture         =   "frmRealLevel1.frx":1A97F
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Image imgRockLeftFurther 
      Height          =   1215
      Index           =   3
      Left            =   -720
      Picture         =   "frmRealLevel1.frx":1BB22
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Image imgRockLeftFurther 
      Height          =   1215
      Index           =   2
      Left            =   -720
      Picture         =   "frmRealLevel1.frx":1CCC5
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image imgRockLeftFurther 
      Height          =   1215
      Index           =   1
      Left            =   -720
      Picture         =   "frmRealLevel1.frx":1DE68
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image imgRockLeftFurther 
      Height          =   1215
      Index           =   0
      Left            =   -720
      Picture         =   "frmRealLevel1.frx":1F00B
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image imgRockDownHigher 
      Height          =   1095
      Index           =   5
      Left            =   4200
      Picture         =   "frmRealLevel1.frx":201AE
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Image imgRockDownHigher 
      Height          =   1095
      Index           =   4
      Left            =   5400
      Picture         =   "frmRealLevel1.frx":21351
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Image imgRockDownHigher 
      Height          =   1095
      Index           =   3
      Left            =   6600
      Picture         =   "frmRealLevel1.frx":224F4
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Image imgRockDownHigher 
      Height          =   1095
      Index           =   2
      Left            =   7800
      Picture         =   "frmRealLevel1.frx":23697
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Image imgRockDownHigher 
      Height          =   1095
      Index           =   1
      Left            =   9000
      Picture         =   "frmRealLevel1.frx":2483A
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Image imgRockDownHigher 
      Height          =   1095
      Index           =   0
      Left            =   10200
      Picture         =   "frmRealLevel1.frx":259DD
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   12
      Left            =   7800
      Picture         =   "frmRealLevel1.frx":26B80
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgRockUpRight 
      Height          =   1215
      Left            =   14760
      Picture         =   "frmRealLevel1.frx":27A88
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Image imgRockDownLeft 
      Height          =   1095
      Left            =   11400
      Picture         =   "frmRealLevel1.frx":28C2B
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Image imgCharacter 
      Height          =   1215
      Left            =   12840
      Picture         =   "frmRealLevel1.frx":29DCE
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Image imgRockRight 
      Height          =   1215
      Index           =   2
      Left            =   14760
      Picture         =   "frmRealLevel1.frx":2AC5F
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image imgRockRight 
      Height          =   1215
      Index           =   1
      Left            =   14760
      Picture         =   "frmRealLevel1.frx":2BE02
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image imgRockRight 
      Height          =   1215
      Index           =   0
      Left            =   14760
      Picture         =   "frmRealLevel1.frx":2CFA5
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   12
      Left            =   14400
      Picture         =   "frmRealLevel1.frx":2E148
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   11
      Left            =   13200
      Picture         =   "frmRealLevel1.frx":2F2EB
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   10
      Left            =   12000
      Picture         =   "frmRealLevel1.frx":3048E
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   9
      Left            =   10800
      Picture         =   "frmRealLevel1.frx":31631
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   8
      Left            =   9600
      Picture         =   "frmRealLevel1.frx":327D4
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   7
      Left            =   8400
      Picture         =   "frmRealLevel1.frx":33977
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   6
      Left            =   7200
      Picture         =   "frmRealLevel1.frx":34B1A
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   5
      Left            =   6000
      Picture         =   "frmRealLevel1.frx":35CBD
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   4
      Left            =   4800
      Picture         =   "frmRealLevel1.frx":36E60
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   3
      Left            =   3600
      Picture         =   "frmRealLevel1.frx":38003
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   2
      Left            =   2400
      Picture         =   "frmRealLevel1.frx":391A6
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   1
      Left            =   1200
      Picture         =   "frmRealLevel1.frx":3A349
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   0
      Left            =   0
      Picture         =   "frmRealLevel1.frx":3B4EC
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   23
      Left            =   10200
      Picture         =   "frmRealLevel1.frx":3C68F
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   22
      Left            =   10200
      Picture         =   "frmRealLevel1.frx":3D597
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   21
      Left            =   10200
      Picture         =   "frmRealLevel1.frx":3E49F
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   19
      Left            =   9000
      Picture         =   "frmRealLevel1.frx":3F3A7
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   18
      Left            =   9000
      Picture         =   "frmRealLevel1.frx":402AF
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   17
      Left            =   9000
      Picture         =   "frmRealLevel1.frx":411B7
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   16
      Left            =   9000
      Picture         =   "frmRealLevel1.frx":420BF
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   15
      Left            =   7800
      Picture         =   "frmRealLevel1.frx":42FC7
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   14
      Left            =   7800
      Picture         =   "frmRealLevel1.frx":43ECF
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   13
      Left            =   7800
      Picture         =   "frmRealLevel1.frx":44DD7
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   11
      Left            =   6600
      Picture         =   "frmRealLevel1.frx":45CDF
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   10
      Left            =   6600
      Picture         =   "frmRealLevel1.frx":46BE7
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   9
      Left            =   6600
      Picture         =   "frmRealLevel1.frx":47AEF
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   8
      Left            =   6600
      Picture         =   "frmRealLevel1.frx":489F7
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   7
      Left            =   5400
      Picture         =   "frmRealLevel1.frx":498FF
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   6
      Left            =   5400
      Picture         =   "frmRealLevel1.frx":4A807
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   5
      Left            =   5400
      Picture         =   "frmRealLevel1.frx":4B70F
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   4
      Left            =   5400
      Picture         =   "frmRealLevel1.frx":4C617
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   3
      Left            =   4200
      Picture         =   "frmRealLevel1.frx":4D51F
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   2
      Left            =   4200
      Picture         =   "frmRealLevel1.frx":4E427
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   1
      Left            =   4200
      Picture         =   "frmRealLevel1.frx":4F32F
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   0
      Left            =   4200
      Picture         =   "frmRealLevel1.frx":50237
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgRockRightFurther 
      Height          =   1095
      Index           =   3
      Left            =   3000
      Picture         =   "frmRealLevel1.frx":5113F
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image imgRockRightFurther 
      Height          =   1095
      Index           =   2
      Left            =   3000
      Picture         =   "frmRealLevel1.frx":522E2
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image imgRockRightFurther 
      Height          =   1095
      Index           =   1
      Left            =   3000
      Picture         =   "frmRealLevel1.frx":53485
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgRockDownRight 
      Height          =   1095
      Left            =   3000
      Picture         =   "frmRealLevel1.frx":54628
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Image imgRockRightFurther 
      Height          =   1095
      Index           =   0
      Left            =   3000
      Picture         =   "frmRealLevel1.frx":557CB
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgRockLeft 
      Height          =   1095
      Index           =   2
      Left            =   11400
      Picture         =   "frmRealLevel1.frx":5696E
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgRockLeft 
      Height          =   1095
      Index           =   1
      Left            =   11400
      Picture         =   "frmRealLevel1.frx":57B11
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgRockLeft 
      Height          =   1095
      Index           =   0
      Left            =   11400
      Picture         =   "frmRealLevel1.frx":58CB4
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   2
      Left            =   11400
      Picture         =   "frmRealLevel1.frx":59E57
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   1
      Left            =   12600
      Picture         =   "frmRealLevel1.frx":5AFFA
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   0
      Left            =   13800
      Picture         =   "frmRealLevel1.frx":5C19D
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   20
      Left            =   10200
      Picture         =   "frmRealLevel1.frx":5D340
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmRealLevel1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Player Movement Variables (Public so that all levels can access these movement values regardless whether or not they're used in that particular level)
Dim playerUpSpeed As Integer
Dim playerDownSpeed As Integer
Dim playerLeftSpeed As Integer
Dim playerRightSpeed As Integer

'Enemy Movement Variables
Dim enemy1Speed As Integer

'Enemy Animation Variables
Dim enemy1Animation As Integer



'Dragon Initial Growth Variables
Dim dragonGrowth As Integer


'Exit animation variables
Dim exitCount As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 119 Then
    imgCharacter.Top = imgCharacter.Top - playerUpSpeed
End If

If KeyAscii = 97 Then
    imgCharacter.Left = imgCharacter.Left - playerLeftSpeed
    timerLeftMovement.Enabled = True
    timerRightMovement.Enabled = False

End If

If KeyAscii = 115 Then
    imgCharacter.Top = imgCharacter.Top + playerDownSpeed
End If

If KeyAscii = 100 Then
    imgCharacter.Left = imgCharacter.Left + playerRightSpeed
    timerRightMovement.Enabled = True
    timerLeftMovement.Enabled = False

End If
End Sub

Private Sub Form_Load()
'Assigns values to the Movement variables
playerUpSpeed = 50
playerDownSpeed = 50
playerLeftSpeed = 75
playerRightSpeed = 75
enemy1Speed = 10

End Sub

Private Sub timerCollision_Timer()
'REQUIRED so that the character can move after moving away from collision bounds
playerUpSpeed = 50
playerDownSpeed = 50
playerLeftSpeed = 75
playerRightSpeed = 75
'Checks if the player is colliding with the rocks that the user can collide with while moving upwards
For upRock = 0 To 11
    If imgCharacter.Top <= imgRockUp(upRock).Top + imgRockUp(upRock).Height And imgCharacter.Left + imgCharacter.Width >= imgRockUp(upRock).Left And imgCharacter.Left + imgCharacter.Width <= imgRockUp(upRock).Left + imgRockUp(upRock).Width Then
        playerUpSpeed = 0
        playerDownSpeed = 50
    End If
Next upRock

'Checks the collision for the rocks that the character can collide with while moving left
For leftRock = 0 To 2
    If imgCharacter.Left <= imgRockLeft(leftRock).Left + imgRockLeft(leftRock).Width And imgCharacter.Left >= imgRockLeft(leftRock).Left And imgCharacter.Top + imgCharacter.Height >= imgRockDownLeft.Top Then
        playerLeftSpeed = 0
        playerRightSpeed = 75
    End If
Next leftRock

For leftRockFurther = 0 To 5
    If imgCharacter.Left <= imgRockLeftFurther(leftRockFurther).Left + imgRockLeftFurther(leftRockFurther).Width And imgCharacter.Top Then
        playerLeftSpeed = 0
        playerRightSpeed = 75
    End If
Next leftRockFurther
    

'Checks the collision for the rocks that the character can collide with while moving down
For downRock = 0 To 1
    If imgCharacter.Top + imgCharacter.Height >= imgRockDown(downRock).Top And imgCharacter.Left >= imgRockDown(1).Left Then
        playerDownSpeed = 0
        playerUpSped = 50
    End If
Next downRock

'Checks collision for the rock that the user can collide with while moving either upwards or right
If imgCharacter.Top <= imgRockUpRight.Top + imgRockUpRight.Height And imgCharacter.Left + imgCharacter.Width >= imgRockUpRight.Left Then
    playerUpSpeed = 0
    playerDownSpeed = 50
End If

If imgCharacter.Left + imgCharacter.Width >= imgRockUpRight.Width And imgCharacter.Top >= imgRockUpRight.Top And imgCharacter.Top + imgCharacter.Height <= imgRockUpRight.Top + imgRockUpRight.Height Then
    playerRightSpeed = 0
    playerLeftSpeed = 75
End If

'Checks collision for the rocks that the user can collide with while moving right
For rightRock = 0 To 2
    If imgCharacter.Left + imgCharacter.Width >= imgRockRight(rightRock).Left And imgCharacter.Top >= imgRockRight(0).Top Then
        playerRightSpeed = 0
        playerLeftSpeed = 75
    End If
Next rightRock

'Checks collision for the rocks that the user can collide with while moving downward
For downRock = 0 To 5
    If imgCharacter.Top + imgCharacter.Height >= imgRockDownHigher(downRock).Top And imgCharacter.Left <= imgRockDownHigher(0).Left + imgRockDownHigher(0).Width And imgCharacter.Left + imgCharacter.Width >= imgRockDownRight.Left Then
        playerDownSpeed = 0
        playerUpSpeed = 75
    End If
Next downRock

'Checks collision for the rocks that the user can collide with while moving right
For rightRock = 0 To 3
    If imgCharacter.Left + imgCharacter.Width >= imgRockRightFurther(rightRock).Left And imgCharacter.Left <= imgRockRightFurther(rightRock).Left + imgRockRightFurther(rightRock).Width And imgCharacter.Top >= imgRockDownRight.Top Then
        playerRightSpeed = 0
        playerLeftSpeed = 75
    End If
Next rightRock



End Sub

Private Sub timerEnemy1Animation_Timer()
enemy1Animation = enemy1Animation + 1

Select Case enemy1Animation
    
    Case Is = 1
        imgEnemy1.Picture = frmPictures.imgGhostRight1.Picture
    
    Case Is = 2
        imgEnemy1.Picture = frmPictures.imgGhostRight2.Picture
        
    Case Is = 3
        imgEnemy1.Picture = frmPictures.imgGhostRight3.Picture
    
    
    Case Is = 4
        imgEnemy1.Picture = frmPictures.imgGhostRight4.Picture
        enemy1Animation = 1
End Select

End Sub

Private Sub timerEnemy1Movement_Timer()
imgEnemy1.Top = imgEnemy1.Top + enemy1Speed

For upRock = 0 To 12
    If imgEnemy1.Top <= imgRockUp(upRock).Top + imgRockUp(upRock).Height Then
        enemy1Speed = -enemy1Speed
    End If
Next upRock

For downRockHigher = 0 To 5
    If imgEnemy1.Top + imgEnemy1.Height >= imgRockDownHigher(downRockHigher).Top Then
        enemy1Speed = -10
    End If
Next downRockHigher

    
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
