VERSION 5.00
Begin VB.Form frmTutorialLevel1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dragonic - Tutorial Level 1"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmTutorialLevel1.frx":0000
   ScaleHeight     =   7440
   ScaleWidth      =   15000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerDragonGrowth 
      Interval        =   250
      Left            =   11160
      Top             =   6720
   End
   Begin VB.Timer timerRightMovement 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   14400
      Top             =   120
   End
   Begin VB.Timer timerCollisionDetection 
      Interval        =   1
      Left            =   7200
      Top             =   120
   End
   Begin VB.Timer timerLeftMovement 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   120
      Top             =   120
   End
   Begin VB.Image imgRockLeftRight 
      Height          =   1095
      Index           =   2
      Left            =   10800
      Picture         =   "frmTutorialLevel1.frx":197DC
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Image imgRockLeftRight 
      Height          =   1095
      Index           =   1
      Left            =   10800
      Picture         =   "frmTutorialLevel1.frx":1A97F
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image imgRockLeftRight 
      Height          =   1095
      Index           =   0
      Left            =   10800
      Picture         =   "frmTutorialLevel1.frx":1BB22
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image imgRockDownLeftRight 
      Height          =   1095
      Left            =   10800
      Picture         =   "frmTutorialLevel1.frx":1CCC5
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Shape shapeExitSquare 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   14640
      Top             =   6960
      Width           =   135
   End
   Begin VB.Shape shapeExitSquare 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   14640
      Top             =   7200
      Width           =   135
   End
   Begin VB.Shape shapeExitSquare 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   13200
      Top             =   6840
      Width           =   135
   End
   Begin VB.Shape shapeExitSquare 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   12960
      Top             =   7080
      Width           =   135
   End
   Begin VB.Shape shapeExitSquare 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   12720
      Top             =   6960
      Width           =   135
   End
   Begin VB.Shape shapeExitSquare 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   12480
      Top             =   7200
      Width           =   135
   End
   Begin VB.Shape shapeExitHighlight 
      BorderStyle     =   0  'Transparent
      DrawMode        =   6  'Mask Pen Not
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   10800
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   5775
   End
   Begin VB.Shape shapeExitSquare 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   12360
      Top             =   6960
      Width           =   135
   End
   Begin VB.Image imgRockRightFarther 
      Height          =   1215
      Index           =   6
      Left            =   14880
      Picture         =   "frmTutorialLevel1.frx":1DE68
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   975
   End
   Begin VB.Image imgRockRightFarther 
      Height          =   1095
      Index           =   5
      Left            =   14880
      Picture         =   "frmTutorialLevel1.frx":1F00B
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   975
   End
   Begin VB.Image imgRockRightFarther 
      Height          =   1095
      Index           =   4
      Left            =   14880
      Picture         =   "frmTutorialLevel1.frx":201AE
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   975
   End
   Begin VB.Image imgRockRightFarther 
      Height          =   1095
      Index           =   3
      Left            =   14880
      Picture         =   "frmTutorialLevel1.frx":21351
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image imgRockRightFarther 
      Height          =   1095
      Index           =   2
      Left            =   14880
      Picture         =   "frmTutorialLevel1.frx":224F4
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   975
   End
   Begin VB.Image imgRockRightFarther 
      Height          =   1095
      Index           =   1
      Left            =   14880
      Picture         =   "frmTutorialLevel1.frx":23697
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image imgRockRightFarther 
      Height          =   1095
      Index           =   0
      Left            =   14880
      Picture         =   "frmTutorialLevel1.frx":2483A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   975
   End
   Begin VB.Image imgRockUpHigher 
      Height          =   1095
      Index           =   6
      Left            =   14280
      Picture         =   "frmTutorialLevel1.frx":259DD
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUpHigher 
      Height          =   1095
      Index           =   5
      Left            =   13200
      Picture         =   "frmTutorialLevel1.frx":26B80
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUpHigher 
      Height          =   1095
      Index           =   4
      Left            =   12000
      Picture         =   "frmTutorialLevel1.frx":27D23
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUpHigher 
      Height          =   1095
      Index           =   3
      Left            =   10800
      Picture         =   "frmTutorialLevel1.frx":28EC6
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUpHigher 
      Height          =   1095
      Index           =   2
      Left            =   9600
      Picture         =   "frmTutorialLevel1.frx":2A069
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUpHigher 
      Height          =   1095
      Index           =   1
      Left            =   8400
      Picture         =   "frmTutorialLevel1.frx":2B20C
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgRockUpHigher 
      Height          =   1095
      Index           =   0
      Left            =   7200
      Picture         =   "frmTutorialLevel1.frx":2C3AF
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Image imgCharacter 
      Height          =   675
      Left            =   2520
      Picture         =   "frmTutorialLevel1.frx":2D552
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   900
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   18
      Left            =   4800
      Picture         =   "frmTutorialLevel1.frx":2E2EC
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   17
      Left            =   3600
      Picture         =   "frmTutorialLevel1.frx":2F1F4
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   16
      Left            =   2400
      Picture         =   "frmTutorialLevel1.frx":300FC
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   15
      Left            =   4800
      Picture         =   "frmTutorialLevel1.frx":31004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   14
      Left            =   3600
      Picture         =   "frmTutorialLevel1.frx":31F0C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   13
      Left            =   2400
      Picture         =   "frmTutorialLevel1.frx":32E14
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   12
      Left            =   1200
      Picture         =   "frmTutorialLevel1.frx":33D1C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   11
      Left            =   1200
      Picture         =   "frmTutorialLevel1.frx":34C24
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   10
      Left            =   0
      Picture         =   "frmTutorialLevel1.frx":35B2C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   9
      Left            =   0
      Picture         =   "frmTutorialLevel1.frx":36A34
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   8
      Left            =   4680
      Picture         =   "frmTutorialLevel1.frx":3793C
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   7
      Left            =   3480
      Picture         =   "frmTutorialLevel1.frx":38844
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   6
      Left            =   2280
      Picture         =   "frmTutorialLevel1.frx":3974C
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   5
      Left            =   1080
      Picture         =   "frmTutorialLevel1.frx":3A654
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgDarkRock 
      Height          =   1095
      Index           =   4
      Left            =   -120
      Picture         =   "frmTutorialLevel1.frx":3B55C
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgRockUpLeft 
      Height          =   1095
      Left            =   6000
      Picture         =   "frmTutorialLevel1.frx":3C464
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image imgRockLeft 
      Height          =   1095
      Index           =   2
      Left            =   6000
      Picture         =   "frmTutorialLevel1.frx":3D607
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   1215
   End
   Begin VB.Image imgRockLeft 
      Height          =   1095
      Index           =   1
      Left            =   6000
      Picture         =   "frmTutorialLevel1.frx":3E7AA
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgRockLeft 
      Height          =   1095
      Index           =   0
      Left            =   6000
      Picture         =   "frmTutorialLevel1.frx":3F94D
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   8
      Left            =   9600
      Picture         =   "frmTutorialLevel1.frx":40AF0
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   7
      Left            =   8400
      Picture         =   "frmTutorialLevel1.frx":41C93
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   6
      Left            =   7200
      Picture         =   "frmTutorialLevel1.frx":42E36
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   4
      Left            =   4800
      Picture         =   "frmTutorialLevel1.frx":43FD9
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   3
      Left            =   3600
      Picture         =   "frmTutorialLevel1.frx":4517C
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   2
      Left            =   2400
      Picture         =   "frmTutorialLevel1.frx":4631F
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   1
      Left            =   1200
      Picture         =   "frmTutorialLevel1.frx":474C2
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image imgRockUp 
      Height          =   1095
      Index           =   0
      Left            =   0
      Picture         =   "frmTutorialLevel1.frx":48665
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   5
      Left            =   6000
      Picture         =   "frmTutorialLevel1.frx":49808
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   4
      Left            =   4800
      Picture         =   "frmTutorialLevel1.frx":4A9AB
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   3
      Left            =   3600
      Picture         =   "frmTutorialLevel1.frx":4BB4E
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   2
      Left            =   2400
      Picture         =   "frmTutorialLevel1.frx":4CCF1
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   1
      Left            =   1200
      Picture         =   "frmTutorialLevel1.frx":4DE94
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image imgRockDown 
      Height          =   1095
      Index           =   0
      Left            =   0
      Picture         =   "frmTutorialLevel1.frx":4F037
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
End
Attribute VB_Name = "frmTutorialLevel1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ICS 2O1
'Reezan Visram
'ISU Game - Dragonic
'A user-controlled dragon makes its way through a maze of tunnels and mazes, avoiding enemies and collecting powerups to make it out of the cave
'Started On: May 28th 2019

'Player Movement Variables (Public so that all levels can access these movement values regardless whether or not they're used in that particular level)
Dim playerUpSpeed As Integer
Dim playerDownSpeed As Integer
Dim playerLeftSpeed As Integer
Dim playerRightSpeed As Integer

'Sprite Animation Variables (Public so that all levels can access these sprite animation variables)
Public leftSprites As Integer
Public rightSprites As Integer

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

'Assigns values to the Sprite Animation variables
leftSprites = 0
rightSprites = 0

End Sub

Private Sub imgRockUpUp_Click(Index As Integer)

End Sub

Private Sub imgRockUpHigh_Click(Index As Integer)
End Sub

Private Sub timerCollisionDetection_Timer()
'REQUIRED so that the character can move after moving away from collision bounds
playerUpSpeed = 50
playerDownSpeed = 50
playerLeftSpeed = 75
playerRightSpeed = 75

'Checks the collision for the rocks that the character can collide with while moving upwards (5 in the left half of the level, 7 in the right half)
For upRock = 0 To 4
    If imgCharacter.Top <= imgRockUp(upRock).Top + imgRockUp(upRock).Height And imgCharacter.Left + imgCharacter.Width >= imgRockUp(upRock).Left And imgCharacter.Left + imgCharacter.Width <= imgRockUp(upRock).Left + imgRockUp(upRock).Width Then
        playerUpSpeed = 0
        playerDownSpeed = 50
    End If
Next upRock

For upRockHigher = 0 To 6
    If imgCharacter.Top <= imgRockUpHigher(upRockHigher).Top + imgRockUpHigher(upRockHigher).Height And imgCharacter.Left + imgCharacter.Width >= imgRockUpHigher(upRockHigher).Left And imgCharacter.Left + imgCharacter.Width <= imgRockUpHigher(upRockHigher).Left + imgRockUpHigher(upRockHigher).Width Then
        playerUpSpeed = 0
        playerDownSpeed = 50
    End If
Next upRockHigher

'Checks the collision for the rocks that the character can collide with while moving downwards (9 in the left half of the level, 1 in the right half of the level )
For downRock = 0 To 8
    If imgCharacter.Top + imgCharacter.Height >= imgRockDown(downRock).Top And imgCharacter.Left + imgCharacter.Width >= imgRockDown(downRock).Left And imgCharacter.Left + imgCharacter.Width <= imgRockDown(downRock).Left + imgRockDown(downRock).Width Then
        playerDownSpeed = 0
        playerUpSpeed = 50
        timerLeftMovement.Enabled = False
        timerRightMovement.Enabled = False
    End If
Next downRock

'Checks the collision for the rocks that the character can collide with while moving left (3 in the left half of the level, 4 in the right half)
For leftRock = 0 To 2
    If imgCharacter.Left <= imgRockLeft(leftRock).Left + imgRockLeft(leftRock).Width And imgCharacter.Top <= imgRockLeft(leftRock).Top + imgRockLeft(leftRock).Height Then
        playerLeftSpeed = 0
        playerRightSpeed = 75
    End If
Next leftRock


         
'Checks the collision for the rocks that the character cab collide with while moving right (3 in the left half of the level, 7 in the right half of the level)
For rightRockFarther = 0 To 6
    If imgCharacter.Left + imgCharacter.Width >= imgRockRightFarther(rightRockFarther).Left And imgCharacter.Top <= imgRockRightFarther(rightRockFarther).Top + imgRockRightFarther(rightRockFarther).Height And imgCharacter.Top + imgCharacter.Height >= imgRockRightFarther(rightRockFarther).Top Then
        playerRightSpeed = 0
        playerLeftSpeed = 75
    End If
Next rightRockFarther

'Checks the collision for the rock in the left half that the character can collide with either moving upwards or moving left
If imgCharacter.Top <= imgRockUpLeft.Top + imgRockUpLeft.Height And imgCharacter.Left + imgCharacter.Width >= imgRockUpLeft.Left And imgCharacter.Left <= imgRockUpLeft.Left + imgRockUpLeft.Width Then
    playerUpSpeed = 0
    playerDownSpeed = 50
    
ElseIf imgCharacter.Left <= imgRockUpLeft.Left + imgRockUpLeft.Width And imgCharacter.Top <= imgRockUpLeft.Top + imgRockUpLeft.Height Then
    playerLeftSpeed = 0
    playerRightSpeed = 75
End If

If imgCharacter.Top >= frmTutorialLevel1.ScaleHeight Then
    Unload frmTutorialLevel1
    frmTutorialLevel2.Show
End If

'Checks the collision for the rocks that the character can collide with either moving left or right
If timerLeftMovement.Enabled = True Then
    For leftRightRock = 0 To 2
        If imgCharacter.Left <= imgRockLeftRight(leftRightRock).Left + imgRockLeftRight(leftRightRock).Width And imgCharacter.Left >= imgRockLeftRight(leftRightRock).Left And imgCharacter.Top + imgCharacter.Height >= imgRockLeftRight(leftRightRock).Top Then
            playerLeftSpeed = 0
            playerRightSpeed = 75
        End If
    Next leftRightRock
End If


If timerRightMovement.Enabled = True Then
    For leftRightRock = 0 To 2
        If imgCharacter.Left + imgCharacter.Width >= imgRockLeftRight(leftRightRock).Left And imgCharacter.Left <= imgRockLeftRight(leftRightRock).Left + imgRockLeftRight(leftRightRock).Width And imgCharacter.Top + imgCharacter.Height >= imgRockLeftRight(leftRightRock).Top Then
            playerRightSpeed = 0
            playerLeftSpeed = 75
        End If
    Next leftRightRock
End If



'Checks the collision for the rock that the character can collide with either moving down, left or right
If imgCharacter.Top + imgCharacter.Height >= imgRockDownLeftRight.Top And imgCharacter.Top <= imgRockDownLeftRight.Top + imgRockDownLeftRight.Height And imgRockDownLeftRight.Left >= imgCharacter.Left And imgRockDownLeftRight.Left + imgRockDownLeftRight.Width <= imgCharacter.Left + imgCharacter.Width Then
    playerDownSpeed = 0
    playerUpSpeed = 50
End If

If timerLeftMovement.Enabled = True Then
    If imgCharacter.Left <= imgRockDownLeftRight.Left + imgRockDownLeftRight.Width And imgCharacter.Left >= imgRockDownLeftRight.Left And imgCharacter.Top + imgCharacter.Height >= imgRockDownLeftRight.Top Then
        playerLeftSpeed = 0
        playerRightSpeed = 75
    End If
End If

If timerRightMovement.Enabled = True Then
    If imgCharacter.Left + imgCharacter.Width >= imgRockDownLeftRight.Left And imgCharacter.Left <= imgRockDownLeftRight.Left + imgRockDownLeftRight.Width And imgCharacter.Top + imgCharacter.Height >= imgRockDownLeftRight.Top Then
        playerRightSpeed = 0
        playerLeftSpeed = 75
    End If
End If



End Sub

Private Sub timerExitIndicator_Timer()

End Sub

Private Sub timerDragonGrowth_Timer()
dragonGrowth = dragonGrowth + 1

Select Case dragonGrowth
Case Is = 1
    imgCharacter.Width = imgCharacter.Width + 100
    imgCharacter.Height = imgCharacter.Height + 100
    
Case Is = 2
    imgCharacter.Width = imgCharacter.Width + 100
    imgCharacter.Height = imgCharacter.Height + 100
    
Case Is = 3
    imgCharacter.Width = imgCharacter.Width + 100
    imgCharacter.Height = imgCharacter.Height + 100

Case Is = 4
    imgCharacter.Width = imgCharacter.Width + 100
    imgCharacter.Height = imgCharacter.Height + 100

Case Is = 5
    imgCharacter.Width = imgCharacter.Width + 100
    imgCharacter.Height = imgCharacter.Height + 100
    
Case Is = 6
    imgCharacter.Width = imgCharacter.Width + 100
    imgCharacter.Height = imgCharacter.Height + 100
    
Case Is = 7
    imgCharacter.Width = imgCharacter.Width + 100
    imgCharacter.Height = imgCharacter.Height + 100
    
Case Is = 8
    imgCharacter.Width = imgCharacter.Width + 100
    imgCharacter.Height = imgCharacter.Height + 100
    
Case Is = 9
    imgCharacter.Width = imgCharacter.Width + 100
    imgCharacter.Height = imgCharacter.Height + 100
    timerDragonGrowth.Enabled = False
    
End Select
End Sub

Private Sub timerLeftMovement_Timer()
'Increments the animation for the sprites facing left
leftSprites = leftSprites + 1

'Loops through the sprites
Select Case leftSprites
    Case Is = 1
        imgCharacter.Picture = frmPictures.imgDragonLeft1.Picture
        
    Case Is = 2
        imgCharacter.Picture = frmPictures.imgDragonLeft2.Picture
        
    Case Is = 3
        imgCharacter.Picture = frmPictures.imgDragonLeft3.Picture
        leftSprites = 0
End Select
End Sub

Private Sub timerRightMovement_Timer()
'Increments the animation for the sprites facing right
rightSprites = rightSprites + 1

'Loops through the sprites
Select Case rightSprites
    Case Is = 1
        imgCharacter.Picture = frmPictures.imgDragonRight1.Picture
    
    Case Is = 2
        imgCharacter.Picture = frmPictures.imgDragonRight2.Picture
        
    Case Is = 3
        imgCharacter.Picture = frmPictures.imgDragonRight3.Picture
        rightSprites = 0
End Select
End Sub

