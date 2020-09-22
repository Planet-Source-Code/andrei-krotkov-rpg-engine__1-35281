VERSION 5.00
Begin VB.Form frmIntro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Introduction"
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIntro.frx":0000
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   379
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWin 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   1440
   End
   Begin VB.Timer tmrTestResult 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   1680
   End
   Begin VB.Timer tmrTalk 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   2280
      Top             =   1680
   End
   Begin VB.PictureBox picMessage 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   2820
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Label lblGuy 
         BackStyle       =   0  'Transparent
         Caption         =   "Dragon:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblMessage 
         BackStyle       =   0  'Transparent
         Caption         =   "Hey, get back here!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   1920
         TabIndex        =   5
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.Timer tmrDragon 
      Interval        =   100
      Left            =   4200
      Top             =   2520
   End
   Begin VB.PictureBox sprDragon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   2880
      Picture         =   "frmIntro.frx":D1052
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   186
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.PictureBox mskDragon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   2880
      Picture         =   "frmIntro.frx":E12B4
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   186
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Timer tmrMove 
      Interval        =   100
      Left            =   5160
      Top             =   2880
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   1200
      Picture         =   "frmIntro.frx":F1516
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   120
      Picture         =   "frmIntro.frx":F7558
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xpos As Integer
Dim ypos As Integer
Dim LeftX As Integer
Dim TopY As Integer
Dim frame As Integer
Dim dypos As Integer


Private Sub Form_Load()
    
    modRPG.StartRPG
    xpos = -20
    ypos = 161 - 32
    frame = 0
    dypos = -120
    Call MoveIt
        
End Sub


Sub MoveIt()
Dim intCOunter As Integer
'Cls clears all the pixels from the screen.
frmIntro.Cls
'32 here is the length and width of each Wolf Man frame.  He is 32 x 32
frmIntro.Line (0, 161)-(379, 261), , BF
Call BitBlt(frmIntro.Hdc, 100, dypos, mskDragon.Width, mskDragon.Height, mskDragon.Hdc, 0, 0, SRCAND)
Call BitBlt(frmIntro.Hdc, 100, dypos, sprDragon.Width, sprDragon.Height, sprDragon.Hdc, 0, 0, SRCINVERT)
If frame = 1 Then
    Call BitBlt(frmIntro.Hdc, xpos, ypos, 32, 32, picMask.Hdc, 0, 32, SRCAND)
    Call BitBlt(frmIntro.Hdc, xpos, ypos, 32, 32, picSprite.Hdc, 0, 32, SRCINVERT)
ElseIf frame = 0 Then
    Call BitBlt(frmIntro.Hdc, xpos, ypos, 32, 32, picMask.Hdc, 32, 32, SRCAND)
    Call BitBlt(frmIntro.Hdc, xpos, ypos, 32, 32, picSprite.Hdc, 32, 32, SRCINVERT)
Else
    Call BitBlt(frmIntro.Hdc, xpos, ypos, 32, 32, picMask.Hdc, 32, 96, SRCAND)
    Call BitBlt(frmIntro.Hdc, xpos, ypos, 32, 32, picSprite.Hdc, 32, 96, SRCINVERT)
End If
'Refresh must be always be written to tell the program you want to see all the picture information
frmRPG.Refresh
End Sub

Private Sub picMessage_Paint()
    
    For intCOunter = 0 To 255
        
        picMessage.Line (0, intCOunter)-(379, intCOunter), RGB(0, 0, 255 - intCOunter)
    
    Next
    
    
    
End Sub

Private Sub tmrDragon_Timer()

    dypos = dypos + 5
    MoveIt
    If dypos = 60 Then
    tmrDragon.Enabled = False
    picMessage.Visible = True
    End If
    
End Sub

Private Sub tmrMove_Timer()
    
    frame = frame + 1
    If frame = 2 Then frame = 0
    xpos = xpos + 5
    MoveIt
    If xpos = 200 Then
    tmrMove.Enabled = False
    frame = 2
    MoveIt
    lblGuy.Caption = modRPG.Name & ":"
    lblMessage.Caption = "Why?"
tmrTalk.Enabled = True
    End If
    
End Sub

Private Sub tmrTalk_Timer()
Static Talk
Talk = Talk + 1
Select Case Talk
Case 1
    lblGuy.Caption = "Dragon:"
    lblMessage.Caption = "Why? Why? Why you little punk!" & vbCrLf & "I will kill you!"
Case 2
    lblGuy.Caption = modRPG.Name & ":"
    lblMessage.Caption = "You're going down..."
Case 3
    modRPG.BattleResult = 0
    modRPG.EnemyStrength = 15
    modRPG.EnemyWeakness = 25
    modRPG.Enemyhp = 100
    frmBattle.mskEnemy = frmGraphics.msDragon.Picture
    frmBattle.sprEnemy = frmGraphics.sprDragon.Picture
    frmBattle.lblEnemy.Caption = "Dragon Mini-boss"
    frmBattle.Show 1
    tmrTestResult.Enabled = True
End Select

End Sub

Private Sub tmrTestResult_Timer()
    If modRPG.BattleResult = 2 Then
    lblGuy.Caption = modRPG.Name & ":"
    lblMessage.Caption = "Ha!!!! I beat you! Now go away!"
    For intCOunter = 1 To 1000000
DoEvents
Next
    lblGuy.Caption = "Dragon:"
    lblMessage.Caption = "You haven't beat me yet!"
    tmrWin.Enabled = True
    End If
    If modRPG.BattleResult = 1 Then
    lblGuy.Caption = "Dragon:"
    lblMessage.Caption = "Mu-ha-ha-ha-ha-ha-ha"
        For intCOunter = 1 To 1000000
DoEvents
Next
    MsgBox "Game Over." & vbCrLf & "Thanks for playing..."
    Unload Me
    Unload frmBattle
    Unload frmRPG
    End
    End If
End Sub

Private Sub tmrWin_Timer()
    frame = frame + 1
    If frame >= 2 Then frame = 0
    dypos = dypos - 5
    xpos = xpos + 5
    MoveIt
    If xpos >= frmIntro.ScaleWidth Then
        tmrWin.Enabled = False
        Me.Hide
        frmRPG.Show
    End If
End Sub
