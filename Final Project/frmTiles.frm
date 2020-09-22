VERSION 5.00
Begin VB.Form frmTiles 
   Caption         =   "Tile Viewer"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   ScaleHeight     =   143
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   197
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTilesBottom 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   1440
      Picture         =   "frmTiles.frx":0000
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   1
      Top             =   120
      Width           =   1440
   End
   Begin VB.PictureBox picView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblDown 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblLeft 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblRight 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblInfo 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblUp 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line12 
      X1              =   40
      X2              =   48
      Y1              =   64
      Y2              =   72
   End
   Begin VB.Line Line11 
      X1              =   56
      X2              =   48
      Y1              =   64
      Y2              =   72
   End
   Begin VB.Line Line10 
      X1              =   48
      X2              =   48
      Y1              =   48
      Y2              =   72
   End
   Begin VB.Line Line9 
      X1              =   24
      X2              =   16
      Y1              =   48
      Y2              =   40
   End
   Begin VB.Line Line8 
      X1              =   24
      X2              =   16
      Y1              =   32
      Y2              =   40
   End
   Begin VB.Line Line7 
      X1              =   40
      X2              =   16
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Line Line6 
      X1              =   72
      X2              =   80
      Y1              =   48
      Y2              =   40
   End
   Begin VB.Line Line5 
      X1              =   72
      X2              =   80
      Y1              =   32
      Y2              =   40
   End
   Begin VB.Line Line4 
      X1              =   56
      X2              =   80
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Line Line3 
      X1              =   56
      X2              =   48
      Y1              =   16
      Y2              =   8
   End
   Begin VB.Line Line2 
      X1              =   40
      X2              =   48
      Y1              =   16
      Y2              =   8
   End
   Begin VB.Line Line1 
      X1              =   48
      X2              =   48
      Y1              =   8
      Y2              =   32
   End
End
Attribute VB_Name = "frmTiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xpos As Integer
Dim ypos As Integer

Sub Main()
    
    picView.Cls
    'Call BitBlt(picView.Hdc, 0, 0, 32, 32, picBlank.Hdc, 0, 0, SRCINVERT)
    Call BitBlt(picView.Hdc, 0, 0, 32, 32, picTilesBottom.Hdc, xpos, ypos, SRCPAINT)
    picView.Refresh
    
    
End Sub

Private Sub Form_Activate()
    
    xpos = 0
    ypos = 0
    Call Main
    
End Sub

Private Sub lblDown_Click()

    ypos = ypos + 32
    If ypos > 4 * 32 Then ypos = 4 * 32
    Call Main

End Sub

Private Sub lblLeft_Click()

    xpos = xpos - 32
    If xpos <= 0 Then xpos = 0
    Call Main
    
End Sub

Private Sub lblRight_Click()
    
    xpos = xpos + 32
    If xpos >= 64 Then xpos = 64
    Call Main
    
End Sub

Private Sub lblUp_Click()
    
    ypos = ypos - 32
    If ypos < 0 Then ypos = 0
    Call Main
    
End Sub
