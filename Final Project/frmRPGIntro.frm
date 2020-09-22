VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   1080
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   240
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Width           =   4215
      Begin VB.Label lblMove 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRPGIntro.frx":0000
         ForeColor       =   &H00FFFFFF&
         Height          =   4815
         Left            =   0
         TabIndex        =   2
         Top             =   2520
         Width           =   4215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RPG Demo: A Knight's Tale"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Load frmRPG
    frmRPG.ZOrder 0
    frmRPG.Show
    Me.Hide
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Load frmRPG
    frmRPG.ZOrder 0
    frmRPG.Show
    Me.Hide
End Sub

Private Sub Timer1_Timer()
lblMove.Top = lblMove.Top - 1
If lblMove.Top + lblMove.Height <= 0 Then lblMove.Top = Picture1.Height
End Sub
