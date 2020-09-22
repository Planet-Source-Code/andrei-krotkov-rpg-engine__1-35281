VERSION 5.00
Begin VB.Form frmBackground 
   BackColor       =   &H00000004&
   BorderStyle     =   0  'None
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3120
      Top             =   2040
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmBackground.frx":0000
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "The story of a knight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Andrei's RPG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload frmBattle
    Unload frmGraphics
    Unload frmIntro
    Unload frmName
    Unload frmRPG
    Unload frmStart
    Unload frmBackground
End Sub

Private Sub Form_Activate()
   With frmBackground
       .Left = 0
        .Top = 0
        .Width = Screen.Width
        .Height = Screen.Height
    End With
End Sub

Private Sub tmrUpdate_Timer()
    
    lblHP.Caption = modRPG.HP
    lblMP.Caption = modRPG.MP
    lblCash.Caption = modRPG.cash
    
End Sub

Private Sub Timer1_Timer()
Label3.Caption = cash
End Sub
