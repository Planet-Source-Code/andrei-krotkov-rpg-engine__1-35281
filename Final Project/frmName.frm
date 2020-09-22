VERSION 5.00
Begin VB.Form frmName 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Knight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Hero Name:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii > 64 And KeyAscii < 123 Or KeyAscii = 32 Then
        lblName.Caption = lblName.Caption & Chr(KeyAscii)
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
    
    If KeyCode = 8 Then
        lblName.Caption = Left(lblName.Caption, Len(lblName.Caption) - 1)
    ElseIf KeyCode = 13 Then
        modRPG.Name = lblName.Caption
        Unload Me
        Load frmIntro
        frmIntro.Show
        frmIntro.ZOrder 0
    End If
        
End Sub
