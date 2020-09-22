VERSION 5.00
Begin VB.Form frmShop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "What would you like to buy today?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   304
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   320
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Label lblClose 
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   4200
      Width           =   855
   End
   Begin VB.Shape shpClose 
      BorderWidth     =   5
      Height          =   375
      Left            =   3675
      Top             =   4200
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblDisplay 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Shape shpMove 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3240
      Picture         =   "frmShop.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "x0"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Image imgHPotion 
      Height          =   480
      Left            =   120
      Picture         =   "frmShop.frx":0C42
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgMPotion 
      Height          =   480
      Left            =   600
      Picture         =   "frmShop.frx":1884
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgBEmpty 
      Height          =   480
      Left            =   120
      Picture         =   "frmShop.frx":24C6
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "frmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim xleft
Dim ytop
Dim intItemC(999) As Integer
Sub PicShow()

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    lblMessage = ""
    Select Case KeyCode
        
        Case 37
            If xleft > 0 And shpClose.Visible = False Then xleft = xleft - 1
        
        Case 38
            If ytop > 0 And shpClose.Visible = False Then
            ytop = ytop - 1
            Else
            shpMove.Visible = True
            shpClose.Visible = False
            End If
        
        Case 39
        
            If xleft < 5 And shpClose.Visible = False Then xleft = xleft + 1
        
        Case 40
        
            If ytop < 1 And shpClose.Visible = False Then
            ytop = ytop + 1
            Else
            shpMove.Visible = False
            shpClose.Visible = True
            End If
            
        Case 32
            If shpClose.Visible = True Then
                Me.Hide
                Exit Sub
            End If
            If modRPG.strItems(xleft & ytop) <> "Blank" And Not (xleft >= 1 And ytop = 1) Then
            If cash >= intItemC(xleft & ytop) Then
                modRPG.intItemsOwn(xleft & ytop) = modRPG.intItemsOwn(xleft & ytop) + 1
                cash = cash - intItemC(xleft & ytop)
                Label1.Caption = "x" & cash
                lblMessage.Caption = "Bought " & modRPG.strItems(xleft & ytop)
            Else
                lblMessage.Caption = "Not enough cash."
            End If
            End If
    End Select
    If modRPG.strItems(xleft & ytop) <> "Blank" And Not (xleft >= 1 And ytop = 1) Then
    lblItem.Caption = modRPG.strItems(xleft & ytop) & vbCrLf & " Owned: x" & modRPG.intItemsOwn(xleft & ytop) & vbCrLf & " Cost: " & intItemC(xleft & ytop)
    Else
    lblItem = ""
    End If
    shpMove.Move xleft * 32 + 7, ytop * 32 + 8
    Call PicShow
End Sub

Private Sub Form_Load()
    xleft = 0
    ytop = 0
    intCounter = 0
    Open App.Path & "/items.txt" For Input As 10
    Do
    Line Input #10, strLine
    strItems(intCounter) = strLine
    intCounter = intCounter + 1
    Loop Until EOF(10) = True
    Close 10
    intCounter = 0
    Open App.Path & "/itemc.txt" For Input As 2
    Do
    Line Input #2, strLine
    intItemC(intCounter) = CInt(strLine)
    intCounter = intCounter + 1
    Loop Until EOF(2) = True
    Close 2
Label1.Caption = "x" & cash

End Sub

Private Sub tmrUpdate_Timer()
Call PicShow
End Sub
