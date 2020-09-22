VERSION 5.00
Begin VB.Form frmItems 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MousePointer    =   12  'No Drop
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUpdate 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   304
      Y1              =   104
      Y2              =   104
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
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   320
      Y1              =   168
      Y2              =   168
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
      Top             =   2160
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
   Begin VB.Image imgBWater 
      Height          =   480
      Left            =   2520
      Picture         =   "frmItems.frx":0000
      Top             =   600
      Visible         =   0   'False
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3240
      Picture         =   "frmItems.frx":0C42
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgBFaerie 
      Height          =   480
      Left            =   2040
      Picture         =   "frmItems.frx":1884
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMedal 
      Height          =   480
      Left            =   600
      Picture         =   "frmItems.frx":24C6
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLetter 
      Height          =   480
      Left            =   120
      Picture         =   "frmItems.frx":3108
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBMPotion 
      Height          =   480
      Left            =   1560
      Picture         =   "frmItems.frx":3D4A
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBHPotion 
      Height          =   480
      Left            =   1080
      Picture         =   "frmItems.frx":498C
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBBPotion 
      Height          =   480
      Left            =   600
      Picture         =   "frmItems.frx":55CE
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBEmpty 
      Height          =   480
      Left            =   120
      Picture         =   "frmItems.frx":6210
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMPotion 
      Height          =   480
      Left            =   600
      Picture         =   "frmItems.frx":6E52
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHPotion 
      Height          =   480
      Left            =   120
      Picture         =   "frmItems.frx":7A94
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xleft
Dim ytop
Sub PicShow()
    If intItemsOwn(0) > 0 Then imgHPotion.Visible = True Else imgHPotion.Visible = False
    If intItemsOwn(10) > 0 Then imgMPotion.Visible = True Else imgMPotion.Visible = False
    
    If intItemsOwn(1) > 0 Then imgBEmpty.Visible = True Else imgBEmpty.Visible = False
    If intItemsOwn(11) > 0 Then imgBBPotion.Visible = True Else imgBBPotion.Visible = False
    If intItemsOwn(21) > 0 Then imgBHPotion.Visible = True Else imgBHPotion.Visible = False
    If intItemsOwn(31) > 0 Then imgBMPotion.Visible = True Else imgBMPotion.Visible = False
    If intItemsOwn(41) > 0 Then imgBFaerie.Visible = True Else imgBFaerie.Visible = False
    If intItemsOwn(51) > 0 Then imgBWater.Visible = True Else imgBWater.Visible = False
    
    If intItemsOwn(2) > 0 Then imgLetter.Visible = True Else imgLetter.Visible = False
    If intItemsOwn(12) > 0 Then imgMedal.Visible = True Else imgMedal.Visible = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    lblDisplay = ""
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
        
            If ytop < 2 And shpClose.Visible = False Then
            ytop = ytop + 1
            Else
            shpMove.Visible = False
            shpClose.Visible = True
            End If
            
        Case 32
            If shpClose.Visible = True Then
                Me.Hide
            End If
            
            If intItemsOwn(xleft & ytop) > 0 Then
                Select Case xleft & ytop
                    Case 0 'HP Potion (No bottle)
                        HP = HP + 20
                        lblDisplay = "Healed 20 HP!"
                        If HP > MaxHP Then HP = MaxHP
                        intItemsOwn(xleft & ytop) = intItemsOwn(xleft & ytop) - 1
                    Case 10 'MP Potion (No Bottle)
                        MP = HP + 20
                        lblDisplay = "Healed 20 MP!"
                        If MP > MaxMP Then MP = MaxMP
                        intItemsOwn(xleft & ytop) = intItemsOwn(xleft & ytop) - 1
                    Case 11
                        MP = HP + 20
                        HP = HP + 20
                        lblDisplay = "Healed 20 HP and MP!"
                        If HP > MaxHP Then HP = MaxHP
                        If MP > MaxMP Then MP = MaxMP
                        intItemsOwn(xleft & ytop) = intItemsOwn(xleft & ytop) - 1
                        intItemsOwn(1) = intItemsOwn(1) + 1
                    Case 21
                        HP = HP + 20
                        lblDisplay = "Healed 20 HP!"
                        If HP > MaxHP Then HP = MaxHP
                        intItemsOwn(xleft & ytop) = intItemsOwn(xleft & ytop) - 1
                        intItemsOwn(1) = intItemsOwn(1) + 1
                    Case 31
                        MP = HP + 20
                        lblDisplay = "Healed 20 MP!"
                        If MP > MaxMP Then MP = MaxMP
                        intItemsOwn(xleft & ytop) = intItemsOwn(xleft & ytop) - 1
                        intItemsOwn(1) = intItemsOwn(1) + 1
                End Select
            End If
    End Select
    If intItemsOwn(xleft & ytop) > 0 Then
    lblItem.Caption = modRPG.strItems(xleft & ytop) & vbCrLf & " x " & intItemsOwn(xleft & ytop)
    Else
    lblItem.Caption = ""
    End If
    shpMove.Move xleft * 32 + 7, ytop * 32 + 8
    Call PicShow
End Sub

Private Sub Form_Load()

    xleft = 0
    ytop = 0
    intCounter = 0
    Open App.Path & "/items.txt" For Input As 1
    Do
    Line Input #1, strLine
    strItems(intCounter) = strLine
    intCounter = intCounter + 1
    Loop Until EOF(1) = True
    Close 1


End Sub


Private Sub tmrUpdate_Timer()
Call PicShow
End Sub
