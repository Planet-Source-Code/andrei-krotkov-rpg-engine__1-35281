VERSION 5.00
Begin VB.Form frmBattle 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   12  'No Drop
   Picture         =   "frmBattle.frx":0000
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUpdate 
      Interval        =   1
      Left            =   1800
      Top             =   2160
   End
   Begin VB.PictureBox mskDark 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   4560
      Picture         =   "frmBattle.frx":57B2
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   3120
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   1635
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   360
         ScaleHeight     =   240
         ScaleMode       =   0  'User
         ScaleWidth      =   100
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   360
         Width           =   1080
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "HP"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   360
            TabIndex        =   30
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lblEHP 
            BackColor       =   &H00FF8080&
            Height          =   255
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   5
         X1              =   120
         X2              =   1680
         Y1              =   90.667
         Y2              =   90.667
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Dragon"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   1635
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleMode       =   0  'User
         ScaleWidth      =   100
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   360
         Width           =   1080
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "HP"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lblHP 
            BackColor       =   &H008080FF&
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleMode       =   0  'User
         ScaleWidth      =   100
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   720
         Width           =   1080
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "MP"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lblMP 
            BackColor       =   &H0080FF80&
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   5
         X1              =   0
         X2              =   1560
         Y1              =   59.13
         Y2              =   59.13
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Knight"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox sprDragon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1830
      Left            =   4200
      Picture         =   "frmBattle.frx":70F4
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   186
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox mskDragon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1830
      Left            =   4560
      Picture         =   "frmBattle.frx":17356
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   186
      TabIndex        =   18
      Top             =   4680
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox mskSword 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1110
      Left            =   4560
      Picture         =   "frmBattle.frx":275B8
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox sprSword 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1110
      Left            =   4560
      Picture         =   "frmBattle.frx":2B90A
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   16
      Top             =   4680
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox mskIce 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1425
      Left            =   4680
      Picture         =   "frmBattle.frx":2FC5C
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox mskCross 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   4560
      Picture         =   "frmBattle.frx":33506
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox sprCross 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   4440
      Picture         =   "frmBattle.frx":38F48
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox sprIce 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1425
      Left            =   4680
      Picture         =   "frmBattle.frx":3E98A
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox mskFire 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   4560
      Picture         =   "frmBattle.frx":42234
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox sprFire 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   4560
      Picture         =   "frmBattle.frx":44366
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   4320
      Picture         =   "frmBattle.frx":46498
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   4320
      Picture         =   "frmBattle.frx":4C4DA
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Message 
      BackColor       =   &H00000000&
      Height          =   1335
      Left            =   0
      MousePointer    =   12  'No Drop
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   316
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3480
      Width           =   4800
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ice Crystal 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fireball 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Healing 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Run"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Magic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         Height          =   255
         Left            =   120
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Physical"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblMessage 
         BackStyle       =   0  'Transparent
         Caption         =   "(Enemy) wants to fight!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xguy As Integer
Dim frame As Integer
Dim dire As Integer '32 = right; 96 = left
Dim iend As Integer

Private Sub Form_GotFocus()
Call Stats
End Sub

Private Sub Message_Paint()
    On Error Resume Next
    For a = 0 To 255
    
        Message.Line (0, a)-(Message.ScaleWidth, a), RGB(0, 0, 255 - a)
        Picture1.Line (0, a)-(Picture1.ScaleWidth, a), RGB(0, 0, 255 - a)
        Picture3.Line (0, a)-(Picture1.ScaleWidth, a), RGB(0, 0, 255 - a)
    
    Next
    
End Sub

Sub PaintScreen()
    frmBattle.Cls
    Call BitBlt(frmBattle.Hdc, xguy, 132, 32, 32, picMask.Hdc, frame, dire, SRCAND)
    Call BitBlt(frmBattle.Hdc, xguy, 132, 32, 32, picSprite.Hdc, frame, dire, SRCINVERT)
    Call BitBlt(frmBattle.Hdc, 192, 60, mskDragon.Width, mskDragon.Height, mskDragon.Hdc, 0, 0, SRCAND)
    Call BitBlt(frmBattle.Hdc, 192, 60, mskDragon.Width, mskDragon.Height, sprDragon.Hdc, 0, 0, SRCINVERT)
End Sub

Sub Stats()
On Error Resume Next
    lblMP.Width = 100 * (MP / MaxMP)
    lblHP.Width = 100 * (HP / MaxHP)
    lblEHP.Width = 100 * (EHP / eMaxHP)
End Sub


Private Sub Form_Activate()
    
    Call Stats

    dire = 32
    Label4.Caption = "Heal " & hLevel
    Label5.Caption = "Fireball " & Flevel
    Label6.Caption = "Ice Crystal " & ilevel
    For xguy = -32 To 10 Step 5
    If frame = 32 Then frame = 0 Else frame = 32
    Call PaintScreen
    frmBattle.Refresh
    For b = 0 To 5000000
    Next
    Next
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Shape1.Visible = True
    lblMessage.Visible = False
    
End Sub

Private Sub Message_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Randomize
    If lblMessage.Visible = False Then
    
    Select Case KeyCode
        
        Case 38 'Up
        
            Shape1.Top = Shape1.Top - 72
            
            If Shape1.Top < 24 Then Shape1.Top = 168
            
        Case 40 ' Down
        
            Shape1.Top = Shape1.Top + 72
            
            If Shape1.Top > 168 Then Shape1.Top = 24
            
        Case 32
                        
            If Shape1.Top = 96 And Shape1.Left <> 112 Then
                Label6.Visible = True
                Label4.Visible = True
                Label5.Visible = True
                Shape1.Top = 24
                Shape1.Left = 112
                Exit Sub
            End If
            'RUN
            If Shape1.Top = 168 And Shape1.Left <> 112 Then
                Call HideAll
                lblMessage.Caption = "You try to run away..."
                If modRPG.Canrun = True Then
                Call Run
                Me.Hide
                Else
                lblMessage.Caption = lblMessage.Caption & "You failed..."
                Call DetAttack
                End If
                Call ShowMain
            End If
            'PHYSICAL
            If Shape1.Top = 24 And Shape1.Left <> 112 Then
                Call HideAll
                Call Physical
                Call Stats
                Call DetAttack
                Call Stats
                Call ShowMain
                Call PaintScreen
            End If
            'FIRE
            If Shape1.Top = 96 And Shape1.Left = 112 And MP >= Flevel Then
                Call HideAll
                Call Fire
                Call Stats
                Call DetAttack
                Call Stats
                Call ShowAll
                Call PaintScreen
            End If
            'HEAL
            If Shape1.Top = 24 And Shape1.Left = 112 And MP >= hLevel Then
               Call HideAll
               Call Heal
                Call Stats
                Call DetAttack
                Call Stats
                Call ShowAll
                Call PaintScreen
            End If
            'ICE
            If Shape1.Top = 168 And Shape1.Left = 112 And MP >= ilevel Then
                Call HideAll
                Call Ice
                Call Stats
                Call DetAttack
                Call Stats
                Call ShowAll
                Call PaintScreen
            End If
            If EHP <= 0 Then
                Me.Hide
            End If
                If HP <= 0 Then
                    MsgBox "I'm sorry, brave knight. You were killed in combat..."
                    Unload Me
                    End
                End If
            Case 27
            
            Label6.Visible = False
            Label4.Visible = False
            Label5.Visible = False
            Shape1.Top = 24
            Shape1.Left = 8
        
        End Select
    End If
    Stats
End Sub



Public Sub Physical()
                Damage = Int((Str / edef) * Str / 2) + Int(Rnd * 3 + 1)
                EHP = EHP - Damage
                lblMessage.Caption = "You go up to the enemy and use your sword..." & vbCrLf & Damage & " damage."
                lblMessage.Visible = True
                dire = 32
                For xguy = 10 To 145 Step 5
                    If frame = 32 Then frame = 0 Else frame = 32
                    Call PaintScreen
                    frmBattle.Refresh
                    For b = 0 To 750000
                      Next
                Next
                For a = 0 To 264 - 160 + 40 Step 5
                    PaintScreen
                    Call BitBlt(frmBattle.Hdc, 160 + a, 32 + a, mskSword.Width, mskSword.Height, mskSword.Hdc, 0, 0, SRCAND)
                    Call BitBlt(frmBattle.Hdc, 160 + a, 32 + a, mskSword.Width, mskSword.Height, sprSword.Hdc, 0, 0, SRCINVERT)
                    frmBattle.Refresh
                    For b = 0 To 750000
                    Next
                Next
                dire = 96
                For xguy = 145 To 10 Step -5
                    If frame = 32 Then frame = 0 Else frame = 32
                    Call PaintScreen
                    frmBattle.Refresh
                    For b = 0 To 1000000
                      Next
                Next
                dire = 32
                Call PaintScreen
                
                frmBattle.Refresh
                abc = True
End Sub

Public Sub Run()

                dire = 96
                For xguy = 10 To -32 Step -5
                If frame = 32 Then frame = 0 Else frame = 32
                Call PaintScreen
                frmBattle.Refresh
                For b = 0 To 2500000
                Next
                Next
                Exit Sub
                
End Sub

Public Sub Fire()
                If Weakness = 1 Then Damage = Flevel * 2 * 3
                If Weakness = 2 Then Damage = Int(Flevel * 2 / 3)
                If Weakness = 3 Then Damage = Flevel * 2
                If Damage <> 0 Then Damage = Damage + Int(Rnd * 4) + 1
                EHP = EHP - Damage
                MP = MP - Flevel
                lblMessage.Caption = "Level " & Flevel & " Fire attack!" & vbCrLf & Damage & " damage!"
                lblMessage.Visible = True
                If hLevel > MP Then Label4.ForeColor = vbRed
                If Flevel > MP Then Label5.ForeColor = vbRed
                If ilevel > MP Then Label6.ForeColor = vbRed
                abc = False
                y = 50
                For xpos = 0 To frmBattle.ScaleWidth Step 1
                    If xpos <= frmBattle.ScaleWidth / 2 - sprFire.Width / 2 Then y = y - 0.5 Else y = y + 0.5
                        frmBattle.Cls
                        Rand = Int(Rnd * 4)
                    PaintScreen
                    Call BitBlt(frmBattle.Hdc, xpos, 50 + y + Rand, sprFire.Width, sprFire.Height, mskFire.Hdc, 0, 0, SRCAND)
                    Call BitBlt(frmBattle.Hdc, xpos, 50 + y + Rnd * 2, sprFire.Width, sprFire.Height, sprFire.Hdc, 0, 0, SRCINVERT)
                    frmBattle.Refresh
                    For a = 0 To 5000
                    Next
                Next
                abc = True
End Sub

Public Sub Heal()
            Damage = hLevel * 2
            If Damage <> 0 Then Damage = Damage + Int(Rnd * 4) + 1
            If HP + Damage < MaxHP Then HP = HP + Damage Else HP = MaxHP

            MP = MP - hLevel
            If hLevel > MP Then Label4.ForeColor = vbRed
            If Flevel > MP Then Label5.ForeColor = vbRed
            If ilevel > MP Then Label6.ForeColor = vbRed
            lblMessage.Caption = "Level " & hLevel & " Healing!" & vbCrLf & "You recovered " & Damage & " damage..."
            abc = False
            y = 50
                       
            For xpos = Message.Top To -1 * sprCross.Height - 50 Step -1
            PaintScreen
            Call BitBlt(frmBattle.Hdc, -10, xpos, sprCross.Width, sprCross.Height, mskCross.Hdc, 0, 0, SRCAND)
            Call BitBlt(frmBattle.Hdc, -10, xpos, sprCross.Width, sprCross.Height, sprCross.Hdc, 0, 0, SRCINVERT)
            frmBattle.Refresh
            
            Next
           
            abc = True

End Sub
Public Sub Ice()
                If Weakness = 2 Then Damage = ilevel * 6
                If Weakness = 1 Then Damage = Int(ilevel)
                If Weakness = 3 Then Damage = ilevel * 2
                If Damage <> 0 Then Damage = Damage + Int(Rnd * 4) + 1
            EHP = EHP - Damage
            MP = MP - ilevel
            If hLevel > MP Then Label4.ForeColor = vbRed
            If Flevel > MP Then Label5.ForeColor = vbRed
            If ilevel > MP Then Label6.ForeColor = vbRed
            lblMessage.Caption = "Level " & ilevel & " Ice attack!" & vbCrLf & Damage & " damage."
            abc = False
            y = 60
            
            For xpos = -1 * sprIce.Width To frmBattle.ScaleWidth Step 1
            PaintScreen
            Call BitBlt(frmBattle.Hdc, xpos, 50 + y - 30, sprIce.Width, sprIce.Height, mskIce.Hdc, 0, 0, SRCAND)
            Call BitBlt(frmBattle.Hdc, xpos, 50 + y - 30, sprIce.Width, sprIce.Height, sprIce.Hdc, 0, 0, SRCINVERT)
            frmBattle.Refresh
            For a = 0 To 10000
            Next
            
            Next
            abc = True
End Sub

Sub EFire()
                Damage = Int((Estr / def) * Estr / 2) + Int(Rnd * 3 + 1)
                If Damage <> 0 Then Damage = Damage + Int(Rnd * 4) + 1
                HP = HP - Damage
                lblMessage.Caption = "Enemy attacks with Fire" & vbCrLf & Damage & " damage."
                lblMessage.Visible = True
                If hLevel > MP Then Label4.ForeColor = vbRed
                If Flevel > MP Then Label5.ForeColor = vbRed
                If ilevel > MP Then Label6.ForeColor = vbRed
                abc = False
                y = 25
                For xpos = 192 To 0 Step -1
                        If xpos >= frmBattle.ScaleWidth / 2 - sprFire.Width / 2 Then y = y - 0.5 Else y = y + 0.5
                        frmBattle.Cls
                        Rand = Int(Rnd * 4)
                        PaintScreen
                        Call BitBlt(frmBattle.Hdc, xpos, 50 + y + Rand, sprFire.Width, sprFire.Height, mskFire.Hdc, 0, 0, SRCAND)
                        Call BitBlt(frmBattle.Hdc, xpos, 50 + y + Rnd * 2, sprFire.Width, sprFire.Height, sprFire.Hdc, 0, 0, SRCINVERT)
                        frmBattle.Refresh
                        For a = 0 To 10000
                        Next
                Next
End Sub
Sub EDark()
                Damage = Int((Estr / def) * Estr / 2) + Int(Rnd * 3 + 1)
                If Damage <> 0 Then Damage = Damage + Int(Rnd * 4) + 1
                HP = HP - Damage
                lblMessage.Caption = "Enemy attacks" & vbCrLf & Damage & " damage."
                abc = False
                        For a = 0 To 500000
                           Next
End Sub

Sub ShowAll()
            Label6.Visible = True
            Label4.Visible = True
            Label5.Visible = True
            Label1.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Shape1.Visible = True
            lblMessage.Visible = False
End Sub
Sub ShowMain()
            Label1.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Shape1.Visible = True
            lblMessage.Visible = False
End Sub

Sub HideAll()
            Label6.Visible = False
            Label4.Visible = False
            Label5.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            Label3.Visible = False
            Shape1.Visible = False
            lblMessage.Visible = True
End Sub

Sub DetAttack()
            If Eanim = 1 Then
                Call EFire
            ElseIf Eanim = 2 Then
                Call EIce
            ElseIf Eanim = 3 Then
                Call EDark
            Else
                Call Ephys
            End If
End Sub
Sub EIce()
                Damage = Int((Estr / def) * Estr / 2) + Int(Rnd * 3 + 1)
                If Damage <> 0 Then Damage = Damage + Int(Rnd * 4) + 1
                HP = HP - Damage
                lblMessage.Caption = "Enemy attacks with an Ice Storm!" & vbCrLf & Damage & " damage."
                abc = False
                y = 60
            
            For xpos = frmBattle.ScaleWidth To -1 * sprIce.Width Step -1
            PaintScreen
            Call BitBlt(frmBattle.Hdc, xpos, 50 + y - 30, sprIce.Width, sprIce.Height, mskIce.Hdc, 0, 0, SRCAND)
            Call BitBlt(frmBattle.Hdc, xpos, 50 + y - 30, sprIce.Width, sprIce.Height, sprIce.Hdc, 0, 0, SRCINVERT)
            frmBattle.Refresh
            For a = 0 To 10000
            Next
            
            Next
End Sub
Sub Ephys()
                Damage = Int((Estr / def) * Estr / 2) + Int(Rnd * 3 + 1)
                If Damage <> 0 Then Damage = Damage + Int(Rnd * 4) + 1
                HP = HP - Damage
                lblMessage.Caption = "Enemy attacks with Darkness" & vbCrLf & Damage & " damage."
                abc = False
                For a = 0 To 1000000
                Next
End Sub

Private Sub tmrUpdate_Timer()
Call Stats
End Sub
