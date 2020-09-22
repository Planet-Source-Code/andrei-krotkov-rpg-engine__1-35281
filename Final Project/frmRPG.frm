VERSION 5.00
Begin VB.Form frmRPG 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MousePointer    =   12  'No Drop
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   317
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Shape shpChoose 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   255
         Left            =   0
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   " No"
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
         TabIndex        =   10
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " Yes"
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
         TabIndex        =   9
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label lblQuestion 
         BackStyle       =   0  'Transparent
         Caption         =   "Question"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.PictureBox sprppl 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   600
      Picture         =   "frmRPG.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox Message 
      BackColor       =   &H00000000&
      Height          =   1335
      Left            =   0
      MousePointer    =   12  'No Drop
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   316
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   4800
      Begin VB.Timer tmrNext 
         Interval        =   1000
         Left            =   2760
         Top             =   840
      End
      Begin VB.Image imgNext 
         Height          =   225
         Left            =   2160
         Picture         =   "frmRPG.frx":C042
         Top             =   960
         Width           =   300
      End
      Begin VB.Label lblMessage 
         BackStyle       =   0  'Transparent
         Caption         =   "Help Me!"
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
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Who 
         BackStyle       =   0  'Transparent
         Caption         =   "Voice from above"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.PictureBox mskppl 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   600
      Picture         =   "frmRPG.frx":C408
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Timer battres 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   120
      Picture         =   "frmRPG.frx":1844A
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   120
      Picture         =   "frmRPG.frx":1E48C
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Shape guy 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   5
      Height          =   480
      Left            =   1440
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmRPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xpos As Integer, ypos As Integer, frame As Integer
Dim intColmap As Integer, intRowMap As Integer
Dim intFileNum As Integer, intFrame As Integer, intDir As Integer
Dim Level(10, 10) As Integer
Dim xleft As Integer, ytop As Integer
Dim abc As Boolean
Dim def As Boolean
Dim yes As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Static steps As Integer
    If KeyCode = 37 Then
        steps = steps + 1
        If xpos - 1 = 0 Then
            intColmap = intColmap - 1
            xpos = 10
            ytop = 96
        If xleft = 0 Then xleft = 32 Else xleft = 0
            GoTo 1
        End If
        If Level(xpos - 1, ypos) = 1 Then
            xpos = xpos - 1
        End If
        ytop = 96
        If xleft = 0 Then xleft = 32 Else xleft = 0
    End If
    If KeyCode = 39 Then
    steps = steps + 1
        If xpos + 1 > 10 Then
            intColmap = intColmap + 1
            xpos = 1
                    ytop = 32
        If xleft = 0 Then xleft = 32 Else xleft = 0
            GoTo 1
        End If
        If Level(xpos + 1, ypos) = 1 Then
            xpos = xpos + 1
        End If
        ytop = 32
        If xleft = 0 Then xleft = 32 Else xleft = 0
    End If
    If KeyCode = 38 Then
    steps = steps + 1
        If ypos - 1 = 0 Then
            intRowMap = intRowMap - 1
            ypos = 10
                    ytop = 0
        If xleft = 0 Then xleft = 32 Else xleft = 0
            GoTo 1
        End If
        If Level(xpos, ypos - 1) = 1 Then
            ypos = ypos - 1
        End If
        ytop = 0
        If xleft = 0 Then xleft = 32 Else xleft = 0
    End If
    If KeyCode = 40 Then
    steps = steps + 1
        If ypos + 1 > 10 Then
            intRowMap = intRowMap + 1
            ypos = 1

                    ytop = 64
        If xleft = 0 Then xleft = 32 Else xleft = 0
             GoTo 1
             End If
        If Level(xpos, ypos + 1) = 1 Then
            ypos = ypos + 1
        End If
        ytop = 64
        If xleft = 0 Then xleft = 32 Else xleft = 0
    End If
    If KeyCode = 32 Then abc = True
    If KeyCode = 27 Then
    frmItems.Show 1
    Do
    
    
    
    Loop Until frmItems.Visible = False
    End If
1
    Debug.Print xpos, ypos
    guy.Left = (xpos * 32) - 32
    guy.Top = (ypos * 32) - 32
    PaintScreen
End Sub

Private Sub Form_Load()
    
    xpos = 4
    ypos = 5
    frame = 0
    intColmap = 2
    intRowMap = 2
    PaintScreen
        
End Sub

Sub PaintScreen()
    
    'On Error Resume Next
    
    Dim strLine As String
    Dim intTile As Integer
    Dim intTile2 As Integer
        guy.Left = (xpos * 32) - 32
    guy.Top = (ypos * 32) - 32
    frmRPG.Cls
    'For x = 0 To frmRPG.ScaleWidth Step 32
    '    frmRPG.Line (x, 0)-(x, frmRPG.ScaleHeight)
    'Next
    
    'For y = 0 To frmRPG.ScaleHeight Step 32
    '    frmRPG.Line (0, y)-(frmRPG.ScaleWidth, y)
    'Next
        redo = False
    
    
    Open App.Path & "\maps\frame" & Format(intColmap, "00") & Format(intRowMap, "00") & ".txt" For Input As 1
    For intCounter = 1 To 10
    Line Input #1, strLine
    For intTile = 1 To 10
    intTile2 = CInt(Mid(strLine, (intTile * 2) - 1, 2))
    Select Case intTile2
    
        Case 1
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMain.Hdc, 0, 0, SRCCOPY)
        Case 2
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMain.Hdc, 32, 0, SRCCOPY)
        Case 3
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMain.Hdc, 64, 0, SRCCOPY)
        Case 4
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMain.Hdc, 0, 32, SRCCOPY)
        Case 5
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMain.Hdc, 32, 32, SRCCOPY)
        Case 6
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMain.Hdc, 64, 32, SRCCOPY)
        Case 7
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMain.Hdc, 0, 64, SRCCOPY)
        Case 8
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMain.Hdc, 32, 64, SRCCOPY)
        Case 9
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMain.Hdc, 64, 64, SRCCOPY)
        Case 10
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMain.Hdc, 0, 96, SRCCOPY)
        Case 11
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMain.Hdc, 32, 96, SRCCOPY)
        Case 12
            Level(intTile, intCounter) = 0
        Case 13
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 0, 0, SRCCOPY)
        Case 14
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 32, 0, SRCCOPY)
        Case 15
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 64, 0, SRCCOPY)
        Case 16
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 0, 32, SRCCOPY)
        Case 17
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 32, 32, SRCCOPY)
        Case 18
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 64, 32, SRCCOPY)
        Case 19
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 0, 64, SRCCOPY)
        Case 20
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 32, 64, SRCCOPY)
        Case 21
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 64, 64, SRCCOPY)
        Case 22
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 0, 96, SRCCOPY)
        Case 23
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 32, 96, SRCCOPY)
        Case 24
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picField.Hdc, 64, 96, SRCCOPY)
        Case 25
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 0, 0, SRCCOPY)
        Case 26
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 32, 0, SRCCOPY)
        Case 27
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 64, 0, SRCCOPY)
        Case 28
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 0, 32, SRCCOPY)
        Case 29
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 32, 32, SRCCOPY)
        Case 30
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 64, 32, SRCCOPY)
        Case 31
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 0, 64, SRCCOPY)
        Case 32
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 32, 64, SRCCOPY)
        Case 33
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 64, 64, SRCCOPY)
        Case 34
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 0, 96, SRCCOPY)
        Case 35
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 32, 96, SRCCOPY)
        Case 36
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picBridge.Hdc, 64, 96, SRCCOPY)
        Case 37
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 0, 0, SRCCOPY)
        Case 38
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 32, 0, SRCCOPY)
        Case 39
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 64, 0, SRCCOPY)
        Case 40
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 0, 32, SRCCOPY)
        Case 41
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 32, 32, SRCCOPY)
        Case 42
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 64, 32, SRCCOPY)
        Case 43
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 0, 64, SRCCOPY)
        Case 44
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 32, 64, SRCCOPY)
        Case 45
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 64, 64, SRCCOPY)
        Case 46
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 0, 96, SRCCOPY)
        Case 47
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 32, 96, SRCCOPY)
        Case 48
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picwater.Hdc, 64, 96, SRCCOPY)
        Case 49
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 0, 0, SRCCOPY)
        Case 50
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 32, 0, SRCCOPY)
        Case 51
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 64, 0, SRCCOPY)
        Case 52
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 0, 32, SRCCOPY)
        Case 53
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 32, 32, SRCCOPY)
        Case 54
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 64, 32, SRCCOPY)
        Case 55
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 0, 64, SRCCOPY)
        Case 56
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 32, 64, SRCCOPY)
        Case 57
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 64, 64, SRCCOPY)
        Case 58
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 0, 96, SRCCOPY)
        Case 59
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 32, 96, SRCCOPY)
        Case 60
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picMisc.Hdc, 64, 96, SRCCOPY)
        Case 61
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 0, 0, SRCCOPY)
        Case 62
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 32, 0, SRCCOPY)
        Case 63
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 64, 0, SRCCOPY)
        Case 64
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 0, 32, SRCCOPY)
        Case 65
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 32, 32, SRCCOPY)
        Case 66
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 64, 32, SRCCOPY)
        Case 67
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 0, 64, SRCCOPY)
        Case 68
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 32, 64, SRCCOPY)
        Case 69
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 64, 64, SRCCOPY)
        Case 70
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 0, 96, SRCCOPY)
        Case 71
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 32, 96, SRCCOPY)
        Case 72
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCity.Hdc, 64, 96, SRCCOPY)
        Case 73
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 0, 0, SRCCOPY)
        Case 74
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 32, 0, SRCCOPY)
        Case 75
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 64, 0, SRCCOPY)
        Case 76
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 0, 32, SRCCOPY)
        Case 77
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 32, 32, SRCCOPY)
        Case 78
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 64, 32, SRCCOPY)
        Case 79
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 0, 64, SRCCOPY)
        Case 80
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 32, 64, SRCCOPY)
        Case 81
            Level(intTile, intCounter) = 1
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 64, 64, SRCCOPY)
        Case 82
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 0, 96, SRCCOPY)
        Case 83
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 32, 96, SRCCOPY)
        Case 84
            Level(intTile, intCounter) = 0
            Call BitBlt(frmRPG.Hdc, ((intTile * 2 * 32) / 2) - 32, intCounter * 32 - 32, 32, 32, frmGraphics.picCliff.Hdc, 64, 96, SRCCOPY)
  
    End Select
    Next
    Next
    Close #1
    Call BitBlt(frmRPG.Hdc, guy.Left, guy.Top, 32, 32, picMask.Hdc, xleft, ytop, SRCAND)
    Call BitBlt(frmRPG.Hdc, guy.Left, guy.Top, 32, 32, picSprite.Hdc, xleft, ytop, SRCINVERT)
    Open App.Path & "\events\e" & Format(intColmap, "00") & Format(intRowMap, "00") & ".txt" For Input As #1
 Do
  Line Input #1, strLine
  If Left(strLine, 1) = "*" Then
   If modRPG.swa(Int(Mid(strLine, 10, 3))) = Int(Mid(strLine, 14, 3)) Then
    x = (Mid(strLine, 3, 1) + 1)
    y = (Mid(strLine, 5, 1) + 1)
    Call BitBlt(frmRPG.Hdc, (Mid(strLine, 3, 1) + 1) * 32 - 32, (Mid(strLine, 5, 1) + 1) * 32 - 32, 32, 32, mskppl.Hdc, (Mid(strLine, 19, 1)) * 32 - 32, (Mid(strLine, 20, 1)) * 32 - 32, SRCAND)
    Call BitBlt(frmRPG.Hdc, (Mid(strLine, 3, 1) + 1) * 32 - 32, (Mid(strLine, 5, 1) + 1) * 32 - 32, 32, 32, sprppl.Hdc, (Mid(strLine, 19, 1)) * 32 - 32, (Mid(strLine, 20, 1)) * 32 - 32, SRCINVERT)
    If Mid(strLine, 8, 1) = 0 Then
     Level(x, y) = 0
    End If
    Select Case Mid(strLine, 22, 1)
     Case "B"
      If ypos - 1 = y And xpos = x Then blah = True Else blah = False
     Case "T"
      If ypos + 1 = y And xpos = x Then blah = True Else blah = False
     Case "S"
      blah = True
     Case "L"
      If xpos - 1 = x And ypos = y Then blah = True Else blah = False
     Case "R"
      If xpos + 1 = x And ypos = y Then blah = True Else blah = False
    End Select
            
    If ((abc = True And blah = True) Or Mid(strLine, 22, 1) = "S") Then
     Do
      Line Input #1, strLine
      If Mid(strLine, 3, 1) = "F" Then
       MsgBox "There is an if statement: " & vbCrLf & strLine
       BOO = False
       Select Case Mid(strLine, 5, 3)
       Case "$$$"
        Select Case Mid(strLine, 9, 2)
            Case ">>"
             If cash > Mid(strLine, 12, 3) Then BOO = True
            Case "<<"
             If cash < Mid(strLine, 12, 3) Then BOO = True
            Case "=="
             If cash = Mid(strLine, 12, 3) Then BOO = True
            Case ">="
             If cash >= Mid(strLine, 12, 3) Then BOO = True
            Case "<="
             If cash <= Mid(strLine, 12, 3) Then BOO = True
        End Select
       Case "I**"
        Select Case Mid(strLine, 9, 2)
            Case ">>"
             If modRPG.intItemsOwn(Mid(strLine, 6, 2)) > Mid(strLine, 12, 3) Then BOO = True
            Case "<<"
             If modRPG.intItemsOwn(Mid(strLine, 6, 2)) < Mid(strLine, 12, 3) Then BOO = True
            Case "=="
             If modRPG.intItemsOwn(Mid(strLine, 6, 2)) = Mid(strLine, 12, 3) Then BOO = True
            Case ">="
             If modRPG.intItemsOwn(Mid(strLine, 6, 2)) >= Mid(strLine, 12, 3) Then BOO = True
            Case "<="
             If modRPG.intItemsOwn(Mid(strLine, 6, 2)) <= Mid(strLine, 12, 3) Then BOO = True
        End Select
       Case Else
        Select Case Mid(strLine, 9, 2)
            Case ">>"
             If swa(Mid(strLine, 5, 3)) > Mid(strLine, 12, 3) Then BOO = True
            Case "<<"
             If swa(Mid(strLine, 5, 3)) < Mid(strLine, 12, 3) Then BOO = True
            Case "=="
             If swa(Mid(strLine, 5, 3)) = Mid(strLine, 12, 3) Then BOO = True
            Case ">="
             If swa(Mid(strLine, 5, 3)) >= Mid(strLine, 12, 3) Then BOO = True
            Case "<="
             If swa(Mid(strLine, 5, 3)) <= Mid(strLine, 12, 3) Then BOO = True
        End Select
       End Select
       If BOO = True Then
        Do
         Line Input #1, strLine
         If Mid(strLine, 3, 1) = "M" Then
        def = False
        Who = Mid(strLine, 6)
        lblMessage = ""
        For a = 1 To Int(Mid(strLine, 4, 1))
         Line Input #1, strLine
         lblMessage = lblMessage & strLine & vbCrLf
        Next
        Message.Refresh
        If Message.Visible = False Then
         Message.Visible = True
         For a = frmRPG.ScaleHeight To frmRPG.ScaleHeight - Message.Height Step -2
          Message.Top = a
          For b = 0 To 5000
           DoEvents
          Next
         Next
        End If
        Do
         DoEvents
        Loop Until def = True
        strLine = ""
       End If
       If Mid(strLine, 3, 1) = "V" Then
        modRPG.swa(Mid(strLine, 5, 3)) = Mid(strLine, 10)
        redo = True
       End If
       If Mid(strLine, 3, 1) = "C" Then
        For a = Message.Top To frmRPG.ScaleHeight Step 2
         Message.Top = a
         For b = 0 To 5000
          DoEvents
         Next
        Next
        Message.Visible = False
       End If
       If Mid(strLine, 3, 1) = "T" Then
        redo = True
        intRowMap = Mid(strLine, 9, 3)
        intColmap = Mid(strLine, 5, 3)
        xpos = Mid(strLine, 13, 1) + 1
        ypos = Mid(strLine, 15, 1) + 1
       End If
       If Mid(strLine, 3, 1) = "B" Then
        frmBattle.sprDragon = frmGraphics.sprGuy(Int(Mid(strLine, 5, 3))).Picture
        frmBattle.mskDragon = frmGraphics.mskGuy(Int(Mid(strLine, 5, 3))).Picture
        modRPG.Eanim = Int(Mid(strLine, 21, 1))
        modRPG.edef = Int(Mid(strLine, 13, 3))
        modRPG.Estr = Int(Mid(strLine, 9, 3))
        modRPG.EHP = Int(Mid(strLine, 17, 3))
        modRPG.eMaxHP = modRPG.EHP
        frmBattle.Label8.Caption = Mid(strLine, 27)
        frmBattle.Show 1
        Do
        Loop Until frmBattle.Visible = False
        modRPG.cash = modRPG.cash + Int(Mid(strLine, 23, 3))
       End If
       If Mid(strLine, 3, 1) = "H" Then
        If cash >= Int(Mid(strLine, 5, 4)) Then
         cash = cash - Int(Mid(strLine, 5, 4))
         HP = MaxHP
         MP = MaxMP
        End If
       End If
       If Mid(strLine, 3, 1) = "S" Then
        frmShop.Show 1
        Do
        Loop Until frmShop.Visible = False
       End If
       If Mid(strLine, 3, 1) = "I" Then
        intItemsOwn(Int(Mid(strLine, 5, 3))) = intItemsOwn(Int(Mid(strLine, 5, 3))) + Int(Mid(strLine, 9, 3))
       End If
       If Mid(strLine, 3, 1) = "P" Then
        lblQuestion = Mid(strLine, 6)
        Picture1.Visible = True
        Picture1.Refresh
        Picture1.SetFocus
        Do
         DoEvents
        Loop Until Picture1.Visible = False
        If yes = -1 Then
         Do
          Line Input #1, strLine
          If Mid(strLine, 3, 1) = "M" Then
           def = False
           Who = Mid(strLine, 5)
           lblMessage = ""
           For a = 1 To Int(Mid(strLine, 4, 1))
            Line Input #1, strLine
            lblMessage = lblMessage & strLine & vbCrLf
           Next
           Message.Refresh
           If Message.Visible = False Then
            Message.Visible = True
            For a = frmRPG.ScaleHeight To frmRPG.ScaleHeight - Message.Height Step -2
             Message.Top = a
             For b = 0 To 5000
              DoEvents
             Next
            Next
           End If
           Do
            DoEvents
           Loop Until def = True
           strLine = ""
          End If
          If Mid(strLine, 3, 1) = "V" Then
           modRPG.swa(Mid(strLine, 5, 3)) = Mid(strLine, 10)
           redo = True
          End If
          If Mid(strLine, 3, 1) = "C" Then
           For a = Message.Top To frmRPG.ScaleHeight Step 2
            Message.Top = a
            For b = 0 To 5000
             DoEvents
            Next
           Next
           Message.Visible = False
          End If
          If Mid(strLine, 3, 1) = "T" Then
           redo = True
           intRowMap = Mid(strLine, 9, 3)
           intColmap = Mid(strLine, 5, 3)
           xpos = Mid(strLine, 13, 1) + 1
           ypos = Mid(strLine, 15, 1) + 1
          End If
          If Mid(strLine, 3, 1) = "B" Then
           frmBattle.sprDragon = frmGraphics.sprGuy(Int(Mid(strLine, 5, 3))).Picture
           frmBattle.mskDragon = frmGraphics.mskGuy(Int(Mid(strLine, 5, 3))).Picture
           modRPG.Eanim = Int(Mid(strLine, 21, 1))
           modRPG.edef = Int(Mid(strLine, 13, 3))
           modRPG.Estr = Int(Mid(strLine, 9, 3))
           modRPG.EHP = Int(Mid(strLine, 17, 3))
           modRPG.eMaxHP = modRPG.EHP
           frmBattle.Label8.Caption = Mid(strLine, 27)
           frmBattle.Show 1
           Do
           Loop Until frmBattle.Visible = False
           modRPG.cash = modRPG.cash + Int(Mid(strLine, 23, 3))
          End If
          If Mid(strLine, 3, 1) = "H" Then
           If cash >= Int(Mid(strLine, 5, 4)) Then
            cash = cash - Int(Mid(strLine, 5, 4))
            HP = MaxHP
            MP = MaxMP
           End If
          End If
          If Mid(strLine, 3, 1) = "S" Then
           frmShop.Show 1
           Do
           Loop Until frmShop.Visible = False
          End If
          If Mid(strLine, 3, 1) = "I" Then
           intItemsOwn(Int(Mid(strLine, 5, 3))) = intItemsOwn(Int(Mid(strLine, 5, 3))) + Int(Mid(strLine, 9, 3))
          End If
         Loop Until strLine = "==N"
         Do
          Line Input #1, strLine
         Loop Until strLine = "==Q"
        ElseIf yes = False Then
         Do
          Line Input #1, strLine
         Loop Until strLine = "==N"
         Do
          Line Input #1, strLine
          If Mid(strLine, 3, 1) = "M" Then
           def = False
           Who = Mid(strLine, 6)
           lblMessage = ""
           For a = 1 To Int(Mid(strLine, 4, 1))
            Line Input #1, strLine
            lblMessage = lblMessage & strLine & vbCrLf
           Next
           Message.Refresh
           If Message.Visible = False Then
            Message.Visible = True
            For a = frmRPG.ScaleHeight To frmRPG.ScaleHeight - Message.Height Step -2
             Message.Top = a
             For b = 0 To 5000
              DoEvents
             Next
            Next
           End If
           Do
            DoEvents
           Loop Until def = True
           strLine = ""
          End If
          If Mid(strLine, 3, 1) = "V" Then
           modRPG.swa(Mid(strLine, 5, 3)) = Mid(strLine, 10)
           redo = True
          End If
          If Mid(strLine, 3, 1) = "C" Then
           For a = Message.Top To frmRPG.ScaleHeight Step 2
            Message.Top = a
            For b = 0 To 5000
             DoEvents
            Next
           Next
           Message.Visible = False
          End If
          If Mid(strLine, 3, 1) = "T" Then
           redo = True
           intRowMap = Mid(strLine, 9, 3)
           intColmap = Mid(strLine, 5, 3)
           xpos = Mid(strLine, 13, 1) + 1
           ypos = Mid(strLine, 15, 1) + 1
          End If
          If Mid(strLine, 3, 1) = "B" Then
           frmBattle.sprDragon = frmGraphics.sprGuy(Int(Mid(strLine, 5, 3))).Picture
           frmBattle.mskDragon = frmGraphics.mskGuy(Int(Mid(strLine, 5, 3))).Picture
           modRPG.Eanim = Int(Mid(strLine, 21, 1))
           modRPG.edef = Int(Mid(strLine, 13, 3))
           modRPG.Estr = Int(Mid(strLine, 9, 3))
           modRPG.EHP = Int(Mid(strLine, 17, 3))
           modRPG.eMaxHP = modRPG.EHP
           frmBattle.Label8.Caption = Mid(strLine, 27)
           frmBattle.Show 1
           Do
           Loop Until frmBattle.Visible = False
           modRPG.cash = modRPG.cash + Int(Mid(strLine, 23, 3))
          End If
          If Mid(strLine, 3, 1) = "H" Then
           If cash >= Int(Mid(strLine, 5, 4)) Then
            cash = cash - Int(Mid(strLine, 5, 4))
            HP = MaxHP
            MP = MaxMP
           End If
          End If
          If Mid(strLine, 3, 1) = "S" Then
           frmShop.Show 1
           Do
           Loop Until frmShop.Visible = False
          End If
          If Mid(strLine, 3, 1) = "I" Then
           intItemsOwn(Int(Mid(strLine, 5, 3))) = intItemsOwn(Int(Mid(strLine, 5, 3))) + Int(Mid(strLine, 9, 3))
          End If
         Loop Until strLine = "==Q"
        End If
       End If
        Loop Until strLine = "==E"
       Else
        Do
         Line Input #1, strLine
        Loop Until strLine = "==E"
       End If
      End If
       If Mid(strLine, 3, 1) = "M" Then
        def = False
        Who = Mid(strLine, 6)
        lblMessage = ""
        For a = 1 To Int(Mid(strLine, 4, 1))
         Line Input #1, strLine
         lblMessage = lblMessage & strLine & vbCrLf
        Next
        Message.Refresh
        If Message.Visible = False Then
         Message.Visible = True
         For a = frmRPG.ScaleHeight To frmRPG.ScaleHeight - Message.Height Step -2
          Message.Top = a
          For b = 0 To 5000
           DoEvents
          Next
         Next
        End If
        Do
         DoEvents
        Loop Until def = True
        strLine = ""
       End If
       If Mid(strLine, 3, 1) = "V" Then
        modRPG.swa(Mid(strLine, 5, 3)) = Mid(strLine, 10)
        redo = True
       End If
       If Mid(strLine, 3, 1) = "C" Then
        For a = Message.Top To frmRPG.ScaleHeight Step 2
         Message.Top = a
         For b = 0 To 5000
          DoEvents
         Next
        Next
        Message.Visible = False
       End If
       If Mid(strLine, 3, 1) = "T" Then
        redo = True
        intRowMap = Mid(strLine, 9, 3)
        intColmap = Mid(strLine, 5, 3)
        xpos = Mid(strLine, 13, 1) + 1
        ypos = Mid(strLine, 15, 1) + 1
       End If
       If Mid(strLine, 3, 1) = "B" Then
        frmBattle.sprDragon = frmGraphics.sprGuy(Int(Mid(strLine, 5, 3))).Picture
        frmBattle.mskDragon = frmGraphics.mskGuy(Int(Mid(strLine, 5, 3))).Picture
        modRPG.Eanim = Int(Mid(strLine, 21, 1))
        modRPG.edef = Int(Mid(strLine, 13, 3))
        modRPG.Estr = Int(Mid(strLine, 9, 3))
        modRPG.EHP = Int(Mid(strLine, 17, 3))
        modRPG.eMaxHP = modRPG.EHP
        frmBattle.Label8.Caption = Mid(strLine, 27)
        frmBattle.Show 1
        Do
        Loop Until frmBattle.Visible = False
        modRPG.cash = modRPG.cash + Int(Mid(strLine, 23, 3))
       End If
       If Mid(strLine, 3, 1) = "H" Then
        If cash >= Int(Mid(strLine, 5, 4)) Then
         cash = cash - Int(Mid(strLine, 5, 4))
         HP = MaxHP
         MP = MaxMP
        End If
       End If
       If Mid(strLine, 3, 1) = "S" Then
        frmShop.Show 1
        Do
        Loop Until frmShop.Visible = False
       End If
       If Mid(strLine, 3, 1) = "I" Then
        intItemsOwn(Int(Mid(strLine, 5, 3))) = intItemsOwn(Int(Mid(strLine, 5, 3))) + Int(Mid(strLine, 9, 3))
       End If
       If Mid(strLine, 3, 1) = "P" Then
        lblQuestion = Mid(strLine, 6)
        Picture1.Visible = True
        Picture1.Refresh
        Picture1.SetFocus
        Do
         DoEvents
        Loop Until Picture1.Visible = False
        If yes = -1 Then
         Do
          Line Input #1, strLine
          If Mid(strLine, 3, 1) = "M" Then
           def = False
           Who = Mid(strLine, 6)
           lblMessage = ""
           For a = 1 To Int(Mid(strLine, 4, 1))
            Line Input #1, strLine
            lblMessage = lblMessage & strLine & vbCrLf
           Next
           Message.Refresh
           If Message.Visible = False Then
            Message.Visible = True
            For a = frmRPG.ScaleHeight To frmRPG.ScaleHeight - Message.Height Step -2
             Message.Top = a
             For b = 0 To 5000
              DoEvents
             Next
            Next
           End If
           Do
            DoEvents
           Loop Until def = True
           strLine = ""
          End If
          If Mid(strLine, 3, 1) = "V" Then
           modRPG.swa(Mid(strLine, 5, 3)) = Mid(strLine, 10)
           redo = True
          End If
          If Mid(strLine, 3, 1) = "C" Then
           For a = Message.Top To frmRPG.ScaleHeight Step 2
            Message.Top = a
            For b = 0 To 5000
             DoEvents
            Next
           Next
           Message.Visible = False
          End If
          If Mid(strLine, 3, 1) = "T" Then
           redo = True
           intRowMap = Mid(strLine, 9, 3)
           intColmap = Mid(strLine, 5, 3)
           xpos = Mid(strLine, 13, 1) + 1
           ypos = Mid(strLine, 15, 1) + 1
          End If
          If Mid(strLine, 3, 1) = "B" Then
           frmBattle.sprDragon = frmGraphics.sprGuy(Int(Mid(strLine, 5, 3))).Picture
           frmBattle.mskDragon = frmGraphics.mskGuy(Int(Mid(strLine, 5, 3))).Picture
           modRPG.Eanim = Int(Mid(strLine, 21, 1))
           modRPG.edef = Int(Mid(strLine, 13, 3))
           modRPG.Estr = Int(Mid(strLine, 9, 3))
           modRPG.EHP = Int(Mid(strLine, 17, 3))
           modRPG.eMaxHP = modRPG.EHP
           frmBattle.Label8.Caption = Mid(strLine, 27)
           frmBattle.Show 1
           Do
           Loop Until frmBattle.Visible = False
           modRPG.cash = modRPG.cash + Int(Mid(strLine, 23, 3))
          End If
          If Mid(strLine, 3, 1) = "H" Then
           If cash >= Int(Mid(strLine, 5, 4)) Then
            cash = cash - Int(Mid(strLine, 5, 4))
            HP = MaxHP
            MP = MaxMP
           End If
          End If
          If Mid(strLine, 3, 1) = "S" Then
           frmShop.Show 1
           Do
           Loop Until frmShop.Visible = False
          End If
          If Mid(strLine, 3, 1) = "I" Then
           intItemsOwn(Int(Mid(strLine, 5, 3))) = intItemsOwn(Int(Mid(strLine, 5, 3))) + Int(Mid(strLine, 9, 3))
          End If
         Loop Until strLine = "==N"
         Do
          Line Input #1, strLine
         Loop Until strLine = "==Q"
        ElseIf yes = False Then
         Do
          Line Input #1, strLine
         Loop Until strLine = "==N"
         Do
          Line Input #1, strLine
          If Mid(strLine, 3, 1) = "M" Then
           def = False
           Who = Mid(strLine, 6)
           lblMessage = ""
           For a = 1 To Int(Mid(strLine, 4, 1))
            Line Input #1, strLine
            lblMessage = lblMessage & strLine & vbCrLf
           Next
           Message.Refresh
           If Message.Visible = False Then
            Message.Visible = True
            For a = frmRPG.ScaleHeight To frmRPG.ScaleHeight - Message.Height Step -2
             Message.Top = a
             For b = 0 To 5000
              DoEvents
             Next
            Next
           End If
           Do
            DoEvents
           Loop Until def = True
           strLine = ""
          End If
          If Mid(strLine, 3, 1) = "V" Then
           modRPG.swa(Mid(strLine, 5, 3)) = Mid(strLine, 10)
           redo = True
          End If
          If Mid(strLine, 3, 1) = "C" Then
           For a = Message.Top To frmRPG.ScaleHeight Step 2
            Message.Top = a
            For b = 0 To 5000
             DoEvents
            Next
           Next
           Message.Visible = False
          End If
          If Mid(strLine, 3, 1) = "T" Then
           redo = True
           intRowMap = Mid(strLine, 9, 3)
           intColmap = Mid(strLine, 5, 3)
           xpos = Mid(strLine, 13, 1) + 1
           ypos = Mid(strLine, 15, 1) + 1
          End If
          If Mid(strLine, 3, 1) = "B" Then
           frmBattle.sprDragon = frmGraphics.sprGuy(Int(Mid(strLine, 5, 3))).Picture
           frmBattle.mskDragon = frmGraphics.mskGuy(Int(Mid(strLine, 5, 3))).Picture
           modRPG.Eanim = Int(Mid(strLine, 21, 1))
           modRPG.edef = Int(Mid(strLine, 13, 3))
           modRPG.Estr = Int(Mid(strLine, 9, 3))
           modRPG.EHP = Int(Mid(strLine, 17, 3))
           modRPG.eMaxHP = modRPG.EHP
           frmBattle.Label8.Caption = Mid(strLine, 27)
           frmBattle.Show 1
           Do
           Loop Until frmBattle.Visible = False
           modRPG.cash = modRPG.cash + Int(Mid(strLine, 23, 3))
          End If
          If Mid(strLine, 3, 1) = "H" Then
           If cash >= Int(Mid(strLine, 5, 4)) Then
            cash = cash - Int(Mid(strLine, 5, 4))
            HP = MaxHP
            MP = MaxMP
           End If
          End If
          If Mid(strLine, 3, 1) = "S" Then
           frmShop.Show 1
           Do
           Loop Until frmShop.Visible = False
          End If
          If Mid(strLine, 3, 1) = "I" Then
           intItemsOwn(Int(Mid(strLine, 5, 3))) = intItemsOwn(Int(Mid(strLine, 5, 3))) + Int(Mid(strLine, 9, 3))
          End If
         Loop Until strLine = "==Q"
        End If
       End If
    Loop Until strLine = "[EE]"
   End If
  End If
  End If
 Loop Until EOF(1) = True
    Close #1

    def = False
    Message.Visible = False
    abc = False
    frmRPG.Refresh
    If redo = True Then PaintScreen
End Sub

Private Sub Message_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    
    def = True

End If
End Sub

Private Sub Message_Paint()
    
    For a = 0 To 255
    
        Message.Line (0, a)-(Message.ScaleWidth, a), RGB(0, 0, 255 - a)
    
    Next
    
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            If shpChoose.Top = 96 Then shpChoose.Top = 168 Else shpChoose.Top = 96
        Case 40
            If shpChoose.Top = 96 Then shpChoose.Top = 168 Else shpChoose.Top = 96
        Case 32
            If shpChoose.Top = 96 Then
                yes = True
            Else
                yes = False
            End If
            Picture1.Visible = False
    End Select
End Sub

Private Sub Picture1_Paint()
    For a = 0 To 255
    
        Picture1.Line (0, a)-(Message.ScaleWidth, a), RGB(0, 0, 255 - a)
    
    Next
End Sub

Private Sub tmrNext_Timer()
    
    imgNext.Visible = Not (imgNext.Visible)
    
End Sub

Private Sub tmrRefresh_Timer()
PaintScreen
End Sub

