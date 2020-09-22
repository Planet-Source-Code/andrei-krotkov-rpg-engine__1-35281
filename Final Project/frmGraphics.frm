VERSION 5.00
Begin VB.Form frmGraphics 
   Caption         =   "Graphics"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCliff 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   3240
      Picture         =   "frmGraphics.frx":0000
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picBridge 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   1680
      Picture         =   "frmGraphics.frx":9042
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   9
      Top             =   6120
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picCity 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   1680
      Picture         =   "frmGraphics.frx":12084
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picwater 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   120
      Picture         =   "frmGraphics.frx":1B0C6
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   120
      Picture         =   "frmGraphics.frx":24108
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picField 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   120
      Picture         =   "frmGraphics.frx":2D14A
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picMisc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   1680
      Picture         =   "frmGraphics.frx":3618C
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox mskGuy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1830
      Index           =   0
      Left            =   4200
      Picture         =   "frmGraphics.frx":3F1CE
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   186
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox mskGuy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1830
      Index           =   1
      Left            =   3960
      Picture         =   "frmGraphics.frx":4F430
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   186
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox sprGuy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1830
      Index           =   1
      Left            =   4440
      Picture         =   "frmGraphics.frx":5F692
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   186
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox sprGuy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1830
      Index           =   0
      Left            =   120
      Picture         =   "frmGraphics.frx":6F8F4
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   186
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2850
   End
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
