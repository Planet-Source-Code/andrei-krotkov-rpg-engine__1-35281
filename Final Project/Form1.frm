VERSION 5.00
Begin VB.Form frmRPG 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "RPG Game"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTestResult 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      Height          =   1920
      Left            =   2880
      Picture         =   "Form1.frx":57B2
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmRPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblMessage_Click()

End Sub
