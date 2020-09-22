VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   0  'None
   Caption         =   "Invisible"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Label Label2 
      Caption         =   "This game was made in Comp Sci I at River Hill High School in Clarksville, MD. The programmer is Andrei Krotkov."
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Programmer's Notes:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmGraphics.picBridge.Picture = LoadPicture(App.Path & "/graph/Bridge.bmp")
    frmGraphics.picCliff.Picture = LoadPicture(App.Path & "/graph/Cliff.bmp")
    frmGraphics.picMain.Picture = LoadPicture(App.Path & "/graph/main.bmp")
    frmGraphics.picCity.Picture = LoadPicture(App.Path & "/graph/city.bmp")
    frmGraphics.picMisc.Picture = LoadPicture(App.Path & "/graph/misc.bmp")
    frmGraphics.picwater.Picture = LoadPicture(App.Path & "/graph/water.bmp")
    frmGraphics.picField.Picture = LoadPicture(App.Path & "/graph/field.bmp")
    Load frmBackground
    frmBackground.Show
    modRPG.cash = 1000
    ilevel = 1
    Flevel = 1
    hLevel = 1
    HP = 150
    MP = 20
    MaxHP = 150
    MaxMP = 20
    Str = 10
    def = 10
    Form1.Show
End Sub

