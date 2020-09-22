VERSION 5.00
Begin VB.Form frmLevelEdit 
   Caption         =   "Level Editor"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOPEN 
      Caption         =   "Open"
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   615
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   12
      Left            =   1320
      Picture         =   "frmLevelEdit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   11
      Left            =   720
      Picture         =   "frmLevelEdit.frx":0C42
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   495
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   10
      Left            =   120
      Picture         =   "frmLevelEdit.frx":1884
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   9
      Left            =   1320
      Picture         =   "frmLevelEdit.frx":24C6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   495
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   8
      Left            =   720
      Picture         =   "frmLevelEdit.frx":3108
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   7
      Left            =   120
      Picture         =   "frmLevelEdit.frx":3D4A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   6
      Left            =   1320
      Picture         =   "frmLevelEdit.frx":498C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   5
      Left            =   720
      Picture         =   "frmLevelEdit.frx":55CE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   4
      Left            =   120
      Picture         =   "frmLevelEdit.frx":6210
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   3
      Left            =   1320
      Picture         =   "frmLevelEdit.frx":6E52
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   2
      Left            =   720
      Picture         =   "frmLevelEdit.frx":7A94
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.CheckBox Tile 
      Height          =   495
      Index           =   1
      Left            =   120
      Picture         =   "frmLevelEdit.frx":86D6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   12
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":9318
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   11
      Left            =   3240
      Picture         =   "frmLevelEdit.frx":9F5A
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   10
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":AB9C
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   9
      Left            =   3000
      Picture         =   "frmLevelEdit.frx":B7DE
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   8
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":C420
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   7
      Left            =   2760
      Picture         =   "frmLevelEdit.frx":D062
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   6
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":DCA4
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   5
      Left            =   2520
      Picture         =   "frmLevelEdit.frx":E8E6
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   4
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":F528
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "frmLevelEdit.frx":1016A
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   2
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":10DAC
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Cliff 
      Height          =   480
      Index           =   1
      Left            =   2040
      Picture         =   "frmLevelEdit.frx":119EE
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   12
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":12630
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   11
      Left            =   3240
      Picture         =   "frmLevelEdit.frx":13272
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   10
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":13EB4
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   9
      Left            =   3000
      Picture         =   "frmLevelEdit.frx":14AF6
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   8
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":15738
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   7
      Left            =   2760
      Picture         =   "frmLevelEdit.frx":1637A
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   6
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":16FBC
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   5
      Left            =   2520
      Picture         =   "frmLevelEdit.frx":17BFE
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   4
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":18840
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "frmLevelEdit.frx":19482
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   2
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":1A0C4
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bridge 
      Height          =   480
      Index           =   1
      Left            =   2040
      Picture         =   "frmLevelEdit.frx":1AD06
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   12
      Left            =   7680
      Picture         =   "frmLevelEdit.frx":1B948
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   11
      Left            =   7560
      Picture         =   "frmLevelEdit.frx":1C58A
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   10
      Left            =   7440
      Picture         =   "frmLevelEdit.frx":1D1CC
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   9
      Left            =   7320
      Picture         =   "frmLevelEdit.frx":1DE0E
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   8
      Left            =   7200
      Picture         =   "frmLevelEdit.frx":1EA50
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   7
      Left            =   7080
      Picture         =   "frmLevelEdit.frx":1F692
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   6
      Left            =   6960
      Picture         =   "frmLevelEdit.frx":202D4
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   5
      Left            =   6840
      Picture         =   "frmLevelEdit.frx":20F16
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   4
      Left            =   6720
      Picture         =   "frmLevelEdit.frx":21B58
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   3
      Left            =   6600
      Picture         =   "frmLevelEdit.frx":2279A
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   2
      Left            =   6480
      Picture         =   "frmLevelEdit.frx":233DC
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image City 
      Height          =   480
      Index           =   1
      Left            =   6360
      Picture         =   "frmLevelEdit.frx":2401E
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   1080
      Left            =   7320
      Picture         =   "frmLevelEdit.frx":24C60
      Stretch         =   -1  'True
      Top             =   120
      Width           =   840
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   6360
      Picture         =   "frmLevelEdit.frx":2DCA2
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   6360
      Picture         =   "frmLevelEdit.frx":36CE4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   840
   End
   Begin VB.Line Line4 
      Index           =   10
      X1              =   128
      X2              =   288
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Line Line4 
      Index           =   9
      X1              =   128
      X2              =   288
      Y1              =   152
      Y2              =   152
   End
   Begin VB.Line Line4 
      Index           =   8
      X1              =   128
      X2              =   288
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line Line4 
      Index           =   7
      X1              =   128
      X2              =   288
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line4 
      Index           =   6
      X1              =   128
      X2              =   288
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Line Line4 
      Index           =   5
      X1              =   128
      X2              =   288
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Line Line4 
      Index           =   4
      X1              =   128
      X2              =   288
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Line Line4 
      Index           =   3
      X1              =   128
      X2              =   288
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   128
      X2              =   288
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   128
      X2              =   288
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   128
      X2              =   288
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line Line3 
      Index           =   9
      X1              =   288
      X2              =   288
      Y1              =   8
      Y2              =   168
   End
   Begin VB.Line Line3 
      Index           =   8
      X1              =   272
      X2              =   272
      Y1              =   8
      Y2              =   168
   End
   Begin VB.Line Line3 
      Index           =   7
      X1              =   256
      X2              =   256
      Y1              =   8
      Y2              =   168
   End
   Begin VB.Line Line3 
      Index           =   6
      X1              =   128
      X2              =   128
      Y1              =   8
      Y2              =   168
   End
   Begin VB.Line Line3 
      Index           =   5
      X1              =   240
      X2              =   240
      Y1              =   8
      Y2              =   168
   End
   Begin VB.Line Line3 
      Index           =   4
      X1              =   224
      X2              =   224
      Y1              =   8
      Y2              =   168
   End
   Begin VB.Line Line3 
      Index           =   3
      X1              =   208
      X2              =   208
      Y1              =   8
      Y2              =   168
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   192
      X2              =   192
      Y1              =   8
      Y2              =   168
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   176
      X2              =   176
      Y1              =   8
      Y2              =   168
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   160
      X2              =   160
      Y1              =   8
      Y2              =   168
   End
   Begin VB.Line Line2 
      X1              =   144
      X2              =   144
      Y1              =   168
      Y2              =   8
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   12
      Left            =   5760
      Picture         =   "frmLevelEdit.frx":3FD26
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   11
      Left            =   5640
      Picture         =   "frmLevelEdit.frx":40968
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   10
      Left            =   5520
      Picture         =   "frmLevelEdit.frx":415AA
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   9
      Left            =   5400
      Picture         =   "frmLevelEdit.frx":421EC
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   8
      Left            =   5280
      Picture         =   "frmLevelEdit.frx":42E2E
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   7
      Left            =   5160
      Picture         =   "frmLevelEdit.frx":43A70
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   6
      Left            =   5040
      Picture         =   "frmLevelEdit.frx":446B2
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   5
      Left            =   4920
      Picture         =   "frmLevelEdit.frx":452F4
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   4
      Left            =   4800
      Picture         =   "frmLevelEdit.frx":45F36
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   3
      Left            =   4680
      Picture         =   "frmLevelEdit.frx":46B78
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   2
      Left            =   4560
      Picture         =   "frmLevelEdit.frx":477BA
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image WAter 
      Height          =   480
      Index           =   1
      Left            =   4440
      Picture         =   "frmLevelEdit.frx":483FC
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   99
      Left            =   4080
      Picture         =   "frmLevelEdit.frx":4903E
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   98
      Left            =   3840
      Picture         =   "frmLevelEdit.frx":49C80
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   97
      Left            =   3600
      Picture         =   "frmLevelEdit.frx":4A8C2
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   96
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":4B504
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   95
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":4C146
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   94
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":4CD88
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   93
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":4D9CA
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   92
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":4E60C
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   91
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":4F24E
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   90
      Left            =   1920
      Picture         =   "frmLevelEdit.frx":4FE90
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   89
      Left            =   4080
      Picture         =   "frmLevelEdit.frx":50AD2
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   88
      Left            =   3840
      Picture         =   "frmLevelEdit.frx":51714
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   87
      Left            =   3600
      Picture         =   "frmLevelEdit.frx":52356
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   86
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":52F98
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   85
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":53BDA
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   84
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":5481C
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   83
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":5545E
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   82
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":560A0
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   81
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":56CE2
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   80
      Left            =   1920
      Picture         =   "frmLevelEdit.frx":57924
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   79
      Left            =   4080
      Picture         =   "frmLevelEdit.frx":58566
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   78
      Left            =   3840
      Picture         =   "frmLevelEdit.frx":591A8
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   77
      Left            =   3600
      Picture         =   "frmLevelEdit.frx":59DEA
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   76
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":5AA2C
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   75
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":5B66E
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   74
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":5C2B0
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   73
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":5CEF2
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   72
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":5DB34
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   71
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":5E776
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   70
      Left            =   1920
      Picture         =   "frmLevelEdit.frx":5F3B8
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   69
      Left            =   4080
      Picture         =   "frmLevelEdit.frx":5FFFA
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   68
      Left            =   3840
      Picture         =   "frmLevelEdit.frx":60C3C
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   67
      Left            =   3600
      Picture         =   "frmLevelEdit.frx":6187E
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   66
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":624C0
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   65
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":63102
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   64
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":63D44
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   63
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":64986
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   62
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":655C8
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   61
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":6620A
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   60
      Left            =   1920
      Picture         =   "frmLevelEdit.frx":66E4C
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   59
      Left            =   4080
      Picture         =   "frmLevelEdit.frx":67A8E
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   58
      Left            =   3840
      Picture         =   "frmLevelEdit.frx":686D0
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   57
      Left            =   3600
      Picture         =   "frmLevelEdit.frx":69312
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   56
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":69F54
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   55
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":6AB96
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   54
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":6B7D8
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   53
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":6C41A
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   52
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":6D05C
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   51
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":6DC9E
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   50
      Left            =   1920
      Picture         =   "frmLevelEdit.frx":6E8E0
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   49
      Left            =   4080
      Picture         =   "frmLevelEdit.frx":6F522
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   48
      Left            =   3840
      Picture         =   "frmLevelEdit.frx":70164
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   47
      Left            =   3600
      Picture         =   "frmLevelEdit.frx":70DA6
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   46
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":719E8
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   45
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":7262A
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   44
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":7326C
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   43
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":73EAE
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   42
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":74AF0
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   41
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":75732
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   40
      Left            =   1920
      Picture         =   "frmLevelEdit.frx":76374
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   39
      Left            =   4080
      Picture         =   "frmLevelEdit.frx":76FB6
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   38
      Left            =   3840
      Picture         =   "frmLevelEdit.frx":77BF8
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   37
      Left            =   3600
      Picture         =   "frmLevelEdit.frx":7883A
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   36
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":7947C
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   35
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":7A0BE
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   34
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":7AD00
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   33
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":7B942
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   32
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":7C584
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   31
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":7D1C6
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   30
      Left            =   1920
      Picture         =   "frmLevelEdit.frx":7DE08
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   29
      Left            =   4080
      Picture         =   "frmLevelEdit.frx":7EA4A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   28
      Left            =   3840
      Picture         =   "frmLevelEdit.frx":7F68C
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   27
      Left            =   3600
      Picture         =   "frmLevelEdit.frx":802CE
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   26
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":80F10
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   25
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":81B52
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   24
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":82794
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   23
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":833D6
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   22
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":84018
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   21
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":84C5A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   20
      Left            =   1920
      Picture         =   "frmLevelEdit.frx":8589C
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   19
      Left            =   4080
      Picture         =   "frmLevelEdit.frx":864DE
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   18
      Left            =   3840
      Picture         =   "frmLevelEdit.frx":87120
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   17
      Left            =   3600
      Picture         =   "frmLevelEdit.frx":87D62
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   16
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":889A4
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   15
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":895E6
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   14
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":8A228
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   13
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":8AE6A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   12
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":8BAAC
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   11
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":8C6EE
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   10
      Left            =   1920
      Picture         =   "frmLevelEdit.frx":8D330
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   9
      Left            =   4080
      Picture         =   "frmLevelEdit.frx":8DF72
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   8
      Left            =   3840
      Picture         =   "frmLevelEdit.frx":8EBB4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   7
      Left            =   3600
      Picture         =   "frmLevelEdit.frx":8F7F6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   6
      Left            =   3360
      Picture         =   "frmLevelEdit.frx":90438
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   5
      Left            =   3120
      Picture         =   "frmLevelEdit.frx":9107A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   4
      Left            =   2880
      Picture         =   "frmLevelEdit.frx":91CBC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   3
      Left            =   2640
      Picture         =   "frmLevelEdit.frx":928FE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   2
      Left            =   2400
      Picture         =   "frmLevelEdit.frx":93540
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   1
      Left            =   2160
      Picture         =   "frmLevelEdit.frx":94182
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgField 
      Height          =   1080
      Left            =   4440
      Picture         =   "frmLevelEdit.frx":94DC4
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   840
   End
   Begin VB.Image imgWater 
      Height          =   1080
      Left            =   5400
      Picture         =   "frmLevelEdit.frx":9DE06
      Stretch         =   -1  'True
      Top             =   120
      Width           =   840
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   12
      Left            =   5760
      Picture         =   "frmLevelEdit.frx":A6E48
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   11
      Left            =   5640
      Picture         =   "frmLevelEdit.frx":A7A8A
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   10
      Left            =   5520
      Picture         =   "frmLevelEdit.frx":A86CC
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   9
      Left            =   5400
      Picture         =   "frmLevelEdit.frx":A930E
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   8
      Left            =   5280
      Picture         =   "frmLevelEdit.frx":A9F50
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   7
      Left            =   5160
      Picture         =   "frmLevelEdit.frx":AAB92
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   6
      Left            =   5040
      Picture         =   "frmLevelEdit.frx":AB7D4
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   5
      Left            =   4920
      Picture         =   "frmLevelEdit.frx":AC416
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   4
      Left            =   4800
      Picture         =   "frmLevelEdit.frx":AD058
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   3
      Left            =   4680
      Picture         =   "frmLevelEdit.frx":ADC9A
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   2
      Left            =   4560
      Picture         =   "frmLevelEdit.frx":AE8DC
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Main 
      Height          =   480
      Index           =   1
      Left            =   4440
      Picture         =   "frmLevelEdit.frx":AF51E
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   12
      Left            =   5760
      Picture         =   "frmLevelEdit.frx":B0160
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   11
      Left            =   5640
      Picture         =   "frmLevelEdit.frx":B0DA2
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   10
      Left            =   5520
      Picture         =   "frmLevelEdit.frx":B19E4
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   9
      Left            =   5400
      Picture         =   "frmLevelEdit.frx":B2626
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   8
      Left            =   5280
      Picture         =   "frmLevelEdit.frx":B3268
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   7
      Left            =   5160
      Picture         =   "frmLevelEdit.frx":B3EAA
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   6
      Left            =   5040
      Picture         =   "frmLevelEdit.frx":B4AEC
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   5
      Left            =   4920
      Picture         =   "frmLevelEdit.frx":B572E
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   4
      Left            =   4800
      Picture         =   "frmLevelEdit.frx":B6370
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   3
      Left            =   4680
      Picture         =   "frmLevelEdit.frx":B6FB2
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   2
      Left            =   4560
      Picture         =   "frmLevelEdit.frx":B7BF4
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image misc 
      Height          =   480
      Index           =   1
      Left            =   4440
      Picture         =   "frmLevelEdit.frx":B8836
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   12
      Left            =   5760
      Picture         =   "frmLevelEdit.frx":B9478
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   11
      Left            =   5640
      Picture         =   "frmLevelEdit.frx":BA0BA
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   10
      Left            =   5520
      Picture         =   "frmLevelEdit.frx":BACFC
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   9
      Left            =   5400
      Picture         =   "frmLevelEdit.frx":BB93E
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   8
      Left            =   5280
      Picture         =   "frmLevelEdit.frx":BC580
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   7
      Left            =   5160
      Picture         =   "frmLevelEdit.frx":BD1C2
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   6
      Left            =   5040
      Picture         =   "frmLevelEdit.frx":BDE04
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   5
      Left            =   4920
      Picture         =   "frmLevelEdit.frx":BEA46
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   4
      Left            =   4680
      Picture         =   "frmLevelEdit.frx":BF688
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   3
      Left            =   4560
      Picture         =   "frmLevelEdit.frx":C02CA
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   2
      Left            =   4680
      Picture         =   "frmLevelEdit.frx":C0F0C
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image field 
      Height          =   480
      Index           =   1
      Left            =   4440
      Picture         =   "frmLevelEdit.frx":C1B4E
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMisc 
      Height          =   1080
      Left            =   5400
      Picture         =   "frmLevelEdit.frx":C2790
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   840
   End
   Begin VB.Shape Border 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Height          =   495
      Left            =   120
      Top             =   240
      Width           =   495
   End
   Begin VB.Image imgMain 
      Height          =   1080
      Left            =   4440
      Picture         =   "frmLevelEdit.frx":CB7D2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   840
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   120
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Image Map 
      Height          =   255
      Index           =   0
      Left            =   1920
      Picture         =   "frmLevelEdit.frx":D4814
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmLevelEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tilechoose As Integer
Dim MapA(10, 10) As Integer
Dim Mapc As Integer

Private Sub cmdFill_Click()
a = MsgBox("Are you sure?", vbYesNo, "Fill")
If a = vbYes Then
For b = 0 To 99
Map(b).Picture = Tile(tilechoose).Picture
a = Format(b, "00")
x = Right(a, 1) + 1
y = Left(a, 1) + 1
c = tilechoose + (12 * Mapc)
MapA(x, y) = c
Next
End If
End Sub


Private Sub Command1_Click()

End Sub

Private Sub cmdOPEN_Click()
intNum = FreeFile
xrow = InputBox("Row?")
ycol = InputBox("Col?")
file = App.Path & "/maps/Frame" & Format(ycol, "00") & Format(xrow, "00") & ".txt"
Open file For Input As #intNum
For intCounter = 0 To 9
Line Input #intNum, strInput
For x = 1 To 19 Step 2
mmm = Mid(strInput, x, 2)
MapA((x + 1) / 2, intCounter + 1) = mmm
intA = (intCounter * 10) + ((x - 1) / 2)
Select Case Mid(strInput, x, 2)
    Case 1 To 12
        Map(intA).Picture = Main(mmm).Picture
    Case 13 To 24
        mmm = mmm - 12
        Map(intA).Picture = field(mmm).Picture
    Case 25 To 36
        mmm = mmm - 24
        Map(intA).Picture = Bridge(mmm).Picture
    Case 37 To 48
        mmm = mmm - 36
        Map(intA).Picture = WAter(mmm).Picture
    Case 49 To 60
        mmm = mmm - 48
        Map(intA).Picture = misc(mmm).Picture
    Case 61 To 72
        mmm = mmm - 60
        Map(intA).Picture = City(mmm).Picture
    Case 73 To 84
        mmm = mmm - 72
        Map(intA).Picture = Cliff(mmm).Picture
End Select
Next
Next
Error:
Close intNum
End Sub

Private Sub cmdSave_Click()
On Error GoTo Error
intNum = FreeFile
xrow = InputBox("Row?")
ycol = InputBox("Col?")
file = App.Path & "/maps/Frame" & Format(ycol, "00") & Format(xrow, "00") & ".txt"
Open file For Output As #intNum
For b = 1 To 10
For a = 1 To 10
strArr = strArr & Format(MapA(a, b), "00")
Next
Print #intNum, strArr
strArr = ""
Next
Close #intNum
MsgBox "Saved"
Exit Sub
Error:
MsgBox "An error has occured, Andrei..." & vbCrLf & Err.Description & vbCrLf & "Error number " & Err.Number, vbCritical
End Sub

Private Sub Form_Load()
Mapc = 0
tilechoose = 1
For z = 0 To 99
a = Format(z, "00")
x = Right(a, 1) + 1
y = Left(a, 1) + 1
MapA(x, y) = 1
Next
End Sub

Private Sub Image1_Click()
For a = 1 To 12
Tile(a).Picture = Bridge(a).Picture
Next
Mapc = 2
End Sub

Private Sub Image2_Click()
For a = 1 To 12
Tile(a).Picture = City(a).Picture
Next
Mapc = 5
End Sub

Private Sub Image3_Click()
For a = 1 To 12
Tile(a).Picture = Cliff(a).Picture
Next
Mapc = 6
End Sub

Private Sub imgField_Click()
For a = 1 To 12
Tile(a).Picture = field(a).Picture
Next
Mapc = 1
End Sub

Private Sub imgMain_Click()
For a = 1 To 12
Tile(a).Picture = Main(a).Picture
Next
Mapc = 0
End Sub


Private Sub imgMisc_Click()
For a = 1 To 12
Tile(a).Picture = misc(a).Picture
Next
Mapc = 4
End Sub


Private Sub imgWater_Click()
For a = 1 To 12
Tile(a).Picture = WAter(a).Picture
Next
Mapc = 3
End Sub

Private Sub Map_Click(Index As Integer)
On Error Resume Next
Map(Index).Picture = Tile(tilechoose).Picture
'NEED HELP OVER HERE!!!
a = Format(Index, "00")
x = Right(a, 1) + 1
y = Left(a, 1) + 1
b = tilechoose + (12 * Mapc)
MapA(x, y) = b


'END HELP
End Sub

Private Sub Tile_Click(Index As Integer)
On Error Resume Next
tilechoose = Index
Border.Left = Tile(Index).Left
Border.Top = Tile(Index).Top
For a = 1 To 12
Tile(a).Value = vbUnchecked
Next
End Sub

