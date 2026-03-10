VERSION 5.00
Begin VB.Form Intro 
   Caption         =   "Intro"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox pNegDinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   5880
      Picture         =   "Intro.frx":0000
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   18
      Top             =   6720
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pDinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   4800
      Picture         =   "Intro.frx":3AC2
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   17
      Top             =   6720
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pNegKinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   5880
      Picture         =   "Intro.frx":7584
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   16
      Top             =   8040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pKinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   4800
      Picture         =   "Intro.frx":B046
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   15
      Top             =   8040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pNegHinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   8400
      Picture         =   "Intro.frx":EB08
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   14
      Top             =   8040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pHinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   7320
      Picture         =   "Intro.frx":125CA
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   13
      Top             =   8040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pNegWinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   8400
      Picture         =   "Intro.frx":1608C
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   12
      Top             =   6720
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pWinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   7320
      Picture         =   "Intro.frx":19B4E
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   11
      Top             =   6720
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pNegStinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   8400
      Picture         =   "Intro.frx":1D610
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pStinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   7320
      Picture         =   "Intro.frx":210D2
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   9
      Top             =   5400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pNegSlinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   8400
      Picture         =   "Intro.frx":24B94
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pSlinky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   7320
      Picture         =   "Intro.frx":28656
      ScaleHeight     =   1170
      ScaleWidth      =   960
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pWith 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   330
      Left            =   120
      Picture         =   "Intro.frx":2C118
      ScaleHeight     =   270
      ScaleWidth      =   870
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox pNegStarring 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   120
      Picture         =   "Intro.frx":2CDBA
      ScaleHeight     =   300
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox pStarring 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   120
      Picture         =   "Intro.frx":2EF1C
      ScaleHeight     =   300
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox pNegTurbo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   855
      Left            =   120
      Picture         =   "Intro.frx":3107E
      ScaleHeight     =   795
      ScaleWidth      =   3750
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.PictureBox pTurbo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   855
      Left            =   120
      Picture         =   "Intro.frx":3AC70
      ScaleHeight     =   795
      ScaleWidth      =   3750
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.PictureBox NegTitle 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1995
      Left            =   120
      Picture         =   "Intro.frx":44862
      ScaleHeight     =   1935
      ScaleWidth      =   7830
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   7890
   End
   Begin VB.PictureBox Title 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1995
      Left            =   120
      Picture         =   "Intro.frx":75EC4
      ScaleHeight     =   1935
      ScaleWidth      =   7830
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7890
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture3_Click()

End Sub
