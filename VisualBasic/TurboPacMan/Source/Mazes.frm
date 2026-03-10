VERSION 5.00
Begin VB.Form Mazes 
   AutoRedraw      =   -1  'True
   Caption         =   "Mazes"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox MazeA 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   4
      Left            =   840
      Picture         =   "Mazes.frx":0000
      ScaleHeight     =   8520
      ScaleWidth      =   8070
      TabIndex        =   19
      Top             =   2880
      Width           =   8130
   End
   Begin VB.PictureBox MazeA 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   3
      Left            =   600
      Picture         =   "Mazes.frx":E01C2
      ScaleHeight     =   8520
      ScaleWidth      =   8355
      TabIndex        =   18
      Top             =   3240
      Width           =   8415
   End
   Begin VB.PictureBox MazeA 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   2
      Left            =   960
      Picture         =   "Mazes.frx":1C7FC4
      ScaleHeight     =   8520
      ScaleWidth      =   8070
      TabIndex        =   17
      Top             =   3000
      Width           =   8130
   End
   Begin VB.PictureBox MazeA 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   1
      Left            =   840
      Picture         =   "Mazes.frx":2A8186
      ScaleHeight     =   8520
      ScaleWidth      =   8070
      TabIndex        =   16
      Top             =   2880
      Width           =   8130
   End
   Begin VB.PictureBox MazeA 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   0
      Left            =   840
      Picture         =   "Mazes.frx":388348
      ScaleHeight     =   8520
      ScaleWidth      =   8295
      TabIndex        =   15
      Top             =   2880
      Width           =   8355
   End
   Begin VB.PictureBox NegDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   4
      Left            =   960
      Picture         =   "Mazes.frx":46E6AA
      ScaleHeight     =   8520
      ScaleWidth      =   8070
      TabIndex        =   14
      Top             =   3000
      Width           =   8130
   End
   Begin VB.PictureBox NegDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   3
      Left            =   840
      Picture         =   "Mazes.frx":54E86C
      ScaleHeight     =   8520
      ScaleWidth      =   8070
      TabIndex        =   13
      Top             =   2880
      Width           =   8130
   End
   Begin VB.PictureBox NegDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   2
      Left            =   960
      Picture         =   "Mazes.frx":62EA2E
      ScaleHeight     =   8520
      ScaleWidth      =   8070
      TabIndex        =   12
      Top             =   3000
      Width           =   8130
   End
   Begin VB.PictureBox NegDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   1
      Left            =   960
      Picture         =   "Mazes.frx":70EBF0
      ScaleHeight     =   8520
      ScaleWidth      =   8475
      TabIndex        =   11
      Top             =   2880
      Width           =   8535
   End
   Begin VB.PictureBox NegDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   0
      Left            =   960
      Picture         =   "Mazes.frx":7F9F32
      ScaleHeight     =   8520
      ScaleWidth      =   8070
      TabIndex        =   10
      Top             =   3000
      Width           =   8130
   End
   Begin VB.PictureBox PosDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   4
      Left            =   960
      Picture         =   "Mazes.frx":8DA0F4
      ScaleHeight     =   8520
      ScaleWidth      =   8385
      TabIndex        =   9
      Top             =   3000
      Width           =   8445
   End
   Begin VB.PictureBox PosDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   3
      Left            =   840
      Picture         =   "Mazes.frx":9C30B6
      ScaleHeight     =   8520
      ScaleWidth      =   8625
      TabIndex        =   8
      Top             =   2880
      Width           =   8685
   End
   Begin VB.PictureBox PosDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   2
      Left            =   960
      Picture         =   "Mazes.frx":AB2AF8
      ScaleHeight     =   8520
      ScaleWidth      =   8505
      TabIndex        =   7
      Top             =   2880
      Width           =   8565
   End
   Begin VB.PictureBox PosDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   1
      Left            =   840
      Picture         =   "Mazes.frx":B9EFFA
      ScaleHeight     =   8520
      ScaleWidth      =   8475
      TabIndex        =   6
      Top             =   2880
      Width           =   8535
   End
   Begin VB.PictureBox PosDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   0
      Left            =   840
      Picture         =   "Mazes.frx":C8A33C
      ScaleHeight     =   8520
      ScaleWidth      =   8295
      TabIndex        =   5
      Top             =   2760
      Width           =   8355
   End
   Begin VB.PictureBox Maze 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   4
      Left            =   720
      Picture         =   "Mazes.frx":D7069E
      ScaleHeight     =   8520
      ScaleWidth      =   8070
      TabIndex        =   4
      Top             =   2760
      Width           =   8130
   End
   Begin VB.PictureBox Maze 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   3
      Left            =   960
      Picture         =   "Mazes.frx":E50860
      ScaleHeight     =   8520
      ScaleWidth      =   8355
      TabIndex        =   3
      Top             =   2760
      Width           =   8415
   End
   Begin VB.PictureBox Maze 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   2
      Left            =   840
      Picture         =   "Mazes.frx":F38662
      ScaleHeight     =   8520
      ScaleWidth      =   8070
      TabIndex        =   2
      Top             =   2760
      Width           =   8130
   End
   Begin VB.PictureBox Maze 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   1
      Left            =   0
      Picture         =   "Mazes.frx":1018824
      ScaleHeight     =   8520
      ScaleWidth      =   8070
      TabIndex        =   1
      Top             =   0
      Width           =   8130
   End
   Begin VB.PictureBox Maze 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8580
      Index           =   0
      Left            =   0
      Picture         =   "Mazes.frx":10F89E6
      ScaleHeight     =   8520
      ScaleWidth      =   8295
      TabIndex        =   0
      Top             =   0
      Width           =   8355
   End
End
Attribute VB_Name = "Mazes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
