VERSION 5.00
Begin VB.Form Dots 
   Caption         =   "Dots"
   ClientHeight    =   8625
   ClientLeft      =   11385
   ClientTop       =   2805
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   Begin VB.PictureBox Maze 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   8580
      Left            =   6240
      Picture         =   "Dots.frx":0000
      ScaleHeight     =   568
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   8355
   End
   Begin VB.PictureBox PosDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   8580
      Left            =   8160
      Picture         =   "Dots.frx":4D5E2
      ScaleHeight     =   568
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   8355
   End
   Begin VB.PictureBox NegDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   8580
      Left            =   3600
      Picture         =   "Dots.frx":9ABC4
      ScaleHeight     =   568
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   538
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   8130
   End
   Begin VB.PictureBox DotsBegin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   8580
      Left            =   2400
      Picture         =   "Dots.frx":E5E26
      ScaleHeight     =   568
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   8355
   End
   Begin VB.PictureBox Maze2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   8580
      Left            =   120
      Picture         =   "Dots.frx":133408
      ScaleHeight     =   8520
      ScaleWidth      =   8295
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   8355
   End
End
Attribute VB_Name = "Dots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
