VERSION 5.00
Begin VB.Form Characters 
   Caption         =   "Characters"
   ClientHeight    =   10260
   ClientLeft      =   -285
   ClientTop       =   1215
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   10260
   ScaleWidth      =   8835
   Begin VB.PictureBox Chars2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   10245
      Left            =   4440
      Picture         =   "Characters.frx":0000
      ScaleHeight     =   679
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   290
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4410
   End
   Begin VB.PictureBox Chars 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   10245
      Left            =   0
      Picture         =   "Characters.frx":9091A
      ScaleHeight     =   679
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   290
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4410
   End
End
Attribute VB_Name = "Characters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
