VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Turbo PacMan"
   ClientHeight    =   9030
   ClientLeft      =   3720
   ClientTop       =   1575
   ClientWidth     =   8130
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":0442
   ScaleHeight     =   602
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   Begin VB.Timer Timer3 
      Left            =   1560
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   0
      ScaleHeight     =   569
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   545
      TabIndex        =   2
      Top             =   480
      Width           =   8175
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   2160
         Top             =   120
      End
      Begin VB.Label lblStart 
         BackStyle       =   0  'Transparent
         Caption         =   "PRESS          F1       F2"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2310
         TabIndex        =   7
         Top             =   4200
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblReady 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "READY!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3120
         TabIndex        =   5
         Top             =   5325
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblTurbo 
         BackStyle       =   0  'Transparent
         Caption         =   "TURBO"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label GameOver 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Game Over"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3075
         TabIndex        =   3
         Top             =   4080
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   0
   End
   Begin VB.Label text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   5760
      TabIndex        =   6
      Top             =   -120
      Width           =   2295
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblScoreHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pellets(1 To 6) As PowerPellet
Dim Ghosts(1 To 6) As Ghost
Dim PACMAN As Pac
Dim MSPACMAN As Pac
Dim Food As Bonus
Dim levels(1 To 15) As GameLevel
Dim Score As Long
Dim Level As Long

Dim secret As String          'This string is used to award bonus life
Dim secret2 As String         'This string is for unlocking unlimited TURBO

Dim FRAME_INDEX As Byte       'Flips Power Pellets on/off, Boosts Red ghost
Dim DEATH_TOLL As Byte        'Keeps track of how many Ghosts have been eaten
Dim EXPLODE_FRAME As Byte     'Used to blowup ghosts when bomb is eaten
Dim BONUS_MULTIPLIER          'Doubles points when 2x is eaten

Dim CAN_START As Boolean      'keeps player from starting before intro has finished
Dim STARTED As Boolean        'Goes to TRUE after game begins
Dim TURBO_ON As Boolean       'detemines TURBO status
Dim GHOSTS_PISSED As Boolean  'If all 6 ghosts are eaten, they get P.O.'d (faster)
Dim LIGHTNING As Boolean      'for undepletable turbo
Dim CHILLED As Boolean        'for slowing down ghosts
Dim FROZEN As Boolean         'for freezing ghosts
Dim RECEIVED_BONUS As Boolean
Dim PAUSE As Boolean          'For pausing the game

Dim DIRECTION_KEY As Integer  'Keeps track of which key has been pressed
Dim FRUIT_FRAME As Integer    'Used to make the FOOD/Fruit scoot when it moves
Dim CHANGE_FRAME As Integer   'Flips scared ghosts between blue and white
Dim SLOW_DOWN As Integer      'Slows down scared ghosts
Dim INTRO_COUNTER As Integer  'Determines what part of the Intro to perform
Dim TITLE_X As Integer        'Keeps Track of Intro Title location for scroll
Dim TITLE_Y As Integer        'Keeps Track of Intro Title location for scroll
Dim TIME_TIME As Integer      'Time interval for timer, used for quick change in code
Dim DOT_MAX As Integer        'Holds number of dots in a level
Dim TURBO As Integer          'Keeps track o how much turbo player has
Dim SCARED_TIME As Integer    'Amount of time Power Pellets are effective
Dim EATEN_TIMER As Integer    'PAUSEs everything when a Ghost is eaten
Dim DOT_COUNTER As Integer    'Keeps track of how many dots have been eaten
Dim PELLET_COUNTER As Integer 'Keeps track of how many pellets have been eaten

Dim RUNNING_SCARED As Long    'Timer for ghosts after power pellet has been eaten
Dim BEGIN_TIMER As Long       'Used for delay between levels
Dim EXTRA As Long             'Time to add to power pellet timer when special item eaten
Dim FREEZE_TIMER As Long      'Timer for ghost freeze or slowdown
Dim CHAR_SOURCE As Long       'Determines which Image to get the ghost graphics graphics
Dim PAC_SOURCE As Long        'Detemines which image to get Pac graphics
Dim START_TIMER As Long       'Used for time release of Ghosts

Dim DOT_COLOR As Long         'Defines the color of a dot, can vary between computers
Dim BLACK_COLOR As Long       'Defines the background color
Dim ALMOST_BLACK As Long      'Defines the color used to round the maze's corners
Dim WALL_COLOR As Long        'Defines the Maze's Wall color
Dim DOOR As Long              'Defines the  Ghost Base Door color

Const East = 0                'Directions for Pacman, Ghosts, Bonus. These are
Const West = 1                'Also used to determine what frame to animate the
Const North = 2               'Character with as well as what the directional
Const South = 3               'Vector of the character is
Const bonus_score = 30000     'This is the score that extra Pac is awarded

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Picture1_KeyDown KeyCode, Shift
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 Picture1_KeyUp KeyCode, Shift
End Sub

Private Sub Form_Load()
  PAC_SOURCE = Characters.Chars.hDC
  TURBO_ON = True
  CAN_START = False
  INTRO_COUNTER = 0
  TITLE_X = 10
  TITLE_Y = Picture1.Height '-1 * Intro.Title.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
sndPlaySound "die.wav", SND_NODEFAULT Or SND_ASYNC
 End
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
PAUSE = False
If (KeyCode = 32) Then
  PACMAN.TURBO = True
Else
  DIRECTION_KEY = KeyCode
End If
End Sub
Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 1 Then
  secret2 = secret2 + Chr(KeyCode)
  If Len(secret2) >= 5 Then
    If secret2 = "TURBO" Then TURBO_ON = False
    secret2 = ""
  End If
End If

If (KeyCode = 88) And (Shift = 7) Then
  PACMAN.lives = PACMAN.lives + 1
  Display_Lives
ElseIf (KeyCode = 65) And (Shift = 7) Then
  DOT_COUNTER = levels(Level).DotMax
  PELLET_COUNTER = 6
ElseIf (KeyCode = 32) Then
  PACMAN.TURBO = False
ElseIf (KeyCode = 112) Then
  If CAN_START Then
    PAC_SOURCE = Characters.Chars.hDC
    Timer2.Interval = 0
    GameOver.Visible = False
    PACMAN.MS = False
    Picture1.Visible = False
    Startup
    Timer1.Interval = TIME_TIME
    lblReady.Visible = True
    Picture1.Refresh
    sndPlaySound "theme.wav", SND_NODEFAULT Or SND_SYNC
  End If
ElseIf (KeyCode = 113) Then
  If CAN_START Then
    PAC_SOURCE = Characters.Chars.hDC
    Timer2.Interval = 0
    GameOver.Visible = False
    PACMAN.MS = True
    Picture1.Visible = False
    Startup
    Timer1.Interval = TIME_TIME
    lblReady.Visible = True
    Picture1.Refresh
    sndPlaySound "mstheme.wav", SND_NODEFAULT Or SND_SYNC
  End If
ElseIf (KeyCode = 114) Then
  If CAN_START Then
    PAC_SOURCE = Characters.Chars2.hDC
    Timer2.Interval = 0
    GameOver.Visible = False
    PACMAN.MS = False
    Picture1.Visible = False
    Startup
    Timer1.Interval = TIME_TIME
    lblReady.Visible = True
    Picture1.Refresh
    sndPlaySound "theme.wav", SND_NODEFAULT Or SND_SYNC
  End If
ElseIf (KeyCode = 115) Then
  If CAN_START Then
    PAC_SOURCE = Characters.Chars2.hDC
    Timer2.Interval = 0
    GameOver.Visible = False
    PACMAN.MS = True
    Picture1.Visible = False
    Startup
    Timer1.Interval = TIME_TIME
    lblReady.Visible = True
    Picture1.Refresh
    sndPlaySound "mstheme.wav", SND_NODEFAULT Or SND_SYNC
  End If
ElseIf (KeyCode = 80) Then
  PAUSE = True
  sndPlaySound 0, SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
Else
  If KeyCode = DIRECTION_KEY Then DIRECTION_KEY = 0
End If
End Sub

Private Sub Timer1_Timer()
If STARTED And Not PAUSE Then
  FRUIT_FRAME = FRUIT_FRAME + 1
  Food.Timer = Food.Timer + 1
  
  If PACMAN.Poisoned Then
    PACMAN.PoisonCounter = PACMAN.PoisonCounter - 1
    If PACMAN.PoisonCounter <= 0 Then PACMAN.Poisoned = False
  End If
  
  If FRUIT_FRAME >= 30 Then FRUIT_FRAME = 0
  If DOT_COUNTER >= DOT_MAX And PELLET_COUNTER = 6 Then
    If DOT_COUNTER = DOT_MAX Then
      sndPlaySound 0, SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
      Erase_Ghosts
      Timer1.Interval = 250
    End If
    DOT_COUNTER = DOT_COUNTER + 1
    CHANGE_FRAME = -CHANGE_FRAME
    If DOT_COUNTER > DOT_MAX + 10 Then
      If Level < 15 Then Level = Level + 1
      Food.Timer = 0
      Restore_Food
      Timer1.Interval = 0
      Build_Maze
      Initialize
      Picture1.Refresh
      DOT_COUNTER = 0
      Timer1.Interval = TIME_TIME
      Exit Sub
    End If

    If CHANGE_FRAME > 0 Then
      BitBlt Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, Dots.Maze.hDC, 0, 0, &HCC0020
      Draw_Pacman
    Else
      BitBlt Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, Dots.Maze2.hDC, 0, 0, &HCC0020
      Draw_Pacman
    End If
    Picture1.Refresh
    Exit Sub
  End If

  If Food.Timer = 2500 Then
    Food.Eaten = False
  ElseIf PACMAN.MS And Food.Timer > 4000 Then
    Food.Timer = 2500
  ElseIf Food.Timer = 4000 Then
    Food.Eaten = True
    Food.Timer = 0
  End If

  If EATEN_TIMER > 0 Then
    If EATEN_TIMER = 70 Then
      Erase_Ghosts
      Draw_Ghosts
      Picture1.Refresh
    End If
    EATEN_TIMER = EATEN_TIMER - 1
  Else
    If START_TIMER < 10000 Then START_TIMER = START_TIMER + 1
    If PACMAN.Dying Then
      If PACMAN.Frame < 0 Then
        If PACMAN.MS Then
          Timer1.Interval = 80
        Else
          Timer1.Interval = 70
        End If
      End If
      If PACMAN.Frame = 0 Then
        Restore_Food
        Erase_Ghosts
      ElseIf PACMAN.Frame = -4 Then
        sndPlaySound "die.wav", SND_NODEFAULT Or SND_ASYNC
      End If
  
      If PACMAN.MS Then
        If PACMAN.Frame >= 0 And PACMAN.Frame < 15 Then
          PACMAN.Direction = PACMAN.Direction + 1
          If PACMAN.Direction > 3 Then PACMAN.Direction = 1
          BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, PAC_SOURCE, 35, PACMAN.Src_Y + (17 * PACMAN.Direction), &HCC0020
        ElseIf PACMAN.Frame = 15 Then
          BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, Dots.PosDots.hDC, PACMAN.X, PACMAN.Y, &HCC0020
        End If
      Else
        If PACMAN.Frame >= 0 And PACMAN.Frame < 15 Then
          BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, PAC_SOURCE, 1 + (PACMAN.Frame * 17), 544 + (PACMAN.Direction * 17), &HCC0020
        ElseIf PACMAN.Frame = 15 Then
          BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, Dots.PosDots.hDC, PACMAN.X, PACMAN.Y, &HCC0020
        End If
      End If
      PACMAN.Frame = PACMAN.Frame + 1
      If PACMAN.Frame > 30 Then
        If PACMAN.lives < 1 Then
          GameOver.Visible = True
          Timer1.Interval = 0
          Exit Sub
        End If
        PACMAN.Frame = 1
        PACMAN.Dying = False
        PACMAN.Poisoned = False
        PACMAN.PoisonCounter = 0
        GHOSTS_PISSED = False
        Initialize
        Timer1.Interval = TIME_TIME
        Display_Lives
      End If
      Picture1.Refresh
    Else
      SLOW_DOWN = SLOW_DOWN - 1
      If SLOW_DOWN < 0 Then SLOW_DOWN = 2
      If (RUNNING_SCARED > 0) Then
        RUNNING_SCARED = RUNNING_SCARED - 1
        If (RUNNING_SCARED = 0) Then
          sndPlaySound 0, SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY
          secret = ""
          Ghosts(1).Scared = False
          Ghosts(2).Scared = False
          Ghosts(3).Scared = False
          Ghosts(4).Scared = False
          Ghosts(5).Scared = False
          Ghosts(6).Scared = False
          DEATH_TOLL = 0
        End If
      End If
      Erase_Ghosts
      If (FRAME_INDEX = 15) Then
        FRAME_INDEX = 0
        CHANGE_FRAME = -CHANGE_FRAME
      End If

      FRAME_INDEX = FRAME_INDEX + 1
      Detect_Dot
      If Not PACMAN.Poisoned Or (PACMAN.Poisoned And ((FRAME_INDEX Mod 2) = 0)) Then
        Move_Pac
      End If
      Detect_Dot
      lblScore = Score
 
      If PACMAN.TURBO And TURBO >= 0 Then
        If TURBO_ON And Not LIGHTNING Then TURBO = TURBO - 1
        If Not PACMAN.Poisoned Or (PACMAN.Poisoned And ((FRAME_INDEX Mod 2) = 0)) Then
          Move_Pac
        End If
        Detect_Dot
        lblScore = Score
      End If

      If (FRAME_INDEX = 7) Then
        CHANGE_FRAME = -CHANGE_FRAME
      End If

      Draw_TURBO
      
      If FROZEN Or CHILLED Then FREEZE_TIMER = FREEZE_TIMER - 1
      If FROZEN Then
        If FRAME_INDEX = 1 Then sndPlaySound "count.wav", SND_NODEFAULT Or SND_ASYNC Or SND_NOSTOP
        If FREEZE_TIMER <= 0 Then
          FROZEN = False
          FREEZE_TIMER = 0
          CHILLED = False
          sndPlaySound 0, SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
        End If
      End If
      If FREEZE_TIMER <= 0 Then CHILLED = False
      
      If Not FROZEN Then
        If Not CHILLED Or (CHILLED And ((FRAME_INDEX Mod 2) = 0)) Then Move_Ghosts
      End If
      If (GHOSTS_PISSED) And ((FRAME_INDEX Mod 5) = 0) And Not FROZEN And Not CHILLED Then Move_Ghosts
      Detect_Dot
      Draw_Pacman
      Draw_Ghosts
      If Score > bonus_score And Not RECEIVED_BONUS Then
        PACMAN.lives = PACMAN.lives + 1
        RECEIVED_BONUS = True
        Display_Lives
        sndPlaySound "xtra.wav", SND_NODEFAULT Or SND_ASYNC
      End If
   
      Detect_Collision
      Picture1.Refresh
    End If
  End If
Else
  lblReady.Visible = True
  BEGIN_TIMER = BEGIN_TIMER + 1
  If BEGIN_TIMER > 50 Then
    STARTED = True
    lblReady.Visible = False
  End If
End If
End Sub
Private Sub Move_Pac()
Dim pixel_1 As Long
Dim pixel_2 As Long
'See if Pacman can/needs turning
If (DIRECTION_KEY = 37) Then
  pixel_1 = GetPixel(Dots.Maze.hDC, PACMAN.X - 1, PACMAN.Y)
  pixel_2 = GetPixel(Dots.Maze.hDC, PACMAN.X - 1, PACMAN.Y + 15)
  If (pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR) Then
    PACMAN.Direction = West
    PACMAN.Y_vector = 0
    PACMAN.X_vector = -1
  End If
ElseIf (DIRECTION_KEY = 38) Then
  pixel_1 = GetPixel(Dots.Maze.hDC, PACMAN.X, PACMAN.Y - 1)
  pixel_2 = GetPixel(Dots.Maze.hDC, PACMAN.X + 15, PACMAN.Y - 1)
  If (pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR) And ((PACMAN.X > 4) And ((PACMAN.X) < (Picture1.Width - 23))) Then
    PACMAN.Direction = North
    PACMAN.Y_vector = -1
    PACMAN.X_vector = 0
  End If
ElseIf (DIRECTION_KEY = 39) Then
  pixel_1 = GetPixel(Dots.Maze.hDC, PACMAN.X + 16, PACMAN.Y)
  pixel_2 = GetPixel(Dots.Maze.hDC, PACMAN.X + 16, PACMAN.Y + 15)
  If (pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR) Then
    PACMAN.Direction = East
    PACMAN.Y_vector = 0
    PACMAN.X_vector = 1
  End If
ElseIf DIRECTION_KEY = 40 Then
  pixel_1 = GetPixel(Dots.Maze.hDC, PACMAN.X, PACMAN.Y + 16)
  pixel_2 = GetPixel(Dots.Maze.hDC, PACMAN.X + 15, PACMAN.Y + 16)
  If (pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR) And ((PACMAN.X > 4) And ((PACMAN.X) < (Picture1.Width - 23))) Then
    PACMAN.Direction = South
    PACMAN.Y_vector = 1
    PACMAN.X_vector = 0
  End If
End If
'See if Pacman can move in the direction he's facing
If (PACMAN.Direction = East) Then
  pixel_1 = GetPixel(Dots.Maze.hDC, PACMAN.X + PACMAN.X_vector + 15, PACMAN.Y + PACMAN.Y_vector)
  pixel_2 = GetPixel(Dots.Maze.hDC, PACMAN.X + PACMAN.X_vector + 15, PACMAN.Y + PACMAN.Y_vector + 15)
ElseIf PACMAN.Direction = West Then
  pixel_1 = GetPixel(Dots.Maze.hDC, PACMAN.X + PACMAN.X_vector, PACMAN.Y + PACMAN.Y_vector)
  pixel_2 = GetPixel(Dots.Maze.hDC, PACMAN.X + PACMAN.X_vector, PACMAN.Y + PACMAN.Y_vector + 15)
ElseIf PACMAN.Direction = North Then
  pixel_1 = GetPixel(Dots.Maze.hDC, PACMAN.X + PACMAN.X_vector, PACMAN.Y + PACMAN.Y_vector)
  pixel_2 = GetPixel(Dots.Maze.hDC, PACMAN.X + PACMAN.X_vector + 15, (PACMAN.Y + PACMAN.Y_vector))
ElseIf PACMAN.Direction = South Then
  pixel_1 = GetPixel(Dots.Maze.hDC, PACMAN.X + PACMAN.X_vector, PACMAN.Y + PACMAN.Y_vector + 15)
  pixel_2 = GetPixel(Dots.Maze.hDC, PACMAN.X + PACMAN.X_vector + 15, PACMAN.Y + PACMAN.Y_vector + 15)
End If

'There might be a btter way to check this. It might be simpler
'To check the empty maze at the pac location for black pixels
If (pixel_1 <> DOOR And pixel_2 <> DOOR) And (pixel_1 <> WALL_COLOR And pixel_2 <> WALL_COLOR) And (pixel_1 <> ALMOST_BLACK And pixel_2 <> ALMOST_BLACK) Then
  PACMAN.Y = PACMAN.Y + PACMAN.Y_vector
  PACMAN.X = PACMAN.X + PACMAN.X_vector
  If (PACMAN.X < -15) Then
    PACMAN.X = Dots.Maze.Width + 15
  ElseIf (PACMAN.X > Dots.Maze.Width + 15) Then
    PACMAN.X = -15
  End If
  'See if there is a Dot to be eaten
End If
PACMAN.Frame = PACMAN.Frame + PACMAN.frame_val
If PACMAN.Frame = 4 Or PACMAN.Frame = 0 Then
   PACMAN.frame_val = -PACMAN.frame_val
End If
End Sub
Private Sub Detect_Dot()
Dim pixel_1 As Long
pixel_1 = GetPixel(Dots.PosDots.hDC, PACMAN.X + 7, PACMAN.Y + 7)
  
If (pixel_1 = DOT_COLOR) Then
 If RUNNING_SCARED < 0 Or RUNNING_SCARED < SCARED_TIME - (1000 - (100 * Level)) Then
  sndPlaySound "dot.wav", SND_NODEFAULT Or SND_ASYNC
 Else
  sndPlaySound "scaredstart.wav", SND_NODEFAULT Or SND_ASYNC Or SND_NOSTOP Or SND_LOOP
 End If
 Score = Score + (BONUS_MULTIPLIER * 10)
 DOT_COUNTER = DOT_COUNTER + 1
 If TURBO < 596 Then TURBO = TURBO + 5
 BitBlt Picture1.hDC, PACMAN.X + 2, PACMAN.Y + 2, 14, 14, PAC_SOURCE, 239, 0, &HCC0020
 BitBlt Dots.PosDots.hDC, PACMAN.X + 2, PACMAN.Y + 2, 14, 14, PAC_SOURCE, 239, 0, &HCC0020
Else
If RUNNING_SCARED < 0 Or RUNNING_SCARED < SCARED_TIME - (900 - (100 * Level)) Then
  If Not FROZEN Then sndPlaySound "siren.wav", SND_NODEFAULT Or SND_ASYNC Or SND_NOSTOP Or SND_LOOP
Else
  sndPlaySound "scaredstart.wav", SND_NODEFAULT Or SND_ASYNC Or SND_NOSTOP Or SND_LOOP
End If
End If
  
End Sub
Private Sub Draw_Pacman()
  BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, PAC_SOURCE, ((PACMAN.Frame + 5) * 16) + (PACMAN.Frame + 5), PACMAN.Src_Y + (PACMAN.Direction * 16) + PACMAN.Direction, vbSrcAnd
  BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, PAC_SOURCE, (PACMAN.Frame * 16) + (PACMAN.Frame + 1), PACMAN.Src_Y + (PACMAN.Direction * 16) + PACMAN.Direction, vbSrcPaint
End Sub
Private Sub Draw_Ghosts()
Dim itr As Byte
Dim Src_X As Long
Dim Src_Y As Long

For itr = 1 To 6
 If (Ghosts(itr).Eaten = False) Then
  If (Ghosts(itr).Scared = True) Then
    If (RUNNING_SCARED < (2 * (SCARED_TIME / 10))) And CHANGE_FRAME < 0 Then
      'White ghost
      BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 171 + (Ghosts(itr).Frame * 17), 51, vbSrcAnd
      BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 137 + (Ghosts(itr).Frame * 17), Ghosts(itr).Src_Y, vbSrcPaint
    Else
      'Blue ghost
      BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 171 + (Ghosts(itr).Frame * 17), 51, vbSrcAnd
      BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 86 + (Ghosts(itr).Frame * 17), Ghosts(itr).Src_Y, vbSrcPaint
    End If
  Else
    BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 171 + (Ghosts(itr).Frame * 17), 51, vbSrcAnd
    BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, Ghosts(itr).Src_X + (Ghosts(itr).Frame * 17), Ghosts(itr).Src_Y + (Ghosts(itr).Direction * 17), vbSrcPaint
  End If
  If (CHANGE_FRAME < 0) Then Ghosts(itr).Frame = Ghosts(itr).Frame + 1
  If Ghosts(itr).Frame > 2 Then Ghosts(itr).Frame = 0
 End If
Next itr

For itr = 1 To 6
 If (Ghosts(itr).Eaten = True) Then
    If Ghosts(itr).JustEaten Then
      'Score graphics
      Ghosts(itr).JustEaten = False
      BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 1 + ((DEATH_TOLL - 1) * 17) + ((BONUS_MULTIPLIER - 1) * 17), 493, vbSrcAnd
      BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 1 + ((DEATH_TOLL - 1) * 17) + ((BONUS_MULTIPLIER - 1) * 17), 476, vbSrcPaint
    Else
      'Eyeballs
      BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 52 + 17, Ghosts(itr).Src_Y + (Ghosts(itr).Direction * 17), vbSrcAnd
      BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 52, Ghosts(itr).Src_Y + (Ghosts(itr).Direction * 17), vbSrcPaint
    End If
  If (CHANGE_FRAME < 0) Then Ghosts(itr).Frame = Ghosts(itr).Frame + 1
  If Ghosts(itr).Frame > 2 Then Ghosts(itr).Frame = 0
 End If
Next itr
End Sub
Public Sub Move_Ghosts()
  Dim itr As Byte
  Dim reverse As Boolean
  Dim rnd_num As Long
  
  rnd_num = Int(4000 * Rnd())
  
  If rnd_num = 10 Then
     reverse = True
  Else
     reverse = False
  End If
  
  For itr = 1 To 6
    Move_Ghost itr, reverse
    If GHOSTS_PISSED = False And FRAME_INDEX > Ghosts(itr).Speed Then Move_Ghost itr, False
  Next itr
End Sub

Public Sub Move_Ghost(Index As Byte, reverse As Boolean)
  Dim pixel_1 As Long
  Dim pixel_2 As Long
  Dim pixel_3 As Long
  Dim pixel_4 As Long
  Dim pixel_5 As Long
  Dim pixel_6 As Long
  Dim pixel_7 As Long
  Dim pixel_8 As Long
  
  Dim door_color As Long
  Dim rand_dir As Long
  Dim seed As Long
  
If (Ghosts(Index).Scared = False) Or SLOW_DOWN = 0 Then
  door_color = DOOR
  
  If reverse And Ghosts(Index).Eaten = False And Ghosts(Index).Scared = False Then
    'we'll need to change direction, preferably reverse
    'north
    pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
    pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
    
    'south
    pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
    pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
    
    'east
    pixel_5 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
    pixel_6 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
    
    'west
    pixel_7 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
    pixel_8 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
    
    If Ghosts(Index).Direction = North Then
      'try to go south first
      If pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
        Ghosts(Index).Direction = South
        Ghosts(Index).Y = Ghosts(Index).Y + 1
      ElseIf pixel_5 = BLACK_COLOR And pixel_6 = BLACK_COLOR Then
        Ghosts(Index).Direction = East
        Ghosts(Index).X = Ghosts(Index).X + 1
      ElseIf pixel_7 = BLACK_COLOR And pixel_8 = BLACK_COLOR Then
        Ghosts(Index).Direction = West
        Ghosts(Index).X = Ghosts(Index).X - 1
      ElseIf pixel_1 = BLACK_COLOR And pixel_1 = BLACK_COLOR Then
        Ghosts(Index).Direction = North
        Ghosts(Index).Y = Ghosts(Index).Y - 1
      End If
    ElseIf Ghosts(Index).Direction = South Then
      'try to go north first
      If pixel_1 = BLACK_COLOR And pixel_1 = BLACK_COLOR Then
        Ghosts(Index).Direction = North
        Ghosts(Index).Y = Ghosts(Index).Y - 1
      ElseIf pixel_5 = BLACK_COLOR And pixel_6 = BLACK_COLOR Then
        Ghosts(Index).Direction = East
        Ghosts(Index).X = Ghosts(Index).X + 1
      ElseIf pixel_7 = BLACK_COLOR And pixel_8 = BLACK_COLOR Then
        Ghosts(Index).Direction = West
        Ghosts(Index).X = Ghosts(Index).X - 1
      ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
        Ghosts(Index).Direction = South
        Ghosts(Index).Y = Ghosts(Index).Y + 1
      End If
    ElseIf Ghosts(Index).Direction = East Then
      'try to go west first
      If pixel_7 = BLACK_COLOR And pixel_8 = BLACK_COLOR Then
        Ghosts(Index).Direction = West
        Ghosts(Index).X = Ghosts(Index).X - 1
      ElseIf pixel_1 = BLACK_COLOR And pixel_1 = BLACK_COLOR Then
        Ghosts(Index).Direction = North
        Ghosts(Index).Y = Ghosts(Index).Y - 1
      ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
        Ghosts(Index).Direction = South
        Ghosts(Index).Y = Ghosts(Index).Y + 1
      ElseIf pixel_5 = BLACK_COLOR And pixel_6 = BLACK_COLOR Then
        Ghosts(Index).Direction = East
        Ghosts(Index).X = Ghosts(Index).X + 1
      End If
    ElseIf Ghosts(Index).Direction = West Then
      'try to go east first
      If pixel_5 = BLACK_COLOR And pixel_6 = BLACK_COLOR Then
        Ghosts(Index).Direction = East
        Ghosts(Index).X = Ghosts(Index).X + 1
      ElseIf pixel_1 = BLACK_COLOR And pixel_1 = BLACK_COLOR Then
        Ghosts(Index).Direction = North
        Ghosts(Index).Y = Ghosts(Index).Y - 1
      ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
        Ghosts(Index).Direction = South
        Ghosts(Index).Y = Ghosts(Index).Y + 1
      ElseIf pixel_7 = BLACK_COLOR And pixel_8 = BLACK_COLOR Then
        Ghosts(Index).Direction = West
        Ghosts(Index).X = Ghosts(Index).X - 1
      End If
    End If
  Else
    If Ghosts(Index).Direction = North Then
      pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
      pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
    ElseIf Ghosts(Index).Direction = South Then
      pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
      pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
    ElseIf Ghosts(Index).Direction = East Then
      pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
      pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
    Else
      pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
      pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
    End If
  
    'This first bit leads the ghost (eye balls) back home
    If Ghosts(Index).Eaten Then
      If (Ghosts(Index).X >= Picture1.Width - 4) Then
        Ghosts(Index).X = Ghosts(Index).X - 1
      ElseIf (Ghosts(Index).X <= -12) Then
        Ghosts(Index).X = Ghosts(Index).X + 1
      ElseIf (Ghosts(Index).Y > 260) And Ghosts(Index).Direction <> South Then
        'then see if we can go North
        pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
        pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
        If pixel_1 <> 16711680 And pixel_1 <> 526344 And pixel_2 <> 16711680 And pixel_2 <> 526344 Then
          Ghosts(Index).Direction = North
          Ghosts(Index).Y = Ghosts(Index).Y - 1
        Else
          pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
          pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
          pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
          pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
          'see if we can go East or West
          If Ghosts(Index).Direction <> East And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = West
            Ghosts(Index).X = Ghosts(Index).X - 1
          ElseIf Ghosts(Index).Direction <> West And pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
            Ghosts(Index).Direction = East
            Ghosts(Index).X = Ghosts(Index).X + 1
          ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = West
            Ghosts(Index).X = Ghosts(Index).X - 1
          ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
            Ghosts(Index).Direction = East
            Ghosts(Index).X = Ghosts(Index).X + 1
          End If
        End If
      ElseIf (Ghosts(Index).X > 261) And Ghosts(Index).Direction <> East Then
        'see if we can go West
        pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
        pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
        If pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
          Ghosts(Index).Direction = West
          Ghosts(Index).X = Ghosts(Index).X - 1
        Else
          pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
          pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
          pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
          pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
          'see if we can go North or South
          If Ghosts(Index).Direction <> South And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = North
            Ghosts(Index).Y = Ghosts(Index).Y - 1
          ElseIf Ghosts(Index).Direction <> North And (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
            Ghosts(Index).Direction = South
            Ghosts(Index).Y = Ghosts(Index).Y + 1
          ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = North
            Ghosts(Index).Y = Ghosts(Index).Y - 1
          ElseIf (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
            Ghosts(Index).Direction = South
            Ghosts(Index).Y = Ghosts(Index).Y + 1
          End If
        End If
      ElseIf (Ghosts(Index).Y < 228) And Ghosts(Index).Direction <> North Then
        'see if we can go South
        pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
        pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
        If (pixel_1 = BLACK_COLOR Or pixel_1 = door_color) And (pixel_2 = BLACK_COLOR Or pixel_2 = door_color) Then
          Ghosts(Index).Direction = South
          Ghosts(Index).Y = Ghosts(Index).Y + 1
        Else
          pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
          pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
          pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
          pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
          'see if we can go East or West
          If Ghosts(Index).Direction <> East And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = West
            Ghosts(Index).X = Ghosts(Index).X - 1
          ElseIf Ghosts(Index).Direction <> West And pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
            Ghosts(Index).Direction = East
            Ghosts(Index).X = Ghosts(Index).X + 1
          ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = West
            Ghosts(Index).X = Ghosts(Index).X - 1
          ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
            Ghosts(Index).Direction = East
            Ghosts(Index).X = Ghosts(Index).X + 1
          End If
        End If
      ElseIf (Ghosts(Index).X < 261) And Ghosts(Index).Direction <> West Then
        'see if we can go East
        pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
        pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
        If pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
          Ghosts(Index).Direction = East
          Ghosts(Index).X = Ghosts(Index).X + 1
        Else
          pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
          pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
          pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
          pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
          'see if we can go North or South
          If Ghosts(Index).Direction <> South And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = North
            Ghosts(Index).Y = Ghosts(Index).Y - 1
          ElseIf Ghosts(Index).Direction <> North And (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
            Ghosts(Index).Direction = South
            Ghosts(Index).Y = Ghosts(Index).Y + 1
          ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = North
            Ghosts(Index).Y = Ghosts(Index).Y - 1
          ElseIf (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
            Ghosts(Index).Direction = South
            Ghosts(Index).Y = Ghosts(Index).Y + 1
          End If
        End If
      ElseIf (Ghosts(Index).Y > 260) Then
        'then see if we can go North
        pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
        pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
        If pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR And Ghosts(Index).Direction <> South Then
          Ghosts(Index).Direction = North
          Ghosts(Index).Y = Ghosts(Index).Y - 1
        Else
          pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
          pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
          pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
          pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
          'see if we can go East or West
          If Ghosts(Index).Direction <> East And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = West
            Ghosts(Index).X = Ghosts(Index).X - 1
          ElseIf Ghosts(Index).Direction <> West And pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
            Ghosts(Index).Direction = East
            Ghosts(Index).X = Ghosts(Index).X + 1
          ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = West
            Ghosts(Index).X = Ghosts(Index).X - 1
          ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
            Ghosts(Index).Direction = East
            Ghosts(Index).X = Ghosts(Index).X + 1
          Else
            pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
            pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
            pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
            pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
            
            If Ghosts(Index).Direction = North And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
              Ghosts(Index).Direction = North
              Ghosts(Index).Y = Ghosts(Index).Y - 1
            ElseIf Ghosts(Index).Direction = South And (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
              Ghosts(Index).Direction = South
              Ghosts(Index).Y = Ghosts(Index).Y + 1
            ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
              Ghosts(Index).Direction = North
              Ghosts(Index).Y = Ghosts(Index).Y - 1
            ElseIf (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
              Ghosts(Index).Direction = South
              Ghosts(Index).Y = Ghosts(Index).Y + 1
            End If
        End If
      End If
    ElseIf (Ghosts(Index).X > 261) Then
      'see if we can go West
      pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
      pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
      If pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR And Ghosts(Index).Direction <> East Then
        Ghosts(Index).Direction = West
        Ghosts(Index).X = Ghosts(Index).X - 1
      Else
        pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
        pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
        pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
        pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
        'see if we can go North or South
        If Ghosts(Index).Direction <> South And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
          Ghosts(Index).Direction = North
          Ghosts(Index).Y = Ghosts(Index).Y - 1
        ElseIf Ghosts(Index).Direction <> North And (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
          Ghosts(Index).Direction = South
          Ghosts(Index).Y = Ghosts(Index).Y + 1
        ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
          Ghosts(Index).Direction = North
          Ghosts(Index).Y = Ghosts(Index).Y - 1
        ElseIf (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
          Ghosts(Index).Direction = South
          Ghosts(Index).Y = Ghosts(Index).Y + 1
        Else
            pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
            pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
            pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
            pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
            
            If Ghosts(Index).Direction = West And pixel_1 <> 16711680 And pixel_1 <> 526344 And pixel_2 <> 16711680 And pixel_2 <> 526344 Then
              Ghosts(Index).Direction = West
              Ghosts(Index).X = Ghosts(Index).X - 1
            ElseIf Ghosts(Index).Direction = East And pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
              Ghosts(Index).Direction = East
              Ghosts(Index).X = Ghosts(Index).X + 1
            ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
              Ghosts(Index).Direction = West
              Ghosts(Index).X = Ghosts(Index).X - 1
            ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
              Ghosts(Index).Direction = East
              Ghosts(Index).X = Ghosts(Index).X + 1
            End If
        End If
      End If
    ElseIf (Ghosts(Index).Y < 228) Then
      'see if we can go South
      pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
      pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
      If (pixel_1 = BLACK_COLOR Or pixel_1 = door_color) And (pixel_2 = BLACK_COLOR Or pixel_2 = door_color) And Ghosts(Index).Direction <> North Then
        Ghosts(Index).Direction = South
        Ghosts(Index).Y = Ghosts(Index).Y + 1
      Else
        pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
        pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
        pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
        pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
        'see if we can go East or West
        If Ghosts(Index).Direction <> East And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
          Ghosts(Index).Direction = West
          Ghosts(Index).X = Ghosts(Index).X - 1
        ElseIf Ghosts(Index).Direction <> West And pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
          Ghosts(Index).Direction = East
          Ghosts(Index).X = Ghosts(Index).X + 1
        ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
          Ghosts(Index).Direction = West
          Ghosts(Index).X = Ghosts(Index).X - 1
        ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
          Ghosts(Index).Direction = East
          Ghosts(Index).X = Ghosts(Index).X + 1
        Else
            pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
            pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
            pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
            pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
            If Ghosts(Index).Direction = South And (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
              Ghosts(Index).Direction = South
              Ghosts(Index).Y = Ghosts(Index).Y + 1
            ElseIf Ghosts(Index).Direction = North And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
              Ghosts(Index).Direction = North
              Ghosts(Index).Y = Ghosts(Index).Y - 1
            ElseIf (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
              Ghosts(Index).Direction = South
              Ghosts(Index).Y = Ghosts(Index).Y + 1
            ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
              Ghosts(Index).Direction = North
              Ghosts(Index).Y = Ghosts(Index).Y - 1
            End If
        End If
      End If
    ElseIf (Ghosts(Index).X < 261) Then
      'see if we can go East
      pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
      pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
      If pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR And Ghosts(Index).Direction <> West Then
        Ghosts(Index).Direction = East
        Ghosts(Index).X = Ghosts(Index).X + 1
      Else
        pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
        pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
        pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
        pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
        'see if we can go North or South
        If Ghosts(Index).Direction <> South And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
          Ghosts(Index).Direction = North
          Ghosts(Index).Y = Ghosts(Index).Y - 1
        ElseIf Ghosts(Index).Direction <> North And (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
          Ghosts(Index).Direction = South
          Ghosts(Index).Y = Ghosts(Index).Y + 1
        ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
          Ghosts(Index).Direction = North
          Ghosts(Index).Y = Ghosts(Index).Y - 1
        ElseIf (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
          Ghosts(Index).Direction = South
          Ghosts(Index).Y = Ghosts(Index).Y + 1
        Else
            pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
            pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
            pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
            pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
            If Ghosts(Index).Direction = East And pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
              Ghosts(Index).Direction = East
              Ghosts(Index).X = Ghosts(Index).X + 1
            ElseIf Ghosts(Index).Direction = West And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
              Ghosts(Index).Direction = West
              Ghosts(Index).X = Ghosts(Index).X - 1
            ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
              Ghosts(Index).Direction = East
              Ghosts(Index).X = Ghosts(Index).X + 1
            ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
              Ghosts(Index).Direction = West
              Ghosts(Index).X = Ghosts(Index).X - 1
            End If
        End If
      End If
    Else
      If (Ghosts(Index).Y < 260) Then
        'see if we can go South
        pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
        pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
        If (pixel_1 = BLACK_COLOR Or pixel_1 = door_color) And (pixel_2 = BLACK_COLOR Or pixel_2 = door_color) And Ghosts(Index).Direction <> North Then
          Ghosts(Index).Direction = South
          Ghosts(Index).Y = Ghosts(Index).Y + 1
        Else
          pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
          pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
          pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
          pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
          'see if we can go East or West
          If Ghosts(Index).Direction <> East And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = West
            Ghosts(Index).X = Ghosts(Index).X - 1
          ElseIf Ghosts(Index).Direction <> West And pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
            Ghosts(Index).Direction = East
            Ghosts(Index).X = Ghosts(Index).X + 1
          ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
            Ghosts(Index).Direction = West
            Ghosts(Index).X = Ghosts(Index).X - 1
          ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
            Ghosts(Index).Direction = East
            Ghosts(Index).X = Ghosts(Index).X + 1
          Else
            pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
            pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
            pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
            pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
            If Ghosts(Index).Direction = South And (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
              Ghosts(Index).Direction = South
              Ghosts(Index).Y = Ghosts(Index).Y + 1
            ElseIf Ghosts(Index).Direction = North And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
              Ghosts(Index).Direction = North
              Ghosts(Index).Y = Ghosts(Index).Y - 1
            ElseIf (pixel_3 = BLACK_COLOR Or pixel_3 = door_color) And (pixel_4 = BLACK_COLOR Or pixel_4 = door_color) Then
              Ghosts(Index).Direction = South
              Ghosts(Index).Y = Ghosts(Index).Y + 1
            ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
              Ghosts(Index).Direction = North
              Ghosts(Index).Y = Ghosts(Index).Y - 1
            End If
          End If
        End If
      Else 'We found the spot, restore the Ghost
        Ghosts(Index).Eaten = False
        Ghosts(Index).Free = False
      End If
    End If
  'This next bit frees the Ghost from its Home
  ElseIf Ghosts(Index).Eaten = False And _
            (Ghosts(Index).Free = False) And _
            (START_TIMER > (((Index - 1) * 1000) / (Level + 1))) Then
    If (Ghosts(Index).X < 261) Then
      Ghosts(Index).Direction = East
      Ghosts(Index).X = Ghosts(Index).X + 1
    ElseIf (Ghosts(Index).X > 261) Then
      Ghosts(Index).Direction = West
      Ghosts(Index).X = Ghosts(Index).X - 1
    ElseIf (Ghosts(Index).Y > 232) Then
      Ghosts(Index).Direction = North
      Ghosts(Index).Y = Ghosts(Index).Y - 1
    Else
      Ghosts(Index).Free = True
    End If
  Else 'We'll Move the Ghost, changing direction if necessary
    If (Ghosts(Index).X < 324 And Ghosts(Index).X > 212 And Ghosts(Index).Y < 337 And Ghosts(Index).Y > 260) Then
      'If they are still in the box, we don't want them
      'dancing around like they've had too much coffee
      seed = 550000
    Else
      seed = 4000
    End If
  
    rand_dir = Int(seed * Rnd())
    If (Ghosts(Index).X > 4) And (Ghosts(Index).X < (Picture1.Width - 23)) Then
      If (rand_dir < 3000) Or (pixel_1 <> BLACK_COLOR Or pixel_2 <> BLACK_COLOR) Then
        'We'll try to change direction if possible
        If (Ghosts(Index).Direction = North) Or (Ghosts(Index).Direction = South) Then
          'Try to go east or west
          pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
          pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
          pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
          pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
          If (Ghosts(Index).X = PACMAN.X) And _
           (((Ghosts(Index).Direction = North) And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR) Or _
            ((Ghosts(Index).Direction = South) And pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR)) Then
          Else
            pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
            pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
            pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
            pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
            If pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR And pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
              If (PACMAN.X >= Ghosts(Index).X) Then
                Ghosts(Index).Direction = East
              ElseIf (PACMAN.X < Ghosts(Index).X) Then
                Ghosts(Index).Direction = West
              End If
            ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
              Ghosts(Index).Direction = West
            ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
              Ghosts(Index).Direction = East
            Else
              'continue in the same direction
            End If
          End If
        Else
          'try to go north or south
          pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y)
          pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 16, Ghosts(Index).Y + 15)
          pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y)
          pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X - 1, Ghosts(Index).Y + 15)
          If (Ghosts(Index).Y = PACMAN.Y) And _
           (((Ghosts(Index).Direction = East) And pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR) Or _
            ((Ghosts(Index).Direction = West) And pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR)) Then
          Else
            pixel_1 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y - 1)
            pixel_2 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y - 1)
            pixel_3 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X, Ghosts(Index).Y + 16)
            pixel_4 = GetPixel(Dots.Maze.hDC, Ghosts(Index).X + 15, Ghosts(Index).Y + 16)
            If pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR And pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
              If (PACMAN.Y >= Ghosts(Index).Y) Then
                Ghosts(Index).Direction = South
              ElseIf (PACMAN.Y < Ghosts(Index).Y) Then
                Ghosts(Index).Direction = North
              End If
            ElseIf pixel_1 = BLACK_COLOR And pixel_2 = BLACK_COLOR Then
              Ghosts(Index).Direction = North
            ElseIf pixel_3 = BLACK_COLOR And pixel_4 = BLACK_COLOR Then
              Ghosts(Index).Direction = South
            Else
           
            End If 'If pixel_1 <> 16711680...
          End If
        End If 'If (ghosts(index).Direction = North) Or ...
      Else 'If (rand_dir = 50) Then
        
      End If 'If (rand_dir = 50) Then
    End If
    If (Ghosts(Index).Direction = North) Then
      Ghosts(Index).Y = Ghosts(Index).Y - 1
    ElseIf (Ghosts(Index).Direction = South) Then
      Ghosts(Index).Y = Ghosts(Index).Y + 1
    ElseIf (Ghosts(Index).Direction = East) Then
      Ghosts(Index).X = Ghosts(Index).X + 1
      If Ghosts(Index).X > (Picture1.Width + 15) Then Ghosts(Index).X = -15
    ElseIf (Ghosts(Index).Direction = West) Then
      Ghosts(Index).X = Ghosts(Index).X - 1
      If Ghosts(Index).X < -15 Then Ghosts(Index).X = Picture1.Width + 15
    End If
  End If
End If
End If
End Sub
Private Sub Erase_Ghosts()
Dim itr As Byte
BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, Dots.PosDots.hDC, PACMAN.X, PACMAN.Y, &HCC0020
BitBlt Picture1.hDC, Food.X, Food.Y, 16, 16, Dots.PosDots.hDC, Food.X, Food.Y, &HCC0020
If Food.Eaten = True Then
  If Food.Timer >= 0 And Food.Timer < 3000 Then  'Erase Score
    Restore_Food
  End If
ElseIf Food.Timer >= 0 And Food.Timer < 3000 Then
  If Food.Timer >= 0 And Food.Timer < 3 Then Restore_Food
End If

For itr = 1 To 6
  BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, Dots.PosDots.hDC, Ghosts(itr).X, Ghosts(itr).Y, &HCC0020
Next itr

If (FRAME_INDEX = 15) Then
 For itr = 1 To 6
  If Pellets(itr).Frame < 0 Then
    BitBlt Picture1.hDC, Pellets(itr).X, Pellets(itr).Y, 16, 16, Dots.PosDots.hDC, 239, 0, &HCC0020
    BitBlt Dots.PosDots.hDC, Pellets(itr).X, Pellets(itr).Y, 16, 16, Dots.PosDots.hDC, 239, 0, &HCC0020
  Else
    If Pellets(itr).Eaten = False Then
      BitBlt Picture1.hDC, Pellets(itr).X, Pellets(itr).Y, 16, 16, PAC_SOURCE, Pellets(itr).Src_X, Pellets(itr).Src_Y, &HCC0020
      BitBlt Dots.PosDots.hDC, Pellets(itr).X, Pellets(itr).Y, 16, 16, PAC_SOURCE, Pellets(itr).Src_X, Pellets(itr).Src_Y, &HCC0020
    End If
  End If
  Pellets(itr).Frame = -Pellets(itr).Frame
 Next itr
End If

If Food.Eaten = True Then
  If Food.Timer < 0 And Not Food.Special Then   'Draw Score
    BitBlt Picture1.hDC, Food.X, Food.Y, 16, 16, PAC_SOURCE, Food.Source_X + 51 + ((BONUS_MULTIPLIER - 1) * 34), Food.Source_Y, vbSrcAnd
    BitBlt Picture1.hDC, Food.X, Food.Y, 16, 16, PAC_SOURCE, Food.Source_X + 34 + ((BONUS_MULTIPLIER - 1) * 34), Food.Source_Y, vbSrcPaint
  End If
Else 'If Food.Timer = 3000 Or PACMAN.MS = True Then 'Draw Food
  If (PACMAN.MS = True) Then
    If (FRUIT_FRAME > 10) Then
      MoveFood
    End If
    Food.Timer = Food.Timer - 1
  End If
  BitBlt Picture1.hDC, Food.X, Food.Y, 16, 16, PAC_SOURCE, Food.Source_X + 17, Food.Source_Y, vbSrcAnd
  BitBlt Picture1.hDC, Food.X, Food.Y, 16, 16, PAC_SOURCE, Food.Source_X, Food.Source_Y, vbSrcPaint
End If

End Sub

Private Sub Detect_Collision()
Dim itr As Byte
Dim pellet_index As Byte
pellet_index = 0
'Detect Pellets collision first
If (PACMAN.X < 14) And (PACMAN.X > 0) Then
  If (PACMAN.Y < 47) Then
    If Pellets(1).Eaten = False Then
      pellet_index = 1
      Pellets(1).Eaten = True
      Score = Score + (BONUS_MULTIPLIER * 100)
      PELLET_COUNTER = PELLET_COUNTER + 1
      secret = ""
      For itr = 1 To 6
        If Ghosts(itr).Eaten = False Then Ghosts(itr).Scared = True
      Next itr
    End If
  ElseIf (PACMAN.Y > 291) And PACMAN.Y < 301 And PACMAN.X > -11 Then
    If Pellets(3).Eaten = False Then
      pellet_index = 3
      Pellets(3).Eaten = True
      Score = Score + (BONUS_MULTIPLIER * 100)
      PELLET_COUNTER = PELLET_COUNTER + 1
      secret = ""
      For itr = 1 To 6
        If Ghosts(itr).Eaten = False Then Ghosts(itr).Scared = True
      Next itr
    End If
  ElseIf (PACMAN.Y > 538) Then
    If Pellets(5).Eaten = False Then
      pellet_index = 5
      Pellets(5).Eaten = True
      Score = Score + (BONUS_MULTIPLIER * 100)
      PELLET_COUNTER = PELLET_COUNTER + 1
      secret = ""
      For itr = 1 To 6
        If Ghosts(itr).Eaten = False Then Ghosts(itr).Scared = True
      Next itr
    End If
  End If
ElseIf (PACMAN.X > Picture1.Width - 34) And (PACMAN.X < Picture1.Width - 24) Then
  If (PACMAN.Y < 50) Then
    If Pellets(2).Eaten = False Then
      pellet_index = 2
      Pellets(2).Eaten = True
      Score = Score + (BONUS_MULTIPLIER * 100)
      PELLET_COUNTER = PELLET_COUNTER + 1
      secret = ""
      For itr = 1 To 6
        If Ghosts(itr).Eaten = False Then Ghosts(itr).Scared = True
      Next itr
    End If
  ElseIf (PACMAN.Y > 291) And PACMAN.Y < 301 And PACMAN.X > -11 Then
    If Pellets(4).Eaten = False Then
      pellet_index = 4
      Pellets(4).Eaten = True
      Score = Score + (BONUS_MULTIPLIER * 100)
      PELLET_COUNTER = PELLET_COUNTER + 1
      secret = ""
      For itr = 1 To 6
        If Ghosts(itr).Eaten = False Then Ghosts(itr).Scared = True
      Next itr
    End If
  ElseIf (PACMAN.Y > 538) Then
    If Pellets(6).Eaten = False Then
      pellet_index = 6
      Pellets(6).Eaten = True
      Score = Score + (BONUS_MULTIPLIER * 100)
      PELLET_COUNTER = PELLET_COUNTER + 1
      secret = ""
      For itr = 1 To 6
        If Ghosts(itr).Eaten = False Then Ghosts(itr).Scared = True
      Next itr
    End If
  End If
End If

'Detect Food collision
If Food.Eaten = False Then
  If Food.X = PACMAN.X Then
    If Food.Y < PACMAN.Y Then
       If Food.Y > PACMAN.Y - 8 Then
         'collision
         Food.Eaten = True
         Food.Timer = -300
       End If
    ElseIf Food.Y > PACMAN.Y Then
       If Food.Y < PACMAN.Y + 8 Then
         'collision
         Food.Eaten = True
         Food.Timer = -300
       End If
    Else
      'collision
      Food.Eaten = True
      Food.Timer = -300
    End If
  ElseIf Food.Y = PACMAN.Y Then
    If Food.X < PACMAN.X Then
       If Food.X > PACMAN.X - 8 Then
         'collision
         Food.Eaten = True
         Food.Timer = -300
       End If
    ElseIf Food.X > PACMAN.X Then
       If Food.X < PACMAN.X + 8 Then
         'collision
         Food.Eaten = True
         Food.Timer = -300
       End If
    Else
      Food.Eaten = True
      Food.Timer = -300
    End If
  End If
  If Food.Eaten Then
    If Food.Special = True Then
      If (Food.SpecialType = 1) Then
        sndPlaySound 0, SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
        FROZEN = True
        FREEZE_TIMER = 1000
        Restore_Food
      ElseIf (Food.SpecialType = 2) Then
        Timer1.Interval = 0
        EXPLODE_FRAME = 0
        sndPlaySound "explode.wav", SND_NODEFAULT Or SND_ASYNC
        Timer3.Interval = 43
        Restore_Food
      ElseIf (Food.SpecialType = 3) Then
        sndPlaySound 0, SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
        FROZEN = True
        FREEZE_TIMER = 500
        Restore_Food
      ElseIf (Food.SpecialType = 4) Then
        PACMAN.lives = PACMAN.lives + 1
        Display_Lives
        sndPlaySound "pop.wav", SND_NODEFAULT Or SND_ASYNC
        Restore_Food
      ElseIf (Food.SpecialType = 5) Then
        LIGHTNING = True
        sndPlaySound "voltage.wav", SND_NODEFAULT Or SND_ASYNC
        Restore_Food
      ElseIf (Food.SpecialType = 6) Then
        CHILLED = True
        FREEZE_TIMER = 1000
        GHOSTS_PISSED = False
        sndPlaySound "ice.wav", SND_NODEFAULT Or SND_ASYNC
        Restore_Food
      ElseIf (Food.SpecialType = 7) Then
        Timer1.Interval = 0
        EXPLODE_FRAME = 0
        sndPlaySound "crackers.wav", SND_NODEFAULT Or SND_ASYNC
        Timer3.Interval = 43
        Restore_Food
      ElseIf (Food.SpecialType = 8) Then
        EXTRA = EXTRA + 300
        sndPlaySound "gulp.wav", SND_NODEFAULT Or SND_ASYNC
        If RUNNING_SCARED > 0 Then RUNNING_SCARED = RUNNING_SCARED + EXTRA
        Restore_Food
      ElseIf (Food.SpecialType = 9) Then
        sndPlaySound "poison.wav", SND_NODEFAULT Or SND_ASYNC
        PACMAN.Poisoned = True
        PACMAN.PoisonCounter = 500
        Restore_Food
      ElseIf (Food.SpecialType = 10) Then
        sndPlaySound "cashreg.wav", SND_NODEFAULT Or SND_ASYNC
        BONUS_MULTIPLIER = 2
        Restore_Food
      End If
    Else
      Score = Score + (BONUS_MULTIPLIER * Food.Value)
      sndPlaySound "fruit.wav", SND_NODEFAULT Or SND_ASYNC
    End If
  End If
End If

If pellet_index > 0 Then
  sndPlaySound "scaredstart.wav", SND_NODEFAULT Or SND_ASYNC 'Or SND_LOOP
  secret = ""
  BitBlt Picture1.hDC, Pellets(pellet_index).X, Pellets(pellet_index).Y, 16, 16, PAC_SOURCE, 239, 0, &HCC0020
  BitBlt Dots.PosDots.hDC, Pellets(pellet_index).X, Pellets(pellet_index).Y, 16, 16, PAC_SOURCE, 239, 0, &HCC0020
  'Reverse ghost direction
  DEATH_TOLL = 0
  GHOSTS_PISSED = False
  CHAR_SOURCE = Characters.Chars.hDC
  For itr = 1 To 6
    If Ghosts(itr).Eaten = False Then
      If Ghosts(itr).Free And Ghosts(itr).Direction = West Then
       Ghosts(itr).Direction = East
      ElseIf Ghosts(itr).Free And Ghosts(itr).Direction = East Then
       Ghosts(itr).Direction = West
      ElseIf Ghosts(itr).Free And Ghosts(itr).Direction = South Then
       Ghosts(itr).Direction = North
      ElseIf Ghosts(itr).Free And Ghosts(itr).Direction = North Then
       Ghosts(itr).Direction = South
      End If
    End If
  Next itr
  
  RUNNING_SCARED = SCARED_TIME + EXTRA
End If

'Detect Ghost Collisions
For itr = 1 To 6
  If Ghosts(itr).X = PACMAN.X Then
    If Ghosts(itr).Y < PACMAN.Y Then
       If Ghosts(itr).Y > PACMAN.Y - 8 Then
         'collision
         If (Ghosts(itr).Scared) Then
           secret = secret + Ghosts(itr).Letter
           Ghosts(itr).Scared = False
           Ghosts(itr).Eaten = True
           Ghosts(itr).JustEaten = True
           DEATH_TOLL = DEATH_TOLL + 1
           Score = Score + (BONUS_MULTIPLIER * ((2 ^ (DEATH_TOLL)) * 100))
           EATEN_TIMER = 70
           sndPlaySound "ghosts.wav", SND_NODEFAULT Or SND_ASYNC
         ElseIf (Ghosts(itr).Eaten) Then
           'ignore
         Else
           PACMAN.Dying = True
           PACMAN.Frame = -5
           PACMAN.lives = PACMAN.lives - 1
         End If
       End If
    ElseIf Ghosts(itr).Y > PACMAN.Y Then
       If Ghosts(itr).Y < PACMAN.Y + 8 Then
         'collision
         If (Ghosts(itr).Scared) Then
           secret = secret + Ghosts(itr).Letter
           Ghosts(itr).Scared = False
           Ghosts(itr).Eaten = True
           Ghosts(itr).JustEaten = True
           DEATH_TOLL = DEATH_TOLL + 1
           Score = Score + (BONUS_MULTIPLIER * ((2 ^ (DEATH_TOLL)) * 100))
           EATEN_TIMER = 70
           sndPlaySound "ghosts.wav", SND_NODEFAULT Or SND_ASYNC
         ElseIf (Ghosts(itr).Eaten) Then
           'ignore
         Else
           PACMAN.Dying = True
           PACMAN.Frame = -5
           PACMAN.lives = PACMAN.lives - 1
         End If
       End If
    Else
      If (Ghosts(itr).Scared) Then
        secret = secret + Ghosts(itr).Letter
        Ghosts(itr).Scared = False
        Ghosts(itr).Eaten = True
        Ghosts(itr).JustEaten = True
        DEATH_TOLL = DEATH_TOLL + 1
        Score = Score + (BONUS_MULTIPLIER * ((2 ^ (DEATH_TOLL)) * 100))
        EATEN_TIMER = 70
        sndPlaySound "ghosts.wav", SND_NODEFAULT Or SND_ASYNC
      ElseIf (Ghosts(itr).Eaten) Then
        'ignore
      Else
        PACMAN.Dying = True
        PACMAN.Frame = -5
        PACMAN.lives = PACMAN.lives - 1
      End If
    End If
  ElseIf Ghosts(itr).Y = PACMAN.Y Then
    If Ghosts(itr).X < PACMAN.X Then
       If Ghosts(itr).X > PACMAN.X - 8 Then
         'collision
         If (Ghosts(itr).Scared) Then
           secret = secret + Ghosts(itr).Letter
           Ghosts(itr).Scared = False
           Ghosts(itr).Eaten = True
           Ghosts(itr).JustEaten = True
           DEATH_TOLL = DEATH_TOLL + 1
           Score = Score + (BONUS_MULTIPLIER * ((2 ^ (DEATH_TOLL)) * 100))
           EATEN_TIMER = 70
           sndPlaySound "ghosts.wav", SND_NODEFAULT Or SND_ASYNC
         ElseIf (Ghosts(itr).Eaten) Then
           'ignore
         Else
           PACMAN.Dying = True
           PACMAN.Frame = -5
           PACMAN.lives = PACMAN.lives - 1
         End If
       End If
    ElseIf Ghosts(itr).X > PACMAN.X Then
       If Ghosts(itr).X < PACMAN.X + 8 Then
         'collision
         If (Ghosts(itr).Scared) Then
           secret = secret + Ghosts(itr).Letter
           Ghosts(itr).Scared = False
           Ghosts(itr).Eaten = True
           Ghosts(itr).JustEaten = True
           DEATH_TOLL = DEATH_TOLL + 1
           Score = Score + (BONUS_MULTIPLIER * ((2 ^ (DEATH_TOLL)) * 100))
           EATEN_TIMER = 70
           sndPlaySound "ghosts.wav", SND_NODEFAULT Or SND_ASYNC
         ElseIf (Ghosts(itr).Eaten) Then
           'ignore
         Else
           PACMAN.Dying = True
           PACMAN.Frame = -5
           PACMAN.lives = PACMAN.lives - 1
         End If
       End If
    Else
      If (Ghosts(itr).Scared) Then
        secret = secret + Ghosts(itr).Letter
        Ghosts(itr).Scared = False
        Ghosts(itr).Eaten = True
        DEATH_TOLL = DEATH_TOLL + 1
        Score = Score + (BONUS_MULTIPLIER * ((2 ^ (DEATH_TOLL)) * 100))
        EATEN_TIMER = 70
        sndPlaySound "ghosts.wav", SND_NODEFAULT Or SND_ASYNC
      ElseIf (Ghosts(itr).Eaten) Then
        'ignore
      Else
        PACMAN.Dying = True
        PACMAN.Frame = -5
        PACMAN.lives = PACMAN.lives - 1
      End If
    End If
  End If
Next itr
text1.Caption = secret
If DEATH_TOLL >= 6 Then
  GHOSTS_PISSED = True
  CHAR_SOURCE = Characters.Chars2.hDC
  RUNNING_SCARED = 10
ElseIf DEATH_TOLL > 0 Then
  Dim one_left As Boolean
  one_left = False
  For itr = 1 To 6
    If Ghosts(itr).Scared = True Then one_left = True
  Next itr
  If Not one_left Then RUNNING_SCARED = 10
End If
If secret = "PACMAN" Then
  secret = ""
  PACMAN.lives = PACMAN.lives + 1
  Display_Lives
End If
End Sub
Public Sub Initialize()
Dim itr As Byte

LIGHTNING = False
CHILLED = False
EXTRA = 0
BONUS_MULTIPLIER = 1

FROZEN = False
FREEZE_TIMER = 0
GHOSTS_PISSED = False
CHAR_SOURCE = Characters.Chars.hDC
PAUSE = False
BEGIN_TIMER = 0
CHANGE_FRAME = 1
SLOW_DOWN = 1
DIRECTION_KEY = 39
FRAME_INDEX = 0
EATEN_TIMER = 0
START_TIMER = 0
STARTED = False
secret = ""

RUNNING_SCARED = -1
PACMAN.Direction = East
PACMAN.X = 261
PACMAN.Y = 356

PACMAN.Y_vector = 0
PACMAN.X_vector = 1

PACMAN.TURBO = False
PACMAN.Frame = 0
PACMAN.Dying = False
PACMAN.frame_val = 1
PACMAN.Poisoned = False
PACMAN.PoisonCounter = 0

If PACMAN.MS = False Then
  PACMAN.Src_Y = 0
Else
  PACMAN.Src_Y = 612
End If

For itr = 1 To 6
   Ghosts(itr).Frame = 0
   Ghosts(itr).Eaten = False
   Ghosts(itr).JustEaten = False
   Ghosts(itr).Scared = False
   Ghosts(itr).Free = False
   Ghosts(itr).Direction = North
Next itr

Ghosts(1).X = 251
Ghosts(1).Y = 260
Ghosts(1).Src_X = 1
Ghosts(1).Src_Y = 68
Ghosts(1).Letter = "P"
Ghosts(2).X = 214
Ghosts(2).Y = 260
Ghosts(2).Src_X = 1
Ghosts(2).Src_Y = 136
Ghosts(2).Letter = "A"
Ghosts(3).X = 223
Ghosts(3).Y = 314
Ghosts(3).Src_X = 1
Ghosts(3).Src_Y = 204
Ghosts(3).Letter = "C"
Ghosts(4).X = 262
Ghosts(4).Y = 314
Ghosts(4).Src_X = 1
Ghosts(4).Src_Y = 272
Ghosts(4).Letter = "M"
Ghosts(5).X = 282
Ghosts(5).Y = 260
Ghosts(5).Src_X = 1
Ghosts(5).Src_Y = 340
Ghosts(5).Letter = "A"
Ghosts(6).X = 305
Ghosts(6).Y = 314
Ghosts(6).Src_X = 1
Ghosts(6).Src_Y = 408
Ghosts(6).Letter = "N"
End Sub
Public Sub Build_Maze()
Dim itr As Byte
GHOSTS_PISSED = False
DEATH_TOLL = 0
CHAR_SOURCE = Characters.Chars.hDC
BitBlt Dots.Maze.hDC, 0, 0, Dots.Maze.Width, Dots.Maze.Height, levels(Level).MazeMap, 0, 0, &HCC0020
BitBlt Dots.DotsBegin.hDC, 0, 0, Dots.DotsBegin.Width, Dots.DotsBegin.Height, levels(Level).PositiveDots, 0, 0, &HCC0020
BitBlt Dots.PosDots.hDC, 0, 0, Dots.PosDots.Width, Dots.PosDots.Height, levels(Level).PositiveDots, 0, 0, &HCC0020
BitBlt Dots.NegDots.hDC, 0, 0, Dots.NegDots.Width, Dots.NegDots.Height, levels(Level).NegativeDots, 0, 0, &HCC0020
BitBlt Dots.Maze2.hDC, 0, 0, Dots.Maze2.Width, Dots.Maze2.Height, levels(Level).DoneMap, 0, 0, &HCC0020

DOT_MAX = levels(Level).DotMax

BitBlt Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, Dots.Maze.hDC, 0, 0, &HCC0020
BitBlt Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, Dots.NegDots.hDC, 0, 0, vbSrcAnd
BitBlt Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, Dots.DotsBegin.hDC, 0, 0, vbSrcPaint
BitBlt Dots.PosDots.hDC, 0, 0, Dots.PosDots.Width, Dots.PosDots.Height, Dots.DotsBegin.hDC, 0, 0, &HCC0020

For itr = 1 To 6
   Pellets(itr).Points = 100
   Pellets(itr).Src_X = 188
   Pellets(itr).Src_Y = 0
   Pellets(itr).Eaten = False
   Pellets(itr).Frame = -1
   If Ghosts(itr).Speed > 14 Then Ghosts(itr).Speed = Ghosts(itr).Speed - 1
Next itr
DOT_COUNTER = 0
PELLET_COUNTER = 0
Pellets(1).X = 5
Pellets(1).Y = 36
Pellets(2).X = 517
Pellets(2).Y = 36
Pellets(3).X = 5
Pellets(3).Y = 291
Pellets(4).X = 517
Pellets(4).Y = 291
Pellets(5).X = 5
Pellets(5).Y = 548
Pellets(6).X = 517
Pellets(6).Y = 548

If (Level < 11) Then SCARED_TIME = (Abs(1000 - (100 * Level))) + 20
Display_Lives
Display_Level
End Sub
Public Sub Display_Lives()
Dim itr As Integer

'erase previous
For itr = 0 To PACMAN.lives
  BitBlt Picture1.hDC, 5 + (17 * itr), 0, 16, 16, PAC_SOURCE, 239, 0, &HCC0020
Next itr

'draw current
For itr = 0 To PACMAN.lives - 2
  BitBlt Picture1.hDC, 5 + (17 * itr), 0, 16, 16, PAC_SOURCE, 52, PACMAN.Src_Y, &HCC0020
Next itr

End Sub
Public Sub Display_Level()
Dim itr As Integer
  For itr = 1 To Level
      BitBlt Picture1.hDC, (Picture1.Width - 4) - (17 * itr), 0, 16, 16, PAC_SOURCE, 188, 289 + ((itr - 1) * 17), &HCC0020
  Next itr
End Sub

Public Sub Restore_Food()
BitBlt Picture1.hDC, Food.X, Food.Y, 16, 16, Dots.PosDots.hDC, Food.X, Food.Y, &HCC0020
Food.Eaten = True
FRUIT_FRAME = 0
Dim rnd_num As Integer
If PACMAN.MS Then
Dim temp As Integer
  rnd_num = Int(10 * Rnd())
  Select Case Level
  Case 1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 13, 14
    If (rnd_num > 4) Then
      Food.X_Start = -16
      Food.X_Dest = Picture1.Width
      Food.Direction = East
    Else
      Food.X_Dest = -16
      Food.X_Start = Picture1.Width
      Food.Direction = West
    End If
    Food.Y_Start = 292
    Food.Y_Dest = 292
    Food.X = Food.X_Start
    Food.Y = Food.Y_Start
  Case 9, 10, 15
    If (rnd_num > 7) Then
      Food.X_Start = -16
      Food.X_Dest = Picture1.Width
      Food.Direction = East
      Food.Y_Start = 164
      Food.Y_Dest = Picture1.Height - 148
    ElseIf (rnd_num > 5) Then
      Food.X_Dest = -16
      Food.X_Start = Picture1.Width
      Food.Direction = West
      Food.Y_Start = 164
      Food.Y_Dest = Picture1.Height - 148
    ElseIf (rnd_num > 3) Then
      Food.X_Dest = -16
      Food.X_Start = Picture1.Width
      Food.Direction = West
      Food.Y_Start = Picture1.Height - 148
      Food.Y_Dest = 164
    Else
      Food.X_Start = -16
      Food.X_Dest = Picture1.Width
      Food.Direction = East
      Food.Y_Start = Picture1.Height - 148
      Food.Y_Dest = 164
    End If
    
    Food.X = Food.X_Start
    Food.Y = Food.Y_Start
  End Select
Else
  Food.X = 262
  Food.Y = 356
End If
  If Level < 5 Then
    Food.Value = Level * 100 + ((Level - 1) * 100)
  ElseIf Level < 8 Then
    Food.Value = (Level - 4) * 1000
  ElseIf Level = 8 Then
    Food.Value = 5000
  ElseIf Level = 9 Then
    Food.Value = 7000
  ElseIf Level = 10 Then
    Food.Value = 10000
  ElseIf Level = 11 Then
    Food.Value = 15000
  ElseIf Level = 12 Then
    Food.Value = 20000
  ElseIf Level = 13 Then
    Food.Value = 25000
  Else
    Food.Value = 50000
  End If
rnd_num = Int(20 * Rnd())
  Select Case rnd_num
    Case 1
      Food.Special = True
      Food.Source_X = 1
      Food.Source_Y = 510
      Food.SpecialType = 1
    Case 2
      Food.Special = True
      Food.Source_X = 35
      Food.Source_Y = 510
      Food.SpecialType = 2
    Case 3
      Food.Special = True
      Food.Source_X = 69
      Food.Source_Y = 510
      Food.SpecialType = 3
    Case 4
      Food.Special = True
      Food.Source_X = 103
      Food.Source_Y = 510
      Food.SpecialType = 4
    Case 5
      Food.Special = True
      Food.Source_X = 137
      Food.Source_Y = 510
      Food.SpecialType = 5
    Case 6
      Food.Special = True
      Food.Source_X = 1
      Food.Source_Y = 527
      Food.SpecialType = 6
    Case 7
      Food.Special = True
      Food.Source_X = 35
      Food.Source_Y = 527
      Food.SpecialType = 7
    Case 8
      Food.Special = True
      Food.Source_X = 69
      Food.Source_Y = 527
      Food.SpecialType = 8
    Case 9
      Food.Special = True
      Food.Source_X = 103
      Food.Source_Y = 527
      Food.SpecialType = 9
    Case 10
      Food.Special = True
      Food.Source_X = 137
      Food.Source_Y = 527
      Food.SpecialType = 10
    Case Else
      Food.Special = False
      Food.Source_X = 188
      Food.Source_Y = 289 + ((Level - 1) * 17)
  End Select

End Sub

Public Sub Draw_TURBO()
Dim itr As Integer

For itr = 0 To 60
  BitBlt Picture1.hDC, 225 + itr, 3, 1, 6, PAC_SOURCE, 239, 0, &HCC0020
Next itr

For itr = 0 To Int(TURBO / 10)
  If LIGHTNING Then
    BitBlt Picture1.hDC, 225 + itr, 3, 1, 6, PAC_SOURCE, 239, 69, &HCC0020
  Else
    BitBlt Picture1.hDC, 225 + itr, 3, 1, 6, PAC_SOURCE, 239, 52, &HCC0020
  End If
Next itr
End Sub

Public Sub Startup()
Dim itr As Integer
DOT_COLOR = GetPixel(PAC_SOURCE, 240, 69)
DOOR = GetPixel(PAC_SOURCE, 240, 85)
WALL_COLOR = GetPixel(PAC_SOURCE, 240, 107)
BLACK_COLOR = GetPixel(PAC_SOURCE, 240, 0)
ALMOST_BLACK = GetPixel(PAC_SOURCE, 240, 17)
Me.lblStart.Visible = False
STARTED = False
Timer2.Interval = 0
TIME_TIME = 1
Score = 0
Level = 1
PACMAN.lives = 3
RECEIVED_BONUS = False
Food.Timer = 0

Food.X = 258
Food.Y = 356

levels(1).DoneMap = Mazes.MazeA(0).hDC
levels(1).MazeMap = Mazes.Maze(0).hDC
levels(1).NegativeDots = Mazes.NegDots(0).hDC
levels(1).PositiveDots = Mazes.PosDots(0).hDC
levels(1).DotMax = 451

levels(2).DoneMap = Mazes.MazeA(0).hDC
levels(2).MazeMap = Mazes.Maze(0).hDC
levels(2).NegativeDots = Mazes.NegDots(0).hDC
levels(2).PositiveDots = Mazes.PosDots(0).hDC
levels(2).DotMax = 451

levels(3).DoneMap = Mazes.MazeA(1).hDC
levels(3).MazeMap = Mazes.Maze(1).hDC
levels(3).NegativeDots = Mazes.NegDots(1).hDC
levels(3).PositiveDots = Mazes.PosDots(1).hDC
levels(3).DotMax = 410

levels(4).DoneMap = Mazes.MazeA(1).hDC
levels(4).MazeMap = Mazes.Maze(1).hDC
levels(4).NegativeDots = Mazes.NegDots(1).hDC
levels(4).PositiveDots = Mazes.PosDots(1).hDC
levels(4).DotMax = 410

levels(5).DoneMap = Mazes.MazeA(2).hDC
levels(5).MazeMap = Mazes.Maze(2).hDC
levels(5).NegativeDots = Mazes.NegDots(2).hDC
levels(5).PositiveDots = Mazes.PosDots(2).hDC
levels(5).DotMax = 606

levels(6).DoneMap = Mazes.MazeA(2).hDC
levels(6).MazeMap = Mazes.Maze(2).hDC
levels(6).NegativeDots = Mazes.NegDots(2).hDC
levels(6).PositiveDots = Mazes.PosDots(2).hDC
levels(6).DotMax = 606

levels(7).DoneMap = Mazes.MazeA(3).hDC
levels(7).MazeMap = Mazes.Maze(3).hDC
levels(7).NegativeDots = Mazes.NegDots(3).hDC
levels(7).PositiveDots = Mazes.PosDots(3).hDC
levels(7).DotMax = 595

levels(8).DoneMap = Mazes.MazeA(3).hDC
levels(8).MazeMap = Mazes.Maze(3).hDC
levels(8).NegativeDots = Mazes.NegDots(3).hDC
levels(8).PositiveDots = Mazes.PosDots(3).hDC
levels(8).DotMax = 595

levels(9).DoneMap = Mazes.MazeA(4).hDC
levels(9).MazeMap = Mazes.Maze(4).hDC
levels(9).NegativeDots = Mazes.NegDots(4).hDC
levels(9).PositiveDots = Mazes.PosDots(4).hDC
levels(9).DotMax = 643

levels(10).DoneMap = Mazes.MazeA(4).hDC
levels(10).MazeMap = Mazes.Maze(4).hDC
levels(10).NegativeDots = Mazes.NegDots(4).hDC
levels(10).PositiveDots = Mazes.PosDots(4).hDC
levels(10).DotMax = 643

levels(11).DoneMap = Mazes.MazeA(0).hDC
levels(11).MazeMap = Mazes.Maze(0).hDC
levels(11).NegativeDots = Mazes.NegDots(0).hDC
levels(11).PositiveDots = Mazes.PosDots(0).hDC
levels(11).DotMax = 451

levels(12).DoneMap = Mazes.MazeA(1).hDC
levels(12).MazeMap = Mazes.Maze(1).hDC
levels(12).NegativeDots = Mazes.NegDots(1).hDC
levels(12).PositiveDots = Mazes.PosDots(1).hDC
levels(12).DotMax = 410

levels(13).DoneMap = Mazes.MazeA(2).hDC
levels(13).MazeMap = Mazes.Maze(2).hDC
levels(13).NegativeDots = Mazes.NegDots(2).hDC
levels(13).PositiveDots = Mazes.PosDots(2).hDC
levels(13).DotMax = 606

levels(14).DoneMap = Mazes.MazeA(3).hDC
levels(14).MazeMap = Mazes.Maze(3).hDC
levels(14).NegativeDots = Mazes.NegDots(3).hDC
levels(14).PositiveDots = Mazes.PosDots(3).hDC
levels(14).DotMax = 595

levels(15).DoneMap = Mazes.MazeA(4).hDC
levels(15).MazeMap = Mazes.Maze(4).hDC
levels(15).NegativeDots = Mazes.NegDots(4).hDC
levels(15).PositiveDots = Mazes.PosDots(4).hDC
levels(15).DotMax = 643

lblTurbo.Visible = True
Me.lblScoreHeader.Visible = True
lblScore.Visible = True

Initialize
Build_Maze
Restore_Food

For itr = 1 To 6
  Ghosts(itr).Speed = 15
Next itr
Ghosts(1).Speed = 14
SCARED_TIME = 1100
TURBO = 0
Picture1.Visible = True
Picture1.Refresh
End Sub

Private Sub Timer2_Timer()
Dim itr As Integer
  Select Case INTRO_COUNTER
    Case 0
      If TITLE_Y > 50 Then
        BitBlt Picture1.hDC, TITLE_X, TITLE_Y, Intro.NegTitle.Width, Intro.NegTitle.Height, Intro.NegTitle.hDC, 0, 0, &HCC0020
        TITLE_Y = TITLE_Y - 2
        BitBlt Picture1.hDC, TITLE_X, TITLE_Y, Intro.NegTitle.Width, Intro.NegTitle.Height, Intro.Title.hDC, 0, 0, &HCC0020
      Else
        BitBlt Picture1.hDC, TITLE_X + 15, TITLE_Y - 30, Intro.pTurbo.Width, Intro.pTurbo.Height, Intro.pNegTurbo.hDC, 0, 0, vbSrcAnd
        BitBlt Picture1.hDC, TITLE_X + 15, TITLE_Y - 30, Intro.pTurbo.Width, Intro.pTurbo.Height, Intro.pTurbo.hDC, 0, 0, vbSrcPaint
        INTRO_COUNTER = INTRO_COUNTER + 1
      End If
      Picture1.Refresh
    Case 1
      PACMAN.Direction = East
      PACMAN.X = -16
      PACMAN.Y = 261
      PACMAN.Y_vector = 0
      PACMAN.X_vector = 1
      PACMAN.TURBO = False
      PACMAN.Frame = 0
      PACMAN.Dying = False
      PACMAN.frame_val = 1
      PACMAN.Src_Y = 0
      PACMAN.MS = False
      PACMAN.Poisoned = False
      PACMAN.PoisonCounter = 0

      MSPACMAN.Direction = West
      MSPACMAN.X = Picture1.Width
      MSPACMAN.Y = 261
      MSPACMAN.Y_vector = 0
      MSPACMAN.X_vector = -1
      MSPACMAN.TURBO = False
      MSPACMAN.Frame = 0
      MSPACMAN.Dying = False
      MSPACMAN.frame_val = 1
      MSPACMAN.Src_Y = 612
      MSPACMAN.MS = False
      MSPACMAN.Poisoned = False
      MSPACMAN.PoisonCounter = 0

      BitBlt Picture1.hDC, (Picture1.Width / 2) - 70, 220, Intro.pStarring.Width, Intro.pStarring.Height, Intro.pNegStarring.hDC, 0, 0, vbSrcAnd
      BitBlt Picture1.hDC, (Picture1.Width / 2) - 70, 220, Intro.pStarring.Width, Intro.pStarring.Height, Intro.pStarring.hDC, 0, 0, vbSrcPaint
      INTRO_COUNTER = INTRO_COUNTER + 1
    Case 2
      If (PACMAN.X < ((Picture1.Width / 2) - 30)) Then
          BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, PAC_SOURCE, 239, 0, &HCC0020
          BitBlt Picture1.hDC, MSPACMAN.X, MSPACMAN.Y, 16, 16, PAC_SOURCE, 239, 0, &HCC0020

          MSPACMAN.X = MSPACMAN.X - 1
          PACMAN.X = PACMAN.X + 1

          BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, PAC_SOURCE, ((PACMAN.Frame + 5) * 16) + (PACMAN.Frame + 5), PACMAN.Src_Y + (PACMAN.Direction * 16) + PACMAN.Direction, vbSrcAnd
          BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, PAC_SOURCE, (PACMAN.Frame * 16) + (PACMAN.Frame + 1), PACMAN.Src_Y + (PACMAN.Direction * 16) + PACMAN.Direction, vbSrcPaint
          
          BitBlt Picture1.hDC, MSPACMAN.X, MSPACMAN.Y, 16, 16, PAC_SOURCE, ((MSPACMAN.Frame + 5) * 16) + (MSPACMAN.Frame + 5), MSPACMAN.Src_Y + (MSPACMAN.Direction * 16) + MSPACMAN.Direction, vbSrcAnd
          BitBlt Picture1.hDC, MSPACMAN.X, MSPACMAN.Y, 16, 16, PAC_SOURCE, (MSPACMAN.Frame * 16) + (MSPACMAN.Frame + 1), MSPACMAN.Src_Y + (MSPACMAN.Direction * 16) + MSPACMAN.Direction, vbSrcPaint
          
          PACMAN.Frame = PACMAN.Frame + PACMAN.frame_val
          If PACMAN.Frame = 4 Or PACMAN.Frame = 0 Then
            PACMAN.frame_val = -PACMAN.frame_val
          End If
                    
          MSPACMAN.Frame = MSPACMAN.Frame + MSPACMAN.frame_val
          If MSPACMAN.Frame = 4 Or MSPACMAN.Frame = 0 Then
            MSPACMAN.frame_val = -MSPACMAN.frame_val
          End If
                     
          Picture1.Refresh
      Else
        INTRO_COUNTER = INTRO_COUNTER + 1
      End If
    Case 3
      For itr = 1 To 6
        Ghosts(itr).Frame = 0
        Ghosts(itr).Eaten = False
        Ghosts(itr).JustEaten = False
        Ghosts(itr).Scared = False
        Ghosts(itr).Free = False
        Ghosts(itr).Direction = East
        Ghosts(itr).X = -16
        Ghosts(itr).Y = 400
      Next itr

      Ghosts(1).Src_X = 1
      Ghosts(1).Src_Y = 68
      Ghosts(1).Letter = "P"

      Ghosts(2).Src_X = 1
      Ghosts(2).Src_Y = 136
      Ghosts(2).Letter = "A"

      Ghosts(3).Src_X = 1
      Ghosts(3).Src_Y = 204
      Ghosts(3).Letter = "C"

      Ghosts(4).Src_X = 1
      Ghosts(4).Src_Y = 272
      Ghosts(4).Letter = "M"

      Ghosts(5).Src_X = 1
      Ghosts(5).Src_Y = 340
      Ghosts(5).Letter = "A"

      Ghosts(6).Src_X = 1
      Ghosts(6).Src_Y = 408
      Ghosts(6).Letter = "N"
      INTRO_COUNTER = INTRO_COUNTER + 1
      Timer2.Interval = 400
    Case 4
      INTRO_COUNTER = INTRO_COUNTER + 1
    Case 5
      BitBlt Picture1.hDC, (Picture1.Width / 2) - 28, 320, Intro.pWith.Width, Intro.pWith.Height, Intro.pWith.hDC, 0, 0, &HCC0020
      Picture1.Refresh
      INTRO_COUNTER = INTRO_COUNTER + 1
      
    Case 6
      INTRO_COUNTER = INTRO_COUNTER + 1
    Case 7
      Timer2.Interval = 1
      INTRO_COUNTER = INTRO_COUNTER + 1
    Case 8
      If Ghosts(1).X < 435 Then
        BitBlt Picture1.hDC, Ghosts(1).X, Ghosts(1).Y, 16, 16, PAC_SOURCE, 239, 0, &HCC0020
        Ghosts(1).X = Ghosts(1).X + 2
        BitBlt Picture1.hDC, Ghosts(1).X, Ghosts(1).Y, 16, 16, PAC_SOURCE, 171 + (Ghosts(1).Frame * 17), 51, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(1).X, Ghosts(1).Y, 16, 16, PAC_SOURCE, Ghosts(1).Src_X + (Ghosts(1).Frame * 17), Ghosts(1).Src_Y + (Ghosts(1).Direction * 17), vbSrcPaint
        Ghosts(1).Frame = Ghosts(1).Frame + 1
        If Ghosts(1).Frame > 2 Then Ghosts(1).Frame = 0
        
      Else
        INTRO_COUNTER = INTRO_COUNTER + 1
        BitBlt Picture1.hDC, Ghosts(1).X - 20, Ghosts(1).Y - 35, Intro.pSlinky.Width, Intro.pSlinky.Height, Intro.pNegSlinky.hDC, 0, 0, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(1).X - 20, Ghosts(1).Y - 35, Intro.pSlinky.Width, Intro.pSlinky.Height, Intro.pSlinky.hDC, 0, 0, vbSrcPaint
      End If
      Picture1.Refresh
    Case 9
      If Ghosts(2).X < 365 Then
        BitBlt Picture1.hDC, Ghosts(2).X, Ghosts(2).Y, 16, 16, PAC_SOURCE, 239, 0, &HCC0020
        Ghosts(2).X = Ghosts(2).X + 2
        BitBlt Picture1.hDC, Ghosts(2).X, Ghosts(2).Y, 16, 16, PAC_SOURCE, 171 + (Ghosts(2).Frame * 17), 51, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(2).X, Ghosts(2).Y, 16, 16, PAC_SOURCE, Ghosts(2).Src_X + (Ghosts(2).Frame * 17), Ghosts(2).Src_Y + (Ghosts(2).Direction * 17), vbSrcPaint
        Ghosts(2).Frame = Ghosts(2).Frame + 1
        If Ghosts(2).Frame > 2 Then Ghosts(2).Frame = 0
      Else
        INTRO_COUNTER = INTRO_COUNTER + 1
        BitBlt Picture1.hDC, Ghosts(2).X - 20, Ghosts(2).Y - 35, Intro.pHinky.Width, Intro.pHinky.Height, Intro.pNegHinky.hDC, 0, 0, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(2).X - 20, Ghosts(2).Y - 35, Intro.pHinky.Width, Intro.pHinky.Height, Intro.pHinky.hDC, 0, 0, vbSrcPaint
      End If
      Picture1.Refresh
    Case 10
      itr = 3
      If Ghosts(itr).X < 295 Then
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, 239, 0, &HCC0020
        Ghosts(itr).X = Ghosts(itr).X + 2
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, 171 + (Ghosts(itr).Frame * 17), 51, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, Ghosts(itr).Src_X + (Ghosts(itr).Frame * 17), Ghosts(itr).Src_Y + (Ghosts(itr).Direction * 17), vbSrcPaint
        Ghosts(itr).Frame = Ghosts(itr).Frame + 1
        If Ghosts(itr).Frame > 2 Then Ghosts(itr).Frame = 0
      Else
        INTRO_COUNTER = INTRO_COUNTER + 1
        BitBlt Picture1.hDC, Ghosts(itr).X - 20, Ghosts(itr).Y - 35, Intro.pKinky.Width, Intro.pKinky.Height, Intro.pNegKinky.hDC, 0, 0, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(itr).X - 20, Ghosts(itr).Y - 35, Intro.pKinky.Width, Intro.pKinky.Height, Intro.pKinky.hDC, 0, 0, vbSrcPaint
      End If
      Picture1.Refresh
    Case 11
      itr = 4
      If Ghosts(itr).X < 225 Then
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, 239, 0, &HCC0020
        Ghosts(itr).X = Ghosts(itr).X + 2
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, 171 + (Ghosts(itr).Frame * 17), 51, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, Ghosts(itr).Src_X + (Ghosts(itr).Frame * 17), Ghosts(itr).Src_Y + (Ghosts(itr).Direction * 17), vbSrcPaint
        Ghosts(itr).Frame = Ghosts(itr).Frame + 1
        If Ghosts(itr).Frame > 2 Then Ghosts(itr).Frame = 0
      Else
        INTRO_COUNTER = INTRO_COUNTER + 1
        BitBlt Picture1.hDC, Ghosts(itr).X - 20, Ghosts(itr).Y - 35, Intro.pWinky.Width, Intro.pWinky.Height, Intro.pNegWinky.hDC, 0, 0, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(itr).X - 20, Ghosts(itr).Y - 35, Intro.pWinky.Width, Intro.pWinky.Height, Intro.pWinky.hDC, 0, 0, vbSrcPaint
      End If
      Picture1.Refresh
    Case 12
      itr = 5
      If Ghosts(itr).X < 155 Then
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, 239, 0, &HCC0020
        Ghosts(itr).X = Ghosts(itr).X + 2
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, 171 + (Ghosts(itr).Frame * 17), 51, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, Ghosts(itr).Src_X + (Ghosts(itr).Frame * 17), Ghosts(itr).Src_Y + (Ghosts(itr).Direction * 17), vbSrcPaint
        Ghosts(itr).Frame = Ghosts(itr).Frame + 1
        If Ghosts(itr).Frame > 2 Then Ghosts(itr).Frame = 0
      Else
        INTRO_COUNTER = INTRO_COUNTER + 1
        BitBlt Picture1.hDC, Ghosts(itr).X - 20, Ghosts(itr).Y - 35, Intro.pDinky.Width, Intro.pDinky.Height, Intro.pNegDinky.hDC, 0, 0, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(itr).X - 20, Ghosts(itr).Y - 35, Intro.pDinky.Width, Intro.pDinky.Height, Intro.pDinky.hDC, 0, 0, vbSrcPaint
      End If
      Picture1.Refresh
    Case 13
      itr = 6
      If Ghosts(itr).X < 85 Then
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, 239, 0, &HCC0020
        Ghosts(itr).X = Ghosts(itr).X + 2
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, 171 + (Ghosts(itr).Frame * 17), 51, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, PAC_SOURCE, Ghosts(itr).Src_X + (Ghosts(itr).Frame * 17), Ghosts(itr).Src_Y + (Ghosts(itr).Direction * 17), vbSrcPaint
        Ghosts(itr).Frame = Ghosts(itr).Frame + 1
        If Ghosts(itr).Frame > 2 Then Ghosts(itr).Frame = 0
      Else
        INTRO_COUNTER = INTRO_COUNTER + 1
        BitBlt Picture1.hDC, Ghosts(itr).X - 20, Ghosts(itr).Y - 35, Intro.pStinky.Width, Intro.pStinky.Height, Intro.pNegStinky.hDC, 0, 0, vbSrcAnd
        BitBlt Picture1.hDC, Ghosts(itr).X - 20, Ghosts(itr).Y - 35, Intro.pStinky.Width, Intro.pStinky.Height, Intro.pStinky.hDC, 0, 0, vbSrcPaint
      End If
      Picture1.Refresh
    Case 14
      lblStart.Visible = True
      Timer2.Interval = 0
      CAN_START = True
  End Select
End Sub

Private Sub MoveFood()
Dim pixel_N1 As Long
Dim pixel_N2 As Long
Dim pixel_S1 As Long
Dim pixel_S2 As Long
Dim pixel_E1 As Long
Dim pixel_E2 As Long
Dim pixel_W1 As Long
Dim pixel_W2 As Long
Dim rnd_dir As Integer

rnd_dir = Int(12 * Rnd())

pixel_N1 = GetPixel(Dots.Maze.hDC, Food.X, Food.Y - 1)
pixel_N2 = GetPixel(Dots.Maze.hDC, Food.X + 15, Food.Y - 1)
pixel_S1 = GetPixel(Dots.Maze.hDC, Food.X, Food.Y + 16)
pixel_S2 = GetPixel(Dots.Maze.hDC, Food.X + 15, Food.Y + 16)
pixel_E1 = GetPixel(Dots.Maze.hDC, Food.X + 16, Food.Y)
pixel_E2 = GetPixel(Dots.Maze.hDC, Food.X + 16, Food.Y + 15)
pixel_W1 = GetPixel(Dots.Maze.hDC, Food.X - 1, Food.Y)
pixel_W2 = GetPixel(Dots.Maze.hDC, Food.X - 1, Food.Y + 15)

If Food.X_Dest > Food.X_Start Then
  If pixel_E1 = BLACK_COLOR And pixel_E2 = BLACK_COLOR And _
     pixel_S1 = BLACK_COLOR And pixel_S2 = BLACK_COLOR And _
     pixel_N1 = BLACK_COLOR And pixel_N2 = BLACK_COLOR Then
    If Food.Y > Food.Y_Dest And Food.Direction <> South Then
      If (Food.X > 4 And Food.X < Picture1.Width - 23) And rnd_dir > 5 Then
        Food.Y = Food.Y - 1
        Food.Direction = North
      Else
        Food.X = Food.X + 1
        Food.Direction = East
      End If
    ElseIf Food.Y < Food.Y_Dest And Food.Direction <> North Then
      If (Food.X > 4 And Food.X < Picture1.Width - 23) And rnd_dir > 5 Then
        Food.Y = Food.Y + 1
        Food.Direction = South
      Else
        Food.X = Food.X + 1
        Food.Direction = East
      End If
    Else
      Food.X = Food.X + 1
      Food.Direction = East
    End If
  ElseIf pixel_E1 = BLACK_COLOR And pixel_E2 = BLACK_COLOR And _
         pixel_S1 = BLACK_COLOR And pixel_S2 = BLACK_COLOR Then
    If Food.Direction <> North And Food.Y < Food.Y_Dest And Food.X < Food.X_Dest Then
      If rnd_dir > 5 Then
        Food.Y = Food.Y + 1
        Food.Direction = South
      Else
        Food.X = Food.X + 1
        Food.Direction = East
      End If
    ElseIf Food.Direction = East Or Food.Y = Food.Y_Dest Then
      Food.X = Food.X + 1
      Food.Direction = East
    ElseIf Food.Y < Food.Y_Dest And Food.Direction <> North Then
      Food.Y = Food.Y + 1
      Food.Direction = South
    Else
      Food.X = Food.X + 1
      Food.Direction = East
    End If
  ElseIf pixel_E1 = BLACK_COLOR And pixel_E2 = BLACK_COLOR And _
         pixel_N1 = BLACK_COLOR And pixel_N2 = BLACK_COLOR Then
    If Food.Direction <> South And Food.Y > Food.Y_Dest And Food.X < Food.X_Dest Then
      If rnd_dir > 5 Then
        Food.Y = Food.Y - 1
        Food.Direction = North
      Else
        Food.X = Food.X + 1
        Food.Direction = East
      End If
    ElseIf Food.Direction = East Or Food.Y = Food.Y_Dest Then
      Food.X = Food.X + 1
      Food.Direction = East
    ElseIf (Food.X > 4 And Food.X < Picture1.Width - 23) And Food.Y > Food.Y_Dest And Food.Direction <> South Then
      Food.Y = Food.Y - 1
      Food.Direction = North
    Else
      Food.X = Food.X + 1
      Food.Direction = East
    End If
  ElseIf pixel_N1 = BLACK_COLOR And pixel_N2 = BLACK_COLOR And _
         pixel_S1 = BLACK_COLOR And pixel_S2 = BLACK_COLOR Then
    If Food.Y = Food.Y_Dest And Food.Direction = East Then
      If (rnd_dir > 5) Then
        Food.Y = Food.Y + 1
        Food.Direction = South
      Else
        Food.Y = Food.Y - 1
        Food.Direction = North
      End If
    ElseIf (Food.X > 4 And Food.X < Picture1.Width - 23) And Food.Y < Food.Y_Dest And (Food.Direction <> North) Then
      Food.Y = Food.Y + 1
      Food.Direction = South
    ElseIf (Food.X > 4 And Food.X < Picture1.Width - 23) And Food.Y > Food.Y_Dest And (Food.Direction <> South) Then
      Food.Y = Food.Y - 1
      Food.Direction = North
    ElseIf (Food.X > 4 And Food.X < Picture1.Width - 23) And (Food.Direction <> South) Then
      Food.Y = Food.Y - 1
      Food.Direction = North
    ElseIf (Food.X > 4 And Food.X < Picture1.Width - 23) Then
      Food.Y = Food.Y + 1
      Food.Direction = South
    End If
  ElseIf pixel_E1 = BLACK_COLOR And pixel_E2 = BLACK_COLOR Then
    Food.X = Food.X + 1
    Food.Direction = East
  ElseIf pixel_N1 = BLACK_COLOR And pixel_N2 = BLACK_COLOR Then
    Food.Y = Food.Y - 1
    Food.Direction = North
  ElseIf pixel_S1 = BLACK_COLOR And pixel_S2 = BLACK_COLOR Then
    Food.Y = Food.Y + 1
    Food.Direction = South
  ElseIf Food.X >= Food.X_Dest And Food.Y = Food.Y_Dest Then
    Food.Timer = 0
    Restore_Food
  Else
    Food.X = Food.X + 1
    Food.Direction = East
  End If
ElseIf Food.X_Dest < Food.X_Start Then
  If pixel_W1 = BLACK_COLOR And pixel_W2 = BLACK_COLOR And _
     pixel_S1 = BLACK_COLOR And pixel_S2 = BLACK_COLOR And _
     pixel_N1 = BLACK_COLOR And pixel_N2 = BLACK_COLOR Then
    If Food.Y < Food.Y_Dest And Food.Direction <> North Then
      If (Food.X > 4 And Food.X < Picture1.Width - 23) And rnd_dir > 5 Then
        Food.Y = Food.Y + 1
        Food.Direction = South
      Else
        Food.X = Food.X - 1
        Food.Direction = West
      End If
    ElseIf Food.Y > Food.Y_Dest And Food.Direction <> South Then
      If (Food.X > 4 And Food.X < Picture1.Width - 23) And rnd_dir > 5 Then
        Food.Y = Food.Y - 1
        Food.Direction = North
      Else
        Food.X = Food.X - 1
        Food.Direction = West
      End If
    Else
      Food.X = Food.X - 1
      Food.Direction = West
    End If
  ElseIf pixel_W1 = BLACK_COLOR And pixel_W2 = BLACK_COLOR And _
         pixel_S1 = BLACK_COLOR And pixel_S2 = BLACK_COLOR Then
    If Food.Direction <> North And Food.Y < Food.Y_Dest And Food.X > Food.X_Dest Then
      If rnd_dir > 5 Then
        Food.X = Food.X - 1
        Food.Direction = West
      Else
        Food.Y = Food.Y + 1
        Food.Direction = South
      End If
    ElseIf Food.Direction = West Or Food.Y = Food.Y_Dest Then
      Food.X = Food.X - 1
      Food.Direction = West
    ElseIf Food.Y < Food.Y_Dest And Food.Direction <> North Then
      Food.Y = Food.Y + 1
      Food.Direction = South
    Else
      Food.X = Food.X - 1
      Food.Direction = West
    End If
  ElseIf pixel_W1 = BLACK_COLOR And pixel_W2 = BLACK_COLOR And _
         pixel_N1 = BLACK_COLOR And pixel_N2 = BLACK_COLOR Then
    If Food.Direction <> South And Food.Y > Food.Y_Dest And Food.X > Food.X_Dest Then
      If rnd_dir > 5 Then
        Food.X = Food.X - 1
        Food.Direction = West
      Else
        Food.Y = Food.Y - 1
        Food.Direction = North
      End If
    ElseIf Food.Direction = West Or Food.Y = Food.Y_Dest Then
      Food.X = Food.X - 1
      Food.Direction = West
    ElseIf Food.Y > Food.Y_Dest And Food.Direction <> South Then
      Food.Y = Food.Y - 1
      Food.Direction = North
    Else
      Food.X = Food.X - 1
      Food.Direction = West
    End If
  ElseIf pixel_N1 = BLACK_COLOR And pixel_N2 = BLACK_COLOR And _
         pixel_S1 = BLACK_COLOR And pixel_S2 = BLACK_COLOR Then
    If Food.Y = Food.Y_Dest And Food.Direction = East Then
      If (rnd_dir > 6) Then
        Food.Y = Food.Y + 1
        Food.Direction = South
      Else
        Food.Y = Food.Y - 1
        Food.Direction = North
      End If
    ElseIf Food.Y < Food.Y_Dest And (Food.Direction <> North) Then
      Food.Y = Food.Y + 1
      Food.Direction = South
    ElseIf Food.Y > Food.Y_Dest And (Food.Direction <> South) Then
      Food.Y = Food.Y - 1
      Food.Direction = North
    ElseIf (Food.Direction <> South) Then
      Food.Y = Food.Y - 1
      Food.Direction = North
    Else
      Food.Y = Food.Y + 1
      Food.Direction = South
    End If
  ElseIf pixel_W1 = BLACK_COLOR And pixel_W2 = BLACK_COLOR Then
    Food.X = Food.X - 1
    Food.Direction = West
  ElseIf pixel_N1 = BLACK_COLOR And pixel_N2 = BLACK_COLOR Then
    Food.Y = Food.Y - 1
    Food.Direction = North
  ElseIf pixel_S1 = BLACK_COLOR And pixel_S2 = BLACK_COLOR Then
    Food.Y = Food.Y + 1
    Food.Direction = South
  ElseIf Food.X <= Food.X_Dest And Food.Y = Food.Y_Dest Then
    Food.Timer = 0
    Restore_Food
  Else
    Food.X = Food.X - 1
    Food.Direction = West
  End If
Else
  Food.Timer = 0
  Restore_Food
End If
End Sub

Private Sub Timer3_Timer()
Dim itr As Byte

BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, Dots.PosDots.hDC, PACMAN.X, PACMAN.Y, &HCC0020

For itr = 1 To 6
    BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, Dots.PosDots.hDC, Ghosts(itr).X, Ghosts(itr).Y, &HCC0020
Next itr

BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, PAC_SOURCE, ((PACMAN.Frame + 5) * 16) + (PACMAN.Frame + 5), PACMAN.Src_Y + (PACMAN.Direction * 16) + PACMAN.Direction, vbSrcAnd
BitBlt Picture1.hDC, PACMAN.X, PACMAN.Y, 16, 16, PAC_SOURCE, (PACMAN.Frame * 16) + (PACMAN.Frame + 1), PACMAN.Src_Y + (PACMAN.Direction * 16) + PACMAN.Direction, vbSrcPaint

If EXPLODE_FRAME < 6 Then
For itr = 1 To 6
 If Ghosts(itr).Eaten = False Then
  If (Ghosts(itr).Scared = True) Then
    BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, Ghosts(itr).Src_X + (5 * 17) + (EXPLODE_FRAME * 17), Ghosts(itr).Src_Y + (34), vbSrcAnd
    BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, Ghosts(itr).Src_X + (5 * 17) + (EXPLODE_FRAME * 17), Ghosts(itr).Src_Y + (17 + 34), vbSrcPaint
  Else
    BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, Ghosts(itr).Src_X + (5 * 17) + (EXPLODE_FRAME * 17), Ghosts(itr).Src_Y + (34), vbSrcAnd
    BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, Ghosts(itr).Src_X + (5 * 17) + (EXPLODE_FRAME * 17), Ghosts(itr).Src_Y + (17), vbSrcPaint
  End If
 Else
   BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 52 + 17, Ghosts(itr).Src_Y + (Ghosts(itr).Direction * 17), vbSrcAnd
   BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 52, Ghosts(itr).Src_Y + (Ghosts(itr).Direction * 17), vbSrcPaint
 End If
Next itr
End If

EXPLODE_FRAME = EXPLODE_FRAME + 1
If EXPLODE_FRAME = 6 Then
  For itr = 1 To 6
    If (Ghosts(itr).Scared = True) And Ghosts(itr).Eaten = False Then
      secret = secret + Ghosts(itr).Letter
      text1.Caption = secret
      If DEATH_TOLL < 6 Then DEATH_TOLL = DEATH_TOLL + 1
      Score = Score + (BONUS_MULTIPLIER * ((2 ^ (DEATH_TOLL)) * 100))
      BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 1 + ((DEATH_TOLL - 1) * 17), 493, vbSrcAnd
      BitBlt Picture1.hDC, Ghosts(itr).X, Ghosts(itr).Y, 16, 16, CHAR_SOURCE, 1 + ((DEATH_TOLL - 1) * 17), 476, vbSrcPaint
      EATEN_TIMER = 45
    End If
    Ghosts(itr).Scared = False
    Ghosts(itr).Eaten = True
    RUNNING_SCARED = 0
  Next itr
  If secret = "PACMAN" Then
    secret = ""
    PACMAN.lives = PACMAN.lives + 1
    Display_Lives
  Else
    secret = ""
    text1.Caption = secret
  End If
  Timer3.Interval = 0
  Timer1.Interval = TIME_TIME
End If
Picture1.Refresh
End Sub
