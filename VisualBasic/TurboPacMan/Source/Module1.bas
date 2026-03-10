Attribute VB_Name = "Module1"
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X _
                                     As Long, ByVal Y As Long, ByVal nWidth _
                                     As Long, ByVal nHeight As Long, _
                                     ByVal hSrcDC As Long, ByVal xSrc _
                                     As Long, ByVal ySrc As Long, ByVal _
                                     dwRop As Long) As Long
                                     
                                     
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
                                                 (ByVal lpszSoundName As String, _
                                                  ByVal wFlags As Long) As Long

Public Const SND_ASYNC = &H1
Public Const SND_SYNC = &H0        ' Play synchronously (default)
Public Const SND_NODEFAULT = &H2   ' Don't use default sound
Public Const SND_MEMORY = &H4      ' lpszSoundName points to a
                                   ' memory file.
Public Const SND_LOOP = &H8        ' Loop the sound until next
                                   ' sndPlaySound.
Public Const SND_NOSTOP = &H10     ' Don't stop any currently
                                   ' playing sound.

               
Declare Function GetPixel Lib "gdi32" (ByVal handle As Long, _
                                       ByVal nXpos As Long, _
                                       ByVal nYpos As Long) As Long

Public Type PowerPellet
  X As Integer
  Y As Integer
  Points As Integer
  Frame As Integer
  Eaten As Boolean
  Src_X As Integer
  Src_Y As Integer
End Type

Public Type GameLevel
  MazeMap As Long
  DoneMap As Long
  PositiveDots As Long
  NegativeDots As Long
  DotMax As Integer
End Type

Public Type Ghost
  X As Integer
  Y As Integer
  Src_X As Integer
  Src_Y As Integer
  Frame As Integer
  Direction As Byte
  JustEaten As Boolean
  Eaten As Boolean
  Scared As Boolean
  Free As Boolean
  Speed As Byte
  Letter As String
End Type

Public Type Bonus
  X As Integer
  Y As Integer
  Eaten As Boolean
  Value As Long
  Timer As Long
  X_Dest As Integer
  Y_Dest As Integer
  X_Start As Integer
  Y_Start As Integer
  Direction As Byte
  Special As Boolean
  SpecialType As Byte
  Source_X As Long
  Source_Y As Long
End Type

Public Type Pac
  X As Integer
  Y As Integer
  X_vector As Integer
  Y_vector As Integer
  Src_X As Integer
  Src_Y As Integer
  Frame As Integer
  frame_val As Integer
  lives As Integer
  Direction As Byte
  TURBO As Boolean
  Dying As Boolean
  MS As Boolean
  Poisoned As Boolean
  PoisonCounter As Integer
End Type

