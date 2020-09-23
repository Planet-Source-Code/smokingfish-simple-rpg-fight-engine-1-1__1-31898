Attribute VB_Name = "mod1"
'Die API's
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As RasterOpConstants) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const MERGEPAINT = &HBB0226
Public Const SRCCOPY = &HCC0020
'Variablen
Public Player1X As Long
Public Player2X As Long
Public Player1Y As Long
Public Sprite As Integer
Public Player2Y As Long
Public Dice As Integer
Public Player1HIT As Boolean
Public Player2HIT As Boolean
Public Player1MAGIC As Boolean
Public Player2MAGIC As Boolean
Public Pause As Boolean
Public FlamePlayer As Integer
Public ParticleOFF As Boolean
Public Player1SPRITE As Integer
Public Player2SPRITE As Integer
Public Player1HP As Integer
Public Player2HP As Integer
Public Player1MP As Integer
Public Player2MP As Integer
Public GameEnd As Boolean
Public FPS As Long
Public MenuItemX As Integer
Public MenuItemY As Integer
Public ShowMenu As Boolean
Public Player1MOVE As Boolean
Public Player2MOVE As Boolean
Public Player1BACK As Boolean
Public Player2BACK As Boolean
