VERSION 5.00
Begin VB.Form game 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6105
   ClientLeft      =   6960
   ClientTop       =   4425
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "game.frx":0442
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   564
   Begin VB.PictureBox exp1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1200
      Picture         =   "game.frx":D07C
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   6
      Top             =   3600
      Width           =   900
   End
   Begin VB.PictureBox exp2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   2040
      Picture         =   "game.frx":F57E
      ScaleHeight     =   840
      ScaleWidth      =   840
      TabIndex        =   5
      Top             =   3600
      Width           =   900
   End
   Begin VB.PictureBox menuitem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      Picture         =   "game.frx":F788
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   4
      Top             =   4800
      Width           =   1110
   End
   Begin VB.PictureBox menu 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   0
      Picture         =   "game.frx":F9DE
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   3
      Top             =   4080
      Width           =   1110
   End
   Begin VB.Timer anim 
      Interval        =   100
      Left            =   2640
      Top             =   2400
   End
   Begin VB.Timer fpstimer 
      Interval        =   1000
      Left            =   2040
      Top             =   3000
   End
   Begin VB.Timer GameTimer 
      Interval        =   1
      Left            =   2040
      Top             =   2400
   End
   Begin VB.PictureBox button 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      Picture         =   "game.frx":106E4
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   2
      Top             =   3600
      Width           =   1110
   End
   Begin VB.PictureBox charmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   0
      Picture         =   "game.frx":10B0E
      ScaleHeight     =   480
      ScaleWidth      =   1920
      TabIndex        =   1
      Top             =   3000
      Width           =   1980
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   0
      Picture         =   "game.frx":13B50
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      Top             =   2400
      Width           =   1980
   End
End
Attribute VB_Name = "game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Draw()
'char1.
'wenn pause an ist wird nicht mehr weiterge"malt"
If Pause = False Then
'bildschirm löschen
Me.Cls

'einfache animation
If Player1SPRITE = 1 Then
BitBlt game.hdc, Player1X, Player1Y, 32, 32, charmask.hdc, 0, 0, MERGEPAINT
BitBlt game.hdc, Player1X, Player1Y, 32, 32, char.hdc, 0, 0, SRCAND
Else
BitBlt game.hdc, Player1X, Player1Y, 32, 32, charmask.hdc, 32, 0, MERGEPAINT
BitBlt game.hdc, Player1X, Player1Y, 32, 32, char.hdc, 32, 0, SRCAND
End If
'char2
'einfache animation
If Player2SPRITE = 1 Then
BitBlt game.hdc, Player2X, Player2Y, 32, 32, charmask.hdc, 64, 0, MERGEPAINT
BitBlt game.hdc, Player2X, Player2Y, 32, 32, char.hdc, 64, 0, SRCAND
Else
BitBlt game.hdc, Player2X, Player2Y, 32, 32, charmask.hdc, 96, 0, MERGEPAINT
BitBlt game.hdc, Player2X, Player2Y, 32, 32, char.hdc, 96, 0, SRCAND
End If
BitBlt Me.hdc, 0, 0, button.Width, button.Height, button.hdc, 0, 0, vbSrcPaint
BitBlt Me.hdc, 250, 0, button.Width, button.Height, button.hdc, 0, 0, vbSrcPaint
'HP und MP der Player anzeigen
textout 9, 1, "HP: " & Player1HP
textout 9, 12, "MP: " & Player1MP
textout 259, 1, "HP: " & Player2HP
textout 259, 12, "MP: " & Player2MP
If GetAsyncKeyState(vbKeyS) < 0 Then
ShowMenu = True
End If
If ShowMenu = True Then
'load menu
BitBlt Me.hdc, 0, 0, menu.Width, menu.Height, menu.hdc, 0, 0, vbSrcPaint
BitBlt Me.hdc, MenuItemX, MenuItemY, menuitem.Width, menuitem.Height, menuitem.hdc, 0, 0, vbSrcPaint
textout 9, 30, "Fight"
textout 9, 50, "Magic"
textout 9, 70, "Run"
End If
If Player1MAGIC = True Then
Pause = True
anim.Enabled = False
GameTimer.Enabled = False
fpstimer.Enabled = False
PlayWav App.Path & "\blowup.wav"
BitBlt game.hdc, 80, 45, 56, 56, exp2.hdc, 0, 0, MERGEPAINT
BitBlt game.hdc, 80, 45, 56, 56, exp1.hdc, 0, 0, SRCAND
BitBlt game.hdc, 185, 45, 56, 56, exp2.hdc, 0, 0, MERGEPAINT
BitBlt game.hdc, 185, 45, 56, 56, exp1.hdc, 0, 0, SRCAND
BitBlt Me.hdc, 125, 60, button.Width, button.Height, button.hdc, 0, 0, vbSrcPaint
textout 125, 65, "not done yet"
Wait 2
Pause = False
anim.Enabled = True
GameTimer.Enabled = True
fpstimer.Enabled = True
Player1MAGIC = False
End If
If Player2HIT = True Then
ShowMenu = False
If Player2BACK = True Then
If Player2X = 140 Then
Randomize
Dice = Int((20 * Rnd) + 1)
Player1HP = Player1HP - Dice
If Player1HP < 0 Then
Player1HP = 0
Pause = True
anim.Enabled = False
GameTimer.Enabled = False
fpstimer.Enabled = False
PlayWav App.Path & "\blowup.wav"
BitBlt game.hdc, 80, 45, 56, 56, exp2.hdc, 0, 0, MERGEPAINT
BitBlt game.hdc, 80, 45, 56, 56, exp1.hdc, 0, 0, SRCAND
BitBlt game.hdc, 185, 45, 56, 56, exp2.hdc, 0, 0, MERGEPAINT
BitBlt game.hdc, 185, 45, 56, 56, exp1.hdc, 0, 0, SRCAND
BitBlt Me.hdc, 125, 60, button.Width, button.Height, button.hdc, 0, 0, vbSrcPaint
textout 125, 65, "Looser"
Wait 2
Pause = False
anim.Enabled = True
GameTimer.Enabled = True
fpstimer.Enabled = True
Player1MAGIC = False
End
End If
End If
End If
If Player2X = 56 Then
Pause = True
anim.Enabled = False
GameTimer.Enabled = False
fpstimer.Enabled = False
PlayWav App.Path & "\blowup.wav"
BitBlt game.hdc, Player1X - 20, Player1Y - 20, 56, 56, exp2.hdc, 0, 0, MERGEPAINT
BitBlt game.hdc, Player1X - 20, Player1Y - 20, 56, 56, exp1.hdc, 0, 0, SRCAND
Wait 1
Pause = False
anim.Enabled = True
GameTimer.Enabled = True
fpstimer.Enabled = True
End If
If Player2MOVE Then
Player2X = Player2X - 2
If Player2X > 60 Then
textout Player2X - 5, Player2Y - 20, "ATTACK"
End If
If Player2X < Player1X + 2 Then
Player2BACK = True
Player2MOVE = False
End If
End If
If Player2BACK = True Then
If Player2X > 60 Then
textout Player2X - 5, Player2Y - 20, "WELL DONE"
End If
Player2X = Player2X + 2
If Player2X = 230 Then
Player2BACK = False
Player2HIT = False
Player2MOVE = False
ShowMenu = True
End If
End If
End If
If Player1MAGIC Then
End If
If Player1HIT = True Then
'the attack
ShowMenu = False
If Player1BACK = True Then
If Player1X < 50 Then
Player1BACK = False
ShowMenu = False
Player1HIT = False
Player2HIT = True
Player2MOVE = True
End If
If Player1X = 140 Then
Randomize
Dice = Int((20 * Rnd) + 1)
Player2HP = Player2HP - Dice
If Player2HP < 0 Then
Player2HP = 0
Pause = True
anim.Enabled = False
GameTimer.Enabled = False
fpstimer.Enabled = False
PlayWav App.Path & "\blowup.wav"
BitBlt game.hdc, 80, 45, 56, 56, exp2.hdc, 0, 0, MERGEPAINT
BitBlt game.hdc, 80, 45, 56, 56, exp1.hdc, 0, 0, SRCAND
BitBlt game.hdc, 185, 45, 56, 56, exp2.hdc, 0, 0, MERGEPAINT
BitBlt game.hdc, 185, 45, 56, 56, exp1.hdc, 0, 0, SRCAND
BitBlt Me.hdc, 125, 60, button.Width, button.Height, button.hdc, 0, 0, vbSrcPaint
textout 125, 65, "Winner"
Wait 2
Pause = False
anim.Enabled = True
GameTimer.Enabled = True
fpstimer.Enabled = True
Player1MAGIC = False
End
End If
End If
If Player1X < 200 Then
textout Player1X - 5, Player1Y - 20, "MUHAHA"
End If
Player1X = Player1X - 2
End If
If Player1MOVE = True Then
If Player1X < 200 Then
textout Player1X - 5, Player1Y - 20, "KILL YA!!"
End If
Player1X = Player1X + 2
End If

'THE MEETING -- >
If Player1X = Player2X - 10 Then
Pause = True
anim.Enabled = False
GameTimer.Enabled = False
fpstimer.Enabled = False
PlayWav App.Path & "\blowup.wav"
BitBlt game.hdc, Player2X - 20, Player2Y - 20, 56, 56, exp2.hdc, 0, 0, MERGEPAINT
BitBlt game.hdc, Player2X - 20, Player2Y - 20, 56, 56, exp1.hdc, 0, 0, SRCAND
Wait 1
Pause = False
anim.Enabled = True
GameTimer.Enabled = True
fpstimer.Enabled = True
End If
If Player1X = Player2X - 4 Then
Player1MOVE = False
Player1BACK = True
End If
End If
'frames pro sekunde aufschreiben
FPS = FPS + 1
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
End
End If
If ShowMenu = True Then
If KeyCode = vbKeyUp Then
If MenuItemY = 30 Then MenuItemY = 90
MenuItemY = MenuItemY - 20
End If
If KeyCode = vbKeyDown Then
If MenuItemY = 70 Then MenuItemY = 10
MenuItemY = MenuItemY + 20
End If
If KeyCode = vbKeyReturn Then
If MenuItemY = 30 Then
Player1HIT = True
Player1MOVE = True
End If
If MenuItemY = 50 Then
Player1MAGIC = True
End If
If MenuItemY = 70 Then
End
End If
End If
End If
End Sub

Private Sub gametimer_Timer()
'hauptsub starten
Draw
End Sub

Private Sub Form_Load()
Sprite = 0
'X und Y Position zuordnen. x = links - rechts , y = oben - unten
'variablen einen wert zuweisen
Player1X = 50
Player1Y = 120
Player2X = 230
Player2Y = 120
Player1HP = 100
Player2HP = 100
Player1MP = 100
Player2MP = 100
ParticleOFF = True
Player1SPRITE = 1
Player2SPRITE = 1
Player2BACK = False
Player2MOVE = True
ShowMenu = False
MenuItemX = 0
MenuItemY = 30
'hauptsub starten
Draw

End Sub


Private Sub fpstimer_Timer()
'die fps rate in der leiste anzeigen
Me.Caption = "FightEngine 1.1 , FPS: " & FPS
'fps wieder reseten
FPS = 0
End Sub

Private Sub anim_Timer()
'einfache animation
If Player1SPRITE = 1 Then Player1SPRITE = 2 Else Player1SPRITE = 1
If Player2SPRITE = 1 Then Player2SPRITE = 2 Else Player2SPRITE = 1
End Sub

Private Sub textout(x As Integer, y As Integer, text As String)
CurrentX = x
CurrentY = y
Print text
End Sub

Private Sub Wait(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

