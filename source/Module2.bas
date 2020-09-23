Attribute VB_Name = "Module2"
Option Explicit


Public Declare Function mciGetErrorString Lib "winmm" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PlaySound Lib "winmm" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


Public Enum SNDPLAYSOUND
    SND_ASYNC = &H1
    SND_LOOP = &H8
    SND_NODEFAULT = &H2
    SND_NOSTOP = &H10
    SND_NOWAIT = &H2000
    SND_PURGE = &H40
    SND_SYNC = &H0
End Enum


Public Function PlayShortWav(ByVal sndName As String) As Long
    TerminateAllSound
    PlayShortWav = PlaySound(sndName, 0, SNDPLAYSOUND.SND_NODEFAULT Or SNDPLAYSOUND.SND_SYNC)
End Function

Public Function PlayWav(ByVal StrName As String) As Long
TerminateAllSound
Dim RetStr As String, CallBack As Long, ShortName
    RetStr = Space$(128)
    ShortName = GetShortFileName(StrName)
    PlayWav = mciSendString("open waveaudio!" & ShortName & " alias wav", RetStr, 128, CallBack)
    PlayWav = mciSendString("play wav", RetStr, 128, CallBack)
End Function

Public Function PlayMidi(ByVal StrName As String) As Long
TerminateAllSound
Dim RetStr As String, CallBack As Long, ShortName
    RetStr = Space$(128)
    ShortName = GetShortFileName(StrName)
    PlayMidi = mciSendString("open sequencer!" & ShortName & " alias midi", RetStr, 128, CallBack)
    PlayMidi = mciSendString("play midi", RetStr, 128, 1)
End Function

Public Function StopWav() As Long
Dim RetStr As String, CallBack As Long
    RetStr = Space$(128)
    StopWav = mciSendString("stop wav", RetStr, 128, CallBack)
    StopWav = mciSendString("close wav", RetStr, 128, CallBack)
End Function

Public Function StopMidi() As Long
Dim RetStr As String, CallBack As Long
    RetStr = Space$(128)
    StopMidi = mciSendString("stop midi", RetStr, 128, CallBack)
    StopMidi = mciSendString("close midi", RetStr, 128, CallBack)
End Function

Public Function TerminateAllSound() As Long
    StopWav
    StopMidi
End Function

Public Function GetShortFileName(ByVal FileName As String) As String
Dim rc As Long
Dim ShortPath As String
    Const PATH_LEN& = 164
ShortPath = String$(PATH_LEN + 1, 0)
rc = GetShortPathName(FileName, ShortPath, PATH_LEN)
GetShortFileName = Left$(ShortPath, rc)
End Function


Public Function SetProgress(ByVal ProgressBar As PictureBox, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long, ByVal Alias As String) As String
Dim RetStr As String, CallBack As Long
Dim qTemp As Single, rTemp As Long, sTemp As Long
Dim Num$
RetStr = Space$(128)
SetProgress = mciSendString("status " & Alias & " length", RetStr, 128, CallBack)
ProgressBar.Cls
ProgressBar.ForeColor = RGB(Green, Blue, Red)
If Val(RetStr) Then
    rTemp = Val(RetStr)
    qTemp = Val(RetStr)
    SetProgress = mciSendString("status " & Alias & " position", RetStr, 128, CallBack)
    sTemp = Val(RetStr)
    qTemp = Val(RetStr) / qTemp
    Num$ = Format$(qTemp * 100, "###") + "%"
    ProgressBar.Line (0, 0)-(ProgressBar.Width * qTemp, ProgressBar.Height), RGB(Red, Green, Blue), BF
    ProgressBar.CurrentX = (ProgressBar.ScaleWidth - ProgressBar.TextWidth(Num$)) / 2
    ProgressBar.CurrentY = (ProgressBar.ScaleHeight - ProgressBar.TextHeight(Num$)) / 2
    ProgressBar.Print Num$
Else
    Num$ = Format$(100, "###") + "%"
    ProgressBar.Line (0, 0)-(ProgressBar.Width, ProgressBar.Height), ProgressBar.BackColor, BF
    ProgressBar.CurrentX = (ProgressBar.ScaleWidth - ProgressBar.TextWidth(Num$)) / 2
    ProgressBar.CurrentY = (ProgressBar.ScaleHeight - ProgressBar.TextHeight(Num$)) / 2
    ProgressBar.Print Num$
End If
If sTemp = rTemp Then
    ProgressBar.Cls
    SetProgress = "Done"
    TerminateAllSound
End If
End Function




