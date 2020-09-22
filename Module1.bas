Attribute VB_Name = "Module1"

Option Explicit
Public Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYN = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H18
Public Const SND_NOSTOP = &H10

