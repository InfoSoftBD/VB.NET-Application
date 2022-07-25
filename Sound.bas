Attribute VB_Name = "Module2"
Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias " GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Global gDBpath As String
Global gDBName As String

#If Win32 Then
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long
#Else
Declare Function sndPlaySound Lib "MMSYSTEM.DLL" (ByVal lpszSoundName As Any, ByVal wFlags As Integer) As Integer
#End If
Declare Function SetWindowPos Lib "User32" (ByVal h&, ByVal hb&, ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal f&) As Long
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_MEMORY = &H4

Global SoundBuffer As String
Sub BeginPlaySound(ByVal ResourceId As Integer)
Dim Ret As Variant
#If Win32 Then
SoundBuffer = StrConv(LoadResData(ResourceId, "WAVE"), vbUnicode)
#Else
SounfBuffer = LoadResData(ResourceId, "WAVE")
#End If
Ret = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
DoEvents
End Sub
Sub EndPlaySound()
Dim Ret As Variant
Ret = sndPlaySound(0&, 0&)
End Sub


