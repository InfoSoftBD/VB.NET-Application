Attribute VB_Name = "Module5"
Public Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function TransparentForm Lib "SkinForm.dll" (ByVal sPathFile As String) As Long
Public wdt As Integer
Public ht As Integer
