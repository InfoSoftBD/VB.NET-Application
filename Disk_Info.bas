Attribute VB_Name = "Module6"
Private Declare Function GetWindowsVersion Lib "kernel32" Alias "GetVersion" () As Long
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)

Public Function GetVersion() As Long
    On Error GoTo ErrTrap
    GetVersion& = GetWindowsVersion&
Exit Function
ErrTrap:
    GetVersion& = 0&
End Function
Public Function GetHardDiskSerial(Optional sDrive As String) As Long
    On Error GoTo ErrTrap
    Dim lNumber As Long, sBuffer As String * 255
    If sDrive$ = "" Then sDrive$ = "C"
    Call GetVolumeInformation(sDrive$ & ":\", sBuffer$, 255, lNumber&, 0&, 0&, sBuffer$, 255)
    GetHardDiskSerial& = lNumber&
Exit Function
ErrTrap:
    Call MsgBox("Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error")
    HardDiskSerial& = 0&
End Function
