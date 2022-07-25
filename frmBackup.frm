VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackup 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Backup"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   150
      ScaleHeight     =   315
      ScaleWidth      =   5085
      TabIndex        =   13
      Top             =   4845
      Width           =   5145
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   150
      ScaleHeight     =   705
      ScaleWidth      =   5145
      TabIndex        =   6
      Top             =   180
      Width           =   5175
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BACKUP PROCESS"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   975
         TabIndex        =   7
         Top             =   120
         Width           =   3510
      End
      Begin VB.Image Image3 
         Height          =   735
         Left            =   0
         Picture         =   "frmBackup.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5205
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   150
      ScaleHeight     =   675
      ScaleWidth      =   5085
      TabIndex        =   3
      Top             =   5325
      Width           =   5145
      Begin VB.CommandButton Command1 
         Caption         =   "Start Backup"
         Height          =   510
         Left            =   360
         TabIndex        =   5
         Top             =   90
         Width           =   1545
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   510
         Left            =   3150
         TabIndex        =   4
         Top             =   90
         Width           =   1545
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   3660
      Left            =   150
      ScaleHeight     =   3600
      ScaleWidth      =   5085
      TabIndex        =   0
      Top             =   1080
      Width           =   5145
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   2400
         TabIndex        =   12
         Top             =   1320
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtDesti 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   1
         Text            =   "BMS_Backup"
         Top             =   480
         Width           =   2520
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup FIle List:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   15
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Directory Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drive Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Folder:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   2
         Top             =   120
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prog As Integer
Dim prgMax As Integer
Dim strCurrentFile As String
Dim strPattern As String
Dim strSourceFile As String
Private Sub CopyOrMoveFiles(pstrSourcePath As String, pstrTargetPath As String, pstrExtension As String, pstrCopyOrMoveCode As String)

If Right$(pstrSourcePath, 1) <> "\" Then
pstrSourcePath = pstrSourcePath & "\"
End If

If Right$(pstrTargetPath, 1) <> "\" Then
pstrTargetPath = pstrTargetPath & "\"
End If

strPattern = pstrSourcePath & "*." & pstrExtension
strCurrentFile = Dir$(strPattern)

Do Until strCurrentFile = ""

    strSourceFile = pstrSourcePath & strCurrentFile
    strTargetFile = pstrTargetPath & strCurrentFile

    FileCopy strSourceFile, strTargetFile
    List1.AddItem strCurrentFile + " Copied."
    List1.Refresh
    
    prog = prog + 1
    ProgressBar1.Min = 0
    ProgressBar1.Max = prgMax
    ProgressBar1.Value = prog
    strCurrentFile = Dir$()
Loop

If pstrCopyOrMoveCode = "M" Then
    Kill strPattern
End If

End Sub
Private Sub CountFiles(pstrSourcePath As String, pstrTargetPath As String, pstrExtension As String, pstrCopyOrMoveCode As String)
If Right$(pstrSourcePath, 1) <> "\" Then
    pstrSourcePath = pstrSourcePath & "\"
End If
strPattern = pstrSourcePath & "*." & pstrExtension
strCurrentFile = Dir$(strPattern)

Do Until strCurrentFile = ""

strSourceFile = pstrSourcePath & strCurrentFile

List1.AddItem strCurrentFile + " Copied."
prgMax = prgMax + 1
strCurrentFile = Dir$()
Loop
End Sub
Private Sub Command1_Click()
On Error GoTo cmdTryIt_Click_Error
    conn.Close
    List1.Clear
    prog = 0
    prgMax = 0
    ProgressBar1.Value = 0

    Dim Location As String
    Dim Photo As String
    Dim Thumb As String
    Dim Vedio As String
    Dim Sound As String
    
        Location = Dir1.Path + "\" + txtDesti.Text
        Photo = Location + "\Photo"
        Thumb = Location + "\Thumb"
        Vedio = Location + "\Vedio"
        'Sound = Location + "\Sound"
        
    If Dir$(Location, vbDirectory) = "" Then
       MkDir (Location)
       CountFiles App.Path, Location, "*.*", "C"
       CopyOrMoveFiles App.Path, Location, "*.*", "C"
       List1.AddItem strSourceFile
       
       If Dir$(Photo, vbDirectory) = "" Then
        MkDir (Photo)
        CountFiles App.Path + "\Photo", Photo, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Photo", Photo, "*.*", "C"
        List1.AddItem strSourceFile
       Else
        CountFiles App.Path + "\Photo", Photo, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Photo", Photo, "*.*", "C"
        List1.AddItem strSourceFile
       End If
       
       If Dir$(Thumb, vbDirectory) = "" Then
        MkDir (Thumb)
        CountFiles App.Path + "\Thumb", Thumb, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Thumb", Thumb, "*.*", "C"
        List1.AddItem strSourceFile
       Else
        CountFiles App.Path + "\Thumb", Thumb, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Thumb", Thumb, "*.*", "C"
        List1.AddItem strSourceFile
       End If
       
       If Dir$(Vedio, vbDirectory) = "" Then
        MkDir (Vedio)
        CountFiles App.Path + "\Vedio", Vedio, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Vedio", Vedio, "*.*", "C"
        List1.AddItem strSourceFile
       Else
        CountFiles App.Path + "\Vedio", Vedio, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Vedio", Vedio, "*.*", "C"
        List1.AddItem strSourceFile
       End If
       
    conn.Open cnStr
    MsgBox "Backup Successfully Completed! " & prgMax & " No.s File copied.", vbInformation, "Successful!"
        
Else
        CountFiles App.Path, Location, "*.*", "C"
        CopyOrMoveFiles App.Path, Location, "*.*", "C"
    
        If Dir$(Photo, vbDirectory) = "" Then
        MkDir (Photo)
        CountFiles App.Path + "\Photo", Photo, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Photo", Photo, "*.*", "C"
        List1.AddItem strSourceFile
       Else
        CountFiles App.Path + "\Photo", Photo, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Photo", Photo, "*.*", "C"
        List1.AddItem strSourceFile
       End If
       
       If Dir$(Thumb, vbDirectory) = "" Then
        MkDir (Thumb)
        CountFiles App.Path + "\Thumb", Thumb, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Thumb", Thumb, "*.*", "C"
        List1.AddItem strSourceFile
       Else
        CountFiles App.Path + "\Thumb", Thumb, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Thumb", Thumb, "*.*", "C"
        List1.AddItem strSourceFile
       End If
       
       If Dir$(Vedio, vbDirectory) = "" Then
        MkDir (Vedio)
        CountFiles App.Path + "\Vedio", Vedio, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Vedio", Vedio, "*.*", "C"
        List1.AddItem strSourceFile
       Else
        CountFiles App.Path + "\Vedio", Vedio, "*.*", "C"
        CopyOrMoveFiles App.Path + "\Vedio", Vedio, "*.*", "C"
        List1.AddItem strSourceFile
       End If
    conn.Open cnStr
    MsgBox "Backup Successfully Completed! " & prgMax & " No.s File copied.", vbInformation, "Successful!"

Exit Sub

cmdTryIt_Click_Error:

MsgBox "The following error has occurred:" & vbNewLine & "Error # " & Err.Number & " - " & Err.Description, vbCritical, "File System Commands Demo - Error"

End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
txtDesti.Text = "BMS_Backup-" & Today
End Sub

