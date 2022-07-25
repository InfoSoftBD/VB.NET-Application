VERSION 5.00
Begin VB.Form frmAuthorization 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Authorization"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   Icon            =   "frmAuthorization.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   180
      ScaleHeight     =   765
      ScaleWidth      =   6945
      TabIndex        =   7
      Top             =   2745
      Width           =   6975
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   525
         Left            =   585
         TabIndex        =   9
         Top             =   135
         Width           =   1830
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   525
         Left            =   4680
         TabIndex        =   8
         Top             =   135
         Width           =   1830
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      Left            =   180
      ScaleHeight     =   1530
      ScaleWidth      =   6885
      TabIndex        =   2
      Top             =   990
      Width           =   6945
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   2205
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   855
         Width           =   4515
      End
      Begin VB.TextBox txtAutho_Code 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2250
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   180
         Width           =   4515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1140
         TabIndex        =   6
         Top             =   945
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authorization Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   5
         Top             =   270
         Width           =   1830
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   180
      ScaleHeight     =   660
      ScaleWidth      =   6945
      TabIndex        =   0
      Top             =   135
      Width           =   6975
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AUTHORIZATION"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   1680
         TabIndex        =   1
         Top             =   45
         Width           =   4155
      End
      Begin VB.Image Image3 
         Height          =   645
         Left            =   0
         Picture         =   "frmAuthorization.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6990
      End
   End
End
Attribute VB_Name = "frmAuthorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String

Private Sub cmdClose_Click()
Unload Me
frmPassword.Show 1
End Sub

Private Sub Command2_Click()
 Dim strCode As Long
    Dim cnt As Integer
    Dim cntSt As Integer
    Dim cntLn As Integer
    Dim Autho As String
    Autho = ""
    cnt = 0
    cntSt = 0
    cntLn = 1
    Do While Not cnt = 5
    txtAutho_Code.SelStart = cntSt
    txtAutho_Code.SelLength = cntLn
    strCode = Val(Asc(txtAutho_Code.SelText)) * 136
    cnt = cnt + 1
    cntSt = cntSt + 1
    cntLn = cntLn + 1
    If Autho = "" Then
    Autho = strCode
    Else
    Autho = Autho & "-" & strCode
    End If
    Loop

If txtPassword.Text = Autho Then
    Set rsN = New ADODB.Recordset
        str = "select * from Others"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        rsN.MoveFirst
    
    If rsN!Autho_Code = "0" Then
        rsN!Autho_Code = txtAutho_Code.Text
        rsN.Update
        rsN.Close
        Else
            If rsN!Autho_Code1 = "0" Then
            rsN!Autho_Code1 = txtAutho_Code.Text
            rsN.Update
            rsN.Close
            Else
                If rsN!Autho_Code2 = "0" Then
                rsN!Autho_Code2 = txtAutho_Code.Text
                rsN.Update
                rsN.Close
                End If
            End If
        End If
        
        MsgBox "Authorization Successfull!", vbInformation, "Authorization!"
        txtAutho_Code.Text = ""
        txtPassword.Text = ""
    Else
    MsgBox "Invalid Password", vbCritical, "Authorization!"
    txtAutho_Code.Text = ""
    txtPassword.Text = ""
    Exit Sub
End If
End Sub

Private Sub Form_Load()
 Dim Code As String
 Dim CPU As String
 Dim BIOS As String
 Dim HARD_DISK As String
    CPU = GetWmiDeviceSingleValue("Win32_Processor", "ProcessorID")
    BIOS = GetWmiDeviceSingleValue("Win32_BIOS", "SerialNumber")
    HARD_DISK = GetHardDiskSerial("C")
    Code = CPU + "-" + BIOS + "-" + HARD_DISK
    txtAutho_Code.Text = ""
    txtPassword.Text = ""
    txtAutho_Code.Text = Code
    'MsgBox GetVersion
Command2.Enabled = False
End Sub

Private Sub txtAutho_Code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 On Error Resume Next
   Dim strCode As Long
   Dim Autho As String
    Dim cnt As Integer
    Dim cntSt As Integer
    Dim cntLn As Integer
    Autho = ""
    cnt = 0
    cntSt = 0
    cntLn = 1
    Do While Not cnt = 5
    txtAutho_Code.SelStart = cntSt
    txtAutho_Code.SelLength = cntLn
    strCode = Val(Asc(txtAutho_Code.SelText)) * 136
    cnt = cnt + 1
    cntSt = cntSt + 1
    cntLn = cntLn + 1
    
    If Autho = "" Then
    Autho = strCode
    Else
    Autho = Autho & "-" & strCode
    End If
    Loop
    'Text1.Text = Autho
    
txtPassword.SetFocus
End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    On Error Resume Next
    Dim strCode As Long
    Dim Autho As String
    Dim cnt As Integer
    Dim cntSt As Integer
    Dim cntLn As Integer
    Autho = ""
    cnt = 0
    cntSt = 0
    cntLn = 1
    Do While Not cnt = 5
    txtAutho_Code.SelStart = cntSt
    txtAutho_Code.SelLength = cntLn
    strCode = Val(Asc(txtAutho_Code.SelText)) * 136
    cnt = cnt + 1
    cntSt = cntSt + 1
    cntLn = cntLn + 1
    If Autho = "" Then
    Autho = strCode
    Else
    Autho = Autho & "-" & strCode
    End If
    Loop
    If txtPassword.Text = Autho Then
    Command2.Enabled = True
    Command2.SetFocus
    Else
    MsgBox "Invalid Authorization Code!", vbCritical, "Authorization!"
    Exit Sub
    End If
End If
End Sub

