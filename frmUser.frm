VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUser 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Management"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5475
   FillColor       =   &H00004000&
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3705
      Left            =   180
      ScaleHeight     =   3645
      ScaleWidth      =   5040
      TabIndex        =   3
      Top             =   135
      Width           =   5100
      Begin VB.TextBox txtDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1260
         TabIndex        =   14
         Top             =   180
         Width           =   1680
      End
      Begin VB.ComboBox cmbType 
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
         ItemData        =   "frmUser.frx":0442
         Left            =   1260
         List            =   "frmUser.frx":044C
         TabIndex        =   12
         Top             =   1395
         Width           =   1725
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1260
         TabIndex        =   10
         Top             =   1935
         Width           =   3570
      End
      Begin VB.TextBox txtConfirm 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   1260
         PasswordChar    =   "@"
         TabIndex        =   8
         Top             =   3105
         Width           =   2220
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1260
         PasswordChar    =   "@"
         TabIndex        =   5
         Top             =   2520
         Width           =   2235
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1260
         TabIndex        =   4
         Top             =   765
         Width           =   1680
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click to Add Photo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   3255
         MouseIcon       =   "frmUser.frx":0466
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   1650
         Width           =   1545
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Left            =   3240
         Picture         =   "frmUser.frx":08A8
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1590
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   720
         TabIndex        =   15
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Type"
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
         Left            =   210
         TabIndex        =   13
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2070
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm"
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
         Left            =   405
         TabIndex        =   9
         Top             =   3195
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
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
         Left            =   480
         TabIndex        =   7
         Top             =   855
         Width           =   675
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   6
         Top             =   2610
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   180
      ScaleHeight     =   720
      ScaleWidth      =   5040
      TabIndex        =   0
      Top             =   4005
      Width           =   5100
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   2565
         TabIndex        =   18
         Top             =   135
         Width           =   1080
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   495
         Left            =   1350
         TabIndex        =   17
         Top             =   135
         Width           =   1080
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   495
         Left            =   3825
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1080
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Create User"
         Height          =   495
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   135
         Width           =   1080
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Photo As String
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String

Private Sub cmbType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtName.SetFocus
End If
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from User_Info where User_Id like '" & txtID.Text & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        If MsgBox("Do you really want to delete?", vbCritical + vbYesNo, "Delete!") = vbYes Then
        str = "delete from User_Info where User_Id like '" & txtID.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
        MsgBox "User successfully deleted!", vbInformation, "Delete!"
        txtDate.Text = Date
        txtID.Text = ""
        txtName.Text = ""
        cmbType.Text = ""
        Image1.Picture = LoadPicture("")
        txtPassword.Text = ""
        txtConfirm.Text = ""
        txtID.SetFocus
        Else
        txtID.SetFocus
        Exit Sub
    End If
    End If
End Sub
Private Sub cmdOK_Click()
If txtPassword.Text = txtConfirm.Text Then

On Error Resume Next
    Set rsN = New ADODB.Recordset
        str = "select * from User_Info where User_Id like '" & txtID.Text & "'"
        rsN.Open str, conn
    
    If rsN.EOF Then
        rsN.Close
        
        rsN.Open "User_Info", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!User_Id = txtID.Text
        rsN!Name = txtName.Text
        rsN!Password = txtPassword.Text
        rsN!User_Type = cmbType.Text
        rsN!Photo = Photo
        rsN.Update
        rsN.Close
        MsgBox "User creat successfull!", vbInformation, "Accepted!"
        txtDate.Text = Date
        txtID.Text = ""
        txtName.Text = ""
        cmbType.Text = ""
        Image1.Picture = LoadPicture("")
        txtPassword.Text = ""
        txtConfirm.Text = ""
        txtID.SetFocus

    Else
        MsgBox "User ID already exist. Please try another User ID", 64, "User Error"
        rsN.Close
        txtDate.Text = Date
        txtID.Text = ""
        txtName.Text = ""
        cmbType.Text = ""
        Image1.Picture = LoadPicture("")
        txtPassword.Text = ""
        txtConfirm.Text = ""
        txtID.SetFocus
    End If

Else
    MsgBox "Password Miss Match.Try again", 17, "Password Error"
    txtPassword.Text = ""
    txtConfirm.Text = ""
    txtPassword.SetFocus
End If
Exit Sub
   Resume Next
End Sub

Private Sub cmdUpdate_Click()
Set rsN = New ADODB.Recordset
        str = "select * from User_Info where User_Id like '" & txtID.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
        If Not rsN.EOF Then
        rsN!Date = txtDate.Text
        rsN!User_Id = txtID.Text
        rsN!Name = txtName.Text
        rsN!Password = txtPassword.Text
        rsN!User_Type = cmbType.Text
        rsN!Photo = Photo
        rsN.Update
        rsN.Close
        MsgBox "User update successfull!", vbInformation, "Accepted!"
        txtDate.Text = Date
        txtID.Text = ""
        txtName.Text = ""
        cmbType.Text = ""
        Image1.Picture = LoadPicture("")
        txtPassword.Text = ""
        txtConfirm.Text = ""
        txtID.SetFocus
        End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
cmdOK.Enabled = False
cmdUpdate.Enabled = False
cmdDelete.Enabled = False
txtDate.Text = Date
End Sub

Private Sub Label34_Click()
CommonDialog1.ShowOpen
CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Bitmaps|*.BMP|GIF Images|*.GIF|JPEG Images|*.JPG|All Images|*.BMP;*.GIF;*.JPG"
Image1.Picture = LoadPicture(App.Path + "\Photo\" & CommonDialog1.FileTitle)
Photo = CommonDialog1.FileTitle
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOK.SetFocus
End If
End Sub

Private Sub txtConfirm_LostFocus()
If txtPassword.Text = txtConfirm.Text Then
Exit Sub
Else
MsgBox "Password Miss Match.Try again", 17, "Password Error"
txtPassword.Text = ""
txtConfirm.Text = ""
txtPassword.SetFocus
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbType.SetFocus
End If
End Sub

Private Sub txtId_LostFocus()
If txtID.Text = "" Then
Exit Sub
End If

Set rs = New ADODB.Recordset
        str = "select * from User_Info where User_Id like '" & txtID.Text & "'"
        rs.Open str, conn
    On Error Resume Next
If Not rs.EOF Then
    If MsgBox("User exist! Do you want add new User?", vbInformation + vbYesNo, "Add New") = vbYes Then
        rs.Close
        txtID.Text = ""
        txtName.Text = ""
        txtPassword.Text = ""
        txtConfirm.Text = ""
        txtID.SetFocus
        Image1.Picture = LoadPicture("")
        cmdOK.Enabled = True
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
        Else
        txtID.Text = rs!User_Id
        txtName.Text = rs!Name
        cmbType.Text = rs!User_Type
        txtPassword.Text = rs!Password
        txtConfirm.Text = rs!Password
        Image1.Picture = LoadPicture(App.Path + "\Photo\" & rs!Photo)
        Photo = rs!Photo
        rs.Close
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = True
        cmdOK.Enabled = False
    End If
Else '

rs.Close
    If MsgBox("Invalid User Id! Do you want to ad new user?", vbInformation + vbYesNo, "Add New") = vbYes Then
    txtName.Text = ""
    cmbType.Text = ""
    Image1.Picture = LoadPicture("")
    txtPassword.Text = ""
    txtConfirm.Text = ""
    cmdOK.Enabled = True
Exit Sub
Else
txtID.Text = ""
txtName.Text = ""
cmbType.Text = ""
Image1.Picture = LoadPicture("")
txtPassword.Text = ""
txtConfirm.Text = ""
cmdOK.Enabled = False
txtID.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPassword.SetFocus
End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtConfirm.SetFocus
End If
End Sub
