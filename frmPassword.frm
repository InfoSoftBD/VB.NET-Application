VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Password"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   10125
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPassword.frx":0442
   ScaleHeight     =   6690
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2565
      Picture         =   "frmPassword.frx":11BD0
      ScaleHeight     =   225
      ScaleWidth      =   5580
      TabIndex        =   4
      Top             =   6435
      Width           =   5580
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   435
         Picture         =   "frmPassword.frx":13B75
         ScaleHeight     =   135
         ScaleWidth      =   4575
         TabIndex        =   5
         Top             =   45
         Width           =   4575
         Begin VB.Image Image3 
            Height          =   210
            Left            =   -4680
            Picture         =   "frmPassword.frx":15641
            Top             =   -30
            Width           =   4710
         End
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5895
      Width           =   990
   End
   Begin VB.TextBox txtPassword 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   900
      PasswordChar    =   "@"
      TabIndex        =   2
      Top             =   5400
      Width           =   2835
   End
   Begin VB.TextBox txtID 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   900
      TabIndex        =   1
      Top             =   4635
      Width           =   2835
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   300
      Left            =   9630
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6390
      Width           =   405
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   2760
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1725
      Left            =   990
      Stretch         =   -1  'True
      Top             =   2340
      Width           =   1830
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String

Private Sub Command1_Click()
Dim Code As String
 Dim CPU As String
 Dim BIOS As String
 Dim HARD_DISK As String
 Dim Chkdate As Date
 Dim Day As Date
    Day = Date
    CPU = GetWmiDeviceSingleValue("Win32_Processor", "ProcessorID")
    BIOS = GetWmiDeviceSingleValue("Win32_BIOS", "SerialNumber")
    HARD_DISK = GetHardDiskSerial("C")
    Code = CPU + "-" + BIOS + "-" + HARD_DISK
    
    Set rs = New ADODB.Recordset
            str = "select Autho_Code, Autho_Code1, Autho_Code2 from Others where Autho_Code like '" & Code & "' Or Autho_Code1 like '" & Code & "' Or Autho_Code2 like '" & Code & "'"
        rs.Open str, conn
    'On Error Resume Next
    If Not rs.EOF Then
        rs.Close
    Else
        rs.Close
        MsgBox "Authorization Required!", vbCritical, "Authorization Error!"
        Unload Me
        frmAuthorization.Show 1
    End If
    
    Set rs = New ADODB.Recordset
        str = "select * from User_Info where User_Id like '" & txtId.Text & "' and Password like '" & txtPassword.Text & "'"
        rs.Open str, conn
        
    'On Error Resume Next
    If Not rs.EOF Then
        txtId.Text = rs!User_Id
        txtPassword.Text = rs!Password
        Image1.Picture = LoadPicture(App.Path + "\Photo\" & rs!Photo)
        
        Set rsN = New ADODB.Recordset
            str = "select * from Others"
            rsN.Open str, conn, adOpenDynamic, adLockOptimistic
            rsN.MoveFirst
        
            rsN!Pass_Date = "31-12-2030"
            rsN!User_Id = rs!User_Id
            rsN!Password = rs!Password
            rsN!Name = rs!Name
            rsN!User_Type = rs!User_Type
            rsN.Update
            
            Today = CDate(rsN!Today)
            Chkdate = CDate(rsN!Pass_Date)
            Power = rsN!Power
            User_Id = rsN!User_Id
            User_Name = rsN!Name
            U_Type = rsN!User_Type
            Org_Name = rsN!Org_Name
            Branch_Code = rsN!Branch_Code
            Branch_Name = rsN!Branch_Name
            Branch_Address = rsN!Branch_Address
            
            If Today > Chkdate Or Day > Chkdate Then
              MsgBox "Contact Period Over!", vbCritical, "Error!"
            Exit Sub
            End If
            rsN.Close
           
        rs.Close
    
        Timer1.Enabled = True
        
    Else
    If MsgBox("Invalid user name or password.", 17, "Access Denied!") <> 1 Then Exit Sub
        rs.Close
        txtId.SetFocus
        txtId.Text = ""
        txtPassword.Text = ""
        Image1.Picture = LoadPicture("")
    End If
End Sub

Private Sub Command2_Click()
End
End Sub
Private Sub Form_Load()
'Call SetWindowRgn(Me.hwnd, TransparentForm(App.Path + "\Photo\login Round.bmp"), True)
End Sub



Private Sub Picture3_Click()

End Sub

Private Sub Timer1_Timer()
' ProgressBar1.Visible = True
'    ProgressBar1.Max = 100
'    ProgressBar1.Value = ProgressBar1.Value + 1
'If ProgressBar1.Value = ProgressBar1.Max Then
'    Timer1.Enabled = False
'    ProgressBar1.Visible = False
'    Unload Me
'   frmMain.Show
'End If

Image3.Left = Image3.Left + 20
If Image3.Left = 0 Then
    Timer1.Enabled = False
    Unload Me
    frmMain.Show
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPassword.SetFocus
End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Dim Code As String
 Dim CPU As String
 Dim BIOS As String
 Dim HARD_DISK As String
 Dim Chkdate As Date
 Dim Day As Date
    Day = Date
    CPU = GetWmiDeviceSingleValue("Win32_Processor", "ProcessorID")
    BIOS = GetWmiDeviceSingleValue("Win32_BIOS", "SerialNumber")
    HARD_DISK = GetHardDiskSerial("C")
    Code = CPU + "-" + BIOS + "-" + HARD_DISK
    
    Set rs = New ADODB.Recordset
        str = "select Autho_Code, Autho_Code1, Autho_Code2 from Others where Autho_Code like '" & Code & "' Or Autho_Code1 like '" & Code & "' Or Autho_Code2 like '" & Code & "'"
        rs.Open str, conn
    'On Error Resume Next
    If Not rs.EOF Then
        rs.Close
    Else
    
        rs.Close
        MsgBox "Authorization Required!", vbCritical, "Authorization Error!"
        Unload Me
        frmAuthorization.Show 1
    End If
    
    Set rs = New ADODB.Recordset
        str = "select * from User_Info where User_Id like '" & txtId.Text & "' and Password like '" & txtPassword.Text & "'"
        rs.Open str, conn
        
    'On Error Resume Next
    If Not rs.EOF Then
        txtId.Text = rs!User_Id
        txtPassword.Text = rs!Password
        Image1.Picture = LoadPicture(App.Path + "\Photo\" & rs!Photo)
        
        Set rsN = New ADODB.Recordset
            str = "select * from Others"
            rsN.Open str, conn, adOpenDynamic, adLockOptimistic
            rsN.MoveFirst
        
            rsN!Pass_Date = "31-12-2030"
            rsN!User_Id = rs!User_Id
            rsN!Password = rs!Password
            rsN!Name = rs!Name
            rsN!User_Type = rs!User_Type
            rsN.Update
            
            Today = CDate(rsN!Today)
            Chkdate = CDate(rsN!Pass_Date)
            Power = rsN!Power
            User_Id = rsN!User_Id
            User_Name = rsN!Name
            U_Type = rsN!User_Type
            Org_Name = rsN!Org_Name
            Branch_Code = rsN!Branch_Code
            Branch_Name = rsN!Branch_Name
            Branch_Address = rsN!Branch_Address
            
            If Today > Chkdate Or Day > Chkdate Then
              MsgBox "Contact Period Over!", vbCritical, "Error!"
            Exit Sub
            End If
            rsN.Close
           
        rs.Close
    
        Timer1.Enabled = True
        
    Else
    If MsgBox("Invalid user name or password.", 17, "Access Denied!") <> 1 Then Exit Sub
        rs.Close
        On Error Resume Next
        txtId.SetFocus
        txtId.Text = ""
        txtPassword.Text = ""
        Image1.Picture = LoadPicture("")
    End If
End If
End Sub




