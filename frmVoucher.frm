VERSION 5.00
Begin VB.Form frmGLAccount 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GL Account Opening"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   Icon            =   "frmVoucher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   165
      ScaleHeight     =   705
      ScaleWidth      =   6285
      TabIndex        =   1
      Top             =   5505
      Width           =   6315
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   435
         Left            =   4785
         TabIndex        =   5
         Top             =   150
         Width           =   1350
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   435
         Left            =   135
         TabIndex        =   4
         Top             =   150
         Width           =   1350
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   435
         Left            =   3225
         TabIndex        =   3
         Top             =   150
         Width           =   1350
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Update"
         Height          =   435
         Left            =   1680
         TabIndex        =   2
         Top             =   150
         Width           =   1350
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4335
      Index           =   0
      Left            =   165
      ScaleHeight     =   4305
      ScaleWidth      =   6240
      TabIndex        =   0
      Top             =   1020
      Width           =   6270
      Begin VB.ComboBox txtName 
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
         Left            =   1920
         TabIndex        =   21
         Text            =   "Combo1"
         Top             =   1050
         Width           =   3930
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GL Account Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   180
         TabIndex        =   12
         Top             =   135
         Width           =   5865
         Begin VB.ComboBox txtAccount 
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
            Left            =   1740
            TabIndex        =   22
            Text            =   "Combo1"
            Top             =   420
            Width           =   1545
         End
         Begin VB.ComboBox cmbH_Group 
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
            Left            =   1710
            TabIndex        =   19
            Text            =   "Combo1"
            Top             =   1980
            Width           =   3930
         End
         Begin VB.TextBox txtDate 
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
            Height          =   375
            Left            =   4095
            TabIndex        =   18
            Text            =   "Text4"
            Top             =   360
            Width           =   1545
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
            Left            =   1725
            TabIndex        =   17
            Text            =   "Combo1"
            Top             =   1455
            Width           =   1545
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Group"
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
            TabIndex        =   20
            Top             =   2025
            Width           =   1410
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Type"
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
            Left            =   360
            TabIndex        =   16
            Top             =   1515
            Width           =   1290
         End
         Begin VB.Label Label1 
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
            Left            =   3510
            TabIndex        =   15
            Top             =   420
            Width           =   435
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
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
            Left            =   270
            TabIndex        =   14
            Top             =   945
            Width           =   1380
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GL Account No."
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
            Left            =   180
            TabIndex        =   13
            Top             =   420
            Width           =   1470
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contra Account Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   180
         TabIndex        =   7
         Top             =   2700
         Width           =   5865
         Begin VB.ComboBox cmbContra_Ac 
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
            Left            =   1710
            TabIndex        =   10
            Text            =   "Combo1"
            Top             =   405
            Width           =   1545
         End
         Begin VB.ComboBox cmbContra_Name 
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
            Left            =   1710
            TabIndex        =   9
            Text            =   "Combo1"
            Top             =   990
            Width           =   3930
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
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
            TabIndex        =   11
            Top             =   1035
            Width           =   1380
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account No."
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
            TabIndex        =   8
            Top             =   450
            Width           =   1140
         End
      End
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "GL ACCOUNT OPENING"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   945
      TabIndex        =   6
      Top             =   270
      Width           =   4725
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   210
      Picture         =   "frmVoucher.frx":0442
      Stretch         =   -1  'True
      Top             =   180
      Width           =   6255
   End
End
Attribute VB_Name = "frmGLAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Today As Date
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub clearTextboxes()
        txtAccount.Text = ""
        txtDate.Text = Date
        txtName.Text = ""
        cmbType.Text = ""
        cmbH_Group.Text = ""
        cmbContra_Ac.Text = ""
        cmbContra_Name.Text = ""
End Sub
Private Sub CmbAc()
        txtAccount.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT AC_No FROM GL_Master"
        rs.Open str, conn
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        txtAccount.AddItem rs!AC_No
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub CmbContraAc()
        cmbContra_Ac.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT AC_No FROM GL_Master"
        rs.Open str, conn
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbContra_Ac.AddItem rs!AC_No
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub CmbName()
        txtName.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Head_Name FROM GL_Master"
        rs.Open str, conn
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        txtName.AddItem rs!Head_Name
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub CmbContraName()
        cmbContra_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Head_Name FROM GL_Master"
        rs.Open str, conn
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbContra_Name.AddItem rs!Head_Name
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub cmbGroup()
        cmbH_Group.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Group FROM GL_Master"
        rs.Open str, conn
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbH_Group.AddItem rs!Group
        rs.MoveNext
        Loop
        rs.Close
End Sub

Private Sub cmbContra_Ac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbContra_Name.SetFocus
End If
End Sub

Private Sub cmbContra_Ac_LostFocus()
Dim Contra_Ac As String
Contra_Ac = cmbContra_Ac.Text
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & Contra_Ac & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
        On Error Resume Next
        cmbContra_Name.Text = rs!Head_Name
    End If
End Sub

Private Sub cmbContra_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Command2.Enabled = True Then
Command2.SetFocus
Else
Exit Sub
End If
End If
End Sub

Private Sub cmbContra_Name_LostFocus()
Dim Contra_Name As String
Contra_Name = cmbContra_Name.Text
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where Head_Name like '" & Contra_Name & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
        On Error Resume Next
        cmbContra_Ac.Text = rs!AC_No
    End If
End Sub

Private Sub cmbGroup_Change()

End Sub

Private Sub cmbH_Group_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbContra_Ac.SetFocus
End If
End Sub

Private Sub cmbType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbH_Group.SetFocus
End If
End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim ID As String
    ID = txtAccount.Text
    
     Set rsN = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & ID & "'"
        rsN.Open str, conn
    
    If rsN.EOF Then
        rsN.Close
        
        rsN.Open "GL_Master", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!AC_No = txtAccount.Text
        rsN!Date = txtDate.Text
        rsN!Head_Name = txtName.Text
        rsN!Head_Type = cmbType.Text
        rsN!Group = cmbH_Group.Text
        rsN!Contra_Ac = cmbContra_Ac.Text
        rsN!Contra_Name = cmbContra_Name.Text
        rsN!Balance = 0
        
        If cmbType.Text = "ASSET" Then
        rsN!H_Code = 1
        Else
        If cmbType.Text = "LIABILITY" Then
        rsN!H_Code = 2
        Else
        If cmbType.Text = "INCOME" Then
        rsN!H_Code = 3
        Else
        If cmbType.Text = "EXPENSE" Then
        rsN!H_Code = 4
        End If
        End If
        End If
        End If
        rsN.Update
        rsN.Close
    Else
        MsgBox "There exist a record with this Account no.: " & rsN!ID
        rsN.Close
    End If
    
    Call clearTextboxes
    txtDate.Text = Today
    txtAccount.SetFocus
    Command2.Enabled = False
    Command3.Enabled = False
    Command5.Enabled = False
    Exit Sub
End Sub

Private Sub Command4_Click()
 
Dim sort_Asset As Integer
Dim sort_Liab As Integer
      
    sort_Asset = 0
    sort_Liab = 0


On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Trial_Balance where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Trial_Balance where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * from TRL1"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Trial_Balance", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
            
                rsN!sl = 1
                rsN!Asset_Sl = sort_Asset + 1
                rsN!Asset_AC = rs!AC_No
                rsN!Asset_Head = rs!Head_Name
                rsN!Asset_Balance = rs!Balance
                sort_Asset = sort_Asset + 1
            
            
            'If rs!Head_Type = "Liability" Then
               ' rsN!sl = 1
              '  rsN!Liab_Sl = sort_Liab + 1
               ' rsN!Liab_AC = rs!AC_No
               ' rsN!Liab_Head = rs!Head_Name
               ' rsN!Liab_Balance = rs!Balance
               ' sort_Liab = sort_Liab + 1
            'End If
            
        rs.MoveNext
        rsN.Update
    
        Loop
        rs.Close
        rsN.Close
        
    'rptStatement.rsStatement.ConnectionString = cnStr
   ' rptStatement.rsStatement.Source = str

'Timer1.Enabled = True
End Sub

Private Sub Command5_Click()

Set rsN = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & txtAccount.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsN.EOF Then
        rsN!AC_No = txtAccount.Text
        rsN!Date = txtDate.Text
        rsN!Head_Name = txtName.Text
        rsN!Head_Type = cmbType.Text
        rsN!Group = cmbH_Group.Text
        rsN!Contra_Ac = cmbContra_Ac.Text
        rsN!Contra_Name = cmbContra_Name.Text
        
        If cmbType.Text = "ASSET" Then
        rsN!H_Code = 1
        Else
        If cmbType.Text = "LIABILITY" Then
        rsN!H_Code = 2
        Else
        If cmbType.Text = "INCOME" Then
        rsN!H_Code = 3
        Else
        If cmbType.Text = "EXPENSE" Then
        rsN!H_Code = 4
        End If
        End If
        End If
        End If
        rsN.Update
        rsN.Close
    Else
        MsgBox "There is no record found with this Account no.: " & rsN!ID
        rsN.Close
    End If
    
    Call clearTextboxes
    txtAccount.SetFocus
    Command2.Enabled = False
    Command3.Enabled = False
    Command5.Enabled = False
End Sub

Private Sub Form_Load()
 On Error Resume Next
       Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
           On Error Resume Next
           Today = rs!Today
           rs.Close
           txtDate.Text = Today
        End If
Call clearTextboxes
Call CmbContraAc
Call CmbAc
Call CmbContraName
Call CmbName
Call cmbGroup
cmbType.AddItem "ASSET"
cmbType.AddItem "LIABILITY"
cmbType.AddItem "INCOME"
cmbType.AddItem "EXPENSE"

Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtName.SetFocus
End If
End Sub

Private Sub txtAccount_LostFocus()
Dim ID As String
    ID = txtAccount.Text
Dim mid As String
    mid = 0
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
    Call clearTextboxes
    On Error Resume Next
        txtAccount.Text = rs!AC_No
        txtDate.Text = rs!Date
        cmbH_Group.Text = rs!Group
        txtName.Text = rs!Head_Name
        cmbType.Text = rs!Head_Type
        cmbContra_Ac.Text = rs!Contra_Ac
        cmbContra_Name.Text = rs!Contra_Name
        rs.Close
        
        Command5.Enabled = True
        Command3.Enabled = True
    Else
    If MsgBox("Do you want add new GL Account?", vbInformation + vbYesNo, "Add New") = vbYes Then
        
        Set rsU = New ADODB.Recordset
        str = "select * from GL_Master order by AC_No"
        rsU.Open str, conn
        rsU.MoveFirst
        
    Do While Not rsU.EOF
        mid = Val(rsU!AC_No)
        rsU.MoveNext
    Loop
        rsU.Close
        mid = mid + 1
        
        
        Call clearTextboxes
        txtDate.Text = Today
        txtAccount.Text = mid
        Command2.Enabled = True
    Else
        Call clearTextboxes
        txtName.SetFocus
    End If
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbType.SetFocus
End If
End Sub
