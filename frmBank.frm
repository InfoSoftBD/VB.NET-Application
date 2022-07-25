VERSION 5.00
Begin VB.Form frmBank_Tran 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Transaction Posting"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   Icon            =   "frmBank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3435
      Left            =   180
      ScaleHeight     =   3405
      ScaleWidth      =   5970
      TabIndex        =   7
      Top             =   990
      Width           =   6000
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transaction Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3120
         Left            =   135
         TabIndex        =   8
         Top             =   90
         Width           =   5685
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
            Left            =   1320
            TabIndex        =   26
            Text            =   "Combo1"
            Top             =   360
            Width           =   2220
         End
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4095
            TabIndex        =   24
            Top             =   2025
            Width           =   1410
         End
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4095
            TabIndex        =   22
            Top             =   360
            Width           =   1410
         End
         Begin VB.TextBox txtDescription 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1305
            TabIndex        =   19
            Top             =   2565
            Width           =   4155
         End
         Begin VB.ComboBox cmbType 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmBank.frx":0442
            Left            =   1305
            List            =   "frmBank.frx":044C
            TabIndex        =   18
            Text            =   "Combo2"
            Top             =   1485
            Width           =   1320
         End
         Begin VB.TextBox txtChq_Amnt 
            Alignment       =   1  'Right Justify
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
            Left            =   1305
            TabIndex        =   12
            Text            =   "Text7"
            Top             =   2025
            Width           =   1320
         End
         Begin VB.ComboBox cmbBranch 
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
            Left            =   4095
            TabIndex        =   11
            Text            =   "Combo2"
            Top             =   945
            Width           =   1410
         End
         Begin VB.ComboBox cmbBank 
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
            Left            =   1305
            TabIndex        =   10
            Text            =   "Combo1"
            Top             =   945
            Width           =   2040
         End
         Begin VB.TextBox txtChq_No 
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
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1485
            Width           =   1410
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
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
            Left            =   3330
            TabIndex        =   25
            Top             =   2070
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   3585
            TabIndex        =   23
            Top             =   435
            Width           =   405
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tran Type"
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
            Left            =   135
            TabIndex        =   21
            Top             =   1575
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Description"
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
            Left            =   135
            TabIndex        =   20
            Top             =   2640
            Width           =   1035
         End
         Begin VB.Label lblBank 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
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
            Left            =   135
            TabIndex        =   17
            Top             =   990
            Width           =   1020
         End
         Begin VB.Label lblBranch 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch"
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
            Left            =   3375
            TabIndex        =   16
            Top             =   990
            Width           =   615
         End
         Begin VB.Label lblChq_Amnt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount "
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
            Left            =   135
            TabIndex        =   15
            Top             =   2070
            Width           =   690
         End
         Begin VB.Label lblChq_No 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque No."
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
            Left            =   3060
            TabIndex        =   14
            Top             =   1530
            Width           =   1020
         End
         Begin VB.Label lblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account No."
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
            Left            =   135
            TabIndex        =   13
            Top             =   405
            Width           =   1080
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   180
      ScaleHeight     =   630
      ScaleWidth      =   5955
      TabIndex        =   2
      Top             =   4635
      Width           =   5985
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Save"
         Height          =   435
         Left            =   255
         TabIndex        =   6
         Top             =   105
         Width           =   1020
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   435
         Left            =   1725
         TabIndex        =   5
         Top             =   105
         Width           =   1020
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   435
         Left            =   4635
         TabIndex        =   4
         Top             =   90
         Width           =   1020
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   435
         Left            =   3195
         TabIndex        =   3
         Top             =   90
         Width           =   1020
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   180
      ScaleHeight     =   615
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   180
      Width           =   5985
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BANK TRANSACTION"
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
         Height          =   465
         Left            =   705
         TabIndex        =   1
         Top             =   45
         Width           =   4635
      End
      Begin VB.Image Image3 
         Height          =   645
         Left            =   0
         Picture         =   "frmBank.frx":0463
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6045
      End
   End
End
Attribute VB_Name = "frmBank_Tran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub Bank_Cr()
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100101 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs!Date = txtDate.Text
        rs.Update
            
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C No. " & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
        
        
    Set rs = New ADODB.Recordset
        str = "select * from Bank_Master where AC_No like '" & txtAccount.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic

        rs!Withdraw = rs!Withdraw + Val(txtChq_Amnt.Text)
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs!Date = txtDate.Text
        rs.Update
        
        Set rsN = New ADODB.Recordset
        rsN.Open "Bank_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = txtAccount.Text
        rsN!Bank_Name = cmbBank.Text
        rsN!Branch_Name = cmbBranch.Text
        rsN!Chq_No = "CN" & txtChq_No.Text
        rsN!Description = txtDescription.Text
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Bank_Cr = rsU!Bank_Cr + Val(txtChq_Amnt.Text)
        rsU!Bank_Close = rsU!Bank_Close - Val(txtChq_Amnt.Text)
        rsU.Update
        rsU.Close
        
End Sub
Private Sub Bank_Dr()
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100101 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtChq_Amnt.Text)
        rs!Date = txtDate.Text
        rs.Update
            
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C No. " & txtAccount.Text
        rsN!Dr = Val(txtChq_Amnt.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
        
    Set rs = New ADODB.Recordset
        str = "select * from Bank_Master where AC_No like '" & txtAccount.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rs.EOF Then
        rs!Deposit = rs!Deposit + Val(txtChq_Amnt.Text)
        rs!Balance = rs!Balance + Val(txtChq_Amnt.Text)
        rs!Date = txtDate.Text
        rs.Update
        
        Set rsN = New ADODB.Recordset
        rsN.Open "Bank_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = txtAccount.Text
        rsN!Bank_Name = cmbBank.Text
        rsN!Branch_Name = cmbBranch.Text
        rsN!Chq_No = "VN" & txtChq_No.Text
        rsN!Description = txtDescription.Text
        rsN!Dr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Bank_Dr = rsU!Bank_Dr + Val(txtChq_Amnt.Text)
        rsU!Bank_Close = rsU!Bank_Close + Val(txtChq_Amnt.Text)
        rsU.Update
        rsU.Close
        
    Else
        
    Set rsN = New ADODB.Recordset
           rsN.Open "Bank_Master", conn, adOpenDynamic, adLockOptimistic, -1
           rsN.AddNew
       
           rsN!Date = txtDate.Text
           rsN!AC_No = txtAccount.Text
           rsN!Bank_Name = cmbBank.Text
           rsN!Branch_Name = cmbBranch.Text
           rsN!Open_Bal = 0
           rsN!Deposit = Val(txtChq_Amnt.Text)
           rsN!Withdraw = 0
           rsN!Balance = Val(txtChq_Amnt.Text)
           rsN.Update
           rsN.Close
        
            Set rsN = New ADODB.Recordset
            rsN.Open "Bank_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!AC_No = txtAccount.Text
            rsN!Bank_Name = cmbBank.Text
            rsN!Branch_Name = cmbBranch.Text
            rsN!Chq_No = "VN" & txtChq_No.Text
            rsN!Description = txtDescription.Text
            rsN!Dr = Val(txtChq_Amnt.Text)
            rsN!Balance = Val(txtChq_Amnt.Text)
            rsN.Update
            rsN.Close
            rs.Close
        
        Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Bank_Dr = rsU!Bank_Dr + Val(txtChq_Amnt.Text)
        rsU!Bank_Close = rsU!Bank_Close + Val(txtChq_Amnt.Text)
        rsU.Update
        rsU.Close
End If
    
End Sub
Private Sub Cash_Dr()
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtChq_Amnt.Text)
        rs!Date = txtDate.Text
        rs.Update
            
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C No. " & txtAccount.Text
        rsN!Dr = Val(txtChq_Amnt.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        
        Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!MR_No = "CN" & txtChq_No.Text
        rsN!Name = rs!Head_Name
        rsN!Description = txtAccount.Text
        rsN!Dr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Cash_Dr = rsU!Cash_Dr + Val(txtChq_Amnt.Text)
        rsU!Cash_Close = rsU!Cash_Close + Val(txtChq_Amnt.Text)
        rsU.Update
        rsU.Close
End Sub
Private Sub Cash_Cr()
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs!Date = txtDate.Text
        rs.Update
            
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C No. " & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        
        Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!MR_No = "VN " & txtChq_No.Text
        rsN!Name = rs!Head_Name
        rsN!Description = txtAccount.Text
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Cash_Cr = rsU!Cash_Cr + Val(txtChq_Amnt.Text)
        rsU!Cash_Close = rsU!Cash_Close - Val(txtChq_Amnt.Text)
        rsU.Update
        rsU.Close
End Sub

Private Sub Branch_Name()
On Error Resume Next
        cmbBranch.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Branch_Name FROM Bank_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbBranch.AddItem rsN!Branch_Name
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub clearTextboxes()
        txtAccount.Text = ""
        cmbBank.Text = ""
        cmbBranch.Text = ""
        cmbType.Text = ""
        txtChq_No.Text = ""
        txtBalance.Text = ""
        txtChq_Amnt.Text = ""
        txtDescription.Text = ""
End Sub
Private Sub Account_Name()
        txtAccount.Clear
         Set rsN = New ADODB.Recordset
            str = "SELECT DISTINCT AC_NO FROM Bank_Master"
            rsN.Open str, conn

    If Not rsN.EOF Then
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        txtAccount.AddItem rsN!AC_No
        rsN.MoveNext
        Loop
        rsN.Close
        Else
        rsN.Close
        End If
End Sub
Private Sub Account_No()
       On Error Resume Next
        txtAccount.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT AC_No FROM Bank_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        txtAccount.AddItem rsN!AC_No
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub Bank_Name()
        On Error Resume Next
        cmbBank.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Bank_Name FROM Bank_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbBank.AddItem rsN!Bank_Name
        rsN.MoveNext
        Loop
        rsN.Close
End Sub

Private Sub cmbBank_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbBranch.SetFocus
End If
End Sub

Private Sub cmbBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbType.SetFocus
End If
End Sub
Private Sub cmbBranch_LostFocus()
'Set rsN = New ADODB.Recordset
'        str = "SELECT * FROM Bank_Master where Bank_Name Like '" & cmbBank.Text & "' AND Branch_Name Like '" & cmbBranch.Text & "' order by Bank_Name"
'        rsN.Open str, conn
'
'        If Not rsN.EOF Then
'            txtAccount.Text = rsN!AC_No
'            cmbBank.Text = rsN!Bank_Name
'            cmbBranch.Text = rsN!Branch_Name
'            txtBalance.Text = Format$(Val(rsN!Balance), "#,##0.00")
'        Else
'            Exit Sub
'        End If
End Sub

Private Sub cmbType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtChq_No.SetFocus
End If
End Sub

Private Sub cmbType_LostFocus()
If cmbType.Text = "Withdraw" Then
lblChq_No.Caption = "Cheque No."
End If
If cmbType.Text = "Deposit" Then
lblChq_No.Caption = "Voucher No."
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdReceive_Click()
If cmbType.Text = "Withdraw" Then
    Call Bank_Cr
    Call Cash_Dr
'-----------------------------------------------------------------------------------
Else
    If cmbType.Text = "Deposit" Then
        Call Bank_Dr
        Call Cash_Cr
    End If
End If
Call Account_No
Call Bank_Name
Call Branch_Name

Call clearTextboxes
End Sub

Private Sub Form_Load()
    Call Account_No
    Call Bank_Name
    Call Branch_Name
    Call clearTextboxes
    
    txtDate.Text = Today
    cmdReceive.Enabled = False
    cmdUpdate.Enabled = False
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbBank.SetFocus
End If
End Sub
Private Sub txtAccount_LostFocus()
    
    Dim Account As String
    Account = txtAccount.Text
    
    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Bank_Master where AC_No Like '" & Account & "'order by AC_No"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            cmbBank.Text = rsN!Bank_Name
            cmbBranch.Text = rsN!Branch_Name
            txtBalance.Text = Format$(Val(rsN!Balance), "#,##0.00")
            rsN.Close
            cmdReceive.Enabled = True

        Else
        
        If MsgBox("Bank Account does not exist! Do you want add new Bank Account?", vbInformation + vbYesNo, "Add New") = vbYes Then
            rsN.Close
            If txtAccount.Text = "" Then
                txtAccount.SetFocus
            Else
                cmbBank.Text = ""
                cmbBranch.Text = ""
                txtBalance.Text = Format$(Val(0#), "#,##0.00")
                txtDate.Text = Today
                cmdReceive.Enabled = True
            End If
        Else
        Call clearTextboxes
            cmdReceive.Enabled = False
        Exit Sub
        End If
End If
End Sub
Private Sub txtChq_Amnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDescription.SetFocus
End If
End Sub

Private Sub txtChq_Amnt_LostFocus()
      On Error Resume Next
'        If cmbType.Text = "Withdraw" Then
'            Set rs = New ADODB.Recordset
'            str = "select * from Bank_Master where AC_No like '" & txtAccount.Text & "'"
'            rs.Open str, conn
'
'            If rs!Balance < Val(txtChq_Amnt.Text) Then
'                MsgBox "Insufficiant Balance.!", vbCritical, "Error!"
'                txtChq_Amnt.SetFocus
'                rs.Close
'            Else
'                txtChq_Amnt.Text = Format$(Val(txtChq_Amnt.Text), "###0.00")
'                rs.Close
'            End If
'
'    Else
'
'        If cmbType.Text = "Deposit" Then
'            Set rs = New ADODB.Recordset
'            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
'            rs.Open str, conn
'
'
'            If rs!Balance < Val(txtChq_Amnt.Text) Then
'                MsgBox "Insufficiant Cash Balance.!", vbCritical, "Error!"
'                txtChq_Amnt.SetFocus
'            Else
'                txtChq_Amnt.Text = Format$(Val(txtChq_Amnt.Text), "###0.00")
'            End If
'        End If
'
'    End If
End Sub

Private Sub txtChq_No_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtChq_Amnt.SetFocus
End If
End Sub
Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdReceive.SetFocus
End If
End Sub
