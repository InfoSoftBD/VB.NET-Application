VERSION 5.00
Begin VB.Form frmBranch_Payment 
   BackColor       =   &H00008000&
   Caption         =   "Branch Payment"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5970
   Icon            =   "frmBranch_Payment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0FFC0&
      Height          =   645
      Left            =   135
      ScaleHeight     =   585
      ScaleWidth      =   5655
      TabIndex        =   33
      Top             =   180
      Width           =   5715
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " PAYMENT RECEIVE FROM BRANCH"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -345
         TabIndex        =   34
         Top             =   90
         Width           =   6270
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   660
      Left            =   135
      ScaleHeight     =   600
      ScaleWidth      =   5610
      TabIndex        =   28
      Top             =   5895
      Width           =   5670
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   435
         Left            =   3015
         TabIndex        =   32
         Top             =   90
         Width           =   1020
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   435
         Left            =   4410
         TabIndex        =   31
         Top             =   90
         Width           =   1020
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   435
         Left            =   1590
         TabIndex        =   30
         Top             =   105
         Width           =   1020
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Save"
         Height          =   435
         Left            =   210
         TabIndex        =   29
         Top             =   105
         Width           =   1020
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Payment Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   135
      TabIndex        =   13
      Top             =   3510
      Width           =   5685
      Begin VB.TextBox txtAccount 
         Alignment       =   2  'Center
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
         Left            =   1215
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   630
         Width           =   1725
      End
      Begin VB.OptionButton optCheque 
         BackColor       =   &H00C0FFC0&
         Caption         =   "By Cheque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1215
         TabIndex        =   20
         Top             =   270
         Width           =   1275
      End
      Begin VB.OptionButton optCash 
         BackColor       =   &H00C0FFC0&
         Caption         =   "By Cash"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3735
         TabIndex        =   19
         Top             =   270
         Width           =   1050
      End
      Begin VB.TextBox txtChq_No 
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
         Left            =   1215
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1710
         Width           =   1725
      End
      Begin VB.TextBox txtCash 
         Alignment       =   1  'Right Justify
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
         Left            =   3735
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   630
         Width           =   1725
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
         Left            =   1215
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   1170
         Width           =   1725
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
         Left            =   3735
         TabIndex        =   15
         Text            =   "Combo2"
         Top             =   1170
         Width           =   1725
      End
      Begin VB.TextBox txtChq_Amnt 
         Alignment       =   1  'Right Justify
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
         Left            =   3735
         TabIndex        =   14
         Text            =   "Text7"
         Top             =   1710
         Width           =   1725
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
         Left            =   90
         TabIndex        =   27
         Top             =   675
         Width           =   1080
      End
      Begin VB.Label lblCash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   2970
         TabIndex        =   26
         Top             =   675
         Width           =   675
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
         Left            =   105
         TabIndex        =   25
         Top             =   1755
         Width           =   1020
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
         Left            =   2985
         TabIndex        =   24
         Top             =   1755
         Width           =   690
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
         Left            =   3060
         TabIndex        =   23
         Top             =   1215
         Width           =   615
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
         Left            =   105
         TabIndex        =   22
         Top             =   1215
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Branch Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   135
      TabIndex        =   0
      Top             =   990
      Width           =   5685
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
         Left            =   4230
         TabIndex        =   6
         Text            =   "Text3"
         Top             =   315
         Width           =   1275
      End
      Begin VB.TextBox txtInvoice 
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
         Left            =   1605
         TabIndex        =   5
         Text            =   "Text4"
         Top             =   315
         Width           =   1275
      End
      Begin VB.TextBox txtVendor_Code 
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
         Left            =   1620
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   810
         Width           =   1275
      End
      Begin VB.TextBox txtVendor_Address 
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
         Left            =   1620
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1800
         Width           =   3840
      End
      Begin VB.ComboBox cmbVendor_Name 
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
         Left            =   1620
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1305
         Width           =   2535
      End
      Begin VB.TextBox txtDue 
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
         Left            =   4230
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   810
         Width           =   1275
      End
      Begin VB.Label Label3 
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
         Left            =   3735
         TabIndex        =   12
         Top             =   405
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
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
         Left            =   405
         TabIndex        =   11
         Top             =   405
         Width           =   1080
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Code"
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
         Left            =   360
         TabIndex        =   10
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
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
         Left            =   300
         TabIndex        =   9
         Top             =   1350
         Width           =   1185
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Address"
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
         Left            =   90
         TabIndex        =   8
         Top             =   1845
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due Amount"
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
         TabIndex        =   7
         Top             =   810
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmBranch_Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Today As Date
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub Cash_Cr()
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtCash.Text)
        rs!Date = txtDate.Text
        rs.Update
            
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Vendor Code-" & txtVendor_Code.Text
        rsN!Dr = Val(txtCash.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        
        Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!MR_No = txtInvoice.Text
        rsN!Description = "Branch Code-" & txtVendor_Code.Text
        rsN!Dr = Val(txtCash.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Cash_Dr = rsU!Cash_Dr + Val(txtCash.Text)
        rsU!Cash_Close = rsU!Cash_Close + Val(txtCash.Text)
        rsU.Update
        rsU.Close
        
        
     Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100102 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtCash.Text)
        rs.Update
        
     Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Vendor Code-" & txtVendor_Code.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtCash.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close

        rs.Close
        
    Set rs = New ADODB.Recordset
        str = "select * from Branch where Branch_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Cr = rs!Cr + Val(txtCash.Text)
        rs!Balance = rs!Balance - Val(txtCash.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Branch_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!Branch_Code = txtVendor_Code.Text
        rsN!Branch_Name = cmbVendor_Name.Text
        'rsN!Branch_Address = txtVendor_Address.Text
        rsN!Description = "Cash Payment"
        rsN!Dr = 0
        rsN!Cr = Val(txtCash.Text)
        rsN!Balance = rs!Balance
        
        rsN.Update
        rsN.Close
        rs.Close
    End If
    
End Sub
Private Sub Bank_Cr()
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
        rsN!Description = "Branch Code-" & txtVendor_Code.Text
        rsN!Dr = Val(txtChq_Amnt.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
        
    Set rs = New ADODB.Recordset
        str = "select * from Bank_Master where AC_No like '" & txtAccount.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
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
        rsN!Chq_No = "Chq-" & txtChq_No.Text
        rsN!Description = "Branch Code-" & txtVendor_Code.Text
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
        
        
     Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100102 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs.Update
        
     Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Vendor Code-" & txtVendor_Code.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close

        rs.Close

Set rs = New ADODB.Recordset
        str = "select * from Branch where Branch_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Cr = rs!Cr + Val(txtChq_Amnt.Text)
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Branch_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!Branch_Code = txtVendor_Code.Text
        rsN!Branch_Name = cmbVendor_Name.Text
        rsN!Branch_Address = txtVendor_Address.Text
        rsN!Description = "Chq-" & txtChq_No.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        
        rsN.Update
        rsN.Close
        rs.Close
    End If
End Sub
Private Sub Vendor_Name()
        cmbVendor_Name.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Branch_Name FROM Branch"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbVendor_Name.AddItem rsN!Branch_Name
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub Bank_Name()
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
Private Sub Branch_Name()
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
        txtInvoice.Text = ""
        txtVendor_Code.Text = ""
        cmbVendor_Name.Text = ""
        txtVendor_Address.Text = ""
        txtDue.Text = ""
                        
        txtCash.Text = ""
        txtAccount.Text = ""
        cmbBank.Text = ""
        cmbBranch.Text = ""
        txtChq_No.Text = ""
        txtChq_Amnt.Text = ""
End Sub

Private Sub cmbBank_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbBranch.SetFocus
End If
End Sub

Private Sub cmbBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtChq_No.SetFocus
End If
End Sub

Private Sub cmbBranch_LostFocus()
Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Bank_Master where Bank_Name Like '" & cmbBank.Text & "' AND Branch_Name Like '" & cmbBranch.Text & "' order by Bank_Name"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtAccount.Text = rsN!AC_No
            cmbBank.Text = rsN!Bank_Name
            cmbBranch.Text = rsN!Branch_Name
        Else
            Exit Sub
        End If
End Sub

Private Sub cmbProd_Model_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty.Text)
    txtQty.SetFocus
End If
End Sub

Private Sub cmbProd_Model_LostFocus()
Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Name Like '" & cmbProd_Name.Text & "' AND Prod_Model Like '" & cmbProd_model.Text & "' order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_model.Text = rsN!Prod_Model
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtSale_Price.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            rsN.Close
        Else
            Exit Sub
        End If
End Sub

Private Sub cmbProd_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProd_model.SetFocus
End If
End Sub

Private Sub cmbProd_Name_LostFocus()
Dim Prod As String
Prod = cmbProd_Name.Text
        
        cmbProd_model.Clear
        Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Model FROM Prod_Master where Prod_Name Like '" & Prod & "'"
        rsN.Open str, conn
        If Not rsN.EOF Then
        
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbProd_model.AddItem rsN!Prod_Model
        rsN.MoveNext
        Loop
        rsN.Close
        Else
        rsN.Close
        End If
    
    Prod = txtProd_Sl.Text

    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_model.Text = rsN!Prod_Model
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtDate.Text = Date
            txtSale_Price.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            rsN.Close
        Else
        Exit Sub
        End If

End Sub

Private Sub cmbVendor_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtVendor_Address.SetFocus
End If
End Sub

Private Sub cmbVendor_Name_LostFocus()
Dim Vendor As String
Dim mid As Integer
Vendor = cmbVendor_Name.Text
        Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Branch where Branch_Name Like '" & Vendor & "'order by Branch_code"
        rsN.Open str, conn
        
If Not rsN.EOF Then
            txtVendor_Code.Text = rsN!Branch_Code
            cmbVendor_Name.Text = rsN!Branch_Name
            txtVendor_Address.Text = rsN!Branch_Address
            txtDue.Text = Format(Val(rsN!Balance), "###0.00")
            rsN.Close
Else
        
    MsgBox "Invalid Branch Code", vbCritical, "Error!"
                
        rsN.Close
        Exit Sub
End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdReceive_Click()
If txtInvoice.Text = "" Then
MsgBox "Please Input Invoice No!", vbCritical
txtInvoice.SetFocus
Exit Sub
End If
    
    
    If optCash.Value = True Then
        Call Cash_Cr
    Else
   
        If optCheque.Value = True Then
            Call Bank_Cr
        End If
    End If
    
    Call clearTextboxes
    txtDate.Text = Today
Exit Sub
   Resume Next
End Sub
Private Sub cmdUpdate_Click()
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Stock where Bom_Code like '" & txtBom.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
        rsU!issue = rsU!issue + Val(txtRqty.Text)
        rsU!Stock = rsU!Opening_Bal + rsU!Receive - rsU!issue
        rsU!Unit_Price = Val(txtPrice.Text)
        rsU!Amount = Val(rsU!Stock) * Val(rsU!Unit_Price)
        rsU.Update
        rsU.Close
    Else
        MsgBox "There is no such Product found in the stock to update", vbCritical + vbOKOnly
        rsU.Close
    End If
    
    Call clearTextboxes
    txtRdate.Text = Date
Exit Sub
   Resume Next
End Sub

Private Sub Form_Load()
On Error Resume Next
   
   
    Call Vendor_Name
    Call Bank_Name
    Call Branch_Name
    Call clearTextboxes
        
        lblCash.Visible = True
        txtCash.Visible = True
        optCash.Value = True
        lblBank.Visible = False
        lblBranch.Visible = False
        cmbBank.Visible = False
        cmbBranch.Visible = False
        lblChq_No.Visible = False
        lblChq_Amnt.Visible = False
        txtChq_No.Visible = False
        txtChq_Amnt.Visible = False
        
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
        
        cmdReceive.Enabled = False
        cmdUpdate.Enabled = False
    Exit Sub
End Sub

Private Sub optCash_Click()
lblCash.Visible = True
txtCash.Visible = True
lblAccount.Visible = False
txtAccount.Visible = False
lblBank.Visible = False
lblBranch.Visible = False
cmbBank.Visible = False
cmbBranch.Visible = False
lblChq_No.Visible = False
lblChq_Amnt.Visible = False
txtChq_No.Visible = False
txtChq_Amnt.Visible = False
End Sub

Private Sub optCheque_Click()
lblCash.Visible = False
txtCash.Visible = False
lblAccount.Visible = True
txtAccount.Visible = True
lblBank.Visible = True
lblBranch.Visible = True
cmbBank.Visible = True
cmbBranch.Visible = True
lblChq_No.Visible = True
lblChq_Amnt.Visible = True
txtChq_No.Visible = True
txtChq_Amnt.Visible = True
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
            rsN.Close
            
            
        Else
        MsgBox "Invalid Account No!", vbCritical, "Error!"
        rsN.Close
        Exit Sub
        End If
        
End Sub

Private Sub txtCash_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmdReceive.Enabled = True Then
cmdReceive.SetFocus
Else
Exit Sub
End If
End If
End Sub

Private Sub txtCash_LostFocus()
txtCash.Text = Format$(Val(txtCash.Text), "###0.00")
End Sub

Private Sub txtChq_Amnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmdReceive.Enabled = True Then
cmdReceive.SetFocus
Else
Exit Sub
End If
End If
End Sub

Private Sub txtChq_Amnt_LostFocus()
txtChq_Amnt.Text = Format$(Val(txtChq_Amnt.Text), "###0.00")
End Sub

Private Sub txtChq_No_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtChq_Amnt.SetFocus
End If
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtInvoice.Text = "" Then
Exit Sub
Else
txtVendor_Code.SetFocus
End If
End If
End Sub
Private Sub txtVendor_Address_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If optCash.Value = True Then
        txtCash.SetFocus
        Else
            If optCheque.Value = True Then
            txtAccount.SetFocus
            End If
    End If
End If
End Sub
Private Sub txtVendor_Code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbVendor_Name.SetFocus
End If
End Sub

Private Sub txtVendor_Code_LostFocus()
Dim Vendor As String
Dim mid As Integer
Vendor = txtVendor_Code.Text
        Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Branch where Branch_Code Like '" & Vendor & "'order by Branch_code"
        rsN.Open str, conn
        
If Not rsN.EOF Then
            txtVendor_Code.Text = rsN!Branch_Code
            cmbVendor_Name.Text = rsN!Branch_Name
            txtVendor_Address.Text = rsN!Branch_Address
            txtDue.Text = Format(Val(rsN!Balance), "###0.00")
            rsN.Close
            cmdReceive.Enabled = True
Else
        
    MsgBox "Invalid Branch Code", vbCritical, "Error!"
                
        rsN.Close
        cmdReceive.Enabled = False
        Exit Sub
End If

End Sub


