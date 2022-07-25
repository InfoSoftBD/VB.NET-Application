VERSION 5.00
Begin VB.Form frmSBClose 
   BackColor       =   &H00008000&
   Caption         =   "Savings Account Closing"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7185
   Icon            =   "frmSBClose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   135
      ScaleHeight     =   675
      ScaleWidth      =   6780
      TabIndex        =   27
      Top             =   180
      Width           =   6840
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SAVINGS ACCOUNT CLOSING"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   825
         TabIndex        =   28
         Top             =   90
         Width           =   5160
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   3630
      Left            =   135
      ScaleHeight     =   3570
      ScaleWidth      =   6780
      TabIndex        =   6
      Top             =   1095
      Width           =   6840
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   5115
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   3060
         Width           =   1455
      End
      Begin VB.TextBox txtOther 
         Alignment       =   1  'Right Justify
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
         Left            =   1635
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   2475
         Width           =   1455
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   1620
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   3045
         Width           =   1455
      End
      Begin VB.TextBox txtTerm 
         Alignment       =   2  'Center
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
         Left            =   5085
         TabIndex        =   15
         Top             =   1305
         Width           =   1455
      End
      Begin VB.TextBox txtInstallment 
         Alignment       =   1  'Right Justify
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
         Left            =   5085
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   1890
         Width           =   1455
      End
      Begin VB.TextBox txtName 
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
         Left            =   1620
         TabIndex        =   13
         Top             =   1305
         Width           =   2715
      End
      Begin VB.TextBox txtAccount 
         Alignment       =   2  'Center
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
         Left            =   1620
         TabIndex        =   12
         Top             =   750
         Width           =   1860
      End
      Begin VB.TextBox txtMR_No 
         Alignment       =   2  'Center
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
         Left            =   1620
         TabIndex        =   11
         Top             =   180
         Width           =   1065
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   1620
         TabIndex        =   10
         Top             =   1890
         Width           =   2355
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
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
         Left            =   5085
         TabIndex        =   9
         Top             =   2460
         Width           =   1455
      End
      Begin VB.TextBox txtType 
         Alignment       =   2  'Center
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
         Left            =   5085
         TabIndex        =   8
         Top             =   765
         Width           =   1455
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
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
         Left            =   5085
         TabIndex        =   7
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Left            =   3735
         TabIndex        =   32
         Top             =   3135
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction"
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
         Left            =   630
         TabIndex        =   30
         Top             =   2565
         Width           =   870
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Installment"
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
         Left            =   4035
         TabIndex        =   26
         Top             =   1980
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
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
         Left            =   4545
         TabIndex        =   25
         Top             =   1350
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Member"
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
         Left            =   30
         TabIndex        =   24
         Top             =   1395
         Width           =   1485
      End
      Begin VB.Label Label2 
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
         Left            =   435
         TabIndex        =   23
         Top             =   825
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MR No."
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
         Left            =   855
         TabIndex        =   22
         Top             =   225
         Width           =   660
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   540
         TabIndex        =   21
         Top             =   1980
         Width           =   975
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
         Left            =   4275
         TabIndex        =   20
         Top             =   2520
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
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
         Left            =   3780
         TabIndex        =   19
         Top             =   810
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Tk."
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
         Left            =   465
         TabIndex        =   18
         Top             =   3120
         Width           =   1050
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
         Left            =   4575
         TabIndex        =   17
         Top             =   270
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0FFC0&
      Height          =   780
      Left            =   135
      ScaleHeight     =   720
      ScaleWidth      =   6780
      TabIndex        =   0
      Top             =   4905
      Width           =   6840
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   435
         Left            =   5475
         TabIndex        =   5
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   435
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   435
         Left            =   1440
         TabIndex        =   3
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Print"
         Height          =   435
         Left            =   4110
         TabIndex        =   2
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Update"
         Height          =   435
         Left            =   2805
         TabIndex        =   1
         Top             =   150
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmSBClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Today As Date
Dim ac As String
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub clearTextboxes()
        txtAccount.Text = ""
        txtDate.Text = ""
        txtName.Text = ""
        txtType.Text = ""
        txtTerm.Text = ""
        txtInstallment.Text = "0.00"
        txtBalance.Text = "0.00"
        txtNet.Text = "0.00"
End Sub
Private Sub ColumnWidth()
        DataGrid1.Columns(0).Width = 800
        DataGrid1.Columns(1).Width = 1200
        DataGrid1.Columns(2).Width = 1500
        DataGrid1.Columns(3).Width = 1000
        DataGrid1.Columns(4).Width = 500
        DataGrid1.Columns(5).Width = 500
        DataGrid1.Columns(6).Width = 800
        DataGrid1.Columns(7).Width = 1000
        DataGrid1.Columns(8).Width = 500
        DataGrid1.Columns(9).Width = 500
        DataGrid1.Columns(10).Width = 500

End Sub
Private Sub Check1_Click()
On Error Resume Next
If Check1.Value = 1 Then
Picture3.Visible = False
Image1.Visible = True
Else
Image1.Visible = False
Picture3.Visible = True
End If
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Command2.SetFocus
End If
If Command2.Enabled = False Then
Command5.SetFocus
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    If Combo1.Text = "" Then
        MsgBox "Please Input Transaction Code", 65, "Transaction Error"
        Combo1.SetFocus
        Exit Sub
    End If
    If Combo1.Text = 101 Or Combo1.Text = 202 Or Combo1.Text = 303 Then
        txtAccount.SelStart = 0
        txtAccount.SelLength = Len(txtAccount.Text)
        txtAccount.SetFocus
    Else
        MsgBox "Wrong Transaction Code", 65, "Transaction Error"
        Combo1.Text = ""
        Combo1.SetFocus
    End If
End If
End Sub

Private Sub Combo1_LostFocus()
     On Error Resume Next
     If Combo1.Text = 101 Then
        lblTransaction.Caption = "Cash Transaction"
    End If
    If Combo1.Text = 202 Then
        lblTransaction.Caption = "Clearing Transaction"
    End If
    If Combo1.Text = 303 Then
        lblTransaction.Caption = "Transfer Transaction"
    End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Dim vn As String
    Dim Week As Date

    Set rsU = New ADODB.Recordset
        str = "select * from Deposit_Master where AC_No like '" & txtAccount.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
        rsU!Amount = rsU!Amount - Val(txtTotal.Text)
        rsU!Fine = rsU!Fine - Val(txtOther.Text)
        rsU!C_lose = "Yes"
    
        rsU!Withdraw = rsU!Withdraw + Val(txtTotal.Text)
        rsU!Daily_Draw = rsU!Daily_Draw + Val(txtTotal.Text)
        rsU!Weekly_Draw = rsU!Weekly_Draw + Val(txtTotal.Text)
        rsU!Monthly_Draw = rsU!Monthly_Draw + Val(txtTotal.Text)
        rsU!Yearly_Draw = rsU!Yearly_Draw + Val(txtTotal.Text)
        rsU.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = txtAccount.Text
        rsN!MR_No = rsN!sl
            vn = rsN!sl
        rsN!Name = txtName.Text
        rsN!Description = txtDescription.Text + "VN" & Format$(Val(rsN!sl), "###0")
        rsN!Dr = txtNet.Text
        rsN!Fine = Val(txtOther.Text)
        rsN!Balance = rsU!Amount
        rsN!Term = txtTerm.Text
        rsN!Center_name = rsU!Center_name
        rsN!Center_code = rsU!Center_code
        rsN!Samity_Name = rsU!Samity_Name
        rsN!Samity_Code = rsU!Samity_Code
        rsN!FO_Name = rsU!FO_Name
        rsN!FO_Code = rsU!FO_Code
        rsN!DPO_Name = rsU!DPO_Name
        rsN!DPO_Code = rsU!DPO_Code
        rsN.Update
        rsN.Close
        rsU.Close
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!MR_No = vn
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Cash_Cr = rsU!Cash_Cr + Val(txtNet.Text)
        rsU!Cash_Close = rsU!Cash_Close - Val(txtNet.Text)
        rsU.Update
        rsU.Close

'----------------------------------------------------------------------------
    If txtType.Text = "BM-D" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "BM-W" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If
    
    If txtType.Text = "BM-M" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If
    
    If txtType.Text = "BM-Y" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If
    
    If txtType.Text = "D2Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "D3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "D4Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "D5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "W3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100105 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "W5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100105 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "W7Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100105 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "M5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "M10Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If
 
    If txtType.Text = "M15Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If
 
    If txtType.Text = "M20Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If
 
    If txtType.Text = "Y10Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100128 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "LP3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "LP5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "LP8Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "LP10Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "LP12Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "FD2Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "FD3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "FD4Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "FD5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "FD5.5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "MFD24Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "MFD36Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "MFD48Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "MFD60Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "MFD66Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100112 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtOther.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If
'----------------------------------------------------------------------------
    End If
    
    Set rs = New ADODB.Recordset
        str = "select * from Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by AC_No"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
    
    Call clearTextboxes
    Command2.Enabled = False
    
    txtDate.Text = Today
    txtDescription.Text = "Savings Withdraw"
    txtAccount.Text = ac
    txtAccount.SelStart = 7
    txtAccount.SelLength = Len(txtAccount.Text)
    txtAccount.SetFocus

Exit Sub
   Resume Next
Exit Sub
   Resume Next
End Sub

'Private Sub Command3_Click()
'On Error Resume Next
 'Dim Today As Date
 'Today = txtDate.Text
 'str = "select * from Palli_Bill where tr_no like '" & txtTran.Text & "'"
  '  Set rs = New ADODB.Recordset
   ' rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    'If Not rs.EOF Then
    'On Error Resume Next
     '   txtTran.Text = rsN!Tr_no
      '  txtAccount.Text = rsN!Account_no
       ' txtAccbill.Text = rsN!Ac_Bill
       ' txtSer.Text = rsN!Ser_charge
       ' txtOther.Text = rsN!Others
       ' txtVat.Text = rsN!vat
       ' txtNetbill.Text = rsN!Net_bill
       ' Check1.Value = rsN!Revenue
       ' txtDate.Text = rsN!Date
       ' rs.Close
     
    ' If MsgBox("Really want to delete?", vbCritical + vbYesNo) = vbYes Then
        
     '   str = "delete from Palli_Bill where Tr_no like '" & txtTran.Text & "'"
      '  rs.Open str, conn, adOpenDynamic, adLockOptimistic
       ' rs.Close
    'Call clearTextboxes
        
     '   str = "select * from Palli_Bill where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by Tr_no Desc"
      '  rs.Open str, conn, adOpenKeyset, adLockReadOnly
        
       ' Adodc1.RecordSource = str
        'Adodc1.Refresh
        'DataGrid1.Refresh
        'Call ColumnWidth
        'End If
    'Else
     '   MsgBox "There is no such Transaction no. found: " & txtTran.Text
      '  rs.Close
   ' End If
    '    Command3.Enabled = False
     '   Command5.Enabled = False
      '  Command2.Enabled = True
    'Exit Sub
    'Resume Next
'End Sub

Private Sub Command4_Click()
Unload Me
frmPalli_Print.Show 1
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim Today As Date
Today = txtDate.Text

Set rsU = New ADODB.Recordset
    str = "select * from Palli_Bill where Tr_no like '" & txtTran.Text & "'"
    rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    If Not rsU.EOF Then
        rsU!Tr_type = Combo1.Text
        rsU!Account_no = txtAccount.Text
        rsU!Ac_Bill = Val(txtAccbill.Text)
        rsU!Ser_charge = Val(txtSer.Text)
        rsU!Others = Val(txtOther.Text)
        rsU!vat = Val(txtVat.Text)
        rsU!Net_bill = Val(txtNetbill.Text)
        rsU!Revenue = Check1.Value
        If Check1.Value = 1 Then
        rsU!Stamp = 5
        End If
        rsU!Date = txtDate.Text
        rsU.Update
        rsU.Close
    Call clearTextboxes
           
        
    Else
        MsgBox "There is no such Transaction No. found.", 64, "Update Error"
        rsU.Close
        Exit Sub
    End If
        
    Set rs = New ADODB.Recordset
        str = "select * from Palli_Bill where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by Tr_no Desc"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
    Call ColumnWidth
    Call clearTextboxes
        Command3.Enabled = False
        Command5.Enabled = False
        Command2.Enabled = True
    Exit Sub
    Resume Next
End Sub

Private Sub Form_Load()
   On Error Resume Next
  Dim suf As String
 suf = 101
 On Error Resume Next
       Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
           On Error Resume Next
           Today = rs!Today
           ac = rs!Branch_Code + suf
           rs.Close
           txtDate.Text = Today
           txtAccount.Text = ac
          txtDescription.Text = "Account Close"
        End If
 
        Command2.Enabled = False
        Command3.Enabled = False
        Command5.Enabled = False
    Exit Sub
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription.Text)
    txtDescription.SetFocus
End If
End Sub

Private Sub txtAccount_LostFocus()
If txtAccount.Text = "" Then
Exit Sub
End If

Dim ID As String
Dim L_Id As String
    ID = txtAccount.Text
    
    txtAccount.SelStart = 7
    txtAccount.SelLength = Len(txtAccount.Text)
    L_Id = txtAccount.SelText
    
    
    Set rs = New ADODB.Recordset
        str = "select * from Deposit_Master where AC_No like '" & ID & "' Or Old_Ac like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
        If rs!C_lose = "Yes" Then
            MsgBox "Account already close!", vbCritical, "Close Info"
            Call clearTextboxes
            txtAccount.SetFocus
        Else
        
            Set rsU = New ADODB.Recordset
            str = "select * from Loan_Master where Customer like '" & L_Id & "'"
            rsU.Open str, conn
                
                If Not rsU.EOF Then
                    If rsU!C_lose = "Yes" Then
                        rsU.Close
                    Else
                        MsgBox "Customer Has a Loan A/C: " & rsU!AC_No & " Balance: " & Format$(Val(rsU!Balance), "###0.00"), vbInformation, "Loan Info!"
                        rsU.Close
                    End If
                End If
                
                Set rsU = New ADODB.Recordset
                str = "select * from Loan_Info where G_Id1 like '" & L_Id & "' Or G_Id2 like '" & L_Id & "' Or G_Id3 like '" & L_Id & "' Or G_Id4 like '" & L_Id & "' Or G_Id5 like '" & L_Id & "' Or G_Id6 like '" & L_Id & "' Or G_Id7 like '" & L_Id & "' Or G_Id8 like '" & L_Id & "' Or G_Id9 like '" & L_Id & "' Or G_Id10 like '" & L_Id & "'"
                rsU.Open str, conn

                If Not rsU.EOF Then
                    
                    Set rsN = New ADODB.Recordset
                    str = "select * from Loan_Master where Ac_No like '" & rsU!AC_No & "'"
                    rsN.Open str, conn

                        If Not rsN.EOF Then
                            
                            If rsN!C_lose = "Yes" Then
                                rsN.Close
                                rsU.Close
                             Else
                                 MsgBox "Customer is a Guarantor of Loan A/C: " & rsU!AC_No, vbInformation, "Loan Info!"
                                 rsN.Close
                                 rsU.Close
                             End If
                        End If
                End If
            
            Call clearTextboxes
            On Error Resume Next
                txtAccount.Text = rs!AC_No
                txtDate.Text = Today
                txtName.Text = rs!Name
                txtType.Text = rs!Type
                txtTerm.Text = rs!Term
                txtInstallment.Text = rs!Installment
                txtBalance.Text = rs!Amount
                txtNet.Text = "0"
                rs.Close
        
                txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
                txtInstallment.Text = Format$(Val(txtInstallment.Text), "###0.00")
                txtBalance.Text = Format$(Val(txtBalance.Text), "###0.00")
                Command2.Enabled = True
                
            End If
        
    Else
    MsgBox "There is no such Account no. found.,", vbCritical
        rs.Close
        txtAccount.Text = ac
        txtAccount.SelStart = 7
        txtAccount.SelLength = Len(txtAccount.Text)
        txtAccount.SetFocus
        txtDate.Text = Today
    End If
    Exit Sub
End Sub

Private Sub txtNetbill_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    Check1.SetFocus
End If
End Sub

Private Sub txtNetbill_LostFocus()
On Error Resume Next
Dim Net As Integer
    txtAccbill.Text = Format$(Val(txtAccbill.Text), "###0.00")
    txtSer.Text = Format$(Val(txtSer.Text), "###0.00")
    txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
    txtVat.Text = Format$(Val(txtVat.Text), "###0.00")
    txtNetbill.Text = Format$(Val(txtNetbill.Text), "###0.00")
   
    Net = Val(txtNetbill.Text)
If Net >= 200 Then
    Check1.Value = 1
    Check1.SetFocus
    Picture3.Visible = False
    Image1.Visible = True
   
    Else
    Check1.Value = 0
    Image1.Visible = False
    Picture3.Visible = True
    Command2.SetFocus
    If Command2.Enabled = False Then
    Command5.SetFocus
    End If
End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtOther.SelStart = 0
    txtOther.SelLength = Len(txtOther.Text)
    txtOther.SetFocus
End If
End Sub

Private Sub txtMR_No_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAccount.SetFocus
End If
End Sub

Private Sub txtMR_No_LostFocus()
If txtMR_No.Text = "" Then
Exit Sub
End If

On Error Resume Next
Dim Tran As String
    Tran = txtMR_No.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Tran where sl like '" & Tran & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        txtMR_No.Text = rs!sl
        txtAccount.Text = rs!AC_No
        txtDate.Text = rs!Date
        txtName.Text = rs!Name
        txtType.Text = rs!Type
        txtTerm.Text = rs!Term
        txtDescription.Text = rs!Description
        txtNet.Text = rs!Dr
        txtOther.Text = rs!Fine
        
        txtBalance.Text = rs!Balance
        rs.Close
        
        txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
        txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
        txtTotal.Text = Format$(Val(txtTotal.Text), "###0.00")
        
        Command3.Enabled = True
        Command4.Enabled = True
        'Command5.Enabled = True
        Command2.Enabled = False
        
    Else
    If MsgBox("There is no such Transaction no found! Verify Transaction no.?", vbCritical + vbYesNo) = vbYes Then
        Call clearTextboxes
        rs.Close
        txtMR_No.SetFocus
        txtAccount.Text = ac
        txtDate.Text = Today
    Else
        Call clearTextboxes
        txtAccount.Text = ac
        txtDate.Text = Today
        txtAccount.SetFocus
    End If
    End If
    Exit Sub
End Sub

Private Sub txtNet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command2.SetFocus
End If
End Sub

Private Sub txtNet_LostFocus()
If Val(txtNet.Text) > Val(txtBalance.Text) Then
MsgBox "Insufficient Balance", vbCritical, "Error!"
txtNet.Text = ""
txtNet.SetFocus
Else
txtTotal.Text = Format$(Val(txtNet.Text) + Val(txtOther.Text), "###0.00")
txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
End If
End Sub

Private Sub txtOther_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNet.SelStart = 0
    txtNet.SelLength = Len(txtNet.Text)
    txtNet.SetFocus
End If
End Sub

Private Sub txtOther_LostFocus()
On Error Resume Next
    txtNet.Text = Format$(Val(txtBalance.Text) - Val(txtOther.Text), "###0.00")
    txtInstallment.Text = Format$(Val(txtInstallment.Text), "###0.00")
    txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
    txtNet.SelStart = 0
    txtNet.SelLength = Len(txtNet.Text)
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If txtSearch.Text = "" Then
Exit Sub
End If

    Dim search As String
        search = txtSearch.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Tran where Name Like '" & "%" & search & "%" & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
        Call ColumnWidth
        'Else
        
    'On Error Resume Next
     '   Dim Today As Date
      '  Today = Date
       ' On Error GoTo Last
       ' txtDate.Text = Today
               
    'Set rs = New ADODB.Recordset
     '   str = "select * from Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') order by Tr_no Desc"
      '  rs.Open str, conn
        
       ' Adodc1.ConnectionString = cnStr
        'Adodc1.RecordSource = str
        'Adodc1.Refresh
        'rs.Close
               
        Call ColumnWidth
    End If
    Exit Sub
Last:
    MsgBox ("Database Connection error: " + Err.Description)
 
End Sub

Private Sub txtSer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtOther.SelStart = 0
    txtOther.SelLength = Len(txtOther.Text)
    txtOther.SetFocus
End If
End Sub

Private Sub txtSer_LostFocus()
On Error Resume Next
Dim a, b, c, d As Integer
    a = Val(txtAccbill.Text)
    b = Val(txtSer.Text)
    c = Val(txtOther.Text)
    d = Val(txtVat.Text)
    txtNetbill.Text = Format$(Round(a + b + c + d), "###0.00")
    txtAccbill.Text = Format$(Val(txtAccbill.Text), "###0.00")
    txtSer.Text = Format$(Val(txtSer.Text), "###0.00")
    txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
    txtVat.Text = Format$(Val(txtVat.Text), "###0.00")
    txtOther.SelStart = 0
    txtOther.SelLength = Len(txtOther.Text)
End Sub

Private Sub txtTran_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
    Combo1.SetFocus
End If
End Sub

Private Sub txtTran_LostFocus()
On Error Resume Next
Dim Tran As String
    Tran = txtTran.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Palli_Bill where Tr_no like '" & Tran & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        txtTran.Text = rs!Tr_no
        Combo1.Text = rs!Tr_type
        txtAccount.Text = rs!Account_no
        txtAccbill.Text = rs!Ac_Bill
        txtSer.Text = rs!Ser_charge
        txtOther.Text = rs!Others
        txtVat.Text = rs!vat
        txtNetbill.Text = rs!Net_bill
        Check1.Value = rs!Revenue
        txtDate.Text = rs!Date
        rs.Close
        Command3.Enabled = True
        Command5.Enabled = True
        Command2.Enabled = True
        
    Else
    If MsgBox("There is no such Transaction no found, Do you want add new employee?", vbCritical + vbYesNo) = vbYes Then
        Call clearTextboxes
        rs.Close
    Else
        Call clearTextboxes
    End If
    End If
    Exit Sub
Last:
    MsgBox ("Database Connection error: " + Err.Description)
End Sub

Private Sub txtVat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSer.SelStart = 0
    txtSer.SelLength = Len(txtSer.Text)
    txtSer.SetFocus
End If
End Sub

Private Sub txtVat_LostFocus()
On Error Resume Next
Dim a, b, c, d, e As Double
    a = Val(txtAccbill.Text)
    b = Val(txtSer.Text)
    c = Val(txtOther.Text)

    d = Val(txtVat.Text)
    
    txtNetbill.Text = Format$(Round(a + b + c + d), "###0.00")
    txtAccbill.Text = Format$(Val(txtAccbill.Text), "###0.00")
    txtSer.Text = Format$(Val(txtSer.Text), "###0.00")
    txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
    txtVat.Text = Format$(Val(txtVat.Text), "###0.00")
    
    txtSer.SelStart = 0
    txtSer.SelLength = Len(txtSer.Text)
End Sub





