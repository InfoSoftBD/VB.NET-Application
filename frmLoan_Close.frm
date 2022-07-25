VERSION 5.00
Begin VB.Form frmLoan_Close 
   BackColor       =   &H00008000&
   Caption         =   "Loan Account Closing"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   Icon            =   "frmLoan_Close.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0FFC0&
      Height          =   780
      Left            =   135
      ScaleHeight     =   720
      ScaleWidth      =   6780
      TabIndex        =   25
      Top             =   4905
      Width           =   6840
      Begin VB.CommandButton Command5 
         Caption         =   "Update"
         Height          =   435
         Left            =   2805
         TabIndex        =   30
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Print"
         Height          =   435
         Left            =   4110
         TabIndex        =   29
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   435
         Left            =   1440
         TabIndex        =   28
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   435
         Left            =   120
         TabIndex        =   27
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   435
         Left            =   5475
         TabIndex        =   26
         Top             =   150
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   3630
      Left            =   135
      ScaleHeight     =   3570
      ScaleWidth      =   6780
      TabIndex        =   2
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
         Left            =   5070
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   3060
         Width           =   1500
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
         TabIndex        =   13
         Top             =   195
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
         TabIndex        =   12
         Top             =   765
         Width           =   1455
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5085
         TabIndex        =   11
         Top             =   2460
         Width           =   1455
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
         TabIndex        =   9
         Top             =   180
         Width           =   1065
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
         TabIndex        =   8
         Top             =   750
         Width           =   1860
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
         TabIndex        =   7
         Top             =   1305
         Width           =   2715
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
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   1890
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
         TabIndex        =   5
         Top             =   1305
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
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   3045
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
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   2475
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
         Left            =   3690
         TabIndex        =   32
         Top             =   3135
         Width           =   1275
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
         TabIndex        =   24
         Top             =   270
         Width           =   405
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
         TabIndex        =   23
         Top             =   3120
         Width           =   1050
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
         TabIndex        =   22
         Top             =   810
         Width           =   1200
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
         TabIndex        =   21
         Top             =   2520
         Width           =   705
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
         TabIndex        =   20
         Top             =   1980
         Width           =   975
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
         TabIndex        =   19
         Top             =   225
         Width           =   660
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
         TabIndex        =   18
         Top             =   825
         Width           =   1080
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
         TabIndex        =   17
         Top             =   1395
         Width           =   1485
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
         TabIndex        =   16
         Top             =   1350
         Width           =   435
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
         TabIndex        =   15
         Top             =   1980
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Charge"
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
         Left            =   180
         TabIndex        =   14
         Top             =   2565
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   135
      ScaleHeight     =   675
      ScaleWidth      =   6780
      TabIndex        =   0
      Top             =   180
      Width           =   6840
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOAN ACCOUNT CLOSING"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   975
         TabIndex        =   1
         Top             =   90
         Width           =   4860
      End
   End
End
Attribute VB_Name = "frmLoan_Close"
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
        str = "select * from Loan_Master where AC_No like '" & txtAccount.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
        rsU!C_lose = "Yes"
        rsU!Balance = rsU!Balance + Val(txtNet.Text)
        rsU!Fine = rsU!Fine + Val(txtOther.Text)
        rsU!Inst_Date = txtDate.Text
        rsU!Paid = rsU!Paid + Val(txtNet.Text)
        rsU!Daily_Pay = rsU!Daily_Pay + Val(txtNet.Text)
        rsU!Weekly_Pay = rsU!Weekly_Pay + Val(txtNet.Text)
        rsU!Monthly_Pay = rsU!Monthly_Pay + Val(txtNet.Text)
        rsU!Yearly_Pay = rsU!Yearly_Pay + Val(txtNet.Text)
        rsU!Inst_Paid = rsU!Inst_Paid + 1
        rsU!Due = 0
        rsU!Inst_Due = 0
        rsU!Term_Fail = 0
        
        If rsU!Week_1 = 0 Then
            rsU!Week_1 = rsU!Week_1 + Val(txtNet.Text)
            Else
                If rsU!Week_2 = 0 Then
                    rsU!Week_2 = rsU!Week_2 + Val(txtNet.Text)
                    Else
                    If rsU!Week_3 = 0 Then
                        rsU!Week_3 = rsU!Week_3 + Val(txtNet.Text)
                        Else
                        If rsU!Week_4 = 0 Then
                            rsU!Week_4 = rsU!Week_4 + Val(txtNet.Text)
                            Else
                            If rsU!Week_5 = 0 Then
                                rsU!Week_5 = rsU!Week_5 + Val(txtNet.Text)
                                Else
        
                            End If
                        End If
                    End If
                End If
            End If
        
        rsU.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Loan_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        
        rsN!Date = txtDate.Text
        rsN!AC_No = txtAccount.Text
        rsN!MR_No = rsN!Sl
            vn = rsN!Sl
        rsN!Name = txtName.Text
        rsN!Description = txtDescription.Text + "-MN:" & Format$(Val(rsN!Sl), "###0")
        rsN!Cr = txtNet.Text
        rsN!Fine = txtOther.Text
        rsN!Balance = rsU!Balance
        rsN!Advance = rsU!Advance
        rsN!Term = rsU!Term
        rsN!Type = rsU!Type
        rsN!Daily_Pay = Val(txtNet.Text)
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
    
    
    If txtType.Text = "Product" Then
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100107 & "'"
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
            rs.Close
            
            
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtTotal.Text)
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
    
        Set rsN = New ADODB.Recordset
            rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!MR_No = vn
            rsN!Description = "A/C:" & txtAccount.Text
            rsN!Dr = Val(txtTotal.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
    
        Set rsU = New ADODB.Recordset
            str = "select * from Others"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
            rsU!Cash_Dr = rsU!Cash_Dr + Val(txtTotal.Text)
            rsU!Cash_Close = rsU!Cash_Close + Val(txtTotal.Text)
            rsU.Update
            rsU.Close
    
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
            rsN!Description = "Fine form A/C:" & txtAccount.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtOther.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        End If

End If
'----------------------------------------------------------------------
If txtType.Text = "Cash" Then
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100108 & "'"
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
            rs.Close
            
            rs.Close
            
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtTotal.Text)
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
    
        Set rsN = New ADODB.Recordset
            rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!MR_No = vn
            rsN!Description = "A/C:" & txtAccount.Text
            rsN!Dr = Val(txtTotal.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
    
        Set rsU = New ADODB.Recordset
            str = "select * from Others"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
            rsU!Cash_Dr = rsU!Cash_Dr + Val(txtTotal.Text)
            rsU!Cash_Close = rsU!Cash_Close + Val(txtTotal.Text)
            rsU.Update
            rsU.Close
    
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
            rsN!Description = "Fine form A/C:" & txtAccount.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtOther.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        End If
       
    End If
    
    Set rs = New ADODB.Recordset
        str = "select * from Loan_Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by AC_No"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
End If
    Call clearTextboxes
    Command2.Enabled = False
    
    txtDate.Text = Today
    txtDescription.Text = "Installment Paid"
    txtAccount.Text = ac
    txtAccount.SelStart = 7
    txtAccount.SelLength = Len(txtAccount.Text)
    txtAccount.SetFocus
Exit Sub
   Resume Next
End Sub

Private Sub Command3_Click()
On Error Resume Next
 Dim Today As Date
 Today = txtDate.Text
 str = "select * from Palli_Bill where tr_no like '" & txtTran.Text & "'"
    Set rs = New ADODB.Recordset
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
    On Error Resume Next
        txtTran.Text = rsN!Tr_no
        txtAccount.Text = rsN!Account_no
        txtAccbill.Text = rsN!Ac_Bill
        txtSer.Text = rsN!Ser_charge
        txtOther.Text = rsN!Others
        txtVat.Text = rsN!vat
        txtNetbill.Text = rsN!Net_bill
        Check1.Value = rsN!Revenue
        txtDate.Text = rsN!Date
        rs.Close
     
     If MsgBox("Really want to delete?", vbCritical + vbYesNo) = vbYes Then
        
        str = "delete from Palli_Bill where Tr_no like '" & txtTran.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.Close
    Call clearTextboxes
        
        str = "select * from Palli_Bill where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by Tr_no Desc"
        rs.Open str, conn, adOpenKeyset, adLockReadOnly
        
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        Call ColumnWidth
        End If
    Else
        MsgBox "There is no such Transaction no. found: " & txtTran.Text
        rs.Close
    End If
        Command3.Enabled = False
        Command5.Enabled = False
        Command2.Enabled = True
    Exit Sub
    Resume Next
End Sub

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
  suf = 201
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
    ID = txtAccount.Text
 
    Set rs = New ADODB.Recordset
        str = "select * from Loan_Master where AC_No like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
        If rs!C_lose = "Yes" Then
            MsgBox "Account already close!", vbCritical, "Close Info"
            Call clearTextboxes
            txtAccount.Text = ac
            txtAccount.SelStart = 7
            txtAccount.SetFocus
        Else
            
            Call clearTextboxes
            On Error Resume Next
                txtAccount.Text = rs!AC_No
                txtDate.Text = Today
                txtName.Text = rs!Name
                txtType.Text = rs!Type
                txtTerm.Text = rs!Term
                txtInstallment.Text = rs!Installment
                txtBalance.Text = rs!Balance
                txtNet.Text = "0.00"
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
        txtMR_No.Text = rs!Sl
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
If Val(txtNet.Text) < -(Val(txtBalance.Text)) Then
MsgBox "Insufficient Amount", vbCritical, "Error!"
txtNet.Text = ""
txtNet.SetFocus
Else
txtTotal.Text = Format$((Val(txtNet.Text) + Val(txtOther.Text)), "###0.00")
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
    txtNet.Text = Format$(-(Val(txtBalance.Text)), "###0.00")
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
    
               
        Call ColumnWidth
    End If
    Exit Sub
Last:
    MsgBox ("Database Connection error: " + Err.Description)
 
End Sub












