VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSBBack 
   BackColor       =   &H00008000&
   Caption         =   "Old Data Posting"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11715
   Icon            =   "frmSBBack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0FFC0&
      Height          =   5550
      Left            =   7155
      ScaleHeight     =   5490
      ScaleWidth      =   4275
      TabIndex        =   32
      Top             =   135
      Width           =   4335
      Begin VB.TextBox txtSearch 
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
         Left            =   2280
         TabIndex        =   33
         ToolTipText     =   "Enter Consumer no."
         Top             =   180
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmSBBack.frx":0442
         Height          =   4560
         Left            =   240
         TabIndex        =   34
         Top             =   735
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   8043
         _Version        =   393216
         BackColor       =   12648384
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Transaction"
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
         TabIndex        =   35
         Top             =   255
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   135
      ScaleHeight     =   675
      ScaleWidth      =   6780
      TabIndex        =   27
      Top             =   4950
      Width           =   6840
      Begin VB.CommandButton Command5 
         Caption         =   "Update"
         Height          =   435
         Left            =   3525
         TabIndex        =   31
         Top             =   150
         Width           =   1395
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   435
         Left            =   1890
         TabIndex        =   30
         Top             =   150
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   435
         Left            =   210
         TabIndex        =   29
         Top             =   150
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   435
         Left            =   5205
         TabIndex        =   28
         Top             =   150
         Width           =   1395
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         Height          =   375
         Left            =   5085
         TabIndex        =   12
         Top             =   2955
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   765
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
         TabIndex        =   8
         Top             =   1305
         Width           =   2715
      End
      Begin VB.TextBox txtDefaulter 
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
         TabIndex        =   7
         Top             =   2430
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
         Top             =   2415
         Width           =   1455
      End
      Begin VB.TextBox txtOverdue 
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
         Left            =   1620
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   2970
         Width           =   1455
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
         TabIndex        =   26
         Top             =   270
         Width           =   405
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
         TabIndex        =   25
         Top             =   810
         Width           =   1200
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Present Balance"
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
         Left            =   3555
         TabIndex        =   24
         Top             =   3015
         Width           =   1440
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   1395
         Width           =   1485
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defaulter"
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
         Left            =   4215
         TabIndex        =   19
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Balance"
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
         Left            =   450
         TabIndex        =   18
         Top             =   2490
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Overdue"
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
         Left            =   810
         TabIndex        =   17
         Top             =   3045
         Width           =   705
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
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   135
      ScaleHeight     =   675
      ScaleWidth      =   6780
      TabIndex        =   0
      Top             =   135
      Width           =   6840
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SAVINGS OLD DATA POSTING"
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
         Left            =   810
         TabIndex        =   1
         Top             =   90
         Width           =   5190
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7290
      Top             =   5535
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSBBack"
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
        txtDefaulter.Text = ""
        txtInstallment.Text = "0.00"
        txtBalance.Text = "0.00"
        txtDescription.Text = """"
        txtNet.Text = "0.00"
        txtOverdue.Text = "0.00"
        txtMR_No.Text = ""
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

    Dim Week As Date
    On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Deposit_Master where AC_No like '" & txtAccount.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
        rsU!Amount = rsU!Amount + Val(txtNet.Text)
        'rsU!Inst_Date = txtDate.Text
        'rsU!Payment = rsU!Payment + Val(txtNet.Text)
        'rsU!Today_Pay = rsU!Today_Pay + Val(txtNet.Text)
         rsU!Due = Val(txtOverdue.Text)
         rsU!Term_Fail = Val(txtDefaulter.Text)
        
        rsU.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = Today
        rsN!AC_No = txtAccount.Text
        rsN!MR_No = rsN!sl
        rsN!Name = txtName.Text
        rsN!Description = txtDescription.Text + " MN" & Format$(Val(rsN!sl), "###0")
        rsN!Cr = txtNet.Text
        rsN!Fine = 0
        rsN!Balance = rsU!Amount
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
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C No:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
    
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!MR_No = rsN!sl
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Cash_Dr = rsU!Cash_Dr + Val(txtNet.Text)
        rsU!Cash_Close = rsU!Cash_Close + Val(txtNet.Text)
        rsU.Update
        rsU.Close
    
        
        If txtType.Text = "BM-D" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If

    If txtType.Text = "BM-W" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If

    If txtType.Text = "BM-M" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If

    If txtType.Text = "BM-Y" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If

    If txtType.Text = "D2Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If

    If txtType.Text = "D3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
        End If

    If txtType.Text = "D4Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If

    If txtType.Text = "D5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If

    If txtType.Text = "W3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100105 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If

    If txtType.Text = "W5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100105 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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

    End If

    If txtType.Text = "W7Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100105 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If

    If txtType.Text = "M5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If

    If txtType.Text = "M10Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If
 
    If txtType.Text = "M15Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If
 
    If txtType.Text = "M20Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
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
    End If
 
    If txtType.Text = "Y10Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100128 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "LP3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "LP5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "LP8Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "LP10Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "LP12Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "FD2Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "FD3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "FD4Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "FD5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "FD5.5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "MFD24Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "MFD36Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "MFD48Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "MFD60Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If

    If txtType.Text = "MFD66Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
        
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
    txtDescription.Text = "Old Balance"
    txtAccount.Text = ac
    txtAccount.SelStart = 7
    txtAccount.SelLength = Len(txtAccount.Text)
    txtAccount.SetFocus
    
    'Timer1.Enabled = True

Exit Sub
   Resume Next
End Sub

Private Sub Command3_Click()
    
    Set rs = New ADODB.Recordset
        str = "select * from Tran where sl like '" & txtMR_No.Text & "'"
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
        txtNet.Text = rs!Cr
        txtBalance.Text = rs!Balance
        rs.Close
     
     If MsgBox("Really want to delete?", vbCritical + vbYesNo) = vbYes Then
        
        str = "delete from Tran where sl like '" & txtMR_No.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.Close
        
    Set rsU = New ADODB.Recordset
        str = "select * from Deposit_Master where AC_No like '" & txtAccount.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
        rsU!Amount = rsU!Amount - Val(txtNet.Text)
        'rsU!Inst_Date = txtDate.Text
        'rsU!Payment = rsU!Payment + Val(txtNet.Text)
        'rsU!Today_Pay = rsU!Today_Pay + Val(txtNet.Text)
         'rsU!Due = Val(txtOverdue.Text)
         'rsU!Term_Fail = Val(txtDefaulter.Text)
        
        rsU.Update
    
   
    
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
        rsN!Description = "A/C No.-" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtNet.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
    
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!MR_No = txtMR_No.Text
        rsN!Description = "A/C: " + txtAccount.Text
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
    
        
        If txtTerm.Text = "Daily" Then
            
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100104 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(txtNet.Text)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C No.-" & txtAccount.Text
            rsN!Dr = Val(txtNet.Text)
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close

        End If


        If txtTerm.Text = "Weekly" Then
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100105 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(txtNet.Text)
            rs.Update
       Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C No.-" & txtAccount.Text
            rsN!Dr = Val(txtNet.Text)
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        
        End If
        
        If txtTerm.Text = "Monthly" Then
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100106 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(txtNet.Text)
            rs.Update
            
        Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C No.-" & txtAccount.Text
            rsN!Dr = Val(txtNet.Text)
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        End If
        
        If txtTerm.Text = "Yearly" Then
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100128 & "'"
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
            rsN!Dr = Val(txtNet.Text)
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        End If
        
    End If
    
    Call clearTextboxes
        
        str = "select * from Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by sl"
        rs.Open str, conn, adOpenKeyset, adLockReadOnly
        
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        Call ColumnWidth
        End If
    Else
        MsgBox "There is no such Transaction no. found: " & txtMR_No.Text
        rs.Close
    End If
        Command3.Enabled = False
        Command5.Enabled = False
        Command2.Enabled = False
        txtAccount.Text = ac
        txtDate.Text = Today
    Exit Sub
    Resume Next
End Sub

Private Sub Command4_Click()
Unload Me
frmPalli_Print.Show 1
End Sub

Private Sub Command5_Click()

Set rsU = New ADODB.Recordset
        str = "select * from Deposit_Master where AC_No like '" & txtAccount.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
        rsU!Amount = rsU!Amount + Val(txtNet.Text)
        rsU!Due_Date = txtDate.Text
        rsU!Due = Val(txtOverdue.Text)
        rsU!Term_Fail = Val(txtDefaulter.Text)
        
        rsU.Update
        rsU.Close
    Call clearTextboxes
           
        
    Else
        MsgBox "There is no such Account No. found.", 64, "Update Error"
        rsU.Close
        Exit Sub
    End If
        
    Set rs = New ADODB.Recordset
        str = "select * from Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by AC_No"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
    
    Call clearTextboxes
    
    Command5.Enabled = False
    txtDate.Text = Today
    txtDescription.Text = "Old Balance"
    txtAccount.Text = ac
    txtAccount.SelStart = 7
    txtAccount.SelLength = Len(txtAccount.Text)
    txtAccount.SetFocus
End Sub

Private Sub Form_Load()
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
           txtDescription.Text = "Old Balance"
        End If

Set rs = New ADODB.Recordset
        str = "select * from Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by Sl"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close

        Command2.Enabled = False
        Command3.Enabled = False
        Command5.Enabled = False
        Call ColumnWidth
    Exit Sub
End Sub

Private Sub txtAccbill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtVat.SelStart = 0
    txtVat.SelLength = Len(txtVat.Text)
    txtVat.SetFocus
End If
End Sub

Private Sub txtAccbill_LostFocus()
On Error Resume Next
Dim a, b, c, d As Integer
    a = Val(txtAccbill.Text)
    b = Val(txtSer.Text)
    c = Val(txtOther.Text)
txtVat.Text = Format$(Round(Val(a + b) * 0.05), "###0.00")
    d = Val(txtVat.Text)
txtNetbill.Text = Format$(Round(a + b + c + d), "###0.00")
txtAccbill.Text = Format$(Val(txtAccbill.Text), "###0.00")
txtSer.Text = Format$(Val(txtSer.Text), "###0.00")
txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
    txtVat.SelStart = 0
    txtVat.SelLength = Len(txtVat.Text)
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
        str = "select * from Deposit_Master where AC_No like '" & ID & "' Or Old_Ac like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
        If rs!C_lose = "Yes" Then
                MsgBox "Account already close!", vbCritical, "Close Info"
                Call clearTextboxes
            Else
               If rs!Term_Fail > 1 Then
               MsgBox "Member is Defaulter for " & rs!Term_Fail & "Term", vbCritical, "Information!"
               End If
               
               'Call clearTextboxes
                On Error Resume Next
        
                txtAccount.Text = rs!AC_No
                txtDate.Text = Today
                txtName.Text = rs!Name
                txtType.Text = rs!Type
                txtTerm.Text = rs!Term
                txtDefaulter.Text = rs!Term_Fail
                txtOverdue.Text = rs!Due
                txtInstallment.Text = rs!Installment
                txtBalance.Text = rs!Amount
                rs.Close
                
                txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
                txtOverdue.Text = Format$(Val(txtOverdue.Text), "###0.00")
                txtInstallment.Text = Format$(Val(txtInstallment.Text), "###0.00")
                txtBalance.Text = Format$(Val(txtBalance.Text), "###0.00")
        
        Command2.Enabled = True
        Command5.Enabled = True
        End If
    Else
    MsgBox "There is no such Account no. found.,", vbCritical
        rs.Close
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

Private Sub txtDefaulter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command2.SetFocus
End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNet.SelStart = 0
    txtNet.SelLength = Len(txtNet.Text)
    txtNet.SetFocus
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
        txtNet.Text = rs!Cr
        txtBalance.Text = rs!Balance
        rs.Close
        Command3.Enabled = True
        'Command5.Enabled = True
        Command2.Enabled = False
        
    Else
    If MsgBox("There is no such Transaction no found! Verify Transaction no.?", vbCritical + vbYesNo) = vbYes Then
        Call clearTextboxes
        rs.Close
        txtTran.SetFocus
    Else
        Call clearTextboxes
        Combo1.SetFocus
    End If
    End If
    Exit Sub

End Sub

Private Sub txtNet_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtOverdue.SelStart = 0
    txtOverdue.SelLength = Len(txtOverdue.Text)
    txtOverdue.SetFocus
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
    txtNet.Text = Format$(Val(txtInstallment.Text) + Val(txtOther.Text), "###0.00")
    txtInstallment.Text = Format$(Val(txtInstallment.Text), "###0.00")
    txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
    txtNet.SelStart = 0
    txtNet.SelLength = Len(txtNet.Text)
End Sub

Private Sub txtNet_LostFocus()
txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If txtSearch.Text = "" Then
Exit Sub
End If

    Dim search As String
        search = txtSearch.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Tran where AC_no Like '" & "%" & search & "%" & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
        Call ColumnWidth
        
        Else
        
    On Error Resume Next
               
    Set rs = New ADODB.Recordset
        str = "select * from Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') order by sl"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
               
        Call ColumnWidth
    End If
    Exit Sub
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

Private Sub txtOverdue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 txtDefaulter.SelStart = 0
    txtDefaulter.SelLength = Len(txtDefaulter.Text)
    txtDefaulter.SetFocus
End If
End Sub
