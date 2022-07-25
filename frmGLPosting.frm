VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmGLPosting 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Posting"
   ClientHeight    =   6465
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   11025
   Icon            =   "frmGLPosting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   6150
      Left            =   6525
      ScaleHeight     =   6120
      ScaleWidth      =   4305
      TabIndex        =   22
      Top             =   180
      Width           =   4335
      Begin VB.TextBox txtSearch 
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
         Left            =   2280
         TabIndex        =   23
         ToolTipText     =   "Enter Consumer no."
         Top             =   180
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmGLPosting.frx":0442
         Height          =   5265
         Left            =   180
         TabIndex        =   24
         Top             =   735
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   9287
         _Version        =   393216
         BackColor       =   16777215
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
         TabIndex        =   25
         Top             =   255
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   135
      ScaleHeight     =   705
      ScaleWidth      =   6225
      TabIndex        =   16
      Top             =   5580
      Width           =   6255
      Begin VB.CommandButton Command5 
         Caption         =   "Update"
         Height          =   435
         Left            =   2580
         TabIndex        =   21
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Print"
         Height          =   435
         Left            =   3840
         TabIndex        =   20
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   435
         Left            =   1350
         TabIndex        =   19
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   435
         Left            =   120
         TabIndex        =   18
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   435
         Left            =   5070
         TabIndex        =   17
         Top             =   150
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   135
      ScaleHeight     =   4365
      ScaleWidth      =   6225
      TabIndex        =   2
      Top             =   1050
      Width           =   6255
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
         Left            =   1800
         TabIndex        =   28
         Text            =   "Combo1"
         Top             =   1680
         Width           =   4020
      End
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
         Left            =   1770
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   990
         Width           =   1545
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
         ItemData        =   "frmGLPosting.frx":0457
         Left            =   1800
         List            =   "frmGLPosting.frx":0461
         TabIndex        =   26
         Text            =   "Combo2"
         Top             =   2340
         Width           =   1440
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
         Left            =   4365
         TabIndex        =   7
         Top             =   375
         Width           =   1455
      End
      Begin VB.TextBox txtTran 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   330
         Width           =   1470
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
         Left            =   1800
         TabIndex        =   5
         Top             =   3600
         Width           =   4065
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
         Left            =   1785
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   2955
         Width           =   1455
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
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
         Left            =   4365
         TabIndex        =   3
         Top             =   975
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
         Left            =   3900
         TabIndex        =   15
         Top             =   450
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         Left            =   330
         TabIndex        =   14
         Top             =   405
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. Amount"
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
         TabIndex        =   13
         Top             =   3030
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GL Account No."
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
         TabIndex        =   12
         Top             =   1080
         Width           =   1395
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
         Left            =   660
         TabIndex        =   11
         Top             =   3675
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
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
         TabIndex        =   10
         Top             =   1725
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
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
         Left            =   210
         TabIndex        =   9
         Top             =   2385
         Width           =   1485
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
         Left            =   3600
         TabIndex        =   8
         Top             =   1020
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   135
      ScaleHeight     =   705
      ScaleWidth      =   6225
      TabIndex        =   0
      Top             =   180
      Width           =   6255
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSFER POSTING"
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
         Left            =   1020
         TabIndex        =   1
         Top             =   120
         Width           =   4440
      End
      Begin VB.Image Image3 
         Height          =   735
         Left            =   0
         Picture         =   "frmGLPosting.frx":0474
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6255
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6600
      Top             =   6015
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
Attribute VB_Name = "frmGLPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub clearTextboxes()
        txtAccount.Text = ""
        txtName.Text = ""
        cmbType.Text = "DEBIT"
        txtDate.Text = Date
        txtDescription.Text = ""
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
Private Sub cmbType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNet.SelStart = 0
    txtNet.SelLength = Len(txtNet.Text)
    txtNet.SetFocus
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
   If cmbType.Text = "" Then
   MsgBox "Please Select Transaction Type DEBIT/CREDIT", vbCritical
   Exit Sub
   End If
   
    Set rsU = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & txtAccount.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
If Not rsU.EOF Then
'-----------------------------------------------------------------------
        
    If cmbType.Text = "CREDIT" Then
         If rsU!Head_Type = "LIABILITY" Or rsU!Head_Type = "INCOME" Then
         
            rsU!Balance = rsU!Balance + Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = txtAccount.Text
                rsN!Name = txtName.Text
                rsN!Description = txtDescription.Text
                rsN!Dr = 0
                rsN!Cr = Val(txtNet.Text)
                rsN!Balance = rsU!Balance
                rsN.Update
                rsN.Close
            Else
    
            If rsU!Head_Type = "ASSET" Or rsU!Head_Type = "EXPENSE" Then
             
                rsU!Balance = rsU!Balance - Val(txtNet.Text)
                rsU!Date = txtDate.Text
                rsU.Update
                
                Set rsN = New ADODB.Recordset
                rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
                rsN.AddNew
                    rsN!Date = txtDate.Text
                    rsN!AC_No = txtAccount.Text
                    rsN!Name = txtName.Text
                    rsN!Description = txtDescription.Text
                    rsN!Dr = 0
                    rsN!Cr = Val(txtNet.Text)
                    rsN!Balance = rsU!Balance
                    rsN.Update
                    rsN.Close
              End If
              End If
    Else
    '--------------------------------------------------------------------------------
    
    If cmbType.Text = "DEBIT" Then
          
         If rsU!Head_Type = "LIABILITY" Or rsU!Head_Type = "INCOME" Then
         
            rsU!Balance = rsU!Balance - Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = txtAccount.Text
                rsN!Name = txtName.Text
                rsN!Description = txtDescription.Text
                rsN!Dr = Val(txtNet.Text)
                rsN!Cr = 0
                rsN!Balance = rsU!Balance
                rsN.Update
                rsN.Close
        Else
        If rsU!Head_Type = "ASSET" Or rsU!Head_Type = "EXPENSE" Then
            rsU!Balance = rsU!Balance + Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = txtAccount.Text
                rsN!Name = txtName.Text
                rsN!Description = txtDescription.Text
                rsN!Dr = Val(txtNet.Text)
                rsN!Cr = 0
                rsN!Balance = rsU!Balance
                rsN.Update
                rsN.Close
        End If
        End If
    '-------------------------------------------------------------------
    End If
End If
Else
MsgBox "Invalid Account no.", vbCritical, "Posting Error!"
rsU.Close
Exit Sub
End If
    Set rs = New ADODB.Recordset
        str = "select * from GL_Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by AC_No"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
    
    Call clearTextboxes
    txtDate.Text = Today
    Command2.Enabled = False
    txtAccount.SetFocus
    'Timer1.Enabled = True
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
        txtDate.Text = Today
        
        cmbType.Text = "DEBIT"
        
    Set rs = New ADODB.Recordset
        str = "select * from GL_Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by AC_No"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
       
        Call CmbAc
        Call CmbName
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
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    txtName.SetFocus
End If
End Sub

Private Sub txtAccount_LostFocus()
If txtAccount.Text = "" Then
Exit Sub
End If

Dim ID As String
    ID = txtAccount.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
    Call clearTextboxes
    On Error Resume Next
        txtAccount.Text = rs!AC_No
        txtDate.Text = Today
        txtName.Text = rs!Head_Name
        txtBalance.Text = rs!Balance
        rs.Close
        
        txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
        
        txtBalance.Text = Format$(Val(txtBalance.Text), "###0.00")
        Command2.Enabled = True
        
    Else
    MsgBox "There is no such Account no. found.,", vbCritical
        rs.Close
        Call clearTextboxes
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
    Command2.SetFocus
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbType.SelStart = 0
    cmbType.SelLength = Len(cmbType.Text)
    cmbType.SetFocus
End If
End Sub

Private Sub txtName_LostFocus()
If txtName.Text = "" Then
Exit Sub
End If

Dim ID As String
    ID = txtName.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where Head_Name like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
    Call clearTextboxes
    On Error Resume Next
        txtAccount.Text = rs!AC_No
        txtDate.Text = Today
        txtName.Text = rs!Head_Name
        txtBalance.Text = rs!Balance
        rs.Close
        
        txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
        
        txtBalance.Text = Format$(Val(txtBalance.Text), "###0.00")
        Command2.Enabled = True
        
    Else
    MsgBox "There is no such Account no. found.,", vbCritical
        rs.Close
        Call clearTextboxes
        txtAccount.SetFocus
        txtDate.Text = Today
    End If
    Exit Sub

End Sub

Private Sub txtNet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDescription.SetFocus
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
        str = "select * from Tran where Name like ' % search % '"
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
    txtAccount.SelStart = 0
    txtAccount.SelLength = Len(txtAccount.Text)
    txtAccount.SetFocus
End If
End Sub

Private Sub txtTran_LostFocus()
On Error Resume Next
Dim Tran As String
    Tran = txtTran.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Tran where Sl like '" & Tran & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        
        txtTran.Text = rs!sl
        txtDate.Text = rs!Date
        txtAccount.Text = rs!AC_No
        txtName.Text = rs!Name
        
        If rs!Cr > 0 Then
            txtNet.Text = rs!Cr
            cmbType.Text = "CREDIT"
        Else
            txtNet.Text = rs!Dr
            cmbType.Text = "DEBIT"
        End If
        
        txtDescription.Text = rs!Description
        txtBalance.Text = rs!Balance
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




