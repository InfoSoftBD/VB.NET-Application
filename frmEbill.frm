VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmEbill 
   BackColor       =   &H00004000&
   Caption         =   "PDB Bill Collection"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   Icon            =   "frmEbill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6750
      Top             =   6870
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
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0FFC0&
      Height          =   6945
      Left            =   6720
      ScaleHeight     =   6885
      ScaleWidth      =   4275
      TabIndex        =   27
      Top             =   240
      Width           =   4335
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmEbill.frx":0442
         Height          =   6045
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   10663
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         HeadLines       =   1
         RowHeight       =   19
         RowDividerStyle =   4
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         TabIndex        =   13
         ToolTipText     =   "Enter Consumer no."
         Top             =   120
         Width           =   1815
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
         Left            =   360
         TabIndex        =   33
         Top             =   240
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   6195
      TabIndex        =   25
      Top             =   6450
      Width           =   6255
      Begin VB.CommandButton Command5 
         Caption         =   "&Update"
         Height          =   435
         Left            =   2580
         TabIndex        =   10
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Print"
         Height          =   435
         Left            =   3840
         TabIndex        =   11
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete"
         Height          =   435
         Left            =   1350
         TabIndex        =   9
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   435
         Left            =   120
         TabIndex        =   8
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   435
         Left            =   5070
         TabIndex        =   12
         Top             =   150
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   5025
      Left            =   240
      ScaleHeight     =   4965
      ScaleWidth      =   6195
      TabIndex        =   17
      Top             =   1200
      Width           =   6255
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frmEbill.frx":0457
         Left            =   2550
         List            =   "frmEbill.frx":0464
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   840
         Width           =   1065
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
         Left            =   4440
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0FFC0&
         Height          =   1095
         Left            =   4440
         ScaleHeight     =   1035
         ScaleWidth      =   1395
         TabIndex        =   29
         Top             =   1350
         Width           =   1455
      End
      Begin VB.TextBox txtTran 
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
         Left            =   2565
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Revenue Stamp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4440
         Picture         =   "frmEbill.frx":0477
         TabIndex        =   7
         Top             =   2550
         Width           =   1695
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
         Left            =   2565
         TabIndex        =   1
         Top             =   1380
         Width           =   1455
      End
      Begin VB.TextBox txtAccbill 
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
         Left            =   2565
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1980
         Width           =   1455
      End
      Begin VB.TextBox txtSer 
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
         Left            =   2565
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   2580
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
         Left            =   2565
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   3180
         Width           =   1455
      End
      Begin VB.TextBox txtVat 
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
         Left            =   2565
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   3780
         Width           =   1455
      End
      Begin VB.TextBox txtNetbill 
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
         Left            =   2565
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   4380
         Width           =   1455
      End
      Begin VB.Label lblTransaction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Transaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3705
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Label Label10 
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
         Height          =   375
         Left            =   765
         TabIndex        =   31
         Top             =   930
         Width           =   1485
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
         Left            =   3840
         TabIndex        =   30
         Top             =   360
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   4440
         Picture         =   "frmEbill.frx":DF4C
         Stretch         =   -1  'True
         Top             =   1350
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   885
         TabIndex        =   26
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Bill Amount"
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
         Left            =   765
         TabIndex        =   23
         Top             =   4500
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
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
         Left            =   1875
         TabIndex        =   22
         Top             =   3900
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Others"
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
         Left            =   1665
         TabIndex        =   21
         Top             =   3300
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ser Charge"
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
         Left            =   1275
         TabIndex        =   20
         Top             =   2700
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Bill Amount"
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
         TabIndex        =   19
         Top             =   2100
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Castomer Account No."
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
         Left            =   270
         TabIndex        =   18
         Top             =   1500
         Width           =   1980
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   6195
      TabIndex        =   16
      Top             =   240
      Width           =   6255
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PDB BILL POSTING"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   24
         Top             =   90
         Width           =   3780
      End
   End
End
Attribute VB_Name = "frmEbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub clearTextboxes()
        txtTran.Text = ""
        txtAccount.Text = ""
        txtAccbill.Text = "0.00"
        txtSer.Text = "0.00"
        txtOther.Text = "0.00"
        txtVat.Text = "0.00"
        txtNetbill.Text = "0.00"
        Check1.Value = 0
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
Dim Today As Date
Today = txtDate.Text
On Error Resume Next
    Set rsN = New ADODB.Recordset
        str = "select * from Ebill where Tr_no='" & txtTran.Text & "'"
        rsN.Open str, conn
    If rsN.EOF Then
        rsN.Close
        rsN.Open "Ebill", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Tr_no = txtTran.Text
        rsN!Tr_type = Combo1.Text
        rsN!Account_no = txtAccount.Text
        rsN!Ac_Bill = Val(txtAccbill.Text)
        rsN!Ser_charge = Val(txtSer.Text)
        rsN!Others = Val(txtOther.Text)
        rsN!vat = Val(txtVat.Text)
        rsN!Net_bill = Val(txtNetbill.Text)
        rsN!Revenue = Check1.Value
        If Check1.Value = 1 Then
        rsN!Stamp = 5
        End If
        rsN!Date = txtDate.Text
        rsN.Update
        rsN.Close
    Call clearTextboxes
    Call ColumnWidth
    Else
        rsN.Close
    End If

    Set rs = New ADODB.Recordset
        str = "select * from Ebill where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') order by Tr_no Desc"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Call ColumnWidth
    
    Combo1.SetFocus
    Command3.Enabled = False
    Command5.Enabled = False
Exit Sub
   Resume Next
End Sub
Private Sub Command3_Click()
On Error Resume Next
 Dim Today As Date
 Today = txtDate.Text
 str = "select * from Ebill where tr_no like '" & txtTran.Text & "'"
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
        
        str = "delete from Ebill where Tr_no like '" & txtTran.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.Close
    Call clearTextboxes
        
        str = "select * from Ebill where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by Tr_no Desc"
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
frmPrint.Show 1
End Sub
Private Sub Command5_Click()
On Error Resume Next
Dim Today As Date
Today = txtDate.Text

Set rsU = New ADODB.Recordset
    str = "select * from Ebill where Tr_no like '" & txtTran.Text & "'"
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
        str = "select * from Ebill where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by Tr_no Desc"
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
Dim Today As Date
Today = Date
On Error GoTo Last
        txtDate.Text = Today
        Combo1.Text = 101
        lblTransaction.Visible = True
        lblTransaction.Caption = "Cash Transaction"
        
    Set rs = New ADODB.Recordset
        str = "select * from Ebill where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by Tr_no Desc"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
        Command3.Enabled = False
        Command5.Enabled = False
        Call ColumnWidth
    Exit Sub
Last:
    MsgBox ("Database Connection error: " + Err.Description)
End Sub
Private Sub txtAccbill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSer.SelStart = 0
    txtSer.SelLength = Len(txtSer.Text)
    txtSer.SetFocus
End If
End Sub
Private Sub txtAccbill_LostFocus()
On Error Resume Next
Dim a, b, c, d As Double
    a = Val(txtAccbill.Text)
    b = Val(txtSer.Text)
    c = Val(txtOther.Text)
txtVat.Text = Format$(Val(a + b + c) * 0.05, "###0.00")
    d = Val(txtVat.Text)
txtNetbill.Text = Format$(Round(a + b + c + d), "###0.00")
txtAccbill.Text = Format$(Val(txtAccbill.Text), "###0.00")
txtSer.Text = Format$(Val(txtSer.Text), "###0.00")
txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
    txtSer.SelStart = 0
    txtSer.SelLength = Len(txtSer.Text)
End Sub
Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtAccbill.SelStart = 0
    txtAccbill.SelLength = Len(txtAccbill.Text)
    txtAccbill.SetFocus
End If
End Sub
Private Sub txtNetbill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check1.SetFocus
End If
End Sub
Private Sub txtNetbill_LostFocus()
On Error Resume Next
Dim Net As Double
    txtAccbill.Text = Format$(Val(txtAccbill.Text), "###0.00")
    txtSer.Text = Format$(Val(txtSer.Text), "###0.00")
    txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
    txtVat.Text = Format$(Val(txtVat.Text), "###0.00")
    txtNetbill.Text = Format$(Val(txtNetbill.Text), "###0.00")
   
    Net = Val(txtNetbill.Text)
If Net >= 500 Then
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
Private Sub txtOther_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtVat.SelStart = 0
    txtVat.SelLength = Len(txtVat.Text)
    txtVat.SetFocus
End If
End Sub
Private Sub txtOther_LostFocus()
On Error Resume Next
Dim a, b, c, d As Integer
    a = Val(txtAccbill.Text)
    b = Val(txtSer.Text)
    c = Val(txtOther.Text)
txtVat.Text = Format$(Val(a + b + c) * 0.05, "###0.00")
    d = Val(txtVat.Text)
txtNetbill.Text = Format$(Round(a + b + c + d), "###0.00")
txtAccbill.Text = Format$(Val(txtAccbill.Text), "###0.00")
txtSer.Text = Format$(Val(txtSer.Text), "###0.00")
txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
    txtVat.SelStart = 0
    txtVat.SelLength = Len(txtVat.Text)
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If txtSearch.Text = "" Then
Exit Sub
End If

    Dim search As String
        search = txtSearch.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Ebill where Account_no like '" & search & "'"
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
        Dim Today As Date
        Today = Date
        On Error GoTo Last
        txtDate.Text = Today
               
    Set rs = New ADODB.Recordset
        str = "select * from Ebill where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') order by Tr_no Desc"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
               
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
txtVat.Text = Format$(Val(a + b + c) * 0.05, "###0.00")
    d = Val(txtVat.Text)
txtNetbill.Text = Format$(Round(a + b + c + d), "###0.00")
txtAccbill.Text = Format$(Val(txtAccbill.Text), "###0.00")
txtSer.Text = Format$(Val(txtSer.Text), "###0.00")
txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
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
        str = "select * from Ebill where Tr_no like '" & Tran & "'"
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
Last:
    MsgBox ("Database Connection error: " + Err.Description)
End Sub
Private Sub txtVat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNetbill.SelStart = 0
    txtNetbill.SelLength = Len(txtNetbill.Text)
    txtNetbill.SetFocus
End If
End Sub
Private Sub txtVat_LostFocus()
On Error Resume Next
Dim a, b, c, d, e As Double
    a = Val(txtAccbill.Text)
    b = Val(txtSer.Text)
    c = Val(txtOther.Text)

txtVat.Text = Format$(Val(txtVat.Text), "###0.00")
    d = Val(txtVat.Text)
    e = (Val(d) / 0.05)
    
    txtNetbill.Text = Format$(Round(d + e), "###0.00")
    txtAccbill.Text = Format$(Val(txtAccbill.Text), "###0.00")
    txtSer.Text = Format$(Val(txtSer.Text), "###0.00")
    txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
    txtNetbill.SelStart = 0
    txtNetbill.SelLength = Len(txtNetbill.Text)
End Sub
