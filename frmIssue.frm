VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "Msdatgrd.ocx"
Begin VB.Form frmIssue 
   BackColor       =   &H80000001&
   Caption         =   "Issue Form"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   Icon            =   "frmIssue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      Height          =   3345
      Left            =   315
      ScaleHeight     =   3285
      ScaleWidth      =   7050
      TabIndex        =   12
      Top             =   1035
      Width           =   7110
      Begin VB.ComboBox txtStation 
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
         Left            =   5085
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1575
         Width           =   1725
      End
      Begin VB.TextBox txtbom 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1500
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtRdate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5085
         TabIndex        =   20
         Text            =   "Text3"
         Top             =   375
         Width           =   1695
      End
      Begin VB.TextBox txtReq 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1485
         TabIndex        =   19
         Text            =   "Text4"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtRqty 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5085
         TabIndex        =   18
         Text            =   "Text7"
         Top             =   990
         Width           =   1695
      End
      Begin VB.ComboBox txtModel 
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
         Left            =   1485
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   1575
         Width           =   2130
      End
      Begin VB.ComboBox txtParts 
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
         Left            =   1485
         TabIndex        =   16
         Text            =   "Combo2"
         Top             =   2115
         Width           =   2130
      End
      Begin VB.TextBox txtStock 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1485
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2655
         Width           =   1695
      End
      Begin VB.TextBox txtPrice 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5085
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   2115
         Width           =   1695
      End
      Begin VB.TextBox txtAmount 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5085
         TabIndex        =   13
         Text            =   "Text3"
         Top             =   2655
         Width           =   1695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Station"
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
         TabIndex        =   10
         Top             =   1620
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BOM Code"
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
         TabIndex        =   11
         Top             =   465
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Date"
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
         Left            =   3990
         TabIndex        =   29
         Top             =   495
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Req No."
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
         Left            =   585
         TabIndex        =   28
         Top             =   1035
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Name"
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
         Left            =   195
         TabIndex        =   27
         Top             =   1635
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parts Name"
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
         Left            =   255
         TabIndex        =   26
         Top             =   2175
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Qty"
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
         Left            =   4080
         TabIndex        =   25
         Top             =   1095
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Stock"
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
         TabIndex        =   24
         Top             =   2745
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
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
         Left            =   4065
         TabIndex        =   23
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Label10 
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
         Left            =   4245
         TabIndex        =   22
         Top             =   2745
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000013&
      Height          =   5055
      Left            =   7605
      ScaleHeight     =   4995
      ScaleWidth      =   3150
      TabIndex        =   7
      Top             =   225
      Width           =   3210
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmIssue.frx":0442
         Height          =   4740
         Left            =   135
         TabIndex        =   8
         Top             =   135
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   8361
         _Version        =   393216
         BackColor       =   -2147483629
         HeadLines       =   1
         RowHeight       =   19
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   7770
      Top             =   4995
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   767
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
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000013&
      Height          =   645
      Left            =   315
      ScaleHeight     =   585
      ScaleWidth      =   7050
      TabIndex        =   5
      Top             =   210
      Width           =   7110
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT ISSUE FORM"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   105
         TabIndex        =   6
         Top             =   0
         Width           =   6945
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      Height          =   750
      Left            =   315
      ScaleHeight     =   690
      ScaleWidth      =   7050
      TabIndex        =   0
      Top             =   4515
      Width           =   7110
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   435
         Left            =   4320
         TabIndex        =   30
         Top             =   135
         Width           =   1155
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Issue"
         Height          =   435
         Left            =   165
         TabIndex        =   4
         Top             =   135
         Width           =   1155
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   435
         Left            =   1545
         TabIndex        =   3
         Top             =   135
         Width           =   1155
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   435
         Left            =   2925
         TabIndex        =   2
         Top             =   135
         Width           =   1155
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   435
         Left            =   5700
         TabIndex        =   1
         Top             =   135
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String

Private Sub clearTextboxes()
        txtBom.Text = ""
        txtReq.Text = ""
        txtModel.Text = ""
        txtParts.Text = ""
        txtStock.Text = ""
        txtRdate.Text = ""
        txtRqty.Text = ""
        txtPrice.Text = ""
        txtAmount.Text = ""
        txtStation.Text = ""
End Sub


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
 On Error Resume Next
    If txtReq.Text = "" Then
        MsgBox "Please Enter a Req. no. to be delete", vbCritical + vbOKOnly
        txtReq.SetFocus
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
        str = "select * from Issue where Req_no like '" & txtReq.Text & "'"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
    On Error Resume Next
        txtBom.Text = rs!Bom_code
        txtRdate.Text = rsN!Out_date
        txtReq.Text = rsN!Req_no
        txtModel.Text = rs!Model
        txtParts.Text = rs!Parts
        txtRqty.Text = rs!Qty
        txtStation.Text = rs!Station
        txtPrice.Text = rs!Unit_Price
        txtAmount.Text = rs!Amount
        rs.Close
     
     If MsgBox("Really want to delete?", vbCritical + vbYesNo) = vbYes Then
        
        str = "delete from Issue where Req_no like '" & txtReq.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.Close
    
    Set rsU = New ADODB.Recordset
        str = "select * from Stock where Bom_Code like '" & txtBom.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
        On Error Resume Next
        rsU!issue = rsU!issue - Val(txtRqty.Text)
        rsU!Stock = rsU!Opening_Bal + rsU!Receive - rsU!issue
        rsU!Unit_Price = Val(txtPrice.Text)
        rsU!Amount = Val(rsU!Stock) * Val(rsU!Unit_Price)
        rsU.Update
        rsU.Close
    
    Set rsN = New ADODB.Recordset
        str = "select * from Issue order by Out_date"
        rsN.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rsN.Close
    
    End If
    End If
    
    Else
        MsgBox "There is no such Req. no. found to be delete", vbCritical + vbOKOnly
        rs.Close
    End If
    Call clearTextboxes
    txtRdate.Text = Date
Exit Sub
Resume Next
End Sub
Private Sub cmdReceive_Click()

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
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Issue", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Bom_code = txtBom.Text
        rsN!Out_date = txtRdate.Text
        rsN!Req_no = txtReq.Text
        rsN!Model_name = txtModel.Text
        rsN!Parts_name = txtParts.Text
        rsN!Qty = txtRqty.Text
        rsN!Station = txtStation.Text
        rsN!Unit_Price = Val(txtPrice.Text)
        rsN!Amount = Val(txtRqty.Text) * Val(txtPrice.Text)
        rsN.Update
        rsN.Close

    Set rs = New ADODB.Recordset
        str = "select * from Issue order by Out_date"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    
Else
    
    MsgBox "There is no such product found in the stock. you can not issue without receive", vbCritical + vbOKOnly
    
    End If
    Call clearTextboxes
    txtRdate.Text = Date
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

Private Sub Command1_Click()
frmIssuesearch.Show 1
End Sub

Private Sub Form_Load()
 On Error Resume Next
Set rsU = New ADODB.Recordset
        str = "select * from Issue order by Out_date"
        rsU.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rsU.Close
    
    Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Parts FROM Stock"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtParts.AddItem rs!Parts
        rs.MoveNext
        Loop
        rs.Close
 
    Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Model FROM Stock"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        txtModel.AddItem rsN!Model
        rsN.MoveNext
        Loop
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Station FROM Issue"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtStation.AddItem rs!Station
        rs.MoveNext
        Loop
        rs.Close
    
    Call clearTextboxes
        txtRdate.Text = Date
        cmdReceive.Enabled = False
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
    Exit Sub
End Sub

Private Sub Picture5_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtBom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtReq.SetFocus
End If
End Sub

Private Sub txtbom_LostFocus()
If txtBom.Text = "" Then
Exit Sub
End If

Dim Bom As String
    Bom = txtBom.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Stock where Bom_Code like '" & Bom & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        txtBom.Text = rs!Bom_code
        txtModel.Text = rs!Model
        txtParts.Text = rs!Parts
        txtStock.Text = rs!Stock
        txtPrice.Text = rs!Unit_Price
        txtAmount.Text = rs!Amount
        rs.Close
            
        cmdReceive.Enabled = True
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = True
    Else
    If MsgBox("There is no such BOM Code found, Do you want add new Product?", vbCritical + vbYesNo) = vbYes Then
        'Call clearTextboxes
        txtRdate.Text = Date
        cmdReceive.Enabled = True
        rs.Close
    Else
        Call clearTextboxes
    End If
    End If
    Exit Sub

End Sub

Private Sub txtModel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtParts.SetFocus
End If
End Sub

Private Sub txtParts_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtStock.SetFocus
End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAmount.Text = Val(txtRqty.Text) * Val(txtPrice.Text)
txtAmount.SetFocus
End If
End Sub

Private Sub txtRdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtRqty.SetFocus
End If
End Sub

Private Sub txtReq_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtModel.SetFocus
End If
End Sub

Private Sub txtReq_LostFocus()
If txtReq.Text = "" Then
    Exit Sub
    End If

Dim Req As String
    Req = txtReq.Text
    
    Set rsN = New ADODB.Recordset
        str = "select * from Issue where Req_no like '" & Req & "'"
        rsN.Open str, conn
    
    If Not rsN.EOF Then
    On Error Resume Next
        MsgBox "This Req.No. is already exist, Please enter new Req. no.", vbCritical + vbOKOnly
        txtBom.Text = rsN!Bom_code
        txtRdate.Text = rsN!Out_date
        txtReq.Text = rsN!Req_no
        txtModel.Text = rsN!Model_name
        txtParts.Text = rsN!Parts_name
        txtRqty.Text = rsN!Qty
        txtStation.Text = rsN!Station
        txtPrice.Text = rsN!Unit_Price
        txtAmount.Text = rsN!Amount
        rs.Close
    Else
       txtModel.SetFocus
    End If
    Exit Sub
End Sub

Private Sub txtRqty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAmount.Text = Val(txtRqty.Text) * Val(txtPrice.Text)
txtStation.SetFocus
End If
End Sub

Private Sub txtStation_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPrice.SetFocus
End If
End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtRdate.SetFocus
End If
End Sub
