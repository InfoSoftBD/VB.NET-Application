VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmIssue 
   BackColor       =   &H00008000&
   Caption         =   "Issue Form"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   Icon            =   "frmProdIssue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   2985
      Left            =   135
      ScaleHeight     =   2925
      ScaleWidth      =   6060
      TabIndex        =   11
      Top             =   900
      Width           =   6120
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
         Left            =   4185
         TabIndex        =   27
         Text            =   "Text3"
         Top             =   180
         Width           =   1695
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
         Left            =   1395
         TabIndex        =   26
         Text            =   "Text4"
         Top             =   158
         Width           =   1275
      End
      Begin VB.TextBox txtProd_Sl 
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
         Left            =   1395
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   720
         Width           =   1275
      End
      Begin VB.ComboBox cmbProd_Model 
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
         Left            =   1395
         TabIndex        =   20
         Text            =   "Combo2"
         Top             =   1800
         Width           =   1725
      End
      Begin VB.ComboBox cmbProd_Name 
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
         Left            =   1395
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   1260
         Width           =   1725
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
         Height          =   375
         Left            =   1395
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2295
         Width           =   1275
      End
      Begin VB.ComboBox txtBranch 
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
         Left            =   4185
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   720
         Width           =   1725
      End
      Begin VB.TextBox txtQty 
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
         Left            =   4185
         TabIndex        =   14
         Text            =   "Text7"
         Top             =   1260
         Width           =   1695
      End
      Begin VB.TextBox txtSale_Price 
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
         Left            =   4185
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   1800
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
         Height          =   375
         Left            =   4185
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   2340
         Width           =   1695
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
         Left            =   3135
         TabIndex        =   29
         Top             =   225
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invice No."
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
         Left            =   465
         TabIndex        =   28
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
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
         Left            =   120
         TabIndex        =   25
         Top             =   765
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
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
         Left            =   60
         TabIndex        =   24
         Top             =   1305
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Model"
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
         Left            =   45
         TabIndex        =   23
         Top             =   1845
         Width           =   1260
      End
      Begin VB.Label Label10 
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
         Index           =   1
         Left            =   105
         TabIndex        =   22
         Top             =   2340
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Branch"
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
         Left            =   2925
         TabIndex        =   10
         Top             =   765
         Width           =   1140
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
         Left            =   3225
         TabIndex        =   17
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Price"
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
         Left            =   3165
         TabIndex        =   16
         Top             =   1845
         Width           =   900
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
         Index           =   0
         Left            =   3390
         TabIndex        =   15
         Top             =   2430
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0FFC0&
      Height          =   4695
      Left            =   6435
      ScaleHeight     =   4635
      ScaleWidth      =   3150
      TabIndex        =   7
      Top             =   135
      Width           =   3210
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4425
         Left            =   135
         TabIndex        =   8
         Top             =   135
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   7805
         _Version        =   393216
         BackColor       =   12648384
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
      Left            =   6645
      Top             =   4500
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
      BackColor       =   &H00C0FFC0&
      Height          =   645
      Left            =   135
      ScaleHeight     =   585
      ScaleWidth      =   6060
      TabIndex        =   5
      Top             =   120
      Width           =   6120
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT ISSUE "
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   330
         TabIndex        =   6
         Top             =   0
         Width           =   5685
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   750
      Left            =   135
      ScaleHeight     =   690
      ScaleWidth      =   6060
      TabIndex        =   0
      Top             =   4065
      Width           =   6120
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
         Left            =   1680
         TabIndex        =   3
         Top             =   135
         Width           =   1155
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   435
         Left            =   3195
         TabIndex        =   2
         Top             =   135
         Width           =   1155
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   435
         Left            =   4665
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
Dim Today As Date
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub Prod_Name()
        cmbProd_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name FROM Prod_Master"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbProd_Name.AddItem rs!Prod_Name
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub Prod_Model()
        cmbProd_Model.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Model FROM Prod_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbProd_Model.AddItem rsN!Prod_Model
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub Branch()
        txtBranch.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Branch_Name FROM Branch"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        txtBranch.AddItem rsN!Branch_Name
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub Br_Tran()
Dim Branch As String
Bran = txtBranch.Text

Set rsU = New ADODB.Recordset
        str = "select * from Branch where Branch_Name like '" & Bran & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU!Dr = rsU!Dr + (Val(txtQty.Text) * Val(txtSale_Price.Text))
        rsU!Balance = rsU!Balance + (Val(txtQty.Text) * Val(txtSale_Price.Text))
        rsU!Date = txtDate.Text
        rsU.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Branch_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!Branch_Code = rsU!Branch_Code
        rsN!Branch_Name = rsU!Branch_Name
        rsN!Description = "Product Issue-" & txtProd_Sl.Text
        rsN!Dr = (Val(txtQty.Text) * Val(txtSale_Price.Text))
        rsN!Cr = 0
        rsN!Balance = rsU!Balance
        rsN.Update
        rsN.Close
        rsU.Close
End Sub
Private Sub GL_Tran()

Dim Prod As String
Prod = txtProd_Sl.Text

Set rsU = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Code like '" & Prod & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic

                
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100114 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - (Val(txtQty.Text) * Val(rsU!Prod_Price))
        rs!Date = txtDate.Text
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Product Code-" & txtProd_Sl.Text
        rsN!Dr = 0
        rsN!Cr = (Val(txtQty.Text) * Val(rsU!Prod_Price))
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
       

Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100115 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + ((Val(txtQty.Text) * Val(txtSale_Price.Text)) - (Val(txtQty.Text) * Val(rsU!Prod_Price)))
        rs!Date = txtDate.Text
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Product Code-" & txtProd_Sl.Text
        rsN!Dr = 0
        rsN!Cr = ((Val(txtQty.Text) * Val(txtSale_Price.Text)) - (Val(txtQty.Text) * Val(rsU!Prod_Price)))
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

            
'---------------------------------------------------------------------
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100102 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + (Val(txtQty.Text) * Val(txtSale_Price.Text))
        rs!Date = txtDate.Text
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Issue to-" & txtBranch.Text
        rsN!Dr = (Val(txtQty.Text) * Val(txtSale_Price.Text))
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
End Sub

Private Sub clearTextboxes()
        txtInvoice.Text = ""
        cmbProd_Name.Text = ""
        cmbProd_Model.Text = ""
        txtProd_Sl.Text = ""
        txtDate.Text = ""
        txtBranch.Text = ""
        txtQty.Text = ""
        txtAmount.Text = ""
        txtSale_Price.Text = ""
        txtStock.Text = ""
End Sub


Private Sub cmbProd_Model_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtBranch.SetFocus
End If
End Sub

Private Sub cmbProd_Model_LostFocus()
Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Name Like '" & cmbProd_Name.Text & "' AND Prod_Model Like '" & cmbProd_Model.Text & "' order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtSale_Price.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            rsN.Close
            cmdReceive.Enabled = True
        Else
            Exit Sub
        End If
End Sub

Private Sub cmbProd_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProd_Model.SetFocus
End If
End Sub

Private Sub cmbProd_Name_LostFocus()
Dim Prod As String
Prod = cmbProd_Name.Text
        
        cmbProd_Model.Clear
        Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Model FROM Prod_Master where Prod_Name Like '" & Prod & "'"
        rsN.Open str, conn
        If Not rsN.EOF Then
        
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbProd_Model.AddItem rsN!Prod_Model
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
            cmbProd_Model.Text = rsN!Prod_Model
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtDate.Text = Today
            txtSale_Price.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            rsN.Close
        Else
        Exit Sub
        End If

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
If txtInvoice.Text = "" Then
MsgBox "Please Input Invoice No!", vbCritical
txtInvoice.SetFocus
Exit Sub
End If

Dim Prod As String
Prod = txtProd_Sl.Text

Set rs = New ADODB.Recordset
        str = "select * from Branch where Branch_Name like '" & txtBranch.Text & "'"
        rs.Open str, conn

Set rsU = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Code like '" & Prod & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
If Not rsU.EOF Then
        rsU!Sale = rsU!Sale + Val(txtQty.Text)
        rsU!Stock = rsU!Stock - Val(txtQty.Text)
        rsU!Amount = rsU!Stock * rsU!Prod_Price
        rsU.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Prod_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Invo_No = txtInvoice.Text
        rsN!Prod_Name = cmbProd_Name.Text
        rsN!Prod_Model = cmbProd_Model.Text
        rsN!Prod_Code = txtProd_Sl.Text
        rsN!Purchase = 0
        rsN!Return = 0
        rsN!Sale = 0
        rsN!issue = txtQty.Text
        rsN!Stock = rsU!Stock
        rsN!Amount = Val(txtQty.Text) * Val(txtSale_Price.Text)
        rsN!Date = txtDate.Text
        rsN!Sale_Price = txtSale_Price.Text
        rsN!Branch_Code = rs!Branch_Code
        rsN!Branch_Name = rs!Branch_Name
        
        rsN.Update
        rsN.Close
    
        rsU.Close
        rs.Close
        
        Call GL_Tran
        Call Br_Tran
End If

Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran order by Sl"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
        
        
    Call Prod_Name
    Call Prod_Model
    
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

Private Sub Command1_Click()
frmIssuesearch.Show 1
End Sub

Private Sub Form_Load()
 On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Prod_Tran order by Sl"
        rsU.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rsU.Close
   
    Call Prod_Name
    Call Prod_Model
    Call Branch
    Call clearTextboxes
              
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




Private Sub txtBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtQty.SetFocus
End If
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtInvoice.Text = "" Then
Exit Sub
Else
txtProd_Sl.SetFocus
End If
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAmount.Text = Val(txtQty.Text) * Val(txtSale_Price.Text)
txtSale_Price.SetFocus
End If
End Sub

Private Sub txtProd_Sl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProd_Name.SetFocus
End If
End Sub

Private Sub txtProd_Sl_LostFocus()
Dim mid As Integer
    Dim Prod As String
    Prod = txtProd_Sl.Text
    mid = 0
    
    
    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtDate.Text = Today
            txtSale_Price.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            rsN.Close
            cmdReceive.Enabled = True
            cmdUpdate.Enabled = False
            cmdDelete.Enabled = False
        Else
        MsgBox "Invalid Product Code!", vbCritical, "Error"
        Call clearTextboxes
            txtDate.Text = Today
            txtProd_Sl.Text = mid
            txtStock.Text = "0.00"
            txtSale_Price.Text = "0.00"
            cmdReceive.Enabled = False
            cmbProd_Name.SetFocus
        End If
        Exit Sub
End Sub

Private Sub txtSale_Price_Change()
If KeyAscii = 13 Then
cmdReceive.SetFocus
End If
End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDate.SetFocus
End If
End Sub
