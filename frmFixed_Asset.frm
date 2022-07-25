VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmFixed_Asset 
   Appearance      =   0  'Flat
   BackColor       =   &H00004080&
   Caption         =   "Fixed Asset Schedule"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10125
   Icon            =   "frmFixed_Asset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3525
      Left            =   135
      ScaleHeight     =   3465
      ScaleWidth      =   5925
      TabIndex        =   9
      Top             =   960
      Width           =   5985
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Asset Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   135
         TabIndex        =   10
         Top             =   90
         Width           =   5685
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
            Left            =   1275
            TabIndex        =   29
            Text            =   "Text4"
            Top             =   360
            Width           =   1275
         End
         Begin VB.ComboBox cmbProd_Group 
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
            Left            =   1245
            TabIndex        =   27
            Text            =   "Combo1"
            Top             =   1425
            Width           =   1860
         End
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
            TabIndex        =   25
            Text            =   "Text3"
            Top             =   360
            Width           =   1275
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
            Left            =   4230
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   900
            Width           =   1275
         End
         Begin VB.TextBox txtQty 
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
            Left            =   4230
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   1418
            Width           =   1275
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
            Left            =   1260
            TabIndex        =   15
            Text            =   "Combo1"
            Top             =   1980
            Width           =   1815
         End
         Begin VB.TextBox txtProd_Price 
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
            Left            =   4230
            TabIndex        =   14
            Text            =   "Text2"
            Top             =   2520
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
            Left            =   1260
            TabIndex        =   13
            Text            =   "Combo2"
            Top             =   2527
            Width           =   1815
         End
         Begin VB.TextBox txtProd_Cost 
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
            Left            =   4230
            TabIndex        =   12
            Text            =   "Text7"
            Top             =   1980
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
            Left            =   1260
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   900
            Width           =   1275
         End
         Begin VB.Label Label11 
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
            Left            =   345
            TabIndex        =   30
            Top             =   405
            Width           =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Asset Group"
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
            TabIndex        =   28
            Top             =   1485
            Width           =   1095
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
            TabIndex        =   26
            Top             =   405
            Width           =   405
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
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
            Left            =   3825
            TabIndex        =   24
            Top             =   1485
            Width           =   315
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
            Left            =   2940
            TabIndex        =   23
            Top             =   990
            Width           =   1200
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Cost"
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
            Left            =   3255
            TabIndex        =   22
            Top             =   2587
            Width           =   885
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Cost"
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
            TabIndex        =   21
            Top             =   2040
            Width           =   810
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Asset Model"
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
            TabIndex        =   20
            Top             =   2587
            Width           =   1095
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Asset Name"
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
            TabIndex        =   19
            Top             =   2040
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Asset Code"
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
            Left            =   165
            TabIndex        =   18
            Top             =   945
            Width           =   1020
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   135
      ScaleHeight     =   600
      ScaleWidth      =   5925
      TabIndex        =   4
      Top             =   4635
      Width           =   5985
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Save"
         Height          =   435
         Left            =   255
         TabIndex        =   8
         Top             =   105
         Width           =   1020
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   435
         Left            =   1725
         TabIndex        =   7
         Top             =   105
         Width           =   1020
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   435
         Left            =   4635
         TabIndex        =   6
         Top             =   90
         Width           =   1020
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   435
         Left            =   3195
         TabIndex        =   5
         Top             =   90
         Width           =   1020
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0FFFF&
      Height          =   645
      Left            =   135
      ScaleHeight     =   585
      ScaleWidth      =   5925
      TabIndex        =   2
      Top             =   165
      Width           =   5985
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FIXED ASSET ENTRY FORM"
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
         Height          =   450
         Left            =   555
         TabIndex        =   3
         Top             =   45
         Width           =   4785
      End
      Begin VB.Image Image3 
         Height          =   600
         Left            =   0
         Picture         =   "frmFixed_Asset.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5940
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   5145
      Left            =   6255
      ScaleHeight     =   5085
      ScaleWidth      =   3645
      TabIndex        =   0
      Top             =   135
      Width           =   3705
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFixed_Asset.frx":A4D0
         Height          =   4830
         Left            =   135
         TabIndex        =   1
         Top             =   135
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   8520
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         ForeColor       =   -2147483647
         HeadLines       =   1
         RowHeight       =   19
         RowDividerStyle =   3
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
      Height          =   420
      Left            =   6705
      Top             =   4590
      Visible         =   0   'False
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   741
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
Attribute VB_Name = "frmFixed_Asset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub Prod_Tran()
Set rsN = New ADODB.Recordset
        rsN.Open "Fixed_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Prod_Group = cmbProd_Group.Text
        rsN!Prod_Name = cmbProd_Name.Text
        rsN!Prod_Model = cmbProd_Model.Text
        rsN!Prod_Code = txtProd_Sl.Text
        rsN!Purchase = txtQty.Text
        rsN!Disposal = 0
        rsN!Stock = rsU!Stock
        rsN!Date = txtDate.Text
        rsN!Prod_Price = txtProd_Cost.Text
        rsN!Amount = Val(txtQty.Text) * Val(txtProd_Cost.Text)
        rsN.Update
        rsN.Close
End Sub
Private Sub GL_Dr()
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100122 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + (Val(txtQty.Text) * Val(txtProd_Cost.Text))
        rs!Date = txtDate.Text
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Asset Code-" & txtProd_Sl.Text
        rsN!Dr = (Val(txtQty.Text) * Val(txtProd_Cost.Text))
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
            
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + (Val(txtQty.Text) * Val(txtProd_Cost.Text))
        rs!Date = txtDate.Text
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Asset Code-" & txtProd_Sl.Text
        rsN!Dr = (Val(txtQty.Text) * Val(txtProd_Cost.Text))
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
End Sub
Private Sub Prod_Group()
        cmbProd_Group.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Group FROM Fixed_Master"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbProd_Group.AddItem rs!Prod_Group
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub Prod_Name()
        cmbProd_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name FROM Fixed_Master"
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
        str = "SELECT DISTINCT Prod_Model FROM Fixed_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbProd_Model.AddItem rsN!Prod_Model
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub clearTextboxes()
        txtInvoice.Text = ""
        cmbProd_Group.Text = ""
        cmbProd_Name.Text = ""
        cmbProd_Model.Text = ""
        txtProd_Sl.Text = ""
        txtDate.Text = ""
        txtProd_Cost.Text = ""
        txtQty.Text = ""
        txtProd_Price.Text = ""
        txtStock.Text = ""
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

Private Sub cmbProd_Group_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProd_Name.SetFocus
End If
End Sub

Private Sub cmbProd_Group_LostFocus()
Dim Prod As String
Prod = cmbProd_Group.Text
        
        cmbProd_Name.Clear
        Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name From Fixed_Master where Prod_Group Like '" & Prod & "'"
        rsN.Open str, conn
        If Not rsN.EOF Then
        
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbProd_Name.AddItem rsN!Prod_Name
        rsN.MoveNext
        Loop
        rsN.Close
        Else
        rsN.Close
        End If
    
cmbProd_Model.Clear

        Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Model From Fixed_Master where Prod_Group Like '" & Prod & "'"
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
        str = "SELECT * FROM Fixed_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Group = rsN!Prod_Group
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtDate.Text = Date
            rsN.Close
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
        str = "SELECT * FROM Fixed_Master where Prod_Name Like '" & cmbProd_Name.Text & "' AND Prod_Model Like '" & cmbProd_Model.Text & "' order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Group = rsN!Prod_Group
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            rsN.Close
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
        str = "SELECT DISTINCT Prod_Model FROM Fixed_Master where Prod_Name Like '" & Prod & "'"
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
        str = "SELECT * FROM Fixed_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Group = rsN!Prod_Group
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtDate.Text = Date
            rsN.Close
        Else
        Exit Sub
        End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdReceive_Click()
Dim Prod As String
Prod = txtProd_Sl.Text

Set rsU = New ADODB.Recordset
        str = "select * from Fixed_Master where Prod_Code like '" & Prod & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
If Not rsU.EOF Then
'Cash Transaction open---------------------------------------------------
     
        rsU!Purchase = rsU!Purchase + Val(txtQty.Text)
        rsU!Stock = rsU!Stock + Val(txtQty.Text)
        rsU!Prod_Price = Val(txtProd_Cost.Text)
        rsU!Amount = rsU!Stock * rsU!Prod_Price
        rsU.Update
        
    Call Prod_Tran
    
        rsU.Close
    
    Call GL_Dr
    
Else
rsU.Close
    
    'Product Table-----------------------------------------------------------
        Set rsN = New ADODB.Recordset
        rsN.Open "Fixed_Master", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Prod_Code = txtProd_Sl.Text
        rsN!Prod_Group = cmbProd_Group.Text
        rsN!Prod_Name = cmbProd_Name.Text
        rsN!Prod_Model = cmbProd_Model.Text
        rsN!Open_Bal = 0
        rsN!Purchase = Val(txtQty.Text)
        rsN!Disposal = 0
        rsN!Stock = Val(txtQty.Text)
        rsN!Prod_Price = Val(txtProd_Cost.Text)
        rsN!Amount = rsN!Stock * rsN!Prod_Price
        rsN.Update
        rsN.Close
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Fixed_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Prod_Name = cmbProd_Name.Text
        rsN!Prod_Model = cmbProd_Model.Text
        rsN!Prod_Group = cmbProd_Group.Text
        rsN!Prod_Code = txtProd_Sl.Text
        rsN!Purchase = txtQty.Text
        rsN!Disposal = 0
        rsN!Stock = txtQty.Text
        rsN!Date = txtDate.Text
        rsN!Prod_Price = txtProd_Cost.Text
        rsN!Amount = Val(txtQty.Text) * Val(txtProd_Cost.Text)
        
        rsN.Update
        rsN.Close
    
    Call GL_Dr
    
    
End If

Set rs = New ADODB.Recordset
        str = "select * from Fixed_Tran order by Sl"
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

Private Sub Form_Load()
On Error Resume Next
Set rsU = New ADODB.Recordset
        str = "select * from Fixed_Tran order by Sl"
        rsU.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rsU.Close
   
    Call Prod_Group
    Call Prod_Name
    Call Prod_Model
    Call clearTextboxes
        
        
        
        txtDate.Text = Date
        cmdReceive.Enabled = False
        cmdUpdate.Enabled = False
    Exit Sub
End Sub
Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbBank.SetFocus
End If
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtInvoice.Text = "" Then
Exit Sub
Else
cmdReceive.Enabled = True
txtProd_Sl.SetFocus
End If
End If
End Sub
Private Sub txtProd_Cost_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Price.SelStart = 0
txtProd_Price.SelLength = Len(txtProd_Price.Text)
txtProd_Price.SetFocus
End If
End Sub

Private Sub txtProd_Cost_LostFocus()
txtProd_Price.Text = Format$(Val(txtQty.Text) * Val(txtProd_Cost.Text), "###0.00")
txtProd_Cost.Text = Format$(Val(txtProd_Cost.Text), "###0.00")
End Sub
Private Sub txtProd_Price_LostFocus()
txtProd_Price.Text = Format$(Val(txtProd_Price.Text), "###0.00")
End Sub

Private Sub txtProd_Sl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProd_Group.SetFocus
End If
End Sub

Private Sub txtProd_Sl_LostFocus()
On Error Resume Next
Dim mid As Integer
    Dim Prod As String
    Prod = txtProd_Sl.Text
    mid = 0
    
    
    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Fixed_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Group.Text = rsN!Prod_Group
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtDate.Text = Date
            txtProd_Cost.Text = Format$(Val(rsN!Prod_Price), "###0.00")
            rsN.Close
            cmdUpdate.Enabled = True
            cmdReceive.Enabled = True
            'cmdPrint.Enabled = True
        
        
        Else
        If MsgBox("Do you want add new Asset?", vbInformation + vbYesNo, "Add New") = vbYes Then
        
        Set rsU = New ADODB.Recordset
            str = "select * from Fixed_Master order by Prod_Code"
            rsU.Open str, conn
        
        If Not rsU.EOF Then
           rsU.MoveFirst
        
        Do While Not rsU.EOF = True
            mid = Val(rsU!Prod_Code)
            rsU.MoveNext
        Loop
            rsU.Close
            mid = mid + 1
        Else
        rsU.Close
        mid = 101
        End If
        Call clearTextboxes
            txtDate.Text = Date
            txtProd_Sl.Text = mid
            txtStock.Text = "0.00"
            txtProd_Cost.Text = "0.00"
            txtProd_Price.Text = "0.00"
            cmdReceive.Enabled = True
        Else
            Call clearTextboxes
            cmbProd_Name.SetFocus
        End If
    End If
    Exit Sub

End Sub
Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Cost.SelStart = 0
txtProd_Cost.SelLength = Len(txtProd_Cost.Text)
txtProd_Cost.SetFocus
End If
End Sub

