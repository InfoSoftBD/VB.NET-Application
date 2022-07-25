VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmGD_Rtn 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return from Godown"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11370
   Icon            =   "frmGD_Rtn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   105
      ScaleHeight     =   660
      ScaleWidth      =   11085
      TabIndex        =   40
      Top             =   90
      Width           =   11115
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT RETURN FROM GODOWN"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   105
         TabIndex        =   41
         Top             =   45
         Width           =   10845
      End
      Begin VB.Image Image3 
         Height          =   2085
         Left            =   0
         Picture         =   "frmGD_Rtn.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11100
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   90
      ScaleHeight     =   720
      ScaleWidth      =   11085
      TabIndex        =   35
      Top             =   8295
      Width           =   11115
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   525
         Left            =   6075
         TabIndex        =   39
         Top             =   90
         Width           =   1830
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   525
         Left            =   9045
         TabIndex        =   38
         Top             =   90
         Width           =   1830
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   525
         Left            =   3120
         TabIndex        =   37
         Top             =   105
         Width           =   1830
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Save"
         Height          =   525
         Left            =   270
         TabIndex        =   36
         Top             =   90
         Width           =   1830
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   7260
      Left            =   90
      ScaleHeight     =   7230
      ScaleWidth      =   11100
      TabIndex        =   0
      Top             =   915
      Width           =   11130
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Factory Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   135
         TabIndex        =   22
         Top             =   135
         Width           =   10815
         Begin VB.TextBox txtRef_No 
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
            Left            =   4530
            TabIndex        =   28
            Text            =   "Text4"
            Top             =   323
            Width           =   1410
         End
         Begin VB.ComboBox txtVendor_Code 
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
            Left            =   7320
            TabIndex        =   27
            Text            =   "Combo1"
            Top             =   330
            Width           =   1320
         End
         Begin VB.TextBox txtDate 
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
            Left            =   9210
            TabIndex        =   26
            Text            =   "Text3"
            Top             =   300
            Width           =   1380
         End
         Begin VB.TextBox txtInvoice 
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
            Left            =   1155
            TabIndex        =   25
            Text            =   "Text4"
            Top             =   323
            Width           =   1425
         End
         Begin VB.TextBox txtVendor_Address 
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
            Left            =   5160
            TabIndex        =   24
            Text            =   "Text2"
            Top             =   840
            Width           =   5445
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
            Left            =   1155
            TabIndex        =   23
            Text            =   "Combo1"
            Top             =   840
            Width           =   3000
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referance/Batch No."
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
            Left            =   2685
            TabIndex        =   34
            Top             =   390
            Width           =   1800
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
            Left            =   8700
            TabIndex        =   33
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Challan No."
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
            TabIndex        =   32
            Top             =   390
            Width           =   1005
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Godown Code"
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
            Left            =   6000
            TabIndex        =   31
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            TabIndex        =   30
            Top             =   870
            Width           =   510
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
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
            Left            =   4350
            TabIndex        =   29
            Top             =   870
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5385
         Left            =   150
         TabIndex        =   1
         Top             =   1590
         Width           =   10815
         Begin VB.ComboBox txtProd_Sl 
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
            Left            =   210
            TabIndex        =   10
            Top             =   600
            Width           =   1320
         End
         Begin VB.TextBox cmbPercent 
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
            Left            =   9135
            TabIndex        =   9
            Text            =   "Text2"
            Top             =   593
            Width           =   585
         End
         Begin VB.TextBox txtCommission 
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
            Left            =   9825
            TabIndex        =   8
            Text            =   "Text2"
            Top             =   593
            Width           =   795
         End
         Begin VB.TextBox txtProd_Cost 
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
            Left            =   6690
            TabIndex        =   7
            Text            =   "Text7"
            Top             =   593
            Width           =   1065
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
            Left            =   3585
            TabIndex        =   6
            Text            =   "Combo2"
            Top             =   600
            Width           =   2250
         End
         Begin VB.TextBox txtProd_Price 
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
            Left            =   7860
            TabIndex        =   5
            Text            =   "Text2"
            Top             =   593
            Width           =   1185
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
            Left            =   1590
            TabIndex        =   4
            Top             =   600
            Width           =   1905
         End
         Begin VB.TextBox txtQty 
            Alignment       =   2  'Center
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
            Left            =   5895
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   593
            Width           =   690
         End
         Begin VB.ComboBox cmbType 
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
            TabIndex        =   2
            Top             =   1455
            Visible         =   0   'False
            Width           =   1365
         End
         Begin MSDataGridLib.DataGrid grdProd 
            Height          =   3780
            Left            =   210
            TabIndex        =   11
            Top             =   1095
            Width           =   10395
            _ExtentX        =   18336
            _ExtentY        =   6668
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
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
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   9330
            TabIndex        =   21
            Top             =   270
            Width           =   180
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Discount"
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
            Left            =   9825
            TabIndex        =   20
            Top             =   270
            Width           =   765
         End
         Begin VB.Label lblNet 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount Tk. 0.00"
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
            Left            =   8460
            TabIndex        =   19
            Top             =   5010
            Width           =   2055
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
            Left            =   225
            TabIndex        =   18
            Top             =   270
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
            Left            =   1590
            TabIndex        =   17
            Top             =   270
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
            Left            =   3585
            TabIndex        =   16
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label Label7 
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
            Left            =   6690
            TabIndex        =   15
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Price"
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
            Left            =   7875
            TabIndex        =   14
            Top             =   270
            Width           =   930
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
            Left            =   6090
            TabIndex        =   13
            Top             =   270
            Width           =   315
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "R.M. Type"
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
            Left            =   1440
            TabIndex        =   12
            Top             =   1125
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frmGD_Rtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Prod_Sl As Integer
Dim netAmnt As Double
Dim stk As Double
Dim rsProd As ADODB.Recordset
Dim Prod As String
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim sqlStr As String
Dim strloan As String
Dim str As String
Private Sub Col_Prod()
With grdProd
    .Columns(0).Width = 500
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1200
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1500
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1500
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 500
    .Columns(4).Alignment = dbgRight
    .Columns(5).Width = 1200
    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 1200
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 1200
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 1200
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 1200
    .Columns(9).Alignment = dbgRight
    
End With
End Sub
Private Sub grd_Prod()
Prod_Sl = 0
netAmnt = 0
strloan = Format$(Val(netAmnt), "###0.00")
lblNet.Caption = "Net Loan Amount Tk. " + strloan
    ' Create an initial recordset, just for demonstration purposes,
    ' and assign it to the DataGrid control's DataSource property.
    Set rsProd = New ADODB.Recordset
    With rsProd
        .Fields.Append "Sl", adBSTR
        .Fields.Append "Product Code", adBSTR
        .Fields.Append "Product Name", adBSTR
        .Fields.Append "Product Type", adBSTR
        .Fields.Append "Description", adBSTR
        .Fields.Append "Qty", adBSTR
        .Fields.Append "Unit Price", adBSTR
        .Fields.Append "Total Price", adBSTR
        .Fields.Append "Commission", adBSTR
        .Fields.Append "Net Amount", adBSTR
        .Open
   
    End With
    Set grdProd.DataSource = rsProd
    
End Sub
Private Sub Prod_Tran()

    Set rsN = New ADODB.Recordset
        rsN.Open "Prod_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Invo_No = txtInvoice.Text
        rsN!Ref_no = txtRef_No.Text
        rsN!Ref_Code = txtVendor_Code.Text
        rsN!Ref_Name = cmbVendor_Name.Text
        rsN!User_Id = User_Id
        rsN!User_Name = User_Name
        rsN!D_ate = txtDate.Text
        rsN!Prod_Code = rsProd![Product Code]
'        rsN!Prod_Type = rsProd![Product Type]
        rsN!Prod_Name = rsProd![Product Name]
        rsN!Prod_Model = rsProd![Description]
        rsN!Purchase = 0
        rsN!Sale = 0
        rsN!Return = Val(rsProd!Qty)
        rsN!Stock = rsU!Stock
        rsN!Prod_Price = rsProd![Unit Price]
        rsN!Com = Val(rsProd![Commission])
        rsN!Amount = rsU!Amount
        rsN.Update
        rsN.Close
 End Sub
Private Sub Fact_Cr()

    Set rsN = New ADODB.Recordset
        rsN.Open "Godown_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Invo_No = txtInvoice.Text
        rsN!Ref_no = txtRef_No.Text
        rsN!Ref_Code = txtVendor_Code.Text
        rsN!Ref_Name = cmbVendor_Name.Text
        rsN!User_Id = User_Id
        rsN!User_Name = User_Name
        rsN!D_ate = txtDate.Text
        rsN!Prod_Code = rsProd![Product Code]
        rsN!Prod_Name = rsProd![Product Name]
        rsN!Prod_Model = rsProd![Description]
        rsN!Return = Val(rsProd!Qty)
        rsN!Sale = 0
        rsN!Lift = 0
        rsN!Stock = rs!Stock
        rsN!Prod_Price = rsProd![Unit Price]
        rsN!Com = Val(rsProd![Commission])
        rsN!Amount = rs!Amount
        rsN.Update
        rsN.Close
 End Sub
Private Sub Fact_Tran()
On Error Resume Next
 Set rs = New ADODB.Recordset
        str = "select * from Godown_Master where Prod_Code like '" & Prod & "' And Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rs.EOF Then
             rs!Return = rs!Return + Val(rsProd!Qty)
             rs!Stock = rs!Stock - Val(rsProd!Qty)
             rs!Prod_Price = Val(rsProd![Unit Price])
             rs!Com = rs!Com + Val(rsProd![Commission])
             rs!Amount = rs!Stock * rs!Prod_Price
             rs.Update
             Call Fact_Cr
             'Call GL_Dr
             'Call Vendor_Cr
             'rsProd.MoveNext
             rs.Close
                    
    Else
        rs.Close
    
    
    
    Set rs = New ADODB.Recordset
        rs.Open "Godown_Master", conn, adOpenDynamic, adLockOptimistic, -1
        rs.AddNew
        
        rs!Prod_Code = rsProd![Product Code]
        rs!Prod_Name = rsProd![Product Name]
        rs!Prod_Model = rsProd![Description]
        rs!Open_Bal = 0
        rs!Purchase = 0
        rs!Sale = 0
        rs!Return = Val(rsProd!Qty)
        rs!Return = 0
        rs!Stock = rs!Stock - Val(rsProd!Qty)
        rs!Prod_Price = Val(rsProd![Unit Price])
        rs!Com = Val(rsProd![Commission])
        rs!Amount = rs!Stock * rs!Prod_Price
        rs!Sale_Price = Val(rsProd![Sale Price])
        rs!Dealer_Price = Val(rsProd![Dealer Price])
        rs.Update
      
'        Call Prod_Tran
'        Call GL_Dr
'        Call Vendor_Cr
        'rsProd.MoveNext
        rs.Close
    End If

End Sub
Private Sub Vendor_Cr()
On Error Resume Next

Set rs = New ADODB.Recordset
        str = "select * from Godown_Master where Prod_Code like '" & Prod & "' And Vendor_Code Like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
   
        rs!Return = rs!Return + Val(rsProd!Qty)
        rs!Stock = rs!Stock - Val(rsProd!Qty)
        rs!Prod_Price = Val(rsProd![Unit Price])
        rs!Amount = rs!Stock * rs!Prod_Price
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Godown_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Invo_No = txtInvoice.Text
        rsN!Ref_no = txtRef_No.Text
        rsN!Ref_Code = txtVendor_Code.Text
        rsN!Ref_Name = cmbVendor_Name.Text
        rsN!User_Id = User_Id
        rsN!User_Name = User_Name
        rsN!D_ate = txtDate.Text
        rsN!Prod_Code = rsProd![Product Code]
        rsN!Prod_Type = rsProd![Product Type]
        rsN!Prod_Name = rsProd![Product Name]
        rsN!Prod_Model = rsProd![Description]
        rsN!Return = Val(rsProd!Qty)
        rsN!Sale = 0
        rsN!Lift = 0
        rsN!Stock = rs!Stock
        rsN!Prod_Price = rsProd![Unit Price]
        rsN!Com = Val(rsProd![Commission])
        rsN!Amount = rs!Amount
        rsN.Update
        rsN.Close
        rs.Close
    
    Else
            rs.Close
    
    
    Set rsN = New ADODB.Recordset
            rsN.Open "Godown_Master", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew

            rsN!Vendor_Code = txtVendor_Code.Text
            rsN!Vendor_Name = cmbVendor_Name.Text
            rsN!Vendor_Address = txtVendor_Address.Text
            rsN!Prod_Code = rsProd![Product Code]
'            rsN!Prod_Type = rsProd![Product Type]
            rsN!Prod_Name = rsProd![Product Name]
            rsN!Prod_Model = rsProd![Description]
            rsN!Open_Bal = 0
            rsN!Purchase = 0
            rsN!Sale = 0
            rsN!Lift = 0
            rsN!Return = Val(rsProd!Qty)
            rsN!Stock = rsN!Stock - Val(rsProd!Qty)
            rsN!Prod_Price = Val(rsProd![Unit Price])
            rsN!Com = Val(rsProd![Commission])
            rsN!Amount = rsN!Stock * rsN!Prod_Price
            rsN!Sale_Price = Val(rsProd![Sale Price])
            rsN!Dealer_Price = Val(rsProd![Dealer Price])
            rsN.Update


        Set rs = New ADODB.Recordset
            rs.Open "Godown_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew

            rs!Invo_No = txtInvoice.Text
            rs!Ref_no = txtRef_No.Text
            rs!Ref_Code = txtVendor_Code.Text
            rs!Ref_Name = cmbVendor_Name.Text
            rs!User_Id = User_Id
            rs!User_Name = User_Name
            rs!D_ate = txtDate.Text
            rs!Prod_Code = rsProd![Product Code]
            rs!Prod_Type = rsProd![Product Type]
            rs!Prod_Name = rsProd![Product Name]
            rs!Prod_Model = rsProd![Description]
            rs!Return = Val(rsProd!Qty)
            rs!Sale = 0
            rs!Lift = 0
            rs!Stock = rsN!Stock
            rs!Prod_Price = rsProd![Unit Price]
            rs!Com = Val(rsProd![Commission])
            rs!Amount = rsN!Amount
            rs.Update
            rs.Close
    End If
End Sub
Private Sub Vendor_Rev()
Set rs = New ADODB.Recordset
        str = "select * from Godown_Master where Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Dr = rs!Dr + Val(rsProd![Total Price])
        rs!Balance = rs!Balance - Val(rsProd![Total Price])
        rs.Update
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Godown_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Vendor_Code = txtVendor_Code.Text
            rsN!Vendor_Name = cmbVendor_Name.Text
            rsN!Vendor_Address = txtVendor_Address.Text
            rsN!Description = rsProd![Product Name] + "," + rsProd!Description
            rsN!Dr = Val(rsProd![Total Price])
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
    Else
            rs.Close
    End If
End Sub
Private Sub Cash_Cr()
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtCash.Text)
        rs!Date = txtDate.Text
        rs.Update
            
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "SCode-" & txtVendor_Code.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtCash.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        
        Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!MR_No = txtInvoice.Text
        rsN!Name = cmbVendor_Name.Text
        rsN!Description = "SCode-" & txtVendor_Code.Text
        rsN!Cr = Val(txtCash.Text)
        rsN!Balance = rs!Balance
        rsN!Tk1000 = Val(txt1000.Text)
        rsN!Tk500 = Val(txt500.Text)
        rsN!Tk100 = Val(txt100.Text)
        rsN!Tk50 = Val(txt50.Text)
        rsN!Tk20 = Val(txt20.Text)
        rsN!Tk10 = Val(txt10.Text)
        rsN!Tk5 = Val(txt5.Text)
        rsN!Tk2 = Val(txt2.Text)
        rsN!Tk1 = Val(txt1.Text)
        rsN.Update
        rsN.Close
        rs.Close
       
       
    
            Set rsU = New ADODB.Recordset
                str = "select * from Cash_Master where Code like '" & 1000 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                
                rsU!Cr = rsU!Cr + Val(txt1000.Text)
                rsU!Balance = rsU!Balance - Val(txt1000.Text)
                rsU!Cash_Cr = rsU!Cash_Cr + (Val(txt1000.Text) * 1000)
                rsU!Cash_Close = rsU!Cash_Close - (Val(txt1000.Text) * 1000)
                rsU.Update
                rsU.Close
                
                
            Set rsU = New ADODB.Recordset
                str = "select * from Cash_Master where Code like '" & 500 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                
                rsU!Cr = rsU!Cr + Val(txt500.Text)
                rsU!Balance = rsU!Balance - Val(txt500.Text)
                rsU!Cash_Cr = rsU!Cash_Cr + (Val(txt500.Text) * 500)
                rsU!Cash_Close = rsU!Cash_Close - (Val(txt500.Text) * 500)
                rsU.Update
                rsU.Close
            
            Set rsU = New ADODB.Recordset
                str = "select * from Cash_Master where Code like '" & 100 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                
                rsU!Cr = rsU!Cr + Val(txt100.Text)
                rsU!Balance = rsU!Balance - Val(txt100.Text)
                rsU!Cash_Cr = rsU!Cash_Cr + (Val(txt100.Text) * 100)
                rsU!Cash_Close = rsU!Cash_Close - (Val(txt100.Text) * 100)
                rsU.Update
                rsU.Close
            
            
            Set rsU = New ADODB.Recordset
                str = "select * from Cash_Master where Code like '" & 50 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                
                rsU!Cr = rsU!Cr + Val(txt50.Text)
                rsU!Balance = rsU!Balance - Val(txt50.Text)
                rsU!Cash_Cr = rsU!Cash_Cr + (Val(txt50.Text) * 50)
                rsU!Cash_Close = rsU!Cash_Close - (Val(txt50.Text) * 50)
                rsU.Update
                rsU.Close
            
            Set rsU = New ADODB.Recordset
                str = "select * from Cash_Master where Code like '" & 20 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                
                rsU!Cr = rsU!Cr + Val(txt20.Text)
                rsU!Balance = rsU!Balance - Val(txt20.Text)
                rsU!Cash_Cr = rsU!Cash_Cr + (Val(txt20.Text) * 20)
                rsU!Cash_Close = rsU!Cash_Close - (Val(txt20.Text) * 20)
                rsU.Update
                rsU.Close
        
            Set rsU = New ADODB.Recordset
                str = "select * from Cash_Master where Code like '" & 10 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                
                rsU!Cr = rsU!Cr + Val(txt10.Text)
                rsU!Balance = rsU!Balance - Val(txt10.Text)
                rsU!Cash_Cr = rsU!Cash_Cr + (Val(txt10.Text) * 10)
                rsU!Cash_Close = rsU!Cash_Close - (Val(txt10.Text) * 10)
                rsU.Update
                rsU.Close
        
            Set rsU = New ADODB.Recordset
                str = "select * from Cash_Master where Code like '" & 5 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                
                rsU!Cr = rsU!Cr + Val(txt5.Text)
                rsU!Balance = rsU!Balance - Val(txt5.Text)
                rsU!Cash_Cr = rsU!Cash_Cr + (Val(txt5.Text) * 5)
                rsU!Cash_Close = rsU!Cash_Close - (Val(txt5.Text) * 5)
                rsU.Update
                rsU.Close
            
            Set rsU = New ADODB.Recordset
                str = "select * from Cash_Master where Code like '" & 2 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                
                rsU!Cr = rsU!Cr + Val(txt2.Text)
                rsU!Balance = rsU!Balance - Val(txt2.Text)
                rsU!Cash_Cr = rsU!Cash_Cr + (Val(txt2.Text) * 2)
                rsU!Cash_Close = rsU!Cash_Close - (Val(txt2.Text) * 2)
                rsU.Update
                rsU.Close
        
            Set rsU = New ADODB.Recordset
                str = "select * from Cash_Master where Code like '" & 1 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                
                rsU!Cr = rsU!Cr + Val(txt1.Text)
                rsU!Balance = rsU!Balance - Val(txt1.Text)
                rsU!Cash_Cr = rsU!Cash_Cr + (Val(txt1.Text) * 1)
                rsU!Cash_Close = rsU!Cash_Close - (Val(txt1.Text) * 1)
                rsU.Update
                rsU.Close
    
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Cash_Cr = rsU!Cash_Cr + Val(txtCash.Text)
        rsU!Cash_Close = rsU!Cash_Close - Val(txtCash.Text)
        rsU.Update
        rsU.Close
        
        
     Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100150 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtCash.Text)
        rs.Update
        
     Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "SCode-" & txtVendor_Code.Text
        rsN!Dr = Val(txtCash.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close

        rs.Close
        
    Set rs = New ADODB.Recordset
        str = "select * from Godown_Master where Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Dr = rs!Dr + Val(txtCash.Text)
        rs!Balance = rs!Balance - Val(txtCash.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Godown_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!Vendor_Code = txtVendor_Code.Text
        rsN!Vendor_Name = cmbVendor_Name.Text
        rsN!Vendor_Address = txtVendor_Address.Text
        rsN!Description = "Cash Payment"
        rsN!Dr = Val(txtCash.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        
        rsN.Update
        rsN.Close
        rs.Close
    End If
End Sub
Private Sub Cash_Rev()
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
        rsN!Description = "Entry Reversed"
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
        rsN!Description = "Entry Reversed"
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
        str = "select * from GL_Master where AC_No like '" & 100103 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtCash.Text)
        rs.Update
        
     Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Entry Reversed"
        rsN!Dr = 0
        rsN!Cr = Val(txtCash.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close

        rs.Close
        
    Set rs = New ADODB.Recordset
        str = "select * from Godown_Master where Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Cr = rs!Cr + Val(txtCash.Text)
        rs!Balance = rs!Balance + Val(txtCash.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Godown_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!Vendor_Code = txtVendor_Code.Text
        rsN!Vendor_Name = cmbVendor_Name.Text
        rsN!Vendor_Address = txtVendor_Address.Text
        rsN!Description = "Entry Reversed"
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
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
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
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
        
    Set rs = New ADODB.Recordset
        str = "select * from Bank_Master where AC_No like '" & txtAccount.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs!Withdraw = rs!Withdraw + Val(txtChq_Amnt.Text)
        rs!Date = txtDate.Text
        rs.Update
        
        Set rsN = New ADODB.Recordset
        rsN.Open "Bank_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = txtAccount.Text
        rsN!Bank_Name = cmbBank.Text
        rsN!Branch_Name = cmbBranch.Text
        rsN!MR_No = txtInvoice.Text
        rsN!Chq_No = "CN" & txtChq_No.Text
        rsN!Description = "SCode-" & txtVendor_Code.Text
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
        
        
     Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100150 & "'"
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
        rsN!Dr = Val(txtChq_Amnt.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close

        rs.Close

Set rs = New ADODB.Recordset
        str = "select * from Godown_Master where Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Dr = rs!Dr + Val(txtChq_Amnt.Text)
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Godown_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!Vendor_Code = txtVendor_Code.Text
        rsN!Vendor_Name = cmbVendor_Name.Text
        rsN!Vendor_Address = txtVendor_Address.Text
        rsN!Description = "CN" & txtChq_No.Text
        rsN!Dr = Val(txtChq_Amnt.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        
        rsN.Update
        rsN.Close
        rs.Close
    End If
End Sub
Private Sub Bank_Rev()
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
        rsN!Description = "Entry Reversed"
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
        rsN!Chq_No = "CN" & txtChq_No.Text
        rsN!Description = "Entry Reversed"
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
        str = "select * from GL_Master where AC_No like '" & 100103 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtChq_Amnt.Text)
        rs.Update
        
     Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Entry Reversed"
        rsN!Dr = 0
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close

        rs.Close

Set rs = New ADODB.Recordset
        str = "select * from Godown_Master where Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Dr = rs!Dr + Val(txtChq_Amnt.Text)
        rs!Balance = rs!Balance + Val(txtChq_Amnt.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Godown_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!Vendor_Code = txtVendor_Code.Text
        rsN!Vendor_Name = cmbVendor_Name.Text
        rsN!Vendor_Address = txtVendor_Address.Text
        rsN!Description = "Entry Reversed"
        rsN!Dr = 0
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        
        rsN.Update
        rsN.Close
        rs.Close
    End If
End Sub
Private Sub GL_Dr()
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100103 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - rsProd![Total Price]
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "PCode-" & rsProd![Product Code]
        rsN!Dr = 0
        rsN!Cr = rsProd![Total Price]
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100102 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + rsProd![Total Price]
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "FCode-" & txtVendor_Code.Text
        rsN!Dr = rsProd![Total Price]
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

End Sub
Private Sub GL_Rev()
If rsProd![Product Code] = 101 Then
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100103 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - (Val(rsProd![Total Price]) + Val(rsProd![Commission]))
        rs!Date = txtDate.Text
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "PCode-" & rsProd![Product Code]
        rsN!Dr = 0
        rsN!Cr = Val(rsProd![Total Price]) + Val(rsProd![Commission])
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100150 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(rsProd![Total Price])
        rs!Date = txtDate.Text
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "SCode-" & txtVendor_Code.Text
        rsN!Dr = Val(rsProd![Total Price])
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100143 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(rsProd![Commission])
        rs!Date = txtDate.Text
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "SCode-" & txtVendor_Code.Text
        rsN!Dr = Val(rsProd![Commission])
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
Else
    
    If rsProd![Product Code] = "BKS01" Then
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100105 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - (Val(rsProd![Total Price]) + Val(rsProd![Commission]))
        rs!Date = txtDate.Text
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "PCode-" & rsProd![Product Code]
        rsN!Dr = 0
        rsN!Cr = Val(rsProd![Total Price]) + Val(rsProd![Commission])
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100150 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(rsProd![Total Price])
        rs!Date = txtDate.Text
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "SCode-" & txtVendor_Code.Text
        rsN!Dr = Val(rsProd![Total Price])
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100143 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(rsProd![Commission])
        rs!Date = txtDate.Text
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "SCode-" & txtVendor_Code.Text
        rsN!Dr = Val(rsProd![Commission])
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
Else
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100102 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - rsProd![Total Price]
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "PCode-" & rsProd![Product Code]
        rsN!Dr = 0
        rsN!Cr = rsProd![Total Price]
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100150 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - rsProd![Total Price]
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "SCode-" & txtVendor_Code.Text
        rsN!Dr = rsProd![Total Price]
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

    End If
End If
End Sub
Private Sub Total_Cash()
    Total = (Val(txt1000.Text) * 1000) + (Val(txt500.Text) * 500) + (Val(txt100.Text) * 100) + (Val(txt50.Text) * 50) + (Val(txt20.Text * 20)) + (Val(txt10.Text) * 10) + (Val(txt5.Text) * 5) + (Val(txt2.Text) * 2) + (Val(txt1.Text) * 1) + Val(txtChq_Amnt.Text)
    txtCash.Text = Format(Total - Val(txtChq_Amnt.Text), "###0.00")
    lblPmt.Caption = "Total Paid Amount Tk. " + Format$(Val(Total), "###0.00")
End Sub
Private Sub Prod_Code()
        txtProd_Sl.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Code FROM Prod_Master"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtProd_Sl.AddItem rs!Prod_Code
        rs.MoveNext
        Loop
        rs.Close
End Sub
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

Private Sub Vendor_Code()
        txtVendor_Code.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Vendor_Code FROM Godown_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        txtVendor_Code.AddItem rsN!Vendor_Code
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub Vendor_Name()
        cmbVendor_Name.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Vendor_Name FROM Godown_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbVendor_Name.AddItem rsN!Vendor_Name
        rsN.MoveNext
        Loop
        rsN.Close
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
        txtRef_No.Text = ""
        txtDate.Text = ""
        txtVendor_Code.Text = ""
        cmbVendor_Name.Text = ""
        txtVendor_Address.Text = ""
        
        txtProd_Sl.Text = ""
        cmbProd_Name.Text = ""
        cmbType.Text = ""
        cmbProd_Model.Text = ""
        
        txtProd_Cost.Text = "0.00"
        txtQty.Text = "0"
        txtProd_Price.Text = "0.00"
        txtCommission.Text = "0.00"
        cmbPercent.Text = "0.00"
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
Private Sub cmbBrand_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbOrigin.SelStart = 0
    cmbOrigin.SelLength = Len(cmbOrigin.Text)
    cmbOrigin.SetFocus
End If
End Sub
Private Sub cmbColor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSl.SelStart = 0
    txtSl.SelLength = Len(txtSl.Text)
    txtSl.SetFocus
End If
End Sub

Private Sub cmbOrigin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbColor.SelStart = 0
    cmbColor.SelLength = Len(cmbColor.Text)
    cmbColor.SetFocus
End If
End Sub

Private Sub cmbPercent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtCommission.SetFocus
End If
End Sub
Private Sub cmbPercent_LostFocus()
If txtProd_Sl.Text = "101" Or txtProd_Sl.Text = "BKS01" Then
txtCommission.Text = Format$(Val(txtProd_Cost.Text) * (Val(cmbPercent.Text) / 100), "###0.00")
txtProd_Price.Text = Format$(Val(txtProd_Cost.Text) - (Val(txtProd_Cost.Text) * (Val(cmbPercent.Text) / 100)), "###0.00")
    txtCommission.SelStart = 0
    txtCommission.SelLength = Len(txtCommission.Text)
    txtCommission.SetFocus
Else
txtCommission.Text = Format$(Val(txtProd_Price.Text) * (Val(cmbPercent.Text) / 100), "###0.00")
'txtProd_Price.Text = Format$(Val(txtProd_Cost.Text) - (Val(txtProd_Cost.Text) * (Val(cmbPercent.Text) / 100)), "###0.000")
    txtCommission.SelStart = 0
    txtCommission.SelLength = Len(txtCommission.Text)
    txtCommission.SetFocus
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
        str = "SELECT * FROM Prod_Master where Prod_Name Like '" & cmbProd_Name.Text & "' AND Prod_Model Like '" & cmbProd_Model.Text & "' order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            cmbBrand.Text = rsN!Brand_Name
            cmbOrigin.Text = rsN!Origin
           ' txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtS_Price.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            txtD_Price.Text = Format$(Val(rsN!Dealer_Price), "###0.00")
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
        str = "SELECT DISTINCT Prod_Name, Prod_Model FROM Prod_Master where Prod_Name Like '" & Prod & "'"
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
mid = 0
        
        Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Godown_Master where Vendor_Name Like '" & Vendor & "'order by Vendor_code"
        rsN.Open str, conn
        
If Not rsN.EOF Then
            txtVendor_Code.Text = rsN!Vendor_Code
            cmbVendor_Name.Text = rsN!Vendor_Name
            txtVendor_Address.Text = rsN!Vendor_Address
            rsN.Close
Else
        
    If MsgBox("Do you want add new Vendor?", vbInformation + vbYesNo, "Add New") = vbYes Then
        
        Set rsU = New ADODB.Recordset
            str = "select * from Godown_Master order by Vendor_Code"
            rsU.Open str, conn
        
        If Not rsU.EOF Then
           rsU.MoveFirst
        
            Do While Not rsU.EOF = True
                mid = Val(rsU!Vendor_Code)
                rsU.MoveNext
            Loop
                rsU.Close
                mid = mid + 1
            Else
                rsU.Close
                mid = 101
            End If
            txtVendor_Code.Text = mid
            'cmbVendor_Name.Text = ""
            txtVendor_Address.Text = ""
     Else
        rsN.Close
        Exit Sub
    End If
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Really want to delete?", vbCritical + vbYesNo) = vbYes Then

    Dim Tran As String
    Dim Prod As String
        Tran = txtInvoice.Text
    
        rsProd.MoveFirst
    
    Do While Not rsProd.EOF
        Prod = rsProd![Product Code]
        
    
        Set rsU = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Code like '" & Prod & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
        If rsProd![Product Code] = "101" Or rsProd![Product Code] = "BKS01" Then
             rsU!Purchase = rsU!Purchase - Val(rsProd![Total Price])
             rsU!Stock = rsU!Stock - Val(rsProd![Total Price])
             rsU!Com = rsU!Com - Val(rsProd![Commission])
             rsU!Amount = rsU!Amount - (Val(rsProd![Total Price]) + Val(rsProd![Commission]))
             rsU.Update
             rsU.Close
        Else
            rsU!Purchase = rsU!Purchase - Val(rsProd!Qty)
             rsU!Stock = rsU!Stock - Val(rsProd!Qty)
             rsU!Com = rsU!Com - Val(rsProd![Commission])
             rsU!Amount = rsU!Stock * rsU!Prod_Price
             rsU.Update
             rsU.Close
        End If
    Else
    MsgBox "Invalid Product Code!", vbCritical, "Error!"
    rsU.Close
    Exit Sub
    End If
    
    Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where Invo_No like '" & Tran & "' and Prod_Code like '" & Prod & "'"
        rs.Open str, conn

    If Not rs.EOF Then
        rs.Close
        
        Set rs = New ADODB.Recordset
        str = "delete * from Prod_Tran where Invo_No like '" & Tran & "' and Prod_Code like '" & Prod & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
'        rs.Close
    Else
    Exit Sub
    End If
    
    
    Call GL_Rev
    Call Vendor_Rev
  rsProd.MoveNext
  
Loop
    Call clearTextboxes
    Call grd_Prod
    Call Col_Prod
Else
Exit Sub
End If
End Sub

Private Sub cmdPrint_Click()
If txtInvoice.Text = "" Then
Exit Sub
End If
Dim Tran As String
    Tran = txtInvoice.Text

    Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where Invo_No like '" & Tran & "'"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly

        If Not rs.EOF Then
            rptInvoice.rsInvoice.ConnectionString = cnStr
            rptInvoice.rsInvoice.Source = str
            
            rptInvoice.Field19.DataField = "Purchase"
            rptInvoice.Field16.DataField = "Prod_Price"
            rptInvoice.lblTitle.Caption = "PURCHASE INVOICE"
            rptInvoice.txtTotal.Text = Format(Val(strloan), "#,##0.00")
            rs.Close
        Else
            MsgBox "There is no such Invoice found, ", vbCritical + vbOKOnly
            rs.Close
        End If
    rptInvoice.Show 1
    cmdPrint.Enabled = False
End Sub

Private Sub cmdReceive_Click()
On Error Resume Next
If txtInvoice.Text = "" Then
MsgBox "Please Input Invoice No!", vbCritical
txtInvoice.SetFocus
Exit Sub
End If





rsProd.MoveFirst
    
    Do While Not rsProd.EOF
        Prod = rsProd![Product Code]
        
        Set rsU = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Code like '" & Prod & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
             Call Vendor_Cr
             rsU!Return = rsU!Return + Val(rsProd!Qty)
             rsU!Stock = rsU!Stock + Val(rsProd!Qty)
             rsU!Prod_Price = Val(rsProd![Unit Price])
             rsU!Com = rsU!Com + Val(rsProd![Commission])
             rsU!Amount = rsU!Stock * rsU!Prod_Price
             rsU.Update
             
             Call Prod_Tran
             Call GL_Dr
             
             'Call Fact_Tran
             'Call Fact_cr
             rsProd.MoveNext
             rsU.Close
                    
    Else
        rsU.Close
    
    Call Vendor_Cr
    
    Set rsU = New ADODB.Recordset
        rsU.Open "Prod_Master", conn, adOpenDynamic, adLockOptimistic, -1
        rsU.AddNew

        rsU!Prod_Code = rsProd![Product Code]
'        rsU!Prod_Type = rsProd![Product Type]
        rsU!Prod_Name = rsProd![Product Name]
        rsU!Prod_Model = rsProd![Description]
        rsU!Open_Bal = 0
        rsU!Purchase = 0
        rsU!Sale = 0
        rsU!Lift = Val(rsProd!Qty)
        rsU!Return = 0
        rsU!Stock = rsU!Stock - Val(rsProd!Qty)
        rsU!Prod_Price = Val(rsProd![Unit Price])
        rsU!Com = Val(rsProd![Commission])
        rsU!Amount = rsU!Stock * rsU!Prod_Price
        rsU!Sale_Price = Val(rsProd![Sale Price])
        rsU!Dealer_Price = Val(rsProd![Dealer Price])
        rsU.Update

        Call Prod_Tran
        Call GL_Dr
        
'        Call Fact_Tran
'        Call Fact_Cr
        rsProd.MoveNext
        rsU.Close
    End If


    Loop
   
    Call Prod_Name
    Call Prod_Model
    Call Prod_Code
    
    Call Vendor_Name
    Call Vendor_Code
    Call clearTextboxes
    Call grd_Prod
    Call Col_Prod
    
    txtDate.Text = Today
    txtInvoice.SetFocus
    cmdReceive.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
Exit Sub
   Resume Next
End Sub

Private Sub cmdUpdate_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
   
    Call Prod_Name
    Call Prod_Code
    Call Prod_Model
   
    
    Call grd_Prod
    Call Col_Prod
    Call Vendor_Name
    Call Vendor_Code
    Call Account_Name
    Call Bank_Name
    Call Branch_Name
    
    Call clearTextboxes
        
        txtDate.Text = Today
        
        cmdReceive.Enabled = False
        cmdUpdate.Enabled = False
    Exit Sub
End Sub

Private Sub grdProd_DblClick()

If MsgBox("Do you want to delete record?", vbCritical + vbYesNo, "Delete") = vbYes Then

With rsProd
        .Delete (adAffectCurrent)
End With
    rsProd.MoveFirst
    Prod_Sl = 0
    netAmnt = 0
    Do While Not rsProd.EOF
        Prod_Sl = Prod_Sl + 1
        rsProd!sl = Prod_Sl
        netAmnt = netAmnt + Val(rsProd![Total Price])
        rsProd.MoveNext
    Loop
    
    strloan = Format$(Val(netAmnt), "###0.00")
    lblNet.Caption = "Total Amount Tk. " + strloan
Else
Exit Sub
End If
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
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn
        
        If rs!Balance < Val(txtCash.Text) Then
        MsgBox "Insufficiant Balance.!", vbCritical, "Error!"
        txtChq_Amnt.SetFocus
        Else
        txtCash.Text = Format$(Val(txtCash.Text), "###0.00")
        End If
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
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100101 & "'"
        rs.Open str, conn
        
        If rs!Balance < Val(txtChq_Amnt.Text) Then
        MsgBox "Insufficiant Balance.!", vbCritical, "Error!"
        txtChq_Amnt.SetFocus
        Else
        txtChq_Amnt.Text = Format$(Val(txtChq_Amnt.Text), "###0.00")
        End If

End Sub

Private Sub txtChq_No_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtChq_Amnt.SetFocus
End If
End Sub

Private Sub txtCommission_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCommission.Text = Format$(Val(txtCommission.Text), "###0.00")
    txtProd_Sl.SetFocus
End If
End Sub

Private Sub txtEngine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty.Text)
    txtQty.SetFocus
End If
End Sub

Private Sub txtCommission_LostFocus()
txtProd_Price.Text = Format$(Val(txtProd_Price.Text), "###0.00")
        rsProd.AddNew
        rsProd!sl = Prod_Sl
        rsProd![Product Code] = txtProd_Sl.Text
'        rsProd![Product Type] = cmbType.Text
        rsProd![Product Name] = cmbProd_Name.Text
        rsProd![Description] = cmbProd_Model.Text
        rsProd!Qty = txtQty.Text
        rsProd![Unit Price] = Format$(Val(txtProd_Cost.Text), "###0.00")
        rsProd![Total Price] = Format$(Val(txtProd_Price.Text), "###0.00")
        rsProd![Commission] = Format$(Val(txtCommission.Text), "###0.00")
        rsProd![Net Amount] = Format$((Val(txtProd_Price) + Val(txtCommission.Text)), "###0.00")
        rsProd.Update
        
        rsProd.MoveFirst
        Prod_Sl = 0
        netAmnt = 0
        Do While Not rsProd.EOF
            Prod_Sl = Prod_Sl + 1
            rsProd!sl = Prod_Sl
            
            netAmnt = netAmnt + Val(rsProd![Total Price]) - Val(txtCommission.Text)
            rsProd.MoveNext
        Loop
        strloan = 0
        strloan = Format$(Round(Val(netAmnt)), "###0.00")
        lblNet.Caption = "Total Amount Payable Tk. " + strloan
        
    Set grdProd.DataSource = rsProd
    grdProd.Refresh
    
    Call Col_Prod
            txtProd_Sl.Text = ""
            cmbProd_Name.Text = ""
            cmbType.Text = ""
            cmbProd_Model.Text = ""
            txtQty.Text = ""
            txtProd_Cost.Text = "0.00"
            txtProd_Price.Text = "0.00"
            txtCommission.Text = "0.00"
            cmbPercent.Text = "0.00"

End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtInvoice.Text = "" Then
Exit Sub
Else
cmdReceive.Enabled = True
txtRef_No.SetFocus
End If
End If
End Sub

Private Sub txtInvoice_LostFocus()
On Error Resume Next
Dim Tran As String
    Tran = txtInvoice.Text

    Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where Invo_No like '" & Tran & "'"
        rs.Open str, conn

 If Not rs.EOF Then
    MsgBox "Duplicate invoice No.", vbCritical, "Erroe!"
            rs.MoveFirst
                txtVendor_Code.Text = rs!Ref_Code
                cmbVendor_Name.Text = rs!Ref_Name
                Date = rs!D_ate
            
                Prod_Sl = 0
                netAmnt = 0

            Do While Not rs.EOF
                Prod_Sl = Prod_Sl + 1
                rsProd.AddNew
        
                rsProd!sl = Prod_Sl
                rsProd![Product Code] = rs!Prod_Code
                rsProd![Product Name] = rs!Prod_Name
                rsProd![Description] = rs!Prod_Model
                rsProd!Qty = rs!Purchase
                rsProd![Unit Price] = Format$(Val(rs!Prod_Price), "###0.000")
                rsProd![Total Price] = Format$(Val(rs!Prod_Price * rs!Purchase), "###0.000")
                rsProd![Commission] = Format$(Val(rs!Com), "###0.000")
                rsProd![Net Amount] = Format$(Val(rs!Purchase * rs!Prod_Price), "###0.000")
                netAmnt = netAmnt + Val(rs!Purchase * rs!Prod_Price)
        
                
                
                Set grdProd.DataSource = rsProd
                grdProd.Refresh
                Call Col_Prod
                rs.MoveNext
            Loop
        
    rs.Close
    
        strloan = Format$(Round(Val(netAmnt)), "###0.000")
        lblNet.Caption = "Total Amount Payable Tk. " + strloan
    
    
        Set rs = New ADODB.Recordset
               str = "select * from Cash_Book where MR_No like '" & Tran & "'"
               rs.Open str, conn
           
           If Not rs.EOF Then
               txtCash.Text = rs!Cr
               rs.Close
           Else
               rs.Close
'               txtCash.Text = 0
           End If
    
            Set rs = New ADODB.Recordset
               str = "select * from Bank_Tran where MR_No like '" & Tran & "'"
               rs.Open str, conn
           
           If Not rs.EOF Then
               txtChq_Amnt.Text = rs!Cr
               rs.Close
           Else
               rs.Close
'               txtChq_Amnt.Text = 0
           End If
            
    
        cmdPrint.Enabled = True
        cmdReceive.Enabled = False
        cmdDelete.Enabled = True
    Else
    rs.Close
    Call clearTextboxes
    txtInvoice.Text = Tran
    txtDate.Text = Today
    Exit Sub
  End If
End Sub

Private Sub txtQty_LostFocus()
'Dim Prod As String
'        Prod = txtProd_Sl.Text
'
'
'    Set rsN = New ADODB.Recordset
'        str = "SELECT * FROM Prod_Master where Prod_Code Like '" & Prod & "' order by Prod_code"
'        rsN.Open str, conn
'
'        If Not rsN.EOF Then
'
'                If rsN!Stock >= Val(txtQty.Text) Then
'                    txtProd_Cost.SelStart = 0
'                    txtProd_Cost.SelLength = Len(txtProd_Cost.Text)
'                    txtProd_Cost.SetFocus
'                Else
'                    MsgBox "Stock not available!", vbCritical, "Sales Info!"
'                    txtQty.Text = 0
'                    txtProd_Cost.SelStart = 0
'                    txtProd_Cost.SelLength = Len(txtProd_Cost.Text)
'                    txtProd_Cost.SetFocus
'                End If
'            rsN.Close
'        End If

End Sub

Private Sub txtRef_No_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtVendor_Code.SetFocus
End If
End Sub

Private Sub txtProd_Cost_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Price.SetFocus
End If
End Sub

Private Sub txtProd_Cost_LostFocus()
If txtProd_Sl.Text = "101" Then

    txtProd_Price.Text = Format$((Val(txtProd_Cost.Text) - (Val(txtProd_Cost.Text) * 0.04023)), "###0.00")
    txtProd_Cost.Text = Format$(Val(txtProd_Cost.Text), "###0.00")
    cmbPercent.Text = Format$(Val(4.023), "###0.00")
    txtCommission.Text = Format$(Val(txtProd_Cost.Text) * 0.04023, "###0.00")
    txtCommission.SelStart = 0
    txtCommission.SelLength = Len(txtCommission.Text)
    txtCommission.SetFocus

    Else

    If txtProd_Sl.Text = "BKS01" Then

    txtProd_Price.Text = Format$((Val(txtProd_Cost.Text)), "###0.00")
    txtProd_Cost.Text = Format$(Val(txtProd_Cost.Text), "###0.00")
    txtProd_Price.SelStart = 0
    txtProd_Price.SelLength = Len(txtProd_Price.Text)
    txtCommission.SetFocus
    Else
        txtProd_Price.Text = Format$(Val(txtQty.Text) * Val(txtProd_Cost.Text), "###0.00")
        txtProd_Cost.Text = Format$(Val(txtProd_Cost.Text), "###0.00")
        
        txtCommission.SelStart = 0
        txtCommission.SelLength = Len(txtCommission.Text)
        txtCommission.SetFocus
    End If
 End If
End Sub

Private Sub txtProd_Price_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtProd_Price.Text = Format$(Val(txtProd_Price.Text), "###0.00")
    cmbPercent.SelStart = 0
    cmbPercent.SelLength = Len(cmbPercent.Text)
    cmbPercent.SetFocus
End If
End Sub

Private Sub txtProd_Price_LostFocus()
txtProd_Price.Text = Format$(Val(txtProd_Price.Text), "###0.00")
End Sub

Private Sub txtProd_Sl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtQty.SetFocus
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Cost.SelStart = 0
txtProd_Cost.SelLength = Len(txtProd_Cost.Text)
txtProd_Cost.SetFocus
End If
End Sub
Private Sub txtProd_Sl_LostFocus()
If txtProd_Sl.Text = "" Then
cmdReceive.SetFocus
Exit Sub
End If
    Dim Prod As String
    Prod = txtProd_Sl.Text
    
    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
'            cmbType.Text = rsN!Prod_Type
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            txtProd_Cost.Text = Format$(Val(rsN!Prod_Price), "###0.00")
            rsN.Close
            txtQty.SelStart = 0
            txtQty.SelLength = Len(txtQty.Text)
            txtQty.SetFocus
            
        Else
        If MsgBox("Do you want add new Product?", vbInformation + vbYesNo, "Add New") = vbYes Then
        cmbProd_Name.SetFocus
        End If
        End If
    Exit Sub
End Sub
Private Sub txtSearch_KeyPress(KeyAscii As Integer)
On Error Resume Next
If txtSearch.Text = "" Then
Exit Sub
End If

    Dim search As String
        search = txtSearch.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where Prod_code Like '" & "%" & search & "%" & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
               
       ' Call ColumnWidth
    End If
    Exit Sub
Last:
    MsgBox ("Database Connection error: " + Err.Description)
End Sub

Private Sub txtD_Price_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtD_Price.Text = Format$(Val(txtD_Price.Text), "###0.00")
    
End If
End Sub

Private Sub txtD_Price_LostFocus()
End Sub
Private Sub txtS_Price_KeyPress(KeyAscii As Integer)
End Sub

Private Sub txtSl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtEngine.SelStart = 0
    txtEngine.SelLength = Len(txtEngine.Text)
    txtEngine.SetFocus
End If
End Sub

Private Sub txtVendor_Address_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Sl.SetFocus
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
mid = 0
        
        Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Godown_Master where Vendor_Code Like '" & Vendor & "'order by Vendor_code"
        rsN.Open str, conn
        
If Not rsN.EOF Then
            txtVendor_Code.Text = rsN!Vendor_Code
            cmbVendor_Name.Text = rsN!Vendor_Name
            txtVendor_Address.Text = rsN!Vendor_Address
            rsN.Close
Else
        
    If MsgBox("Do you want add new Vendor?", vbInformation + vbYesNo, "Add New") = vbYes Then
        
        Set rsU = New ADODB.Recordset
            str = "select * from Godown_Master order by Vendor_Code"
            rsU.Open str, conn
        
        If Not rsU.EOF Then
           rsU.MoveFirst
        
            Do While Not rsU.EOF = True
                mid = Val(rsU!Vendor_Code)
                rsU.MoveNext
            Loop
                rsU.Close
                mid = mid + 1
            Else
                rsU.Close
                mid = 101
            End If
            txtVendor_Code.Text = mid
            cmbVendor_Name.Text = ""
            txtVendor_Address.Text = ""
     Else
        rsN.Close
        Exit Sub
    End If
End If

End Sub


