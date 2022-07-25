VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLoanAccount 
   BackColor       =   &H00008000&
   Caption         =   "Loan Account Opening"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11205
   Icon            =   "frmLoan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Guarantor Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   5490
      TabIndex        =   87
      Top             =   3825
      Width           =   5550
      Begin VB.ComboBox cmbGtr_Id 
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
         Left            =   180
         TabIndex        =   90
         Text            =   "Combo2"
         Top             =   615
         Width           =   1320
      End
      Begin VB.ComboBox cmbGtr_Name 
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
         Left            =   1755
         TabIndex        =   89
         Text            =   "Combo3"
         Top             =   615
         Width           =   3615
      End
      Begin MSDataGridLib.DataGrid grdGtr 
         Height          =   1770
         Left            =   180
         TabIndex        =   88
         Top             =   1125
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   3122
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
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guarantor ID"
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
         TabIndex        =   92
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guarantor Name"
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
         Left            =   1755
         TabIndex        =   91
         Top             =   315
         Width           =   1425
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
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
      Height          =   2805
      Left            =   5490
      TabIndex        =   65
      Top             =   900
      Width           =   5550
      Begin MSDataGridLib.DataGrid grdProd 
         Height          =   1410
         Left            =   180
         TabIndex        =   79
         Top             =   1080
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   2487
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
      Begin VB.TextBox txtNet 
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
         Left            =   3825
         TabIndex        =   69
         Text            =   "Text2"
         Top             =   585
         Width           =   1545
      End
      Begin VB.ComboBox cmbProd_model 
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
         Left            =   1665
         TabIndex        =   68
         Text            =   "Combo1"
         Top             =   592
         Width           =   1320
      End
      Begin VB.ComboBox cmbProduct_Name 
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
         Left            =   180
         TabIndex        =   67
         Text            =   "Combo1"
         Top             =   592
         Width           =   1320
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
         Left            =   3150
         TabIndex        =   66
         Text            =   "0"
         Top             =   585
         Width           =   510
      End
      Begin VB.Label lblNet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Loan Amount Tk. 0.00"
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
         Left            =   2910
         TabIndex        =   80
         Top             =   2520
         Width           =   2430
      End
      Begin VB.Label Label37 
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
         Left            =   3240
         TabIndex        =   73
         Top             =   270
         Width           =   315
      End
      Begin VB.Label Label6 
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
         Left            =   180
         TabIndex        =   72
         Top             =   270
         Width           =   1245
      End
      Begin VB.Label Label42 
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
         Left            =   1665
         TabIndex        =   71
         Top             =   270
         Width           =   1260
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Price"
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
         TabIndex        =   70
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   135
      TabIndex        =   34
      Top             =   4545
      Width           =   5190
      Begin VB.TextBox txtContact_Office 
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
         Left            =   135
         TabIndex        =   45
         Text            =   "Text17"
         Top             =   4095
         Width           =   1455
      End
      Begin VB.ComboBox cmbReligion 
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
         Left            =   3465
         TabIndex        =   44
         Text            =   "Combo3"
         Top             =   1215
         Width           =   1545
      End
      Begin VB.TextBox txtBirthday 
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
         Left            =   3465
         TabIndex        =   43
         Text            =   "Text18"
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox txtPermanent 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   180
         TabIndex        =   42
         Text            =   "Text16"
         Top             =   3330
         Width           =   4815
      End
      Begin VB.TextBox txtPresent 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   180
         TabIndex        =   41
         Text            =   "Text15"
         Top             =   2565
         Width           =   4815
      End
      Begin VB.TextBox txtFather 
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
         Left            =   180
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   1215
         Width           =   2700
      End
      Begin VB.TextBox txtMother 
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
         Left            =   180
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   1890
         Width           =   2700
      End
      Begin VB.TextBox txtName 
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
         Left            =   180
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   540
         Width           =   2715
      End
      Begin VB.TextBox txtContact_Home 
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
         Left            =   1755
         TabIndex        =   37
         Text            =   "Text17"
         Top             =   4095
         Width           =   1455
      End
      Begin VB.TextBox txtMobile 
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
         Left            =   3375
         TabIndex        =   36
         Text            =   "Text17"
         Top             =   4095
         Width           =   1590
      End
      Begin VB.ComboBox cmbNationality 
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
         Left            =   3465
         TabIndex        =   35
         Text            =   "Combo3"
         Top             =   1890
         Width           =   1545
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Office"
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
         TabIndex        =   56
         Top             =   3780
         Width           =   1230
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth/Age"
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
         Left            =   3465
         TabIndex        =   55
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
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
         Left            =   3465
         TabIndex        =   54
         Top             =   945
         Width           =   690
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permanent Address"
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
         TabIndex        =   53
         Top             =   3015
         Width           =   1725
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Present Address"
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
         TabIndex        =   52
         Top             =   2295
         Width           =   1455
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fathers Name"
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
         TabIndex        =   51
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mothers Name"
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
         TabIndex        =   50
         Top             =   1620
         Width           =   1275
      End
      Begin VB.Label Label16 
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
         Left            =   180
         TabIndex        =   49
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Home"
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
         Left            =   1755
         TabIndex        =   48
         Top             =   3780
         Width           =   1245
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No."
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
         Left            =   3375
         TabIndex        =   47
         Top             =   3780
         Width           =   930
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality"
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
         Left            =   3465
         TabIndex        =   46
         Top             =   1620
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00C0FFC0&
      Height          =   840
      Left            =   135
      ScaleHeight     =   780
      ScaleWidth      =   5130
      TabIndex        =   24
      Top             =   9315
      Width           =   5190
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   465
         Left            =   2610
         TabIndex        =   28
         Top             =   180
         Width           =   1170
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Save"
         Height          =   465
         Left            =   90
         TabIndex        =   27
         Top             =   180
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   465
         Left            =   3870
         TabIndex        =   26
         Top             =   180
         Width           =   1170
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   465
         Left            =   1350
         TabIndex        =   25
         Top             =   180
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Account Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   135
      TabIndex        =   13
      Top             =   900
      Width           =   5190
      Begin VB.TextBox txtForm 
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
         Left            =   90
         TabIndex        =   83
         Text            =   "Text4"
         Top             =   2925
         Width           =   1545
      End
      Begin VB.TextBox txtInsurance 
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
         Left            =   1800
         TabIndex        =   82
         Text            =   "Text2"
         Top             =   2925
         Width           =   1455
      End
      Begin VB.TextBox txtDown_Payment 
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
         Left            =   3420
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   2925
         Width           =   1545
      End
      Begin VB.TextBox txtD_Date 
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
         Left            =   3465
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   2160
         Width           =   1545
      End
      Begin VB.ComboBox cmbTerm 
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
         Left            =   3465
         TabIndex        =   59
         Text            =   "Combo1"
         Top             =   1395
         Width           =   1545
      End
      Begin VB.TextBox txtSecurity 
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
         Left            =   1845
         TabIndex        =   57
         Text            =   "Text2"
         Top             =   585
         Width           =   1455
      End
      Begin VB.TextBox txtInst_No 
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
         Left            =   1845
         TabIndex        =   32
         Text            =   "Text2"
         Top             =   1395
         Width           =   1455
      End
      Begin VB.TextBox txtInstallment 
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
         Left            =   135
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   1395
         Width           =   1545
      End
      Begin VB.TextBox txtMaturity 
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
         Left            =   1845
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtOpen 
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
         Left            =   135
         TabIndex        =   16
         Text            =   "Text4"
         Top             =   2160
         Width           =   1545
      End
      Begin VB.ComboBox cmbLType 
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
         Left            =   3465
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   585
         Width           =   1545
      End
      Begin VB.TextBox txtID 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000#"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   135
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   585
         Width           =   1545
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Down Payment"
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
         Left            =   3420
         TabIndex        =   86
         Top             =   2610
         Width           =   1320
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Form Fee"
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
         TabIndex        =   85
         Top             =   2610
         Width           =   1320
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance"
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
         Left            =   1800
         TabIndex        =   84
         Top             =   2610
         Width           =   840
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Term"
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
         Left            =   3465
         TabIndex        =   60
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security Deposit"
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
         Left            =   1845
         TabIndex        =   58
         Top             =   270
         Width           =   1440
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Installment"
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
         Left            =   1845
         TabIndex        =   31
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date"
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
         Left            =   1845
         TabIndex        =   23
         Top             =   1845
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Date"
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
         Left            =   135
         TabIndex        =   22
         Top             =   1845
         Width           =   1185
      End
      Begin VB.Label Label1 
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
         Left            =   135
         TabIndex        =   21
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Account"
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
         Left            =   3465
         TabIndex        =   20
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
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
         Left            =   135
         TabIndex        =   19
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1st Due Date"
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
         Left            =   3465
         TabIndex        =   18
         Top             =   1845
         Width           =   1140
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Official Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   5490
      TabIndex        =   0
      Top             =   7110
      Width           =   5550
      Begin VB.ComboBox cmbFO_Code 
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
         TabIndex        =   78
         Text            =   "Combo1"
         Top             =   1440
         Width           =   960
      End
      Begin VB.ComboBox cmbDPO_Code 
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
         TabIndex        =   77
         Text            =   "Combo2"
         Top             =   1980
         Width           =   960
      End
      Begin VB.ComboBox cmbCenter_Code 
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
         TabIndex        =   76
         Text            =   "20101"
         Top             =   345
         Width           =   960
      End
      Begin VB.ComboBox cmbSamity_Code 
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
         TabIndex        =   75
         Text            =   "Combo6"
         Top             =   900
         Width           =   960
      End
      Begin VB.ComboBox cmbAM_Code 
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
         TabIndex        =   74
         Text            =   "Combo2"
         Top             =   2520
         Width           =   960
      End
      Begin VB.ComboBox cmbAM_Name 
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
         Left            =   3555
         TabIndex        =   62
         Text            =   "Combo2"
         Top             =   2520
         Width           =   1860
      End
      Begin VB.ComboBox cmbSamity_name 
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
         Left            =   3555
         TabIndex        =   4
         Text            =   "Combo2"
         Top             =   900
         Width           =   1860
      End
      Begin VB.ComboBox cmbCenter_Name 
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
         Left            =   3555
         TabIndex        =   3
         Text            =   "Combo2"
         Top             =   345
         Width           =   1860
      End
      Begin VB.ComboBox cmbFO_Name 
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
         Left            =   3555
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   1440
         Width           =   1860
      End
      Begin VB.ComboBox cmbDPO_Name 
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
         Left            =   3555
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   1980
         Width           =   1860
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AM Code"
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
         Left            =   375
         TabIndex        =   64
         Top             =   2610
         Width           =   810
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AM Name"
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
         Left            =   2610
         TabIndex        =   63
         Top             =   2565
         Width           =   870
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Samity Code"
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
         TabIndex        =   33
         Top             =   945
         Width           =   1125
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D.F.O"
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
         Left            =   -270
         TabIndex        =   12
         Top             =   2985
         Width           =   135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Samity Name"
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
         Left            =   2295
         TabIndex        =   11
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Center Code"
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
         TabIndex        =   10
         Top             =   405
         Width           =   1080
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Center Name"
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
         Left            =   2340
         TabIndex        =   9
         Top             =   405
         Width           =   1140
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.O. Code"
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
         Left            =   285
         TabIndex        =   8
         Top             =   1485
         Width           =   900
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.O. Name"
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
         Left            =   2520
         TabIndex        =   7
         Top             =   1485
         Width           =   960
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D.P.O. Name"
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
         Left            =   2310
         TabIndex        =   6
         Top             =   2025
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D.P.O. Code"
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
         Left            =   75
         TabIndex        =   5
         Top             =   2025
         Width           =   1110
      End
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOAN ACCOUNT OPENING"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   135
      TabIndex        =   29
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "frmLoanAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Photo As String
Dim Thumb As String
Dim stk As Integer
Dim strloan As String
Dim Prod_Sl As Integer
Dim Gtr_Sl As Integer
Dim netAmnt As Integer
Dim rsGtr As ADODB.Recordset
Dim rsProd As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub L_Type()
        cmbLType.Clear
        cmbLType.AddItem "Product"
        cmbLType.AddItem "Cash"
End Sub
Private Sub grd_Gtr()
Gtr_Sl = 0
    ' Create an initial recordset, just for demonstration purposes,
    ' and assign it to the DataGrid control's DataSource property.
    Set rsGtr = New ADODB.Recordset
    With rsGtr
        .Fields.Append "Sl", adBSTR
        .Fields.Append "Guarantor ID", adBSTR
        .Fields.Append "Guarantor Name", adBSTR
        .Open
   
    End With
    Set grdGtr.DataSource = rsGtr
End Sub
Private Sub Col_Gtr()
With grdGtr
    .Columns(0).Width = 500
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1350
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 3000
    .Columns(2).Alignment = dbgLeft
End With
End Sub
Private Sub Col_Prod()
With grdProd
    .Columns(0).Width = 500
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1500
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1500
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 500
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 850
    .Columns(4).Alignment = dbgRight
End With
End Sub

Private Sub Prod_Name()
        cmbProduct_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_name FROM Prod_master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbProduct_Name.AddItem rs!Prod_Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
        cmbProduct_Name.AddItem "Show All"
End Sub
Private Sub Prod_Model()

        cmbProd_Model.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name, Prod_Model FROM Prod_Master where Prod_Name like '" & cmbProduct_Name.Text & "'"
        rs.Open str, conn
        
        If Not rs.EOF Then
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbProd_Model.AddItem rs!Prod_Model
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
        cmbProd_Model.AddItem "Show All"
End Sub
Private Sub Prod_Price()
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name, Prod_Model, Prod_Price, Sale_Price, Stock FROM Prod_Master where Prod_Name like '" & cmbProduct_Name.Text & "' and Prod_Model like '" & cmbProd_Model.Text & "' and Stock > 0"
        rs.Open str, conn
        
        If Not rs.EOF Then
        txtNet.Text = rs!Sale_Price
        txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
        rs.Close
        Else
        MsgBox "Product not available!", vbCritical, "Error!"
        rs.Close
        Exit Sub
        End If
End Sub
Private Sub Prod_Stock()
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name, Prod_Model, Sale_Price, Stock FROM Prod_Master where Prod_Name like '" & cmbProduct_Name.Text & "' and Prod_Model like '" & cmbProd_Model.Text & "' and Stock >0"
        rs.Open str, conn
        
        If Not rs.EOF Then
            If rs!Stock >= stk Then
            txtNet.Text = rs!Sale_Price * stk
            rs.Close
            Else
            MsgBox "Only " & stk & " Products available!", vbCritical, "Product Error"
            rs.Close
            End If
        Else
            rs.Close
            MsgBox "No Products available!", vbCritical, "Product Error"
        End If
        Exit Sub
End Sub
Private Sub ComboTerm()
        cmbTerm.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Term FROM Parameter"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbTerm.AddItem rs!Term
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
        
End Sub
Private Sub ComboCenter_Code()
        On Error Resume Next
        cmbCenter_Code.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Branch_Code FROM Branch order by Branch_Code"
        rs.Open str, conn
        If Not rs.EOF Then
        
            rs.MoveFirst
            Do While Not rs.EOF
            On Error Resume Next
            cmbCenter_Code.AddItem rs!Branch_Code
            rs.MoveNext
            Loop
        Else
        rs.Close
        End If
        rs.Close
End Sub
Private Sub ComboCenter_Name()
        cmbCenter_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Branch_Name FROM Branch Order by Branch_Name"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbCenter_Name.AddItem rs!Branch_Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
End Sub
Private Sub ComboSamity_Name()
        cmbSamity_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Samity_Name FROM Samity"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbSamity_Name.AddItem rs!Samity_Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
End Sub
Private Sub ComboSamity_Code()
        cmbSamity_Code.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Samity_Code FROM Samity"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbSamity_Code.AddItem rs!Samity_Code
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
End Sub
Private Sub ComboAM_Code()
Dim AM As String
AM = "AM"
        
        cmbDPO_Code.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_Id, Designation FROM Employee Where Designation like '" & AM & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbAM_Code.AddItem rs!Staff_Id
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
End Sub
Private Sub ComboAM_Name()
Dim AM As String
AM = "AM"
        
        cmbDPO_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name, Designation FROM Employee Where Designation like '" & AM & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbAM_Name.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
End Sub
Private Sub ComboFO_Code()
Dim FO As String
FO = "FO"
        cmbFO_Code.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_ID, Designation FROM Employee Where Designation like '" & FO & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbFO_Code.AddItem rs!Staff_Id
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
End Sub
Private Sub cmbG_ID()
        cmbGtr_Id.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer FROM Deposit_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbGtr_Id.AddItem rs!Customer
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub
Private Sub cmbG_Name()
        cmbGtr_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name FROM Deposit_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbGtr_Name.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub

Private Sub ComboFO_Name()
Dim FO As String
FO = "FO"
        cmbFO_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name, Designation FROM Employee Where Designation like '" & FO & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbFO_Name.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub
Private Sub ComboDPO_Code()
Dim DPO As String
DPO = "DPO"
        
        cmbDPO_Code.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_Id, Designation FROM Employee Where Designation like '" & DPO & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbDPO_Code.AddItem rs!Staff_Id
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub
Private Sub ComboDPO_Name()
Dim DPO As String
DPO = "DPO"
        
        cmbDPO_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name, Designation FROM Employee Where Designation like '" & DPO & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbDPO_Name.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub
Private Sub ComboReligion()
cmbReligion.AddItem "Islam"
cmbReligion.AddItem "Hindu"
cmbReligion.AddItem "Christian"
cmbReligion.AddItem "Others"
End Sub
Private Sub ComboNationality()
cmbNationality.AddItem "Bangladeshi"
cmbNationality.AddItem "Others"
End Sub
Private Sub clearTextboxes()
        txtId.Text = ""
        cmbLType.Text = ""
        txtSecurity.Text = ""
        
        txtNet.Text = "0.00"
        txtInstallment.Text = "0.00"
        txtInst_No.Text = "0"
        cmbTerm.Text = ""
        
        txtOpen.Text = ""
        txtMaturity.Text = ""
        txtD_Date.Text = ""
        
        txtForm.Text = "0.00"
        txtInsurance.Text = "0.00"
        txtDown_Payment.Text = "0.00"
        
        txtName.Text = ""
        txtFather.Text = ""
        txtMother.Text = ""
        txtBirthday.Text = ""
        cmbReligion.Text = ""
        cmbNationality.Text = ""
        txtPresent.Text = ""
        txtPermanent.Text = ""
        txtContact_Office.Text = ""
        txtContact_Home.Text = ""
        txtMobile.Text = ""
        
        cmbProduct_Name.Text = ""
        cmbProd_Model.Text = ""
        cmbGtr_Id.Text = ""
        cmbGtr_Name.Text = ""
                       
        cmbCenter_Name.Text = ""
        cmbCenter_Code.Text = ""
        cmbSamity_Code.Text = ""
        cmbSamity_Name.Text = ""
        
        cmbFO_Name.Text = ""
        cmbFO_Code.Text = ""
        cmbDPO_Name.Text = ""
        cmbDPO_Code.Text = ""
        cmbAM_Name.Text = ""
        cmbAM_Code.Text = ""
End Sub

Private Sub cmbGtr_Id_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbGtr_Name.SetFocus
End If
End Sub
Private Sub cmbGtr_Id_LostFocus()
    If cmbGtr_Id.Text = "" Then
        Exit Sub
    End If
    Dim msgsl As Integer
    Dim code As String
    msgsl = 0
    code = cmbGtr_Id.Text
    On Error Resume Next
            Set rsU = New ADODB.Recordset
            str = "SELECT * FROM Loan_Info where G_ID1 like '" & code & "' or G_ID2 like '" & code & "' or G_ID3 like '" & code & "' or G_ID4 like '" & code & "' or G_ID5 like '" & code & "' or G_ID6 like '" & code & "' or G_ID7 like '" & code & "' or G_ID8 like '" & code & "' or G_ID9 like '" & code & "' or G_ID10 like '" & code & "'"
            rsU.Open str, conn
            If Not rsU.EOF Then
            rsU.MoveFirst
            Do While Not rsU.EOF
            msgsl = msgsl + 1
            rsU.MoveNext
            Loop
            rsU.Close
            End If
            
            Set rs = New ADODB.Recordset
            str = "SELECT Customer, Name, Amount FROM Deposit_Master where Customer like '" & code & "'"
            rs.Open str, conn
            
            If Not rs.EOF Then
                If msgsl > 0 Then
                MsgBox "Customer is already Guarantor for: " & msgsl & " Loan and Balance: " & Format$(Val(rs!Amount), "###0.00"), vbInformation, "Guarantor Info!"
                Else
                MsgBox "Guarantor Balance: " & Format$(Val(rs!Amount), "###0.00"), vbInformation, "Guarantor Info!"
                End If
                cmbGtr_Name.Text = rs!Name
                
                rs.Close
                
            End If
            
End Sub

Private Sub cmbAM_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdAdd.SetFocus
End If
End Sub

Private Sub cmbFO_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbDPO_Code.SetFocus
End If
End Sub
Private Sub cmbGarantorID_1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbGarantorName_1.SetFocus
End If
End Sub

Private Sub cmbGarantorID_1_LostFocus()
If cmbGarantorID_1.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbGarantorID_1.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer, Name FROM Personal_Info where Customer like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbGarantorName_1.Text = rs!Name
        rs.Close
        End If
End Sub


Private Sub cmbGarantorID_2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbGarantorName_2.SetFocus
End If
End Sub

Private Sub cmbGarantorID_2_LostFocus()
If cmbGarantorID_2.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbGarantorID_2.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer, Name FROM Personal_Info where Customer like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbGarantorName_2.Text = rs!Name
        rs.Close
        End If
End Sub




Private Sub cmbGarantorID_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbGarantorName_3.SetFocus
End If
End Sub

Private Sub cmbGarantorID_3_LostFocus()
If cmbGarantorID_3.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbGarantorID_3.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer, Name FROM Personal_Info where Customer like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbGarantorName_3.Text = rs!Name
        rs.Close
        End If
End Sub

Private Sub cmbGarantorID_4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbGarantorName_4.SetFocus
End If
End Sub

Private Sub cmbGarantorID_4_LostFocus()
If cmbGarantorID_4.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbGarantorID_4.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer, Name FROM Personal_Info where Customer like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbGarantorName_4.Text = rs!Name
        rs.Close
        End If
End Sub

Private Sub cmbGarantorName_1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbGarantorID_2.SetFocus
End If
End Sub

Private Sub cmbGarantorName_1_LostFocus()
If cmbGarantorName_1.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbGarantorName_1.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer, Name FROM Personal_Info where Name like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbGarantorID_1.Text = rs!Customer
        rs.Close
        End If
End Sub

Private Sub cmbGarantorName_2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbGarantorID_3.SetFocus
End If
End Sub

Private Sub cmbGarantorName_2_LostFocus()
If cmbGarantorName_2.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbGarantorName_2.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer, Name FROM Personal_Info where Name like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbGarantorID_2.Text = rs!Customer
        rs.Close
        End If
End Sub

Private Sub cmbGarantorName_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbGarantorID_4.SetFocus
End If
End Sub

Private Sub cmbGarantorName_3_LostFocus()
If cmbGarantorName_3.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbGarantorName_3.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer, Name FROM Personal_Info where Name like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbGarantorID_3.Text = rs!Customer
        rs.Close
        End If
End Sub

Private Sub cmbGarantorName_4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbCenter_Code.SetFocus
End If
End Sub

Private Sub cmbGarantorName_4_LostFocus()
If cmbGarantorName_4.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbGarantorName_4.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer, Name FROM Personal_Info where Name like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbGarantorID_4.Text = rs!Customer
        rs.Close
        End If
End Sub

Private Sub cmbGtr_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmbGtr_Name.Text = "" Then
        Exit Sub
    End If
   cmbCenter_Code.SetFocus
End If
End Sub
Private Sub cmbGtr_Name_LostFocus()
If cmbGtr_Id.Text = "" Then
Exit Sub
End If
If cmbGtr_Name.Text = "" Then
Exit Sub
End If
Gtr_Sl = Gtr_Sl + 1
If Val(Gtr_Sl) > 10 Then
    MsgBox "No more Guarantor Allowed!", vbCritical, "Error!"
    Exit Sub
Else
    rsGtr.AddNew
    rsGtr!sl = Gtr_Sl
    rsGtr![Guarantor ID] = cmbGtr_Id.Text
    rsGtr![Guarantor Name] = cmbGtr_Name.Text
        
    Set grdGtr.DataSource = rsGtr
    grdGtr.Refresh
    Call Col_Gtr
End If
End Sub

Private Sub cmbLType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmbLType.Text = "Cash" Then
netAmnt = 0
txtNet.SetFocus
End If
If cmbLType.Text = "Product" Then
cmbProduct_Name.Enabled = True
cmbProduct_Name.SetFocus
Call grd_Prod
Call Col_Prod
End If
End If
End Sub

Private Sub cmbLType_LostFocus()
    If cmbLType.Text = "Cash" Then
        cmbProduct_Name.Enabled = False
        cmbProd_Model.Enabled = False
        Label18.Caption = "Loan Amount"
    End If
    If cmbLType.Text = "Product" Then
        cmbProduct_Name.Enabled = True
        cmbProd_Model.Enabled = True
        Label18.Caption = "Product Price"
    End If
End Sub

Private Sub cmbProd_Model_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    txtQty.Text = "1"
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty.Text)
    txtQty.SetFocus

End If
End Sub

Private Sub cmbProd_Model_LostFocus()
Call Prod_Price
End Sub

Private Sub cmbProduct_Name_Click()
If cmbProduct_Name.Text = "" Then
Exit Sub
End If

If cmbProduct_Name.Text = "Show All" Then

Set rs = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master Order by Prod_Name"
        rs.Open str, conn
        
        frmProd_Search.Adodc1.ConnectionString = cnStr
        frmProd_Search.Adodc1.RecordSource = str
        frmProd_Search.Adodc1.Refresh
        rs.Close
frmProd_Search.Show 1
End If
End Sub

Private Sub cmbProduct_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProd_Model.SetFocus
End If
End Sub

Private Sub cmbProduct_Name_LostFocus()
Call Prod_Model
End Sub


Private Sub cmbProduct_Sl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbGarantorID_1.SetFocus
End If
End Sub

Private Sub cmbProduct_Sl_LostFocus()
Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where Prod_Sl like '" & cmbProduct_Sl.Text & "'"
        rs.Open str, conn
        
        If Not rs.EOF Then
            cmbProduct_Name.Text = rs!Prod_Name
            cmbProd_Model.Text = rs!Prod_Model
            cmbProduct_Sl.Text = rs!Prod_Sl
            
            txtPrice.Text = rs!Prod_Price
            rs.Close
        End If
End Sub

Private Sub cmbTerm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtOpen.SetFocus
End If
End Sub

Private Sub cmbUnion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbFO_Code.SetFocus
End If
End Sub

Private Sub cmbTerm_LostFocus()
Dim Due_Date As Date
Due_Date = CDate(txtOpen.Text)
If cmbTerm.Text = "Daily" Then
txtD_Date.Text = Due_Date + 1
End If

If cmbTerm.Text = "Weekly" Then
txtD_Date.Text = Due_Date + 7
End If

If cmbTerm.Text = "Monthly" Then
txtD_Date.Text = Due_Date + 30
End If
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
    Dim suf As String
    Dim Prif As String
    suf = 201

    If cmbCenter_Code.Text = "" Then
        MsgBox "Please Input Center Code", vbCritical
        cmbCenter_Code.SetFocus
        Exit Sub
    Else
        Prif = cmbCenter_Code.Text
    End If
    
    
    Set rs = New ADODB.Recordset
        str = "select * from Loan_Info where Customer like '" & txtId.Text & "'"
        rs.Open str, conn

If Not rs.EOF Then
    
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        suf = suf + 1
        rs.MoveNext
        Loop
        rs.Close
    
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Loan_info", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Customer = txtId.Text
        rsN!AC_No = Prif + suf + txtId.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Open_Date = txtOpen.Text
        rsN!Mat_Date = txtMaturity.Text
        rsN!D_Date = txtD_Date.Text
        
        rsN!Security = txtSecurity.Text
        rsN!Form = txtForm.Text
        rsN!Insurance = txtInsurance.Text
        rsN!Down_Payment = txtDown_Payment.Text
        rsN!Net_Loan = Val(netAmnt)
                
        rsN!Name = txtName.Text
        rsN!Fathers_Name = txtFather.Text
        rsN!Mothers_Name = txtMother.Text
        rsN!Age = txtBirthday.Text
        rsN!Religion = cmbReligion.Text
        rsN!Nationality = cmbNationality.Text
        rsN!Present_Address = txtPresent.Text
        rsN!Permanent_Address = txtPermanent.Text
        rsN!Contact_Office = txtContact_Office.Text
        rsN!Contact_Home = txtContact_Home.Text
        rsN!Mobile = txtMobile.Text
        rsN!Photo = Photo
        rsN!Thumb = Thumb
        
        rsN!G_ID1 = 0
        rsN!G_ID2 = 0
        rsN!G_ID3 = 0
        rsN!G_ID4 = 0
        rsN!G_ID5 = 0
        rsN!G_ID6 = 0
        rsN!G_ID7 = 0
        rsN!G_ID8 = 0
        rsN!G_ID9 = 0
        rsN!G_ID10 = 0
        
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        rsN!Samity_Name = cmbSamity_Name.Text
        rsN!Samity_Code = cmbSamity_Code.Text
        rsN!FO_Name = cmbFO_Name.Text
        rsN!FO_Code = cmbFO_Code.Text
        rsN!DPO_Name = cmbDPO_Name.Text
        rsN!DPO_Code = cmbDPO_Code.Text
        rs!AM_Name = cmbAM_Name.Text
        rs!AM_Code = cmbAM_Code.Text
        
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_Model.Text
                       
    rsGtr.MoveFirst
        
        Do While Not rsGtr.EOF
            If rsN!G_ID1 = 0 Then
                rsN!G_ID1 = rsGtr![Guarantor ID]
                rsN!G_Name1 = rsGtr![Guarantor Name]
                Else
                If rsN!G_ID2 = 0 Then
                    rsN!G_ID2 = rsGtr![Guarantor ID]
                    rsN!G_Name2 = rsGtr![Guarantor Name]
                    Else
                    If rsN!G_ID3 = 0 Then
                        rsN!G_ID3 = rsGtr![Guarantor ID]
                        rsN!G_Name3 = rsGtr![Guarantor Name]
                        Else
                        If rsN!G_ID4 = 0 Then
                            rsN!G_ID4 = rsGtr![Guarantor ID]
                            rsN!G_Name4 = rsGtr![Guarantor Name]
                            Else
                            If rsN!G_ID5 = 0 Then
                                rsN!G_ID5 = rsGtr![Guarantor ID]
                                rsN!G_Name5 = rsGtr![Guarantor Name]
                                Else
                                If rsN!G_ID6 = 0 Then
                                    rsN!G_ID6 = rsGtr![Guarantor ID]
                                    rsN!G_Name6 = rsGtr![Guarantor Name]
                                    Else
                                    If rsN!G_ID7 = 0 Then
                                        rsN!G_ID7 = rsGtr![Guarantor ID]
                                        rsN!G_Name7 = rsGtr![Guarantor Name]
                                        Else
                                        If rsN!G_ID8 = 0 Then
                                            rsN!G_ID8 = rsGtr![Guarantor ID]
                                            rsN!G_Name8 = rsGtr![Guarantor Name]
                                            Else
                                            If rsN!G_ID9 = 0 Then
                                                rsN!G_ID9 = rsGtr![Guarantor ID]
                                                rsN!G_Name9 = rsGtr![Guarantor Name]
                                                Else
                                                If rsN!G_ID10 = 0 Then
                                                    rsN!G_ID10 = rsGtr![Guarantor ID]
                                                    rsN!G_Name10 = rsGtr![Guarantor Name]
                                                    Else
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        rsGtr.MoveNext
        Loop
        
        
        rsN.Update
        rsN.Close
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Loan_Master", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Open_Date = txtOpen.Text
        rsN!D_Date = txtD_Date.Text
        rsN!Inst_Date = ""
        rsN!Mat_Date = txtMaturity.Text
        rsN!AC_No = Prif + suf + txtId.Text
        rsN!Customer = txtId.Text
        rsN!Name = txtName.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_Model.Text
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Inst_Due = txtInst_No.Text
        rsN!Inst_Paid = 0
        rsN!Amount = Val(netAmnt)
        rsN!Balance = -Val(netAmnt) + Val(txtDown_Payment.Text)
        rsN!Due = 0
        rsN!Paid = Val(txtDown_Payment.Text)
        rsN!Advance = 0
        rsN!Fine = 0
        rsN!Term_Fail = 0
        rsN!Daily_Pay = 0
        rsN!Weekly_Pay = 0
        rsN!Monthly_Pay = 0
        rsN!Yearly_Pay = 0
        
        rsN!Daily_Draw = 0
        rsN!Weekly_Draw = 0
        rsN!Monthly_Draw = 0
        rsN!Yearly_Draw = 0
        
        rsN!Daily_Bal = 0
        rsN!Weekly_Bal = 0
        rsN!Monthly_Bal = 0
        rsN!Yearly_Bal = 0
        
        
        rsN!Week_1 = 0
        rsN!Week_2 = 0
        rsN!Week_3 = 0
        rsN!Week_4 = 0
        rsN!Week_5 = 0
        
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        rsN!Samity_Code = cmbSamity_Code.Text
        rsN!Samity_Name = cmbSamity_Name.Text
        rsN!FO_Name = cmbFO_Name.Text
        rsN!FO_Code = cmbFO_Code.Text
        rsN!DPO_Name = cmbDPO_Name.Text
        rsN!DPO_Code = cmbDPO_Code.Text
        rsN!AM_Name = cmbAM_Name.Text
        rsN!AM_Code = cmbAM_Code.Text
        rsN!C_lose = "No"
        rsN.Update
        rsN.Close
          
        Set rsN = New ADODB.Recordset
            rsN.Open "Loan_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = Prif + suf + txtId.Text
            rsN!MR_No = txtMR_No.Text
            rsN!Name = txtName.Text
            rsN!Description = "Loan Disbursement"
            rsN!Dr = Val(netAmnt)
            rsN!Fine = ""
            rsN!Balance = -Val(netAmnt)
            rsN.Update
            rsN.Close
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Loan_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = Prif + suf + txtId.Text
            rsN!MR_No = txtMR_No.Text
            rsN!Name = txtName.Text
            rsN!Description = "Downpayment Realized"
            rsN!Cr = Val(txtDown_Payment.Text)
            rsN!Fine = ""
            rsN!Balance = -Val(netAmnt) + Val(txtDown_Payment.Text)
            rsN.Update
            rsN.Close
'-----------------------------------------------------------------------------------
If cmbLType.Text = "Product" Then
    
    rsProd.MoveFirst
    
Do While Not rsProd.EOF
        
    Set rsU = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Name like '" & rsProd![Product Name] & "' and Prod_Model like '" & rsProd![Product Model] & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic

    If Not rsU.EOF Then
            rsU!Sale = rsU!Sale + Val(rsProd!Qty)
            rsU!Stock = rsU!Stock - Val(rsProd!Qty)
            rsU!Amount = rsU!Stock * rsU!Prod_Price
            rsU.Update
    
        Set rsN = New ADODB.Recordset
            rsN.Open "Prod_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            
            rsN!Invo_No = Prif + suf + txtId.Text
            rsN!Prod_Name = rsProd![Product Name]
            rsN!Prod_Model = rsProd![Product Model]
            rsN!Prod_Code = rsU!Prod_Code
            rsN!Purchase = 0
            rsN!Sale = Val(rsProd!Qty)
            rsN!Stock = rsU!Stock
            rsN!Date = txtOpen.Text
            rsN!Prod_Price = rsU!Prod_Price
            rsN!Sale_Price = Val(rsProd!Amount)
            rsN!Amount = Val(rsProd!Amount) * Val(rsProd!Qty)
                   
            rsN.Update
            rsN.Close

'DR-------------------------------------------------------------------
        Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100107 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(rsProd!Amount)
        rs.Update
        
        Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtOpen.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C-" & Prif + suf + txtId.Text
        rsN!Dr = Val(rsProd!Amount)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
            
'Cr----------------------------------------------------------------------
        Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100114 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(rsU!Prod_Price)
            rs!Date = txtOpen.Text
            rs.Update
    
        Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = 0
            rsN!Cr = Val(rsU!Prod_Price)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100115 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + (Val(rsProd!Amount) - Val(rsU!Prod_Price))
        rs!Date = txtOpen.Text
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtOpen.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C-" & Prif + suf + txtId.Text
        rsN!Dr = 0
        rsN!Cr = (Val(rsProd!Amount) - Val(rsU!Prod_Price))
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
 

'------------------------------------------------------------
   End If
   rsProd.MoveNext
Loop
        
    Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100107 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(txtDown_Payment.Text)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtDown_Payment.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtDown_Payment.Text)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = Val(txtDown_Payment.Text)
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            
            
            Set rsN = New ADODB.Recordset
            rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!Description = "A/C: " + Prif + suf + txtId.Text
            rsN!Dr = Val(txtDown_Payment.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rsU = New ADODB.Recordset
            str = "select * from Others"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
            rsU!Cash_Dr = rsU!Cash_Dr + Val(txtDown_Payment.Text)
            rsU!Cash_Close = rsU!Cash_Close + Val(txtDown_Payment.Text)
            rsU.Update
            rsU.Close

End If
        
If cmbLType.Text = "Cash" Then
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100108 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(netAmnt)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = Val(netAmnt)
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(netAmnt)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = 0
            rsN!Cr = Val(netAmnt)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            
            
            Set rsN = New ADODB.Recordset
            rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!Description = "A/C-" + Prif + suf + txtId.Text
            rsN!Cr = Val(netAmnt)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rsU = New ADODB.Recordset
            str = "select * from Others"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
            rsU!Cash_Cr = rsU!Cash_Cr + Val(netAmnt)
            rsU!Cash_Close = rsU!Cash_Close - Val(netAmnt)
            rsU.Update
            rsU.Close
        End If
 
Else
'---------------------------------------------------------------------
    Set rsN = New ADODB.Recordset
    rsN.Open "Loan_info", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Customer = txtId.Text
        rsN!AC_No = Prif + suf + txtId.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Open_Date = txtOpen.Text
        rsN!Mat_Date = txtMaturity.Text
        rsN!D_Date = txtD_Date.Text
        
        rsN!Security = txtSecurity.Text
        rsN!Form = txtForm.Text
        rsN!Insurance = txtInsurance.Text
        rsN!Down_Payment = txtDown_Payment.Text
        rsN!Net_Loan = Val(netAmnt)
        
        rsN!Name = txtName.Text
        rsN!Fathers_Name = txtFather.Text
        rsN!Mothers_Name = txtMother.Text
        rsN!Age = txtBirthday.Text
        rsN!Religion = cmbReligion.Text
        rsN!Nationality = cmbNationality.Text
        rsN!Present_Address = txtPresent.Text
        rsN!Permanent_Address = txtPermanent.Text
        rsN!Contact_Office = txtContact_Office.Text
        rsN!Contact_Home = txtContact_Home.Text
        rsN!Mobile = txtMobile.Text
        rsN!Photo = Photo
        rsN!Thumb = Thumb
        
        rsN!G_ID1 = 0
        rsN!G_ID2 = 0
        rsN!G_ID3 = 0
        rsN!G_ID4 = 0
        rsN!G_ID5 = 0
        rsN!G_ID6 = 0
        rsN!G_ID7 = 0
        rsN!G_ID8 = 0
        rsN!G_ID9 = 0
        rsN!G_ID10 = 0
        
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        rsN!Samity_Name = cmbSamity_Name.Text
        rsN!Samity_Code = cmbSamity_Code.Text
        rsN!FO_Name = cmbFO_Name.Text
        rsN!FO_Code = cmbFO_Code.Text
        rsN!DPO_Name = cmbDPO_Name.Text
        rsN!DPO_Code = cmbDPO_Code.Text
        rs!AM_Name = cmbAM_Name.Text
        rs!AM_Code = cmbAM_Code.Text
        
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_Model.Text
            
        rsGtr.MoveFirst
        
        Do While Not rsGtr.EOF
            If rsN!G_ID1 = 0 Then
                rsN!G_ID1 = rsGtr![Guarantor ID]
                rsN!G_Name1 = rsGtr![Guarantor Name]
                Else
                If rsN!G_ID2 = 0 Then
                    rsN!G_ID2 = rsGtr![Guarantor ID]
                    rsN!G_Name2 = rsGtr![Guarantor Name]
                    Else
                    If rsN!G_ID3 = 0 Then
                        rsN!G_ID3 = rsGtr![Guarantor ID]
                        rsN!G_Name3 = rsGtr![Guarantor Name]
                        Else
                        If rsN!G_ID4 = 0 Then
                            rsN!G_ID4 = rsGtr![Guarantor ID]
                            rsN!G_Name4 = rsGtr![Guarantor Name]
                            Else
                            If rsN!G_ID5 = 0 Then
                                rsN!G_ID5 = rsGtr![Guarantor ID]
                                rsN!G_Name5 = rsGtr![Guarantor Name]
                                Else
                                If rsN!G_ID6 = 0 Then
                                    rsN!G_ID6 = rsGtr![Guarantor ID]
                                    rsN!G_Name6 = rsGtr![Guarantor Name]
                                    Else
                                    If rsN!G_ID7 = 0 Then
                                        rsN!G_ID7 = rsGtr![Guarantor ID]
                                        rsN!G_Name7 = rsGtr![Guarantor Name]
                                        Else
                                        If rsN!G_ID8 = 0 Then
                                            rsN!G_ID8 = rsGtr![Guarantor ID]
                                            rsN!G_Name8 = rsGtr![Guarantor Name]
                                            Else
                                            If rsN!G_ID9 = 0 Then
                                                rsN!G_ID9 = rsGtr![Guarantor ID]
                                                rsN!G_Name9 = rsGtr![Guarantor Name]
                                                Else
                                                If rsN!G_ID10 = 0 Then
                                                    rsN!G_ID10 = rsGtr![Guarantor ID]
                                                    rsN!G_Name10 = rsGtr![Guarantor Name]
                                                    Else
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        rsGtr.MoveNext
        Loop
           
        rsN.Update
        rsN.Close
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Loan_Master", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Open_Date = txtOpen.Text
        rsN!D_Date = txtD_Date.Text
        rsN!Inst_Date = ""
        rsN!Mat_Date = txtMaturity.Text
        rsN!AC_No = Prif + suf + txtId.Text
        rsN!Customer = txtId.Text
        rsN!Name = txtName.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_Model.Text
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Inst_Due = txtInst_No.Text
        rsN!Inst_Paid = 0
        rsN!Amount = Val(netAmnt)
        rsN!Balance = -Val(netAmnt) + Val(txtDown_Payment.Text)
        rsN!Due = 0
        rsN!Paid = Val(txtDown_Payment.Text)
        rsN!Advance = 0
        rsN!Fine = 0
        rsN!Term_Fail = 0
        
        rsN!Daily_Pay = 0
        rsN!Weekly_Pay = 0
        rsN!Monthly_Pay = 0
        rsN!Yearly_Pay = 0
        
        rsN!Daily_Draw = 0
        rsN!Weekly_Draw = 0
        rsN!Monthly_Draw = 0
        rsN!Yearly_Draw = 0
        
        rsN!Daily_Bal = 0
        rsN!Weekly_Bal = 0
        rsN!Monthly_Bal = 0
        rsN!Yearly_Bal = 0
        
        
        rsN!Week_1 = 0
        rsN!Week_2 = 0
        rsN!Week_3 = 0
        rsN!Week_4 = 0
        rsN!Week_5 = 0
        
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        rsN!Samity_Code = cmbSamity_Code.Text
        rsN!Samity_Name = cmbSamity_Name.Text
        rsN!FO_Name = cmbFO_Name.Text
        rsN!FO_Code = cmbFO_Code.Text
        rsN!DPO_Name = cmbDPO_Name.Text
        rsN!DPO_Code = cmbDPO_Code.Text
        rsN!AM_Name = cmbAM_Name.Text
        rsN!AM_Code = cmbAM_Code.Text
        rsN!C_lose = "No"
        rsN.Update
        rsN.Close
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Loan_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = Prif + suf + txtId.Text
            rsN!MR_No = txtMR_No.Text
            rsN!Name = txtName.Text
            rsN!Description = "Loan Disbursement"
            rsN!Dr = Val(netAmnt)
            rsN!Fine = ""
            rsN!Balance = -Val(netAmnt)
            rsN.Update
            rsN.Close
            
        Set rsN = New ADODB.Recordset
            rsN.Open "Loan_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = Prif + suf + txtId.Text
            rsN!MR_No = txtMR_No.Text
            rsN!Name = txtName.Text
            rsN!Description = "Downpayment Realized"
            rsN!Cr = Val(txtDown_Payment.Text)
            rsN!Fine = ""
            rsN!Balance = -Val(netAmnt) + Val(txtDown_Payment.Text)
            rsN.Update
            rsN.Close


    If cmbLType.Text = "Product" Then
    
        rsProd.MoveFirst
    
            Do While Not rsProd.EOF
        
            Set rsU = New ADODB.Recordset
                str = "select * from Prod_Master where Prod_Name like '" & rsProd![Product Name] & "' and Prod_Model like '" & rsProd![Product Model] & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
            If Not rsU.EOF Then
                    rsU!Sale = rsU!Sale + Val(rsProd!Qty)
                    rsU!Stock = rsU!Stock - Val(rsProd!Qty)
                    rsU!Amount = rsU!Stock * rsU!Prod_Price
                    rsU.Update
        
            Set rsN = New ADODB.Recordset
                rsN.Open "Prod_Tran", conn, adOpenDynamic, adLockOptimistic, -1
                rsN.AddNew
                
                rsN!Invo_No = Prif + suf + txtId.Text
                rsN!Prod_Name = rsProd![Product Name]
                rsN!Prod_Model = rsProd![Product Model]
                rsN!Prod_Code = rsU!Prod_Code
                rsN!Purchase = 0
                rsN!Sale = Val(rsProd!Qty)
                rsN!Stock = rsU!Stock
                rsN!Date = txtOpen.Text
                rsN!Prod_Price = rsU!Prod_Price
                rsN!Sale_Price = Val(rsProd!Amount)
                rsN!Amount = Val(rsProd!Amount) * Val(rsProd!Qty)
                       
                rsN.Update
                rsN.Close

'DR-------------------------------------------------------------------
            Set rs = New ADODB.Recordset
                str = "select * from GL_Master where AC_No like '" & 100107 & "'"
                rs.Open str, conn, adOpenDynamic, adLockOptimistic
                rs!Balance = rs!Balance + Val(rsProd!Amount)
                rs.Update
            
            Set rsN = New ADODB.Recordset
                rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
                rsN.AddNew
                rsN!Date = txtOpen.Text
                rsN!AC_No = rs!AC_No
                rsN!Name = rs!Head_Name
                rsN!Description = "A/C-" & Prif + suf + txtId.Text
                rsN!Dr = Val(rsProd!Amount)
                rsN!Cr = 0
                rsN!Balance = rs!Balance
                rsN.Update
                rsN.Close
                rs.Close
            
'Cr----------------------------------------------------------------------
            Set rs = New ADODB.Recordset
                str = "select * from GL_Master where AC_No like '" & 100114 & "'"
                rs.Open str, conn, adOpenDynamic, adLockOptimistic
                rs!Balance = rs!Balance - Val(rsU!Prod_Price)
                rs!Date = txtOpen.Text
                rs.Update
        
            Set rsN = New ADODB.Recordset
                rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
                rsN.AddNew
                rsN!Date = txtOpen.Text
                rsN!AC_No = rs!AC_No
                rsN!Name = rs!Head_Name
                rsN!Description = "A/C-" & Prif + suf + txtId.Text
                rsN!Dr = 0
                rsN!Cr = Val(rsU!Prod_Price)
                rsN!Balance = rs!Balance
                rsN.Update
                rsN.Close
                rs.Close

            Set rs = New ADODB.Recordset
                str = "select * from GL_Master where AC_No like '" & 100115 & "'"
                rs.Open str, conn, adOpenDynamic, adLockOptimistic
                rs!Balance = rs!Balance + (Val(rsProd!Amount) - Val(rsU!Prod_Price))
                rs!Date = txtOpen.Text
                rs.Update
            
            Set rsN = New ADODB.Recordset
                rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
                rsN.AddNew
                rsN!Date = txtOpen.Text
                rsN!AC_No = rs!AC_No
                rsN!Name = rs!Head_Name
                rsN!Description = "A/C-" & Prif + suf + txtId.Text
                rsN!Dr = 0
                rsN!Cr = (Val(rsProd!Amount) - Val(rsU!Prod_Price))
                rsN!Balance = rs!Balance
                rsN.Update
                rsN.Close
                rs.Close
 

'------------------------------------------------------------
            End If
            rsProd.MoveNext
            Loop
        
        Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100107 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(txtDown_Payment.Text)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtDown_Payment.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtDown_Payment.Text)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = Val(txtDown_Payment.Text)
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            
            
            Set rsN = New ADODB.Recordset
            rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!Description = "A/C: " + Prif + suf + txtId.Text
            rsN!Dr = Val(txtDown_Payment.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rsU = New ADODB.Recordset
            str = "select * from Others"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
            rsU!Cash_Dr = rsU!Cash_Dr + Val(txtDown_Payment.Text)
            rsU!Cash_Close = rsU!Cash_Close + Val(txtDown_Payment.Text)
            rsU.Update
            rsU.Close

    End If
        
    If cmbLType.Text = "Cash" Then
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100108 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(netAmnt)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = Val(netAmnt)
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(netAmnt)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = 0
            rsN!Cr = Val(netAmnt)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            
            
            Set rsN = New ADODB.Recordset
            rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!Description = "A/C-" + Prif + suf + txtId.Text
            rsN!Cr = Val(netAmnt)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rsU = New ADODB.Recordset
            str = "select * from Others"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
            rsU!Cash_Cr = rsU!Cash_Cr + Val(netAmnt)
            rsU!Cash_Close = rsU!Cash_Close - Val(netAmnt)
            rsU.Update
            rsU.Close
    End If

        
End If
            
    Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100135 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtForm.Text)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtForm.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtForm.Text)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = Val(txtForm.Text)
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            
            Set rsN = New ADODB.Recordset
            rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!Description = "A/C-" + Prif + suf + txtId.Text
            rsN!Dr = Val(txtForm.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rsU = New ADODB.Recordset
            str = "select * from Others"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
            rsU!Cash_Dr = rsU!Cash_Dr + Val(txtForm.Text)
            rsU!Cash_Close = rsU!Cash_Close + Val(txtForm.Text)
            rsU.Update
            rsU.Close
        
            
    Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100139 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtInsurance.Text)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtInsurance.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtInsurance.Text)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C-" & Prif + suf + txtId.Text
            rsN!Dr = Val(txtInsurance.Text)
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            
            Set rsN = New ADODB.Recordset
            rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtOpen.Text
            rsN!Description = "A/C-" + Prif + suf + txtId.Text
            rsN!Dr = Val(txtInsurance.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
            
            Set rsU = New ADODB.Recordset
            str = "select * from Others"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
            rsU!Cash_Dr = rsU!Cash_Dr + Val(txtInsurance.Text)
            rsU!Cash_Close = rsU!Cash_Close + Val(txtInsurance.Text)
            rsU.Update
            rsU.Close
            
                
    cmdAdd.Enabled = False
    
    MsgBox "Account Open Successfull! Your Account no.: " & Prif + suf + txtId.Text, vbInformation, "Loan Info"
    Call clearTextboxes
    Call grd_Prod
    Call Col_Prod
    Call grd_Gtr
    Call Col_Gtr
    Exit Sub
End Sub

Private Sub cmdPrint_Click()
On Error Resume Next
If txtId.Text = "" Then
Exit Sub
Else

Set rs = New ADODB.Recordset
    str = "select * from Loan_Info where AC_No like '" & txtId.Text & "'"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        rs.Close
        rptLoan_Info.lblBranch.Caption = Branch_Name & " Branch, " & Branch_Address
        rptLoan_Info.lblUser_Id.Caption = User_Id
        rptLoan_Info.lblUser_Name.Caption = User_Name
        rptLoan_Info.rsLoan_info.ConnectionString = cnStr
        rptLoan_Info.rsLoan_info.Source = str
    Else
    MsgBox "No such account Found", vbCritical, "Error!"
    End If
    
    rptLoan_Info.Show 1
    cmdPrint.Enabled = False
    cmdUpdate.Enabled = False
End If
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
    
  Set rsN = New ADODB.Recordset
        str = "select * from Loan_Info where AC_No like '" & txtId.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsN.EOF Then
        
        rsN!AC_No = txtId.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Open_Date = txtOpen.Text
        rsN!Mat_Date = txtMaturity.Text
        rsN!D_Date = txtD_Date.Text
        
        rsN!Security = txtSecurity.Text
        rsN!Form = txtForm.Text
        rsN!Insurance = txtInsurance.Text
        rsN!Down_Payment = txtDown_Payment.Text
        rsN!Net_Loan = Val(netAmnt)
                
        rsN!Name = txtName.Text
        rsN!Fathers_Name = txtFather.Text
        rsN!Mothers_Name = txtMother.Text
        rsN!Age = txtBirthday.Text
        rsN!Religion = cmbReligion.Text
        rsN!Nationality = cmbNationality.Text
        rsN!Present_Address = txtPresent.Text
        rsN!Permanent_Address = txtPermanent.Text
        rsN!Contact_Office = txtContact_Office.Text
        rsN!Contact_Home = txtContact_Home.Text
        rsN!Mobile = txtMobile.Text
        rsN!Photo = Photo
        rsN!Thumb = Thumb
        
        rsN!G_ID1 = 0
        rsN!G_ID2 = 0
        rsN!G_ID3 = 0
        rsN!G_ID4 = 0
        rsN!G_ID5 = 0
        rsN!G_ID6 = 0
        rsN!G_ID7 = 0
        rsN!G_ID8 = 0
        rsN!G_ID9 = 0
        rsN!G_ID10 = 0
        
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        rsN!Samity_Name = cmbSamity_Name.Text
        rsN!Samity_Code = cmbSamity_Code.Text
        rsN!FO_Name = cmbFO_Name.Text
        rsN!FO_Code = cmbFO_Code.Text
        rsN!DPO_Name = cmbDPO_Name.Text
        rsN!DPO_Code = cmbDPO_Code.Text
        rs!AM_Name = cmbAM_Name.Text
        rs!AM_Code = cmbAM_Code.Text
        
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_Model.Text
                       
    rsGtr.MoveFirst
        
        Do While Not rsGtr.EOF
            If rsN!G_ID1 = 0 Then
                rsN!G_ID1 = rsGtr![Guarantor ID]
                rsN!G_Name1 = rsGtr![Guarantor Name]
                Else
                If rsN!G_ID2 = 0 Then
                    rsN!G_ID2 = rsGtr![Guarantor ID]
                    rsN!G_Name2 = rsGtr![Guarantor Name]
                    Else
                    If rsN!G_ID3 = 0 Then
                        rsN!G_ID3 = rsGtr![Guarantor ID]
                        rsN!G_Name3 = rsGtr![Guarantor Name]
                        Else
                        If rsN!G_ID4 = 0 Then
                            rsN!G_ID4 = rsGtr![Guarantor ID]
                            rsN!G_Name4 = rsGtr![Guarantor Name]
                            Else
                            If rsN!G_ID5 = 0 Then
                                rsN!G_ID5 = rsGtr![Guarantor ID]
                                rsN!G_Name5 = rsGtr![Guarantor Name]
                                Else
                                If rsN!G_ID6 = 0 Then
                                    rsN!G_ID6 = rsGtr![Guarantor ID]
                                    rsN!G_Name6 = rsGtr![Guarantor Name]
                                    Else
                                    If rsN!G_ID7 = 0 Then
                                        rsN!G_ID7 = rsGtr![Guarantor ID]
                                        rsN!G_Name7 = rsGtr![Guarantor Name]
                                        Else
                                        If rsN!G_ID8 = 0 Then
                                            rsN!G_ID8 = rsGtr![Guarantor ID]
                                            rsN!G_Name8 = rsGtr![Guarantor Name]
                                            Else
                                            If rsN!G_ID9 = 0 Then
                                                rsN!G_ID9 = rsGtr![Guarantor ID]
                                                rsN!G_Name9 = rsGtr![Guarantor Name]
                                                Else
                                                If rsN!G_ID10 = 0 Then
                                                    rsN!G_ID10 = rsGtr![Guarantor ID]
                                                    rsN!G_Name10 = rsGtr![Guarantor Name]
                                                    Else
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        rsGtr.MoveNext
        Loop
        
        
        rsN.Update
        rsN.Close
       
    Set rsN = New ADODB.Recordset
        str = "select * from Loan_Master where Customer like '" & txtId.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
        rsN!Open_Date = txtOpen.Text
        rsN!D_Date = txtD_Date.Text
        rsN!Mat_Date = txtMaturity.Text
        rsN!AC_No = txtId.Text
        rsN!Customer = txtId.Text
        rsN!Name = txtName.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_Model.Text
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Security = Val(txtDown_Payment.Text)
        rsN.Update
        rsN.Close
        
    MsgBox "Account Update Successfull! ", vbInformation, "Personal Info"
    Call clearTextboxes
    Call grd_Prod
    Call Col_Prod
    Call grd_Gtr
    Call Col_Gtr
    
    cmdUpdate.Enabled = False
    cmdPrint.Enabled = False
    Else
    rsN.Close
    MsgBox "Invalid Customer ID", vbCritical, "Error"
    
    cmdUpdate.Enabled = False
    cmdPrint.Enabled = False
    End If
    
    Exit Sub
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
        Call L_Type
        Call ComboTerm
        Call ComboReligion
        Call ComboNationality
        Call ComboCenter_Code
        Call ComboCenter_Name
        Call ComboFO_Code
        Call ComboFO_Name
        Call ComboDPO_Code
        Call ComboDPO_Name
        Call ComboAM_Code
        Call ComboAM_Name
        Call ComboSamity_Name
        Call ComboSamity_Code
        Call Prod_Name
        Call cmbG_ID
        Call cmbG_Name
        Call grd_Gtr
        Call Col_Gtr
        Call clearTextboxes
            txtOpen.Text = Today
'-----------------------------------------
cmbProduct_Name.Enabled = False
cmbProd_Model.Enabled = False
'--------------------------------------------
        
        cmdAdd.Enabled = False
        'cmdDelete.Enabled = False
        cmdPrint.Enabled = False
        cmdUpdate.Enabled = False
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDown.SetFocus
End If
End Sub

Private Sub Label63_Click()

End Sub

Private Sub txtBirthday_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbReligion.SetFocus
End If
End Sub

Private Sub txtContact_Home_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMobile.SetFocus
End If
End Sub

Private Sub txtContact_Office_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtContact_Home.SetFocus
End If
End Sub

Private Sub txtD_Date_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtForm.SelStart = 0
   txtForm.Text = Format$(Val(txtForm.Text), "###0.00")
   txtForm.SelLength = Len(txtForm.Text)
   txtForm.SetFocus
End If
End Sub

Private Sub txtDown_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtService.SetFocus
End If
End Sub

Private Sub txtDue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmbProduct_Name.Enabled = True Then
cmbProduct_Name.SetFocus
Else
cmbGarantorID_1.SetFocus
End If
End If
End Sub

Private Sub txtDown_Payment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbGtr_Id.SetFocus
End If
End Sub

Private Sub txtFather_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMother.SetFocus
End If
End Sub

Private Sub txtForm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtInsurance.SelStart = 0
   txtInsurance.Text = Format$(Val(txtInsurance.Text), "###0.00")
   txtInsurance.SelLength = Len(txtInsurance.Text)
   txtInsurance.SetFocus
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbLType.SetFocus
End If
End Sub

Private Sub txtId_LostFocus()
Dim mid As Integer
    mid = 0
Dim ID As String
    ID = txtId.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Loan_Info where Customer like '" & ID & "' or AC_No Like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
        If MsgBox("Loan already exist! Do you want to add another Loan?", vbInformation + vbYesNo, "Loan Info") = vbYes Then
            rs.Close
            Call clearTextboxes
            
            Set rs = New ADODB.Recordset
                str = "select * from Personal_Info where Customer like '" & ID & "'"
                rs.Open str, conn
                
               Call clearTextboxes
                txtOpen.Text = Today
                txtId.Text = rs!Customer
                txtName.Text = rs!Name
                txtFather.Text = rs!Fathers_Name
                txtMother.Text = rs!Mothers_Name
                txtBirthday.Text = rs!Age
                cmbReligion.Text = rs!Religion
                cmbNationality.Text = rs!Nationality
                txtPresent.Text = rs!Present_Address
                txtPermanent.Text = rs!Permanent_Address
                txtContact_Office.Text = rs!Contact_Office
                txtContact_Home.Text = rs!Contact_Home
                txtMobile.Text = rs!Mobile
                
                Photo = rs!Photo
                Thumb = rs!Thumb
                
                cmbCenter_Name.Text = rs!Center_name
                cmbCenter_Code.Text = rs!Center_code
                cmbSamity_Code.Text = rs!Samity_Code
                cmbSamity_Name.Text = rs!Samity_Name
                cmbFO_Name.Text = rs!FO_Name
                cmbFO_Code.Text = rs!FO_Code
                cmbDPO_Name.Text = rs!DPO_Name
                cmbDPO_Code.Text = rs!DPO_Code
                cmbAM_Name.Text = rs!AM_Name
                cmbAM_Code.Text = rs!AM_Code
                rs.Close
                
        
                Set rs = New ADODB.Recordset
                str = "select * from Deposit_Master where Customer like '" & ID & "'"
                rs.Open str, conn
                
                txtSecurity.Text = rs!Amount
                rs.Close
                
                Call grd_Gtr
                Call Col_Gtr
                cmdUpdate.Enabled = False
                cmdAdd.Enabled = True
                'cmdDelete.Enabled = True
                cmdPrint.Enabled = False
        Else
        
        Call clearTextboxes
        Call grd_Prod
    
        Set rsU = New ADODB.Recordset
            str = "select * from Prod_Tran where Invo_No like '" & ID & "'"
            rsU.Open str, conn
       
        If Not rsU.EOF Then
            rsU.MoveFirst
        
            Do While Not rsU.EOF
    
                Prod_Sl = Prod_Sl + 1
            
                rsProd.AddNew
                rsProd!sl = Prod_Sl
                rsProd![Product Name] = rsU!Prod_Name
                rsProd![Product Model] = rsU!Prod_Model
                rsProd!Qty = rsU!Sale
                rsProd!Amount = rsU!Sale_Price
    
                Set grdProd.DataSource = rsProd
                grdProd.Refresh
                Call Col_Prod
                rsU.MoveNext
            Loop
        End If
        
        On Error Resume Next
        txtId.Text = rs!AC_No
        cmbLType.Text = rs!Type
        cmbTerm.Text = rs!Term
        
        txtNet.Text = rs!Net_Loan
        netAmnt = rs!Net_Loan
        txtInstallment.Text = rs!Installment
        txtInst_No.Text = rs!Inst_no
        txtOpen.Text = rs!Open_Date
        txtD_Date.Text = rs!D_Date
        txtMaturity.Text = rs!Mat_Date
        
        txtSecurity.Text = rs!Security
        txtForm.Text = rs!Form
        txtInsurance.Text = rs!Insurance
        txtDown_Payment.Text = rs!Down_Payment
        
        txtName.Text = rs!Name
        txtFather.Text = rs!Fathers_Name
        txtMother.Text = rs!Mothers_Name
        txtBirthday.Text = rs!Age
        cmbReligion.Text = rs!Religion
        cmbNationality.Text = rs!Nationality
        txtPresent.Text = rs!Present_Address
        txtPermanent.Text = rs!Permanent_Address
        txtContact_Office.Text = rs!Contact_Office
        txtContact_Home.Text = rs!Contact_Home
        txtMobile.Text = rs!Mobile
        
        cmbCenter_Name.Text = rs!Center_name
        cmbCenter_Code.Text = rs!Center_code
        cmbSamity_Code.Text = rs!Samity_Code
        cmbSamity_Name.Text = rs!Samity_Name
        cmbFO_Name.Text = rs!FO_Name
        cmbFO_Code.Text = rs!FO_Code
        cmbDPO_Name.Text = rs!DPO_Name
        cmbDPO_Code.Text = rs!DPO_Code
        cmbAM_Name.Text = rs!AM_Name
        cmbAM_Code.Text = rs!AM_Code
        
        cmbProduct_Name.Text = rs!Prod_Name
        cmbProd_Model.Text = rs!Prod_Model
   
        Call grd_Gtr
        
        If Not rs!G_ID1 = 0 Then
            rsGtr.AddNew
            rsGtr!sl = 1
            rsGtr![Guarantor ID] = rs!G_ID1
            rsGtr![Guarantor Name] = rs!G_Name1
         End If
         If Not rs!G_ID2 = 0 Then
                rsGtr.AddNew
                rsGtr!sl = 2
                rsGtr![Guarantor ID] = rs!G_ID2
                rsGtr![Guarantor Name] = rs!G_Name2
         End If
         If Not rs!G_ID3 = 0 Then
                rsGtr.AddNew
                rsGtr!sl = 3
                rsGtr![Guarantor ID] = rs!G_ID3
                rsGtr![Guarantor Name] = rs!G_Name3
        End If
        If Not rs!G_ID4 = 0 Then
                rsGtr.AddNew
                rsGtr!sl = 4
                rsGtr![Guarantor ID] = rs!G_ID4
                rsGtr![Guarantor Name] = rs!G_Name4
        End If
        If Not rs!G_ID5 = 0 Then
                rsGtr.AddNew
                rsGtr!sl = 5
                rsGtr![Guarantor ID] = rs!G_ID5
                rsGtr![Guarantor Name] = rs!G_Name5
        End If
        If Not rs!G_ID6 = 0 Then
                rsGtr.AddNew
                rsGtr!sl = 6
                rsGtr![Guarantor ID] = rs!G_ID6
                rsGtr![Guarantor Name] = rs!G_Name6
        End If
        If Not rs!G_ID7 = 0 Then
                rsGtr.AddNew
                rsGtr!sl = 7
                rsGtr![Guarantor ID] = rs!G_ID7
                rsGtr![Guarantor Name] = rs!G_Name7
        End If
        If Not rs!G_ID8 = 0 Then
               rsGtr.AddNew
               rsGtr!sl = 8
               rsGtr![Guarantor ID] = rs!G_ID8
               rsGtr![Guarantor Name] = rs!G_Name8
        End If
        If Not rs!G_ID9 = 0 Then
               rsGtr.AddNew
               rsGtr!sl = 9
               rsGtr![Guarantor ID] = rs!G_ID9
               rsGtr![Guarantor Name] = rs!G_Name9
        End If
        If Not rs!G_ID10 = 0 Then
               rsGtr.AddNew
               rsGtr!sl = 10
               rsGtr![Guarantor ID] = rs!G_ID10
               rsGtr![Guarantor Name] = rs!G_Name10
        End If
        rs.Close
        Set grdGtr.DataSource = rsGtr
        grdGtr.Refresh
        Call Col_Gtr
        
        cmdUpdate.Enabled = True
        cmdAdd.Enabled = False
        cmdPrint.Enabled = True
    End If
    
Else
    rs.Close
    
If MsgBox("Do you want open new Loan Account?", vbInformation + vbYesNo, "Add New") = vbYes Then
        
        Set rs = New ADODB.Recordset
        str = "select * from Personal_Info where Customer like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
        Call clearTextboxes
        txtId.Text = rs!Customer
        
        
        txtName.Text = rs!Name
        txtFather.Text = rs!Fathers_Name
        txtMother.Text = rs!Mothers_Name
        txtBirthday.Text = rs!Age
        cmbReligion.Text = rs!Religion
        cmbNationality.Text = rs!Nationality
        txtPresent.Text = rs!Present_Address
        txtPermanent.Text = rs!Permanent_Address
        txtContact_Office.Text = rs!Contact_Office
        txtContact_Home.Text = rs!Contact_Home
        txtMobile.Text = rs!Mobile
        
        Photo = rs!Photo
        Thumb = rs!Thumb
        
        cmbCenter_Name.Text = rs!Center_name
        cmbCenter_Code.Text = rs!Center_code
        cmbSamity_Code.Text = rs!Samity_Code
        cmbSamity_Name.Text = rs!Samity_Name
        cmbFO_Name.Text = rs!FO_Name
        cmbFO_Code.Text = rs!FO_Code
        cmbDPO_Name.Text = rs!DPO_Name
        cmbDPO_Code.Text = rs!DPO_Code
        cmbAM_Name.Text = rs!AM_Name
        cmbAM_Code.Text = rs!AM_Code
        rs.Close
        
        Set rs = New ADODB.Recordset
        str = "select * from Deposit_Master where Customer like '" & ID & "'"
        rs.Open str, conn
        
        txtSecurity.Text = rs!Amount
        rs.Close
        
        Call grd_Gtr
        Call Col_Gtr
        cmdUpdate.Enabled = False
        cmdAdd.Enabled = True
       
        cmdPrint.Enabled = False
      
        txtOpen.Text = Today
        
 
    Else
        MsgBox "Invalid Customer ID!", vbCritical + vbOKOnly, "Loan Info"
        Call clearTextboxes
        cmbLType.SetFocus
        
    End If
    End If
  End If
End Sub


Private Sub cmbCenter_Code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbCenter_Name.SetFocus
End If
End Sub

Private Sub cmbCenter_Code_LostFocus()
If cmbCenter_Code.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbCenter_Code.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Center_Code, Center_name FROM Personal_Info where Center_Code like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbCenter_Name.Text = rs!Center_name
        rs.Close
        End If
End Sub

Private Sub cmbCenter_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbSamity_Code.SetFocus
End If
End Sub

Private Sub cmbCenter_Name_LostFocus()
If cmbCenter_Name.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbCenter_Name.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Center_Code, Center_name FROM Personal_Info where Center_Name like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbCenter_Code.Text = rs!Center_code
        rs.Close
        End If
End Sub
Private Sub cmbDPO_Code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbDPO_Name.SetFocus
End If
End Sub
Private Sub cmbDPO_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbAM_Code.SetFocus
End If
End Sub
Private Sub cmbFO_Code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbFO_Name.SetFocus
End If
End Sub
Private Sub cmbFO_Code_LostFocus()
If cmbFO_Code.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbFO_Code.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_ID, Name, S_Code, S_Name FROM Employee where Staff_Id like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbFO_Name.Text = rs!Name
        cmbDPO_Code.Text = rs!S_Code
        cmbDPO_Name.Text = rs!S_Name
        rs.Close
        End If
End Sub


Private Sub cmbFO_Name_LostFocus()
If cmbFO_Name.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbFO_Name.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_ID, Name, S_Code, S_Name FROM Employee where Name like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbFO_Code.Text = rs!Staff_Id
        cmbDPO_Code.Text = rs!S_Code
        cmbDPO_Name.Text = rs!S_Name
        rs.Close
        End If
End Sub

Private Sub txtInst_No_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbTerm.SetFocus
End If
End Sub

Private Sub txtInstallment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtInst_No.SelStart = 0
   txtInst_No.Text = Format$(Val(txtInst_No.Text), "###0.00")
   txtInst_No.SelLength = Len(txtInst_No.Text)
   txtInst_No.SetFocus
End If
End Sub

Private Sub txtInsurance_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtDown_Payment.SelStart = 0
   txtDown_Payment.Text = Format$(Val(txtDown_Payment.Text), "###0.00")
   txtDown_Payment.SelLength = Len(txtDown_Payment.Text)
   txtDown_Payment.SetFocus
End If
End Sub

Private Sub txtMaturity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtD_Date.SetFocus
End If
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
        .Fields.Append "Product Name", adBSTR
        .Fields.Append "Product Model", adBSTR
        .Fields.Append "Qty", adBSTR
        .Fields.Append "Amount", adBSTR
        .Open
   
    End With
    Set grdProd.DataSource = rsProd
    
End Sub
Private Sub txtNet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtInstallment.SelStart = 0
   txtInstallment.Text = Format$(Val(txtInstallment.Text), "###0.00")
   txtInstallment.SelLength = Len(txtInstallment.Text)
   txtInstallment.SetFocus
End If
End Sub

Private Sub txtNet_LostFocus()
If cmbLType.Text = "Cash" Then
    netAmnt = 0
    netAmnt = netAmnt + Val(txtNet.Text)
    txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
    strloan = Format$(Val(netAmnt), "###0.00")
    lblNet.Caption = "Net Loan Amount Tk. " + strloan
End If

If cmbLType.Text = "Product" Then
    Prod_Sl = Prod_Sl + 1
    netAmnt = netAmnt + Val(txtNet.Text)
    txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
    strloan = Format$(Val(netAmnt), "###0.00")
    lblNet.Caption = "Net Loan Amount Tk. " + strloan
        rsProd.AddNew
        rsProd!sl = Prod_Sl
        rsProd![Product Name] = cmbProduct_Name.Text
        rsProd![Product Model] = cmbProd_Model.Text
        rsProd!Qty = txtQty.Text
        rsProd!Amount = txtNet.Text

    Set grdProd.DataSource = rsProd
    grdProd.Refresh
   Call Col_Prod
End If
End Sub

Private Sub txtOpen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMaturity.SetFocus
End If
End Sub

Private Sub txtService_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtNet.SetFocus
End If
End Sub

Private Sub txtService_LostFocus()
txtNet.Text = (Val(txtAmount.Text) + Val(txtService.Text)) - Val(txtDown.Text)
End Sub
Private Sub txtMobile_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtBirthday.SetFocus
End If
End Sub
Private Sub txtMother_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPresent.SetFocus
End If
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtFather.SetFocus
End If
End Sub
Private Sub txtPermanent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtContact_Office.SetFocus
End If
End Sub
Private Sub txtPresent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPermanent.SetFocus
End If
End Sub

Private Sub cmbNationality_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbGtr_Id.SetFocus
End If
End Sub
Private Sub cmbReligion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbNationality.SetFocus
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
stk = Val(txtQty.Text)
Call Prod_Stock
   txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
   txtNet.SelStart = 0
   txtNet.SelLength = Len(txtNet.Text)
   txtNet.SetFocus
End If
End Sub
