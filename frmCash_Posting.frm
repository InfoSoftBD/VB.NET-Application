VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmCash_Posting 
   Appearance      =   0  'Flat
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Posting"
   ClientHeight    =   7170
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   11730
   Icon            =   "frmCash_Posting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   180
      ScaleHeight     =   705
      ScaleWidth      =   6885
      TabIndex        =   24
      Top             =   180
      Width           =   6915
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CASH POSTING"
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
         Height          =   540
         Left            =   1650
         TabIndex        =   25
         Top             =   90
         Width           =   3900
      End
      Begin VB.Image Image3 
         Height          =   690
         Left            =   0
         Picture         =   "frmCash_Posting.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6990
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5190
      Left            =   180
      ScaleHeight     =   5160
      ScaleWidth      =   6855
      TabIndex        =   9
      Top             =   1020
      Width           =   6885
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
         Left            =   300
         TabIndex        =   47
         Text            =   "Combo1"
         Top             =   2490
         Width           =   3630
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
         Left            =   330
         TabIndex        =   46
         Text            =   "Combo1"
         Top             =   1620
         Width           =   1545
      End
      Begin VB.TextBox txtCash 
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
         Left            =   5580
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   4560
         Width           =   990
      End
      Begin VB.TextBox txt1000 
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
         Left            =   5550
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox txt500 
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
         Left            =   5550
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   720
         Width           =   1005
      End
      Begin VB.TextBox txt100 
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
         Left            =   5550
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1170
         Width           =   1005
      End
      Begin VB.TextBox txt50 
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
         Left            =   5550
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1650
         Width           =   1005
      End
      Begin VB.TextBox txt20 
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
         Left            =   5550
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   2130
         Width           =   1005
      End
      Begin VB.TextBox txt10 
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
         Left            =   5550
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2610
         Width           =   1005
      End
      Begin VB.TextBox txt5 
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
         Left            =   5550
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   3060
         Width           =   1005
      End
      Begin VB.TextBox txt2 
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
         Left            =   5550
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   3510
         Width           =   1005
      End
      Begin VB.TextBox txt1 
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
         Left            =   5550
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   3960
         Width           =   1005
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
         Left            =   2250
         TabIndex        =   15
         Top             =   1620
         Width           =   1665
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   2250
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   3360
         Width           =   1665
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
         Left            =   360
         TabIndex        =   13
         Top             =   4320
         Width           =   3525
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
         Left            =   360
         TabIndex        =   12
         Top             =   690
         Width           =   1665
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
         Left            =   2250
         TabIndex        =   11
         Top             =   690
         Width           =   1665
      End
      Begin VB.ComboBox cmbType 
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
         Height          =   360
         ItemData        =   "frmCash_Posting.frx":A4D0
         Left            =   360
         List            =   "frmCash_Posting.frx":A4DA
         TabIndex        =   10
         Text            =   "Combo2"
         Top             =   3390
         Width           =   1665
      End
      Begin VB.Shape Shape3 
         Height          =   4845
         Left            =   150
         Top             =   150
         Width           =   3915
      End
      Begin VB.Shape Shape2 
         Height          =   555
         Left            =   4200
         Top             =   4470
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         Height          =   4245
         Left            =   4200
         Top             =   150
         Width           =   2505
      End
      Begin VB.Label lblCash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Taka"
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
         Left            =   4350
         TabIndex        =   45
         Top             =   4620
         Width           =   1005
      End
      Begin VB.Label lbl1000 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 1000.....X "
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
         TabIndex        =   44
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lbl500 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 500......X "
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
         TabIndex        =   43
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label lbl100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 100......X "
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
         TabIndex        =   42
         Top             =   1230
         Width           =   1170
      End
      Begin VB.Label lbl50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 50........X "
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
         TabIndex        =   41
         Top             =   1710
         Width           =   1155
      End
      Begin VB.Label lbl20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 20........X "
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
         TabIndex        =   40
         Top             =   2190
         Width           =   1185
      End
      Begin VB.Label lbl10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 10........X "
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
         TabIndex        =   39
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label lbl5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 5..........X "
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
         TabIndex        =   38
         Top             =   3120
         Width           =   1170
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 2..........X "
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
         TabIndex        =   37
         Top             =   3570
         Width           =   1170
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 1..........X "
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
         TabIndex        =   36
         Top             =   4020
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
         Left            =   2250
         TabIndex        =   23
         Top             =   1230
         Width           =   705
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
         Left            =   360
         TabIndex        =   22
         Top             =   3000
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Head Name"
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
         Left            =   360
         TabIndex        =   21
         Top             =   2130
         Width           =   1020
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
         Left            =   360
         TabIndex        =   20
         Top             =   3900
         Width           =   1035
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   19
         Top             =   1230
         Width           =   1080
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
         Left            =   2250
         TabIndex        =   18
         Top             =   3000
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LC/Ref No."
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
         Left            =   360
         TabIndex        =   17
         Top             =   300
         Width           =   945
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
         Left            =   2250
         TabIndex        =   16
         Top             =   300
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   180
      ScaleHeight     =   705
      ScaleWidth      =   6855
      TabIndex        =   4
      Top             =   6315
      Width           =   6885
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   435
         Left            =   5280
         TabIndex        =   8
         Top             =   150
         Width           =   1350
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   150
         Width           =   1350
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   435
         Left            =   1875
         TabIndex        =   6
         Top             =   150
         Width           =   1350
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Print"
         Height          =   435
         Left            =   3585
         TabIndex        =   5
         Top             =   150
         Width           =   1350
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   6870
      Left            =   7230
      ScaleHeight     =   6840
      ScaleWidth      =   4305
      TabIndex        =   0
      Top             =   150
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
         TabIndex        =   1
         ToolTipText     =   "Enter Consumer no."
         Top             =   180
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCash_Posting.frx":A4F0
         Height          =   5955
         Left            =   240
         TabIndex        =   2
         Top             =   735
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   10504
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
         TabIndex        =   3
         Top             =   255
         Width           =   1845
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6645
      Top             =   5325
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
Attribute VB_Name = "frmCash_Posting"
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
        cmbType.Text = "PAYMENT"
        txtDate.Text = Date
        txtDescription.Text = ""
        txtBalance.Text = "0.00"
        txtNet.Text = "0.00"
        txtCash.Text = "0.00"
        txt1000.Text = "0"
        txt500.Text = "0"
        txt100.Text = "0"
        txt50.Text = "0"
        txt20.Text = "0"
        txt10.Text = "0"
        txt5.Text = "0"
        txt2.Text = "0"
        txt1.Text = "0"
End Sub
Private Sub Total_Cash()
    Dim Total As Double
    Total = (Val(txt1000.Text) * 1000) + (Val(txt500.Text) * 500) + (Val(txt100.Text) * 100) + (Val(txt50.Text) * 50) + (Val(txt20.Text * 20)) + (Val(txt10.Text) * 10) + (Val(txt5.Text) * 5) + (Val(txt2.Text) * 2) + (Val(txt1.Text) * 1)
    txtCash.Text = Format(Total, "###0.00")
'    lblPmt.Caption = "Total Paid Amount Tk. " + Format$(Val(Total), "###0.00")
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
Private Sub Cash_Dr()
    Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!MR_No = mn
            rsN!Name = txtName.Text
            rsN!Description = txtName.Text & "-" & txtDescription.Text
            rsN!Dr = Val(txtNet.Text)
            rsN!Balance = rsU!Balance
            
            
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
            rsU.Close
        
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 1000 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt1000.Text)
            rsU!Balance = rsU!Balance + Val(txt1000.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt1000.Text) * 1000)
            rsU!Cash_Close = rsU!Cash_Close + (Val(txt1000.Text) * 1000)
            rsU.Update
            rsU.Close
            
            
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 500 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt500.Text)
            rsU!Balance = rsU!Balance + Val(txt500.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt500.Text) * 500)
            rsU!Cash_Close = rsU!Cash_Close + (Val(txt500.Text) * 500)
            rsU.Update
            rsU.Close
        
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 100 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt100.Text)
            rsU!Balance = rsU!Balance + Val(txt100.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt100.Text) * 100)
            rsU!Cash_Close = rsU!Cash_Close + (Val(txt100.Text) * 100)
            rsU.Update
            rsU.Close
        
        
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 50 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt50.Text)
            rsU!Balance = rsU!Balance + Val(txt50.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt50.Text) * 50)
            rsU!Cash_Close = rsU!Cash_Close + (Val(txt50.Text) * 50)
            rsU.Update
            rsU.Close
        
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 20 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt20.Text)
            rsU!Balance = rsU!Balance + Val(txt20.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt20.Text) * 20)
            rsU!Cash_Close = rsU!Cash_Close + (Val(txt20.Text) * 20)
            rsU.Update
            rsU.Close
    
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 10 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt10.Text)
            rsU!Balance = rsU!Balance + Val(txt10.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt10.Text) * 10)
            rsU!Cash_Close = rsU!Cash_Close + (Val(txt10.Text) * 10)
            rsU.Update
            rsU.Close
    
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 5 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt5.Text)
            rsU!Balance = rsU!Balance + Val(txt5.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt5.Text) * 5)
            rsU!Cash_Close = rsU!Cash_Close + (Val(txt5.Text) * 5)
            rsU.Update
            rsU.Close
        
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 2 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt2.Text)
            rsU!Balance = rsU!Balance + Val(txt2.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt2.Text) * 2)
            rsU!Cash_Close = rsU!Cash_Close + (Val(txt2.Text) * 2)
            rsU.Update
            rsU.Close
    
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 1 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt1.Text)
            rsU!Balance = rsU!Balance + Val(txt1.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt1.Text) * 1)
            rsU!Cash_Close = rsU!Cash_Close + (Val(txt1.Text) * 1)
            rsU.Update
            rsU.Close
                        
            
    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.MoveFirst
        rs!Cash_Dr = rs!Cash_Dr + Val(txtNet.Text)
        rs!Cash_Close = rs!Cash_Close + Val(txtNet.Text)
        rs.Update
        rs.Close
End Sub
Private Sub Cash_Cr()
Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!MR_No = mn
            rsN!Name = txtName.Text
            rsN!Description = txtDescription.Text
            rsN!Cr = Val(txtNet.Text)
            rsN!Balance = rsU!Balance
                            
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
            rsU.Close

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
                    
    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.MoveFirst
        rs!Cash_Cr = rs!Cash_Cr + Val(txtNet.Text)
        rs!Cash_Close = rs!Cash_Close - Val(txtNet.Text)
        rs.Update
        rs.Close
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
Dim mn As String
On Error Resume Next
   If cmbType.Text = "" Then
   MsgBox "Please Select Transaction Type RECEIVE/PAYMENT", vbCritical
   Exit Sub
   End If
   
    Set rsU = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & txtAccount.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
If Not rsU.EOF Then
'-----------------------------------------------------------------------
If cmbType.Text = "PAYMENT" Then
    If rsU!Head_Type = "EXPENSE" Or rsU!Head_Type = "ASSET" Then
    
          
            rsU!Balance = rsU!Balance + Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = txtAccount.Text
                rsN!Name = txtTran.Text
                rsN!Description = txtDescription.Text
                rsN!Dr = Val(txtNet.Text)
                rsN!Cr = 0
                rsN!Balance = rsU!Balance
                rsN.Update
                rsN.Close
                
            Set rsU = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Balance = rsU!Balance - Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = rsU!AC_No
                rsN!Name = rsU!Head_Name
                rsN!Description = txtDescription.Text
                rsN!Dr = 0
                rsN!Cr = Val(txtNet.Text)
                rsN!Balance = rsU!Balance
                rsN.Update
                mn = rsN!sl
                rsN.Close
                    
                Call Cash_Cr
'------------------------------------------------------------------------------------
    Else
    If rsU!Head_Type = "INCOME" Or rsU!Head_Type = "LIABILITY" Then
    
          
            rsU!Balance = rsU!Balance - Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = txtAccount.Text
                rsN!Name = txtTran.Text
                rsN!Description = txtDescription.Text
                rsN!Dr = Val(txtNet.Text)
                rsN!Cr = 0
                rsN!Balance = rsU!Balance
                rsN.Update
                rsN.Close
                
            Set rsU = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Balance = rsU!Balance - Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = rsU!AC_No
                rsN!Name = rsU!Head_Name
                rsN!Description = txtDescription.Text
                rsN!Dr = 0
                rsN!Cr = Val(txtNet.Text)
                rsN!Balance = rsU!Balance
                rsN.Update
                mn = rsN!sl
                rsN.Close
                    
                Call Cash_Cr
           
        End If
    End If
Else
'----------------------------------------------------------------------------------
If cmbType.Text = "RECEIVE" Then
    If rsU!Head_Type = "EXPENSE" Or rsU!Head_Type = "ASSET" Then
    
          
            rsU!Balance = rsU!Balance - Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = txtAccount.Text
                rsN!Name = txtTran.Text
                rsN!Description = txtDescription.Text
                rsN!Dr = 0
                rsN!Cr = Val(txtNet.Text)
                rsN!Balance = rsU!Balance
                rsN.Update
                rsN.Close
                
            Set rsU = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Balance = rsU!Balance + Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = rsU!AC_No
                rsN!Name = rsU!Head_Name
                rsN!Description = txtDescription.Text
                rsN!Dr = Val(txtNet.Text)
                rsN!Cr = 0
                rsN!Balance = rsU!Balance
                rsN.Update
                mn = rsN!sl
                rsN.Close
                    
                Call Cash_Dr
'------------------------------------------------------------------------------------
    Else
    If rsU!Head_Type = "INCOME" Or rsU!Head_Type = "LIABILITY" Then
    
          
            rsU!Balance = rsU!Balance + Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = txtAccount.Text
                rsN!Name = txtTran.Text
                rsN!Description = txtDescription.Text
                rsN!Dr = 0
                rsN!Cr = Val(txtNet.Text)
                rsN!Balance = rsU!Balance
                rsN.Update
                rsN.Close
                
            Set rsU = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Balance = rsU!Balance + Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = rsU!AC_No
                rsN!Name = rsU!Head_Name
                rsN!Description = txtDescription.Text
                rsN!Dr = Val(txtNet.Text)
                rsN!Cr = 0
                rsN!Balance = rsU!Balance
                rsN.Update
                mn = rsN!sl
                rsN.Close
                    
                Call Cash_Dr
       End If
    End If
End If
End If
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
    txtAccount.SetFocus
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
Exit Sub
   Resume Next
End Sub

Private Sub Command2_GotFocus()
'If Val(txtNet.Text) <> Val(txtCash.Text) Then
'MsgBox "Denomination Amount differs from Net Amount" & Val(txtNet.Text) - Val(txtCash.Text), vbCritical, "Sale Info!"
'txtNet.SetFocus
'Command2.Enabled = False
'Else
'Command2.SetFocus
'End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If txtTran.Text = "" Then
Exit Sub
End If

On Error Resume Next
Dim Tran As String
    Tran = txtTran.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Cash_Book where Mr_No like '" & Tran & "'"
        rs.Open str, conn
    
If Not rs.EOF Then
    On Error Resume Next
        
        txtTran.Text = rs!MR_No
        txtDate.Text = rs!Date
        txtAccount.Text = rs!Description
        txtName.Text = rs!Name
        txtDescription.Text = "Posting Reversed"
        txtBalance.Text = rs!Balance
        If rs!Cr > 0 Then
        txtNet.Text = rs!Cr
        cmbType.Text = "C"
        Else
        txtNet.Text = rs!Dr
        cmbType.Text = "D"
        End If
        rs.Close
     
    If MsgBox("Really want to delete?", vbCritical + vbYesNo) = vbYes Then
        
        str = "delete from Cash_Book where Mr_No like '" & txtTran.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.Close
        
   
    Set rsU = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & txtAccount.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
    '-----------------------------------------------------------------------
        
        If cmbType.Text = "C" Then
          
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
                
            Set rsU = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Balance = rsU!Balance + Val(txtNet.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = rsU!AC_No
                rsN!Name = rsU!Head_Name
                rsN!Description = txtDescription.Text
                rsN!Dr = Val(txtNet.Text)
                rsN!Cr = 0
                rsN!Balance = rsU!Balance
                rsN.Update
                rsN.Close
  
            Set rs = New ADODB.Recordset
                str = "select * from Others"
                rs.Open str, conn, adOpenDynamic, adLockOptimistic
                rs.MoveFirst
                rs!Cash_Cr = rs!Cash_Cr - Val(txtNet.Text)
                rs!Cash_Close = rs!Cash_Close + Val(txtNet.Text)
                rs.Update
                rs.Close
'------------------------------------------------------------------------------------
    Else
    
    If cmbType.Text = "D" Then
          
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
                
'------------------------------------------------------------------------------------
            
                Set rsU = New ADODB.Recordset
                str = "select * from GL_Master where AC_No like '" & 100100 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
                rsU!Balance = rsU!Balance - Val(txtNet.Text)
                rsU!Date = txtDate.Text
                rsU.Update
                
                Set rsN = New ADODB.Recordset
                rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
                rsN.AddNew
                    rsN!Date = txtDate.Text
                    rsN!AC_No = rsU!AC_No
                    rsN!Name = rsU!Head_Name
                    rsN!Description = txtDescription.Text
                    rsN!Dr = 0
                    rsN!Cr = Val(txtNet.Text)
                    rsN!Balance = rsU!Balance
                    rsN.Update
                    rsN.Close
            
            Set rs = New ADODB.Recordset
                str = "select * from Others"
                rs.Open str, conn, adOpenDynamic, adLockOptimistic
                rs.MoveFirst
                rs!Cash_Dr = rs!Cash_Dr - Val(txtNet.Text)
                rs!Cash_Close = rs!Cash_Close - Val(txtNet.Text)
                rs.Update
                rs.Close
            End If
        End If
    End If
End If
    Set rs = New ADODB.Recordset
        str = "select * from Cash_Book where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by AC_No"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
        Call ColumnWidth
        Call clearTextboxes
Else
        MsgBox "There is no such Transaction no. found: " & txtTran.Text
        rs.Close
        Call clearTextboxes
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
        rsU!Account_No = txtAccount.Text
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
Private Sub txt1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCash.SelStart = 0
    txtCash.SelLength = Len(txtCash.Text)
    txtCash.SetFocus
End If
End Sub

Private Sub txt1_LostFocus()
    Call Total_Cash
End Sub

Private Sub txt10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt5.SelStart = 0
    txt5.SelLength = Len(txt5.Text)
    txt5.SetFocus
End If
End Sub

Private Sub txt10_LostFocus()
    Call Total_Cash
End Sub

Private Sub txt100_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt50.SelStart = 0
    txt50.SelLength = Len(txt50.Text)
    txt50.SetFocus
End If
End Sub

Private Sub txt100_LostFocus()
Call Total_Cash
End Sub

Private Sub txt1000_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt500.SelStart = 0
    txt500.SelLength = Len(txt500.Text)
    txt500.SetFocus
End If
End Sub
Private Sub txt1000_LostFocus()
Call Total_Cash
End Sub
Private Sub txt2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt1.SelStart = 0
    txt1.SelLength = Len(txt1.Text)
    txt1.SetFocus
End If
End Sub
Private Sub txt2_LostFocus()
    Call Total_Cash
End Sub
Private Sub txt20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt10.SelStart = 0
    txt10.SelLength = Len(txt10.Text)
    txt10.SetFocus
End If
End Sub
Private Sub txt20_LostFocus()
    Call Total_Cash
End Sub
Private Sub txt5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt2.SelStart = 0
    txt2.SelLength = Len(txt2.Text)
    txt2.SetFocus
End If
End Sub

Private Sub txt5_LostFocus()
    Call Total_Cash
End Sub

Private Sub txt50_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt20.SelStart = 0
    txt20.SelLength = Len(txt20.Text)
    txt20.SetFocus
End If
End Sub

Private Sub txt50_LostFocus()
    Call Total_Cash
End Sub

Private Sub txt500_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt100.SelStart = 0
    txt100.SelLength = Len(txt100.Text)
    txt100.SetFocus
End If
End Sub

Private Sub txt500_LostFocus()
Call Total_Cash
End Sub

Private Sub Form_Load()
On Error Resume Next
     Call clearTextboxes
     txtDate.Text = Today
     cmbType.Text = "PAYMENT"
        
    Set rs = New ADODB.Recordset
        str = "select * from Gl_Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by AC_No"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
        Call CmbAc
        Call CmbName
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
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



Private Sub txtCash_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command2.Enabled = True
txtDescription.SetFocus
End If
End Sub

Private Sub txtCash_LostFocus()
Call Total_Cash
If Val(txtNet.Text) <> Val(txtCash.Text) Then
MsgBox "Denomination Amount differs from Net Amount" & Val(txtNet.Text) - Val(txtCash.Text), vbCritical, "Sale Info!"
Command2.Enabled = True
txtDescription.SetFocus
Else
Command2.Enabled = True
txtDescription.SetFocus
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
    On Error Resume Next
        txtAccount.Text = rs!AC_No
        txtDate.Text = Today
        txtName.Text = rs!Head_Name
        txtBalance.Text = rs!Balance
        rs.Close
        
        txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
        
        txtBalance.Text = Format$(Val(txtBalance.Text), "###0.00")

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
    txt1000.SelStart = 0
    txt1000.SelLength = Len(txtNet.Text)
    txt1000.SetFocus
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
If cmbType.Text = "PAYMENT" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn
        
        If rs!Balance < Val(txtNet.Text) Then
        MsgBox "Insufficiant Balance.!", vbCritical, "Error!"
        txtNet.SetFocus
        Else
        txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
        End If
Else
txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
End If
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
If txtTran.Text = "" Then
Exit Sub
End If

On Error Resume Next
Dim Tran As String
    Tran = txtTran.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Cash_Book where Mr_No like '" & Tran & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    
        txtTran.Text = rs!MR_No
        txtDate.Text = rs!Date
        txtAccount.Text = rs!Description
        txtName.Text = rs!Name
        txtDescription.Text = "Posting Reversed"
        txtBalance.Text = rs!Balance
        If rs!Cr > 0 Then
        txtNet.Text = rs!Cr
        cmbType.Text = "C"
        Else
        txtNet.Text = rs!Dr
        cmbType.Text = "D"
        End If
        rs.Close
        
        txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
        txtBalance.Text = Format$(Val(txtBalance.Text), "###0.00")
        Command3.Enabled = True
        Command4.Enabled = True
        'Command5.Enabled = True
        Command2.Enabled = False
        
    Else
'    If MsgBox("There is no such Transaction no found! Verify Transaction no.?", vbCritical + vbYesNo) = vbYes Then
'        Call clearTextboxes
'        rs.Close
'        txtTran.SetFocus
'        'txtAccount.Text = ac
'        txtDate.Text = Today
'    Else
'        Call clearTextboxes
'        'txtAccount.Text = ac
        txtDate.Text = Today
        txtAccount.SetFocus
'    End If
    End If
    Exit Sub
End If
End Sub







