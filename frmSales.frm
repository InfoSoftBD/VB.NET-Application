VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmSales 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Sales Entry"
   ClientHeight    =   10905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10905
   ScaleWidth      =   14025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cash Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4485
      Left            =   11040
      TabIndex        =   69
      Top             =   960
      Width           =   2835
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
         Left            =   1710
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   210
         Width           =   975
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
         Left            =   1710
         TabIndex        =   78
         Text            =   "Text1"
         Top             =   630
         Width           =   975
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
         Left            =   1710
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   1050
         Width           =   975
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
         Left            =   1710
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   1470
         Width           =   975
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
         Left            =   1710
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   1890
         Width           =   975
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
         Left            =   1710
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   2310
         Width           =   975
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
         Left            =   1710
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   2730
         Width           =   975
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
         Left            =   1710
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   3150
         Width           =   975
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
         Left            =   1710
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   3570
         Width           =   975
      End
      Begin VB.TextBox txtCash 
         Alignment       =   1  'Right Justify
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
         Left            =   1530
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   3990
         Width           =   1170
      End
      Begin VB.Label Label32 
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
         Left            =   180
         TabIndex        =   89
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label33 
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
         Left            =   180
         TabIndex        =   88
         Top             =   690
         Width           =   1320
      End
      Begin VB.Label Label34 
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
         Left            =   180
         TabIndex        =   87
         Top             =   1110
         Width           =   1320
      End
      Begin VB.Label Label35 
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
         Left            =   180
         TabIndex        =   86
         Top             =   1530
         Width           =   1305
      End
      Begin VB.Label Label36 
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
         Left            =   180
         TabIndex        =   85
         Top             =   1950
         Width           =   1335
      End
      Begin VB.Label Label37 
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
         Left            =   180
         TabIndex        =   84
         Top             =   2370
         Width           =   1305
      End
      Begin VB.Label Label38 
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
         Left            =   180
         TabIndex        =   83
         Top             =   2790
         Width           =   1320
      End
      Begin VB.Label Label39 
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
         Left            =   180
         TabIndex        =   82
         Top             =   3210
         Width           =   1320
      End
      Begin VB.Label Label40 
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
         Left            =   180
         TabIndex        =   81
         Top             =   3630
         Width           =   1320
      End
      Begin VB.Label lblCash 
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
         Left            =   180
         TabIndex        =   80
         Top             =   4080
         Width           =   1275
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bank Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   11040
      TabIndex        =   58
      Top             =   5550
      Width           =   2835
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
         Left            =   810
         TabIndex        =   63
         Text            =   "Combo1"
         Top             =   270
         Width           =   1905
      End
      Begin VB.ComboBox cmbBank 
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
         Height          =   360
         Left            =   825
         TabIndex        =   62
         Text            =   "Combo1"
         Top             =   705
         Width           =   1890
      End
      Begin VB.ComboBox cmbBranch 
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
         Height          =   360
         Left            =   810
         TabIndex        =   61
         Text            =   "Combo2"
         Top             =   1140
         Width           =   1905
      End
      Begin VB.TextBox txtChq_No 
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
         Left            =   1260
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   1590
         Width           =   1410
      End
      Begin VB.TextBox txtChq_Amnt 
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
         Left            =   1260
         TabIndex        =   59
         Text            =   "Text7"
         Top             =   2040
         Width           =   1410
      End
      Begin VB.Label lblAccount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C No."
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
         TabIndex        =   68
         Top             =   330
         Width           =   690
      End
      Begin VB.Label lblBank 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
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
         TabIndex        =   67
         Top             =   780
         Width           =   450
      End
      Begin VB.Label lblBranch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
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
         TabIndex        =   66
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label lblChq_No 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chq No."
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
         TabIndex        =   65
         Top             =   1650
         Width           =   675
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chq Amnt"
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
         Left            =   90
         TabIndex        =   64
         Top             =   2100
         Width           =   930
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Card Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   11040
      TabIndex        =   51
      Top             =   8220
      Width           =   2835
      Begin VB.ComboBox cmbCard_Name 
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
         Height          =   360
         Left            =   1110
         TabIndex        =   54
         Text            =   "Combo1"
         Top             =   270
         Width           =   1620
      End
      Begin VB.TextBox txtCard_No 
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
         Height          =   390
         Left            =   1140
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   720
         Width           =   1560
      End
      Begin VB.TextBox txtCard_Amnt 
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
         Left            =   1170
         TabIndex        =   52
         Text            =   "Text7"
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Card Name"
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
         TabIndex        =   57
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Card No."
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
         TabIndex        =   56
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Card Amnt"
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
         Left            =   90
         TabIndex        =   55
         Top             =   1230
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   11040
      ScaleHeight     =   660
      ScaleWidth      =   2775
      TabIndex        =   49
      Top             =   10050
      Width           =   2835
      Begin VB.Label lblPmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Tk. 0.00"
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
         Left            =   1215
         TabIndex        =   50
         Top             =   180
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   165
      ScaleHeight     =   690
      ScaleWidth      =   10695
      TabIndex        =   23
      Top             =   10020
      Width           =   10755
      Begin VB.CommandButton cmdCash_Memo 
         Caption         =   "Agent Memo"
         Height          =   525
         Left            =   1845
         TabIndex        =   93
         Top             =   90
         Width           =   1650
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Sales Invoice"
         Height          =   525
         Left            =   3615
         TabIndex        =   90
         Top             =   90
         Width           =   1650
      End
      Begin VB.CommandButton cmdSales 
         Caption         =   "Save"
         Height          =   525
         Left            =   90
         TabIndex        =   27
         Top             =   90
         Width           =   1650
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Pos Memo"
         Height          =   525
         Left            =   5400
         TabIndex        =   26
         Top             =   90
         Width           =   1650
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   525
         Left            =   8955
         TabIndex        =   25
         Top             =   90
         Width           =   1650
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   525
         Left            =   7170
         TabIndex        =   24
         Top             =   90
         Width           =   1650
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
      Height          =   5715
      Left            =   165
      TabIndex        =   9
      Top             =   4200
      Width           =   10740
      Begin VB.TextBox txtCost 
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
         Height          =   360
         Left            =   765
         TabIndex        =   98
         Text            =   "0.00"
         Top             =   4950
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox txtStock 
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
         Left            =   6120
         TabIndex        =   96
         Text            =   "0.00"
         Top             =   555
         Width           =   570
      End
      Begin VB.TextBox txtCharge 
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
         Left            =   9015
         TabIndex        =   94
         Text            =   "Text7"
         Top             =   4950
         Width           =   1560
      End
      Begin VB.TextBox txtTerms 
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
         Left            =   150
         TabIndex        =   91
         Text            =   "Text7"
         Top             =   4950
         Width           =   8760
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
         Left            =   1785
         TabIndex        =   45
         Top             =   570
         Width           =   1875
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
         Left            =   3750
         TabIndex        =   44
         Text            =   "Combo2"
         Top             =   570
         Width           =   2310
      End
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
         Left            =   180
         TabIndex        =   43
         Text            =   "Combo1"
         Top             =   570
         Width           =   1545
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
         Left            =   6795
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   555
         Width           =   480
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
         Left            =   8310
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   555
         Width           =   945
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
         Left            =   7380
         TabIndex        =   12
         Text            =   "Text7"
         Top             =   555
         Width           =   825
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
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   555
         Width           =   735
      End
      Begin VB.TextBox txtPercent 
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
         Left            =   9360
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   555
         Width           =   405
      End
      Begin MSDataGridLib.DataGrid grdProd 
         Height          =   3585
         Left            =   150
         TabIndex        =   15
         Top             =   1050
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   6324
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
      Begin VB.Label lblCost 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cost Tk. 0.00"
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
         Left            =   195
         TabIndex        =   99
         Top             =   5400
         Width           =   1725
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock"
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
         Left            =   6150
         TabIndex        =   97
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label20 
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
         Left            =   9045
         TabIndex        =   95
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terms and Condition"
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
         TabIndex        =   92
         Top             =   4680
         Width           =   1785
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Description"
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
         Left            =   3795
         TabIndex        =   48
         Top             =   240
         Width           =   1710
      End
      Begin VB.Label Label8 
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
         Left            =   1785
         TabIndex        =   47
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   195
         TabIndex        =   46
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label13 
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
         Left            =   6870
         TabIndex        =   21
         Top             =   225
         Width           =   315
      End
      Begin VB.Label Label12 
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
         Left            =   8340
         TabIndex        =   20
         Top             =   225
         Width           =   930
      End
      Begin VB.Label Label11 
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
         Left            =   7380
         TabIndex        =   19
         Top             =   225
         Width           =   855
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
         Left            =   8520
         TabIndex        =   18
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label Label5 
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
         TabIndex        =   17
         Top             =   225
         Width           =   765
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
         Left            =   9450
         TabIndex        =   16
         Top             =   225
         Width           =   180
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Information"
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
      Left            =   180
      TabIndex        =   2
      Top             =   960
      Width           =   10725
      Begin VB.ComboBox cmbAgent_Code 
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
         Left            =   1305
         TabIndex        =   102
         Text            =   "Combo1"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ComboBox cmbAgent_Name 
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
         TabIndex        =   101
         Text            =   " "
         Top             =   2520
         Width           =   3255
      End
      Begin VB.ComboBox cmbP_Type 
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
         Left            =   8775
         TabIndex        =   100
         Text            =   " "
         Top             =   2520
         Width           =   1725
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
         Left            =   1200
         TabIndex        =   41
         Text            =   "Text4"
         Top             =   360
         Width           =   1605
      End
      Begin VB.ComboBox txtType 
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
         Height          =   360
         ItemData        =   "frmSales.frx":0442
         Left            =   6360
         List            =   "frmSales.frx":044F
         TabIndex        =   40
         Text            =   "Combo1"
         Top             =   1110
         Width           =   1560
      End
      Begin VB.TextBox txtMobile 
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
         Height          =   405
         Left            =   6360
         TabIndex        =   37
         Text            =   "Text17"
         Top             =   1830
         Width           =   1545
      End
      Begin VB.TextBox txtAddress 
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
         Height          =   405
         Left            =   270
         TabIndex        =   36
         Text            =   "Text15"
         Top             =   1830
         Width           =   5940
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   9540
         TabIndex        =   34
         Text            =   "Text3"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtTotal_Due 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   9540
         TabIndex        =   31
         Text            =   "Text3"
         Top             =   1830
         Width           =   975
      End
      Begin VB.TextBox txtSales_Due 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   9540
         TabIndex        =   30
         Text            =   "Text3"
         Top             =   1140
         Width           =   975
      End
      Begin VB.ComboBox txtID 
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
         Left            =   270
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   1110
         Width           =   1365
      End
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
         Left            =   1740
         TabIndex        =   28
         Text            =   "Combo1"
         Top             =   1125
         Width           =   4515
      End
      Begin VB.TextBox txtDate 
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
         Left            =   6360
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   360
         Width           =   1545
      End
      Begin VB.ComboBox cmbType 
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
         Height          =   360
         ItemData        =   "frmSales.frx":046C
         Left            =   3990
         List            =   "frmSales.frx":0479
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   360
         Width           =   1800
      End
      Begin VB.Shape Shape1 
         Height          =   555
         Left            =   135
         Top             =   2430
         Width           =   10455
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Godown ID"
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
         Left            =   315
         TabIndex        =   105
         Top             =   2580
         Width           =   945
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Godown Name"
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
         Left            =   2790
         TabIndex        =   104
         Top             =   2580
         Width           =   1275
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Type"
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
         Left            =   7470
         TabIndex        =   103
         Top             =   2580
         Width           =   1200
      End
      Begin VB.Shape Shape4 
         Height          =   2085
         Left            =   8160
         Top             =   240
         Width           =   2445
      End
      Begin VB.Shape Shape3 
         Height          =   2115
         Left            =   120
         Top             =   240
         Width           =   7935
      End
      Begin VB.Label Label6 
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
         Left            =   270
         TabIndex        =   42
         Top             =   420
         Width           =   840
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
         Left            =   6360
         TabIndex        =   39
         Top             =   1530
         Width           =   930
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Address"
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
         TabIndex        =   38
         Top             =   1530
         Width           =   1620
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance B/F"
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
         Left            =   8250
         TabIndex        =   35
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Payable"
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
         Left            =   8280
         TabIndex        =   33
         Top             =   1890
         Width           =   1065
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Amount"
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
         Left            =   8280
         TabIndex        =   32
         Top             =   1170
         Width           =   1125
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Type"
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
         Left            =   6360
         TabIndex        =   22
         Top             =   810
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Customer"
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
         Left            =   1740
         TabIndex        =   8
         Top             =   840
         Width           =   1620
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
         Left            =   270
         TabIndex        =   7
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Type"
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
         Left            =   2880
         TabIndex        =   6
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label4 
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
         Left            =   5850
         TabIndex        =   5
         Top             =   420
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   150
      ScaleHeight     =   660
      ScaleWidth      =   13710
      TabIndex        =   0
      Top             =   135
      Width           =   13740
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT SALES ENTRY"
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
         Left            =   3360
         TabIndex        =   1
         Top             =   45
         Width           =   6945
      End
      Begin VB.Image Image3 
         Height          =   2085
         Left            =   0
         Picture         =   "frmSales.frx":04A3
         Stretch         =   -1  'True
         Top             =   0
         Width           =   13710
      End
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Prod_Sl As Integer
Dim netAmnt As Double
Dim cst As Double

Dim Total As Double
Dim stk As Double
Dim strloan As String
Dim sqlStr As String
Dim rsProd As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Dim inword As String
Dim StrD As String
Dim Value As Double
Dim result As Double
Dim StrPs As String
Dim Str1 As String
Dim Str10 As String
Dim Str100 As String
Dim Str1000 As String
Dim Str100000 As String
Dim Str10000000 As String
Dim Tk As String
Dim ps As String
Dim Only As String
Dim n0, n1, n2, n3, n4, n5, n6, n7, n8, n9, n10 As String
Dim n11, n12, n13, n14, n15, n16, n17, n18, n19, n20 As String
Dim n30, n40, n50, n60, n70, n80, n90, n100, n1000, n100000, n10000000 As String
Private Sub Case_Result()
    Value = result
    n0 = "Zero ": n1 = "One ": n2 = "Two ": n3 = "Three ": n4 = "Four "
    n5 = "Five ": n6 = "Six ": n7 = "Seven ": n8 = "Eight ": n9 = "Nine "
    n8 = "Eight ": n9 = "Nine ": n10 = "Ten ": n11 = "Eleven ": n12 = "Twelve ": n13 = "Thirteen "
    n14 = "Fourteen ": n15 = "Fifteen ": n16 = "Sixteen ": n17 = "Seventeen ": n18 = "Eighteen ": n19 = "Nineteen "
    n20 = "Twenty ": n30 = "Thirty ": n40 = "Forty ": n50 = "Fifty ": n60 = "Sixty ": n70 = "Seventy "
    n80 = "Eighty ": n90 = "Ninety ": n100 = "Hundred ": n1000 = "Thousand ": n100000 = "Lakh ": n10000000 = "Crore "
    Tk = "Taka ": ps = "Paisa ": Only = "Only"
    If Value = 0 Then
        str = n0
    End If
    If Value = 1 Then
        str = n1
    End If
    If Value = 2 Then
        str = n2
    End If
    If Value = 3 Then
        str = n3
    End If
    If Value = 4 Then
        str = n4
    End If
    If Value = 5 Then
        str = n5
    End If
    If Value = 6 Then
        str = n6
    End If
    If Value = 7 Then
        str = n7
    End If
    If Value = 8 Then
        str = n8
    End If
    If Value = 9 Then
        str = n9
    End If
    If Value = 10 Then
        str = n10
    End If
    If Value = 11 Then
        str = n11
    End If
    If Value = 12 Then
        str = n12
    End If
    If Value = 13 Then
        str = n13
    End If
    If Value = 14 Then
        str = n14
    End If
    If Value = 15 Then
        str = n15
    End If
    If Value = 16 Then
        str = n16
    End If
    If Value = 17 Then
        str = n17
    End If
    If Value = 18 Then
        str = n18
    End If
    If Value = 19 Then
        str = n19
    End If
    If Value = 20 Then
        str = n20
    End If
    If Value = 21 Then
        str = n20 + n1
    End If
    If Value = 22 Then
        str = n20 + n2
    End If
    If Value = 23 Then
        str = n20 + n3
    End If
    If Value = 24 Then
        str = n20 + n4
    End If
    If Value = 25 Then
        str = n20 + n5
    End If
    If Value = 26 Then
        str = n20 + n6
    End If
    If Value = 27 Then
        str = n20 + n7
    End If
    If Value = 28 Then
        str = n20 + n8
    End If
    If Value = 29 Then
        str = n20 + n9
    End If
    If Value = 30 Then
        str = n30
    End If
    If Value = 31 Then
        str = n30 + n1
    End If
    If Value = 32 Then
        str = n30 + n2
    End If
    If Value = 33 Then
        str = n30 + n3
    End If
    If Value = 34 Then
        str = n30 + n4
    End If
    If Value = 35 Then
        str = n30 + n5
    End If
    If Value = 36 Then
        str = n30 + n6
    End If
    If Value = 37 Then
        str = n30 + n7
    End If
    If Value = 38 Then
        str = n30 + n8
    End If
    If Value = 39 Then
        str = n30 + n9
    End If
    If Value = 40 Then
        str = n40
    End If
    If Value = 41 Then
        str = n40 + n1
    End If
    If Value = 42 Then
        str = n40 + n2
    End If
    If Value = 43 Then
        str = n40 + n3
    End If
    If Value = 44 Then
        str = n40 + n4
    End If
    If Value = 45 Then
        str = n40 + n5
    End If
    If Value = 46 Then
        str = n40 + n6
    End If
    If Value = 47 Then
        str = n40 + n7
    End If
    If Value = 48 Then
        str = n40 + n8
    End If
    If Value = 49 Then
        str = n40 + n9
    End If
    If Value = 50 Then
        str = n50
    End If
    If Value = 51 Then
        str = n50 + n1
    End If
    If Value = 52 Then
        str = n50 + n2
    End If
    If Value = 53 Then
        str = n50 + n3
    End If
    If Value = 54 Then
        str = n50 + n4
    End If
    If Value = 55 Then
        str = n50 + n5
    End If
    If Value = 56 Then
        str = n50 + n6
    End If
    If Value = 57 Then
        str = n50 + n7
    End If
    If Value = 58 Then
        str = n50 + n8
    End If
    If Value = 59 Then
        str = n50 + n9
    End If
    If Value = 60 Then
        str = n60
    End If
    If Value = 61 Then
        str = n60 + n1
    End If
    If Value = 62 Then
        str = n60 + n2
    End If
    If Value = 63 Then
        str = n60 + n3
    End If
    If Value = 64 Then
        str = n60 + n4
    End If
    If Value = 65 Then
        str = n60 + n5
    End If
    If Value = 66 Then
        str = n60 + n6
    End If
    If Value = 67 Then
        str = n60 + n7
    End If
    If Value = 68 Then
        str = n60 + n8
    End If
    If Value = 69 Then
    str = n60 + n9
    End If
    If Value = 70 Then
        str = n70
    End If
    If Value = 71 Then
    str = n70 + n1
    End If
    If Value = 72 Then
        str = n70 + n2
    End If
    If Value = 73 Then
    str = n70 + n3
    End If
    If Value = 74 Then
        str = n70 + n4
    End If
    If Value = 75 Then
        str = n70 + n5
    End If
    If Value = 76 Then
        str = n70 + n6
    End If
    If Value = 77 Then
        str = n70 + n7
    End If
    If Value = 78 Then
        str = n70 + n8
    End If
    If Value = 79 Then
        str = n70 + n9
    End If
    If Value = 80 Then
        str = n80
    End If
    If Value = 81 Then
        str = n80 + n1
    End If
    If Value = 82 Then
        str = n80 + n2
    End If
    If Value = 83 Then
        str = n80 + n3
    End If
    If Value = 84 Then
        str = n80 + n4
    End If
    If Value = 85 Then
        str = n80 + n5
    End If
    If Value = 86 Then
        str = n80 + n6
    End If
    If Value = 87 Then
        str = n80 + n7
    End If
    If Value = 88 Then
        str = n80 + n8
    End If
    If Value = 89 Then
        str = n80 + n9
    End If
    If Value = 90 Then
        str = n90
    End If
    If Value = 91 Then
        str = n90 + n1
    End If
    If Value = 92 Then
        str = n90 + n2
    End If
    If Value = 93 Then
        str = n90 + n3
    End If
    If Value = 94 Then
        str = n90 + n4
    End If
    If Value = 95 Then
        str = n90 + n5
    End If
    If Value = 96 Then
        str = n90 + n6
    End If
    If Value = 97 Then
        str = n90 + n7
    End If
    If Value = 98 Then
        str = n90 + n8
    End If
    If Value = 99 Then
        str = n90 + n9
    End If
 End Sub
Private Sub Case0_9()
         result = Value
    Call Case_Result
        Str1 = str
        inword = Tk + Str1 + Only
 End Sub
Private Sub Case10_99()
    result = Value
    Call Case_Result
        Str10 = str
        inword = Tk + Str10 + Only
End Sub
Private Sub Case100_999()
    Dim Mode, Div  As Double
        Mode = Value - 100 * (Int(Value / 100))
        Div = (Value - Mode) / 100
        result = Div
    Call Case_Result
        Str100 = str + n100
    
    If Mode > 0 Then
        result = Mode
        Call Case_Result
        Str100 = Str100 + str
              inword = Tk + Str100 + Only
        
        Else
              inword = Tk + Str100 + Only
        End If
End Sub
Private Sub Case1000_99999()
Dim Mode, Div  As Double
        Mode = Value - 1000 * (Int(Value / 1000))
        Div = (Value - Mode) / 1000
        result = Div
    Call Case_Result
        Str1000 = str + n1000
    If Mode > 0 Then
        If Mode >= 1 And Mode < 10 Then
           Value = Mode
           Call Case0_9
           Str1000 = Str1000 + Str1
           inword = Tk + Str1000 + Only
        End If
        
        If Mode >= 10 And Mode < 100 Then
            Value = Mode
            Call Case10_99
            Str1000 = Str1000 + Str10
            inword = Tk + Str1000 + Only
        End If
        
        If Mode >= 100 And Mode < 1000 Then
            Value = Mode
            Call Case100_999
            Str1000 = Str1000 + Str100
            inword = Tk + Str1000 + Only
        End If
    Else
        inword = Tk + Str1000 + Only
    End If
    
End Sub
Private Sub Case100000_9999999()
Dim Mode, Div  As Double
        Mode = Value - 100000 * (Int(Value / 100000))
        Div = (Value - Mode) / 100000
        result = Div
    Call Case_Result
        Str100000 = str + n100000
    If Mode > 0 Then
        If Mode >= 1 And Mode < 10 Then
            Value = Mode
            Call Case0_9
            Str100000 = Str100000 + Str1
            inword = Tk + Str100000 + Only
        End If
        
        If Mode >= 10 And Mode < 100 Then
            Value = Mode
            Call Case10_99
            Str100000 = Str100000 + Str10
            inword = Tk + Str100000 + Only
        End If
        
        If Mode >= 100 And Mode < 1000 Then
            Value = Mode
            Call Case100_999
            Str100000 = Str100000 + Str100
            inword = Tk + Str100000 + Only
        End If
        
        If Mode >= 1000 And Mode < 100000 Then
            Value = Mode
            Call Case1000_99999
            Str100000 = Str100000 + Str1000
            inword = Tk + Str100000 + Only
        End If
    Else
           inword = Tk + Str100000 + Only
        End If
End Sub
Private Sub Case10000000_999999999()
Dim Mode, Div  As Double
        Mode = Value - 10000000 * (Int(Value / 10000000))
        Div = (Value - Mode) / 10000000
        result = Div
    Call Case_Result
        Str10000000 = str + n10000000
    If Mode > 0 Then
        If Mode >= 1 And Mode < 10 Then
            Value = Mode
            Call Case0_9
            Str10000000 = Str10000000 + Str1
            inword = Tk + Str10000000 + Only
        End If
        
        If Mode >= 10 And Mode < 100 Then
            Value = Mode
            Call Case10_99
            Str10000000 = Str10000000 + Str10
            inword = Tk + Str10000000 + Only
        End If
        
        If Mode >= 100 And Mode < 1000 Then
            Value = Mode
            Call Case100_999
            Str10000000 = Str10000000 + Str100
            inword = Tk + Str10000000 + Only
        End If
        
        If Mode >= 1000 And Mode < 100000 Then
            Value = Mode
            Call Case1000_99999
            Str10000000 = Str10000000 + Str1000
            inword = Tk + Str10000000 + Only
        End If
                
        If Mode >= 100000 And Mode < 10000000 Then
            Value = Mode
            Call Case100000_9999999
            Str10000000 = Str10000000 + Str100000
            inword = Tk + Str10000000 + Only
        End If
    Else
        inword = Tk + Str10000000 + Only
        End If
End Sub

Private Sub Invo_No()
Dim ID As String
Dim mid As Integer
    ID = txtInvoice.Text
    mid = 0
    
    Set rs = New ADODB.Recordset
        str = "select * from Sales_Invoice where Invo_No like '" & ID & "'"
        rs.Open str, conn

    If Not rs.EOF Then
        rs.Close
    Else
        
        Set rsU = New ADODB.Recordset
            str = "select * from Sales_Invoice order by Invo_No"
            rsU.Open str, conn
        
        If Not rsU.EOF Then
           rsU.MoveFirst
        
            Do While Not rsU.EOF = True
                mid = mid + 1
                rsU.MoveNext
            Loop
                rsU.Close
                mid = mid + 1
            Else
                rsU.Close
                mid = 1
            End If
            
            txtInvoice.Text = "IE/" & Year(CDate(Today)) & "/" & Format(mid, "000#")
            txtDate.Text = Today
            txtType.Text = ""
            txtName.Text = ""
            txtAddress.Text = ""
            txtMobile.Text = ""
            'txtID.SetFocus
        rs.Close
    ' cmdSales.Enabled = True
End If
End Sub
Private Sub Vendor_Cr()
On Error Resume Next
Dim Prod As String
Prod = rsProd![Product Code]

Set rs = New ADODB.Recordset
        str = "select * from Godown_Master where Prod_Code like '" & Prod & "' And Vendor_Code Like '" & cmbAgent_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
   
        rs!Sale = rs!Sale + Val(rsProd!Qty)
        rs!Stock = rs!Stock - Val(rsProd!Qty)
        'rs!Prod_Price = Val(rsU!Prod_Price)
        rs!Amount = rs!Stock * rs!Prod_Price
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Godown_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Invo_No = txtInvoice.Text
        rsN!Ref_no = txtInvoice.Text
        rsN!Ref_Code = txtID.Text
        rsN!Ref_Name = txtName.Text
        rsN!User_Id = User_Id
        rsN!User_Name = User_Name
        rsN!D_ate = txtDate.Text
        rsN!Prod_Code = rsProd![Product Code]
        rsN!Prod_Type = rsProd![Product Type]
        rsN!Prod_Name = rsProd![Product Name]
        rsN!Prod_Model = rsProd![Description]
        rsN!Purchase = 0
        rsN!Sale = Val(rsProd!Qty)
        rsN!Lift = 0
        rsN!Stock = rs!Stock
        rsN!Prod_Price = rs!Prod_Price
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

            rsN!Vendor_Code = cmbAgent_Code.Text
            rsN!Vendor_Name = cmbAgent_Name.Text
           ' rsN!Vendor_Address = txtVendor_Address.Text
            rsN!Prod_Code = rsProd![Product Code]
'            rsN!Prod_Type = rsProd![Product Type]
            rsN!Prod_Name = rsProd![Product Name]
            rsN!Prod_Model = rsProd![Description]
            rsN!Open_Bal = 0
            rsN!Purchase = 0
            rsN!Sale = Val(rsProd!Qty)
            rsN!Lift = 0
            rsN!Return = 0
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
            rs!Ref_no = txtInvoice.Text
            rs!Ref_Code = txtID.Text
            rs!Ref_Name = txtName.Text
            rs!User_Id = User_Id
            rs!User_Name = User_Name
            rs!D_ate = txtDate.Text
            rs!Prod_Code = rsProd![Product Code]
            rs!Prod_Type = rsProd![Product Type]
            rs!Prod_Name = rsProd![Product Name]
            rs!Prod_Model = rsProd![Description]
            rs!Purchase = 0
            rs!Sale = Val(rsProd!Qty)
            rs!Lift = 0
            rs!Stock = rsN!Stock
            rs!Prod_Price = rsN!Prod_Price
            rs!Com = Val(rsProd![Commission])
            rs!Amount = rsN!Amount
            rs.Update
            rs.Close
    End If
End Sub
Private Sub Vendor_Rtn()
On Error Resume Next
Dim Prod As String
Prod = rsProd![Product Code]

Set rs = New ADODB.Recordset
        str = "select * from Godown_Master where Prod_Code like '" & Prod & "' And Vendor_Code Like '" & cmbAgent_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
   
        rs!Return = rs!Return + Val(rsProd!Qty)
        rs!Stock = rs!Stock + Val(rsProd!Qty)
        'rs!Prod_Price = Val(rsU!Prod_Price)
        rs!Amount = rs!Stock * rs!Prod_Price
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Godown_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Invo_No = txtInvoice.Text
        rsN!Ref_no = txtInvoice.Text
        rsN!Ref_Code = txtID.Text
        rsN!Ref_Name = txtName.Text
        rsN!User_Id = User_Id
        rsN!User_Name = User_Name
        rsN!D_ate = txtDate.Text
        rsN!Prod_Code = rsProd![Product Code]
        rsN!Prod_Type = rsProd![Product Type]
        rsN!Prod_Name = rsProd![Product Name]
        rsN!Prod_Model = rsProd![Description]
        rsN!Purchase = 0
        rsN!Return = Val(rsProd!Qty)
        rsN!Lift = 0
        rsN!Stock = rs!Stock
        rsN!Prod_Price = rs!Prod_Price
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

            rsN!Vendor_Code = cmbAgent_Code.Text
            rsN!Vendor_Name = cmbAgent_Name.Text
           ' rsN!Vendor_Address = txtVendor_Address.Text
            rsN!Prod_Code = rsProd![Product Code]
'            rsN!Prod_Type = rsProd![Product Type]
            rsN!Prod_Name = rsProd![Product Name]
            rsN!Prod_Model = rsProd![Description]
            rsN!Open_Bal = 0
            rsN!Purchase = 0
            rsN!Sale = 0
            rsN!Lift = 0
            rsN!Return = Val(rsProd!Qty)
            rsN!Stock = rsN!Stock + Val(rsProd!Qty)
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
            rs!Ref_no = txtInvoice.Text
            rs!Ref_Code = txtID.Text
            rs!Ref_Name = txtName.Text
            rs!User_Id = User_Id
            rs!User_Name = User_Name
            rs!D_ate = txtDate.Text
            rs!Prod_Code = rsProd![Product Code]
            rs!Prod_Type = rsProd![Product Type]
            rs!Prod_Name = rsProd![Product Name]
            rs!Prod_Model = rsProd![Description]
            rs!Purchase = 0
            rs!Return = Val(rsProd!Qty)
            rs!Lift = 0
            rs!Stock = rsN!Stock
            rs!Prod_Price = rsN!Prod_Price
            rs!Com = Val(rsProd![Commission])
            rs!Amount = rsN!Amount
            rs.Update
            rs.Close
    End If
End Sub

Private Sub Customer_Add()
On Error Resume Next
Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Dr = rs!Dr + Val(rsProd![Total Price]) - Val(rsProd![Commission])
        rs!Balance = rs!Balance + Val(rsProd![Total Price]) - Val(rsProd![Commission])
        rs.Update
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Description = rsProd![Product Name] + "," + rsProd![Description]
            rsN!Dr = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
    Else
            rs.Close
    
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Master", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
        
            rsN!Date = txtDate.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Agent_Code = cmbAgent_Code.Text
            rsN!Agent_Name = cmbAgent_Name.Text
            rsN!Prod_Type = cmbP_Type.Text
            rsN!Open_Bal = 0
            rsN!Dr = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN!Cr = 0
            rsN!Balance = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN.Update
            rsN.Close
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
        
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Description = rsProd![Product Name] + "," + rsProd![Description]
            rsN!Dr = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN!Cr = 0
            rsN!Balance = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN.Update
            rsN.Close
    End If
End Sub
Private Sub Charge_Add()
On Error Resume Next
Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Cr = rs!Cr + Val(txtCharge.Text)
        rs!Balance = rs!Balance - Val(txtCharge.Text)
        rs.Update
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Description = "Discount Paid"
            rsN!Dr = 0
            rsN!Cr = Val(txtCharge.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
    Else
            rs.Close
    
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Master", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
        
            rsN!Date = txtDate.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Open_Bal = 0
            rsN!Dr = 0
            rsN!Cr = Val(txtCharge.Text)
            rsN!Balance = Val(txtCharge.Text)
            rsN.Update
            rsN.Close
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
        
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Description = "Discount Paid"
            rsN!Dr = 0
            rsN!Cr = Val(txtCharge.Text)
            rsN!Balance = Val(txtCharge.Text)
            rsN.Update
            rsN.Close
    End If



    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtCharge.Text)
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtCharge.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtCharge.Text)
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = Val(txtCharge.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

End Sub
Private Sub Invoice_Add()
On Error Resume Next
'Set rs = New ADODB.Recordset
'        str = "select * from Sales_Invoice where Invo_No like '" & txtInvoice.Text & "'"
'        rs.Open str, conn, adOpenDynamic, adLockOptimistic
'
'    If Not rs.EOF Then
'        MsgBox "Duplicate Invoice No. Please Try Another"
'        rs.Close
'        Exit Sub
'    Else
    
        Set rsN = New ADODB.Recordset
            rsN.Open "Sales_Invoice", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            
            rsN!Invo_No = txtInvoice.Text
            rsN!Sale_Type = cmbType.Text
            rsN!D_ate = txtDate.Text
            rsN!Customer = txtID.Text
            rsN!Name = txtName.Text
            rsN!Type = txtType.Text
            rsN!Address = txtAddress.Text
            rsN!Mobile = txtMobile.Text
            rsN!Balance = Val(txtBalance.Text)
            rsN!Amount = Val(txtSales_Due.Text)
            rsN!Net_Due = Val(txtTotal_Due.Text)
            
            rsN!Cash = txtCash.Text
            rsN!Account = txtAccount.Text
            rsN!Bank_Name = cmbBank.Text
            rsN!Branch_Name = cmbBranch.Text
            rsN!Chq_No = txtChq_No.Text
            rsN!Chq_Amnt = txtChq_Amnt.Text
            rsN!Card_Name = cmbCard_Name.Text
            rsN!Card_No = txtCard_No.Text
            rsN!Card_Amnt = txtCard_Amnt.Text
            rsN!Charge_Amnt = txtCharge.Text
            rsN!Terms = txtTerms.Text
            
            rsN.Update
            rsN.Close
            rs.Close
'    End If
End Sub
Private Sub Invoice_Print()
Dim Tran As String

    Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where Invo_No like '" & Tran & "'"
        rs.Open str, conn

 If Not rs.EOF Then
            rs.MoveFirst
                
                Prod_Sl = 0
                netAmnt = 0

            Do While Not rs.EOF
    
            Prod_Sl = Prod_Sl + 1
            
                rsProd.AddNew
                rsProd!sl = Prod_Sl
                rsProd![Goods Code] = rs!Prod_Code
                rsProd![Goods Name] = rs!Prod_Name
                rsProd![Description] = rs!Prod_Model
                rsProd!Qty = rs!Sale
                rsProd![Unit Price] = Format$(Val(rs!Sale_Price), "###0.0000")
                rsProd![Total Price] = Format$(Val(rs!Sale * rs!Sale_Price), "###0.0000")
                rsProd![Commission] = Format$(Val(rs!Com), "###0.0000")
                netAmnt = Round(netAmnt + Val(rs!Sale * rs!Sale_Price))
                
                Set grdProd.DataSource = rsProd
                grdProd.Refresh
                Call Col_Prod
                rs.MoveNext
            
            
            Loop
    
    rs.Close
    
        rsProd.MoveFirst
        Prod_Sl = 0
        netAmnt = 0
    
    Do While Not rsProd.EOF
        Prod_Sl = Prod_Sl + 1
        rsProd!sl = Prod_Sl
        netAmnt = Round(netAmnt + Val(rsProd![Total Price]))
        rsProd.MoveNext
    Loop
        
        strloan = Format$(Val(netAmnt), "###0.00")
        lblNet.Caption = "Total Amount Receivable Tk. " + strloan
    
    Set grdProd.DataSource = rsProd
        grdProd.Refresh
   
    Call Col_Prod
    
    
        Set rs = New ADODB.Recordset
               str = "select * from Cash_Book where MR_No like '" & Tran & "'"
               rs.Open str, conn
           
           If Not rs.EOF Then
                rs.MoveFirst
           
           Do While Not rs.EOF
           
               txtCash.Text = Format$(Val(txtCash.Text) + Val(rs!Dr), "###0.00")
               txt1000.Text = Val(txt1000.Text) + Val(rs!Tk1000)
               txt500.Text = Val(txt500.Text) + Val(rs!Tk500)
               txt100.Text = Val(txt100.Text) + Val(rs!Tk100)
               txt50.Text = Val(txt50.Text) + Val(rs!Tk50)
               txt20.Text = Val(txt20.Text) + Val(rs!Tk20)
               txt10.Text = Val(txt10.Text) + Val(rs!Tk10)
               txt5.Text = Val(txt5.Text) + Val(rs!Tk5)
               txt2.Text = Val(txt2.Text) + Val(rs!Tk2)
               txt1.Text = Val(txt1.Text) + Val(rs!Tk1)
            rs.MoveNext
            Loop
               Call Total_Cash
               rs.Close
           Else
               rs.Close
               txtCash.Text = 0
               txt1000.Text = 0
               txt500.Text = 0
               txt100.Text = 0
               txt50.Text = 0
               txt20.Text = 0
               txt10.Text = 0
               txt5.Text = 0
               txt2.Text = 0
               txt1.Text = 0
           End If
    
            Set rs = New ADODB.Recordset
               str = "select * from Bank_Tran where MR_No like '" & Tran & "'"
               rs.Open str, conn
           
           If Not rs.EOF Then
               txtChq_Amnt.Text = Format$(Val(rs!Dr), "###0.00")
               txtAccount.Text = rs!AC_No
               cmbBank.Text = rs!Bank_Name
               cmbBranch.Text = rs!Branch_Name
               txtChq_No.Text = rs!Chq_No
               rs.Close
           Else
               rs.Close
               txtChq_Amnt.Text = 0
               txtAccount.Text = ""
               cmbBank.Text = ""
               cmbBranch.Text = ""
               txtChq_No.Text = ""
           End If
            
    
        cmdPrint.Enabled = True
        cmdSales.Enabled = False
        cmdDelete.Enabled = True
    Else
    rs.Close
    'Call clearTextboxes
    'txtInvoice.Text = Tran
    txtDate.Text = Today
    Exit Sub
  End If
End Sub
Private Sub Case_RSPLoad_Issue()
    Set rsN = New ADODB.Recordset
        str = "select * from Customer_Master where Customer like '" & txtID.Text & "' and Prod_Code like '" & rsProd![Goods Code] & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
 If Not rsN.EOF Then
'------------------------------------------------------------------------
        rsN!D_ate = txtDate.Text
        rsN!Receive = rsN!Receive + Val(rsProd![Total Price])
        rsN!Close_Bal = rsN!Close_Bal + Val(rsProd![Total Price])
        rsN!Prod_Price = Val(rsProd![Total Price])
        'rsN!Sale_Amount = Val(rsProd![Total Amount])
        'rsN!Com = Val(rsProd![Commission])
        rsN!Due_Amount = rsN!Due_Amount + Val(rsProd![Total Price])
        rsN.Update
        
        Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Receive = rsProd![Total Price]
            rs!Balance = rsN!Close_Bal
            rs!Com = 0
            rs!Amount = 0
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
Else
'-----------------------------------------------------------------------
       rsN.Close
       
       Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Master", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
       
            rsN!D_ate = txtDate.Text
            rsN!Customer = txtID.Text
            rsN!Name = txtName.Text
            rsN!Type = txtType.Text
            rsN!Prod_Code = rsProd![Goods Code]
            rsN!Prod_Name = rsProd![Goods Name]
            rsN!Prod_Model = rsProd![Description]
            rsN!Description = rsProd![Goods Name] + "-" + rsProd![Description]
            rsN!Open_Bal = 0
            rsN!Receive = Val(rsProd![Total Price])
            rsN!Sale = 0
            rsN!Return = 0
            rsN!Close_Bal = Val(rsProd![Total Price])
            rsN!Prod_Price = Val(rsProd![Total Price])
            rsN!Sale_Amount = 0
            rsN!Com = 0
            rsN!Due_Amount = Val(rsProd![Total Price])
            rsN!C_lose = "N"
            rsN.Update
            
        Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Receive = rsProd![Total Price]
            rs!Balance = rsN!Close_Bal
            rs!Com = 0
            rs!Amount = rsProd![Total Price]
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
End If
End Sub
Private Sub Case_RSPLoad_Sale()
    Set rsN = New ADODB.Recordset
        str = "select * from Customer_Master where Customer like '" & txtID.Text & "' and Prod_Code like '" & rsProd![Goods Code] & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
 If Not rsN.EOF Then
'------------------------------------------------------------------------
            rsN!D_ate = txtDate.Text
            rsN!Sale = rsN!Sale + Val(rsProd![Total Price])
            rsN!Close_Bal = rsN!Close_Bal - Val(rsProd![Total Price])
            rsN!Prod_Price = Val(rsProd![Total Price])
            rsN!Sale_Amount = Val(rsProd![Total Price])
            rsN!Com = Val(rsProd![Commission])
            rsN!Due_Amount = rsN!Due_Amount - Val(rsProd![Total Price])
            rsN.Update
        
        Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Sale = rsProd![Total Price]
            rs!Balance = rsN!Close_Bal
            rs!Com = 0
            rs!Amount = 0
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
Else
'-----------------------------------------------------------------------
       rsN.Close
End If
End Sub

Private Sub CustomerProd_Rev()
    Set rsN = New ADODB.Recordset
        str = "select * from Customer_Master where Customer like '" & txtID.Text & "' and Prod_Code like '" & rsProd![Goods Code] & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic

If Not rsN.EOF Then

    If cmbType.Text = "Sale" Then
        rsN!D_ate = txtDate.Text
        rsN!Sale = rsN!Sale - Val(rsProd!Qty)
        'rsN!Receive = rsN!Receive + Val(rsProd!Qty)
        rsN!Close_Bal = rsN!Close_Bal + Val(rsProd!Qty)
        rsN!Prod_Price = Val(rsProd![Unit Price])
        rsN!Sale_Amount = rsN!Sale_Amount - Val(rsProd![Total Price])
        'rsN!Com = Val(rsProd![Commission])
        rsN!Due_Amount = rsN!Due_Amount + Val(rsProd![Total Price])
        rsN.Update
        
        Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Return = rsProd!Qty
            rs!Payment = 0
            rs!Balance = rsN!Close_Bal
            'rs!Com = rsProd![Commission]
            'rs!Amount = rsProd![Total Price]
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
    End If
Else
rsN.Close
End If
End Sub
Private Sub Cash_Rev()
           
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
        rsN!Description = txtID.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtCash.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        
        
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Cash_Cr = rsU!Cash_Cr + Val(txtCash.Text)
        rsU!Cash_Close = rsU!Cash_Close - Val(txtCash.Text)
        rsU.Update
        rsU.Close
        
        Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!MR_No = txtID.Text
        rsN!Name = txtName.Text
        rsN!Description = "Entry Reversed"
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



End Sub
Private Sub Bank_Rev()
           
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
        rsN!Description = "CID-" & txtID.Text
        rsN!Dr = 0
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
        str = "select * from Bank_Master where AC_No like '" & txtAccount.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs!Date = txtDate.Text
        rs!Withdraw = rs!Withdraw + Val(txtChq_Amnt.Text)
        rs.Update
        
        Set rsN = New ADODB.Recordset
        rsN.Open "Bank_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = txtAccount.Text
        rsN!Bank_Name = cmbBank.Text
        rsN!Branch_Name = cmbBranch.Text
        rsN!Chq_No = "CN" & txtChq_No.Text
        rsN!MR_No = txtID.Text
        rsN!Description = "CID-" & txtID.Text
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
     
End Sub
Private Sub GL_Rev()
If rsProd![Goods Code] = 101 Then
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100103 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + (Val(rsProd![Total Price]) + Val(rsProd![Commission]))
        rs!Date = txtDate.Text
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "PCode-" & rsProd![Goods Code]
        rsN!Dr = Val(rsProd![Total Price]) + Val(rsProd![Commission])
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
        
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
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
        rsN!Description = "CID-" & txtID.Text
        rsN!Dr = 0
        rsN!Cr = Val(rsProd![Commission])
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
Else
    
    If rsProd![Goods Code] = "BKS01" Then
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100105 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + (Val(rsProd![Total Price]) + Val(rsProd![Commission]))
        rs!Date = txtDate.Text
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "PCode-" & rsProd![Goods Code]
        rsN!Dr = Val(rsProd![Total Price]) + Val(rsProd![Commission])
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
       
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
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
        rsN!Description = "CID-" & txtID.Text
        rsN!Dr = 0
        rsN!Cr = Val(rsProd![Commission])
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
Else
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + (Val(rsU!Prod_Price) * Val(rsProd!Qty))
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "PCode-" & rsProd![Goods Code]
        rsN!Dr = rsProd![Total Price]
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100141 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - (Val(rsProd![Total Price]) - (Val(rsU!Prod_Price) * Val(rsProd!Qty)))
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "PID-" & rsProd![Goods Code]
        rsN!Dr = (rsProd![Total Price] - (Val(rsU!Prod_Price) * Val(rsProd!Qty)))
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    End If
End If
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

    If Not rsN.EOF Then
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbBank.AddItem rsN!Bank_Name
        rsN.MoveNext
        Loop
        rsN.Close
        Else
        rsN.Close
        End If
End Sub
Private Sub Br_Name()
        cmbBranch.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Branch_Name FROM Bank_Master"
        rsN.Open str, conn
      If Not rsN.EOF Then
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbBranch.AddItem rsN!Branch_Name
        rsN.MoveNext
        Loop
        rsN.Close
          Else
        rsN.Close
        End If
End Sub
Private Sub Bank_Cr()
Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
    
        rs!Cr = rs!Cr + Val(txtChq_Amnt.Text)
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs.Update
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Description = "Chq Payment-Thanks"
            rsN!Dr = 0
            rsN!Cr = Val(txtChq_Amnt.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close


Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
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
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

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
        rsN!Description = "CID-" & txtID.Text
        rsN!Dr = Val(txtChq_Amnt.Text)
        rsN!Cr = 0
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
        str = "select * from Bank_Master where AC_No like '" & txtAccount.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtChq_Amnt.Text)
        rs!Date = txtDate.Text
        rs!Deposit = rs!Deposit + Val(txtChq_Amnt.Text)
        rs.Update
        
        Set rsN = New ADODB.Recordset
        rsN.Open "Bank_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = txtAccount.Text
        rsN!Bank_Name = cmbBank.Text
        rsN!Branch_Name = cmbBranch.Text
        rsN!Chq_No = "CN" & txtChq_No.Text
        rsN!MR_No = txtID.Text
        rsN!Description = "CID-" & txtID.Text
        rsN!Dr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
End Sub
Private Sub Card_Cr()
Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
    
        rs!Cr = rs!Cr + Val(txtChq_Amnt.Text)
        rs!Balance = rs!Balance + Val(txtCard_Amnt.Text)
        rs.Update
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Description = "Card Payment-Thanks"
            rsN!Dr = 0
            rsN!Cr = Val(txtCard_Amnt.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close



Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100103 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtCard_Amnt.Text)
        rs!Date = txtDate.Text
        rs.Update
            
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "CID-" & txtID.Text & "-" & txtCard_No.Text
        rsN!Dr = Val(txtCard_Amnt.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
        
End Sub
Private Sub Credit_Sale()
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100107 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + (Val(netAmnt) - Val(Total))
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "CID:" & txtID.Text
        rsN!Dr = (Val(netAmnt) - Val(Total))
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
End Sub

Private Sub Cash_Cr()
    Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
    
        rs!Cr = rs!Cr + Val(txtCash.Text)
        rs!Balance = rs!Balance - Val(txtCash.Text)
        rs.Update
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Description = "Payment Received"
            rsN!Dr = 0
            rsN!Cr = Val(txtCash.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
    
    
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
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
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtCash.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    
    
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
        rsN!Description = "CID-" & txtID.Text
        rsN!Dr = Val(txtCash.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Cash_Dr = rsU!Cash_Dr + Val(txtCash.Text)
        rsU!Cash_Close = rsU!Cash_Close + Val(txtCash.Text)
        rsU.Update
        rsU.Close
        
        Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!MR_No = txtInvoice.Text
        rsN!Name = txtName.Text
        rsN!Description = "Sale Amount Received"
        rsN!Dr = Val(txtCash.Text)
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

'Cash Close================================================================

 
End Sub
Private Sub GL_Prof()
If Val(rsU!Prod_Price) = 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100175 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + (rsProd![Total Price] - (Val(rsU!Prod_Price) * Val(rsProd!Qty)))
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "PID-" & rsProd![Product Code]
        rsN!Dr = 0
        rsN!Cr = (rsProd![Total Price] - (Val(rsU!Prod_Price) * Val(rsProd!Qty)))
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    Else
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100141 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + (rsProd![Total Price] - (Val(rsU!Prod_Price) * Val(rsProd!Qty)))
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "PID-" & rsProd![Product Code]
        rsN!Dr = 0
        rsN!Cr = (rsProd![Total Price] - (Val(rsU!Prod_Price) * Val(rsProd!Qty)))
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    
End Sub
Private Sub GL_Issue()
'Load GL Tran Strart=========================================
  
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100103 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - (Val(rsU!Prod_Price) * Val(rsProd!Qty))
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = (Val(rsU!Prod_Price) * Val(rsProd!Qty))
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(rsProd![Total Price])
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rsU!Ref_no
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = 0
        rsN!Cr = Val(rsProd![Total Price])
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close



Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
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
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = Val(rsProd![Commission])
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close


Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(rsProd![Commission])
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = 0
        rsN!Cr = Val(rsProd![Commission])
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
End Sub
Private Sub Prof_Rtn()
        
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100141 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - (rsProd![Total Price] - (Val(rsU!Prod_Price) * Val(rsProd!Qty)))
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "PID-" & rsProd![Product Code]
        rsN!Dr = (rsProd![Total Price] - (Val(rsU!Prod_Price) * Val(rsProd!Qty)))
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
End Sub

Private Sub GL_Return()
'Load GL Tran Strart=========================================
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100103 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + (Val(rsU!Prod_Price) * Val(rsProd!Qty))
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = 0
        rsN!Cr = (Val(rsU!Prod_Price) * Val(rsProd!Qty))
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(rsProd![Total Price])
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rsU!Ref_no
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = Val(rsProd![Total Price])
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close



Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(rsProd![Commission])
        rs!Date = txtDate.Text
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = Val(rsProd![Commission])
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close


    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
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
        rsN!Description = "INV-" & txtInvoice.Text
        rsN!Dr = 0
        rsN!Cr = Val(rsProd![Commission])
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
End Sub

Private Sub Case_DSRLoad_Issue_Sale()
Set rsN = New ADODB.Recordset
        str = "select * from Customer_Master where Customer like '" & txtID.Text & "' and Prod_Code like '" & rsProd![Goods Code] & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
 If Not rsN.EOF Then
'------------------------------------------------------------------------
        rsN!D_ate = txtDate.Text
        rsN!Receive = rsN!Receive + Val(rsProd![Total Price])
        rsN!Close_Bal = rsN!Close_Bal + Val(rsProd![Total Price])
        rsN!Prod_Price = Val(rsProd![Total Price])
        'rsN!Sale_Amount = Val(rsProd![Total Amount])
        'rsN!Com = Val(rsProd![Commission])
        rsN!Due_Amount = rsN!Due_Amount + Val(rsProd![Total Price])
        rsN.Update
        
        Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Receive = rsProd![Total Price]
            rs!Balance = rsN!Close_Bal
            rs!Com = 0
            rs!Amount = 0
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
Else
'-----------------------------------------------------------------------
       rsN.Close
       
       Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Master", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
       
            rsN!D_ate = txtDate.Text
            rsN!Customer = txtID.Text
            rsN!Name = txtName.Text
            rsN!Type = txtType.Text
            rsN!Prod_Code = rsProd![Goods Code]
            rsN!Prod_Name = rsProd![Goods Name]
            rsN!Prod_Model = rsProd![Description]
            rsN!Description = rsProd![Goods Name] + "-" + rsProd![Description]
            rsN!Open_Bal = 0
            rsN!Receive = Val(rsProd![Total Price])
            rsN!Sale = 0
            rsN!Return = 0
            rsN!Close_Bal = Val(rsProd![Total Price])
            rsN!Prod_Price = Val(rsProd![Total Price])
            rsN!Sale_Amount = 0
            rsN!Com = 0
            rsN!Due_Amount = Val(rsProd![Total Price])
            rsN!C_lose = "N"
            rsN.Update
            
        Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Receive = rsProd![Total Price]
            rs!Balance = rsN!Close_Bal
            rs!Com = 0
            rs!Amount = rsProd![Total Price]
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
End If

Set rsN = New ADODB.Recordset
        str = "select * from Customer_Master where Customer like '" & txtID.Text & "' and Prod_Code like '" & rsProd![Goods Code] & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
 If Not rsN.EOF Then
'------------------------------------------------------------------------
            rsN!D_ate = txtDate.Text
            rsN!Sale = rsN!Sale + Val(rsProd![Total Price])
            rsN!Close_Bal = rsN!Close_Bal - Val(rsProd![Total Price])
            rsN!Prod_Price = Val(rsProd![Total Price])
            rsN!Sale_Amount = Val(rsProd![Total Price])
            rsN!Com = Val(rsProd![Commission])
            rsN!Due_Amount = rsN!Due_Amount - Val(rsProd![Total Price])
            rsN.Update
        
        Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Sale = rsProd![Total Price]
            rs!Balance = rsN!Close_Bal
            rs!Com = 0
            rs!Amount = 0
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
Else
'-----------------------------------------------------------------------
       rsN.Close
End If

End Sub
Private Sub CustomerLoad_Rev()
    Set rsN = New ADODB.Recordset
        str = "select * from Customer_Master where Customer like '" & txtID.Text & "' and Prod_Code like '" & rsProd![Goods Code] & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
 If Not rsN.EOF Then
        rsN!D_ate = txtDate.Text
        rsN!Receive = rsN!Receive - (Val(rsProd![Total Price]) + Val(rsProd![Commission]))
        rsN!Sale = rsN!Sale - (Val(rsProd![Total Price]) + Val(rsProd![Commission]))
        'rsN!Close_Bal = rsN!Close_Bal - Val(rsProd![Total Price])
        rsN!Prod_Price = Val(rsProd![Total Price])
        rsN!Sale_Amount = rsN!Sale_Amount - Val(rsProd![Total Price])
        rsN!Com = rsN!Com - Val(rsProd![Commission])
        'rsN!Due_Amount = rsN!Due_Amount - Val(rsProd![Total Price])
        rsN.Update
        
        Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Return = rsProd![Total Price]
            rs!Balance = rsN!Close_Bal
            rs!Com = 0
            rs!Amount = 0
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
    End If
End Sub
Private Sub CustomerProd_Update()
    On Error Resume Next
    Set rsN = New ADODB.Recordset
        str = "select * from Customer_Master where Customer like '" & txtID.Text & "' and Prod_Code like '" & rsProd![Goods Code] & "'"
'        rsN.Open str, conn, adOpenDynamic, adLockOptimistic

'If Not rsN.EOF Then

    If cmbType.Text = "Cash Sale" Then
'        rsN!Date = txtDate.Text
'        rsN!Receive = rsN!Receive + Val(rsProd!Qty)
        'rsN!Close_Bal = rsN!Close_Bal + Val(rsProd!Qty)
        rsN!Prod_Price = Val(rsProd![Unit Price])
        'rsN!Sale_Amount = Val(rsProd![Total Amount])
        'rsN!Com = Val(rsProd![Commission])
        rsN!Due_Amount = rsN!Due_Amount + Val(rsProd![Total Price])
        rsN.Update
        
     Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Receive = rsProd!Qty
            rs!Balance = rsN!Close_Bal
            rs!Com = rsProd![Commission]
            rs!Amount = rsProd![Total Price]
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
    Else
    
    If cmbType.Text = "Return" Then
        rsN!D_ate = txtDate.Text
        rsN!Return = rsN!Return + Val(rsProd!Qty)
        rsN!Close_Bal = rsN!Close_Bal - Val(rsProd!Qty)
        rsN!Prod_Price = Val(rsProd![Unit Price])
        'rsN!Sale_Amount = Val(rsProd![Total Amount])
        'rsN!Com = Val(rsProd![Commission])
        rsN!Due_Amount = rsN!Due_Amount - Val(rsProd![Total Price])
        rsN.Update
        
     Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Return = rsProd!Qty
            rs!Balance = rsN!Close_Bal
            rs!Com = rsProd![Commission]
            rs!Amount = rsProd![Total Price]
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
    Else
    
    If cmbType.Text = "Sale" Then
        rsN!D_ate = txtDate.Text
        rsN!Sale = rsN!Sale + Val(rsProd!Qty)
        rsN!Close_Bal = rsN!Close_Bal - Val(rsProd!Qty)
        rsN!Prod_Price = Val(rsProd![Unit Price])
        rsN!Sale_Amount = rsN!Sale_Amount + Val(rsProd![Total Price])
        'rsN!Com = Val(rsProd![Commission])
        rsN!Due_Amount = rsN!Due_Amount - Val(rsProd![Total Price])
        rsN.Update
        
        Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Sale = rsProd!Qty
            rs!Payment = rsProd![Total Price]
            rs!Balance = rsN!Close_Bal
            rs!Com = rsProd![Commission]
            rs!Amount = rsProd![Total Price]
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
    End If
    End If
    End If
    
'Else
    
    If cmbType.Text = "Cash Sale" Then
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Customer_Master", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!D_ate = txtDate.Text
        rsN!Customer = txtID.Text
        rsN!Name = txtName.Text
        rsN!Type = txtType.Text
        rsN!Prod_Code = rsProd![Goods Code]
        rsN!Prod_Name = rsProd![Goods Name]
        rsN!Prod_Model = rsProd![Description]
        rsN!Description = rsProd![Goods Name] + "-" + rsProd![Description]
        rsN!Open_Bal = 0
        rsN!Receive = Val(rsProd!Qty)
        rsN!Sale = 0
        rsN!Return = 0
        rsN!Close_Bal = Val(rsProd!Qty)
        rsN!Prod_Price = Val(rsProd![Unit Price])
        'rsN!Sale_Amount = Val(rsProd![Total Amount])
        'rsN!Com = Val(rsProd![Commission])
        rsN!Due_Amount = Val(rsProd![Total Price])
        rsN!C_lose = "N"
        rsN.Update
        
        Set rs = New ADODB.Recordset
            rs.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rs.AddNew
            
            rs!D_ate = txtDate.Text
            rs!Customer = txtID.Text
            rs!Name = txtName.Text
            rs!Type = txtType.Text
            rs!Description = rsProd![Goods Name] + "," + rsProd!Description
            rs!Qty = rsProd!Qty
            rs!Unit_Price = rsProd![Unit Price]
            rs!Receive = rsProd!Qty
            rs!Balance = rsN!Close_Bal
            rs!Com = rsProd![Commission]
            rs!Amount = rsProd![Total Price]
            rs!Net_Due = rsN!Due_Amount
            rs.Update
            rs.Close
            rsN.Close
   ' End If
End If
End Sub
Private Sub Customer_Rtn()
On Error Resume Next
Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Cr = rs!Cr + Val(rsProd![Total Price]) - Val(rsProd![Commission])
        rs!Balance = rs!Balance - (Val(rsProd![Total Price]) - Val(rsProd![Commission]))
        rs.Update
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Description = rsProd![Product Name] + "," + rsProd![Description]
            rsN!Dr = 0
            rsN!Cr = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
    Else
            rs.Close
    
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Master", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
        
            rsN!Date = txtDate.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Open_Bal = 0
            rsN!Dr = 0
            rsN!Cr = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN!Balance = -Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN.Update
            rsN.Close
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
        
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtAddress.Text
            rsN!Customer_Type = txtType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Description = rsProd![Product Name] + "," + rsProd![Description]
            rsN!Dr = 0
            rsN!Cr = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN!Balance = -Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN.Update
            rsN.Close
    End If
End Sub

Private Sub Prod_Tran()
If cmbType.Text = "Cash Sale" Then
           
          
           Set rsN = New ADODB.Recordset
               rsN.Open "Prod_Tran", conn, adOpenDynamic, adLockOptimistic, -1
               rsN.AddNew
               
               rsN!Invo_No = txtInvoice.Text
               rsN!Ref_no = rsU!Ref_no
               rsN!Ref_Code = txtID.Text
               rsN!Ref_Name = txtName.Text
               rsN!D_ate = txtDate.Text
               rsN!Prod_Code = rsProd![Product Code]
               rsN!Prod_Name = rsProd![Product Name]
               rsN!Prod_Model = rsProd![Description]
               
               rsN!Purchase = 0
               rsN!Sale = Val(rsProd!Qty)
               rsN!Lift = 0
               rsN!Stock = rsU!Stock
               rsN!Prod_Price = rsU!Prod_Price
               rsN!Com = Val(rsProd![Commission])
               rsN!Amount = (Val(rsProd![Total Price]) - Val(rsProd![Commission]))
               rsN!Sale_Price = rsProd![Unit Price]
               rsN!User_Id = User_Id
               rsN!User_Name = User_Name
               rsN.Update
               rsN.Close
End If

If cmbType.Text = "Credit Sale" Then
           
          
           Set rsN = New ADODB.Recordset
               rsN.Open "Prod_Tran", conn, adOpenDynamic, adLockOptimistic, -1
               rsN.AddNew
               
               rsN!Invo_No = txtInvoice.Text
               rsN!Ref_no = rsU!Ref_no
               rsN!Ref_Code = txtID.Text
               rsN!Ref_Name = txtName.Text
               rsN!D_ate = txtDate.Text
               rsN!Prod_Code = rsProd![Product Code]
               rsN!Prod_Name = rsProd![Product Name]
               rsN!Prod_Model = rsProd![Description]
               
               rsN!Purchase = 0
               rsN!Sale = Val(rsProd!Qty)
               rsN!Lift = 0
               rsN!Stock = rsU!Stock
               rsN!Prod_Price = rsU!Prod_Price
               rsN!Com = Val(rsProd![Commission])
               rsN!Amount = (Val(rsProd![Total Price]) - Val(rsProd![Commission]))
               rsN!Sale_Price = rsProd![Unit Price]
               rsN!User_Id = User_Id
               rsN!User_Name = User_Name
               rsN.Update
               rsN.Close
End If

If cmbType.Text = "Sales Return" Then
        Set rsN = New ADODB.Recordset
               rsN.Open "Prod_Tran", conn, adOpenDynamic, adLockOptimistic, -1
               rsN.AddNew
               
               rsN!Invo_No = txtInvoice.Text
               rsN!Ref_no = rsU!Ref_no
               rsN!Ref_Code = txtID.Text
               rsN!Ref_Name = txtName.Text
               rsN!D_ate = txtDate.Text
               rsN!Prod_Code = rsProd![Product Code]
               rsN!Prod_Name = rsProd![Product Name]
               rsN!Prod_Model = rsProd![Description]
               
               rsN!Purchase = 0
               rsN!Sale = 0
               rsN!Return = Val(rsProd!Qty)
               rsN!Stock = rsU!Stock
               rsN!Prod_Price = 0
               rsN!Com = Val(rsProd![Commission])
               rsN!Amount = (Val(rsProd![Total Price]) - Val(rsProd![Commission]))
               rsN!Sale_Price = rsProd![Unit Price]
               rsN.Update
               rsN.Close
End If
End Sub
Private Sub Col_Prod()
With grdProd
    .Columns(0).Width = 500
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1500
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 2000
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 2000
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 500
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1200
    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 1200
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 1200
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 1200
    .Columns(8).Alignment = dbgRight
End With
End Sub
Private Sub Cust_Id()
        txtID.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer_Code FROM Customer_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtID.AddItem rs!Customer_Code
        rs.MoveNext
        Loop
        rs.Close
    Else
    rs.Close
    End If
End Sub
Private Sub Cust_Name()
        txtName.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer_Name FROM Customer_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtName.AddItem rs!Customer_Name
        rs.MoveNext
        Loop
        rs.Close
    Else
    rs.Close
    End If
End Sub
Private Sub grd_Prod()
Prod_Sl = 0
netAmnt = 0
strloan = Format$(Val(netAmnt), "###0.00")
lblNet.Caption = "Net Amount Receiveable Tk. " + strloan
    ' Create an initial recordset, just for demonstration purposes,
    ' and assign it to the DataGrid control's DataSource property.
    Set rsProd = New ADODB.Recordset
    With rsProd
        .Fields.Append "Sl", adBSTR
        .Fields.Append "Product Code", adBSTR
        .Fields.Append "Product Name", adBSTR
        .Fields.Append "Description", adBSTR
        .Fields.Append "Qty", adBSTR
        .Fields.Append "Unit Price", adBSTR
        .Fields.Append "Total Price", adBSTR
        .Fields.Append "Commission", adBSTR
        .Fields.Append "Cost", adBSTR
        .Open
   
    End With
    Set grdProd.DataSource = rsProd
    
End Sub
Private Sub Prod_Code()
On Error Resume Next
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
        If Not rs.EOF Then
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbProd_Name.AddItem rs!Prod_Name
        rs.MoveNext
        Loop
        rs.Close
    Else
    rs.Close
    End If
End Sub
Private Sub Agent_Id()
On Error Resume Next
        cmbAgent_Code.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Vendor_Code FROM Godown_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbAgent_Code.AddItem rsN!Vendor_Code
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub Agent_Name()
   cmbAgent_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Vendor_Name FROM Godown_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbAgent_Name.AddItem rs!Vendor_Name
        rs.MoveNext
        Loop
        rs.Close
    Else
    rs.Close
    End If
End Sub
Private Sub Prod_Type()
        cmbP_Type.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name FROM Prod_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbP_Type.AddItem rs!Prod_Name
        rs.MoveNext
        Loop
        rs.Close
    Else
    rs.Close
    End If
End Sub

Private Sub Prod_Model()
        cmbProd_Model.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Model FROM Prod_Master"
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



Private Sub ComboType()
        cmbType.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT Type FROM Parameter GROUP BY Type ORDER BY Type ASC"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbType.AddItem rs!Type
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
End Sub
Private Sub Type_Name()
If cmbType.Text = "" Then
Exit Sub
End If
On Error Resume Next
 Dim DPS_Type As String
 DPS_Type = cmbType.Text
   Set rs = New ADODB.Recordset
        str = "select DISTINCT Type, Description from Parameter where Type like '" & DPS_Type & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
        On Error Resume Next
       ' lblDescription.Caption = rs!Description
        rs.Close
    End If
End Sub
Private Sub Cash_Amnt()
txtCash.Text = (Val(txt1000.Text) * 1000) + (Val(txt500.Text) * 500) + (Val(txt100.Text) * 100) + (Val(txt50.Text) * 50) + (Val(txt20.Text * 20)) + (Val(txt10.Text) * 10) + (Val(txt5.Text) * 5) + (Val(txt2.Text) * 2) + (Val(txt1.Text) * 1)
txtCash.Text = Format(Val(txtCash.Text), "###0.00")
End Sub

Private Sub Total_Cash()
    Total = Round(Val(txtCash.Text) + Val(txtChq_Amnt.Text) + Val(txtCard_Amnt.Text))
    txtCash.Text = Format(Total - (Val(txtChq_Amnt.Text) + Val(txtCard_Amnt.Text)), "###0.00")
    txtChq_Amnt.Text = Format(Total - (Val(txtCash.Text) + Val(txtCard_Amnt.Text)), "###0.00")
    txtCard_Amnt.Text = Format(Total - (Val(txtCash.Text) + Val(txtChq_Amnt.Text)), "###0.00")
    lblPmt.Caption = "Total Tk." + Format$(Val(Total), "###0.00")
End Sub

Private Sub clearTextboxes()
        cst = 0
        txtDate.Text = ""
        txtInvoice.Text = ""
        cmbType.Text = ""
        txtID.Text = ""
        txtName.Text = ""
        txtType.Text = ""
        txtAddress.Text = ""
        cmbAgent_Code.Text = ""
        cmbAgent_Name.Text = ""
        cmbP_Type.Text = ""
        txtMobile.Text = ""
        txtBalance.Text = "0.00"
        txtSales_Due.Text = "0.00"
        txtTotal_Due.Text = "0.00"
       
        cmbProd_Name.Text = ""
        cmbProd_Model.Text = ""
        txtProd_Sl.Text = ""
        txtProd_Cost.Text = "0.00"
        txtQty.Text = "0"
         txtStock.Text = "0"
        txtProd_Price.Text = "0.00"
        txtCommission.Text = "0.00"
        txtPercent.Text = "0.00"
        txtTerms.Text = ""
        txtCharge.Text = ""
        
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
        txtAccount.Text = ""
        cmbBank.Text = ""
        cmbBranch.Text = ""
        txtChq_No.Text = ""
        txtChq_Amnt.Text = "0.00"
        cmbCard_Name.Text = ""
        txtCard_No.Text = ""
        txtCard_Amnt.Text = "0.00"
'        txtCom_Des.Text = ""
'        txtCom_Amnt = "0.00"
End Sub

Private Sub cmbBrand_Change()

End Sub

Private Sub cmbOrigin_Change()

End Sub

Private Sub cmbColor_Change()

End Sub

Private Sub cmbTerm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txt1000.SelStart = 0
txt1000.SelLength = Len(txt1000.Text)
txt1000.SetFocus
End If
End Sub

Private Sub cmbAgent_Code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbAgent_Name.SetFocus
End If
End Sub

Private Sub cmbAgent_Code_LostFocus()
Dim ID As String
    ID = cmbAgent_Code.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Godown_Master where Vendor_Code like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
'        Call clearTextboxes
        cmbAgent_Code.Text = rs!Vendor_Code
        cmbAgent_Name.Text = rs!Vendor_Name
        rs.Close
        cmbAgent_Name.SetFocus
    Else
    
    Exit Sub
    End If

End Sub

Private Sub cmbAgent_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Sl.SetFocus
End If
End Sub

Private Sub cmbAgent_Name_LostFocus()
Dim ID As String
    ID = cmbAgent_Name.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Godown_Master where Vendor_Name like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
'        Call clearTextboxes
        cmbAgent_Code.Text = rs!Vendor_Code
        cmbAgent_Name.Text = rs!Vendor_Name
        rs.Close
        'cmbAgent_Name.SetFocus
    Else
    
    Exit Sub
    End If
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

Private Sub cmbCard_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
txtCard_No.SelStart = 0
txtCard_No.SelLength = Len(txtCard_No.Text)
txtCard_No.SetFocus
End If
End Sub

Private Sub cmbCard_Name_LostFocus()
On Error Resume Next
If cmbCard_Name.Text = "" Then
cmdSales.SetFocus
Exit Sub
End If
End Sub

Private Sub cmbP_Type_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Sl.SetFocus
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
On Error Resume Next
Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Godown_Master where Prod_Name Like '" & cmbProd_Name.Text & "' AND Prod_Model Like '" & cmbProd_Model.Text & "' And Vendor_Code Like '" & cmbAgent_Code.Text & "' order by Prod_code"
        rsN.Open str, conn

            If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            'cmbBrand.Text = rsN!Brand_Name
            'cmbOrigin.Text = rsN!Origin
            'cmbColor.Text = rsN!Color
'            txtPercent.Text = 0
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtCost.Text = rsN!Prod_Price
            If txtType.Text = "General" Then
            txtProd_Cost.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            End If
            If txtType.Text = "Dealer" Then
            txtProd_Cost.Text = Format$(Val(rsN!Dealer_Price), "###0.00")
            End If

            If txtType.Text = "Retail" Then
            txtProd_Cost.Text = Format$(Val(rsN!Retail_Price), "###0.00")
            End If
            rsN.Close
            'txtSl.SelStart = 0
          '  txtSl.SelLength = Len(txtSl.Text)
           ' txtSl.SetFocus

        Else

        MsgBox "Invalid Goods Code!", vbCritical, "Sales Info!"
            rsN.Close
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

Private Sub cmbType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtID.SetFocus
End If
End Sub
Private Sub cmbType_LostFocus()
Call Type_Name
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
    
    Set rs = New ADODB.Recordset
        str = "select * from Customer_Info where Customer like '" & txtID.Text & "'"
        rs.Open str, conn
        
     If Not rs.EOF Then
        MsgBox "Customer Already Exist!", vbCritical, "Customer Info!"
        rs.Close
        Exit Sub
    Else
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Customer_Info", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
    
        
        rsN!Customer = txtID.Text
        rsN!Type = cmbType.Text
        'rsN!Term = cmbTerm.Text
        rsN!Com = txtCommission.Text
        rsN!Balance = 0
        rsN!D_ate = txtDate.Text
        
        rsN!Name = txtName.Text
        rsN!Present_Address = txtAddress.Text
        'rsN!Telephone = txtTelephone.Text
        rsN!Mobile = txtMobile.Text
        rsN!C_lose = "N"
        rsN.Update
        rsN.Close
    End If
    
    'cmdsale.Enabled = False
    
    MsgBox "Account Open Successfull! Your Account no.: " & txtID.Text, vbInformation, "Customer Info"
    
    Call clearTextboxes
    Exit Sub

End Sub

Private Sub cmdCash_Memo_Click()
If txtInvoice.Text = "" Then
Exit Sub
End If

Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Dim sort As Integer
Dim Today As Date
    sort = 0
    
    
'On Error Resume Next
'    Set rsU = New ADODB.Recordset
'        str = "select * from Invo_Prev where sl like '" & 1 & "'"
'        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
'
'    If Not rsU.EOF Then
'    On Error Resume Next
'        rsU.Close
'        str = "delete from Invo_Prev where sl like '" & 1 & "'"
'        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
'        rsU.Update
'        rsU.Close
'   End If
'
'rsProd.MoveFirst
'
'Do While Not rsProd.EOF
'On Error Resume Next
'
'Set rsN = New ADODB.Recordset
'        rsN.Open "Invo_Prev", conn, adOpenDynamic, adLockOptimistic, -1
'        rsN.AddNew
'            rsN!sl = 1
'            rsN!Sl_no = sort + 1
'            rsN!Invo_No = txtInvoice.Text
'            rsN!Ref_no = txtInvoice.Text & "-" & rsProd!sl
'            rsN!Ref_Code = txtID.Text
'            rsN!Ref_Name = txtName.Text
'            rsN!D_ate = txtDate.Text
'            rsN!Prod_Code = rsProd![Product Code]
'            rsN!Prod_Name = rsProd![Product Name] & ", " & rsProd![Description]
'            'rsN!Prod_Model = rsProd![Description]
'
'            rsN!Purchase = 0
'            rsN!Sale = Val(rsProd!Qty)
'            rsN!Lift = 0
'            rsN!Stock = rsU!Stock
'            rsN!Prod_Price = rsU!Prod_Price
'            rsN!Com = Val(rsProd![Commission])
'            rsN!Amount = (Val(rsProd![Total Price]) - Val(rsProd![Commission]))
'            rsN!Sale_Price = rsProd![Unit Price]
'            rsN!User_Id = User_Id
'            rsN!User_Name = User_Name
'            rsN.Update
'            rsN.Close
'        sort = sort + 1
'
'            rsProd.MoveNext
'Loop
        
Call Cash_Memo

End Sub

Private Sub cmdPrev_Click()
If txtInvoice.Text = "" Then
Exit Sub
End If

Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Dim sort As Integer
Dim Today As Date
    sort = 0
    
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Invo_Prev where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Invo_Prev where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If

rsProd.MoveFirst

Do While Not rsProd.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Invo_Prev", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
            rsN!sl = 1
            rsN!Sl_no = sort + 1
            rsN!Invo_No = txtInvoice.Text
            rsN!Ref_no = txtInvoice.Text & "-" & rsProd!sl
            rsN!Ref_Code = txtID.Text
            rsN!Ref_Name = txtName.Text
            rsN!D_ate = txtDate.Text
            rsN!Prod_Code = rsProd![Product Code]
            rsN!Prod_Name = rsProd![Product Name] & ", " & rsProd![Description]
            'rsN!Prod_Model = rsProd![Description]
            
            rsN!Purchase = 0
            rsN!Sale = Val(rsProd!Qty)
            rsN!Lift = 0
            rsN!Stock = rsU!Stock
            rsN!Prod_Price = rsU!Prod_Price
            rsN!Com = Val(rsProd![Commission])
            rsN!Amount = (Val(rsProd![Total Price]))
            rsN!Sale_Price = rsProd![Unit Price]
            rsN!User_Id = User_Id
            rsN!User_Name = User_Name
            rsN.Update
            rsN.Close
        sort = sort + 1
            
            rsProd.MoveNext
Loop
        
Call Preview
End Sub

Private Sub cmdPrint_Click()
If txtInvoice.Text = "" Then
Exit Sub
End If

Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Dim sort As Integer
Dim Today As Date
    sort = 0
    
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Invo_Prev where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly

    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Invo_Prev where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If

rsProd.MoveFirst

Do While Not rsProd.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Invo_Prev", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
            rsN!sl = 1
            rsN!Sl_no = sort + 1
            rsN!Invo_No = txtInvoice.Text
            rsN!Ref_no = txtInvoice.Text & "-" & rsProd!sl
            rsN!Ref_Code = txtID.Text
            rsN!Ref_Name = txtName.Text
            rsN!D_ate = txtDate.Text
            rsN!Prod_Code = rsProd![Product Code]
            rsN!Prod_Name = rsProd![Product Name] & ", " & rsProd![Description]
            'rsN!Prod_Model = rsProd![Description]

            rsN!Purchase = 0
            rsN!Sale = Val(rsProd!Qty)
            rsN!Lift = 0
            rsN!Stock = rsU!Stock
            rsN!Prod_Price = rsU!Prod_Price
            rsN!Com = Val(rsProd![Commission])
            rsN!Amount = (Val(rsProd![Total Price]) - Val(rsProd![Commission]))
            rsN!Sale_Price = rsProd![Unit Price]
            rsN!User_Id = User_Id
            rsN!User_Name = User_Name
            rsN.Update
            rsN.Close
        sort = sort + 1

            rsProd.MoveNext
Loop
        
Call Pos_Memo
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Really want to delete?", vbCritical + vbYesNo) = vbYes Then

    Dim Tran As String
    Dim Prod As String
        Tran = txtID.Text
    
        rsProd.MoveFirst
    
    Do While Not rsProd.EOF
        Prod = rsProd![Goods Code]
        
    
        Set rsU = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Code like '" & Prod & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
        If rsProd![Goods Code] = "101" Or rsProd![Goods Code] = "BKS01" Then
             rsU!Sale = rsU!Sale - (Val(rsProd![Total Price]) + Val(rsProd![Commission]))
             rsU!Stock = rsU!Stock + Val(rsProd![Total Price])
             rsU!Com = rsU!Com + Val(rsProd![Commission])
             rsU!Amount = rsU!Amount + (Val(rsProd![Total Price]) + Val(rsProd![Commission]))
             rsU.Update
             rsU.Close
        Else
            rsU!Sale = rsU!Sale - Val(rsProd!Qty)
             'rsU!Lift = rsU!Lift + Val(rsProd!Qty)
             rsU!Com = rsU!Com + Val(rsProd![Commission])
             rsU!Amount = rsU!Stock * rsU!Prod_Price
             rsU.Update
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
    
    If Prod = "101" Or Prod = "BKS01" Then
        Call CustomerLoad_Rev
    Else
        Call CustomerProd_Rev
    End If
    'rsU.Close
  rsProd.MoveNext
  
Loop
    
     If Val(txtCash.Text) > 0 Then
        Call Cash_Rev
     End If
    
    If Val(txtCash.Text) > 0 Then
        
    End If
    
    If Val(txtChq_Amnt.Text) > 0 Then
        Call Bank_Rev
    End If
    
'    If Val(txtCom_Amnt.Text) > 0 Then
'        Call Com_Rev
'    End If
    
'    If Val(txtOther_Amnt.Text) > 0 Then
'        Call Other_Rev
'    End If
    
    Call clearTextboxes
    Call grd_Prod
    Call Col_Prod
Else
Exit Sub
End If
End Sub

Private Sub cmdSales_Click()
On Error Resume Next
If txtInvoice.Text = "" Then
Exit Sub
End If

Call Invoice_Add


Dim Prod As String

rsProd.MoveFirst
    
    Do While Not rsProd.EOF
        Prod = rsProd![Product Code]
        
        Set rsU = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Code like '" & Prod & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rsU.EOF Then
            
            If cmbType.Text = "Cash Sale" Then
                    
               ' rsU!Sale = rsU!Sale + Val(rsProd!Qty)
               ' rsU!Stock = rsU!Stock - Val(rsProd!Qty)
                'rsU!Prod_Price = Val(rsProd![Unit Price])
                'rsU!Com = rsU!Com - Val(rsProd![Commission])
               ' rsU!Amount = rsU!Stock * rsU!Prod_Price
                rsU.Update
                Call Vendor_Cr
                Call Customer_Add
                Call Prod_Tran
                Call GL_Issue
                Call GL_Prof
'
                rsProd.MoveNext
                rsU.Close
            End If
            
            If cmbType.Text = "Credit Sale" Then
                    
                'rsU!Sale = rsU!Sale + Val(rsProd!Qty)
                'rsU!Stock = rsU!Stock - Val(rsProd!Qty)
                'rsU!Prod_Price = Val(rsProd![Unit Price])
                'rsU!Com = rsU!Com - Val(rsProd![Commission])
                'rsU!Amount = rsU!Stock * rsU!Prod_Price
                rsU.Update
                Call Vendor_Cr
                Call Customer_Add
                Call Prod_Tran
                Call GL_Issue
                Call GL_Prof
'                Call CustomerProd_Update
                rsProd.MoveNext
                rsU.Close
            End If
    
                
             If cmbType.Text = "Sales Return" Then
                 Call Vendor_Rtn
'                rsU!Return = rsU!Return + Val(rsProd!Qty)
'                rsU!Stock = rsU!Stock + Val(rsProd!Qty)
                'rsU!Prod_Price = Val(rsProd![Unit Price])
'                rsU!Com = rsU!Com - Val(rsProd![Commission])
'                rsU!Amount = rsU!Stock * rsU!Prod_Price
                rsU.Update
                Call Customer_Rtn
                Call Prod_Tran
                Call GL_Return
                Call Prof_Rtn
                'Call CustomerProd_Update
                rsProd.MoveNext
                rsU.Close
              End If
        
              If cmbType.Text = "" Then
                MsgBox "Invalid Sales Type!", vbCritical, "Sales Error"
                Exit Sub
              End If
              
    Else
       rsU.Close
       MsgBox "Invalied Product Code!", vbCritical, "Sales Error"
       Exit Sub
    End If
        
        Loop
            
    If Val(txtCharge.Text) > 0 Then
           Call Charge_Add
    End If
            
   If Val(txtCash.Text) > 0 Then
           Call Cash_Cr
   End If
       
   If Val(txtChq_Amnt.Text) > 0 Then
       Call Bank_Cr
   End If
'
    If Val(txtCard_Amnt.Text) > 0 Then
       Call Card_Cr
    End If
cmdPrev.Value = True
'        If Val(Total) < Val(netAmnt) Then
'         Call Credit_Sale
'        End If
        
    Call Prod_Name
    Call Prod_Model
    Call clearTextboxes
    Call grd_Prod
    Call Col_Prod
    Call Invo_No
    Total = 0
    lblPmt.Caption = "Total Tk. " + Format$(Val(Total), "###0.00")
    txtDate.Text = Today
    cmbType.SetFocus
    cmdPrev.Enabled = False
    cmdCash_Memo.Enabled = False
    cmdSales.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = True
    cst = 0
Exit Sub
   Resume Next
End Sub

Private Sub cmdSales_GotFocus()
If cmbType.Text = "Sale" Then
On Error Resume Next
Call Total_Cash
Dim ID As String
    ID = txtID.Text
    
Set rs = New ADODB.Recordset
str = "select * from Staff_Master where Staff_Id like '" & ID & "'"
rs.Open str, conn
If Not rs.EOF Then
    If Val(Total) > Val(netAmnt) Then
        txtTotal_Due.Text = Format$(Val(rs!Sales_Due + rs!Advance) - (Val(Total) - Val(netAmnt)), "###0.00")
        MsgBox "Todays Advance Amount: " & Val(Total) - Val(netAmnt) & " Total Dues: " & Val(rs!Sales_Due + rs!Advance) - (Val(Total) - Val(netAmnt)), vbCritical, "Sale Info!"
    Else
    If Val(Total) < Val(netAmnt) Then
        txtTotal_Due.Text = Format$(Val(rs!Sales_Due + rs!Advance) - (Val(Total) - Val(netAmnt)), "###0.00")
        MsgBox "Todays Sales Due Amount: " & Val(Total) - Val(netAmnt) & " Total Dues: " & Val(rs!Sales_Due + rs!Advance) - (Val(Total) - Val(netAmnt)), vbCritical, "Sale Info!"
    End If
    End If
rs.Close
Else
rs.Close
cmdSales.SetFocus
End If
End If
End Sub

Private Sub cmdUpdate_Click()
 On Error Resume Next
    
    Set rsN = New ADODB.Recordset
        str = "select * from Customer_Info where Customer like '" & txtID.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsN.EOF Then
           
        rsN!Customer = txtID.Text
        rsN!Type = cmbType.Text
        'rsN!Term = cmbTerm.Text
        rsN!Com = txtCommission.Text
        rsN!Balance = 0
        rsN!D_ate = txtDate.Text
        
        rsN!Name = txtName.Text
        rsN!Present_Address = txtAddress.Text
        'rsN!Telephone = txtTelephone.Text
        rsN!Mobile = txtMobile.Text
        rsN!C_lose = "N"
        rsN.Update
        rsN.Close
        
        MsgBox "Account Update Successfull! ", vbInformation, "Customer Info"
    
    Call clearTextboxes
    'cmdUpdate.Enabled = False
    cmdPrint.Enabled = True
    Else
        rsN.Close
        MsgBox "Invalid Customer ID", vbCritical, "Error"
    
        'cmdUpdate.Enabled = False
        cmdPrint.Enabled = True
    End If
    
    Exit Sub
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
Private Sub Preview()
If txtID.Text = "" Then
Exit Sub
End If
   
   If Val(netAmnt) >= 1 And Val(netAmnt) < 10 Then
        Value = Val(netAmnt)
        Call Case0_9
    End If
    If Val(netAmnt) >= 10 And Val(netAmnt) < 100 Then
        Value = Val(netAmnt)
        Call Case10_99
    End If
    If Val(netAmnt) >= 100 And Val(netAmnt) < 1000 Then
        Value = Val(netAmnt)
        Call Case100_999
    End If
    If Val(netAmnt) >= 1000 And Val(netAmnt) < 100000 Then
        Value = Val(netAmnt)
        Call Case1000_99999
    End If
    If Val(netAmnt) >= 100000 And Val(netAmnt) < 10000000 Then
        Value = Val(netAmnt)
        Call Case100000_9999999
    End If
    If Val(netAmnt) >= 10000000 Then
        Value = Val(netAmnt)
        Call Case10000000_999999999
    End If



Dim Tran As String
    Tran = txtInvoice.Text

    Set rs = New ADODB.Recordset
        str = "select * from Invo_Prev where Invo_No like '" & Tran & "' and Cdate(D_ate) like '" & txtDate.Text & "' and Sale > 0 Order by Prod_Code"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly

        If Not rs.EOF Then
            rptInvoice.rsInvoice.ConnectionString = cnStr
            rptInvoice.rsInvoice.Source = str
            
            rptInvoice.Field19.DataField = "Sale"
            rptInvoice.Field16.DataField = "Sale_Price"
            rptInvoice.lblTitle.Caption = "SALES INVOICE"
            
            rptInvoice.txtID.Text = txtID.Text
            rptInvoice.txtName.Text = txtName.Text
            rptInvoice.txtAddress.Text = txtAddress.Text
            rptInvoice.txtMobile.Text = txtMobile.Text
            
            rptInvoice.Balance.Text = Format$(Val(txtBalance.Text), "#,##0.00")
            rptInvoice.Sale_Amnt.Text = Format$(Val(txtSales_Due.Text), "#,##0.00")
            
            If Val(txtCash.Text) > 0 Then
                rptInvoice.Pmt_Mode.Text = "Cash Payment"
                rptInvoice.Pmt_Des.Text = Format$(Val(txtCash.Text), "#,##0.00")
                rptInvoice.Paid_Amnt.Text = Format$(Val(txtCash.Text), "#,##0.00")
                rptInvoice.Outstanding.Text = Format$((Val(txtBalance.Text) + Val(netAmnt)) - Val(txtCash.Text), "#,##0.00")
            
            Else
            
                If Val(txtChq_Amnt.Text) > 0 Then
                rptInvoice.Pmt_Mode.Text = "Bank Payment"
                rptInvoice.Pmt_Des.Text = Format$(Val(txtChq_Amnt.Text), "#,##0.00")
                rptInvoice.Paid_Amnt.Text = Format$(Val(txtChq_Amnt.Text), "#,##0.00")
                rptInvoice.Outstanding.Text = Format$((Val(txtBalance.Text) + Val(netAmnt)) - Val(txtChq_Amnt.Text), "#,##0.00")
                
                Else
            
                    If Val(txtCard_Amnt.Text) > 0 Then
                    rptInvoice.Pmt_Mode.Text = "Card Payment"
                    rptInvoice.Pmt_Des.Text = Format$(Val(txtCard_Amnt.Text), "#,##0.00")
                    rptInvoice.Paid_Amnt.Text = Format$(Val(txtCard_Amnt.Text), "#,##0.00")
                    rptInvoice.Outstanding.Text = Format$((Val(txtBalance.Text) + Val(netAmnt)) - Val(txtCard_Amnt.Text), "#,##0.00")
                
                    Else
            
                    rptInvoice.Pmt_Mode.Text = "NIL"
                    rptInvoice.Pmt_Des.Text = Format$(Val(0), "#,##0.00")
                    rptInvoice.Paid_Amnt.Text = Format$(Val(0), "#,##0.00")
                    rptInvoice.Outstanding.Text = Format$((Val(txtBalance.Text) + Val(txtSales_Due.Text)) - Val(0), "#,##0.00")
                    End If
                End If
            End If
'
'            rptInvoice.txtTotal.Text = Format(Val(netAmnt), "#,##0.00")
'            rptInvoice.txtNet.Text = Format(Val(netAmnt), "#,##0.00")
            rptInvoice.txtInword.Text = inword
            rptInvoice.txtTerms.Text = txtTerms.Text
            rptInvoice.txtCharge.Text = Format$(Val(txtCharge.Text), "#,##0.00")
            rptInvoice.txtNet.Text = Format$(Val(netAmnt), "#,##0.00")
           
            
            rs.Close
        Else
            MsgBox "There is no such Invoice found, ", vbCritical + vbOKOnly
            rs.Close
        End If
    rptInvoice.Show 1
   End Sub
Private Sub Pos_Memo()
If txtID.Text = "" Then
Exit Sub
End If
   
   If Val(netAmnt) >= 1 And Val(netAmnt) < 10 Then
        Value = Val(netAmnt)
        Call Case0_9
    End If
    If Val(netAmnt) >= 10 And Val(netAmnt) < 100 Then
        Value = Val(netAmnt)
        Call Case10_99
    End If
    If Val(netAmnt) >= 100 And Val(netAmnt) < 1000 Then
        Value = Val(netAmnt)
        Call Case100_999
    End If
    If Val(netAmnt) >= 1000 And Val(netAmnt) < 100000 Then
        Value = Val(netAmnt)
        Call Case1000_99999
    End If
    If Val(netAmnt) >= 100000 And Val(netAmnt) < 10000000 Then
        Value = Val(netAmnt)
        Call Case100000_9999999
    End If
    If Val(netAmnt) >= 10000000 Then
        Value = Val(netAmnt)
        Call Case10000000_999999999
    End If



Dim Tran As String
    Tran = txtInvoice.Text

    Set rs = New ADODB.Recordset
        str = "select * from Invo_Prev where Invo_No like '" & Tran & "' and Cdate(D_ate) like '" & Today & "' and Sale >0 Order by Prod_code"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly

        If Not rs.EOF Then
            rptCash_Memo.rsInvoice.ConnectionString = cnStr
            rptCash_Memo.rsInvoice.Source = str
            
            rptCash_Memo.Field19.DataField = "Sale"
'            rptCash_Memo.Field30.DataField = "Sale"
            rptCash_Memo.Field16.DataField = "Sale_Price"
            rptCash_Memo.txtDiscount.Text = txtCharge.Text
            rptCash_Memo.lblTitle.Caption = "CASH MEMO"
            rptCash_Memo.txtNet.Text = Format$((Val(netAmnt)), "###0.00")
            
            'rptCash_Memo.txtID.Text = txtID.Text
            'rptInvoice.txtName.Text = txtName.Text
            'rptInvoice.txtAddress.Text = txtAddress.Text
            'rptInvoice.txtMobile.Text = txtMobile.Text
            
            'rptInvoice.Balance.Text = Format(Val(txtBalance.Text), "#,##0.00")
            'rptInvoice.Sale_Amnt.Text = Format(Val(netAmnt), "#,##0.00")
            
            If Val(txtCash.Text) > 0 Then
                rptCash_Memo.Pmt_Mode.Text = "Cash Payment"
                rptCash_Memo.Pmt_Des.Text = Format(Val(txtCash.Text), "#,##0.00")
                rptCash_Memo.Paid_Amnt.Text = Format(Val(txtCash.Text), "#,##0.00")
                rptCash_Memo.Outstanding.Text = Format((Val(txtBalance.Text) + Val(netAmnt)) - Val(txtCash.Text), "#,##0.00")
            
            Else
            
                If Val(txtChq_Amnt.Text) > 0 Then
                rptCash_Memo.Pmt_Mode.Text = "Bank Payment"
                rptCash_Memo.Pmt_Des.Text = txtChq_No.Text
                rptCash_Memo.Paid_Amnt.Text = Format(Val(txtChq_Amnt.Text), "#,##0.00")
                rptCash_Memo.Outstanding.Text = Format((Val(txtBalance.Text) + Val(netAmnt)) - Val(txtChq_Amnt.Text), "#,##0.00")
                
                Else
            
                    If Val(txtCard_Amnt.Text) > 0 Then
                    rptCash_Memo.Pmt_Mode.Text = "Card Payment"
                    rptCash_Memo.Pmt_Des.Text = txtCard_No.Text
                    rptCash_Memo.Paid_Amnt.Text = Format(Val(txtCard_Amnt.Text), "#,##0.00")
                    rptCash_Memo.Outstanding.Text = Format((Val(txtBalance.Text) + Val(netAmnt)) - Val(txtCard_Amnt.Text), "#,##0.00")
                
                    Else
            
                    rptCash_Memo.Pmt_Mode.Text = "NIL"
                    rptCash_Memo.Pmt_Des.Text = Format(Val(0), "#,##0.00")
                    rptCash_Memo.Paid_Amnt.Text = Format(Val(0), "#,##0.00")
                    rptCash_Memo.Outstanding.Text = Format((Val(txtBalance.Text) + Val(txtSales_Due.Text)) - Val(0), "#,##0.00")
                    End If
                End If
            End If
'
'            rptInvoice.txtTotal.Text = Format(Val(netAmnt), "#,##0.00")
'            rptInvoice.txtNet.Text = Format(Val(netAmnt), "#,##0.00")
            
            
            rptCash_Memo.txtInword.Text = inword
            
            'rptCash_Memo.txtTerms.Text = txtTerms.Text
            rs.Close
        Else
            MsgBox "There is no such Invoice found, ", vbCritical + vbOKOnly
            rs.Close
        End If
    If MsgBox("Do you want to print memo now", vbYesNo, "Print") = vbYes Then
    rptCash_Memo.PrintReport True
    Unload rptCash_Memo
    Else
    rptCash_Memo.Show 1
    End If
   End Sub

Private Sub Cash_Memo()
If cmbAgent_Code.Text = "" Then
Exit Sub
End If
   
   If Val(netAmnt) >= 1 And Val(netAmnt) < 10 Then
        Value = Val(netAmnt)
        Call Case0_9
    End If
    If Val(netAmnt) >= 10 And Val(netAmnt) < 100 Then
        Value = Val(netAmnt)
        Call Case10_99
    End If
    If Val(netAmnt) >= 100 And Val(netAmnt) < 1000 Then
        Value = Val(netAmnt)
        Call Case100_999
    End If
    If Val(netAmnt) >= 1000 And Val(netAmnt) < 100000 Then
        Value = Val(netAmnt)
        Call Case1000_99999
    End If
    If Val(netAmnt) >= 100000 And Val(netAmnt) < 10000000 Then
        Value = Val(netAmnt)
        Call Case100000_9999999
    End If
    If Val(netAmnt) >= 10000000 Then
        Value = Val(netAmnt)
        Call Case10000000_999999999
    End If



Dim Tran As String
    Tran = cmbAgent_Code.Text

    Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where Ref_No like '" & Tran & "' and Sale > 0 Order by Prod_Code"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly

        If Not rs.EOF Then
            rptInvoice.rsInvoice.ConnectionString = cnStr
            rptInvoice.rsInvoice.Source = str
            
            rptInvoice.Field19.DataField = "Sale"
            rptInvoice.Field16.DataField = "Sale_Price"
            rptInvoice.Field30.Visible = True
            rptInvoice.lblTitle.Caption = "AGENT INVOICE"
            rptInvoice.lblTitle.Font.Size = 14
            rptInvoice.txtID.Text = cmbAgent_Code.Text
            rptInvoice.txtName.Text = cmbAgent_Name.Text
            rptInvoice.txtAddress.Text = ""
            rptInvoice.txtMobile.Text = ""
            
            rptInvoice.Balance.Text = Format$(Val(0), "#,##0.00")
            rptInvoice.Sale_Amnt.Text = Format$(Val(0), "#,##0.00")
            
            If Val(txtCash.Text) > 0 Then
                rptInvoice.Pmt_Mode.Text = "Cash Payment"
                rptInvoice.Pmt_Des.Text = Format$(Val(0), "#,##0.00")
                rptInvoice.Paid_Amnt.Text = Format$(Val(0), "#,##0.00")
                rptInvoice.Outstanding.Text = Format$(Val(0), "#,##0.00")
            
'            Else
'
'                If Val(txtChq_Amnt.Text) > 0 Then
'                rptInvoice.Pmt_Mode.Text = "Bank Payment"
'                rptInvoice.Pmt_Des.Text = Format$(Val(txtChq_Amnt.Text), "#,##0.00")
'                rptInvoice.Paid_Amnt.Text = Format$(Val(txtChq_Amnt.Text), "#,##0.00")
'                rptInvoice.Outstanding.Text = Format$((Val(txtBalance.Text) + Val(netAmnt)) - Val(txtChq_Amnt.Text), "#,##0.00")
'
'                Else
'
'                    If Val(txtCard_Amnt.Text) > 0 Then
'                    rptInvoice.Pmt_Mode.Text = "Card Payment"
'                    rptInvoice.Pmt_Des.Text = Format$(Val(txtCard_Amnt.Text), "#,##0.00")
'                    rptInvoice.Paid_Amnt.Text = Format$(Val(txtCard_Amnt.Text), "#,##0.00")
'                    rptInvoice.Outstanding.Text = Format$((Val(txtBalance.Text) + Val(netAmnt)) - Val(txtCard_Amnt.Text), "#,##0.00")
'
                    Else
'
                    rptInvoice.Pmt_Mode.Text = "NIL"
                    rptInvoice.Pmt_Des.Text = Format$(Val(0), "#,##0.00")
                    rptInvoice.Paid_Amnt.Text = Format$(Val(0), "#,##0.00")
                    rptInvoice.Outstanding.Text = Format$(Val(0), "#,##0.00")
'                    End If
'                End If
            End If
'
'            rptInvoice.txtTotal.Text = Format(Val(netAmnt), "#,##0.00")
'            rptInvoice.txtNet.Text = Format(Val(netAmnt), "#,##0.00")
            rptInvoice.txtInword.Text = inword
'            rptInvoice.txtTerms.Text = txtTerms.Text
            rptInvoice.txtCharge.Text = Format$(Val(0), "#,##0.00")
            rptInvoice.txtNet.Text = Format$(Val(netAmnt), "#,##0.00")
           
            
            rs.Close
        Else
            MsgBox "There is no such Invoice found, ", vbCritical + vbOKOnly
            rs.Close
        End If
    rptInvoice.Show 1
   End Sub

Private Sub Form_Activate()
cmbType.Text = "Cash Sale"
'txtID.Text = "0001"
'txtName.Text = "General"
'txtType.Text = "General"
'txtAddress = "Khulna"
cmbType.SetFocus
End Sub

Private Sub Form_Load()
    cmdPrev.Enabled = False
    cmdSales.Enabled = False
    cmdPrint.Enabled = True
    cmdDelete.Enabled = False
    cmdCash_Memo.Enabled = True
    Call Cust_Id
    Call Cust_Name
    Call Agent_Id
    Call Agent_Name
    Call Prod_Code
    Call Prod_Name
    Call Prod_Model
    Call Prod_Type
    
    Call grd_Prod
    Call Col_Prod
    'Call Customer_Name
    Call Account_Name
    Call Bank_Name
    Call Br_Name
    
Call clearTextboxes
txtDate.Text = Today
Call Invo_No
lblCost.ForeColor = vbWhite
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblCost.ForeColor = vbWhite
End Sub

Private Sub grdProd_AfterColEdit(ByVal ColIndex As Integer)
rsProd![Total Price] = rsProd!Qty * rsProd![Unit Price]
rsProd.MoveFirst
        Prod_Sl = 0
        netAmnt = 0
        cst = 0
    Do While Not rsProd.EOF
        Prod_Sl = Prod_Sl + 1
        rsProd!sl = Prod_Sl
        netAmnt = netAmnt + (Val(rsProd![Total Price]) - rsProd![Commission])
        cst = Round(cst + (Val(rsProd![Cost]) * Val(rsProd![Qty])))
        rsProd.MoveNext
    Loop
        
        strloan = Format$(Val(netAmnt), "###0.00")
        txtSales_Due.Text = Format$(Round(Val(netAmnt)), "###0.00")
        txtTotal_Due.Text = Format$(Round(Val(txtBalance.Text) + Val(netAmnt)), "###0.00")
        lblNet.Caption = "Total Amount Receivable Tk. " + strloan
        lblCost.Caption = "Total Cost Tk. " & cst
End Sub

Private Sub grdProd_DblClick()

If MsgBox("Do you want to delete record?", vbCritical + vbYesNo, "Delete") = vbYes Then

With rsProd
        .Delete (adAffectCurrent)
End With
    rsProd.MoveFirst
        Prod_Sl = 0
        netAmnt = 0
        cst = 0
    Do While Not rsProd.EOF
        Prod_Sl = Prod_Sl + 1
        rsProd!sl = Prod_Sl
        netAmnt = netAmnt + (Val(rsProd![Total Price]) - rsProd![Commission])
        cst = Round(cst + (Val(rsProd![Cost]) * Val(rsProd![Qty])))
        rsProd.MoveNext
    Loop
        
        strloan = Format$(Val(netAmnt), "###0.00")
        txtSales_Due.Text = Format$(Round(Val(netAmnt)), "###0.00")
        txtTotal_Due.Text = Format$(Round(Val(txtBalance.Text) + Val(netAmnt)), "###0.00")
        lblNet.Caption = "Total Amount Receivable Tk. " + strloan
        lblCost.Caption = "Total Cost Tk. " & cst
        
Else
Exit Sub
End If
End Sub

Private Sub Label7_Click()
frmProd_Search.Show 1
End Sub

Private Sub lblCost_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblCost.ForeColor = vbBlack
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCash.SelStart = 0
    txtCash.SelLength = Len(txtCash.Text)
    txtCash.SetFocus
End If
End Sub

Private Sub txt1_LostFocus()
Call Cash_Amnt
End Sub

Private Sub txt10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt5.SelStart = 0
    txt5.SelLength = Len(txt5.Text)
    txt5.SetFocus
End If
End Sub

Private Sub txt10_LostFocus()
Call Cash_Amnt
End Sub

Private Sub txt100_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt50.SelStart = 0
    txt50.SelLength = Len(txt50.Text)
    txt50.SetFocus
End If
End Sub

Private Sub txt100_LostFocus()
Call Cash_Amnt
End Sub

Private Sub txt1000_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt500.SelStart = 0
    txt500.SelLength = Len(txt500.Text)
    txt500.SetFocus
End If
End Sub

Private Sub txt1000_LostFocus()
Call Cash_Amnt
End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt1.SelStart = 0
    txt1.SelLength = Len(txt1.Text)
    txt1.SetFocus
End If
End Sub

Private Sub txt2_LostFocus()
Call Cash_Amnt
End Sub

Private Sub txt20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt10.SelStart = 0
    txt10.SelLength = Len(txt10.Text)
    txt10.SetFocus
End If
End Sub

Private Sub txt20_LostFocus()
Call Cash_Amnt
End Sub

Private Sub txt5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt2.SelStart = 0
    txt2.SelLength = Len(txt2.Text)
    txt2.SetFocus
End If
End Sub

Private Sub txt5_LostFocus()
Call Cash_Amnt
End Sub

Private Sub txt50_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt20.SelStart = 0
    txt20.SelLength = Len(txt20.Text)
    txt20.SetFocus
End If
End Sub

Private Sub txt50_LostFocus()
Call Cash_Amnt
End Sub

Private Sub txt500_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt100.SelStart = 0
    txt100.SelLength = Len(txt100.Text)
    txt100.SetFocus
End If
End Sub

Private Sub txt500_LostFocus()
Call Cash_Amnt
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbBank.SetFocus
End If
End Sub

Private Sub txtAccount_LostFocus()
    If txtAccount.Text = "" Then
    cmbCard_Name.SetFocus
    Exit Sub
    End If
    
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

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMobile.SetFocus
End If
End Sub

Private Sub txtCard_Amnt_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
cmdSales.SetFocus
End If
End Sub

Private Sub txtCard_Amnt_LostFocus()
On Error Resume Next
Call Total_Cash
End Sub

Private Sub txtCard_No_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCard_Amnt.SelStart = 0
    txtCard_Amnt.SelLength = Len(txtCard_Amnt.Text)
    txtCard_Amnt.SetFocus
End If
End Sub

Private Sub txtCash_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
    On Error Resume Next
    txtAccount.SelStart = 0
    txtAccount.SelLength = Len(txtAccount.Text)
    txtAccount.SetFocus
    End If
End Sub

Private Sub txtCash_LostFocus()
On Error Resume Next
Call Total_Cash
'cmdSales.SetFocus
'If Val(Total) <> Val(netAmnt) Then
'MsgBox "Collection differs from Sale Amount" & Val(Total) - Val(netAmnt), vbCritical, "Sale Info!"
'Else
'txtAccount.SetFocus
'End If
End Sub

Private Sub txtCharge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt1000.SelStart = 0
    txt1000.SelLength = Len(txt1000.Text)
    txt1000.SetFocus
End If
End Sub

Private Sub txtCharge_LostFocus()
    netAmnt = Round(netAmnt - (Val(txtCharge.Text)))
    strloan = 0
    strloan = Format$(Round(Val(netAmnt)), "###0.00")
    txtSales_Due.Text = Format$(Round(Val(netAmnt)), "###0.00")
    txtTotal_Due.Text = Format$(Round(Val(txtBalance.Text) - Val(netAmnt)), "###0.00")
    lblNet.Caption = "Total Amount Receivable Tk. " + strloan
    txtCharge.Text = Format$((Val(txtCharge.Text)), "###0.00")
End Sub

Private Sub txtChq_Amnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
cmbCard_Name.SelStart = 0
cmbCard_Name.SelLength = Len(cmbCard_Name.Text)
cmbCard_Name.SetFocus
End If
End Sub

Private Sub txtChq_Amnt_LostFocus()
On Error Resume Next
Call Total_Cash
'If Val(Total) <> Val(netAmnt) Then
'MsgBox "Collection differs from Sale Amount" & Val(Total) - Val(netAmnt), vbCritical, "Sale Info!"
'Else
'txtCom_Des.SetFocus
'End If
End Sub

Private Sub txtChq_No_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtChq_Amnt.SelStart = 0
txtChq_Amnt.SelLength = Len(txtChq_Amnt.Text)
txtChq_Amnt.SetFocus
End If
End Sub

Private Sub txtCom_Amnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
'txtBcash_Des.SelStart = 0
'txtBcash_Des.SelLength = Len(txtBcash_Des.Text)
'txtBcash_Des.SetFocus
End If
End Sub
'Private Sub txtCom_Amnt_LostFocus()
'On Error Resume Next
'Call Total_Cash
'If Val(Total) <> Val(netAmnt) Then
'MsgBox "Collection differs from Sale Amount" & Val(Total) - Val(netAmnt), vbCritical, "Sale Info!"
'Else
'txtBcash_Des.SetFocus
'End If

Private Sub txtCommission_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Sl.SetFocus
End If
End Sub

Private Sub txtCommission_LostFocus()
    txtProd_Price.Text = Format$(Val(txtProd_Price.Text), "###0.00")
        rsProd.AddNew
        rsProd!sl = Prod_Sl
        rsProd![Product Code] = txtProd_Sl.Text
        rsProd![Product Name] = cmbProd_Name.Text
        rsProd![Description] = cmbProd_Model.Text
        rsProd!Qty = txtQty.Text
        rsProd![Unit Price] = txtProd_Cost.Text
        rsProd![Total Price] = txtProd_Price.Text
        rsProd![Commission] = txtCommission.Text
        rsProd![Cost] = txtCost.Text
        
        rsProd.MoveFirst
        Prod_Sl = 0
        netAmnt = 0
        cst = 0
    Do While Not rsProd.EOF
        Prod_Sl = Prod_Sl + 1
        rsProd!sl = Prod_Sl
        netAmnt = Round(netAmnt + (Val(rsProd![Total Price]) - Val(rsProd![Commission])))
        cst = Round(cst + (Val(rsProd![Cost]) * Val(rsProd![Qty])))
        rsProd.MoveNext
    Loop
        strloan = 0
        strloan = Format$(Round(Val(netAmnt)), "###0.00")
        txtSales_Due.Text = Format$(Round(Val(netAmnt)), "###0.00")
        txtTotal_Due.Text = Format$(Round(Val(txtBalance.Text) + Val(netAmnt)), "###0.00")
        lblNet.Caption = "Total Amount Receivable Tk. " + strloan
        lblCost.Caption = "Total Cost Tk. " & cst
    Set grdProd.DataSource = rsProd
        grdProd.Refresh
   
    Call Col_Prod
        txtProd_Sl.Text = ""
        cmbProd_Name.Text = ""
        cmbProd_Model.Text = ""
        txtQty.Text = "0"
        txtProd_Cost.Text = "0.00"
        txtProd_Price.Text = "0.00"
        'txtPercent.Text = "0.00"
        txtCommission.Text = "0.00"

cmdPrev.Enabled = True
cmdSales.Enabled = True
cmdCash_Memo.Enabled = True
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmbType.Text = "Sale" Then


Dim Tran As String
    Tran = txtID.Text

    Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where Invo_No like '" & Tran & "' And Cdate(D_ate) like '" & Today & "' And Sale >0"
        rs.Open str, conn

 If Not rs.EOF Then
    MsgBox "Duplicate Sales Entry.", vbCritical, "Erroe!"
            rs.MoveFirst
                txtID.Text = rs!Ref_Code
                txtName.Text = rs!Ref_Name
                Date = rs!D_ate
            
                Prod_Sl = 0
                netAmnt = 0

            Do While Not rs.EOF
    
            If rs!Prod_Code = "101" Or rs!Prod_Code = "BKS01" Then
                
                Prod_Sl = Prod_Sl + 1
            
                rsProd.AddNew
                rsProd!sl = Prod_Sl
                rsProd![Goods Code] = rs!Prod_Code
                rsProd![Goods Name] = rs!Prod_Name
                rsProd![Description] = rs!Prod_Model
                rsProd!Qty = 0
                rsProd![Unit Price] = Format$(Val(rs!Sale_Price), "###0.0000")
                rsProd![Total Price] = Format$(Val(rs!Sale_Price) - Val(rs!Com), "###0.0000")
                rsProd![Commission] = Format$(Val(rs!Com), "###0.0000")
                netAmnt = Round(netAmnt + (Val(rs!Sale_Price) + Val(rs!Com)))
                
                Set grdProd.DataSource = rsProd
                grdProd.Refresh
                Call Col_Prod
                rs.MoveNext
            
            Else
            
            Prod_Sl = Prod_Sl + 1
            
                rsProd.AddNew
                rsProd!sl = Prod_Sl
                rsProd![Goods Code] = rs!Prod_Code
                rsProd![Goods Name] = rs!Prod_Name
                rsProd![Description] = rs!Prod_Model
                rsProd!Qty = rs!Sale
                rsProd![Unit Price] = Format$(Val(rs!Sale_Price), "###0.0000")
                rsProd![Total Price] = Format$(Val(rs!Sale * rs!Sale_Price), "###0.0000")
                rsProd![Commission] = Format$(Val(rs!Com), "###0.0000")
                netAmnt = Round(netAmnt + Val(rs!Sale * rs!Sale_Price))
                
                Set grdProd.DataSource = rsProd
                grdProd.Refresh
                Call Col_Prod
                rs.MoveNext
            End If
            
            Loop
    
    rs.Close
    
        rsProd.MoveFirst
        Prod_Sl = 0
        netAmnt = 0
    Do While Not rsProd.EOF
        Prod_Sl = Prod_Sl + 1
        rsProd!sl = Prod_Sl
        netAmnt = Round(netAmnt + Val(rsProd![Total Price]))
        rsProd.MoveNext
    Loop
        
        strloan = Format$(Val(netAmnt), "###0.00")
        lblNet.Caption = "Total Amount Receivable Tk. " + strloan
    
    Set grdProd.DataSource = rsProd
        grdProd.Refresh
   
    Call Col_Prod
    
    
        Set rs = New ADODB.Recordset
               str = "select * from Cash_Book where MR_No like '" & Tran & "' And Cdate(Date) like '" & Today & "'"
               rs.Open str, conn
           
           If Not rs.EOF Then
                rs.MoveFirst
           
           Do While Not rs.EOF
           
               txtCash.Text = Format$(Val(txtCash.Text) + Val(rs!Dr), "###0.00")
               txt1000.Text = Val(txt1000.Text) + Val(rs!Tk1000)
               txt500.Text = Val(txt500.Text) + Val(rs!Tk500)
               txt100.Text = Val(txt100.Text) + Val(rs!Tk100)
               txt50.Text = Val(txt50.Text) + Val(rs!Tk50)
               txt20.Text = Val(txt20.Text) + Val(rs!Tk20)
               txt10.Text = Val(txt10.Text) + Val(rs!Tk10)
               txt5.Text = Val(txt5.Text) + Val(rs!Tk5)
               txt2.Text = Val(txt2.Text) + Val(rs!Tk2)
               txt1.Text = Val(txt1.Text) + Val(rs!Tk1)
            rs.MoveNext
            Loop
               Call Total_Cash
               rs.Close
           Else
               rs.Close
               txtCash.Text = 0
               txt1000.Text = 0
               txt500.Text = 0
               txt100.Text = 0
               txt50.Text = 0
               txt20.Text = 0
               txt10.Text = 0
               txt5.Text = 0
               txt2.Text = 0
               txt1.Text = 0
           End If
    
            Set rs = New ADODB.Recordset
               str = "select * from Bank_Tran where MR_No like '" & Tran & "' And Cdate(Date) like '" & Today & "'"
               rs.Open str, conn
           
           If Not rs.EOF Then
               txtChq_Amnt.Text = Format$(Val(rs!Dr), "###0.00")
               txtAccount.Text = rs!AC_No
               cmbBank.Text = rs!Bank_Name
               cmbBranch.Text = rs!Branch_Name
               txtChq_No.Text = rs!Chq_No
               rs.Close
           Else
               rs.Close
               txtChq_Amnt.Text = 0
               txtAccount.Text = ""
               cmbBank.Text = ""
               cmbBranch.Text = ""
               txtChq_No.Text = ""
           End If
            
    
        cmdPrint.Enabled = True
        cmdSales.Enabled = False
        cmdDelete.Enabled = True
    Else
    rs.Close
    'Call clearTextboxes
    'txtInvoice.Text = Tran
    txtDate.Text = Today
    Exit Sub
  End If
End If
End If
End Sub

Private Sub txtDown_Payment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'txtDue.SelStart = 0
'txtDue.SelLength = Len(txtDue.Text)
'txtDue.SetFocus
End If
End Sub

Private Sub txtDown_Payment_LostFocus()
'txtDue.Text = Val(txtLoan_Amnt.Text) - Val(txtDown_Payment.Text)
'txtDown_Payment.Text = Format$(Val(txtDown_Payment.Text), "###0.00")
'txtDue.Text = Format$(Val(txtDue.Text), "###0.00")
End Sub

Private Sub txtDue_Date_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'cmbTerm.SelStart = 0
'cmbTerm.SelLength = Len(cmbTerm.Text)
'cmbTerm.SetFocus
End If
End Sub

Private Sub txtDue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'txtInst_No.SelStart = 0
'txtInst_No.SelLength = Len(txtInst_No.Text)
'txtInst_No.SetFocus
End If
End Sub

Private Sub txtDue_LostFocus()
'txtDue.Text = Format$(Val(txtDue.Text), "###0.00")
End Sub

Private Sub txtEngine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty.Text)
    txtQty.SetFocus
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtName.SetFocus
End If
End Sub

Private Sub txtId_LostFocus()
On Error Resume Next
'If txtId.Text = "" Then
'txtName.SetFocus
'Exit Sub
'End If

Dim ID As String
Dim mid As Integer
    ID = txtID.Text
    mid = 0
    
    Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & ID & "'"
        rs.Open str, conn

    If Not rs.EOF Then
        'Call clearTextboxes
        txtID.Text = rs!Customer_Code
        'txtDate.Text = Today
        txtType.Text = rs!Customer_Type
        txtName.Text = rs!Customer_Name
        txtAddress.Text = rs!Customer_Address
        txtMobile.Text = rs!Mobile
        cmbAgent_Code.Text = rs!Agent_Code
        cmbAgent_Name.Text = rs!Agent_Name
        txtBalance.Text = Format$(Val(rs!Balance), "###0.00")
        cmbAgent_Code.SetFocus
     rs.Close
     cmdSales.Enabled = True
        
    Else
        
    If MsgBox("Do you want add new Customer?", vbInformation + vbYesNo, "Add New") = vbYes Then
        
        Set rsU = New ADODB.Recordset
            str = "select * from Customer_Master order by Customer_Code"
            rsU.Open str, conn
        
            If Not rsU.EOF Then
                rsU.MoveFirst
        
            Do While Not rsU.EOF = True
                mid = rsU!Customer_Code
                rsU.MoveNext
            Loop
                rsU.Close
                mid = mid + 1
            Else
                rsU.Close
                mid = 1
            End If
            
            txtID.Text = Format(mid, "000#")
            txtDate.Text = Today
            txtType.Text = ""
            txtName.Text = ""
            txtAddress.Text = ""
            txtMobile.Text = ""
            txtName.SetFocus
        rs.Close
     Else
        rs.Close
        Exit Sub
    End If
    txtName.SetFocus
    'cmdSales.Enabled = True
End If
End Sub

Private Sub txtInst_Amnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'txtDue_Date.SelStart = 0
'txtDue_Date.SelLength = Len(txtDue_Date.Text)
'txtDue_Date.SetFocus
End If
End Sub

Private Sub txtInst_Amnt_LostFocus()
'txtInst_Amnt.Text = Format$(Val(txtInst_Amnt.Text), "###0.00")
End Sub

Private Sub txtInst_No_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'txtInst_Amnt.SelStart = 0
'txtInst_Amnt.SelLength = Len(txtInst_Amnt.Text)
'txtInst_Amnt.SetFocus
End If
End Sub

Private Sub txtInst_No_LostFocus()
'txtInst_Amnt.Text = Val(txtDue.Text) / Val(txtInst_No.Text)
'txtInst_No.Text = Format$(Val(txtInst_No.Text), "###0")
'txtInst_Amnt.Text = Format$(Val(txtInst_Amnt.Text), "###0.00")
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbType.SetFocus
End If
End Sub

Private Sub txtInvoice_LostFocus()
On Error Resume Next
If txtInvoice.Text = "" Then
Exit Sub
End If

Dim Tran As String
    Tran = txtInvoice.Text

Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where Invo_No like '" & Tran & "'"
        rs.Open str, conn

If Not rs.EOF Then
    
            rs.MoveFirst
                
                Prod_Sl = 0
                netAmnt = 0
                cst = 0

            Do While Not rs.EOF
    
            Prod_Sl = Prod_Sl + 1
            
                rsProd.AddNew
                rsProd!sl = Prod_Sl
                rsProd![Product Code] = rs!Prod_Code
                rsProd![Product Name] = rs!Prod_Name
                rsProd![Description] = rs!Prod_Model
                rsProd!Qty = rs!Sale
                rsProd![Unit Price] = Format$(Val(rs!Sale_Price), "###0.0000")
                rsProd![Total Price] = Format$(Val(rs!Sale * rs!Sale_Price), "###0.0000")
                rsProd![Commission] = Format$(Val(rs!Com), "###0.0000")
                rsProd![Cost] = Format$(Val(rs!Prod_Price), "###0.0000")
                netAmnt = Round(netAmnt + (Val(rs!Sale * rs!Sale_Price) - Val(rs!Com)))
                cst = Round(cst + (Val(rsProd![Cost]) * Val(rsProd![Qty])))
                Set grdProd.DataSource = rsProd
                grdProd.Refresh
                cmbAgent_Code.Text = rs!Ref_no
                Call Col_Prod
                rs.MoveNext
            Loop
        rs.Close
        
        strloan = 0
        strloan = Format$(Round(Val(netAmnt)), "###0.00")
        lblNet.Caption = "Total Amount Receivable Tk. " + strloan
        lblCost.Caption = "Total Cost Tk. " & cst
        
        
    End If
    
Set rs = New ADODB.Recordset
        str = "select * from Sales_Invoice where Invo_No like '" & Tran & "'"
        rs.Open str, conn

If Not rs.EOF Then
        MsgBox "Duplicate invoice No.", vbCritical, "Erroe!"
            txtInvoice.Text = rs!Invo_No
            cmbType.Text = rs!Sale_Type
            txtDate.Text = rs!D_ate
            
            txtID.Text = rs!Customer
            txtName.Text = rs!Name
            txtType.Text = rs!Type
            txtAddress.Text = rs!Address
            txtMobile.Text = rs!Mobile
            strloan = 0
            txtBalance.Text = Format$(Round(Val(rs!Balance)), "###0.00")
            txtSales_Due.Text = Format$(Round(Val(rs!Amount)), "###0.00")
            txtTotal_Due.Text = Format$(Round(Val(rs!Net_Due)), "###0.00")
            
            txtCash.Text = Format$(Val(rs!Cash), "###0.00")
            
            txtAccount.Text = rs!Account
            cmbBank.Text = rs!Bank_Name
            cmbBranch.Text = rs!Branch_Name
            txtChq_No.Text = rs!Chq_No
            txtChq_Amnt.Text = Format$(Val(rs!Chq_Amnt), "###0.00")
            
            cmbCard_Name.Text = rs!Card_Name
            txtCard_No.Text = rs!Card_No
            txtCard_Amnt.Text = Format$(Val(rs!Card_Amnt), "###0.00")
            txtCharge.Text = Format$(Val(rs!Charge_Amnt), "###0.00")
            txtTerms.Text = rs!Terms
            
            netAmnt = Round(netAmnt - (Val(txtCharge.Text)))
            strloan = 0
            strloan = Format$(Round(Val(netAmnt)), "###0.00")
            lblNet.Caption = "Total Amount Receivable Tk. " + strloan
                       
    rs.Close
cmdPrint.Enabled = True
cmdCash_Memo.Enabled = True
cmdPrev.Enabled = True
cmdSales.Enabled = False
cmdDelete.Enabled = True
Else
rs.Close
'cmdPrev.Enabled = True
'cmdSales.Enabled = True
End If
End Sub

Private Sub txtLoan_Amnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'txtDown_Payment.SelStart = 0
'txtDown_Payment.SelLength = Len(txtDown_Payment.Text)
'txtDown_Payment.SetFocus
End If
End Sub

Private Sub txtLoan_Amnt_LostFocus()
'txtLoan_Amnt.Text = Format$(Val(txtLoan_Amnt.Text), "###0.00")
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbAgent_Code.SetFocus
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtType.SetFocus
End If
End Sub

Private Sub txtPresent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'txtTelephone.SetFocus
End If
End Sub

Private Sub txtTelephone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMobile.SetFocus
End If
End Sub

Private Sub txtName_LostFocus()
On Error Resume Next
'If txtId.Text = "" Then
'txtName.SetFocus
'Exit Sub
'End If

Dim ID As String
Dim mid As Integer
    ID = txtName.Text
    mid = 0
    
    Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Name like '" & ID & "'"
        rs.Open str, conn

    If Not rs.EOF Then
        'Call clearTextboxes
        txtID.Text = rs!Customer_Code
        'txtDate.Text = Today
        txtType.Text = rs!Customer_Type
        txtName.Text = rs!Customer_Name
        txtAddress.Text = rs!Customer_Address
        txtMobile.Text = rs!Mobile
        cmbAgent_Code.Text = rs!Agent_Code
        cmbAgent_Name.Text = rs!Agent_Name
        txtBalance.Text = Format$(Val(rs!Balance), "###0.00")
        cmbAgent_Code.SetFocus
     rs.Close
     cmdSales.Enabled = True
        
    Else
        
    If MsgBox("Do you want add new Customer?", vbInformation + vbYesNo, "Add New") = vbYes Then
        
        Set rsU = New ADODB.Recordset
            str = "select * from Customer_Master order by Customer_Code"
            rsU.Open str, conn
        
            If Not rsU.EOF Then
                rsU.MoveFirst
        
            Do While Not rsU.EOF = True
                mid = rsU!Customer_Code
                rsU.MoveNext
            Loop
                rsU.Close
                mid = mid + 1
            Else
                rsU.Close
                mid = 1
            End If
            
            txtID.Text = Format(mid, "000#")
            txtDate.Text = Today
            txtType.Text = ""
            txtName.Text = txtName.Text
            txtAddress.Text = ""
            txtMobile.Text = ""
            txtType.SetFocus
        rs.Close
     Else
        rs.Close
        Exit Sub
    End If
    txtType.SetFocus
    'cmdSales.Enabled = True
End If
End Sub

Private Sub txtOther_Amnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
cmdSales.SetFocus
End If
End Sub

Private Sub txtOther_Amnt_LostFocus()
On Error Resume Next
Call Total_Cash
'If Val(Total) <> Val(netAmnt) Then
'MsgBox "Collection differs from Sale Amount" & Val(Total) - Val(netAmnt), vbCritical, "Sale Info!"
'Else
'cmdSales.SetFocus
'End If
End Sub

Private Sub txtOther_Des_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'    txtOther_Amnt.SelStart = 0
'    txtOther_Amnt.SelLength = Len(txtOther_Amnt.Text)
'    txtOther_Amnt.SetFocus
End If
End Sub

'Private Sub txtOther_Des_LostFocus()
'If txtOther_Des.Text = "" Then
'cmdSales.SetFocus
'Exit Sub
'End If
'End Sub

Private Sub txtPercent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtCommission.SetFocus
End If
End Sub
Private Sub txtPercent_LostFocus()
    txtCommission.Text = Format$(Val(txtProd_Price.Text) * (Val(txtPercent.Text) / 100), "###0.00")
    txtCommission.SelStart = 0
    txtCommission.SelLength = Len(txtCommission.Text)
    'txtCommission.SetFocus
End Sub

Private Sub txtProd_Cost_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Price.SelStart = 0
txtProd_Price.SelLength = Len(txtProd_Price.Text)
txtProd_Price.SetFocus
End If
End Sub

Private Sub txtProd_Cost_LostFocus()
If txtProd_Sl = "101" Or txtProd_Sl = "BKS01" Then
txtProd_Price.Text = Format$(Val(txtProd_Cost.Text), "###0.00")
txtProd_Cost.Text = Format$(Val(txtProd_Cost.Text), "###0.00")
Else
txtProd_Price.Text = Format$(Val(txtQty.Text) * Val(txtProd_Cost.Text), "###0.00")
txtProd_Cost.Text = Format$(Val(txtProd_Cost.Text), "###0.00")
End If
End Sub

Private Sub txtProd_Price_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPercent.SelStart = 0
    txtPercent.SelLength = Len(txtPercent.Text)
    txtPercent.SetFocus
End If
End Sub

Private Sub txtProd_Price_LostFocus()
txtProd_Price.Text = Format$(Val(txtProd_Price.Text), "###0.00")
End Sub

Private Sub txtProd_Sl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty.Text)
    txtQty.SetFocus
End If
End Sub


Private Sub txtProd_Sl_LostFocus()
If txtProd_Sl.Text = "" Then
    If MsgBox("Do you want to add new product?", vbExclamation + vbYesNo, "Product Info!") = vbYes Then
        cmbProd_Name.SetFocus
        cmbProd_Name.SelStart = 0
        cmbProd_Name.SelLength = Len(cmbProd_Name.Text)
        cmbProd_Name.SetFocus
    Else
        
        txtCash.SelStart = 0
        txtCash.SelLength = Len(txtCash.Text)
        txtCash.SetFocus
    End If
    
    Exit Sub
End If
On Error Resume Next
    Dim Prod As String
    Prod = txtProd_Sl.Text
    
    Set rsN = New ADODB.Recordset
        'str = "SELECT * FROM Prod_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        str = "select * from Godown_Master where Prod_Code like '" & Prod & "' And Vendor_Code Like '" & cmbAgent_Code.Text & "'"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
'            txtPercent.Text = 0
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtCost.Text = rsN!Prod_Price
            If txtType.Text = "General" Then
            txtProd_Cost.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            End If
            If txtType.Text = "Dealer" Then
            txtProd_Cost.Text = Format$(Val(rsN!Dealer_Price), "###0.00")
            End If
            
            If txtType.Text = "Retail" Then
            txtProd_Cost.Text = Format$(Val(rsN!Retail_Price), "###0.00")
            End If
            rsN.Close
            
        Else
        
        MsgBox "Invalid Goods Code!", vbCritical, "Sales Info!"
            rsN.Close
        Exit Sub
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Cost.SelStart = 0
txtProd_Cost.SelLength = Len(txtProd_Cost.Text)
txtProd_Cost.SetFocus
End If
End Sub
Private Sub txtSale_Price_LostFocus()
'txtSale_Price.Text = Format$(Val(txtSale_Price.Text), "###0.00")
End Sub

Private Sub txtQty_LostFocus()
If cmbType.Text = "Cash Sale" Or cmbType.Text = "Credit Sale" Then
Dim Prod As String
        Prod = txtProd_Sl.Text

    Set rsN = New ADODB.Recordset
        str = "select * from Godown_Master where Prod_Code like '" & Prod & "' And Vendor_Code Like '" & cmbAgent_Code.Text & "'"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            If rsN!Stock >= Val(txtQty.Text) Then
                rsN.Close
                txtProd_Cost.SelStart = 0
                txtProd_Cost.SelLength = Len(txtProd_Cost.Text)
                txtProd_Cost.SetFocus
             Else
             
             MsgBox "Stock not available!", vbCritical, "Sales Info!"
                txtQty.Text = 0
                txtProd_Cost.SelStart = 0
                txtProd_Cost.SelLength = Len(txtProd_Cost.Text)
                txtProd_Cost.SetFocus
                rsN.Close
             Exit Sub
            End If
        Else
        rsN.Close
        Exit Sub
        End If
End If

If cmbType.Text = "Sales Return" Then
        Exit Sub
End If
End Sub


Private Sub txtSl_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txtSl_LostFocus()
If txtProd_Sl.Text = "" Then
    txt1000.SelStart = 0
    txt1000.SelLength = Len(txt1000.Text)
    txt1000.SetFocus
    Exit Sub
End If
If cmbType.Text = "Cash Sale" Or cmbType.Text = "Credit Sale" Then
    
    
    Dim Prod As String
    Dim P_Chasis As String
    'P_Chasis = txtSl.Text
    Prod = txtProd_Sl.Text
    
    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Code Like '" & Prod & "'AND Chasis Like '" & P_Chasis & "' AND Stock >0 order by Chasis"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            'cmbBrand.Text = rsN!Brand_Name
           ' cmbOrigin.Text = rsN!Origin
           ' cmbColor.Text = rsN!Color
            'txtSl.Text = rsN!Chasis
            'txtEngine.Text = rsN!Engine
            txtPercent.Text = 0
            'txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            'txtDate.Text = Today
            If txtType.Text = "General" Then
            txtProd_Cost.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            End If
            If txtType.Text = "Dealer" Then
            txtProd_Cost.Text = Format$(Val(rsN!Dealer_Price), "###0.00")
            End If
            rsN.Close
            txtQty.SelStart = 0
            txtQty.SelLength = Len(txtQty.Text)
            txtQty.SetFocus
         Else
         
         MsgBox "Stock not available! or Invalid SL/Chasis no.!", vbCritical, "Sales Info!"
            'txtSl.SelStart = 0
            'txtSl.SelLength = Len(txtSl.Text)
            'txtSl.SetFocus
            rsN.Close
         Exit Sub
        End If
End If

If cmbType.Text = "Sales Return" Then
    
    'P_Chasis = txtSl.Text
    Prod = txtProd_Sl.Text
    
    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Code Like '" & Prod & "'AND Chasis Like '" & P_Chasis & "' order by Chasis"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            'cmbBrand.Text = rsN!Brand_Name
            'cmbOrigin.Text = rsN!Origin
            'cmbColor.Text = rsN!Color
            'txtSl.Text = rsN!Chasis
            'txtEngine.Text = rsN!Engine
            txtPercent.Text = 0
            'txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            'txtDate.Text = Today
            If txtType.Text = "General" Then
            txtProd_Cost.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            End If
            If txtType.Text = "Dealer" Then
            txtProd_Cost.Text = Format$(Val(rsN!Dealer_Price), "###0.00")
            End If
            rsN.Close
            txtQty.SelStart = 0
            txtQty.SelLength = Len(txtQty.Text)
            txtQty.SetFocus
         Else
         
         MsgBox "Stock not available! or Invalid SL/Chasis no.!", vbCritical, "Sales Info!"
            'txtSl.SelStart = 0
            'txtSl.SelLength = Len(txtSl.Text)
            'txtSl.SetFocus
            rsN.Close
         Exit Sub
        End If
End If
End Sub

Private Sub txtTerms_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCharge.SelStart = 0
    txtCharge.SelLength = Len(txtCharge.Text)
    txtCharge.SetFocus
End If
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAddress.SetFocus
End If
End Sub
