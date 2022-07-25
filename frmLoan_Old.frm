VERSION 5.00
Begin VB.Form frmLoan_Old 
   BackColor       =   &H00008000&
   Caption         =   "Loan Old Data Entry"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   Icon            =   "frmLoan_Old.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
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
      Height          =   4245
      Left            =   5490
      TabIndex        =   70
      Top             =   4560
      Width           =   4335
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
         Left            =   1980
         TabIndex        =   80
         Text            =   "Combo2"
         Top             =   2925
         Width           =   2220
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
         Left            =   1980
         TabIndex        =   79
         Text            =   "Combo2"
         Top             =   2160
         Width           =   2220
      End
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
         Left            =   150
         TabIndex        =   78
         Text            =   "Combo1"
         Top             =   2160
         Width           =   1545
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
         Left            =   150
         TabIndex        =   77
         Text            =   "Combo2"
         Top             =   2925
         Width           =   1545
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
         Left            =   1980
         TabIndex        =   76
         Text            =   "Combo2"
         Top             =   585
         Width           =   2220
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
         Left            =   150
         TabIndex        =   75
         Text            =   "Combo1"
         Top             =   585
         Width           =   1545
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
         Left            =   1980
         TabIndex        =   74
         Text            =   "Combo2"
         Top             =   1395
         Width           =   2220
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
         Left            =   150
         TabIndex        =   73
         Text            =   "Combo6"
         Top             =   1395
         Width           =   1500
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
         Left            =   150
         TabIndex        =   72
         Text            =   "Combo2"
         Top             =   3690
         Width           =   1500
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
         Left            =   1980
         TabIndex        =   71
         Text            =   "Combo2"
         Top             =   3690
         Width           =   2220
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
         Left            =   150
         TabIndex        =   91
         Top             =   2610
         Width           =   1110
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
         Left            =   1980
         TabIndex        =   90
         Top             =   2610
         Width           =   1170
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
         Left            =   1980
         TabIndex        =   89
         Top             =   1845
         Width           =   960
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
         Left            =   150
         TabIndex        =   88
         Top             =   1845
         Width           =   900
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
         Left            =   1980
         TabIndex        =   87
         Top             =   270
         Width           =   1140
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
         Left            =   150
         TabIndex        =   86
         Top             =   270
         Width           =   1080
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
         Left            =   1980
         TabIndex        =   85
         Top             =   1035
         Width           =   1185
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
         TabIndex        =   84
         Top             =   2985
         Width           =   135
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
         Left            =   150
         TabIndex        =   83
         Top             =   1035
         Width           =   1125
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
         Left            =   1980
         TabIndex        =   82
         Top             =   3375
         Width           =   870
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
         Left            =   150
         TabIndex        =   81
         Top             =   3375
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   3435
      Left            =   5490
      TabIndex        =   53
      Top             =   915
      Width           =   4335
      Begin VB.ComboBox cmbGarantorID_1 
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
         Left            =   225
         TabIndex        =   61
         Text            =   "Combo2"
         Top             =   585
         Width           =   1410
      End
      Begin VB.ComboBox cmbGarantorName_1 
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
         Left            =   1980
         TabIndex        =   60
         Text            =   "Combo3"
         Top             =   585
         Width           =   2220
      End
      Begin VB.ComboBox cmbGarantorID_2 
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
         Left            =   225
         TabIndex        =   59
         Text            =   "Combo4"
         Top             =   1350
         Width           =   1410
      End
      Begin VB.ComboBox cmbGarantorName_2 
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
         Left            =   1980
         TabIndex        =   58
         Text            =   "Combo5"
         Top             =   1350
         Width           =   2220
      End
      Begin VB.ComboBox cmbGarantorName_4 
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
         Left            =   1980
         TabIndex        =   57
         Text            =   "Combo5"
         Top             =   2880
         Width           =   2220
      End
      Begin VB.ComboBox cmbGarantorID_4 
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
         Left            =   225
         TabIndex        =   56
         Text            =   "Combo4"
         Top             =   2880
         Width           =   1410
      End
      Begin VB.ComboBox cmbGarantorName_3 
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
         Left            =   1980
         TabIndex        =   55
         Text            =   "Combo3"
         Top             =   2115
         Width           =   2220
      End
      Begin VB.ComboBox cmbGarantorID_3 
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
         Left            =   225
         TabIndex        =   54
         Text            =   "Combo2"
         Top             =   2115
         Width           =   1410
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Guarantor ID"
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
         TabIndex        =   69
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label39 
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
         Left            =   1980
         TabIndex        =   68
         Top             =   270
         Width           =   1650
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Guarantor ID"
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
         TabIndex        =   67
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Label Label11 
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
         Left            =   1980
         TabIndex        =   66
         Top             =   1035
         Width           =   1425
      End
      Begin VB.Label Label2 
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
         Left            =   1980
         TabIndex        =   65
         Top             =   2565
         Width           =   1425
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Guarantor ID"
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
         TabIndex        =   64
         Top             =   2565
         Width           =   1320
      End
      Begin VB.Label Label9 
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
         Left            =   1980
         TabIndex        =   63
         Top             =   1800
         Width           =   1650
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Guarantor ID"
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
         TabIndex        =   62
         Top             =   1800
         Width           =   1320
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
      Height          =   3435
      Left            =   135
      TabIndex        =   28
      Top             =   915
      Width           =   5190
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
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   585
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
         ItemData        =   "frmLoan_Old.frx":0442
         Left            =   3465
         List            =   "frmLoan_Old.frx":0444
         TabIndex        =   39
         Text            =   "Combo1"
         Top             =   585
         Width           =   1545
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
         TabIndex        =   38
         Text            =   "Text4"
         Top             =   2880
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
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtInstallment 
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
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   2115
         Width           =   1545
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
         Left            =   3465
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   1350
         Width           =   1545
      End
      Begin VB.TextBox txtInst_No 
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
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   2115
         Width           =   1455
      End
      Begin VB.TextBox txtDown 
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
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   585
         Width           =   1455
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
         TabIndex        =   32
         Text            =   "Combo1"
         Top             =   2115
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
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2880
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
         Left            =   1845
         TabIndex        =   30
         Text            =   "Combo1"
         Top             =   1350
         Width           =   1455
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
         Left            =   135
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   1350
         Width           =   1590
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
         TabIndex        =   52
         Top             =   2565
         Width           =   1140
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
         TabIndex        =   51
         Top             =   270
         Width           =   1080
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
         TabIndex        =   50
         Top             =   270
         Width           =   1410
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
         TabIndex        =   49
         Top             =   1800
         Width           =   945
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
         TabIndex        =   48
         Top             =   2565
         Width           =   1185
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
         TabIndex        =   47
         Top             =   2565
         Width           =   1020
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Amount"
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
         Top             =   1035
         Width           =   1155
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
         TabIndex        =   45
         Top             =   1800
         Width           =   1515
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
         TabIndex        =   44
         Top             =   270
         Width           =   1440
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
         TabIndex        =   43
         Top             =   1800
         Width           =   1275
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
         Left            =   1845
         TabIndex        =   42
         Top             =   1035
         Width           =   1260
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
         Left            =   135
         TabIndex        =   41
         Top             =   1035
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00C0FFC0&
      Height          =   660
      Left            =   5490
      ScaleHeight     =   600
      ScaleWidth      =   4275
      TabIndex        =   23
      Top             =   8970
      Width           =   4335
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   1125
         TabIndex        =   27
         Top             =   135
         Width           =   945
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   375
         Left            =   3195
         TabIndex        =   26
         Top             =   135
         Width           =   945
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Save"
         Height          =   375
         Left            =   90
         TabIndex        =   25
         Top             =   135
         Width           =   945
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   2160
         TabIndex        =   24
         Top             =   135
         Width           =   945
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
      Height          =   5055
      Left            =   90
      TabIndex        =   0
      Top             =   4560
      Width           =   5190
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
         TabIndex        =   11
         Text            =   "Combo3"
         Top             =   2115
         Width           =   1545
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
         TabIndex        =   10
         Text            =   "Text17"
         Top             =   4500
         Width           =   1590
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
         TabIndex        =   9
         Text            =   "Text17"
         Top             =   4500
         Width           =   1455
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
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   585
         Width           =   2715
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
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   2115
         Width           =   2700
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
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1350
         Width           =   2700
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
         TabIndex        =   5
         Text            =   "Text15"
         Top             =   2880
         Width           =   4815
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
         TabIndex        =   4
         Text            =   "Text16"
         Top             =   3690
         Width           =   4815
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
         TabIndex        =   3
         Text            =   "Text18"
         Top             =   585
         Width           =   1545
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
         TabIndex        =   2
         Text            =   "Combo3"
         Top             =   1350
         Width           =   1545
      End
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
         TabIndex        =   1
         Text            =   "Text17"
         Top             =   4500
         Width           =   1455
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
         TabIndex        =   22
         Top             =   1800
         Width           =   915
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
         TabIndex        =   21
         Top             =   4185
         Width           =   930
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
         TabIndex        =   20
         Top             =   4185
         Width           =   1245
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
         TabIndex        =   19
         Top             =   270
         Width           =   1575
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
         TabIndex        =   18
         Top             =   1800
         Width           =   1275
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
         TabIndex        =   17
         Top             =   1035
         Width           =   1230
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
         TabIndex        =   16
         Top             =   2565
         Width           =   1455
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
         TabIndex        =   15
         Top             =   3375
         Width           =   1725
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
         TabIndex        =   14
         Top             =   1035
         Width           =   690
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
         TabIndex        =   13
         Top             =   270
         Width           =   1485
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
         TabIndex        =   12
         Top             =   4185
         Width           =   1230
      End
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OLD LOAN ACCOUNT OPENING"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   90
      TabIndex        =   92
      Top             =   135
      Width           =   9720
   End
End
Attribute VB_Name = "frmLoan_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Today As Date
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub L_Type()
        cmbLType.Clear
        cmbLType.AddItem "Product"
        cmbLType.AddItem "Cash"
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

        cmbProd_model.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name, Prod_Model FROM Prod_Master where Prod_Name like '" & cmbProduct_Name.Text & "'"
        rs.Open str, conn
        
        If Not rs.EOF Then
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbProd_model.AddItem rs!Prod_Model
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
        cmbProd_model.AddItem "Show All"
End Sub
Private Sub Prod_Price()
     
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name, Prod_Model, Sale_Price FROM Prod_Master where Prod_Name like '" & cmbProduct_Name.Text & "' and Prod_Model like '" & cmbProd_model.Text & "' and Stock > 0"
        rs.Open str, conn
        
        If Not rs.EOF Then
        txtNet.Text = rs!Sale_Price
        rs.Close
        Else
        MsgBox "Product not available!", vbCritical, "Error!"
        rs.Close
        Exit Sub
        End If
    
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
        cmbSamity_name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Samity_Name FROM Samity"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbSamity_name.AddItem rs!Samity_Name
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
Private Sub ComboG_ID1()
        cmbGarantorID_1.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer FROM Deposit_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbGarantorID_1.AddItem rs!Customer
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub
Private Sub ComboG_ID2()
        cmbGarantorID_2.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer FROM Deposit_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbGarantorID_2.AddItem rs!Customer
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub
Private Sub ComboG_ID3()
        cmbGarantorID_3.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer FROM Deposit_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbGarantorID_3.AddItem rs!Customer
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub
Private Sub ComboG_ID4()
        cmbGarantorID_4.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer FROM Deposit_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbGarantorID_4.AddItem rs!Customer
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub
Private Sub ComboG_Name1()
        cmbGarantorName_1.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name FROM Deposit_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbGarantorName_1.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub
Private Sub ComboG_Name2()
        cmbGarantorName_2.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name FROM Deposit_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbGarantorName_2.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub
Private Sub ComboG_Name3()
        cmbGarantorName_3.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name FROM Deposit_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbGarantorName_3.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        Exit Sub
        End If
End Sub
Private Sub ComboG_Name4()
        cmbGarantorName_4.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name FROM Deposit_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbGarantorName_4.AddItem rs!Name
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
        txtID.Text = ""
        cmbLType.Text = ""
        txtDown.Text = ""
        
        txtNet.Text = ""
        txtInstallment.Text = ""
        txtInst_No.Text = ""
        cmbTerm.Text = ""
        
        txtOpen.Text = ""
        txtMaturity.Text = ""
        txtD_Date.Text = ""
        
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
        cmbProd_model.Text = ""
        
        cmbGarantorID_1.Text = ""
        cmbGarantorName_1.Text = ""
        cmbGarantorID_2.Text = ""
        cmbGarantorName_2.Text = ""
        cmbGarantorID_3.Text = ""
        cmbGarantorName_3.Text = ""
        cmbGarantorID_4.Text = ""
        cmbGarantorName_4.Text = ""
        
        cmbCenter_Name.Text = ""
        cmbCenter_Code.Text = ""
        cmbSamity_Code.Text = ""
        cmbSamity_name.Text = ""
        
        cmbFO_Name.Text = ""
        cmbFO_Code.Text = ""
        cmbDPO_Name.Text = ""
        cmbDPO_Code.Text = ""
        cmbAM_Name.Text = ""
        cmbAM_Code.Text = ""
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

Private Sub cmbLType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmbLType.Text = "Cash" Then
txtNet.SetFocus
End If
If cmbLType.Text = "Product" Then
cmbProduct_Name.Enabled = True
cmbProduct_Name.SetFocus
End If
End If
End Sub

Private Sub cmbLType_LostFocus()
    If cmbLType.Text = "Cash" Then
        cmbProduct_Name.Enabled = False
        cmbProd_model.Enabled = False
    End If
    If cmbLType.Text = "Product" Then
        cmbProduct_Name.Enabled = True
        cmbProd_model.Enabled = True
        
    End If

End Sub

Private Sub cmbProd_Model_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtNet.SetFocus
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
cmbProd_model.SetFocus
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
            cmbProd_model.Text = rs!Prod_Model
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
        str = "select * from Loan_Info where Customer like '" & txtID.Text & "'"
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
        
        rsN!Customer = txtID.Text
        rsN!AC_No = Prif + suf + txtID.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Open_Date = txtOpen.Text
        rsN!Mat_Date = txtMaturity.Text
        rsN!D_Date = txtD_Date.Text
        rsN!Down_Payment = txtDown.Text
        rsN!Net_Loan = txtNet.Text
                
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
        
        rsN!G_ID1 = cmbGarantorID_1.Text
        rsN!G_Name1 = cmbGarantorName_1.Text
        rsN!G_ID2 = cmbGarantorID_2.Text
        rsN!G_Name2 = cmbGarantorName_2.Text
        rsN!G_ID3 = cmbGarantorID_3.Text
        rsN!G_Name3 = cmbGarantorName_3.Text
        rsN!G_ID4 = cmbGarantorID_4.Text
        rsN!G_Name4 = cmbGarantorName_4.Text
        
        
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        rsN!Samity_Name = cmbSamity_name.Text
        rsN!Samity_Code = cmbSamity_Code.Text
        rsN!FO_Name = cmbFO_Name.Text
        rsN!FO_Code = cmbFO_Code.Text
        rsN!DPO_Name = cmbDPO_Name.Text
        rsN!DPO_Code = cmbDPO_Code.Text
        rs!AM_Name = cmbAM_Name.Text
        rs!AM_Code = cmbAM_Code.Text
        
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_model.Text
                       
        rsN.Update
        rsN.Close
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Loan_Master", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Open_Date = txtOpen.Text
        rsN!D_Date = txtD_Date.Text
        rsN!Inst_Date = ""
        rsN!Mat_Date = txtMaturity.Text
        rsN!AC_No = Prif + suf + txtID.Text
        rsN!Customer = txtID.Text
        rsN!Name = txtName.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_model.Text
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Inst_Due = txtInst_No.Text
        rsN!Inst_Paid = 0
        rsN!Amount = Val(txtNet.Text)
        rsN!Balance = -Val(txtNet.Text)
        rsN!Due = 0
        rsN!Paid = 0
        rsN!Advance = 0
        rsN!Fine = 0
        rsN!Security = Val(txtDown.Text)
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
        rsN!Samity_Name = cmbSamity_name.Text
        rsN!FO_Name = cmbFO_Name.Text
        rsN!FO_Code = cmbFO_Code.Text
        rsN!DPO_Name = cmbDPO_Name.Text
        rsN!DPO_Code = cmbDPO_Code.Text
        rs!AM_Name = cmbAM_Name.Text
        rs!AM_Code = cmbAM_Code.Text
        rsN.Update
        rsN.Close
    
Else
'---------------------------------------------------------------------
    Set rsN = New ADODB.Recordset
    rsN.Open "Loan_info", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Customer = txtID.Text
        rsN!AC_No = Prif + suf + txtID.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Open_Date = txtOpen.Text
        rsN!Mat_Date = txtMaturity.Text
        rsN!D_Date = txtD_Date.Text
        rsN!Down_Payment = txtDown.Text
        rsN!Net_Loan = txtNet.Text
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
        
        rsN!G_ID1 = cmbGarantorID_1.Text
        rsN!G_Name1 = cmbGarantorName_1.Text
        rsN!G_ID2 = cmbGarantorID_2.Text
        rsN!G_Name2 = cmbGarantorName_2.Text
        
        rsN!G_ID3 = cmbGarantorID_3.Text
        rsN!G_Name3 = cmbGarantorName_3.Text
        rsN!G_ID4 = cmbGarantorID_4.Text
        rsN!G_Name4 = cmbGarantorName_4.Text
        
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        rsN!Samity_Name = cmbSamity_name.Text
        rsN!Samity_Code = cmbSamity_Code.Text
        rsN!FO_Name = cmbFO_Name.Text
        rsN!FO_Code = cmbFO_Code.Text
        rsN!DPO_Name = cmbDPO_Name.Text
        rsN!DPO_Code = cmbDPO_Code.Text
        rs!AM_Name = cmbAM_Name.Text
        rs!AM_Code = cmbAM_Code.Text
        
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_model.Text
        
        rsN.Update
        rsN.Close
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Loan_Master", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Open_Date = txtOpen.Text
        rsN!D_Date = txtD_Date.Text
        rsN!Inst_Date = ""
        rsN!Mat_Date = txtMaturity.Text
        rsN!AC_No = Prif + suf + txtID.Text
        rsN!Customer = txtID.Text
        rsN!Name = txtName.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_model.Text
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Inst_Due = txtInst_No.Text
        rsN!Inst_Paid = 0
        rsN!Amount = Val(txtNet.Text)
        rsN!Balance = -Val(txtNet.Text)
        rsN!Due = 0
        rsN!Paid = 0
        rsN!Advance = 0
        rsN!Fine = 0
        rsN!Security = Val(txtDown.Text)
        
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
        rsN!Samity_Name = cmbSamity_name.Text
        rsN!FO_Name = cmbFO_Name.Text
        rsN!FO_Code = cmbFO_Code.Text
        rsN!DPO_Name = cmbDPO_Name.Text
        rsN!DPO_Code = cmbDPO_Code.Text
        rs!AM_Name = cmbAM_Name.Text
        rs!AM_Code = cmbAM_Code.Text
        
        rsN.Update
        rsN.Close
  
End If
    cmdAdd.Enabled = False
    
    MsgBox "Account Open Successfull! Your Account no.: " & Prif + suf + txtID.Text, vbInformation, "Loan Info"
    Call clearTextboxes
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
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
    
  Set rsN = New ADODB.Recordset
        str = "select * from Loan_Info where Customer like '" & txtID.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsN.EOF Then
        
        rsN!Customer = txtID.Text
        rsN!AC_No = Prif + suf + txtID.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Open_Date = txtOpen.Text
        rsN!Mat_Date = txtMaturity.Text
        rsN!D_Date = txtD_Date.Text
        rsN!Down_Payment = txtDown.Text
        rsN!Net_Loan = txtNet.Text
                
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
        
        rsN!G_ID1 = cmbGarantorID_1.Text
        rsN!G_Name1 = cmbGarantorName_1.Text
        rsN!G_ID2 = cmbGarantorID_2.Text
        rsN!G_Name2 = cmbGarantorName_2.Text
        rsN!G_ID3 = cmbGarantorID_3.Text
        rsN!G_Name3 = cmbGarantorName_3.Text
        rsN!G_ID4 = cmbGarantorID_4.Text
        rsN!G_Name4 = cmbGarantorName_4.Text
        
        
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        rsN!Samity_Name = cmbSamity_name.Text
        rsN!Samity_Code = cmbSamity_Code.Text
        rsN!FO_Name = cmbFO_Name.Text
        rsN!FO_Code = cmbFO_Code.Text
        rsN!DPO_Name = cmbDPO_Name.Text
        rsN!DPO_Code = cmbDPO_Code.Text
        rs!AM_Name = cmbAM_Name.Text
        rs!AM_Code = cmbAM_Code.Text
        
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_model.Text
                       
        rsN.Update
        rsN.Close
    
    Set rsN = New ADODB.Recordset
        str = "select * from Loan_Master where Customer like '" & txtID.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
        rsN!Open_Date = txtOpen.Text
        rsN!D_Date = txtD_Date.Text
        rsN!Inst_Date = ""
        rsN!Mat_Date = txtMaturity.Text
        rsN!AC_No = Prif + suf + txtID.Text
        rsN!Customer = txtID.Text
        rsN!Name = txtName.Text
        rsN!Type = cmbLType.Text
        rsN!Term = cmbTerm.Text
        rsN!Prod_Name = cmbProduct_Name.Text
        rsN!Prod_Model = cmbProd_model.Text
        rsN!Installment = txtInstallment.Text
        rsN!Inst_no = txtInst_No.Text
        rsN!Inst_Due = txtInst_No.Text
        rsN!Inst_Paid = 0
        rsN!Amount = Val(txtNet.Text)
        rsN!Balance = -Val(txtNet.Text)
        rsN!Due = 0
        rsN!Paid = 0
        rsN!Advance = 0
        rsN!Fine = 0
        rsN!Security = Val(txtDown.Text)
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
        rsN!Samity_Name = cmbSamity_name.Text
        rsN!FO_Name = cmbFO_Name.Text
        rsN!FO_Code = cmbFO_Code.Text
        rsN!DPO_Name = cmbDPO_Name.Text
        rsN!DPO_Code = cmbDPO_Code.Text
        rs!AM_Name = cmbAM_Name.Text
        rs!AM_Code = cmbAM_Code.Text
        rsN.Update
        rsN.Close
        
    MsgBox "Account Update Successfull! ", vbInformation, "Personal Info"
    Call clearTextboxes
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
        Call ComboG_ID1
        Call ComboG_ID2
        Call ComboG_Name1
        Call ComboG_Name2
        Call ComboG_ID3
        Call ComboG_ID4
        Call ComboG_Name3
        Call ComboG_Name4
        
        Call clearTextboxes
        
Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
            On Error Resume Next
            Today = rs!Today
            txtOpen.Text = Today
            rs.Close
        End If
'-----------------------------------------
cmbProduct_Name.Enabled = False
cmbProd_model.Enabled = False
'--------------------------------------------

cmdAdd.Enabled = False
cmdPrint.Enabled = False
cmdUpdate.Enabled = False
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDown.SetFocus
End If
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
cmbGarantorID_1.SetFocus
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

Private Sub txtFather_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMother.SetFocus
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
    ID = txtID.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Loan_Info where Customer like '" & ID & "'"
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
                txtID.Text = rs!Customer
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
                cmbSamity_name.Text = rs!Samity_Name
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
        
        txtDown.Text = rs!Amount + rs!Advance
        rs.Close
                
        
        cmdUpdate.Enabled = False
        cmdAdd.Enabled = True
        'cmdDelete.Enabled = True
        cmdPrint.Enabled = False
    
    Else
        
    Call clearTextboxes
    On Error Resume Next
        txtID.Text = rs!Customer
        cmbLType.Text = rs!Type
        cmbTerm.Text = rs!Term
        
        txtDown.Text = rs!Down_Payment
        
        txtNet.Text = rs!Net_Loan
        
        txtInstallment.Text = rs!Installment
        txtInst_No.Text = rs!Inst_no
        txtOpen.Text = rs!Open_Date
        txtD_Date.Text = rs!D_Date
        txtMaturity.Text = rs!Mat_Date
        
        
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
        cmbSamity_name.Text = rs!Samity_Name
        cmbFO_Name.Text = rs!FO_Name
        cmbFO_Code.Text = rs!FO_Code
        cmbDPO_Name.Text = rs!DPO_Name
        cmbDPO_Code.Text = rs!DPO_Code
        cmbAM_Name.Text = rs!AM_Name
        cmbAM_Code.Text = rs!AM_Code
        
        cmbProduct_Name.Text = rs!Prod_Name
        cmbProd_model.Text = rs!Prod_Model
   
        
        cmbGarantorID_1.Text = rs!G_ID1
        cmbGarantorName_1.Text = rs!G_Name1
        cmbGarantorID_2.Text = rs!G_ID2
        cmbGarantorName_2.Text = rs!G_Name2
        
        cmbGarantorID_3.Text = rs!G_ID3
        cmbGarantorName_3.Text = rs!G_Name3
        cmbGarantorID_4.Text = rs!G_ID4
        cmbGarantorName_4.Text = rs!G_Name4
        
        
        rs.Close
        
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
        txtID.Text = rs!Customer
        
        
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
        cmbSamity_name.Text = rs!Samity_Name
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
        
        txtDown.Text = rs!Amount
        rs.Close
        
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
txtInst_No.SetFocus
End If
End Sub

Private Sub txtMaturity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtD_Date.SetFocus
End If
End Sub

Private Sub txtNet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtInstallment.SetFocus
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
cmbGarantorID_1.SetFocus
End If
End Sub
Private Sub cmbReligion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbNationality.SetFocus
End If
End Sub

