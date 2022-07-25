VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmployee 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Entry"
   ClientHeight    =   10395
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   9870
   Icon            =   "frmEmployee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Service Record"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   135
      TabIndex        =   70
      Top             =   8055
      Width           =   5280
      Begin VB.TextBox txtS_Record 
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
         Height          =   915
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   71
         Text            =   "frmEmployee.frx":0442
         Top             =   270
         Width           =   5055
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Payment Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   5625
      TabIndex        =   56
      Top             =   7335
      Width           =   4065
      Begin VB.TextBox txtMobile_Bill 
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
         Height          =   390
         Left            =   2670
         TabIndex        =   78
         Text            =   "0.00"
         Top             =   630
         Width           =   1170
      End
      Begin VB.ComboBox cmbPercent 
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
         ItemData        =   "frmEmployee.frx":0448
         Left            =   1665
         List            =   "frmEmployee.frx":0467
         TabIndex        =   63
         Text            =   "0"
         Top             =   1395
         Width           =   645
      End
      Begin VB.TextBox txtSalary 
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
         Height          =   390
         Left            =   135
         TabIndex        =   62
         Text            =   "0.00"
         Top             =   1395
         Width           =   1380
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
         Height          =   390
         Left            =   2475
         TabIndex        =   61
         Text            =   "0.00"
         Top             =   1395
         Width           =   1380
      End
      Begin VB.TextBox txtTarget 
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
         Height          =   390
         Left            =   105
         TabIndex        =   58
         Text            =   "0.00"
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtAchive 
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
         Height          =   390
         Left            =   1350
         TabIndex        =   57
         Text            =   "0.00"
         Top             =   630
         Width           =   1125
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile Bill"
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
         Left            =   2670
         TabIndex        =   79
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label13 
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
         Left            =   1890
         TabIndex        =   66
         Top             =   1170
         Width           =   180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
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
         TabIndex        =   65
         Top             =   1125
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commission"
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
         Left            =   2655
         TabIndex        =   64
         Top             =   1125
         Width           =   1080
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Traget"
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
         TabIndex        =   60
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Achivement"
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
         Left            =   1380
         TabIndex        =   59
         Top             =   315
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
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
      Height          =   4695
      Left            =   135
      TabIndex        =   37
      Top             =   3240
      Width           =   5280
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
         Left            =   1800
         TabIndex        =   46
         Text            =   "Combo3"
         Top             =   4050
         Width           =   1545
      End
      Begin VB.TextBox txtBirthday 
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
         Left            =   135
         TabIndex        =   45
         Text            =   "Text18"
         Top             =   4050
         Width           =   1545
      End
      Begin VB.TextBox txtPermanent 
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
         Left            =   135
         TabIndex        =   44
         Text            =   "Text16"
         Top             =   3195
         Width           =   4905
      End
      Begin VB.TextBox txtPresent 
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
         Left            =   135
         TabIndex        =   43
         Text            =   "Text15"
         Top             =   2340
         Width           =   4905
      End
      Begin VB.TextBox txtFather 
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
         Left            =   135
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   630
         Width           =   2700
      End
      Begin VB.TextBox txtMother 
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
         Left            =   135
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   1485
         Width           =   2700
      End
      Begin VB.TextBox txtContact_Home 
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
         Left            =   3555
         TabIndex        =   40
         Text            =   "Text17"
         Top             =   630
         Width           =   1545
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
         Height          =   375
         Left            =   3555
         TabIndex        =   39
         Text            =   "Text17"
         Top             =   1440
         Width           =   1545
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
         Left            =   3555
         TabIndex        =   38
         Text            =   "Combo3"
         Top             =   4050
         Width           =   1545
      End
      Begin VB.Label Label23 
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
         Left            =   135
         TabIndex        =   55
         Top             =   3690
         Width           =   1485
      End
      Begin VB.Label Label22 
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
         Left            =   1800
         TabIndex        =   54
         Top             =   3690
         Width           =   690
      End
      Begin VB.Label Label21 
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
         Left            =   135
         TabIndex        =   53
         Top             =   2835
         Width           =   1725
      End
      Begin VB.Label Label20 
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
         Left            =   135
         TabIndex        =   52
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fathers / Husband Name"
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
         Width           =   2175
      End
      Begin VB.Label Label18 
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
         Left            =   135
         TabIndex        =   50
         Top             =   1125
         Width           =   1275
      End
      Begin VB.Label Label14 
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
         Left            =   3555
         TabIndex        =   49
         Top             =   270
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
         Left            =   3555
         TabIndex        =   48
         Top             =   1080
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
         Left            =   3555
         TabIndex        =   47
         Top             =   3690
         Width           =   915
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Security Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   5625
      TabIndex        =   28
      Top             =   3240
      Width           =   4065
      Begin VB.TextBox txtDeposit 
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
         Left            =   2115
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   225
         Width           =   1725
      End
      Begin VB.TextBox txtChk 
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
         Left            =   2115
         TabIndex        =   31
         Text            =   "Text7"
         Top             =   720
         Width           =   1725
      End
      Begin VB.ComboBox cmbBranch 
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
         Left            =   2115
         TabIndex        =   30
         Text            =   "Combo2"
         Top             =   1440
         Width           =   1770
      End
      Begin VB.ComboBox cmbBank 
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
         Top             =   1440
         Width           =   1860
      End
      Begin VB.Label lblCash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit Amount"
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
         TabIndex        =   36
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
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
         TabIndex        =   34
         Top             =   1170
         Width           =   1020
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
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
         Left            =   2115
         TabIndex        =   33
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bond Cheque No."
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
         TabIndex        =   32
         Top             =   810
         Width           =   1530
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
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
      Height          =   1860
      Left            =   5625
      TabIndex        =   24
      Top             =   5355
      Width           =   4110
      Begin VB.TextBox txtG_Name3 
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
         Left            =   1305
         TabIndex        =   69
         Text            =   "Text7"
         Top             =   1305
         Width           =   2625
      End
      Begin VB.TextBox txtG_Name2 
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
         Left            =   1305
         TabIndex        =   68
         Text            =   "Text7"
         Top             =   810
         Width           =   2625
      End
      Begin VB.TextBox txtG_Name1 
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
         Left            =   1305
         TabIndex        =   67
         Text            =   "Text7"
         Top             =   270
         Width           =   2670
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Guarantor"
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
         TabIndex        =   27
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Guarantor"
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
         TabIndex        =   26
         Top             =   855
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Guarantor"
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
         TabIndex        =   25
         Top             =   1350
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Supervisor Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   5625
      TabIndex        =   19
      Top             =   9450
      Width           =   4125
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
         Left            =   630
         TabIndex        =   21
         Top             =   270
         Width           =   1110
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
         Left            =   2385
         TabIndex        =   20
         Top             =   270
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         TabIndex        =   23
         Top             =   315
         Width           =   450
      End
      Begin VB.Label Label8 
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
         Left            =   1800
         TabIndex        =   22
         Top             =   315
         Width           =   510
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   135
      TabIndex        =   10
      Top             =   900
      Width           =   5295
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
         Left            =   150
         TabIndex        =   77
         Text            =   "Combo1"
         Top             =   1020
         Width           =   3345
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
         Left            =   1350
         TabIndex        =   76
         Text            =   "Combo1"
         Top             =   270
         Width           =   1605
      End
      Begin VB.TextBox txtDOJ 
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
         Left            =   3630
         TabIndex        =   74
         Top             =   1005
         Width           =   1560
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
         Left            =   2385
         TabIndex        =   72
         Text            =   "Combo1"
         Top             =   1725
         Width           =   1140
      End
      Begin VB.ComboBox cmbDesi 
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
         TabIndex        =   13
         Top             =   1725
         Width           =   2145
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
         Height          =   360
         Left            =   3630
         TabIndex        =   12
         Top             =   240
         Width           =   1560
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
         Left            =   3630
         TabIndex        =   11
         Top             =   1725
         Width           =   1560
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Joining"
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
         Left            =   3630
         TabIndex        =   75
         Top             =   690
         Width           =   1275
      End
      Begin VB.Label Label17 
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
         Left            =   2385
         TabIndex        =   73
         Top             =   1410
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
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
         Top             =   300
         Width           =   1110
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
         Left            =   3060
         TabIndex        =   17
         Top             =   315
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Top             =   675
         Width           =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
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
         TabIndex        =   15
         Top             =   1410
         Width           =   1140
      End
      Begin VB.Label Label7 
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
         Left            =   3630
         TabIndex        =   14
         Top             =   1410
         Width           =   1140
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Photograph"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   5625
      TabIndex        =   7
      Top             =   900
      Width           =   4065
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   1755
         Top             =   495
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1755
         Top             =   1125
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Add Photo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   225
         Left            =   495
         MouseIcon       =   "frmEmployee.frx":048F
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1935
         Width           =   915
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Add Thumb"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   225
         Left            =   2565
         MouseIcon       =   "frmEmployee.frx":08D1
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1935
         Width           =   1005
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1545
         Left            =   2250
         Stretch         =   -1  'True
         Top             =   315
         Width           =   1635
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1545
         Left            =   165
         Stretch         =   -1  'True
         Top             =   315
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   135
      ScaleHeight     =   675
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   9495
      Width           =   5295
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Add New"
         Height          =   435
         Left            =   135
         TabIndex        =   6
         Top             =   135
         Width           =   1155
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   435
         Left            =   1440
         TabIndex        =   5
         Top             =   135
         Width           =   1155
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   435
         Left            =   2700
         TabIndex        =   4
         Top             =   135
         Width           =   1155
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   435
         Left            =   3960
         TabIndex        =   3
         Top             =   135
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   150
      Width           =   9585
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EMPLOYEE ENTRY FORM"
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
         Left            =   1245
         TabIndex        =   1
         Top             =   45
         Width           =   7440
      End
      Begin VB.Image Image3 
         Height          =   645
         Left            =   0
         Picture         =   "frmEmployee.frx":0D13
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9615
      End
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim Today As Date
Dim Photo As String
Dim Thumb As String
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Private Sub Emp_Name()
        txtName.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name FROM Employee"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtName.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
End Sub

Private Sub Emp_Id()
        txtID.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_Id FROM Employee"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtID.AddItem rs!Staff_Id
        rs.MoveNext
        Loop
        rs.Close
End Sub


Private Sub Designation()
cmbDesi.Clear
Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Designation FROM Employee"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbDesi.AddItem rs!Designation
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub Center()
    cmbCenter_Name.Clear
      Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Branch_Name FROM Branch"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbCenter_Name.AddItem rsN!Branch_Name
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub Centercode()
    cmbCenter_Code.Clear
      Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Branch_Code FROM Branch"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbCenter_Code.AddItem rsN!Branch_Code
        rsN.MoveNext
        Loop
        rsN.Close
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
Private Sub DPO_Code()
cmbDPO_Code.Clear
Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_Id FROM Employee"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbDPO_Code.AddItem rs!Staff_Id
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub DPO_Name()
cmbDPO_Name.Clear
Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name FROM Employee"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbDPO_Name.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub clearTextboxes()
        
        txtSalary.Text = ""
        txtCommission.Text = ""
        txtTarget.Text = ""
        txtAchive.Text = ""
        cmbDPO_Code.Text = ""
        cmbDPO_Name.Text = ""
        
        
        txtDate.Text = ""
        txtDOJ.Text = ""
        txtID.Text = ""
        txtName.Text = ""
        cmbDesi.Text = ""
        cmbCenter_Name.Text = ""
        cmbCenter_Code.Text = ""
        
        txtFather.Text = ""
        txtMother.Text = ""
        txtBirthday.Text = ""
        cmbReligion.Text = ""
        cmbNationality.Text = ""
        txtPresent.Text = ""
        txtPermanent.Text = ""
        txtContact_Home.Text = ""
        txtMobile.Text = ""
        
        Photo = ""
        Thumb = ""
        Image1.Picture = LoadPicture("")
        Image2.Picture = LoadPicture("")
        
        txtG_Name1.Text = ""
        txtG_Name2.Text = ""
        txtG_Name3.Text = ""
        txtS_Record.Text = ""
        
        txtDeposit.Text = "0.00"
        txtChk.Text = ""
        cmbBank.Text = ""
        cmbBranch.Text = ""
        txtTarget.Text = "0.00"
        txtAchive.Text = "0.00"
        txtSalary.Text = "0.00"
        txtMobile_Bill.Text = "0.00"
        txtCommission.Text = "0.00"
        cmbPercent.Text = "0"
        cmbDPO_Code.Text = ""
        cmbDPO_Name.Text = ""
End Sub

Private Sub cmbCenter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTarget.SelStart = 0
    txtTarget.SelLength = Len(txtTarget.Text)
    txtTarget.SetFocus
End If
End Sub

Private Sub cmbBank_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbBranch.SetFocus
End If
End Sub

Private Sub cmbBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtG_Name1.SetFocus
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

Dim Code As String
Code = cmbCenter_Code.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Branch_Code, Branch_name FROM Branch where Branch_Code like '" & Code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbCenter_Name.Text = rs!Branch_Name
        rs.Close
        End If
End Sub

Private Sub cmbCenter_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtFather.SetFocus
End If
End Sub

Private Sub cmbCenter_Name_LostFocus()
If cmbCenter_Name.Text = "" Then
Exit Sub
End If

Dim Code As String
Code = cmbCenter_Name.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Branch_Code, Branch_name FROM Branch where Branch_Name like '" & Code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbCenter_Code.Text = rs!Branch_Code
        rs.Close
        End If
End Sub

Private Sub cmbDesi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbCenter_Code.SetFocus
End If
End Sub

Private Sub cmbDPO_Code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbDPO_Name.SetFocus
End If
End Sub

Private Sub cmbDPO_Code_LostFocus()
Dim Search_Id As String
Search_Id = cmbDPO_Code.Text
   
    
    Set rs = New ADODB.Recordset
        str = "select * from Employee where Staff_Id like '" & Search_Id & "' order by Staff_id"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        cmbDPO_Code.Text = rs!Staff_Id
        cmbDPO_Name.Text = rs!Name
        rs.Close
    
    Else
    MsgBox "Invalid Staff code!", vbCritical
    rs.Close
    End If
End Sub

Private Sub cmbDPO_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmdReceive.Enabled = True Then
cmdReceive.SetFocus
Else
Exit Sub
End If
End If
End Sub

Private Sub cmbDPO_Name_LostFocus()
Dim Search_Id As String
Search_Id = cmbDPO_Name.Text
       
    Set rs = New ADODB.Recordset
        str = "select * from Employee where Name like '" & Search_Id & "' order by Staff_id"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        cmbDPO_Code.Text = rs!Staff_Id
        cmbDPO_Name.Text = rs!Name
        rs.Close
    Else
    MsgBox "Invalid Staff code!", vbCritical
    rs.Close
    End If
End Sub

Private Sub cmbNationality_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtContact_Home.SetFocus
End If
End Sub

Private Sub cmbPercent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtCommission.SetFocus
End If
End Sub

Private Sub cmbPercent_LostFocus()
txtCommission.Text = Format$(Val(txtAchive.Text) * (Val(cmbPercent.Text) / 100), "###0.00")
    txtCommission.SelStart = 0
    txtCommission.SelLength = Len(txtCommission.Text)
    txtCommission.SetFocus
End Sub
Private Sub cmbReligion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbNationality.SetFocus
End If
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
cmdReceive.Enabled = False
cmdUpdate.Enabled = False
cmdDelete.Enabled = False
End Sub
Private Sub cmdReceive_Click()
Set rsN = New ADODB.Recordset
        rsN.Open "Employee", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!DOJ = txtDOJ.Text
        rsN!Staff_Id = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesi.Text
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        
        rsN!Fathers_Name = txtFather.Text
        rsN!Mothers_Name = txtMother.Text
        rsN!Age = txtBirthday.Text
        rsN!Religion = cmbReligion.Text
        rsN!Nationality = cmbNationality.Text
        rsN!Present_Address = txtPresent.Text
        rsN!Permanent_Address = txtPermanent.Text
        rsN!Telephone = txtContact_Home.Text
        rsN!Mobile = txtMobile.Text
        
        rsN!Photo = Photo
        rsN!Thumb = Thumb
        
        rsN!G_Name1 = txtG_Name1.Text
        rsN!G_Name2 = txtG_Name2.Text
        rsN!G_Name3 = txtG_Name3.Text
        rsN!S_Record = txtS_Record.Text
        
        rsN!Deposit = txtDeposit.Text
        rsN!Chk_No = txtChk.Text
        rsN!Bank = cmbBank.Text
        rsN!Branch = cmbBranch.Text
        rsN!Target = txtTarget.Text
        rsN!Achivement = txtAchive.Text
        rsN!Mobile_Bill = txtMobile_Bill
        rsN!Salary = txtSalary.Text
        rsN!Commission = txtCommission.Text
        rsN!S_Code = cmbDPO_Code.Text
        rsN!S_Name = cmbDPO_Name.Text
        rsN.Update
        rsN.Close

Set rsN = New ADODB.Recordset
        rsN.Open "Staff_Master", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!D_ate = txtDate.Text
        rsN!Staff_Id = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesi.Text
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        rsN!Salary = 0
        rsN!Commission = 0
        rsN!Draw = 0
        rsN!Advance = 0
        rsN!Sales_Due = 0
        rsN!Balance = 0
        rsN!Salary_Due = txtSalary.Text
        rsN!Mobile_Due = txtMobile_Bill.Text
        rsN!Target = txtTarget.Text
        rsN.Update
        rsN.Close



Call Center
Call Designation
Call Emp_Id
Call Emp_Name
Call clearTextboxes
Call DPO_Code
Call DPO_Name
txtDate.Text = Date
txtID.SetFocus
cmdReceive.Enabled = False
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
Set rsN = New ADODB.Recordset
        str = "select * from Employee where Staff_Id like '" & txtID.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
        If Not rsN.EOF Then
        rsN!Date = txtDate.Text
        rsN!DOJ = txtDOJ.Text
        rsN!Staff_Id = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesi.Text
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        
        rsN!Fathers_Name = txtFather.Text
        rsN!Mothers_Name = txtMother.Text
        rsN!Age = txtBirthday.Text
        rsN!Religion = cmbReligion.Text
        rsN!Nationality = cmbNationality.Text
        rsN!Present_Address = txtPresent.Text
        rsN!Permanent_Address = txtPermanent.Text
        rsN!Telephone = txtContact_Home.Text
        rsN!Mobile = txtMobile.Text
        
        rsN!Photo = Photo
        rsN!Thumb = Thumb
        
        rsN!G_Name1 = txtG_Name1.Text
        rsN!G_Name2 = txtG_Name2.Text
        rsN!G_Name3 = txtG_Name3.Text
        rsN!S_Record = txtS_Record.Text
        
        rsN!Deposit = txtDeposit.Text
        rsN!Chk_No = txtChk.Text
        rsN!Bank = cmbBank.Text
        rsN!Branch = cmbBranch.Text
        rsN!Target = txtTarget.Text
        rsN!Achivement = txtAchive.Text
        rsN!Mobile_Bill = txtMobile_Bill.Text
        rsN!Salary = txtSalary.Text
        rsN!Commission = txtCommission.Text
        rsN!S_Code = cmbDPO_Code.Text
        rsN!S_Name = cmbDPO_Name.Text
        
        rsN.Update
        rsN.Close
        
        
        Set rsN = New ADODB.Recordset
        str = "select * from Staff_Master where Staff_Id like '" & txtID.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
        rsN!D_ate = txtDate.Text
        rsN!Staff_Id = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesi.Text
        rsN!Center_name = cmbCenter_Name.Text
        rsN!Center_code = cmbCenter_Code.Text
        'rsN!Salary = 0
        'rsN!Commission = 0
        'rsN!Draw = 0
        'rsN!Advance = 0
        'rsN!Sales_Due = 0
        'rsN!Balance = 0
        rsN!Salary_Due = txtSalary.Text
        rsN!Mobile_Due = txtMobile_Bill.Text
        rsN!Target = txtTarget.Text
        rsN.Update
        rsN.Close
        
        Else
        rsN.Close
        MsgBox "Invalid Employee ID!"
        txtID.SetFocus
        End If

Call Center
Call Designation
Call clearTextboxes
Call DPO_Code
Call DPO_Name
txtDate.Text = Today
txtID.SetFocus
cmdReceive.Enabled = False
cmdUpdate.Enabled = False
cmdDelete.Enabled = False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
           On Error Resume Next
           Today = rs!Today
           rs.Close
           
        End If
        Call clearTextboxes
        Call Emp_Id
        Call Designation
        Call Center
        Call Centercode
        Call ComboReligion
        Call ComboNationality
        Call DPO_Code
        Call DPO_Name
        txtDate.Text = Today
        cmdReceive.Enabled = False
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
End Sub

Private Sub lblChq_No_Click()
End Sub

Private Sub Label24_Click()

End Sub

Private Sub Label34_Click()
CommonDialog1.ShowOpen
CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "JPEG Images|*.JPG|Bitmaps|*.BMP|GIF Images|*.GIF|All Images|*.BMP;*.GIF;*.JPG"
If CommonDialog1.FileTitle = "" Then
Image1.Picture = LoadPicture("")
Photo = ""
Else
Image1.Picture = LoadPicture(App.Path + "\Photo\" & CommonDialog1.FileTitle)
Photo = CommonDialog1.FileTitle
End If
End Sub

Private Sub Label35_Click()
CommonDialog2.ShowOpen
CommonDialog2.InitDir = App.Path
CommonDialog2.Filter = "JPEG Images|*.JPG|Bitmaps|*.BMP|GIF Images|*.GIF|All Images|*.BMP;*.GIF;*.JPG"
If CommonDialog2.FileTitle = "" Then
Image2.Picture = LoadPicture("")
Thumb = ""
Else
Image2.Picture = LoadPicture(App.Path + "\Thumb\" & CommonDialog2.FileTitle)
Thumb = CommonDialog2.FileTitle
End If
End Sub

Private Sub Text5_Change()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSalary.SelStart = 0
    txtSalary.SelLength = Len(txtSalary.Text)
    txtSalary.SetFocus
End If
End Sub

Private Sub txtBirthday_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbReligion.SetFocus
End If
End Sub

Private Sub txtChk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbBank.SetFocus
End If
End Sub

Private Sub txtCommission_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbDPO_Code.SetFocus
End If
End Sub
Private Sub txtCommission_LostFocus()
txtCommission.Text = Format$(Val(txtCommission.Text), "###0.00")
End Sub

Private Sub txtContact_Home_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMobile.SetFocus
End If
End Sub

Private Sub txtDeposit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtChk.SetFocus
End If
End Sub

Private Sub txtDOJ_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbDesi.SetFocus
End If
End Sub

Private Sub txtFather_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMother.SetFocus
End If
End Sub

Private Sub txtG_Name1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtG_Name2.SetFocus
End If
End Sub

Private Sub txtG_Name2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtG_Name3.SetFocus
End If
End Sub

Private Sub txtG_Name3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTarget.SetFocus
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtName.SetFocus
End If
End Sub

Private Sub txtId_LostFocus()
Dim mid As Integer
Dim Com As Integer
    Search_Id = txtID.Text
    mid = 0
    Com = 0
    
    
    Set rs = New ADODB.Recordset
        str = "select * from Employee where Staff_Id like '" & Search_Id & "' order by Staff_id"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
    Call clearTextboxes
        txtID.Text = rs!Staff_Id
        txtDate.Text = rs!Date
        txtDOJ.Text = rs!DOJ
        txtName.Text = rs!Name
        cmbDesi.Text = rs!Designation
        cmbCenter_Name.Text = rs!Center_name
        cmbCenter_Code.Text = rs!Center_code
        
        txtFather.Text = rs!Fathers_Name
        txtMother.Text = rs!Mothers_Name
        txtBirthday.Text = rs!Age
        cmbReligion.Text = rs!Religion
        cmbNationality.Text = rs!Nationality
        txtPresent.Text = rs!Present_Address
        txtPermanent.Text = rs!Permanent_Address
        txtContact_Home.Text = rs!Telephone
        txtMobile.Text = rs!Mobile
        
        txtG_Name1.Text = rs!G_Name1
        txtG_Name2.Text = rs!G_Name2
        txtG_Name3.Text = rs!G_Name3
        txtS_Record.Text = rs!S_Record
        
        txtDeposit.Text = rs!Deposit
        txtChk.Text = rs!Chk_No
        cmbBank.Text = rs!Bank
        cmbBranch.Text = rs!Branch

        
        Image1.Picture = LoadPicture(App.Path + "\Photo\" & rs!Photo)
        Image2.Picture = LoadPicture(App.Path + "\Thumb\" & rs!Thumb)
        
        Photo = rs!Photo
        Thumb = rs!Thumb
        
        txtTarget.Text = Format$(Val(rs!Target), "###0.00")
        txtSalary.Text = Format$(Val(rs!Salary), "###0.00")
        txtMobile_Bill.Text = Format$(Val(rs!Mobile_Bill), "###0.00")
        txtCommission.Text = Format$(Val(rs!Commission), "###0.00")
        
        cmbDPO_Code.Text = rs!S_Code
        cmbDPO_Name.Text = rs!S_Name
        rs.Close
               
               
               
        cmdReceive.Enabled = False
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = True
    
    Else
    If txtID.Text = "" Then
    
        If MsgBox("Do you want add new Employee with Auto ID", vbInformation + vbYesNo, "Add New") = vbYes Then
            
            Set rsU = New ADODB.Recordset
            str = "select * from Employee order by Staff_Id"
            rsU.Open str, conn
        
                If Not rsU.EOF Then
                    rsU.MoveFirst
                    
                    Do While Not rsU.EOF = True
                        mid = Val(rsU!Staff_Id)
                        rsU.MoveNext
                    Loop
                        rsU.Close
                        mid = mid + 1
                Else
                    rsU.Close
                    mid = "0001"
                End If
        
            
                Call clearTextboxes
                txtDate.Text = Today
                txtID.Text = Format$(Val(mid), "000#")
                cmdReceive.Enabled = True
            Else
                Call clearTextboxes
                txtName.SetFocus
            End If
    Else
    
    If MsgBox("Do you want add new Employee with ID #" & txtID.Text + " ?", vbInformation + vbYesNo, "Add New") = vbYes Then
        mid = txtID.Text
        Call clearTextboxes
        txtID.Text = Format$(Val(mid), "000#")
        txtName.SetFocus
        txtDate.Text = Today
        cmdReceive.Enabled = True
        Else
        Call clearTextboxes
        txtName.SetFocus
    End If
    End If
    End If
    Exit Sub
End Sub

Private Sub txtMobile_Bill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSalary.SelStart = 0
    txtSalary.SelLength = Len(txtSalary.Text)
    txtSalary.SetFocus
End If
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDeposit.SetFocus
End If
End Sub

Private Sub txtMother_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPresent.SetFocus
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDOJ.SetFocus
End If
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtRate.Text = Format$(Val(txtRate.Text), "###0.00")
cmdReceive.SetFocus
End If
End Sub

Private Sub txtPermanent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtBirthday.SetFocus
End If
End Sub


Private Sub txtPresent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPermanent.SetFocus
End If
End Sub

Private Sub txtSalary_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbPercent.SelStart = 0
    cmbPercent.SelLength = Len(cmbPercent.Text)
    cmbPercent.SetFocus
End If

End Sub

Private Sub txtSalary_LostFocus()
txtSalary.Text = Format$(Val(txtSalary.Text), "###0.00")
End Sub

Private Sub txtTarget_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMobile_Bill.SelStart = 0
    txtMobile_Bill.SelLength = Len(txtMobile_Bill.Text)
    txtMobile_Bill.SetFocus
End If
End Sub
