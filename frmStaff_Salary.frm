VERSION 5.00
Begin VB.Form frmSalary 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Payment"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8190
   Icon            =   "frmStaff_Salary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5085
      Left            =   6060
      ScaleHeight     =   5055
      ScaleWidth      =   1965
      TabIndex        =   17
      Top             =   1020
      Width           =   1995
      Begin VB.TextBox txtTarget 
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
         Left            =   240
         TabIndex        =   51
         Text            =   "0.00"
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtAchive 
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
         Left            =   240
         TabIndex        =   50
         Text            =   "0.00"
         Top             =   3810
         Width           =   1455
      End
      Begin VB.TextBox txtAdvance_Paid 
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
         Left            =   240
         TabIndex        =   49
         Text            =   "0.00"
         Top             =   1740
         Width           =   1455
      End
      Begin VB.TextBox txtMobile_Due 
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
         Left            =   240
         TabIndex        =   48
         Text            =   "0.00"
         Top             =   4470
         Width           =   1425
      End
      Begin VB.TextBox txtSalary_Due 
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
         Left            =   240
         TabIndex        =   47
         Text            =   "0.00"
         Top             =   2430
         Width           =   1455
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
         Left            =   240
         TabIndex        =   45
         Text            =   "0.00"
         Top             =   1080
         Width           =   1455
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Text            =   "01-01-2012"
         Top             =   420
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         Height          =   4845
         Left            =   150
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label Label15 
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
         Left            =   240
         TabIndex        =   56
         Top             =   3510
         Width           =   1005
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Traget"
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
         Left            =   240
         TabIndex        =   55
         Top             =   2820
         Width           =   1290
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
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
         Left            =   240
         TabIndex        =   54
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile Siling"
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
         Left            =   240
         TabIndex        =   53
         Top             =   4200
         Width           =   1110
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Salary"
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
         Left            =   240
         TabIndex        =   52
         Top             =   2130
         Width           =   1305
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Due"
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
         Left            =   240
         TabIndex        =   46
         Top             =   810
         Width           =   900
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
         Left            =   240
         TabIndex        =   44
         Top             =   180
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   135
      ScaleHeight     =   705
      ScaleWidth      =   7890
      TabIndex        =   12
      Top             =   6210
      Width           =   7920
      Begin VB.CommandButton Command4 
         Caption         =   "Print"
         Height          =   435
         Left            =   4200
         TabIndex        =   16
         Top             =   150
         Width           =   1485
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   435
         Left            =   2280
         TabIndex        =   15
         Top             =   150
         Width           =   1485
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   435
         Left            =   345
         TabIndex        =   14
         Top             =   150
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   435
         Left            =   6045
         TabIndex        =   13
         Top             =   150
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5100
      Left            =   135
      ScaleHeight     =   5070
      ScaleWidth      =   5790
      TabIndex        =   2
      Top             =   1020
      Width           =   5820
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
         Left            =   4470
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   3930
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
         Left            =   4470
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   3480
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
         Left            =   4470
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   3030
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
         Left            =   4470
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2580
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
         Left            =   4470
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   2100
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
         Left            =   4470
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1620
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
         Left            =   4470
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1140
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
         Left            =   4470
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   690
         Width           =   1005
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
         Left            =   4470
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   210
         Width           =   1005
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
         Left            =   4500
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   4530
         Width           =   990
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
         Left            =   210
         TabIndex        =   21
         Top             =   3900
         Width           =   2715
      End
      Begin VB.ComboBox cmbType 
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
         ItemData        =   "frmStaff_Salary.frx":0442
         Left            =   210
         List            =   "frmStaff_Salary.frx":0452
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   3060
         Width           =   2055
      End
      Begin VB.TextBox txtDesi 
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
         Left            =   210
         TabIndex        =   19
         Top             =   2220
         Width           =   2085
      End
      Begin VB.TextBox txtTotal 
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
         Left            =   1440
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   4530
         Width           =   1455
      End
      Begin VB.TextBox txtMR_No 
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
         Left            =   1650
         TabIndex        =   5
         Top             =   570
         Width           =   1275
      End
      Begin VB.TextBox txtAccount 
         Alignment       =   2  'Center
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
         Left            =   210
         TabIndex        =   4
         Top             =   570
         Width           =   1350
      End
      Begin VB.TextBox txtName 
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
         Left            =   210
         TabIndex        =   3
         Top             =   1410
         Width           =   2715
      End
      Begin VB.Shape Shape4 
         Height          =   4215
         Left            =   120
         Top             =   150
         Width           =   2895
      End
      Begin VB.Shape Shape3 
         Height          =   555
         Left            =   120
         Top             =   4440
         Width           =   2895
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
         Left            =   3270
         TabIndex        =   42
         Top             =   3990
         Width           =   1200
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
         Left            =   3270
         TabIndex        =   41
         Top             =   3540
         Width           =   1170
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
         Left            =   3270
         TabIndex        =   40
         Top             =   3090
         Width           =   1170
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
         Left            =   3270
         TabIndex        =   39
         Top             =   2610
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
         Left            =   3270
         TabIndex        =   38
         Top             =   2160
         Width           =   1185
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
         Left            =   3270
         TabIndex        =   37
         Top             =   1680
         Width           =   1155
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
         Left            =   3270
         TabIndex        =   36
         Top             =   1200
         Width           =   1170
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
         Left            =   3270
         TabIndex        =   35
         Top             =   750
         Width           =   1170
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
         Left            =   3270
         TabIndex        =   34
         Top             =   270
         Width           =   1155
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
         Left            =   3270
         TabIndex        =   33
         Top             =   4590
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         Height          =   4245
         Left            =   3120
         Top             =   120
         Width           =   2505
      End
      Begin VB.Shape Shape2 
         Height          =   555
         Left            =   3120
         Top             =   4440
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
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
         Left            =   210
         TabIndex        =   22
         Top             =   2730
         Width           =   1260
      End
      Begin VB.Label Label26 
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
         Left            =   210
         TabIndex        =   18
         Top             =   1905
         Width           =   1005
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
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
         TabIndex        =   11
         Top             =   4590
         Width           =   1110
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   210
         TabIndex        =   10
         Top             =   3540
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MR No."
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
         Left            =   1680
         TabIndex        =   9
         Top             =   210
         Width           =   660
      End
      Begin VB.Label Label2 
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
         Left            =   210
         TabIndex        =   8
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label3 
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
         Left            =   210
         TabIndex        =   7
         Top             =   1065
         Width           =   1440
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   135
      Picture         =   "frmStaff_Salary.frx":0480
      ScaleHeight     =   705
      ScaleWidth      =   7890
      TabIndex        =   0
      Top             =   180
      Width           =   7920
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STAFF PAYMENT"
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
         Left            =   2190
         TabIndex        =   1
         Top             =   30
         Width           =   4380
      End
      Begin VB.Image Image3 
         Height          =   690
         Left            =   0
         Picture         =   "frmStaff_Salary.frx":5D2E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7935
      End
   End
End
Attribute VB_Name = "frmSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Today As Date
Dim mn As String
Dim Mont As String
Dim yr As String
Dim ac As String
Dim inword As String
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
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

Private Sub clearTextboxes()
        txtAccount.Text = ""
        txtDate.Text = ""
        txtName.Text = ""
        txtDesi.Text = ""
        cmbType.Text = ""
        txtDescription.Text = ""
        txtSales_Due.Text = "0.00"
        txtAdvance_Paid.Text = "0.00"
        txtSalary_Due.Text = "0.00"
        txtMobile_Due.Text = "0.00"
        txtTarget.Text = "0.00"
        txtTotal.Text = "0.00"
        txtMR_No.Text = ""
        
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
Private Sub Cash_Cr()
Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!MR_No = mn
            rsN!Name = txtName.Text
            rsN!Description = txtName.Text
            rsN!Cr = Val(txtTotal.Text)
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
        rs!Cash_Cr = rs!Cash_Cr + Val(txtTotal.Text)
        rs!Cash_Close = rs!Cash_Close - Val(txtTotal.Text)
        rs.Update
        rs.Close
End Sub

Private Sub Command2_GotFocus()
On Error Resume Next
Call Total_Cash
If Val(txtTotal.Text) <> Val(txtCash.Text) Then
MsgBox "Denomination Amount differs from Net Amount" & Val(txtTotal.Text) - Val(txtCash.Text), vbCritical, "Payment Info!"
Command2.Enabled = True
Command2.SetFocus
Else
Command2.Enabled = True
Command2.SetFocus
End If
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
Private Sub M_onth()
        Mont = MonthName(Month(CDate(Today)))
        Mont = UCase(Mont)
        yr = Year(CDate(Today))
End Sub
Private Sub cmbType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription.Text)
    txtDescription.SetFocus
End If
End Sub

Private Sub cmbType_LostFocus()
If cmbType.Text = "Salary" Then
On Error Resume Next
 Call M_onth
    txtDescription.Text = "Salary M/O " & Mont & "-" & yr
End If
If cmbType.Text = "Commission" Then
On Error Resume Next
 Call M_onth
    txtDescription.Text = "Commission M/O " & Mont & "-" & yr
End If
If cmbType.Text = "Mobile Bill" Then
On Error Resume Next
 Call M_onth
    txtDescription.Text = "Mobile Bill M/O " & Mont & "-" & yr
End If
If cmbType.Text = "Advance" Then
On Error Resume Next
 Call M_onth
    txtDescription.Text = "Advance M/O " & Mont & "-" & yr
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Set rs = New ADODB.Recordset
        str = "select * from Staff_Master where Staff_ID like '" & txtAccount.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
    
If Not rs.EOF Then
    If cmbType.Text = "Salary" Then
        rs!Salary = rs!Salary + Val(txtTotal.Text)
        rs!Balance = (rs!Salary + rs!Commission + rs!Mobile) - (rs!Advance + rs!Sales_Due + rs!Draw)
        rs.Update

        Set rsN = New ADODB.Recordset
            rsN.Open "Staff_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!D_ate = txtDate.Text
            rsN!Staff_Id = txtAccount.Text
            rsN!Name = txtName.Text
            rsN!Designation = txtDesi.Text
            rsN!Description = txtDescription.Text
            rsN!Dr = txtTotal.Text
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN!Center_code = Branch_Code
            rsN!Center_name = Branch_Name
            rsN.Update
            rsN.Close
           
        rs!Draw = rs!Draw + Val(txtTotal.Text)
        rs!Balance = (rs!Salary + rs!Commission + rs!Mobile) - (rs!Advance + rs!Sales_Due + rs!Draw)
        rs.Update

        Set rsN = New ADODB.Recordset
            rsN.Open "Staff_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!D_ate = txtDate.Text
            rsN!Staff_Id = txtAccount.Text
            rsN!Name = txtName.Text
            rsN!Designation = txtDesi.Text
            rsN!Description = "Salary Draw"
            rsN!Dr = 0
            rsN!Cr = txtTotal.Text
            rsN!Balance = rs!Balance
            rsN!Center_code = Branch_Code
            rsN!Center_name = Branch_Name
            rsN.Update
            rsN.Close
            rs.Close

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100120 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtTotal.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

        Set rsU = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Balance = rsU!Balance - Val(txtTotal.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = rsU!AC_No
                rsN!Name = rsU!Head_Name
                rsN!Description = "A/C:" & txtAccount.Text
                rsN!Dr = 0
                rsN!Cr = Val(txtTotal.Text)
                rsN!Balance = rsU!Balance
                rsN.Update
                mn = rsN!sl
                rsN.Close
                    
                Call Cash_Cr
            
    End If

    If cmbType.Text = "Commission" Then
        rs!Commission = rs!Commission + Val(txtTotal.Text)
        rs!Balance = (rs!Salary + rs!Commission + rs!Mobile) - (rs!Advance + rs!Sales_Due + rs!Draw)
        rs.Update

        Set rsN = New ADODB.Recordset
            rsN.Open "Staff_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!D_ate = txtDate.Text
            rsN!Staff_Id = txtAccount.Text
            rsN!Name = txtName.Text
            rsN!Designation = txtDesi.Text
            rsN!Description = txtDescription.Text
            rsN!Dr = txtTotal.Text
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN!Center_code = Branch_Code
            rsN!Center_name = Branch_Name
            rsN.Update
            rsN.Close
           
        rs!Draw = rs!Draw + Val(txtTotal.Text)
        rs!Balance = (rs!Salary + rs!Commission + rs!Mobile) - (rs!Advance + rs!Sales_Due + rs!Draw)
        rs.Update

        Set rsN = New ADODB.Recordset
            rsN.Open "Staff_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!D_ate = txtDate.Text
            rsN!Staff_Id = txtAccount.Text
            rsN!Name = txtName.Text
            rsN!Designation = txtDesi.Text
            rsN!Description = "Commission Draw"
            rsN!Dr = 0
            rsN!Cr = txtTotal.Text
            rsN!Balance = rs!Balance
            rsN!Center_code = Branch_Code
            rsN!Center_name = Branch_Name
            rsN.Update
            rsN.Close
            rs.Close

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtTotal.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

        Set rsU = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Balance = rsU!Balance - Val(txtTotal.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = rsU!AC_No
                rsN!Name = rsU!Head_Name
                rsN!Description = "A/C:" & txtAccount.Text
                rsN!Dr = 0
                rsN!Cr = Val(txtTotal.Text)
                rsN!Balance = rsU!Balance
                rsN.Update
                mn = rsN!sl
                rsN.Close
                    
                Call Cash_Cr
            
    End If

    If cmbType.Text = "Mobile Bill" Then
        rs!Mobile = rs!Mobile + Val(txtTotal.Text)
        rs!Balance = (rs!Salary + rs!Commission + rs!Mobile) - (rs!Advance + rs!Sales_Due + rs!Draw)
        rs.Update

        Set rsN = New ADODB.Recordset
            rsN.Open "Staff_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!D_ate = txtDate.Text
            rsN!Staff_Id = txtAccount.Text
            rsN!Name = txtName.Text
            rsN!Designation = txtDesi.Text
            rsN!Description = txtDescription.Text
            rsN!Dr = txtTotal.Text
            rsN!Cr = 0
            rsN!Balance = rs!Balance
            rsN!Center_code = Branch_Code
            rsN!Center_name = Branch_Name
            rsN.Update
            rsN.Close
           
        rs!Draw = rs!Draw + Val(txtTotal.Text)
        rs!Balance = (rs!Salary + rs!Commission + rs!Mobile) - (rs!Advance + rs!Sales_Due + rs!Draw)
        rs.Update

        Set rsN = New ADODB.Recordset
            rsN.Open "Staff_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!D_ate = txtDate.Text
            rsN!Staff_Id = txtAccount.Text
            rsN!Name = txtName.Text
            rsN!Designation = txtDesi.Text
            rsN!Description = "Mobile Bill Draw"
            rsN!Dr = 0
            rsN!Cr = txtTotal.Text
            rsN!Balance = rs!Balance
            rsN!Center_code = Branch_Code
            rsN!Center_name = Branch_Name
            rsN.Update
            rsN.Close
            rs.Close

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100124 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtTotal.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

        Set rsU = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Balance = rsU!Balance - Val(txtTotal.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = rsU!AC_No
                rsN!Name = rsU!Head_Name
                rsN!Description = "A/C:" & txtAccount.Text
                rsN!Dr = 0
                rsN!Cr = Val(txtTotal.Text)
                rsN!Balance = rsU!Balance
                rsN.Update
                mn = rsN!sl
                rsN.Close
                    
                Call Cash_Cr
    End If
    
    If cmbType.Text = "Advance" Then
        rs!Advance = rs!Advance + Val(txtTotal.Text)
        rs!Balance = (rs!Salary + rs!Commission + rs!Mobile) - (rs!Advance + rs!Sales_Due + rs!Draw)
        rs.Update

        Set rsN = New ADODB.Recordset
            rsN.Open "Staff_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!D_ate = txtDate.Text
            rsN!Staff_Id = txtAccount.Text
            rsN!Name = txtName.Text
            rsN!Designation = txtDesi.Text
            rsN!Description = txtDescription.Text
            rsN!Dr = 0
            rsN!Cr = txtTotal.Text
            rsN!Balance = rs!Balance
            rsN!Center_code = Branch_Code
            rsN!Center_name = Branch_Name
            rsN.Update
            rsN.Close
        
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100110 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtTotal.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

        Set rsU = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100100 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Balance = rsU!Balance - Val(txtTotal.Text)
            rsU!Date = txtDate.Text
            rsU.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
                rsN!Date = txtDate.Text
                rsN!AC_No = rsU!AC_No
                rsN!Name = rsU!Head_Name
                rsN!Description = "A/C:" & txtAccount.Text
                rsN!Dr = 0
                rsN!Cr = Val(txtTotal.Text)
                rsN!Balance = rsU!Balance
                rsN.Update
                mn = rsN!sl
                rsN.Close
                    
                Call Cash_Cr
    End If
    
    Else
'        rs!Commission = rs!Commission + Val(txtCommission.Text)
'        rs!Mobile = rs!Mobile + Val(txtMobile.Text)
'        rs!Advance = rs!Mobile + Val(txtMobile.Text)
        
        MsgBox "no record found"
        rs.Close
        Exit Sub
    End If

   
    Call clearTextboxes
    txtDate.Text = Today
    txtAccount.SetFocus
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Exit Sub

End Sub
Private Sub Form_Load()
    Call clearTextboxes
    txtDate.Text = Today
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
    Exit Sub
End Sub
Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbType.SelStart = 0
    cmbType.SelLength = Len(cmbType.Text)
    cmbType.SetFocus
End If
End Sub
Private Sub txtAccount_LostFocus()
If txtAccount.Text = "" Then
Exit Sub
End If
Dim Search_Id As String
Dim Com As Integer
Search_Id = txtAccount.Text
Com = 0
Set rs = New ADODB.Recordset
        str = "select * from Staff_Master where Staff_Id like '" & Search_Id & "' order by Staff_id"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
    Call clearTextboxes
    Call M_onth
        txtAccount.Text = rs!Staff_Id
        txtDate.Text = Today
        txtName.Text = rs!Name
        txtDesi.Text = rs!Designation
        'txtDescription.Text = "Salary M/O " & Mont & "-" & yr
        txtTarget.Text = Format$(Val(rs!Target), "###0.00")
        txtSales_Due.Text = Format$(Val(rs!Sales_Due), "###0.00")
        txtAdvance_Paid.Text = Format$(Val(rs!Advance), "###0.00")
        txtSalary_Due.Text = Format$(Val(rs!Salary_Due), "###0.00")
        txtMobile_Due.Text = Format$(Val(rs!Mobile_Due), "###0.00")
        rs.Close
    
'    Set rs = New ADODB.Recordset
'        str = "select * from Customer_Master where Customer like '" & Search_Id & "'"
'        rs.Open str, conn
'
'        If Not rs.EOF Then
'            rs.MoveFirst
'
'        Do While Not rs.EOF
'            Com = Com + Val(rs!Sale_Amount)
'            rs.MoveNext
'        Loop
'            rs.Close
'            txtAchive.Text = Format$(Val(Com), "###0.00")
'        Else
'        rs.Close
'        End If

        Command2.Enabled = True
    Else
    MsgBox "There is no such Staff Account no. found.,", vbCritical
        rs.Close
    End If
    Exit Sub
End Sub
Private Sub txtCash_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command2.SetFocus
End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTotal.SelStart = 0
    txtTotal.SelLength = Len(txtTotal.Text)
    txtTotal.SetFocus
End If
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt1000.SelStart = 0
    txt1000.SelLength = Len(txt1000.Text)
    txt1000.SetFocus
    Command2.Enabled = True
End If
End Sub
