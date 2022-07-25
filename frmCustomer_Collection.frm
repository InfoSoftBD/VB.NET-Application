VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmCustomer_Collection 
   BackColor       =   &H00004080&
   Caption         =   "Customer Due Collection"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   Icon            =   "frmCustomer_Collection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
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
      Height          =   4800
      Left            =   120
      TabIndex        =   24
      Top             =   2970
      Width           =   7500
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
         Left            =   2025
         TabIndex        =   66
         Text            =   "Text7"
         Top             =   1553
         Width           =   1560
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   180
         TabIndex        =   64
         Text            =   "Text2"
         Top             =   4275
         Width           =   7125
      End
      Begin VB.TextBox txtBank_Balance 
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
         Left            =   5760
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   2190
         Width           =   1560
      End
      Begin VB.TextBox txtCash_Balance 
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
         Left            =   5730
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   240
         Width           =   1560
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
         Left            =   6480
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   1020
         Width           =   675
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
         Left            =   5715
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   1020
         Width           =   675
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
         Left            =   4950
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   1020
         Width           =   675
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
         Left            =   4185
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   1020
         Width           =   675
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
         Left            =   3450
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   1020
         Width           =   675
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
         Left            =   2685
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   1020
         Width           =   675
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
         Left            =   1920
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1020
         Width           =   675
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
         Left            =   1155
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   1020
         Width           =   675
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
         Left            =   330
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1020
         Width           =   735
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
         Left            =   5550
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2955
         Width           =   1560
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
         Left            =   5730
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1553
         Width           =   1440
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
         Left            =   2430
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   2955
         Width           =   1545
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
         Left            =   4035
         TabIndex        =   28
         Text            =   "Combo2"
         Top             =   2955
         Width           =   1380
      End
      Begin VB.TextBox txtChq_Amnt 
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
         Left            =   5550
         TabIndex        =   27
         Text            =   "Text7"
         Top             =   3450
         Width           =   1560
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
         Left            =   270
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   2970
         Width           =   2100
      End
      Begin VB.TextBox txtAdjust 
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
         Left            =   1845
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   3465
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Amount"
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
         Left            =   360
         TabIndex        =   67
         Top             =   1620
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Description"
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
         TabIndex        =   65
         Top             =   4005
         Width           =   1815
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available Bank Balance"
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
         Left            =   3420
         TabIndex        =   63
         Top             =   2250
         Width           =   2280
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available Cash Balance"
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
         Left            =   3390
         TabIndex        =   62
         Top             =   300
         Width           =   2265
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bank Amount"
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
         Left            =   3630
         TabIndex        =   61
         Top             =   3510
         Width           =   1815
      End
      Begin VB.Shape Shape3 
         Height          =   1365
         Left            =   180
         Top             =   2610
         Width           =   7125
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   210
         TabIndex        =   60
         Top             =   2250
         Width           =   1380
      End
      Begin VB.Shape Shape2 
         Height          =   1455
         Left            =   180
         Top             =   660
         Width           =   7125
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   150
         TabIndex        =   59
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label lblPmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total Tk. 0.00"
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
         Left            =   5400
         TabIndex        =   58
         Top             =   4020
         Width           =   1905
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 1"
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
         Left            =   6480
         TabIndex        =   57
         Top             =   750
         Width           =   435
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 2"
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
         Left            =   5715
         TabIndex        =   56
         Top             =   750
         Width           =   435
      End
      Begin VB.Label lbl5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 5"
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
         Left            =   4950
         TabIndex        =   55
         Top             =   750
         Width           =   435
      End
      Begin VB.Label lbl10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 10"
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
         Left            =   4185
         TabIndex        =   54
         Top             =   750
         Width           =   540
      End
      Begin VB.Label lbl20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 20"
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
         Left            =   3450
         TabIndex        =   53
         Top             =   750
         Width           =   540
      End
      Begin VB.Label lbl50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 50"
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
         TabIndex        =   52
         Top             =   750
         Width           =   540
      End
      Begin VB.Label lbl100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 100"
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
         Left            =   1920
         TabIndex        =   51
         Top             =   750
         Width           =   645
      End
      Begin VB.Label lbl500 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 500"
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
         Left            =   1155
         TabIndex        =   50
         Top             =   750
         Width           =   645
      End
      Begin VB.Label lbl1000 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tk. 1000"
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
         TabIndex        =   49
         Top             =   750
         Width           =   750
      End
      Begin VB.Label lblAccount 
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
         Left            =   300
         TabIndex        =   48
         Top             =   2685
         Width           =   1080
      End
      Begin VB.Label lblCash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Amount"
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
         Left            =   3840
         TabIndex        =   47
         Top             =   1620
         Width           =   1800
      End
      Begin VB.Label lblChq_No 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque No."
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
         Left            =   5535
         TabIndex        =   46
         Top             =   2685
         Width           =   1020
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
         Left            =   4065
         TabIndex        =   45
         Top             =   2685
         Width           =   615
      End
      Begin VB.Label lblBank 
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
         Left            =   2490
         TabIndex        =   44
         Top             =   2685
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment "
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
         Left            =   315
         TabIndex        =   43
         Top             =   3510
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   900
      Width           =   7485
      Begin VB.TextBox txtDue 
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
         Left            =   6030
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   600
         Width           =   1275
      End
      Begin VB.ComboBox cmbCustomer_Name 
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
         Left            =   2970
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   600
         Width           =   2940
      End
      Begin VB.TextBox txtCustomer_Address 
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
         Left            =   1530
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   1380
         Width           =   5775
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
         Left            =   135
         TabIndex        =   14
         Text            =   "Text4"
         Top             =   600
         Width           =   1275
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
         Left            =   120
         TabIndex        =   13
         Text            =   "Text3"
         Top             =   1380
         Width           =   1260
      End
      Begin VB.ComboBox txtCustomer_Code 
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
         Left            =   1530
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   607
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Present Due"
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
         Left            =   6030
         TabIndex        =   23
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label Label13 
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
         Left            =   1530
         TabIndex        =   22
         Top             =   1050
         Width           =   1620
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         Left            =   3000
         TabIndex        =   21
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
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
         Left            =   1560
         TabIndex        =   20
         Top             =   270
         Width           =   1350
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
         Left            =   180
         TabIndex        =   19
         Top             =   270
         Width           =   840
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
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   90
      ScaleHeight     =   660
      ScaleWidth      =   7500
      TabIndex        =   6
      Top             =   7860
      Width           =   7530
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Save"
         Height          =   435
         Left            =   270
         TabIndex        =   10
         Top             =   120
         Width           =   1380
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   435
         Left            =   2145
         TabIndex        =   9
         Top             =   120
         Width           =   1320
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   435
         Left            =   5805
         TabIndex        =   8
         Top             =   120
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   435
         Left            =   3975
         TabIndex        =   7
         Top             =   120
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   7455
      TabIndex        =   4
      Top             =   135
      Width           =   7485
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER DUE COLLECTION"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   375
         TabIndex        =   5
         Top             =   60
         Width           =   6615
      End
      Begin VB.Image Image3 
         Height          =   1140
         Left            =   0
         Picture         =   "frmCustomer_Collection.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7470
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      Height          =   8385
      Left            =   7740
      ScaleHeight     =   8325
      ScaleWidth      =   4185
      TabIndex        =   0
      Top             =   135
      Width           =   4245
      Begin VB.TextBox txtSearch 
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
         Bindings        =   "frmCustomer_Collection.frx":A4D0
         Height          =   7485
         Left            =   150
         TabIndex        =   2
         Top             =   735
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   13203
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
      Left            =   7695
      Top             =   7740
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
Attribute VB_Name = "frmCustomer_Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Sub Total_Cash()
    Total = (Val(txt1000.Text) * 1000) + (Val(txt500.Text) * 500) + (Val(txt100.Text) * 100) + (Val(txt50.Text) * 50) + (Val(txt20.Text * 20)) + (Val(txt10.Text) * 10) + (Val(txt5.Text) * 5) + (Val(txt2.Text) * 2) + (Val(txt1.Text) * 1) + Val(txtChq_Amnt.Text)
    txtCash.Text = Format(Total - Val(txtChq_Amnt.Text), "###0.00")
    lblPmt.Caption = "Total Paid Amount Tk. " + Format$(Val(Total), "###0.00")
End Sub

Private Sub Command1_Click()
On Error Resume Next
If MsgBox("Really want to Print?", vbCritical + vbYesNo) = vbYes Then
    
    
    Set rsU = New ADODB.Recordset
        str = "select * from Money_Receipt where Sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
        rsU.Close
        str = "delete from Money_Receipt where Sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

 Set rs = New ADODB.Recordset
        str = "select * from Due_Tran where Invoice like '" & txtInvoice.Text & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        rs.Close
    
    If Val(txtCash.Text) >= 1 And Val(txtCash.Text) < 10 Then
        Value = Val(txtCash.Text)
        Call Case0_9
    End If
    If Val(txtCash.Text) >= 10 And Val(txtCash.Text) < 100 Then
        Value = Val(txtCash.Text)
        Call Case10_99
    End If
    If Val(txtCash.Text) >= 100 And Val(txtCash.Text) < 1000 Then
        Value = Val(txtCash.Text)
        Call Case100_999
    End If
    If Val(txtCash.Text) >= 1000 And Val(txtCash.Text) < 100000 Then
        Value = Val(txtCash.Text)
        Call Case1000_99999
    End If
    If Val(txtCash.Text) >= 100000 And Val(txtCash.Text) < 10000000 Then
        Value = Val(txtCash.Text)
        Call Case100000_9999999
    End If
    If Val(txtCash.Text) >= 10000000 Then
        Value = Val(txtCash.Text)
        Call Case10000000_999999999
    End If
        
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Money_Receipt", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!sl = 1
        rsN!MR_No = txtInvoice.Text
        rsN!Date = txtDate.Text
        rsN!AC_No = txtCustomer_Code.Text
        rsN!Name = cmbCustomer_Name.Text
        rsN!Description = txtDescription.Text
        If txtCash.Text > 0 Then
        rsN!Cr = txtCash.Text
        rsN!Total = txtCash.Text
        End If
        If txtAdjust.Text > 0 Then
        rsN!Cr = txtAdjust.Text
        rsN!Total = txtAdjust.Text
        End If
        
        rsN!inword = inword
        rsN.Update
        rsN.Close
        
    
    
    Call clearTextboxes
    
    Command1.Enabled = False
    
    txtDate.Text = Today
    
    
Set rs = New ADODB.Recordset
    str = "select * from Money_Receipt"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.RecordCount > 0 Then
    rs.Update
    rs.Close

    str = "select * from Money_Receipt"
    rptMoney_Receipt.rsMoney.ConnectionString = cnStr
    rptMoney_Receipt.rsMoney.Source = str
    
    rptMoney_Receipt.Show 1
    End If
Else

Exit Sub
End If
Else

Exit Sub
End If

Resume Next

End Sub
Private Sub Invo_No()
Dim ID As String
Dim mid As Integer
    ID = txtInvoice.Text
    mid = 0
    
    Set rs = New ADODB.Recordset
        str = "select * from Due_Tran where Invoice like '" & ID & "'"
        rs.Open str, conn

    If Not rs.EOF Then
        rs.Close
    Else
        
        Set rsU = New ADODB.Recordset
            str = "select * from Due_Tran order by Invoice"
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
            
            txtInvoice.Text = Format(mid, "000#")
            txtDate.Text = Today
'            txtType.Text = ""
'            txtName.Text = ""
'            txtAddress.Text = ""
'            txtMobile.Text = ""
            'txtID.SetFocus
        rs.Close
    ' cmdSales.Enabled = True
End If
End Sub

Private Sub Form_Activate()
txtInvoice.SetFocus
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

Private Sub Cash_Dr()
    
    Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtCustomer_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
    
        rs!Cr = rs!Cr + Val(txtCash.Text)
        rs!Balance = rs!Balance - Val(txtCash.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!INVOICE = txtInvoice.Text
        rsN!Customer_Code = txtCustomer_Code.Text
        rsN!Customer_Name = cmbCustomer_Name.Text
        rsN!Customer_Address = txtCustomer_Address.Text
        rsN!Description = "Cash Payment " & txtDescription.Text
        rsN!Cr = Val(txtCash.Text)
        rsN!Dr = 0
        rsN!Balance = rs!Balance
        
        rsN.Update
        rsN.Close
        'rs.Close
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Due_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!INVOICE = txtInvoice.Text
        rsN!Customer_Code = txtCustomer_Code.Text
        rsN!Customer_Name = cmbCustomer_Name.Text
        rsN!Customer_Address = txtCustomer_Address.Text
        rsN!Description = "Cash Payment " & txtDescription.Text
        rsN!Cr = Val(txtCash.Text)
        rsN!Dr = 0
        rsN!Balance = rs!Balance
        
        rsN.Update
        rsN.Close
        rs.Close

Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtCash.Text)
        rs.Update
        
     Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "CID-" & txtCustomer_Code.Text
        rsN!Cr = Val(txtCash.Text)
        rsN!Dr = 0
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
        rsN!Description = "CID-" & txtCustomer_Code.Text
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
        rsN!Name = cmbCustomer_Name.Text
        rsN!Description = cmbCustomer_Name.Text
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
            rsU!Balance = rsU!Balance - Val(txt1000.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt1000.Text) * 1000)
            rsU!Cash_Close = rsU!Cash_Close - (Val(txt1000.Text) * 1000)
            rsU.Update
            rsU.Close
            
                
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 500 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt500.Text)
            rsU!Balance = rsU!Balance - Val(txt500.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt500.Text) * 500)
            rsU!Cash_Close = rsU!Cash_Close - (Val(txt500.Text) * 500)
            rsU.Update
            rsU.Close
            
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 100 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt100.Text)
            rsU!Balance = rsU!Balance - Val(txt100.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt100.Text) * 100)
            rsU!Cash_Close = rsU!Cash_Close - (Val(txt100.Text) * 100)
            rsU.Update
            rsU.Close
            
            
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 50 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt50.Text)
            rsU!Balance = rsU!Balance - Val(txt50.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt50.Text) * 50)
            rsU!Cash_Close = rsU!Cash_Close - (Val(txt50.Text) * 50)
            rsU.Update
            rsU.Close
        
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 20 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt20.Text)
            rsU!Balance = rsU!Balance - Val(txt20.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt20.Text) * 20)
            rsU!Cash_Close = rsU!Cash_Close - (Val(txt20.Text) * 20)
            rsU.Update
            rsU.Close
        
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 10 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt10.Text)
            rsU!Balance = rsU!Balance - Val(txt10.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt10.Text) * 10)
            rsU!Cash_Close = rsU!Cash_Close - (Val(txt10.Text) * 10)
            rsU.Update
            rsU.Close
        
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 5 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt5.Text)
            rsU!Balance = rsU!Balance - Val(txt5.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt5.Text) * 5)
            rsU!Cash_Close = rsU!Cash_Close - (Val(txt5.Text) * 5)
            rsU.Update
            rsU.Close
            
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 2 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt2.Text)
            rsU!Balance = rsU!Balance - Val(txt2.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt2.Text) * 2)
            rsU!Cash_Close = rsU!Cash_Close - (Val(txt2.Text) * 2)
            rsU.Update
            rsU.Close
        
        Set rsU = New ADODB.Recordset
            str = "select * from Cash_Master where Code like '" & 1 & "'"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
            rsU!Dr = rsU!Dr + Val(txt1.Text)
            rsU!Balance = rsU!Balance - Val(txt1.Text)
            rsU!Cash_Dr = rsU!Cash_Dr + (Val(txt1.Text) * 1)
            rsU!Cash_Close = rsU!Cash_Close - (Val(txt1.Text) * 1)
            rsU.Update
            rsU.Close
        
    
        Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Cash_Dr = rsU!Cash_Dr + Val(txtCash.Text)
        rsU!Cash_Close = rsU!Cash_Close + Val(txtCash.Text)
        rsU.Update
        rsU.Close
    
End Sub
Private Sub Charge_Add()
On Error Resume Next
Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtCustomer_Code.Text & "'"
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
            rsN!Customer_Code = txtCustomer_Code.Text
            rsN!Customer_Name = cmbCustomer_Name.Text
            rsN!Customer_Address = txtCustomer_Address.Text
'            rsN!Customer_Type = txtType.Text
'            rsN!Mobile = txtMobile.Text
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
            rsN!Customer_Code = txtCustomer_Code.Text
            rsN!Customer_Name = cmbCustomer_Name.Text
            rsN!Customer_Address = txtCustomer_Address.Text
'            rsN!Customer_Type = txtType.Text
'            rsN!Mobile = txtMobile.Text
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
            rsN!Customer_Code = txtCustomer_Code.Text
            rsN!Customer_Name = cmbCustomer_Name.Text
            rsN!Customer_Address = txtCustomer_Address.Text
'            rsN!Customer_Type = txtType.Text
'            rsN!Mobile = txtMobile.Text
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

Private Sub Adjust_Cr()
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100102 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtAdjust.Text)
        rs!Date = txtDate.Text
        rs.Update
            
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "CID-" & txtCustomer_Code.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtAdjust.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        
      
       
        
     Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtAdjust.Text)
        rs.Update
        
     Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "CID-" & txtCustomer_Code.Text
        rsN!Dr = Val(txtAdjust.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close

        rs.Close
        
    Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtCustomer_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Dr = rs!Dr + Val(txtAdjust.Text)
        rs!Balance = rs!Balance + Val(txtAdjust.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!INVOICE = txtInvoice.Text
        rsN!Customer_Code = txtCustomer_Code.Text
        rsN!Customer_Name = cmbCustomer_Name.Text
        rsN!Customer_Address = txtCustomer_Address.Text
        rsN!Description = "Due Adjustment " & txtDescription.Text
        rsN!Dr = Val(txtAdjust.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        
        rsN.Update
        rsN.Close
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Due_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!INVOICE = txtInvoice.Text
        rsN!Customer_Code = txtCustomer_Code.Text
        rsN!Customer_Name = cmbCustomer_Name.Text
        rsN!Customer_Address = txtCustomer_Address.Text
        rsN!Description = "Due Adjustment " & txtDescription.Text
        rsN!Dr = Val(txtAdjust.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        
        rsN.Update
        rsN.Close
        rs.Close
    End If

    
End Sub

Private Sub Bank_Dr()
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
        rsN!Description = "CID-" & txtCustomer_Code.Text
        rsN!Cr = 0
        rsN!Dr = Val(txtChq_Amnt.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
        
    Set rs = New ADODB.Recordset
        str = "select * from Bank_Master where AC_No like '" & txtAccount.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Deposit = rs!Deposit + Val(txtChq_Amnt.Text)
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
        rsN!MR_No = txtInvoice.Text
        rsN!Chq_No = "CN" & txtChq_No.Text
        rsN!Description = cmbCustomer_Name.Text
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
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs.Update
        
     Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "CID-" & txtCustomer_Code.Text
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Dr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close

        rs.Close

Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtCustomer_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Cr = rs!Cr + Val(txtChq_Amnt.Text)
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Customer_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!INVOICE = txtInvoice.Text
        rsN!Customer_Code = txtCustomer_Code.Text
        rsN!Customer_Name = cmbCustomer_Name.Text
        rsN!Customer_Address = txtCustomer_Address.Text
        rsN!Description = "CN" & txtChq_No.Text & txtDescription.Text
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Dr = 0
        rsN!Balance = rs!Balance
        
        rsN.Update
        rsN.Close
        
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Due_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Date = txtDate.Text
        rsN!INVOICE = txtInvoice.Text
        rsN!Customer_Code = txtCustomer_Code.Text
        rsN!Customer_Name = cmbCustomer_Name.Text
        rsN!Customer_Address = txtCustomer_Address.Text
        rsN!Description = "Bank Payment " & txtDescription.Text
        rsN!Cr = Val(txtChq_Amnt.Text)
        rsN!Dr = 0
        rsN!Balance = rs!Balance
        
        rsN.Update
        rsN.Close
        rs.Close
    End If
End Sub
Private Sub Customer_Name()
        cmbCustomer_Name.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Customer_Name FROM Customer_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbCustomer_Name.AddItem rsN!Customer_Name
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub Customer_Code()
        txtCustomer_Code.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Customer_Code FROM Customer_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        txtCustomer_Code.AddItem rsN!Customer_Code
        rsN.MoveNext
        Loop
        rsN.Close
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
Private Sub Account_Name()
        txtAccount.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT AC_No FROM Bank_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        txtAccount.AddItem rsN!AC_No
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub clearTextboxes()
        txtInvoice.Text = ""
        txtCustomer_Code.Text = ""
        cmbCustomer_Name.Text = ""
        txtCustomer_Address.Text = ""
        txtDue.Text = ""
        txtCash.Text = "0.00"
        txtAdjust.Text = "0.00"
        txtCharge.Text = "0.00"
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
        txtCash_Balance.Text = "0.00"
        txtBank_Balance.Text = "0.00"
        txtDescription.Text = ""
        
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
            txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtSale_Price.Text = Format$(Val(rsN!Sale_Price), "###0.00")
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
            txtDate.Text = Date
            txtSale_Price.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            rsN.Close
        Else
        Exit Sub
        End If

End Sub

Private Sub cmbCustomer_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCash.SelStart = 0
    txtCash.SelLength = Len(txtCash.Text)
    txtCash.SetFocus
End If
End Sub

Private Sub cmbCustomer_Name_LostFocus()
If cmbCustomer_Name.Text = "" Then
Exit Sub
End If

Dim Customer As String
Dim mid As Integer
Customer = cmbCustomer_Name.Text
        Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Customer_Master where Customer_Name Like '" & Customer & "'order by Customer_code"
        rsN.Open str, conn
        
If Not rsN.EOF Then
            txtCustomer_Code.Text = rsN!Customer_Code
            cmbCustomer_Name.Text = rsN!Customer_Name
            txtCustomer_Address.Text = rsN!Customer_Address
            txtDue.Text = Format(Val(rsN!Balance), "###0.00")
            rsN.Close
        
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn
       
        If Not rs.EOF Then
        txtCash_Balance.Text = Format$(Val(rs!Balance), "###0.00")
        rs.Close
        End If

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100101 & "'"
        rs.Open str, conn
       
        If Not rs.EOF Then
        txtBank_Balance.Text = Format$(Val(rs!Balance), "###0.00")
        rs.Close
        End If
        
Else
        
    MsgBox "Invalid Customer Code", vbDritical, "Error!"
                
        rsN.Close
        Exit Sub
End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdReceive_Click()
If txtInvoice.Text = "" Then
MsgBox "Please Input Invoice No!", vbDritical
txtInvoice.SetFocus
Exit Sub
End If
    
    If Val(txtCash.Text) > 0 Then
        Call Cash_Dr
    End If
    If Val(txtChq_Amnt.Text) > 0 Then
        Call Bank_Dr
    End If
    If Val(txtAdjust.Text) > 0 Then
        Call Adjust_Cr
    End If
    If Val(txtCharge.Text) > 0 Then
        Call Charge_Add
    End If

    
    Call clearTextboxes
    txtDate.Text = Today
    cmdReceive.Enabled = False
    
    Call Invo_No
    
    Set rs = New ADODB.Recordset
        str = "select * from Due_Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by Id Desc"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
        Call ColumnWidth
        txtInvoice.SetFocus
Exit Sub
   Resume Next
End Sub

Private Sub Form_Load()
On Error Resume Next
   
   
    Call Customer_Name
    Call Customer_Code
    Call Account_Name
    Call Bank_Name
    Call Branch_Name
    Call clearTextboxes
    txtDate.Text = Today
    Call Invo_No
     Set rs = New ADODB.Recordset
        str = "select * from Due_Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by Id Desc"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
        Call ColumnWidth
                
        cmdReceive.Enabled = False
        cmdUpdate.Enabled = False
        Command1.Enabled = False
    Exit Sub
End Sub


Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbBank.SetFocus
End If
End Sub

Private Sub txtAccount_LostFocus()
If txtAccount.Text = "" Then
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
        MsgBox "Invalid Account No!", vbDritical, "Error!"
        rsN.Close
        Exit Sub
        End If
        
End Sub

Private Sub txtAdjust_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAccount.SetFocus
End If
End Sub

Private Sub txtAdjust_LostFocus()
txtAdjust.Text = Format$(Val(txtAdjust.Text), "###0.00")
End Sub

Private Sub txtCash_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAccount.SetFocus
End If
End Sub

Private Sub txtCash_LostFocus()
'Set rs = New ADODB.Recordset
'        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
'        rs.Open str, conn
'
'        If rs!Balance < Val(txtCash.Text) Then
'        MsgBox "Insufficiant Balance.!", vbDritical, "Error!"
'        txtCash.SetFocus
'        Else
        txtCash.Text = Format$(Val(txtCash.Text), "###0.00")
'        End If
End Sub

Private Sub txtCharge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtCash.SelStart = 0
txtCash.SelLength = Len(txtCash.Text)
txtCash.SetFocus
End If
End Sub

Private Sub txtCharge_LostFocus()
txtAdjust.Text = Format$(Val(txtAdjustText), "###0.00")
txtCash.Text = Format$(Val(txtDue.Text) - Val(txtCharge.Text), "###0.00")
txtCash.SelStart = 0
txtCash.SelLength = Len(txtCash.Text)
txtCash.SetFocus
End Sub

Private Sub txtChq_Amnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDescription.SetFocus
End If
End Sub

Private Sub txtChq_Amnt_LostFocus()
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100101 & "'"
        rs.Open str, conn
        rs.Close
'        If rs!Balance < Val(txtChq_Amnt.Text) Then
'        MsgBox "Insufficiant Balance.!", vbDritical, "Error!"
'        txtChq_Amnt.SetFocus
'        Else
        txtChq_Amnt.Text = Format$(Val(txtChq_Amnt.Text), "###0.00")
       ' End If
End Sub

Private Sub txtChq_No_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtChq_Amnt.SetFocus
End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmdReceive.Enabled = True Then
cmdReceive.SetFocus
Else
Exit Sub
End If
End If
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtInvoice.Text = "" Then
Exit Sub
Else
cmdReceive.Enabled = True
cmbCustomer_Name.SetFocus
End If
End If
End Sub
Private Sub txtCustomer_Address_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt1000.SelStart = 0
    txt1000.SelLength = Len(txt500.Text)
    txt1000.SetFocus
End If
End Sub
Private Sub txtCustomer_Code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCash.SelStart = 0
    txtCash.SelLength = Len(txtCash.Text)
    txtCash.SetFocus
End If
End Sub

Private Sub txtCustomer_Code_LostFocus()
If txtCustomer_Code.Text = "" Then
Exit Sub
End If
Dim Customer As String
Dim mid As Integer
Customer = txtCustomer_Code.Text
        Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Customer_Master where Customer_Code Like '" & Customer & "'order by Customer_code"
        rsN.Open str, conn
        
    If Not rsN.EOF Then
            txtCustomer_Code.Text = rsN!Customer_Code
            cmbCustomer_Name.Text = rsN!Customer_Name
            txtCustomer_Address.Text = rsN!Customer_Address
            txtDue.Text = Format(Val(rsN!Balance), "###0.00")
            rsN.Close

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn
       
        If Not rs.EOF Then
        txtCash_Balance.Text = Format$(Val(rs!Balance), "###0.00")
        rs.Close
        End If

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100101 & "'"
        rs.Open str, conn
       
        If Not rs.EOF Then
        txtBank_Balance.Text = Format$(Val(rs!Balance), "###0.00")
        rs.Close
        End If
Else
        
    MsgBox "Invalid Customer Code", vbDritical, "Error!"
        rsN.Close
        Exit Sub
End If

End Sub


Private Sub txtInvoice_LostFocus()
If txtInvoice.Text = "" Then
Exit Sub
End If

On Error Resume Next
Dim Tran As String
    Tran = txtInvoice.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Due_Tran where Invoice like '" & Tran & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        txtDate.Text = rs!Date
        txtInvoice.Text = rs!INVOICE
        txtCustomer_Code.Text = rs!Customer_Code
        cmbCustomer_Name.Text = rs!Customer_Name
        txtCustomer_Address.Text = rs!Customer_Address
        txtDescription.Text = rs!Description
        txtCash.Text = Val(rs!Cr)
        txtAdjust.Text = Val(rs!Dr)
        txtDue.Text = rs!Balance + rs!Cr
        rs.Close
        txtCash.Text = Format$(Val(txtCash.Text), "###0.00")
        txtAdjust.Text = Format$(Val(txtAdjust.Text), "###0.00")
        Command1.Enabled = True
        cmdReceive.Enabled = False
    Else
    rs.Close
    cmbCustomer_Name.SetFocus
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
        str = "select * from Due_Tran where Customer_Code Like '" & "%" & search & "%" & "'"
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

