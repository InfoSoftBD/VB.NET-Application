VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmProd_Receive 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receive Info!"
   ClientHeight    =   10485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   Icon            =   "frmReceive.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   165
      ScaleHeight     =   660
      ScaleWidth      =   11085
      TabIndex        =   6
      Top             =   165
      Width           =   11115
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT RECEIVE ENTRY"
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
         Left            =   2070
         TabIndex        =   7
         Top             =   45
         Width           =   6990
      End
      Begin VB.Image Image3 
         Height          =   690
         Left            =   0
         Picture         =   "frmReceive.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11085
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   150
      ScaleHeight     =   720
      ScaleWidth      =   11085
      TabIndex        =   2
      Top             =   9570
      Width           =   11115
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   525
         Left            =   6075
         TabIndex        =   8
         Top             =   90
         Width           =   1830
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   525
         Left            =   9045
         TabIndex        =   5
         Top             =   90
         Width           =   1830
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   525
         Left            =   3120
         TabIndex        =   4
         Top             =   105
         Width           =   1830
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Save"
         Height          =   525
         Left            =   255
         TabIndex        =   3
         Top             =   105
         Width           =   1830
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8460
      Left            =   150
      ScaleHeight     =   8430
      ScaleWidth      =   11100
      TabIndex        =   1
      Top             =   990
      Width           =   11130
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
         Height          =   2265
         Left            =   180
         TabIndex        =   42
         Top             =   6045
         Width           =   10755
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
            Left            =   1560
            TabIndex        =   74
            Text            =   "Combo1"
            Top             =   1290
            Width           =   1515
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
            Left            =   8790
            TabIndex        =   56
            Text            =   "Text7"
            Top             =   1275
            Width           =   1710
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
            Left            =   5145
            TabIndex        =   55
            Text            =   "Combo2"
            Top             =   1275
            Width           =   1605
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
            Left            =   3135
            TabIndex        =   54
            Text            =   "Combo1"
            Top             =   1275
            Width           =   1950
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
            Left            =   8775
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   570
            Width           =   1710
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
            Left            =   6840
            TabIndex        =   52
            Text            =   "Text1"
            Top             =   1275
            Width           =   1560
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
            Left            =   1560
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   570
            Width           =   735
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
            Left            =   2385
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   570
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
            Left            =   3150
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   570
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
            Left            =   3915
            TabIndex        =   48
            Text            =   "Text1"
            Top             =   570
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
            Left            =   4680
            TabIndex        =   47
            Text            =   "Text1"
            Top             =   570
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
            Left            =   5415
            TabIndex        =   46
            Text            =   "Text1"
            Top             =   570
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
            Left            =   6180
            TabIndex        =   45
            Text            =   "Text1"
            Top             =   570
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
            Left            =   6945
            TabIndex        =   44
            Text            =   "Text1"
            Top             =   570
            Width           =   675
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
            Left            =   7710
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   570
            Width           =   675
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
            Left            =   3150
            TabIndex        =   73
            Top             =   1005
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
            Left            =   5175
            TabIndex        =   72
            Top             =   1005
            Width           =   615
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
            Left            =   6825
            TabIndex        =   71
            Top             =   1005
            Width           =   1020
         End
         Begin VB.Label lblCash 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount"
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
            Left            =   9270
            TabIndex        =   70
            Top             =   300
            Width           =   1155
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
            Left            =   1590
            TabIndex        =   69
            Top             =   1005
            Width           =   1080
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
            Left            =   1560
            TabIndex        =   68
            Top             =   300
            Width           =   750
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
            Left            =   2385
            TabIndex        =   67
            Top             =   300
            Width           =   645
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
            Left            =   3150
            TabIndex        =   66
            Top             =   300
            Width           =   645
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
            Left            =   3915
            TabIndex        =   65
            Top             =   300
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
            Left            =   4680
            TabIndex        =   64
            Top             =   300
            Width           =   540
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
            Left            =   5415
            TabIndex        =   63
            Top             =   300
            Width           =   540
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
            Left            =   6180
            TabIndex        =   62
            Top             =   300
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
            Left            =   6945
            TabIndex        =   61
            Top             =   300
            Width           =   435
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
            Left            =   7710
            TabIndex        =   60
            Top             =   300
            Width           =   435
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
            Left            =   8640
            TabIndex        =   59
            Top             =   1860
            Width           =   1905
         End
         Begin VB.Shape Shape1 
            Height          =   1575
            Left            =   8640
            Top             =   240
            Width           =   1965
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
            Left            =   30
            TabIndex        =   58
            Top             =   600
            Width           =   1365
         End
         Begin VB.Shape Shape2 
            Height          =   1575
            Left            =   1440
            Top             =   240
            Width           =   7065
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
            Left            =   30
            TabIndex        =   57
            Top             =   1350
            Width           =   1380
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vendor Information"
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
         TabIndex        =   23
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
            TabIndex        =   75
            Text            =   "Text4"
            Top             =   323
            Width           =   1545
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
            TabIndex        =   40
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
            TabIndex        =   30
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
            Left            =   1035
            TabIndex        =   0
            Text            =   "Text4"
            Top             =   323
            Width           =   1545
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
            TabIndex        =   25
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
            Left            =   1035
            TabIndex        =   24
            Text            =   "Combo1"
            Top             =   840
            Width           =   3240
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LC No./ Batch No."
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
            TabIndex        =   76
            Top             =   390
            Width           =   1590
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
            TabIndex        =   31
            Top             =   360
            Width           =   405
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
            Left            =   105
            TabIndex        =   29
            Top             =   390
            Width           =   840
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Code"
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
            Left            =   6135
            TabIndex        =   28
            Top             =   390
            Width           =   1125
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
            TabIndex        =   27
            Top             =   870
            Width           =   510
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            TabIndex        =   26
            Top             =   870
            Width           =   720
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
         Height          =   4425
         Left            =   150
         TabIndex        =   9
         Top             =   1590
         Width           =   10815
         Begin VB.TextBox txtR_Price 
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
            Left            =   3510
            TabIndex        =   77
            Top             =   1088
            Width           =   1110
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
            Left            =   210
            TabIndex        =   41
            Top             =   600
            Width           =   1500
         End
         Begin VB.TextBox txtD_Price 
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
            Left            =   5940
            TabIndex        =   38
            Top             =   1088
            Width           =   1110
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
            TabIndex        =   37
            Text            =   "Text2"
            Top             =   555
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
            TabIndex        =   34
            Text            =   "Text2"
            Top             =   555
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
            TabIndex        =   15
            Text            =   "Text7"
            Top             =   555
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
            Left            =   3600
            TabIndex        =   14
            Text            =   "Combo2"
            Top             =   600
            Width           =   2220
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
            TabIndex        =   13
            Text            =   "Text2"
            Top             =   555
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
            Left            =   1770
            TabIndex        =   12
            Top             =   600
            Width           =   1740
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
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   555
            Width           =   690
         End
         Begin VB.TextBox txtS_Price 
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
            TabIndex        =   10
            Top             =   1088
            Width           =   1035
         End
         Begin MSDataGridLib.DataGrid grdProd 
            Height          =   2445
            Left            =   210
            TabIndex        =   32
            Top             =   1575
            Width           =   10395
            _ExtentX        =   18336
            _ExtentY        =   4313
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
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retail Price"
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
            TabIndex        =   78
            Top             =   1155
            Width           =   1005
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dealer Price"
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
            Left            =   4725
            TabIndex        =   39
            Top             =   1155
            Width           =   1065
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
            TabIndex        =   36
            Top             =   255
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
            TabIndex        =   35
            Top             =   255
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
            Left            =   8415
            TabIndex        =   33
            Top             =   4080
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
            TabIndex        =   22
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
            Left            =   1770
            TabIndex        =   21
            Top             =   270
            Width           =   1245
         End
         Begin VB.Label Label6 
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
            Left            =   3600
            TabIndex        =   20
            Top             =   270
            Width           =   1710
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
            TabIndex        =   19
            Top             =   255
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
            TabIndex        =   18
            Top             =   255
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
            TabIndex        =   17
            Top             =   195
            Width           =   315
         End
         Begin VB.Label Label11 
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
            Left            =   225
            TabIndex        =   16
            Top             =   1155
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frmProd_Receive"
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
    .Columns(10).Width = 1200
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 1200
    .Columns(11).Alignment = dbgRight
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
        .Fields.Append "Description", adBSTR
        .Fields.Append "Qty", adBSTR
        .Fields.Append "Unit Price", adBSTR
        .Fields.Append "Total Price", adBSTR
        .Fields.Append "Commission", adBSTR
        .Fields.Append "Net Amount", adBSTR
        .Fields.Append "Sale Price", adBSTR
        .Fields.Append "Dealer Price", adBSTR
        .Fields.Append "Retail Price", adBSTR
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
        rsN!Prod_Name = rsProd![Product Name]
        rsN!Prod_Model = rsProd![Description]
        rsN!Purchase = Val(rsProd!Qty)
        rsN!Sale = 0
        rsN!Lift = 0
        rsN!Stock = rsU!Stock
        rsN!Prod_Price = rsProd![Unit Price]
        rsN!Com = Val(rsProd![Commission])
        rsN!Amount = rsProd![Total Price]
        rsN!Sale_Price = 0
        rsN.Update
        rsN.Close
 End Sub
Private Sub Vendor_Cr()
On Error Resume Next
Set rs = New ADODB.Recordset
        str = "select * from Vendor_Master where Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Cr = rs!Cr + Val(rsProd![Total Price]) - Val(rsProd![Commission])
        rs!Balance = rs!Balance + Val(rsProd![Total Price]) - Val(rsProd![Commission])
        rs.Update
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Vendor_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Vendor_Code = txtVendor_Code.Text
            rsN!Vendor_Name = cmbVendor_Name.Text
            rsN!Vendor_Address = txtVendor_Address.Text
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
            rsN.Open "Vendor_Master", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
        
            rsN!Date = txtDate.Text
            rsN!Vendor_Code = txtVendor_Code.Text
            rsN!Vendor_Name = cmbVendor_Name.Text
            rsN!Vendor_Address = txtVendor_Address.Text
            rsN!Open_Bal = 0
            rsN!Dr = 0
            rsN!Cr = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN!Balance = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN.Update
            rsN.Close
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Vendor_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
        
            rsN!Date = txtDate.Text
            rsN!INVOICE = txtInvoice.Text
            rsN!Vendor_Code = txtVendor_Code.Text
            rsN!Vendor_Name = cmbVendor_Name.Text
            rsN!Vendor_Address = txtVendor_Address.Text
            rsN!Description = rsProd![Product Name] + "," + rsProd![Description]
            rsN!Dr = 0
            rsN!Cr = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN!Balance = Val(rsProd![Total Price]) - Val(rsProd![Commission])
            rsN.Update
            rsN.Close
    End If
End Sub
Private Sub Vendor_Rev()
Set rs = New ADODB.Recordset
        str = "select * from Vendor_Master where Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Dr = rs!Dr + Val(rsProd![Total Price])
        rs!Balance = rs!Balance - Val(rsProd![Total Price])
        rs.Update
        
        Set rsN = New ADODB.Recordset
            rsN.Open "Vendor_Tran", conn, adOpenDynamic, adLockOptimistic, -1
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
        str = "select * from Vendor_Master where Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Dr = rs!Dr + Val(txtCash.Text)
        rs!Balance = rs!Balance - Val(txtCash.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Vendor_Tran", conn, adOpenDynamic, adLockOptimistic, -1
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
        str = "select * from Vendor_Master where Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Cr = rs!Cr + Val(txtCash.Text)
        rs!Balance = rs!Balance + Val(txtCash.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Vendor_Tran", conn, adOpenDynamic, adLockOptimistic, -1
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
        str = "select * from Vendor_Master where Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Dr = rs!Dr + Val(txtChq_Amnt.Text)
        rs!Balance = rs!Balance - Val(txtChq_Amnt.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Vendor_Tran", conn, adOpenDynamic, adLockOptimistic, -1
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
        str = "select * from Vendor_Master where Vendor_Code like '" & txtVendor_Code.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rs.EOF Then
        rs!Dr = rs!Dr + Val(txtChq_Amnt.Text)
        rs!Balance = rs!Balance + Val(txtChq_Amnt.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Vendor_Tran", conn, adOpenDynamic, adLockOptimistic, -1
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
        rsN!Description = "PCode-" & rsProd![Product Code]
        rsN!Dr = rsProd![Total Price]
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close

    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100150 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + rsProd![Total Price] - Val(rsProd![Commission])
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
        rsN!Cr = rsProd![Total Price]
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close


    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100143 & "'"
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
        rsN!Description = "SCode-" & txtVendor_Code.Text
        rsN!Dr = 0
        rsN!Cr = Val(rsProd![Commission])
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
        str = "SELECT DISTINCT Vendor_Code FROM Vendor_Master"
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
        str = "SELECT DISTINCT Vendor_Name FROM Vendor_Master"
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
Private Sub Invo_No()
Dim ID As String
Dim mid As Integer
    ID = txtInvoice.Text
    mid = 0
    
    Set rs = New ADODB.Recordset
        str = "select * from Vendor_Tran where Invoice like '" & ID & "'"
        rs.Open str, conn

    If Not rs.EOF Then
        rs.Close
    Else
        
        Set rsU = New ADODB.Recordset
            str = "select DISTINCT Invoice from Vendor_Tran order by Invoice"
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
            
            txtInvoice.Text = "PI/" & Year(CDate(Today)) & "/" & Format(mid, "000#")
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
Private Sub Preview()
If txtInvoice.Text = "" Then
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
        str = "select * from Invo_Prev where Invo_No like '" & Tran & "' and Sale >0"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly

        If Not rs.EOF Then
            rptInvoice.rsInvoice.ConnectionString = cnStr
            rptInvoice.rsInvoice.Source = str
            
            rptInvoice.Field19.DataField = "Sale"
            rptInvoice.Field16.DataField = "Sale_Price"
            rptInvoice.lblTitle.Caption = "PURCHASE INVOICE"
            rptInvoice.lblTitle.Font.Size = 14
            rptInvoice.txtID.Text = txtVendor_Code.Text
            rptInvoice.txtName.Text = cmbVendor_Name.Text
            rptInvoice.txtAddress.Text = txtVendor_Address.Text
'            rptInvoice.txtMobile.Text = txtMobile.Text
            
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


Private Sub clearTextboxes()
        txtInvoice.Text = ""
        txtRef_No.Text = ""
        txtDate.Text = ""
        txtVendor_Code.Text = ""
        cmbVendor_Name.Text = ""
        txtVendor_Address.Text = ""
        
        txtProd_Sl.Text = ""
        cmbProd_Name.Text = ""
        cmbProd_Model.Text = ""
        
        txtProd_Cost.Text = "0.00"
        txtQty.Text = "0"
        txtProd_Price.Text = "0.00"
        txtS_Price.Text = "0.00"
        txtD_Price.Text = "0.00"
        txtR_Price.Text = "0.00"
        txtCommission.Text = "0.00"
        cmbPercent.Text = "0.00"
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
            'txtStock.Text = Format$(Val(rsN!Stock), "###0.00")
            txtProd_Cost.Text = Format$(Val(rsN!Prod_Price), "###0.00")
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
        str = "SELECT * FROM Vendor_Master where Vendor_Name Like '" & Vendor & "'order by Vendor_code"
        rsN.Open str, conn
        
If Not rsN.EOF Then
            txtVendor_Code.Text = rsN!Vendor_Code
            cmbVendor_Name.Text = rsN!Vendor_Name
            txtVendor_Address.Text = rsN!Vendor_Address
            rsN.Close
Else
        
    If MsgBox("Do you want add new Vendor?", vbInformation + vbYesNo, "Add New") = vbYes Then
        
        Set rsU = New ADODB.Recordset
            str = "select * from Vendor_Master order by Vendor_Code"
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
            rsN!Ref_Code = txtVendor_Code.Text
            rsN!Ref_Name = cmbVendor_Name.Text
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

Private Sub cmdReceive_Click()
On Error Resume Next
If txtInvoice.Text = "" Then
MsgBox "Please Input Invoice No!", vbCritical
txtInvoice.SetFocus
Exit Sub
End If

Dim Prod As String
rsProd.MoveFirst
    
    Do While Not rsProd.EOF
        Prod = rsProd![Product Code]
        
        Set rsU = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Code like '" & Prod & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
             rsU!Purchase = rsU!Purchase + Val(rsProd!Qty)
             'rsU!Prod_Price = Val(rsProd![Unit Price])
             rsU!Prod_Price = Format$(((Val(rsU!Prod_Price) * Val(rsU!Stock)) + (Val(rsProd![Unit Price]) * Val(rsProd!Qty))) / (Val(rsU!Stock) + Val(rsProd!Qty)), "###0.0000")
             rsU!Stock = rsU!Stock + Val(rsProd!Qty)
             rsU!Com = rsU!Com + Val(rsProd![Commission])
             rsU!Amount = rsU!Stock * rsU!Prod_Price
             rsU!Sale_Price = Val(rsProd![Sale Price])
             rsU!Dealer_Price = Val(rsProd![Dealer Price])
             rsU!Retail_Price = Val(rsProd![Retail Price])
             rsU!Ref_no = txtRef_No.Text
             rsU.Update
             Call Prod_Tran
             Call GL_Dr
             Call Vendor_Cr
             rsProd.MoveNext
             rsU.Close
                    
    Else
        rsU.Close
    
    
    
    Set rsU = New ADODB.Recordset
        rsU.Open "Prod_Master", conn, adOpenDynamic, adLockOptimistic, -1
        rsU.AddNew
        
        rsU!Prod_Code = rsProd![Product Code]
        rsU!Prod_Name = rsProd![Product Name]
        rsU!Prod_Model = rsProd![Description]
        rsU!Open_Bal = 0
        rsU!Purchase = Val(rsProd!Qty)
        rsU!Sale = 0
        rsU!Lift = 0
        rsU!Return = 0
        rsU!Stock = Val(rsProd!Qty)
        rsU!Prod_Price = Val(rsProd![Unit Price])
        rsU!Com = Val(rsProd![Commission])
        rsU!Amount = rsU!Stock * rsU!Prod_Price
        rsU!Sale_Price = Val(rsProd![Sale Price])
        rsU!Dealer_Price = Val(rsProd![Dealer Price])
        rsU!Retail_Price = Val(rsProd![Retail Price])
        rsU!Ref_no = txtRef_No.Text
        rsU.Update
      
        Call Prod_Tran
        Call GL_Dr
        Call Vendor_Cr
        rsProd.MoveNext
        rsU.Close
    End If


    Loop
    
            
    If Val(txtCash.Text) > 0 Then
        Call Cash_Cr
    End If
    If Val(txtChq_Amnt.Text) > 0 Then
       Call Bank_Cr
    End If
            

    Call Prod_Name
    Call Prod_Model
    Call Prod_Code
    
    Call Vendor_Name
    Call Vendor_Code
    Call clearTextboxes
    Call grd_Prod
    Call Col_Prod
    Call Invo_No
    txtDate.Text = Today
    txtInvoice.SetFocus
    cmdReceive.Enabled = False
    cmdDelete.Enabled = False
'    cmdUpdate.Enabled = False
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
    Call Invo_No
        txtDate.Text = Today
        
        cmdReceive.Enabled = False
'        cmdUpdate.Enabled = False
    Exit Sub
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
    txtS_Price.SelStart = 0
    txtS_Price.SelLength = Len(txtS_Price.Text)
    txtS_Price.SetFocus
End If
End Sub

Private Sub txtEngine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty.Text)
    txtQty.SetFocus
End If
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
                txtDate.Text = rs!D_ate
            
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
               txtCash.Text = 0
           End If
    
            Set rs = New ADODB.Recordset
               str = "select * from Bank_Tran where MR_No like '" & Tran & "'"
               rs.Open str, conn
           
           If Not rs.EOF Then
               txtChq_Amnt.Text = rs!Cr
               rs.Close
           Else
               rs.Close
               txtChq_Amnt.Text = 0
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

Private Sub txtR_Price_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtR_Price.Text = Format$(Val(txtR_Price.Text), "###0.00")
    txtD_Price.SelStart = 0
    txtD_Price.SelLength = Len(txtD_Price.Text)
    txtD_Price.SetFocus
End If
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
Dim Prod As String
    Prod = txtProd_Sl.Text
    
    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            
            txtProd_Cost.Text = Format$(Val(rsN!Prod_Price), "###0.00")
            txtS_Price.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            txtD_Price.Text = Format$(Val(rsN!Dealer_Price), "###0.00")
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
Private Sub txtD_Price_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtD_Price.Text = Format$(Val(txtD_Price.Text), "###0.00")
    txtProd_Sl.SetFocus
End If
End Sub

Private Sub txtD_Price_LostFocus()
txtProd_Price.Text = Format$(Val(txtProd_Price.Text), "###0.00")
        rsProd.AddNew
        rsProd!sl = Prod_Sl
        rsProd![Product Code] = txtProd_Sl.Text
        rsProd![Product Name] = cmbProd_Name.Text
        rsProd![Description] = cmbProd_Model.Text
        rsProd!Qty = txtQty.Text
        rsProd![Unit Price] = Format$(Val(txtProd_Cost.Text), "###0.00")
        rsProd![Total Price] = Format$(Val(txtProd_Price.Text), "###0.00")
        rsProd![Commission] = Format$(Val(txtCommission.Text), "###0.00")
        rsProd![Net Amount] = Format$((Val(txtProd_Price) + Val(txtCommission.Text)), "###0.00")
        rsProd![Sale Price] = txtS_Price.Text
        rsProd![Dealer Price] = txtD_Price.Text
        rsProd![Retail Price] = txtR_Price.Text
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
            cmbProd_Model.Text = ""
            txtQty.Text = ""
            txtProd_Cost.Text = "0.00"
            txtProd_Price.Text = "0.00"
            txtS_Price.Text = "0.00"
             txtR_Price.Text = "0.00"
            txtD_Price.Text = "0.00"
            txtCommission.Text = "0.00"
            cmbPercent.Text = "0.00"
End Sub

Private Sub txtRef_No_LostFocus()
Dim Tran As String
    Tran = txtRef_No.Text

    Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where Ref_No like '" & Tran & "'"
        rs.Open str, conn

 If Not rs.EOF Then
    MsgBox "Duplicate invoice No.", vbCritical, "Erroe!"
            rs.MoveFirst
                txtInvoice.Text = rs!Invo_No
                txtVendor_Code.Text = rs!Ref_Code
                cmbVendor_Name.Text = rs!Ref_Name
                txtRef_No.Text = rs!Ref_no
                txtDate.Text = rs!D_ate
                
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
               txtCash.Text = 0
           End If
    
            Set rs = New ADODB.Recordset
               str = "select * from Bank_Tran where MR_No like '" & Tran & "'"
               rs.Open str, conn
           
           If Not rs.EOF Then
               txtChq_Amnt.Text = rs!Cr
               rs.Close
           Else
               rs.Close
               txtChq_Amnt.Text = 0
           End If
            
    
        cmdPrint.Enabled = True
        cmdReceive.Enabled = False
        cmdDelete.Enabled = True
    Else
    rs.Close
    Tran = txtRef_No.Text
    'Call clearTextboxes
    
    txtRef_No.Text = Tran
    Call Invo_No
    txtDate.Text = Today
    Exit Sub
  End If


End Sub

Private Sub txtS_Price_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtS_Price.Text = Format$(Val(txtS_Price.Text), "###0.00")
    txtR_Price.SelStart = 0
    txtR_Price.SelLength = Len(txtR_Price.Text)
    txtR_Price.SetFocus
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
        str = "SELECT * FROM Vendor_Master where Vendor_Code Like '" & Vendor & "'order by Vendor_code"
        rsN.Open str, conn
        
If Not rsN.EOF Then
            txtVendor_Code.Text = rsN!Vendor_Code
            cmbVendor_Name.Text = rsN!Vendor_Name
            txtVendor_Address.Text = rsN!Vendor_Address
            rsN.Close
Else
        
    If MsgBox("Do you want add new Vendor?", vbInformation + vbYesNo, "Add New") = vbYes Then
        
        Set rsU = New ADODB.Recordset
            str = "select * from Vendor_Master order by Vendor_Code"
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
