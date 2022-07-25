VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmLC 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "L/C Registrar"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   Icon            =   "frmLC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   150
      ScaleHeight     =   660
      ScaleWidth      =   11445
      TabIndex        =   46
      Top             =   135
      Width           =   11475
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LETTER OF CREDIT ENTRY"
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
         Left            =   1395
         TabIndex        =   47
         Top             =   45
         Width           =   8565
      End
      Begin VB.Image Image3 
         Height          =   690
         Left            =   0
         Picture         =   "frmLC.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11445
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   135
      ScaleHeight     =   720
      ScaleWidth      =   11490
      TabIndex        =   41
      Top             =   9540
      Width           =   11520
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   525
         Left            =   6075
         TabIndex        =   45
         Top             =   90
         Width           =   1830
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   525
         Left            =   9045
         TabIndex        =   44
         Top             =   90
         Width           =   1830
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   525
         Left            =   3120
         TabIndex        =   43
         Top             =   105
         Width           =   1830
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Save"
         Height          =   525
         Left            =   255
         TabIndex        =   42
         Top             =   105
         Width           =   1830
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8460
      Left            =   135
      ScaleHeight     =   8430
      ScaleWidth      =   11460
      TabIndex        =   0
      Top             =   960
      Width           =   11490
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "L/C Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1890
         Left            =   135
         TabIndex        =   22
         Top             =   90
         Width           =   11220
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
            Left            =   1425
            TabIndex        =   31
            Text            =   "Combo1"
            Top             =   840
            Width           =   1950
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
            Left            =   9705
            TabIndex        =   30
            Text            =   "Text3"
            Top             =   323
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
            Left            =   1425
            TabIndex        =   29
            Text            =   "Text4"
            Top             =   323
            Width           =   1950
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
            Left            =   4935
            TabIndex        =   28
            Text            =   "Combo1"
            Top             =   840
            Width           =   3285
         End
         Begin VB.TextBox txtExport_Date 
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
            Left            =   9690
            TabIndex        =   27
            Text            =   "Text3"
            Top             =   833
            Width           =   1380
         End
         Begin VB.TextBox txtShipment_Date 
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
            Left            =   9690
            TabIndex        =   26
            Text            =   "Text3"
            Top             =   1350
            Width           =   1380
         End
         Begin VB.ComboBox cmbBank_Name 
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
            Left            =   4935
            TabIndex        =   25
            Text            =   "Combo1"
            Top             =   330
            Width           =   3300
         End
         Begin VB.ComboBox cmbPort_Name 
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
            Left            =   1425
            TabIndex        =   24
            Text            =   "Combo1"
            Top             =   1357
            Width           =   1950
         End
         Begin VB.ComboBox cmbInsurance 
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
            Left            =   4935
            TabIndex        =   23
            Text            =   "Combo1"
            Top             =   1357
            Width           =   3285
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "L/C Issue Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8295
            TabIndex        =   40
            Top             =   390
            Width           =   1290
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "L/C No."
            BeginProperty Font 
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
            TabIndex        =   39
            Top             =   390
            Width           =   660
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exporter Code"
            BeginProperty Font 
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
            TabIndex        =   38
            Top             =   900
            Width           =   1245
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exporter Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3435
            TabIndex        =   37
            Top             =   900
            Width           =   1305
         End
         Begin VB.Label Label16 
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
            Left            =   8295
            TabIndex        =   36
            Top             =   900
            Width           =   1020
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shipment Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8295
            TabIndex        =   35
            Top             =   1417
            Width           =   1290
         End
         Begin VB.Label Label22 
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
            Left            =   3465
            TabIndex        =   34
            Top             =   390
            Width           =   1020
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port Name"
            BeginProperty Font 
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
            TabIndex        =   33
            Top             =   1417
            Width           =   930
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Insurance Name"
            BeginProperty Font 
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
            TabIndex        =   32
            Top             =   1417
            Width           =   1410
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Commodity Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6225
         Left            =   150
         TabIndex        =   1
         Top             =   2040
         Width           =   11175
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
            Left            =   1530
            TabIndex        =   50
            Top             =   495
            Width           =   2685
         End
         Begin VB.ComboBox cmbPack_Type 
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
            ItemData        =   "frmLC.frx":A4D0
            Left            =   3150
            List            =   "frmLC.frx":A4EF
            TabIndex        =   48
            Text            =   "Combo2"
            Top             =   2115
            Width           =   1050
         End
         Begin VB.TextBox txtBDT_Price 
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
            Left            =   6570
            TabIndex        =   10
            Top             =   2108
            Width           =   1425
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
            Left            =   6570
            TabIndex        =   9
            Text            =   "Text2"
            Top             =   1568
            Width           =   1425
         End
         Begin VB.TextBox txtBank_Charge 
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
            Left            =   9585
            TabIndex        =   8
            Text            =   "Text2"
            Top             =   1035
            Width           =   1335
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
            Left            =   6570
            TabIndex        =   7
            Text            =   "Text7"
            Top             =   488
            Width           =   1425
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
            Left            =   1530
            TabIndex        =   6
            Text            =   "Combo2"
            Top             =   1575
            Width           =   2685
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
            Left            =   6570
            TabIndex        =   5
            Text            =   "Text2"
            Top             =   1028
            Width           =   1425
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
            Left            =   1530
            TabIndex        =   4
            Top             =   1035
            Width           =   2685
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
            Left            =   1530
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   2108
            Width           =   1050
         End
         Begin VB.TextBox txtMargin 
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
            Left            =   9585
            TabIndex        =   2
            Top             =   495
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid grdProd 
            Height          =   3255
            Left            =   180
            TabIndex        =   11
            Top             =   2790
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   5741
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
         Begin VB.Shape Shape4 
            Height          =   825
            Left            =   8415
            Shape           =   4  'Rounded Rectangle
            Top             =   1710
            Width           =   2490
         End
         Begin VB.Shape Shape3 
            Height          =   2310
            Left            =   8280
            Top             =   315
            Width           =   2760
         End
         Begin VB.Shape Shape2 
            Height          =   2310
            Left            =   4500
            Top             =   315
            Width           =   3660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "H/S Code"
            BeginProperty Font 
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
            TabIndex        =   51
            Top             =   555
            Width           =   840
         End
         Begin VB.Shape Shape1 
            Height          =   2310
            Left            =   180
            Top             =   315
            Width           =   4200
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2700
            TabIndex        =   49
            Top             =   2175
            Width           =   345
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "L/C Margin"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8325
            TabIndex        =   21
            Top             =   555
            Width           =   945
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total L/C Value (BDT)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4635
            TabIndex        =   20
            Top             =   2175
            Width           =   1890
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Charge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8325
            TabIndex        =   19
            Top             =   1095
            Width           =   1125
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
            Left            =   8595
            TabIndex        =   18
            Top             =   2025
            Width           =   2055
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Commodity"
            BeginProperty Font 
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
            TabIndex        =   17
            Top             =   1095
            Width           =   990
         End
         Begin VB.Label Label6 
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
            Left            =   315
            TabIndex        =   16
            Top             =   1635
            Width           =   975
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Price (USD)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4635
            TabIndex        =   15
            Top             =   555
            Width           =   1440
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Price (USD)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4635
            TabIndex        =   14
            Top             =   1095
            Width           =   1515
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
            BeginProperty Font 
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
            TabIndex        =   13
            Top             =   2175
            Width           =   735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exchange Rate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4635
            TabIndex        =   12
            Top             =   1635
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Prod_Sl As Integer
Dim netAmnt As Double
Dim strloan As String
Dim rsProd As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
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
lblNet.Caption = "Total Tk. " + strloan
    ' Create an initial recordset, just for demonstration purposes,
    ' and assign it to the DataGrid control's DataSource property.
    Set rsProd = New ADODB.Recordset
    With rsProd
        .Fields.Append "Sl", adBSTR
        .Fields.Append "Product Code", adBSTR
        .Fields.Append "Product Name", adBSTR
        .Fields.Append "Description", adBSTR
        .Fields.Append "Qty", adBSTR
        .Fields.Append "Unit", adBSTR
        .Fields.Append "Unit Price", adBSTR
        .Fields.Append "Total Price", adBSTR
        .Fields.Append "Rate", adBSTR
        .Fields.Append "BDT Amount", adBSTR
        .Fields.Append "Margin", adBSTR
        .Fields.Append "Bank Charge", adBSTR
        .Open
   
    End With
    Set grdProd.DataSource = rsProd
    
End Sub

Private Sub clearTextboxes()
        txtInvoice.Text = ""
        txtDate.Text = ""
        txtExport_Date.Text = ""
        txtShipment_Date.Text = ""
        txtVendor_Code.Text = ""
        cmbVendor_Name.Text = ""
        cmbBank_Name.Text = ""
        cmbPort_Name.Text = ""
        cmbInsurance.Text = ""
        
        txtProd_Sl.Text = ""
        cmbProd_Name.Text = ""
        cmbProd_Model.Text = ""
        txtQty.Text = "0"
        cmbPack_Type.Text = ""
        txtProd_Cost.Text = "0.00"
        txtProd_Price.Text = "0.00"
        
        cmbPercent.Text = "0.00"
        txtBDT_Price.Text = "0.00"
        
        txtBank_Charge.Text = "0.00"
        txtMargin.Text = "0.00"
End Sub
Private Sub Insurance()
On Error Resume Next
        cmbInsurance.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Insurance FROM LC_Tran"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbInsurance.AddItem rs!Insurance
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub P_ort()
On Error Resume Next
        cmbPort_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Port_Name FROM LC_Tran"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbPort_Name.AddItem rs!Port_Name
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub Prod_Code()
On Error Resume Next
        txtProd_Sl.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Code FROM LC_Tran"
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
On Error Resume Next
        cmbProd_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name FROM LC_Tran"
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
On Error Resume Next
        cmbProd_Model.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Model FROM LC_Tran"
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
On Error Resume Next
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
On Error Resume Next
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
Private Sub Bank_Name()
On Error Resume Next
        cmbBank_Name.Clear
         Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Bank_Name FROM Bank_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbBank_Name.AddItem rsN!Bank_Name
        rsN.MoveNext
        Loop
        rsN.Close
End Sub

Private Sub cmbBank_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDate.SetFocus
End If
End Sub

Private Sub cmbInsurance_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtShipment_Date.SetFocus
End If
End Sub



Private Sub cmbPack_Type_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Cost.SelStart = 0
txtProd_Cost.SelLength = Len(txtProd_Cost.Text)
txtProd_Cost.SetFocus
End If
End Sub

Private Sub cmbPercent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtBDT_Price.SelStart = 0
txtBDT_Price.SelLength = Len(txtBDT_Price.Text)
txtBDT_Price.SetFocus
End If
End Sub

Private Sub cmbPercent_LostFocus()
        txtBDT_Price.Text = Format$(Val(cmbPercent.Text) * Val(txtProd_Price.Text), "###0.00")
        txtBDT_Price.SelStart = 0
        txtBDT_Price.SelLength = Len(txtBDT_Price.Text)
        txtBDT_Price.SetFocus
 End Sub

Private Sub cmbPort_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbInsurance.SetFocus
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
'            txtS_Price.Text = Format$(Val(rsN!Sale_Price), "###0.00")
            txtBDT_Price.Text = Format$(Val(rsN!Dealer_Price), "###0.00")
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
txtExport_Date.SetFocus
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
'            txtVendor_Address.Text = rsN!Vendor_Address
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
'            txtVendor_Address.Text = ""
     Else
        rsN.Close
        Exit Sub
    End If
End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdReceive_Click()
'On Error Resume Next
If txtInvoice.Text = "" Then
MsgBox "Please Input Invoice No!", vbCritical
txtInvoice.SetFocus
Exit Sub
End If

Dim Prod As String
rsProd.MoveFirst
    
  
Do While Not rsProd.EOF
   
    Set rsN = New ADODB.Recordset
        rsN.Open "LC_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!Invo_No = txtInvoice.Text
        rsN!LC_No = txtInvoice.Text
        rsN!D_ate = Today
        rsN!LC_Date = txtDate.Text
        rsN!Export_Date = txtExport_Date.Text
        rsN!Shipment_Date = txtShipment_Date.Text
        rsN!Bank_Name = cmbBank_Name.Text
        rsN!Exporter_Code = txtVendor_Code.Text
        rsN!Exporter_Name = cmbVendor_Name.Text
        rsN!Port_Name = cmbPort_Name.Text
        rsN!Insurance = cmbInsurance.Text
        
       
        rsN!Prod_Code = rsProd![Product Code]
        rsN!Prod_Name = rsProd![Product Name]
        rsN!Prod_Model = rsProd![Description]
        rsN!Qty = Val(rsProd!Qty)
        rsN!Unit = rsProd!Unit
        
        rsN!Prod_Price = rsProd![Unit Price]
        rsN!Total_Price = rsProd![Total Price]
        rsN!Rate = Val(rsProd![Rate])
        rsN!BDT_Amount = rsProd![BDT Amount]
        rsN!Margin = rsProd![Margin]
        rsN!Bank_Charge = rsProd![Bank Charge]
        
        rsN!User_Id = User_Id
        rsN!User_Name = User_Name
        rsN.Update
        rsN.Close


        rsProd.MoveNext
        
    

    Loop
    
    
    Call Prod_Name
    Call Prod_Model
    Call Prod_Code
    
    Call Vendor_Name
    Call Vendor_Code
    Call clearTextboxes
    Call grd_Prod
    Call Col_Prod
    
    txtInvoice.SetFocus
    cmdReceive.Enabled = False
    cmdDelete.Enabled = False
'    cmdUpdate.Enabled = False
Exit Sub
   Resume Next

End Sub

Private Sub Form_Load()
    Call Prod_Name
    Call Prod_Code
    Call Prod_Model
    Call P_ort
    Call Insurance

    Call grd_Prod
    Call Col_Prod
    Call Vendor_Name
    Call Vendor_Code
'    Call Account_Name
    Call Bank_Name
'    Call Branch_Name
Call clearTextboxes
End Sub

Private Sub txtBank_Charge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtProd_Sl.SetFocus
End If
End Sub

Private Sub txtBank_Charge_LostFocus()
    txtProd_Price.Text = Format$(Val(txtProd_Price.Text), "###0.00")
        rsProd.AddNew
        rsProd!sl = Prod_Sl
        rsProd![Product Code] = txtProd_Sl.Text
        rsProd![Product Name] = cmbProd_Name.Text
        rsProd![Description] = cmbProd_Model.Text
        rsProd!Qty = txtQty.Text
        rsProd!Unit = cmbPack_Type.Text
        rsProd![Unit Price] = Format$(Val(txtProd_Cost.Text), "###0.00")
        rsProd![Total Price] = Format$(Val(txtProd_Price.Text), "###0.00")
        rsProd![Rate] = cmbPercent.Text
        rsProd![BDT Amount] = Format$(Val(txtBDT_Price.Text), "###0.00")
        rsProd![Margin] = txtMargin.Text
        rsProd![Bank Charge] = txtBank_Charge.Text
        rsProd.Update
        
        rsProd.MoveFirst
        Prod_Sl = 0
        netAmnt = 0
        Do While Not rsProd.EOF
            Prod_Sl = Prod_Sl + 1
            rsProd!sl = Prod_Sl
            
            netAmnt = netAmnt + Val(rsProd![Total Price])
            rsProd.MoveNext
        Loop
        strloan = 0
        strloan = Format$(Round(Val(netAmnt)), "###0.00")
        lblNet.Caption = "Total USD. " + strloan
        
    Set grdProd.DataSource = rsProd
    grdProd.Refresh
    
            Call Col_Prod
            txtProd_Sl.Text = ""
            cmbProd_Name.Text = ""
            cmbProd_Model.Text = ""
            txtQty.Text = "0"
            cmbPack_Type.Text = ""
            txtProd_Cost.Text = "0.00"
            txtProd_Price.Text = "0.00"
            txtBDT_Price.Text = "0.00"
            cmbPercent.Text = "0.00"
            txtMargin.Text = "0.00"
            txtBank_Charge.Text = "0.00"
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtVendor_Code.SetFocus
End If
End Sub

Private Sub txtExport_Date_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbPort_Name.SetFocus
End If
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtInvoice.Text = "" Then
Exit Sub
Else
cmdReceive.Enabled = True
cmbBank_Name.SetFocus
End If
End If

End Sub

Private Sub txtMargin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtBank_Charge.SelStart = 0
  txtBank_Charge.SelLength = Len(txtBank_Charge.Text)
  txtBank_Charge.SetFocus
End If
End Sub

Private Sub txtProd_Cost_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProd_Price.SetFocus
End If
End Sub
Private Sub txtProd_Cost_LostFocus()
        txtProd_Price.Text = Format$(Val(txtQty.Text) * Val(txtProd_Cost.Text), "###0.00")
        txtProd_Cost.Text = Format$(Val(txtProd_Cost.Text), "###0.00")
        
        cmbPercent.SelStart = 0
        cmbPercent.SelLength = Len(cmbPercent.Text)
        cmbPercent.SetFocus
 
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
cmbPack_Type.SetFocus
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
     Exit Sub
    End If
    
End If

Dim Prod As String
    Prod = txtProd_Sl.Text
    
    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM LC_Tran where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            
            txtProd_Cost.Text = Format$(Val(rsN!Prod_Price), "###0.00")
'            txtS_Price.Text = Format$(Val(rsN!Sale_Price), "###0.00")
'            txtBDT_Price.text = Format$(Val(rsN!Dealer_Price), "###0.00")
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
Private Sub txtBDT_Price_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBDT_Price.Text = Format$(Val(txtBDT_Price.Text), "###0.00")
    txtMargin.SelStart = 0
    txtMargin.SelLength = Len(txtMargin.Text)
    txtMargin.SetFocus
End If
End Sub

Private Sub txtShipment_Date_KeyPress(KeyAscii As Integer)
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
'            txtVendor_Address.Text = rsN!Vendor_Address
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
'            txtVendor_Address.Text = ""
     Else
        rsN.Close
        Exit Sub
    End If
End If

End Sub
