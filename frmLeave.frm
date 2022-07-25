VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLeave 
   Appearance      =   0  'Flat
   BackColor       =   &H00004000&
   Caption         =   "Leave Entry"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   Icon            =   "frmLeave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   7200
      ScaleHeight     =   555
      ScaleWidth      =   4155
      TabIndex        =   56
      Top             =   7110
      Width           =   4215
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   120
         Top             =   120
         Width           =   3855
         _ExtentX        =   6800
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLeave.frx":0442
      Height          =   6375
      Left            =   7440
      TabIndex        =   55
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   12648384
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   6855
      Left            =   7200
      ScaleHeight     =   6795
      ScaleWidth      =   4155
      TabIndex        =   54
      Top             =   120
      Width           =   4215
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   135
      ScaleHeight     =   555
      ScaleWidth      =   6885
      TabIndex        =   43
      Top             =   7110
      Width           =   6945
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   135
         TabIndex        =   53
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   2340
         TabIndex        =   50
         Top             =   90
         Width           =   1080
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   4635
         TabIndex        =   49
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   375
         Left            =   3510
         TabIndex        =   48
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   1215
         TabIndex        =   47
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Save"
         Height          =   375
         Left            =   135
         TabIndex        =   46
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   375
         Left            =   5760
         TabIndex        =   45
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   2340
         TabIndex        =   44
         Top             =   90
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Leave Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   135
      TabIndex        =   20
      Top             =   5445
      Width           =   6945
      Begin VB.ComboBox cmbTType 
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
         ItemData        =   "frmLeave.frx":0457
         Left            =   5175
         List            =   "frmLeave.frx":0461
         TabIndex        =   31
         Text            =   "Combo2"
         Top             =   945
         Width           =   1590
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
         ItemData        =   "frmLeave.frx":0471
         Left            =   1665
         List            =   "frmLeave.frx":0490
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   945
         Width           =   2175
      End
      Begin VB.TextBox txtTo 
         Alignment       =   2  'Center
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
         Left            =   3150
         TabIndex        =   23
         Text            =   "Text6"
         Top             =   360
         Width           =   1230
      End
      Begin VB.TextBox txtfrom 
         Alignment       =   2  'Center
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
         Left            =   1125
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   360
         Width           =   1230
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  'Center
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
         Left            =   5535
         TabIndex        =   21
         Text            =   "Text4"
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4005
         TabIndex        =   32
         Top             =   990
         Width           =   945
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Category"
         BeginProperty Font 
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
         TabIndex        =   30
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2415
         TabIndex        =   29
         Top             =   405
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
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
         TabIndex        =   28
         Top             =   405
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Days"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4455
         TabIndex        =   27
         Top             =   405
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Leave Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   135
      TabIndex        =   13
      Top             =   2700
      Width           =   6945
      Begin VB.TextBox txtExLBalance 
         Alignment       =   2  'Center
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
         Left            =   5490
         TabIndex        =   51
         Text            =   "Text13"
         Top             =   1935
         Width           =   1140
      End
      Begin VB.TextBox txtMLBalance 
         Alignment       =   2  'Center
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
         Left            =   4185
         TabIndex        =   41
         Text            =   "Text12"
         Top             =   1935
         Width           =   1140
      End
      Begin VB.TextBox txtSLBalance 
         Alignment       =   2  'Center
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
         Left            =   2880
         TabIndex        =   36
         Text            =   "Text11"
         Top             =   1935
         Width           =   1140
      End
      Begin VB.TextBox txtELBalance 
         Alignment       =   2  'Center
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
         Left            =   1575
         TabIndex        =   35
         Text            =   "Text10"
         Top             =   1935
         Width           =   1140
      End
      Begin VB.TextBox txtCLBalance 
         Alignment       =   2  'Center
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
         Left            =   270
         TabIndex        =   33
         Text            =   "Text9"
         Top             =   1935
         Width           =   1140
      End
      Begin VB.TextBox txtWithout 
         Alignment       =   2  'Center
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
         Left            =   5490
         TabIndex        =   26
         Text            =   "Text8"
         Top             =   900
         Width           =   1140
      End
      Begin VB.TextBox txtHalf 
         Alignment       =   2  'Center
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
         Left            =   4185
         TabIndex        =   25
         Text            =   "Text7"
         Top             =   900
         Width           =   1140
      End
      Begin VB.TextBox txtSLavailed 
         Alignment       =   2  'Center
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
         Left            =   2880
         TabIndex        =   16
         Text            =   "Text3"
         Top             =   900
         Width           =   1140
      End
      Begin VB.TextBox txtELavailed 
         Alignment       =   2  'Center
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
         Left            =   1575
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   900
         Width           =   1140
      End
      Begin VB.TextBox txtCLavailed 
         Alignment       =   2  'Center
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
         Left            =   270
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   900
         Width           =   1140
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ex-L Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5490
         TabIndex        =   52
         Top             =   1620
         Width           =   1170
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "M.L Balance"
         BeginProperty Font 
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
         TabIndex        =   42
         Top             =   1620
         Width           =   1185
      End
      Begin VB.Shape Shape2 
         Height          =   1095
         Left            =   135
         Top             =   315
         Width           =   6675
      End
      Begin VB.Shape Shape1 
         Height          =   870
         Left            =   135
         Top             =   1575
         Width           =   6675
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Leave without Pay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   5445
         TabIndex        =   40
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Leave with Half Pay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   4230
         TabIndex        =   39
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S.L Balance"
         BeginProperty Font 
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
         TabIndex        =   38
         Top             =   1620
         Width           =   1065
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E.L Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1575
         TabIndex        =   37
         Top             =   1620
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.L Balance"
         BeginProperty Font 
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
         TabIndex        =   34
         Top             =   1620
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S.L. Availed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2925
         TabIndex        =   19
         Top             =   495
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E.L. Availed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1620
         TabIndex        =   18
         Top             =   495
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.L. Availed"
         BeginProperty Font 
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
         Top             =   495
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
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
      Height          =   2400
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   6945
      Begin VB.TextBox txtTran 
         Alignment       =   2  'Center
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
         Left            =   5235
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   315
         Width           =   1545
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
         Left            =   1785
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   810
         Width           =   5010
      End
      Begin VB.TextBox txtID 
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
         Left            =   1785
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   315
         Width           =   1545
      End
      Begin VB.ComboBox cmbDesignation 
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
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1335
         Width           =   1545
      End
      Begin VB.ComboBox cmbGrade 
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
         Left            =   5235
         TabIndex        =   3
         Text            =   "Combo2"
         Top             =   1335
         Width           =   1545
      End
      Begin VB.TextBox txtJoinbk 
         Height          =   375
         Left            =   1785
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   1860
         Width           =   1545
      End
      Begin VB.TextBox txtJoinBr 
         Height          =   375
         Left            =   5235
         TabIndex        =   1
         Text            =   "Text4"
         Top             =   1860
         Width           =   1545
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3720
         TabIndex        =   58
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label3 
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
         Left            =   675
         TabIndex        =   12
         Top             =   1350
         Width           =   1020
      End
      Begin VB.Label Label2 
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
         Left            =   1185
         TabIndex        =   11
         Top             =   855
         Width           =   510
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID No."
         BeginProperty Font 
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
         TabIndex        =   10
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grader"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4575
         TabIndex        =   9
         Top             =   1380
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Joining Date"
         BeginProperty Font 
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
         TabIndex        =   8
         Top             =   1905
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Joining Date"
         BeginProperty Font 
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
         TabIndex        =   7
         Top             =   1920
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub clearTextboxes()
        txtID.Text = ""
        txtName.Text = ""
        cmbDesignation.Text = ""
        cmbGrade.Text = ""
        txtJoinbk.Text = ""
        txtJoinBr.Text = ""
        txtCLavailed.Text = ""
        txtELavailed.Text = ""
        txtSLavailed.Text = ""
        txtHalf.Text = ""
        txtWithout.Text = ""
        txtCLBalance.Text = ""
        txtELBalance.Text = ""
        txtSLBalance.Text = ""
        txtMLBalance.Text = ""
        txtExLBalance.Text = ""
        txtDay.Text = ""
        txtfrom.Text = ""
        txtTo.Text = ""
        cmbLType.Text = ""
        cmbTType.Text = ""
End Sub

Private Sub cmbLType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbTType.SetFocus
End If
End Sub

Private Sub cmbTType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmbTType.Text = "" Then
Exit Sub
End If
If cmbTType.Text = "Taken" Then
cmdAdd.Visible = True
cmdSave.Visible = False
cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdAdd.SetFocus
End If
If cmbTType.Text = "Due" Then
cmdAdd.Visible = False
cmdSave.Visible = True
cmdAdd.Enabled = False
cmdSave.Enabled = True
cmdSave.SetFocus
End If
End If

End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
    If cmbLType.Text = "" Then
       Exit Sub
    End If
     
    If cmbLType.Text = "Casual Leave" Then
        
        Set rs = New ADODB.Recordset
        str = "select * from Leave_Status where ID like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
        rs!CL_Availed = Val(txtCLavailed.Text) + Val(txtDay.Text)
        rs!EL_Availed = Val(txtELavailed.Text)
        rs!SL_Availed = Val(txtSLavailed.Text)
        rs!Half_Leave = Val(txtHalf.Text)
        rs!Without_Leave = Val(txtWithout.Text)
        rs!CL_Balance = Val(txtCLBalance.Text) - Val(txtDay.Text)
        rs!EL_Balance = Val(txtELBalance.Text)
        rs!SL_Balance = Val(txtSLBalance.Text)
        rs!ML_Balance = Val(txtMLBalance.Text)
        rs!Exl_Balance = Val(txtExLBalance.Text)
        
             
        Set rsU = New ADODB.Recordset
        str = "select Trn_no from Others where Trn_no like '" & Val(txtTran.Text) - 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        
        rsU!Trn_No = Val(txtTran.Text)
     
        Set rsN = New ADODB.Recordset
        rsN.Open "Leave_Record", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
    
        rsN!ID = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesignation.Text
        rsN!Grade = cmbGrade.Text
        rsN!Bank_Joining = txtJoinbk.Text
        rsN!Branch_Joining = txtJoinBr.Text
        rsN!From_Date = txtfrom.Text
        rsN!To_Date = txtTo.Text
        rsN!CL_Due = 0
        rsN!CL_Taken = Val(txtDay.Text)
        rsN!CL_Balance = Val(txtCLBalance.Text) - Val(txtDay.Text)
        rsN!EL_Due = 0
        rsN!EL_Taken = 0
        rsN!EL_Balance = Val(txtELBalance.Text)
        rsN!SL_Due = 0
        rsN!SL_Taken = 0
        rsN!SL_Balance = Val(txtSLBalance.Text)
        rsN!ML_Due = 0
        rsN!ML_Taken = 0
        rsN!ML_Balance = Val(txtMLBalance.Text)
        rsN!Exl_Due = 0
        rsN!Exl_Taken = 0
        rsN!Exl_Balance = Val(txtExLBalance.Text)
        rsN!Trn_No = Val(txtTran.Text)
        rsN.Update
        rsN.Close
        rs.Update
        rs.Close
        rsU.Update
        rsU.Close
        End If
    End If
    
    
    If cmbLType.Text = "Earn Leave" Then
        
        Set rs = New ADODB.Recordset
        str = "select * from Leave_Status where ID like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
        rs!CL_Availed = Val(txtCLavailed.Text)
        rs!EL_Availed = Val(txtELavailed.Text) + Val(txtDay.Text)
        rs!SL_Availed = Val(txtSLavailed.Text)
        rs!Half_Leave = Val(txtHalf.Text)
        rs!Without_Leave = Val(txtWithout.Text)
        rs!CL_Balance = Val(txtCLBalance.Text)
        rs!EL_Balance = Val(txtELBalance.Text) - Val(txtDay.Text)
        rs!SL_Balance = Val(txtSLBalance.Text)
        rs!ML_Balance = Val(txtMLBalance.Text)
        rs!Exl_Balance = Val(txtExLBalance.Text)
        
        
        Set rsU = New ADODB.Recordset
        str = "select Trn_no from Others where Trn_no like '" & Val(txtTran.Text) - 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        
        rsU!Trn_No = Val(txtTran.Text)
     
        Set rsN = New ADODB.Recordset
        rsN.Open "Leave_Record", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
    
        rsN!ID = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesignation.Text
        rsN!Grade = cmbGrade.Text
        rsN!Bank_Joining = txtJoinbk.Text
        rsN!Branch_Joining = txtJoinBr.Text
        rsN!From_Date = txtfrom.Text
        rsN!To_Date = txtTo.Text
        rsN!CL_Due = 0
        rsN!CL_Taken = 0
        rsN!CL_Balance = Val(txtCLBalance.Text)
        rsN!EL_Due = 0
        rsN!EL_Taken = Val(txtDay.Text)
        rsN!EL_Balance = Val(txtELBalance.Text) - Val(txtDay.Text)
        rsN!SL_Due = 0
        rsN!SL_Taken = 0
        rsN!SL_Balance = Val(txtSLBalance.Text)
        rsN!ML_Due = 0
        rsN!ML_Taken = 0
        rsN!ML_Balance = Val(txtMLBalance.Text)
        rsN!Exl_Due = 0
        rsN!Exl_Taken = 0
        rsN!Exl_Balance = Val(txtExLBalance.Text)
        rsN!Trn_No = Val(txtTran.Text)
        rsN.Update
        rsN.Close
        rs.Update
        rs.Close
        rsU.Update
        rsU.Close
        End If
    End If
    
If cmbLType.Text = "Sick Leave" Then
        
        Set rs = New ADODB.Recordset
        str = "select * from Leave_Status where ID like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
        rs!CL_Availed = Val(txtCLavailed.Text)
        rs!EL_Availed = Val(txtELavailed.Text)
        rs!SL_Availed = Val(txtSLavailed.Text) + Val(txtDay.Text)
        rs!Half_Leave = Val(txtHalf.Text)
        rs!Without_Leave = Val(txtWithout.Text)
        rs!CL_Balance = Val(txtCLBalance.Text)
        rs!EL_Balance = Val(txtELBalance.Text)
        rs!SL_Balance = Val(txtSLBalance.Text) - Val(txtDay.Text)
        rs!ML_Balance = Val(txtMLBalance.Text)
        rs!Exl_Balance = Val(txtExLBalance.Text)
        
        
        Set rsU = New ADODB.Recordset
        str = "select Trn_no from Others where Trn_no like '" & Val(txtTran.Text) - 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        
        rsU!Trn_No = Val(txtTran.Text)
     
        Set rsN = New ADODB.Recordset
        rsN.Open "Leave_Record", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
    
        rsN!ID = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesignation.Text
        rsN!Grade = cmbGrade.Text
        rsN!Bank_Joining = txtJoinbk.Text
        rsN!Branch_Joining = txtJoinBr.Text
        rsN!From_Date = txtfrom.Text
        rsN!To_Date = txtTo.Text
        rsN!CL_Due = 0
        rsN!CL_Taken = 0
        rsN!CL_Balance = Val(txtCLBalance.Text)
        rsN!EL_Due = 0
        rsN!EL_Taken = 0
        rsN!EL_Balance = Val(txtELBalance.Text)
        rsN!SL_Due = 0
        rsN!SL_Taken = Val(txtDay.Text)
        rsN!SL_Balance = Val(txtSLBalance.Text) - Val(txtDay.Text)
        rsN!ML_Due = 0
        rsN!ML_Taken = 0
        rsN!ML_Balance = Val(txtMLBalance.Text)
        rsN!Exl_Due = 0
        rsN!Exl_Taken = 0
        rsN!Exl_Balance = Val(txtExLBalance.Text)
        rsN!Trn_No = Val(txtTran.Text)
        rsN.Update
        rsN.Close
        rs.Update
        rs.Close
        rsU.Update
        rsU.Close
        End If
    End If
    
If cmbLType.Text = "Ex-Leave" Or cmbLType.Text = "Leave Fair Assistance" Then
        
        Set rs = New ADODB.Recordset
        str = "select * from Leave_Status where ID like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
        rs!CL_Availed = Val(txtCLavailed.Text)
        rs!EL_Availed = Val(txtELavailed.Text)
        rs!SL_Availed = Val(txtSLavailed.Text)
        rs!Half_Leave = Val(txtHalf.Text)
        rs!Without_Leave = Val(txtWithout.Text)
        rs!CL_Balance = Val(txtCLBalance.Text)
        rs!EL_Balance = Val(txtELBalance.Text)
        rs!SL_Balance = Val(txtSLBalance.Text)
        rs!ML_Balance = Val(txtMLBalance.Text)
        rs!Exl_Balance = Val(txtExLBalance.Text) - Val(txtDay.Text)
        
        
        Set rsU = New ADODB.Recordset
        str = "select Trn_no from Others where Trn_no like '" & Val(txtTran.Text) - 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        
        rsU!Trn_No = Val(txtTran.Text)
     
        Set rsN = New ADODB.Recordset
        rsN.Open "Leave_Record", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
    
        rsN!ID = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesignation.Text
        rsN!Grade = cmbGrade.Text
        rsN!Bank_Joining = txtJoinbk.Text
        rsN!Branch_Joining = txtJoinBr.Text
        rsN!From_Date = txtfrom.Text
        rsN!To_Date = txtTo.Text
        rsN!CL_Due = 0
        rsN!CL_Taken = 0
        rsN!CL_Balance = Val(txtCLBalance.Text)
        rsN!EL_Due = 0
        rsN!EL_Taken = 0
        rsN!EL_Balance = Val(txtELBalance.Text)
        rsN!SL_Due = 0
        rsN!SL_Taken = 0
        rsN!SL_Balance = Val(txtSLBalance.Text)
        rsN!ML_Due = 0
        rsN!ML_Taken = 0
        rsN!ML_Balance = Val(txtMLBalance.Text)
        rsN!Exl_Due = 0
        rsN!Exl_Taken = Val(txtDay.Text)
        rsN!Exl_Balance = Val(txtExLBalance.Text) - Val(txtDay.Text)
        rsN!Trn_No = Val(txtTran.Text)
        rsN.Update
        rsN.Close
        rs.Update
        rs.Close
        rsU.Update
        rsU.Close
        End If
    End If
    txtTran.Text = Val(txtTran.Text) + 1
    Call clearTextboxes
    cmdAdd.Enabled = False
    txtID.SetFocus
    Exit Sub
End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    If cmbLType.Text = "" Then
       Exit Sub
    End If
     
    If cmbLType.Text = "Casual Leave" Then
        
        Set rs = New ADODB.Recordset
        str = "select * from Leave_Status where ID like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
        rs!CL_Availed = Val(txtCLavailed.Text)
        rs!EL_Availed = Val(txtELavailed.Text)
        rs!SL_Availed = Val(txtSLavailed.Text)
        rs!Half_Leave = Val(txtHalf.Text)
        rs!Without_Leave = Val(txtWithout.Text)
        rs!CL_Balance = Val(txtCLBalance.Text) + Val(txtDay.Text)
        rs!EL_Balance = Val(txtELBalance.Text)
        rs!SL_Balance = Val(txtSLBalance.Text)
        rs!ML_Balance = Val(txtMLBalance.Text)
        rs!Exl_Balance = Val(txtExLBalance.Text)
        
    
        Set rsU = New ADODB.Recordset
        str = "select Trn_no from Others where Trn_no like '" & Val(txtTran.Text) - 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        
        rsU!Trn_No = Val(txtTran.Text)
     
        Set rsN = New ADODB.Recordset
        rsN.Open "Leave_Record", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
    
        rsN!ID = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesignation.Text
        rsN!Grade = cmbGrade.Text
        rsN!Bank_Joining = txtJoinbk.Text
        rsN!Branch_Joining = txtJoinBr.Text
        rsN!From_Date = txtfrom.Text
        rsN!To_Date = txtTo.Text
        rsN!CL_Due = Val(txtDay.Text)
        rsN!CL_Taken = 0
        rsN!CL_Balance = Val(txtCLBalance.Text) + Val(txtDay.Text)
        rsN!EL_Due = 0
        rsN!EL_Taken = 0
        rsN!EL_Balance = Val(txtELBalance.Text)
        rsN!SL_Due = 0
        rsN!SL_Taken = 0
        rsN!SL_Balance = Val(txtSLBalance.Text)
        rsN!ML_Due = 0
        rsN!ML_Taken = 0
        rsN!ML_Balance = Val(txtMLBalance.Text)
        rsN!Exl_Due = 0
        rsN!Exl_Taken = 0
        rsN!Exl_Balance = Val(txtExLBalance.Text)
        rsN!Trn_No = Val(txtTran.Text)
        rsN.Update
        rsN.Close
        rs.Update
        rs.Close
        rsU.Update
        rsU.Close
        End If
    End If
    
    
    If cmbLType.Text = "Earn Leave" Then
        
        Set rs = New ADODB.Recordset
        str = "select * from Leave_Status where ID like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
        rs!CL_Availed = Val(txtCLavailed.Text)
        rs!EL_Availed = Val(txtELavailed.Text)
        rs!SL_Availed = Val(txtSLavailed.Text)
        rs!Half_Leave = Val(txtHalf.Text)
        rs!Without_Leave = Val(txtWithout.Text)
        rs!CL_Balance = Val(txtCLBalance.Text)
        rs!EL_Balance = Val(txtELBalance.Text) + Val(txtDay.Text)
        rs!SL_Balance = Val(txtSLBalance.Text)
        rs!ML_Balance = Val(txtMLBalance.Text)
        rs!Exl_Balance = Val(txtExLBalance.Text)
        
        
        Set rsU = New ADODB.Recordset
        str = "select Trn_no from Others where Trn_no like '" & Val(txtTran.Text) - 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        
        rsU!Trn_No = Val(txtTran.Text)
     
        Set rsN = New ADODB.Recordset
        rsN.Open "Leave_Record", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
    
        rsN!ID = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesignation.Text
        rsN!Grade = cmbGrade.Text
        rsN!Bank_Joining = txtJoinbk.Text
        rsN!Branch_Joining = txtJoinBr.Text
        rsN!From_Date = txtfrom.Text
        rsN!To_Date = txtTo.Text
        rsN!CL_Due = 0
        rsN!CL_Taken = 0
        rsN!CL_Balance = Val(txtCLBalance.Text)
        rsN!EL_Due = Val(txtDay.Text)
        rsN!EL_Taken = 0
        rsN!EL_Balance = Val(txtELBalance.Text) + Val(txtDay.Text)
        rsN!SL_Due = 0
        rsN!SL_Taken = 0
        rsN!SL_Balance = Val(txtSLBalance.Text)
        rsN!ML_Due = 0
        rsN!ML_Taken = 0
        rsN!ML_Balance = Val(txtMLBalance.Text)
        rsN!Exl_Due = 0
        rsN!Exl_Taken = 0
        rsN!Exl_Balance = Val(txtExLBalance.Text)
        rsN!Trn_No = Val(txtTran.Text)
        rsN.Update
        rsN.Close
        rs.Update
        rs.Close
        rsU.Update
        rsU.Close
        End If
    End If
    
If cmbLType.Text = "Sick Leave" Then
        
        Set rs = New ADODB.Recordset
        str = "select * from Leave_Status where ID like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
        rs!CL_Availed = Val(txtCLavailed.Text)
        rs!EL_Availed = Val(txtELavailed.Text)
        rs!SL_Availed = Val(txtSLavailed.Text)
        rs!Half_Leave = Val(txtHalf.Text)
        rs!Without_Leave = Val(txtWithout.Text)
        rs!CL_Balance = Val(txtCLBalance.Text)
        rs!EL_Balance = Val(txtELBalance.Text)
        rs!SL_Balance = Val(txtSLBalance.Text) + Val(txtDay.Text)
        rs!ML_Balance = Val(txtMLBalance.Text)
        rs!Exl_Balance = Val(txtExLBalance.Text)
        
        
        Set rsU = New ADODB.Recordset
        str = "select Trn_no from Others where Trn_no like '" & Val(txtTran.Text) - 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        
        rsU!Trn_No = Val(txtTran.Text)
     
        Set rsN = New ADODB.Recordset
        rsN.Open "Leave_Record", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
    
        
        rsN!ID = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesignation.Text
        rsN!Grade = cmbGrade.Text
        rsN!Bank_Joining = txtJoinbk.Text
        rsN!Branch_Joining = txtJoinBr.Text
        rsN!From_Date = txtfrom.Text
        rsN!To_Date = txtTo.Text
        rsN!CL_Due = 0
        rsN!CL_Taken = 0
        rsN!CL_Balance = Val(txtCLBalance.Text)
        rsN!EL_Due = 0
        rsN!EL_Taken = 0
        rsN!EL_Balance = Val(txtELBalance.Text)
        rsN!SL_Due = Val(txtDay.Text)
        rsN!SL_Taken = 0
        rsN!SL_Balance = Val(txtSLBalance.Text) + Val(txtDay.Text)
        rsN!ML_Due = 0
        rsN!ML_Taken = 0
        rsN!ML_Balance = Val(txtMLBalance.Text)
        rsN!Exl_Due = 0
        rsN!Exl_Taken = 0
        rsN!Exl_Balance = Val(txtExLBalance.Text)
        rsN!Trn_No = Val(txtTran.Text)
        rsN.Update
        rsN.Close
        rs.Update
        rs.Close
        rsU.Update
        rsU.Close
        End If
    End If
    
If cmbLType.Text = "Ex-Leave" Or cmbLType.Text = "Leave Fair Assistance" Then
        
        Set rs = New ADODB.Recordset
        str = "select * from Leave_Status where ID like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
        rs!CL_Availed = Val(txtCLavailed.Text)
        rs!EL_Availed = Val(txtELavailed.Text)
        rs!SL_Availed = Val(txtSLavailed.Text)
        rs!Half_Leave = Val(txtHalf.Text)
        rs!Without_Leave = Val(txtWithout.Text)
        rs!CL_Balance = Val(txtCLBalance.Text)
        rs!EL_Balance = Val(txtELBalance.Text)
        rs!SL_Balance = Val(txtSLBalance.Text)
        rs!ML_Balance = Val(txtMLBalance.Text)
        rs!Exl_Balance = Val(txtExLBalance.Text) + Val(txtDay.Text)
        
        Set rsU = New ADODB.Recordset
        str = "select Trn_no from Others where Trn_no like '" & Val(txtTran.Text) - 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        
        rsU!Trn_No = Val(txtTran.Text)
     
        Set rsN = New ADODB.Recordset
        rsN.Open "Leave_Record", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
    
        rsN!ID = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = cmbDesignation.Text
        rsN!Grade = cmbGrade.Text
        rsN!Bank_Joining = txtJoinbk.Text
        rsN!Branch_Joining = txtJoinBr.Text
        rsN!From_Date = txtfrom.Text
        rsN!To_Date = txtTo.Text
        rsN!CL_Due = 0
        rsN!CL_Taken = 0
        rsN!CL_Balance = Val(txtCLBalance.Text)
        rsN!EL_Due = 0
        rsN!EL_Taken = 0
        rsN!EL_Balance = Val(txtELBalance.Text)
        rsN!SL_Due = 0
        rsN!SL_Taken = 0
        rsN!SL_Balance = Val(txtSLBalance.Text)
        rsN!ML_Due = 0
        rsN!ML_Taken = 0
        rsN!ML_Balance = Val(txtMLBalance.Text)
        rsN!Exl_Due = Val(txtDay.Text)
        rsN!Exl_Taken = 0
        rsN!Exl_Balance = Val(txtExLBalance.Text) + Val(txtDay.Text)
        rsN!Trn_No = Val(txtTran.Text)
        rsN.Update
        rsN.Close
        rs.Update
        rs.Close
        rsU.Update
        rsU.Close
        End If
    End If
    Call clearTextboxes
    cmdSave.Enabled = False
    txtID.SetFocus
    Exit Sub

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Last
    
    Set rs = New ADODB.Recordset
        str = "select * from Leave_Status"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
        
  Set rsU = New ADODB.Recordset
        str = "select Trn_no from Others"
        rsU.Open str, conn
        txtTran.Text = Val(rsU!Trn_No) + 1
        rsU.Close
        
        Call clearTextboxes
        cmdAdd.Visible = True
        cmdAdd.Enabled = False
        cmdSave.Visible = False
        cmdSave.Enabled = False
    Exit Sub
Last:
    MsgBox ("Database Connection error: " + Err.Description)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTo.SetFocus
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtfrom.SetFocus
End If
End Sub

Private Sub txtID_LostFocus()
If txtID.Text = "" Then
Exit Sub
End If

Dim ID As String
    ID = txtID.Text
    
     Set rs = New ADODB.Recordset
        str = "select * from Leave_Status where ID like '" & ID & "'"
        rs.Open str, conn
    
     
    If Not rs.EOF Then
    Call clearTextboxes
    On Error Resume Next
        txtID.Text = rs!ID
        txtName.Text = rs!Name
        cmbDesignation.Text = rs!Designation
        cmbGrade.Text = rs!Grade
        txtJoinbk.Text = rs!Bank_Joining
        txtJoinBr.Text = rs!Branch_Joining
        txtCLavailed.Text = rs!CL_Availed
        txtELavailed.Text = rs!EL_Availed
        txtSLavailed.Text = rs!SL_Availed
        txtHalf.Text = rs!Half_Leave
        txtWithout.Text = rs!Without_Leave
        txtCLBalance.Text = rs!CL_Balance
        txtELBalance.Text = rs!EL_Balance
        txtSLBalance.Text = rs!SL_Balance
        txtMLBalance.Text = rs!ML_Balance
        txtExLBalance.Text = rs!Exl_Balance
        rs.Close
        
    Else
    MsgBox "There is no such ID found, Add employee first?", vbCritical + vbYesNo
        Call clearTextboxes
        txtID.Text = ID
        rs.Close
   
      End If
    Exit Sub
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim fromDate As Date
Dim toDate As Date
fromDate = txtfrom.Text
toDate = txtTo.Text
txtDay.Text = (toDate - fromDate) + 1
cmbLType.SetFocus
End If
End Sub
