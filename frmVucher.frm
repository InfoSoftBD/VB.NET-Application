VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVoucher 
   BackColor       =   &H00008000&
   Caption         =   "Voucher"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13155
   Icon            =   "frmVucher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   13155
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   7455
      Left            =   8400
      ScaleHeight     =   7395
      ScaleWidth      =   4515
      TabIndex        =   33
      Top             =   120
      Width           =   4575
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6975
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   12303
         _Version        =   393216
         BackColor       =   -2147483629
         HeadLines       =   1
         RowHeight       =   19
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9120
      Top             =   7200
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
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0FFC0&
      Height          =   780
      Left            =   150
      ScaleHeight     =   720
      ScaleWidth      =   8010
      TabIndex        =   3
      Top             =   6840
      Width           =   8070
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   435
         Left            =   6675
         TabIndex        =   8
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   435
         Left            =   330
         TabIndex        =   7
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   435
         Left            =   1830
         TabIndex        =   6
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Print"
         Height          =   435
         Left            =   5085
         TabIndex        =   5
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Update"
         Height          =   435
         Left            =   3435
         TabIndex        =   4
         Top             =   150
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0FFC0&
      Height          =   5565
      Left            =   150
      ScaleHeight     =   5505
      ScaleWidth      =   7995
      TabIndex        =   2
      Top             =   1110
      Width           =   8055
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Initial Inforamtion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   7740
         Begin VB.TextBox txtPs 
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
            Left            =   7230
            MaxLength       =   2
            TabIndex        =   29
            Text            =   "00"
            Top             =   300
            Width           =   375
         End
         Begin VB.TextBox txtDate 
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
            Left            =   645
            TabIndex        =   28
            Top             =   300
            Width           =   1410
         End
         Begin VB.TextBox txtTk 
            Alignment       =   1  'Right Justify
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
            Left            =   5955
            TabIndex        =   27
            Top             =   300
            Width           =   1290
         End
         Begin VB.ComboBox cmbVoucher 
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
            Left            =   3585
            TabIndex        =   26
            Text            =   "Single Debit"
            Top             =   307
            Width           =   1680
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Taka"
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
            TabIndex        =   32
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   165
            TabIndex        =   31
            Top             =   360
            Width           =   435
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voucher Type"
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
            Left            =   2190
            TabIndex        =   30
            Top             =   360
            Width           =   1320
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "A/C Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   120
         TabIndex        =   20
         Top             =   990
         Width           =   7740
         Begin VB.TextBox txtTitle 
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
            Left            =   3855
            TabIndex        =   22
            Top             =   255
            Width           =   3735
         End
         Begin VB.ComboBox txtCode 
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
            Left            =   915
            TabIndex        =   21
            Top             =   262
            Width           =   1950
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A/C Title"
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
            Left            =   2940
            TabIndex        =   24
            Top             =   315
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A/C No."
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
            TabIndex        =   23
            Top             =   322
            Width           =   705
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Contra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   135
         TabIndex        =   15
         Top             =   4620
         Width           =   7710
         Begin VB.TextBox txtTran 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6240
            TabIndex        =   35
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cmbContra 
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
            Left            =   900
            TabIndex        =   19
            Top             =   262
            Width           =   1815
         End
         Begin VB.TextBox txtContra 
            Height          =   375
            Left            =   3690
            TabIndex        =   16
            Top             =   255
            Width           =   2310
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A/C Title"
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
            Left            =   2805
            TabIndex        =   18
            Top             =   322
            Width           =   810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A/C No."
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
            Left            =   135
            TabIndex        =   17
            Top             =   322
            Width           =   705
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Amount in word"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   135
         TabIndex        =   13
         Top             =   3600
         Width           =   7710
         Begin VB.TextBox txtStr 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   105
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   285
            Width           =   7485
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Naration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   135
         TabIndex        =   9
         Top             =   1845
         Width           =   7710
         Begin VB.TextBox txtNaration 
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
            TabIndex        =   12
            Top             =   330
            Width           =   7470
         End
         Begin VB.TextBox txtNaration1 
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
            TabIndex        =   11
            Top             =   750
            Width           =   7470
         End
         Begin VB.TextBox txtNaration2 
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
            TabIndex        =   10
            Top             =   1155
            Width           =   7470
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   150
      ScaleHeight     =   795
      ScaleWidth      =   8010
      TabIndex        =   0
      Top             =   120
      Width           =   8070
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VOUCHER GENERATOR"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   780
         TabIndex        =   1
         Top             =   90
         Width           =   6225
      End
   End
End
Attribute VB_Name = "frmVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim StrD As String
Dim Value As Double
Dim Result As Double
Dim str As String
Dim StrPs As String
Dim Str1 As String
Dim Str10 As String
Dim Str100 As String
Dim Str1000 As String
Dim Str100000 As String
Dim Str10000000 As String
Dim Tk As String
Dim Ps As String
Dim Only As String
Dim n0, n1, n2, n3, n4, n5, n6, n7, n8, n9, n10 As String
Dim n11, n12, n13, n14, n15, n16, n17, n18, n19, n20 As String
Dim n30, n40, n50, n60, n70, n80, n90, n100, n1000, n100000, n10000000 As String
Private Sub Case_Result()
    Value = Result
    n0 = "Zero ": n1 = "One ": n2 = "Two ": n3 = "Three ": n4 = "Four "
    n5 = "Five ": n6 = "Six ": n7 = "Seven ": n8 = "Eight ": n9 = "Nine "
    n8 = "Eight ": n9 = "Nine ": n10 = "Ten ": n11 = "Eleven ": n12 = "Twelve ": n13 = "Thirteen "
    n14 = "Fourteen ": n15 = "Fifteen ": n16 = "Sixteen ": n17 = "Seventeen ": n18 = "Eighteen ": n19 = "Nineteen "
    n20 = "Twenty ": n30 = "Thirty ": n40 = "Forty ": n50 = "Fifty ": n60 = "Sixty ": n70 = "Seventy "
    n80 = "Eighty ": n90 = "Ninety ": n100 = "Hundred ": n1000 = "Thousand ": n100000 = "Lakh ": n10000000 = "Crore "
    Tk = "Taka ": Ps = "Paisa ": Only = "Only"
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
Private Sub Case_Ps()
    If Value >= 1 And Value < 100 Then
        Result = Value
        Call Case_Result
        StrPs = str
        txtStr.Text = Ps + StrPs + Only
    End If
 End Sub
Private Sub Case0_9()
         Result = Value
    Call Case_Result
        Str1 = str
    If Val(txtPs.Text) > 0 Then
         Value = Val(txtPs.Text)
    Call Case_Ps
        Str1 = Str1 + "and " + Ps + StrPs
        txtStr.Text = Tk + Str1 + Only
    Else
        txtStr.Text = Tk + Str1 + Only
    End If
 End Sub
Private Sub Case10_99()
    Result = Value
    Call Case_Result
        Str10 = str
    If Val(txtPs.Text) > 0 Then
         Value = Val(txtPs.Text)
    Call Case_Ps
        Str10 = Str10 + "and " + Ps + StrPs
        txtStr.Text = Tk + Str10 + Only
    Else
        txtStr.Text = Tk + Str10 + Only
    End If
End Sub
Private Sub Case100_999()
    Dim Mode, Div  As Double
        Mode = Value - 100 * (Int(Value / 100))
        Div = (Value - Mode) / 100
        Result = Div
    Call Case_Result
        Str100 = str + n100
    
    If Mode > 0 Then
        Result = Mode
        Call Case_Result
        Str100 = Str100 + str
            If Val(txtPs.Text) > 0 Then
                Value = Val(txtPs.Text)
                Call Case_Ps
                Str100 = Str100 + "and " + Ps + StrPs
                txtStr.Text = Tk + Str100 + Only
            Else
                txtStr.Text = Tk + Str100 + Only
            End If
        Else
            If Val(txtPs.Text) > 0 Then
                Value = Val(txtPs.Text)
                Call Case_Ps
                Str100 = Str100 + "and " + Ps + StrPs
                txtStr.Text = Tk + Str100 + Only
            Else
                txtStr.Text = Tk + Str100 + Only
        End If
    End If
End Sub
Private Sub Case1000_99999()
Dim Mode, Div  As Double
        Mode = Value - 1000 * (Int(Value / 1000))
        Div = (Value - Mode) / 1000
        Result = Div
    Call Case_Result
        Str1000 = str + n1000
    If Mode > 0 Then
        If Mode >= 1 And Mode < 10 Then
           Value = Mode
           Call Case0_9
           Str1000 = Str1000 + Str1
           txtStr.Text = Tk + Str1000 + Only
        End If
        
        If Mode >= 10 And Mode < 100 Then
            Value = Mode
            Call Case10_99
            Str1000 = Str1000 + Str10
            txtStr.Text = Tk + Str1000 + Only
        End If
        
        If Mode >= 100 And Mode < 1000 Then
            Value = Mode
            Call Case100_999
            Str1000 = Str1000 + Str100
            txtStr.Text = Tk + Str1000 + Only
        End If
    Else
        If Val(txtPs.Text) > 0 Then
           Value = Val(txtPs.Text)
           Call Case_Ps
           Str1000 = Str1000 + "and " + Ps + StrPs
           txtStr.Text = Tk + Str1000 + Only
        Else
           txtStr.Text = Tk + Str1000 + Only
        End If
    End If
End Sub
Private Sub Case100000_9999999()
Dim Mode, Div  As Double
        Mode = Value - 100000 * (Int(Value / 100000))
        Div = (Value - Mode) / 100000
        Result = Div
    Call Case_Result
        Str100000 = str + n100000
    If Mode > 0 Then
        If Mode >= 1 And Mode < 10 Then
            Value = Mode
            Call Case0_9
            Str100000 = Str100000 + Str1
            txtStr.Text = Tk + Str100000 + Only
        End If
        
        If Mode >= 10 And Mode < 100 Then
            Value = Mode
            Call Case10_99
            Str100000 = Str100000 + Str10
            txtStr.Text = Tk + Str100000 + Only
        End If
        
        If Mode >= 100 And Mode < 1000 Then
            Value = Mode
            Call Case100_999
            Str100000 = Str100000 + Str100
            txtStr.Text = Tk + Str100000 + Only
        End If
        
        If Mode >= 1000 And Mode < 100000 Then
            Value = Mode
            Call Case1000_99999
            Str100000 = Str100000 + Str1000
            txtStr.Text = Tk + Str100000 + Only
        End If
    Else
        If Val(txtPs.Text) > 0 Then
            Value = Val(txtPs.Text)
            Call Case_Ps
            Str100000 = Str100000 + "and " + Ps + StrPs
            txtStr.Text = Tk + Str100000 + Only
        Else
           txtStr.Text = Tk + Str100000 + Only
        End If
    End If
End Sub
Private Sub Case10000000_999999999()
Dim Mode, Div  As Double
        Mode = Value - 10000000 * (Int(Value / 10000000))
        Div = (Value - Mode) / 10000000
        Result = Div
    Call Case_Result
        Str10000000 = str + n10000000
    If Mode > 0 Then
        If Mode >= 1 And Mode < 10 Then
            Value = Mode
            Call Case0_9
            Str10000000 = Str10000000 + Str1
            txtStr.Text = Tk + Str10000000 + Only
        End If
        
        If Mode >= 10 And Mode < 100 Then
            Value = Mode
            Call Case10_99
            Str10000000 = Str10000000 + Str10
            txtStr.Text = Tk + Str10000000 + Only
        End If
        
        If Mode >= 100 And Mode < 1000 Then
            Value = Mode
            Call Case100_999
            Str10000000 = Str10000000 + Str100
            txtStr.Text = Tk + Str10000000 + Only
        End If
        
        If Mode >= 1000 And Mode < 100000 Then
            Value = Mode
            Call Case1000_99999
            Str10000000 = Str10000000 + Str1000
            txtStr.Text = Tk + Str10000000 + Only
        End If
                
        If Mode >= 100000 And Mode < 10000000 Then
            Value = Mode
            Call Case100000_9999999
            Str10000000 = Str10000000 + Str100000
            txtStr.Text = Tk + Str10000000 + Only
        End If
    Else
        If Val(txtPs.Text) > 0 Then
            Value = Val(txtPs.Text)
            Call Case_Ps
            Str10000000 = Str10000000 + "and " + Ps + StrPs
            txtStr.Text = Tk + Str10000000 + Only
        Else
           txtStr.Text = Tk + Str10000000 + Only
        End If
    End If
End Sub

Private Sub cmbContra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtContra.SelStart = 0
    txtContra.SelLength = Len(txtContra.Text)
    txtContra.SetFocus
End If
End Sub

Private Sub cmbVoucher_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTk.SelStart = 0
    txtTk.SelLength = Len(txtTk.Text)
    txtTk.SetFocus
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Set rsN = New ADODB.Recordset
        rsN.Open "Journal", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
    If cmbVoucher.Text = "" Then
        MsgBox "Please Select Voucher Type", vbInformation, "Error"
        cmbVoucher.SetFocus
    Else
        If cmbVoucher.Text = "Single Debit" Then
        rsN!Voucher_Type = "DEBIT"
        End If
        If cmbVoucher.Text = "Single Credit" Then
        rsN!Voucher_Type = "CREDIT"
        End If
        If cmbVoucher.Text = "Cash Debit" Then
        rsN!Voucher_Type = "Cash Debit"
        End If
        If cmbVoucher.Text = "Party Debit" Then
        rsN!Voucher_Type = "Party Debit"
        End If
        If cmbVoucher.Text = "Party Credit" Then
        rsN!Voucher_Type = "Party Credit"
        End If
        rsN!Date = txtDate.Text
        rsN!Tk = txtTk.Text
        rsN!Ps = txtPs.Text
        rsN!AC_Code = txtCode.Text
        rsN!AC_Title = txtTitle.Text
        rsN!Naration = txtNaration.Text
        rsN!Naration1 = txtNaration1.Text
        rsN!Naration2 = txtNaration2.Text
        rsN!String = txtStr.Text
        rsN!Contra_Code = cmbContra.Text
        rsN!Contra_Title = txtContra.Text
        rsN.Update
        rsN.Close
    End If
    
    Set rs = New ADODB.Recordset
        str = "select * from Journal Order by ID Asc"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
        
    Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT AC_Code FROM Journal"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtCode.AddItem rs!AC_Code
        rs.MoveNext
        Loop
        rs.Close
    Exit Sub
    Resume Next
End Sub

Private Sub Command3_Click()
On Error Resume Next
 
 
 str = "select * from Journal where ID like '" & txtTran.Text & "'"
    Set rs = New ADODB.Recordset
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
    On Error Resume Next
        
        rs.Close
     
     If MsgBox("Really want to delete?", vbCritical + vbYesNo) = vbYes Then
        
        str = "delete from Journal where ID like '" & txtTran.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.Close
    
        
        str = "select * from Journal"
        rs.Open str, conn, adOpenKeyset, adLockReadOnly
        
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        
        End If
    Else
        MsgBox "There is no such Transaction no. found: " & txtTran.Text
        rs.Close
    End If
    Exit Sub
    Resume Next
End Sub

Private Sub Command4_Click()
On Error Resume Next
 
 str = "select * from Voucher where Scroll like '" & 1 & "'"
    Set rs = New ADODB.Recordset
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
    On Error Resume Next
        
        rs.Close
                 
        str = "delete from Voucher where Scroll like '" & 1 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.Close
    'Call clearTextboxes
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Voucher", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
    If cmbVoucher.Text = "" Then
        MsgBox "Please Select Voucher Type", vbInformation, "Error"
        cmbVoucher.SetFocus
    Else
        If cmbVoucher.Text = "Single Debit" Then
        rsN!Voucher_Type = "DEBIT"
        End If
        If cmbVoucher.Text = "Single Credit" Then
        rsN!Voucher_Type = "CREDIT"
        End If
        If cmbVoucher.Text = "Cash Debit" Then
        rsN!Voucher_Type = "Cash Debit"
        End If
        If cmbVoucher.Text = "Party Debit" Then
        rsN!Voucher_Type = "Party Debit"
        End If
        If cmbVoucher.Text = "Party Credit" Then
        rsN!Voucher_Type = "Party Credit"
        End If
        rsN!Date = txtDate.Text
        rsN!Tk = txtTk.Text
        rsN!Ps = txtPs.Text
        rsN!AC_Code = txtCode.Text
        rsN!AC_Title = txtTitle.Text
        rsN!Naration = txtNaration.Text
        rsN!Naration1 = txtNaration1.Text
        rsN!Naration2 = txtNaration2.Text
        rsN!String = txtStr.Text
        rsN!Contra_Code = cmbContra.Text
        rsN!Contra_Title = txtContra.Text
        rsN!Scroll = 1
        rsN.Update
    End If
    End If
    Unload Me
   ' frmVoucher_Print.Show 1
Exit Sub
    Resume Next

End Sub

Private Sub Command5_Click()
On Error Resume Next


Set rsN = New ADODB.Recordset
    str = "select * from Journal where ID like '" & txtTran.Text & "'"
    rsN.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsN.EOF Then
        If cmbVoucher.Text = "" Then
        MsgBox "Please Select Voucher Type", vbInformation, "Error"
        cmbVoucher.SetFocus
        End If
        If cmbVoucher.Text = "Single Debit" Then
        rsN!Voucher_Type = "DEBIT"
        End If
        If cmbVoucher.Text = "Single Credit" Then
        rsN!Voucher_Type = "CREDIT"
        End If
        If cmbVoucher.Text = "Cash Debit" Then
        rsN!Voucher_Type = "Cash Debit"
        End If
        If cmbVoucher.Text = "Party Debit" Then
        rsN!Voucher_Type = "Party Debit"
        End If
        If cmbVoucher.Text = "Party Credit" Then
        rsN!Voucher_Type = "Party Credit"
        End If
        rsN!Date = txtDate.Text
        rsN!Tk = txtTk.Text
        rsN!Ps = txtPs.Text
        rsN!AC_Code = txtCode.Text
        rsN!AC_Title = txtTitle.Text
        rsN!Naration = txtNaration.Text
        rsN!Naration1 = txtNaration1.Text
        rsN!Naration2 = txtNaration2.Text
        rsN!String = txtStr.Text
        rsN!Contra_Code = cmbContra.Text
        rsN!Contra_Title = txtContra.Text
        rsN.Update
    Else
        MsgBox "There is no such Voucher No. found.", 64, "Update Error"
        rsU.Close
        Exit Sub
    End If
        
    Set rs = New ADODB.Recordset
        str = "select * from Journal Order by ID Asc"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
        DataGrid1.Refresh
   
    Exit Sub
    Resume Next
End Sub


Private Sub Form_Load()
On Error Resume Next
txtDate.Text = Date
cmbVoucher.AddItem "Single Debit"
cmbVoucher.AddItem "Single Credit"
cmbVoucher.AddItem "Cash Debit"
cmbVoucher.AddItem "Party Debit"
cmbVoucher.AddItem "Party Credit"

Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT AC_Code FROM Journal"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtCode.AddItem rs!AC_Code
        rs.MoveNext
        Loop
        rs.Close

On Error GoTo Last
    Set rs = New ADODB.Recordset
        str = "select * from Journal order by ID Asc"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
        
    Exit Sub
Last:
    MsgBox ("Database Connection error: " + Err.Description)

End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Set rsN = New ADODB.Recordset
        str = "select * from Journal where AC_Code like '" & txtCode.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsN.EOF Then
        txtTitle.Text = rsN!AC_Title
        rsN.Close
        txtTitle.SelStart = 0
        txtTitle.SelLength = Len(txtTitle.Text)
        txtTitle.SetFocus
    End If
End If
End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command2.SetFocus
End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbVoucher.SelStart = 0
    cmbVoucher.SelLength = Len(cmbVoucher.Text)
    cmbVoucher.SetFocus
End If
End Sub
Private Sub txtNaration_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNaration1.SelStart = 0
    txtNaration1.SelLength = Len(txtNaration1.Text)
    txtNaration1.SetFocus
End If
End Sub
Private Sub txtNaration1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNaration2.SelStart = 0
    txtNaration2.SelLength = Len(txtNaration2.Text)
    txtNaration2.SetFocus
End If
End Sub
Private Sub txtNaration2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'txtStr.SelStart = 0
    'txtStr.SelLength = Len(txtStr.Text)
    txtStr.SetFocus
End If
End Sub

Private Sub txtPs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If Val(txtPs.Text) > 0 Then
        Value = Val(txtPs.Text)
        Call Case_Ps
    End If
    If Val(txtTk.Text) >= 1 And Val(txtTk.Text) < 10 Then
        Value = Val(txtTk.Text)
        Call Case0_9
    End If
    If Val(txtTk.Text) >= 10 And Val(txtTk.Text) < 100 Then
        Value = Val(txtTk.Text)
        Call Case10_99
    End If
    If Val(txtTk.Text) >= 100 And Val(txtTk.Text) < 1000 Then
        Value = Val(txtTk.Text)
        Call Case100_999
    End If
    If Val(txtTk.Text) >= 1000 And Val(txtTk.Text) < 100000 Then
        Value = Val(txtTk.Text)
        Call Case1000_99999
    End If
    If Val(txtTk.Text) >= 100000 And Val(txtTk.Text) < 10000000 Then
        Value = Val(txtTk.Text)
        Call Case100000_9999999
    End If
    If Val(txtTk.Text) >= 10000000 Then
        Value = Val(txtTk.Text)
        Call Case10000000_999999999
    End If
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)
    txtCode.SetFocus
End If
End Sub
Private Sub txtStr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbContra.SelStart = 0
    cmbContra.SelLength = Len(cmbContra.Text)
    cmbContra.SetFocus
End If
End Sub
Private Sub txtTitle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNaration.SelStart = 0
    txtNaration.SelLength = Len(txtNaration.Text)
    txtNaration.SetFocus
End If
End Sub

Private Sub txtTk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPs.SelStart = 0
    txtPs.SelLength = Len(txtPs.Text)
    txtPs.SetFocus
End If
End Sub

Private Sub txtTran_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    On Error Resume Next
    Set rsN = New ADODB.Recordset
        str = "select * from Journal where ID like '" & txtTran.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsN.EOF Then
        If rsN!Voucher_Type = "DEBIT" Then
            cmbVoucher.Text = "Single Debit"
        End If
        If rsN!Voucher_Type = "CREDIT" Then
            cmbVoucher.Text = "Single Credit"
        End If
        If rsN!Voucher_Type = "Cash Debit" Then
            cmbVoucher.Text = "Cash Debit"
        End If
        If rsN!Voucher_Type = "Party Debit" Then
            cmbVoucher.Text = "Party Debit"
        End If
        If rsN!Voucher_Type = "Party Credit" Then
            cmbVoucher.Text = "Party Credit"
        End If
        txtDate.Text = rsN!Date
        txtTk.Text = rsN!Tk
        txtPs.Text = rsN!Ps
        txtCode.Text = rsN!AC_Code
        txtTitle.Text = rsN!AC_Title
        txtNaration.Text = rsN!Naration
        txtNaration1.Text = rsN!Naration1
        txtNaration2.Text = rsN!Naration2
        txtStr.Text = rsN!String
        cmbContra.Text = rsN!Contra_Code
        txtContra.Text = rsN!Contra_Title
        rsN.Close
    Else
        MsgBox "There is no such Voucher No. found.", 64, "Update Error"
        rsN.Close
    End If
End If
    Exit Sub
    Resume Next
End Sub
