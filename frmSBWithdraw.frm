VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSBWithdraw 
   BackColor       =   &H00008000&
   Caption         =   "Deposit Withdraw"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11715
   Icon            =   "frmSBWithdraw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0FFC0&
      Height          =   5460
      Left            =   7155
      ScaleHeight     =   5400
      ScaleWidth      =   4275
      TabIndex        =   24
      Top             =   135
      Width           =   4335
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
         TabIndex        =   25
         ToolTipText     =   "Enter Consumer no."
         Top             =   180
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmSBWithdraw.frx":0442
         Height          =   4560
         Left            =   240
         TabIndex        =   26
         Top             =   735
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   8043
         _Version        =   393216
         BackColor       =   12648384
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
         TabIndex        =   27
         Top             =   255
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   135
      ScaleHeight     =   675
      ScaleWidth      =   6780
      TabIndex        =   23
      Top             =   4860
      Width           =   6840
      Begin VB.CommandButton Command4 
         Caption         =   "Print"
         Height          =   435
         Left            =   3480
         TabIndex        =   31
         Top             =   135
         Width           =   1485
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   435
         Left            =   1800
         TabIndex        =   30
         Top             =   135
         Width           =   1485
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   435
         Left            =   135
         TabIndex        =   29
         Top             =   135
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   435
         Left            =   5115
         TabIndex        =   28
         Top             =   135
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   3585
      Left            =   135
      ScaleHeight     =   3525
      ScaleWidth      =   6780
      TabIndex        =   2
      Top             =   1095
      Width           =   6840
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
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
         Left            =   5115
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   3045
         Width           =   1455
      End
      Begin VB.TextBox txtOther 
         Alignment       =   1  'Right Justify
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
         Left            =   1605
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   2475
         Width           =   1455
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
         Left            =   5085
         TabIndex        =   12
         Top             =   195
         Width           =   1455
      End
      Begin VB.TextBox txtType 
         Alignment       =   2  'Center
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
         Left            =   5085
         TabIndex        =   11
         Top             =   765
         Width           =   1455
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
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
         Left            =   5085
         TabIndex        =   10
         Top             =   2460
         Width           =   1455
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   1620
         TabIndex        =   9
         Top             =   1890
         Width           =   2355
      End
      Begin VB.TextBox txtMR_No 
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
         Left            =   1620
         TabIndex        =   8
         Top             =   180
         Width           =   1065
      End
      Begin VB.TextBox txtAccount 
         Alignment       =   2  'Center
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
         Left            =   1620
         TabIndex        =   7
         Top             =   750
         Width           =   1860
      End
      Begin VB.TextBox txtName 
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
         Left            =   1620
         TabIndex        =   6
         Top             =   1305
         Width           =   2715
      End
      Begin VB.TextBox txtInstallment 
         Alignment       =   1  'Right Justify
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
         Left            =   5085
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   1890
         Width           =   1455
      End
      Begin VB.TextBox txtTerm 
         Alignment       =   2  'Center
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
         Left            =   5085
         TabIndex        =   4
         Top             =   1305
         Width           =   1455
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
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
         Left            =   1620
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   3045
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Left            =   3735
         TabIndex        =   35
         Top             =   3120
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Profit/Others"
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
         Left            =   375
         TabIndex        =   34
         Top             =   2565
         Width           =   1095
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
         Left            =   4575
         TabIndex        =   22
         Top             =   270
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Withdrawl Tk"
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
         Left            =   240
         TabIndex        =   21
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
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
         Left            =   3780
         TabIndex        =   20
         Top             =   810
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
         Left            =   4275
         TabIndex        =   19
         Top             =   2520
         Width           =   705
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
         Left            =   540
         TabIndex        =   18
         Top             =   1980
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
         Left            =   855
         TabIndex        =   17
         Top             =   225
         Width           =   660
      End
      Begin VB.Label Label2 
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
         Left            =   435
         TabIndex        =   16
         Top             =   825
         Width           =   1080
      End
      Begin VB.Label Label3 
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
         Left            =   30
         TabIndex        =   15
         Top             =   1395
         Width           =   1485
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
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
         Left            =   4545
         TabIndex        =   14
         Top             =   1350
         Width           =   435
      End
      Begin VB.Label Label15 
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
         Left            =   4035
         TabIndex        =   13
         Top             =   1980
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   135
      ScaleHeight     =   675
      ScaleWidth      =   6780
      TabIndex        =   0
      Top             =   180
      Width           =   6840
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEPOSIT WITHDRAWAL POSTING"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   525
         TabIndex        =   1
         Top             =   90
         Width           =   5760
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7155
      Top             =   5175
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
Attribute VB_Name = "frmSBWithdraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Today As Date
Dim ac As String
Dim inword As String
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Dim StrD As String
Dim Value As Double
Dim Result As Double
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
    Value = Result
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
         Result = Value
    Call Case_Result
        Str1 = str
        inword = Tk + Str1 + Only
 End Sub
Private Sub Case10_99()
    Result = Value
    Call Case_Result
        Str10 = str
        inword = Tk + Str10 + Only
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
              inword = Tk + Str100 + Only
        
        Else
              inword = Tk + Str100 + Only
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
        Result = Div
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
        Result = Div
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
        txtType.Text = ""
        txtTerm.Text = ""
        txtInstallment.Text = "0.00"
        txtBalance.Text = "0.00"
        txtNet.Text = "0.00"
        txtOther.Text = "0.00"
        txtTotal.Text = "0.00"
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
Private Sub Check1_Click()
On Error Resume Next
If Check1.Value = 1 Then
Picture3.Visible = False
Image1.Visible = True
Else
Image1.Visible = False
Picture3.Visible = True
End If
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Command2.SetFocus
End If
If Command2.Enabled = False Then
Command5.SetFocus
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
    Dim vn As String
    Dim Week As Date
    
    On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Deposit_Master where AC_No like '" & txtAccount.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
        rsU!Amount = rsU!Amount - Val(txtNet.Text)
        rsU!Fine = rsU!Fine + Val(txtOther.Text)
        rsU!Withdraw = rsU!Withdraw + Val(txtNet.Text)
        rsU!Daily_Draw = rsU!Daily_Draw + Val(txtNet.Text)
        rsU!Weekly_Draw = rsU!Weekly_Draw + Val(txtNet.Text)
        rsU!Monthly_Draw = rsU!Monthly_Draw + Val(txtNet.Text)
        rsU!Yearly_Draw = rsU!Yearly_Draw + Val(txtNet.Text)
                
        rsU.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = txtAccount.Text
        rsN!MR_No = rsN!sl
            vn = rsN!sl
        rsN!Name = txtName.Text
        rsN!Description = txtDescription.Text + "VN" & Format$(Val(rsN!sl), "###0")
        rsN!Dr = txtNet.Text
        rsN!Fine = -Val(txtOther.Text)
        rsN!Balance = rsU!Amount
        rsN!Term = txtTerm.Text
        rsN!Center_name = rsU!Center_name
        rsN!Center_code = rsU!Center_code
        rsN!Samity_Name = rsU!Samity_Name
        rsN!Samity_Code = rsU!Samity_Code
        rsN!FO_Name = rsU!FO_Name
        rsN!FO_Code = rsU!FO_Code
        rsN!DPO_Name = rsU!DPO_Name
        rsN!DPO_Code = rsU!DPO_Code
        rsN.Update
        rsN.Close
        rsU.Close
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtTotal.Text)
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = 0
        rsN!Cr = Val(txtTotal.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!MR_No = vn
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Cr = Val(txtTotal.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Cash_Cr = rsU!Cash_Cr + Val(txtTotal.Text)
        rsU!Cash_Close = rsU!Cash_Close - Val(txtTotal.Text)
        rsU.Update
        rsU.Close

'---------------------------------------------------------------------------------------
  If txtType.Text = "BM-D" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "BM-W" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "BM-M" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "BM-Y" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100121 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "D2Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100147 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "D3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100147 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "D4Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100147 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "D5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100104 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100147 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "W3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100105 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100148 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "W5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100105 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100148 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "W7Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100105 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100148 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "M5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100149 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "M10Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100149 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If
 
    If txtType.Text = "M15Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100149 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If
 
    If txtType.Text = "M20Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100106 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100149 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If
 
    If txtType.Text = "Y10Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100128 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100150 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "LP3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "LP5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "LP8Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "LP10Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "LP12Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100130 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "FD2Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "FD3Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "FD4Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "FD5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "FD5.5Years" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100129 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "MFD24Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "MFD36Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "MFD48Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "MFD60Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If

    If txtType.Text = "MFD66Months" Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100131 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - Val(txtNet.Text)
        rs.Update
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = Today
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C:" & txtAccount.Text
        rsN!Dr = Val(txtNet.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    If Val(txtOther.Text) > 0 Then
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100151 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtOther.Text)
        rs.Update

    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "Profit A/C:" & txtAccount.Text
        rsN!Dr = Val(txtOther.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    End If
    End If
  
  
  '===================================================================
    End If
    
    Set rs = New ADODB.Recordset
        str = "select * from Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by AC_No"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
    
    Call clearTextboxes
    Command2.Enabled = False
    
    txtDate.Text = Today
    txtDescription.Text = "Savings Withdraw"
    txtAccount.Text = ac
    txtAccount.SelStart = 7
    txtAccount.SelLength = Len(txtAccount.Text)
    txtAccount.SetFocus

Exit Sub
   Resume Next
End Sub

Private Sub Command3_Click()
Set rs = New ADODB.Recordset
        str = "select * from Tran where sl like '" & txtMR_No.Text & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        txtMR_No.Text = rs!sl
        txtAccount.Text = rs!AC_No
        txtDate.Text = rs!Date
        txtName.Text = rs!Name
        txtType.Text = rs!Type
        txtTerm.Text = rs!Term
        txtDescription.Text = rs!Description
        txtNet.Text = rs!Dr
        txtBalance.Text = rs!Balance
        rs.Close
     
     If MsgBox("Really want to delete?", vbCritical + vbYesNo) = vbYes Then
        
        str = "delete from Tran where sl like '" & txtMR_No.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.Close
        
    Set rsU = New ADODB.Recordset
        str = "select * from Deposit_Master where AC_No like '" & txtAccount.Text & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsU.EOF Then
       rsU!Amount = rsU!Amount + Val(txtNet.Text)
        rsU!Withdraw = rsU!Withdraw - Val(txtNet.Text)
        rsU!Daily_Draw = rsU!Daily_Draw - Val(txtNet.Text)
        rsU!Weekly_Draw = rsU!Weekly_Draw - Val(txtNet.Text)
        rsU!Monthly_Draw = rsU!Monthly_Draw - Val(txtNet.Text)
        rsU!Yearly_Draw = rsU!Yearly_Draw - Val(txtNet.Text)
        
        rsU.Update
    
   
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100100 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + Val(txtTotal.Text)
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = "A/C No.-" & txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
    
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Cash_Book", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!MR_No = txtMR_No.Text
        rsN!Description = "A/C:" + txtAccount.Text
        rsN!Dr = Val(txtTotal.Text)
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        rsU!Cash_Dr = rsU!Cash_Dr + Val(txtTotal.Text)
        rsU!Cash_Close = rsU!Cash_Close + Val(txtTotal.Text)
        rsU.Update
        rsU.Close
    
'----------------------------------------------------------------------------
        If txtTerm.Text = "Daily" Then
            
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100104 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtNet.Text)
            rs.Update
            
            Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C No.-" & txtAccount.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtNet.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        If Val(txtOther.Text) > 0 Then
        Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100147 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(txtOther.Text)
            rs.Update
    
        Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "Fine form A/C:" & txtAccount.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtOther.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        End If
        End If

'----------------------------------------------------------------------------
        If txtTerm.Text = "Weekly" Then
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100105 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtNet.Text)
            rs.Update
       Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C No.-" & txtAccount.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtNet.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        If Val(txtOther.Text) > 0 Then
        Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100148 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(txtOther.Text)
            rs.Update
    
        Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "Fine form A/C:" & txtAccount.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtOther.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        End If
        End If
'----------------------------------------------------------------------------
        If txtTerm.Text = "Monthly" Then
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100106 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtNet.Text)
            rs.Update
            
        Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C No.-" & txtAccount.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtNet.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        If Val(txtOther.Text) > 0 Then
        Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100149 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(txtOther.Text)
            rs.Update
    
        Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "Fine form A/C:" & txtAccount.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtOther.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        End If
        End If
'----------------------------------------------------------------------------
        If txtTerm.Text = "Yearly" Then
            Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100128 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance + Val(txtNet.Text)
            rs.Update
            
        Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = Today
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "A/C:" & txtAccount.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtNet.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        
        If Val(txtOther.Text) > 0 Then
        Set rs = New ADODB.Recordset
            str = "select * from GL_Master where AC_No like '" & 100150 & "'"
            rs.Open str, conn, adOpenDynamic, adLockOptimistic
            rs!Balance = rs!Balance - Val(txtOther.Text)
            rs.Update
    
        Set rsN = New ADODB.Recordset
            rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
            rsN!Date = txtDate.Text
            rsN!AC_No = rs!AC_No
            rsN!Name = rs!Head_Name
            rsN!Description = "Fine form A/C:" & txtAccount.Text
            rsN!Dr = 0
            rsN!Cr = Val(txtOther.Text)
            rsN!Balance = rs!Balance
            rsN.Update
            rsN.Close
            rs.Close
        End If
        End If
'----------------------------------------------------------------------------
    End If
    
    Call clearTextboxes
        
        str = "select * from Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by sl"
        rs.Open str, conn, adOpenKeyset, adLockReadOnly
        
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        Call ColumnWidth
        End If
    Else
        MsgBox "There is no such Transaction no. found: " & txtMR_No.Text
        rs.Close
    End If
        Command3.Enabled = False
        Command4.Enabled = False
        Command2.Enabled = False
        txtAccount.Text = ac
        txtDate.Text = Today
    Exit Sub
    Resume Next
End Sub

Private Sub Command4_Click()
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
        str = "select * from Tran where sl like '" & txtMR_No.Text & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        rs.Close
    
    If Val(txtTotal.Text) >= 1 And Val(txtTotal.Text) < 10 Then
        Value = Val(txtTotal.Text)
        Call Case0_9
    End If
    If Val(txtTotal.Text) >= 10 And Val(txtTotal.Text) < 100 Then
        Value = Val(txtTotal.Text)
        Call Case10_99
    End If
    If Val(txtTotal.Text) >= 100 And Val(txtTotal.Text) < 1000 Then
        Value = Val(txtTotal.Text)
        Call Case100_999
    End If
    If Val(txtTotal.Text) >= 1000 And Val(txtTotal.Text) < 100000 Then
        Value = Val(txtTotal.Text)
        Call Case1000_99999
    End If
    If Val(txtTotal.Text) >= 100000 And Val(txtTotal.Text) < 10000000 Then
        Value = Val(txtTotal.Text)
        Call Case100000_9999999
    End If
    If Val(txtTotal.Text) >= 10000000 Then
        Value = Val(txtTotal.Text)
        Call Case10000000_999999999
    End If
    
    Set rsN = New ADODB.Recordset
        rsN.Open "Money_Receipt", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!sl = 1
        rsN!MR_No = txtMR_No.Text
        rsN!Date = txtDate.Text
        rsN!AC_No = txtAccount.Text
        rsN!Name = txtName.Text
        rsN!Description = txtDescription.Text
        rsN!Cr = txtNet.Text
        rsN!Fine = txtOther.Text
        rsN!Total = txtTotal.Text
        rsN!inword = inword
        rsN.Update
        rsN.Close
    
    
    
    Call clearTextboxes
    
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    txtDate.Text = Today
    txtDescription.Text = "Installment Paid"
    txtAccount.Text = ac
    'txtAccount.SelStart = 7
    txtAccount.SelLength = Len(txtAccount.Text)
    'txtAccount.SetFocus

Set rs = New ADODB.Recordset
    str = "select * from Money_Receipt"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.RecordCount > 0 Then
    rs.Update
    rs.Close

    str = "select * from Money_Receipt"
    rptMoney_Receipt.rsMoney.ConnectionString = cnStr
    rptMoney_Receipt.rsMoney.Source = str
    rptMoney_Receipt.Label38.Caption = "WITHDRAWAL VOUCHER"
    rptMoney_Receipt.Label44.Caption = "Received by"
    'rptMoney_Receipt.Label40.Caption = ""
    'rptMoney_Receipt.Label11.Caption = ""
    End If
    rptMoney_Receipt.Show 1
    End If
Else

Exit Sub
End If
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
        rsU!Account_no = txtAccount.Text
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

Private Sub Form_Load()
 Dim suf As String
 suf = 101
 On Error Resume Next
       Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
           On Error Resume Next
           Today = rs!Today
           ac = rs!Branch_Code + suf
           rs.Close
           txtDate.Text = Today
           txtAccount.Text = ac
           txtDescription.Text = "Savings Withdraw"
        End If
 
    Set rs = New ADODB.Recordset
        str = "select * from Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') Order by Sl"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
        
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
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription.Text)
    txtDescription.SetFocus
End If
End Sub

Private Sub txtAccount_LostFocus()
If txtAccount.Text = "" Then
Exit Sub
End If

Dim ID As String
Dim L_Id As String
    ID = txtAccount.Text
    
    txtAccount.SelStart = 7
    txtAccount.SelLength = Len(txtAccount.Text)
    L_Id = txtAccount.SelText
    
    Set rs = New ADODB.Recordset
        str = "select * from Deposit_Master where AC_No like '" & ID & "' Or Old_Ac like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
        If rs!C_lose = "Yes" Then
                MsgBox "Account already close!", vbCritical, "Close Info"
                Call clearTextboxes
                txtAccount.SetFocus
            Else
                Set rsU = New ADODB.Recordset
                str = "select * from Loan_Master where Customer like '" & L_Id & "'"
                rsU.Open str, conn
                
                If Not rsU.EOF Then
                    If rsU!C_lose = "Yes" Then
                        rsU.Close
                    Else
                        MsgBox "Customer Has a Loan A/C: " & rsU!AC_No & " Balance: " & Format$(Val(rsU!Balance), "###0.00"), vbInformation, "Loan Info!"
                        rsU.Close
                    End If
                End If
                
                Set rsU = New ADODB.Recordset
                str = "select * from Loan_Info where G_Id1 like '" & L_Id & "' Or G_Id2 like '" & L_Id & "' Or G_Id3 like '" & L_Id & "' Or G_Id4 like '" & L_Id & "' Or G_Id5 like '" & L_Id & "' Or G_Id6 like '" & L_Id & "' Or G_Id7 like '" & L_Id & "' Or G_Id8 like '" & L_Id & "' Or G_Id9 like '" & L_Id & "' Or G_Id10 like '" & L_Id & "'"
                rsU.Open str, conn

                If Not rsU.EOF Then
                    
                    Set rsN = New ADODB.Recordset
                    str = "select * from Loan_Master where Ac_No like '" & rsU!AC_No & "'"
                    rsN.Open str, conn

                        If Not rsN.EOF Then
                            
                            If rsN!C_lose = "Yes" Then
                            rsN.Close
                            rsU.Close
                         Else
                             MsgBox "Customer is a Guarantor of Loan A/C: " & rsU!AC_No, vbInformation, "Loan Info!"
                             rsN.Close
                             rsU.Close
                         End If
                  End If
          End If
                
               If rs!Term_Fail > 1 Then
               MsgBox "Member is Defaulter for " & rs!Term_Fail & "Term", vbCritical, "Information!"
               End If
               
               'Call clearTextboxes
                On Error Resume Next
        
                txtAccount.Text = rs!AC_No
                txtDate.Text = Today
                txtName.Text = rs!Name
                txtType.Text = rs!Type
                txtTerm.Text = rs!Term
                txtInstallment.Text = rs!Installment
                txtBalance.Text = rs!Amount
                rs.Close
                txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
                txtInstallment.Text = Format$(Val(txtInstallment.Text), "###0.00")
                txtBalance.Text = Format$(Val(txtBalance.Text), "###0.00")
                
        Command2.Enabled = True
        End If
    Else
    MsgBox "There is no such Account no. found.,", vbCritical
        rs.Close
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

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtOther.SelStart = 0
    txtOther.SelLength = Len(txtOther.Text)
    txtOther.SetFocus
End If
End Sub

Private Sub txtMR_No_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtName.SetFocus
End If
End Sub

Private Sub txtMR_No_LostFocus()
If txtMR_No.Text = "" Then
Exit Sub
End If

On Error Resume Next
Dim Tran As String
    Tran = txtMR_No.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Tran where sl like '" & Tran & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        txtMR_No.Text = rs!sl
        txtAccount.Text = rs!AC_No
        txtDate.Text = rs!Date
        txtName.Text = rs!Name
        txtType.Text = rs!Type
        txtTerm.Text = rs!Term
        txtDescription.Text = rs!Description
        txtNet.Text = rs!Dr
        txtOther.Text = rs!Fine
        txtTotal.Text = Val(rs!Dr) + Val(rs!Fine)
        txtBalance.Text = rs!Balance
        rs.Close
        
        txtNet.Text = Format$(Val(txtNet.Text), "###0.00")
        txtTotal.Text = Format$(Val(txtTotal.Text), "###0.00")
        
        Command3.Enabled = True
        Command4.Enabled = True
        'Command5.Enabled = True
        Command2.Enabled = False
        
    Else
    If MsgBox("There is no such Transaction no found! Verify Transaction no.?", vbCritical + vbYesNo) = vbYes Then
        Call clearTextboxes
        rs.Close
        txtMR_No.SetFocus
        txtAccount.Text = ac
        txtDate.Text = Today
    Else
        Call clearTextboxes
        txtAccount.Text = ac
        txtDate.Text = Today
        txtAccount.SetFocus
    End If
    End If
    Exit Sub

End Sub

Private Sub txtNet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Command2.Enabled = True Then
    Command2.SetFocus
    Else
    Exit Sub
    End If
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
If Val(txtNet.Text) > Val(txtBalance.Text) Then
MsgBox "Insufficient Balance", vbCritical, "Error!"
txtNet.Text = ""
txtNet.SetFocus
Else
txtTotal.Text = Format$(Val(txtNet.Text) + Val(txtOther.Text), "###0.00")
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
        str = "select * from Tran where Ac_no Like '" & "%" & search & "%" & "'"
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
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
    Combo1.SetFocus
End If
End Sub

Private Sub txtTran_LostFocus()
On Error Resume Next
Dim Tran As String
    Tran = txtTran.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Palli_Bill where Tr_no like '" & Tran & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        txtTran.Text = rs!Tr_no
        Combo1.Text = rs!Tr_type
        txtAccount.Text = rs!Account_no
        txtAccbill.Text = rs!Ac_Bill
        txtSer.Text = rs!Ser_charge
        txtOther.Text = rs!Others
        txtVat.Text = rs!vat
        txtNetbill.Text = rs!Net_bill
        Check1.Value = rs!Revenue
        txtDate.Text = rs!Date
        rs.Close
        Command3.Enabled = True
        Command5.Enabled = True
        Command2.Enabled = True
        
    Else
    If MsgBox("There is no such Transaction no found, Do you want add new employee?", vbCritical + vbYesNo) = vbYes Then
        Call clearTextboxes
        rs.Close
    Else
        Call clearTextboxes
    End If
    End If
    Exit Sub
Last:
    MsgBox ("Database Connection error: " + Err.Description)
End Sub

Private Sub txtVat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSer.SelStart = 0
    txtSer.SelLength = Len(txtSer.Text)
    txtSer.SetFocus
End If
End Sub

Private Sub txtVat_LostFocus()
On Error Resume Next
Dim a, b, c, d, e As Double
    a = Val(txtAccbill.Text)
    b = Val(txtSer.Text)
    c = Val(txtOther.Text)

    d = Val(txtVat.Text)
    
    txtNetbill.Text = Format$(Round(a + b + c + d), "###0.00")
    txtAccbill.Text = Format$(Val(txtAccbill.Text), "###0.00")
    txtSer.Text = Format$(Val(txtSer.Text), "###0.00")
    txtOther.Text = Format$(Val(txtOther.Text), "###0.00")
    txtVat.Text = Format$(Val(txtVat.Text), "###0.00")
    
    txtSer.SelStart = 0
    txtSer.SelLength = Len(txtSer.Text)
End Sub




