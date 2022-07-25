VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWPrint_Loan 
   BackColor       =   &H00008000&
   Caption         =   "Weekly Loan Account Statement"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   Icon            =   "frmWPrint_Loan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   -135
      Top             =   2475
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Report Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   90
      TabIndex        =   4
      Top             =   135
      Width           =   6270
      Begin VB.ComboBox cmbD_Term 
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
         Left            =   3015
         TabIndex        =   20
         Text            =   "Combo2"
         Top             =   945
         Width           =   1320
      End
      Begin VB.ComboBox cmbP_Term 
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
         Left            =   3015
         TabIndex        =   19
         Text            =   "Combo2"
         Top             =   405
         Width           =   1320
      End
      Begin VB.OptionButton optTransaction 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Weekly Transection Report"
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
         TabIndex        =   18
         Top             =   1035
         Width           =   2670
      End
      Begin VB.OptionButton optPreodic 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Weekly Periodic Report"
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
         Top             =   450
         Width           =   2355
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   420
         Left            =   2430
         TabIndex        =   16
         Top             =   3240
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   420
         Left            =   4455
         TabIndex        =   15
         Top             =   3240
         Width           =   1365
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   420
         Left            =   405
         TabIndex        =   14
         Top             =   3240
         Width           =   1365
      End
      Begin VB.OptionButton optCenter 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Center Report"
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
         Top             =   4185
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.OptionButton optFO 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Weekly FO Report"
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
         TabIndex        =   12
         Top             =   1710
         Width           =   1950
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
         Left            =   2295
         TabIndex        =   11
         Top             =   1665
         Width           =   1095
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
         Left            =   2295
         TabIndex        =   10
         Text            =   "Combo2"
         Top             =   2430
         Width           =   1095
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
         Left            =   3195
         TabIndex        =   9
         Text            =   "Combo2"
         Top             =   4140
         Visible         =   0   'False
         Width           =   1815
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
         Left            =   1890
         TabIndex        =   8
         Top             =   4185
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cmbSamity_Name 
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
         TabIndex        =   7
         Text            =   "Combo2"
         Top             =   2430
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
         Left            =   4815
         TabIndex        =   6
         Text            =   "Combo2"
         Top             =   1665
         Width           =   1095
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
         Left            =   3465
         TabIndex        =   5
         Top             =   1665
         Width           =   1275
      End
      Begin VB.Shape Shape2 
         Height          =   2730
         Left            =   180
         Top             =   270
         Width           =   5895
      End
      Begin VB.Shape Shape1 
         Height          =   600
         Left            =   180
         Top             =   3150
         Width           =   5895
      End
      Begin VB.Label lblSamity_Code 
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
         Left            =   2295
         TabIndex        =   27
         Top             =   2115
         Width           =   1125
      End
      Begin VB.Label lblFo_Name 
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
         Left            =   3465
         TabIndex        =   26
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label lblFO_Code 
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
         Left            =   2295
         TabIndex        =   25
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label lblCenter_Name 
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
         Left            =   3195
         TabIndex        =   24
         Top             =   3870
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblCenter_Code 
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
         Left            =   1890
         TabIndex        =   23
         Top             =   3870
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblSamity_Name 
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
         Left            =   3555
         TabIndex        =   22
         Top             =   2115
         Width           =   1185
      End
      Begin VB.Label Label6 
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
         Left            =   4815
         TabIndex        =   21
         Top             =   1350
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   90
      ScaleHeight     =   180
      ScaleWidth      =   6210
      TabIndex        =   2
      Top             =   4140
      Width           =   6270
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   2265
      Left            =   90
      ScaleHeight     =   2205
      ScaleWidth      =   6210
      TabIndex        =   0
      Top             =   4500
      Width           =   6270
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmWPrint_Loan.frx":0442
         Height          =   1920
         Left            =   135
         TabIndex        =   1
         Top             =   135
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   3387
         _Version        =   393216
         BackColor       =   12648384
         HeadLines       =   1
         RowHeight       =   19
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   420
      Left            =   2340
      Top             =   6210
      Visible         =   0   'False
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   741
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
Attribute VB_Name = "frmWPrint_Loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim sl1 As Integer
    Dim sl2 As Integer
    Dim sort As Integer
    Dim Pagetotal As Double
    Dim sumtotal As Double
    Dim subtotal As Double
    Dim Balance As Double
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub ComboTerm()
        cmbTerm.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT Term FROM Parameter GROUP BY Term Order By Term"
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
        cmbTerm.AddItem "All"
End Sub
Private Sub ComboPTerm()
        cmbP_Term.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT Term FROM Parameter GROUP BY Term Order By Term"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbP_Term.AddItem rs!Term
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
        cmbP_Term.AddItem "All"
End Sub
Private Sub ComboDTerm()
        cmbD_Term.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT Term FROM Parameter GROUP BY Term Order By Term"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbD_Term.AddItem rs!Term
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
        cmbD_Term.AddItem "All"
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
        cmbSamity_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Samity_Name FROM Samity"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbSamity_Name.AddItem rs!Samity_Name
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
        rs.Close
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
        rs.Close
        End If
End Sub
Private Sub cmbFO_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbTerm.SetFocus
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
        rs.Close
        End If
End Sub
Private Sub cmbTerm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbSamity_Code.SetFocus
End If
End Sub

Private Sub Command1_Click()
Dim Today As Date
    Dim Mont As String
    Dim Trm As String
    Dim DTerm As String
    Dim PTerm As String
    Dim Center As String
    Dim FO As String
    Dim Samity As String
        Today = Date
        Mont = MonthName(Month(CDate(Today)))
        FO = cmbFO_Code.Text
        Samity = cmbSamity_Code.Text
        Trm = cmbTerm.Text
        PTerm = cmbP_Term.Text
        DTerm = cmbD_Term.Text
        'Center = cmbCenter_Code.Text
        sort = 0
'----------------------------------------------------------------------
If optPreodic.Value = True Then
    
       
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Loan_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Loan_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        If cmbP_Term.Text = "All" Then
        str = "select * FROM Loan_Master order by Sl_No"
        rptLoan_Collection.Label11.Caption = "WEEKLY LOAN PERIODIC REPORT ALL"
        Else
        str = "select * FROM Loan_Master where Term like '" & PTerm & "' order by Sl_No"
        rptLoan_Collection.Label11.Caption = "WEEKLY LOAN PERIODIC REPORT TERM WISE"
        End If
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Loan_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!Name = rs!Name
        rsN!Customer = rs!Customer
        rsN!AC_No = rs!AC_No
        rsN!Old_AC = rs!Old_AC
        rsN!Type = rs!Type
        rsN!Term = rs!Term
        rsN!Installment = rs!Installment
        rsN!Open_Date = rs!Open_Date
        rsN!Due_Date = rs!Due_Date
        rsN!Inst_Date = rs!Inst_Date
        rsN!Amount = rs!Amount
        rsN!Balance = rs!Balance
        rsN!Paid = rs!Paid
        rsN!Due = rs!Due
        rsN!Advance = rs!Advance
        rsN!Fine = rs!Fine
        rsN!Term_Fail = rs!Term_Fail
        
        rsN!Daily_Pay = rs!Daily_Pay
        rsN!Weekly_Pay = rs!Weekly_Pay
        rsN!Monthly_Pay = rs!Monthly_Pay
        rsN!Yearly_Pay = rs!Yearly_Pay
        
        rsN!Daily_Draw = rs!Daily_Draw
        rsN!Weekly_Draw = rs!Weekly_Draw
        rsN!Monthly_Draw = rs!Monthly_Draw
        rsN!Yearly_Draw = rs!Yearly_Draw
        
        rsN!Daily_Bal = rs!Daily_Bal
        rsN!Weekly_Bal = rs!Weekly_Bal
        rsN!Monthly_Bal = rs!Monthly_Bal
        rsN!Yearly_Bal = rs!Yearly_Bal
        
        rsN!Day_Close = rs!Day_Close
        rsN!Week_Close = rs!Week_Close
        rsN!Month_Close = rs!Month_Close
        rsN!Year_Close = rs!Year_Close
        
        rsN!Week_1 = rs!Week_1
        rsN!Week_2 = rs!Week_2
        rsN!Week_3 = rs!Week_3
        rsN!Week_4 = rs!Week_4
        rsN!Week_5 = rs!Week_5
        
        rsN!Center_name = rs!Center_name
        rsN!Center_code = rs!Center_code
        rsN!Samity_Name = rs!Samity_Name
        rsN!Samity_Code = rs!Samity_Code
        rsN!FO_Name = rs!FO_Name
        rsN!FO_Code = rs!FO_Code
        rsN!DPO_Name = rs!DPO_Name
        rsN!DPO_Code = rs!DPO_Code
        rsN!Month = Mont
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
     
     Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
     End If
     
    Set rs = New ADODB.Recordset
    str = "select * FROM Loan_Print order by Sl_No"
    rs.Open str, conn
        
    If Not rs.EOF Then
        rptLoan_Collection.rsLoan.ConnectionString = cnStr
        rptLoan_Collection.rsLoan.Source = str
         
         rptLoan_Collection.PBalance.DataField = "Weekly_Bal"
         rptLoan_Collection.Payment.DataField = "Weekly_Pay"
         rptLoan_Collection.Due.DataField = "Due"
         rptLoan_Collection.NBalance.DataField = "Balance"
         
         rptLoan_Collection.PBalance_Total.DataField = "Weekly_Bal"
         rptLoan_Collection.Payment_Total.DataField = "Weekly_Pay"
         rptLoan_Collection.Withdraw_Total.DataField = "Due"
         rptLoan_Collection.Nbalance_Total.DataField = "Amount"
        
        Timer1.Enabled = True
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
Command1.Enabled = False
End If

'--------------------------------------------------------------------
If optTransaction.Value = True Then
      
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Loan_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Loan_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If
      
Set rs = New ADODB.Recordset
        If cmbD_Term.Text = "All" Then
        str = "select * FROM Loan_Master where Weekly_Pay > 0 order by Sl_No"
        rptLoan_Collection.Label11.Caption = "WEEKLY LOAN TRANSACTION REPORT ALL"
        Else
        str = "select * FROM Loan_Master where Term like '" & DTerm & "' and Weekly_Pay > 0 order by Sl_No"
        rptLoan_Collection.Label11.Caption = "WEEKLY LOAN TRANSACTION REPORT TERM WISE"
        End If
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Loan_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!Name = rs!Name
        rsN!Customer = rs!Customer
        rsN!AC_No = rs!AC_No
        rsN!Old_AC = rs!Old_AC
        rsN!Type = rs!Type
        rsN!Term = rs!Term
        rsN!Installment = rs!Installment
        rsN!Open_Date = rs!Open_Date
        rsN!Due_Date = rs!Due_Date
        rsN!Inst_Date = rs!Inst_Date
        rsN!Amount = rs!Amount
        rsN!Balance = rs!Balance
        rsN!Paid = rs!Paid
        rsN!Due = rs!Due
        rsN!Advance = rs!Advance
        rsN!Fine = rs!Fine
        rsN!Term_Fail = rs!Term_Fail
        
        rsN!Daily_Pay = rs!Daily_Pay
        rsN!Weekly_Pay = rs!Weekly_Pay
        rsN!Monthly_Pay = rs!Monthly_Pay
        rsN!Yearly_Pay = rs!Yearly_Pay
        
        rsN!Daily_Draw = rs!Daily_Draw
        rsN!Weekly_Draw = rs!Weekly_Draw
        rsN!Monthly_Draw = rs!Monthly_Draw
        rsN!Yearly_Draw = rs!Yearly_Draw
        
        rsN!Daily_Bal = rs!Daily_Bal
        rsN!Weekly_Bal = rs!Weekly_Bal
        rsN!Monthly_Bal = rs!Monthly_Bal
        rsN!Yearly_Bal = rs!Yearly_Bal
        
        rsN!Day_Close = rs!Day_Close
        rsN!Week_Close = rs!Week_Close
        rsN!Month_Close = rs!Month_Close
        rsN!Year_Close = rs!Year_Close
        
        rsN!Week_1 = rs!Week_1
        rsN!Week_2 = rs!Week_2
        rsN!Week_3 = rs!Week_3
        rsN!Week_4 = rs!Week_4
        rsN!Week_5 = rs!Week_5
        
        rsN!Center_name = rs!Center_name
        rsN!Center_code = rs!Center_code
        rsN!Samity_Name = rs!Samity_Name
        rsN!Samity_Code = rs!Samity_Code
        rsN!FO_Name = rs!FO_Name
        rsN!FO_Code = rs!FO_Code
        rsN!DPO_Name = rs!DPO_Name
        rsN!DPO_Code = rs!DPO_Code
        rsN!Month = Mont
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
     
     Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
     End If
     
    Set rs = New ADODB.Recordset
    str = "select * FROM Loan_Print order by Sl_No"
    rs.Open str, conn
        
    If Not rs.EOF Then
        rptLoan_Collection.rsLoan.ConnectionString = cnStr
        rptLoan_Collection.rsLoan.Source = str
         
         rptLoan_Collection.PBalance.DataField = "Weekly_Bal"
         rptLoan_Collection.Payment.DataField = "Weekly_Pay"
         rptLoan_Collection.Due.DataField = "Due"
         rptLoan_Collection.NBalance.DataField = "Balance"
         
         rptLoan_Collection.PBalance_Total.DataField = "Weekly_Bal"
         rptLoan_Collection.Payment_Total.DataField = "Weekly_Pay"
         rptLoan_Collection.Withdraw_Total.DataField = "Due"
         rptLoan_Collection.Nbalance_Total.DataField = "Amount"
        
        Timer1.Enabled = True
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
Command1.Enabled = False
End If

'----------------------------------------------------------------------

If optFO.Value = True Then
     If cmbTerm.Text = "" Then
        Exit Sub
    MsgBox " Please Input Payment Term.", vbCritical, "Error"
    cmbTerm.SetFocus
    End If
    
    If cmbFO_Code.Text = "" Then
        Exit Sub
    MsgBox " Please Input FO code.", vbCritical, "Error"
    cmbFO_Code.SetFocus
    End If
    
    If cmbSamity_Code.Text = "" Then
        Exit Sub
    MsgBox " Please Input Samity code.", vbCritical, "Error"
    cmbSamity_Code.SetFocus
    End If
    
       
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Loan_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Loan_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If



Set rs = New ADODB.Recordset
    If cmbTerm.Text = "All" Then
        str = "select * FROM Loan_Master where FO_Code like '" & FO & "' AND Samity_Code like '" & Samity & "' AND Weekly_Pay > 0 order by Sl_No"
        rptFO_Loan.Label11.Caption = "F.O. WISE WEEKLY LOAN TRANSACTION REPORT ALL"
    Else
        str = "select * FROM Loan_Master where FO_Code like '" & FO & "' AND Samity_Code like '" & Samity & "' AND Term like '" & Trm & "' AND Weekly_Pay > 0 order by Sl_No"
        rptFO_Loan.Label11.Caption = "F.O. WISE WEEKLY LOAN TRANSACTION REPORT TERM WISE"
    End If
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Loan_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!Name = rs!Name
        rsN!Customer = rs!Customer
        rsN!AC_No = rs!AC_No
        rsN!Old_AC = rs!Old_AC
        rsN!Type = rs!Type
        rsN!Term = rs!Term
        rsN!Installment = rs!Installment
        rsN!Open_Date = rs!Open_Date
        rsN!Due_Date = rs!Due_Date
        rsN!Inst_Date = rs!Inst_Date
        rsN!Amount = rs!Amount
        rsN!Balance = rs!Balance
        rsN!Paid = rs!Paid
        rsN!Due = rs!Due
        rsN!Advance = rs!Advance
        rsN!Fine = rs!Fine
        rsN!Term_Fail = rs!Term_Fail
        
        rsN!Daily_Pay = rs!Daily_Pay
        rsN!Weekly_Pay = rs!Weekly_Pay
        rsN!Monthly_Pay = rs!Monthly_Pay
        rsN!Yearly_Pay = rs!Yearly_Pay
        
        rsN!Daily_Draw = rs!Daily_Draw
        rsN!Weekly_Draw = rs!Weekly_Draw
        rsN!Monthly_Draw = rs!Monthly_Draw
        rsN!Yearly_Draw = rs!Yearly_Draw
        
        rsN!Daily_Bal = rs!Daily_Bal
        rsN!Weekly_Bal = rs!Weekly_Bal
        rsN!Monthly_Bal = rs!Monthly_Bal
        rsN!Yearly_Bal = rs!Yearly_Bal
        
        rsN!Day_Close = rs!Day_Close
        rsN!Week_Close = rs!Week_Close
        rsN!Month_Close = rs!Month_Close
        rsN!Year_Close = rs!Year_Close
        
        rsN!Week_1 = rs!Week_1
        rsN!Week_2 = rs!Week_2
        rsN!Week_3 = rs!Week_3
        rsN!Week_4 = rs!Week_4
        rsN!Week_5 = rs!Week_5
        
        rsN!Center_name = rs!Center_name
        rsN!Center_code = rs!Center_code
        rsN!Samity_Name = rs!Samity_Name
        rsN!Samity_Code = rs!Samity_Code
        rsN!FO_Name = rs!FO_Name
        rsN!FO_Code = rs!FO_Code
        rsN!DPO_Name = rs!DPO_Name
        rsN!DPO_Code = rs!DPO_Code
        rsN!Month = Mont
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
     
     Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
     End If
     
     Set rs = New ADODB.Recordset
    str = "select * FROM Loan_Print order by Sl_No"
    rs.Open str, conn
        
    If Not rs.EOF Then
        rptFO_Loan.rsFO_Loan.ConnectionString = cnStr
        rptFO_Loan.rsFO_Loan.Source = str
         
         rptFO_Loan.rsFO_Loan.PBalance.DataField = "Weekly_Bal"
         rptFO_Loan.rsFO_Loan.Payment.DataField = "Weekly_Pay"
         rptFO_Loan.rsFO_Loan.Due.DataField = "Due"
         rptFO_Loan.rsFO_Loan.NBalance.DataField = "Balance"
         
         rptFO_Loan.rsFO_Loan.PBalance_Total.DataField = "Weekly_Bal"
         rptFO_Loan.rsFO_Loan.Payment_Total.DataField = "Weekly_Pay"
         rptFO_Loan.rsFO_Loan.Withdraw_Total.DataField = "Due"
         rptFO_Loan.rsFO_Loan.Nbalance_Total.DataField = "Amount"
        
        
        Timer1.Enabled = True
    Else
        rs.Close
    End If
Command1.Enabled = False
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
 Dim Today As Date
    Dim Mont As String
    Dim Trm As String
    Dim DTerm As String
    Dim PTerm As String
    Dim Center As String
    Dim FO As String
    Dim Samity As String
        Today = Date
        Mont = MonthName(Month(CDate(Today)))
        FO = cmbFO_Code.Text
        Samity = cmbSamity_Code.Text
        Trm = cmbTerm.Text
        PTerm = cmbP_Term.Text
        DTerm = cmbD_Term.Text
        'Center = cmbCenter_Code.Text
        sort = 0
'----------------------------------------------------------------------
If optPreodic.Value = True Then

Set rs = New ADODB.Recordset
        If cmbP_Term.Text = "All" Then
        str = "select * FROM Loan_Master order by Sl_No"
        Else
        str = "select * FROM Loan_Master where Term like '" & PTerm & "' order by Sl_No"
        End If
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        
        If Not rs.EOF Then
            Adodc1.ConnectionString = cnStr
            Adodc1.RecordSource = str
            Adodc1.Refresh
            DataGrid1.Refresh
            rs.Close
            Command1.Enabled = True
     Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
     End If
End If

'--------------------------------------------------------------------
If optTransaction.Value = True Then
      
Set rs = New ADODB.Recordset
        If cmbD_Term.Text = "All" Then
            str = "select * FROM Loan_Master where Weekly_Pay > 0  order by Sl_No"
        Else
            str = "select * FROM Loan_Master where Term like '" & DTerm & "' and Weekly_Pay > 0  order by Sl_No"
        End If
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly

     If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
        
        Command1.Enabled = True
     Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
     End If
End If

'----------------------------------------------------------------------

If optFO.Value = True Then
     If cmbTerm.Text = "" Then
        Exit Sub
    MsgBox " Please Input Payment Term.", vbCritical, "Error"
    cmbTerm.SetFocus
    End If
    
    If cmbFO_Code.Text = "" Then
        Exit Sub
    MsgBox " Please Input FO code.", vbCritical, "Error"
    cmbFO_Code.SetFocus
    End If
    
    If cmbSamity_Code.Text = "" Then
        Exit Sub
    MsgBox " Please Input Samity code.", vbCritical, "Error"
    cmbSamity_Code.SetFocus
    End If

Set rs = New ADODB.Recordset
    If cmbTerm.Text = "All" Then
        str = "select * FROM Loan_Master where FO_Code like '" & FO & "' AND Samity_Code like '" & Samity & "' AND Weekly_Pay > 0  order by Sl_No"
        Else
        str = "select * FROM Loan_Master where FO_Code like '" & FO & "' AND Samity_Code like '" & Samity & "' AND Term like '" & Trm & "' AND Weekly_Pay > 0  order by Sl_No"
    End If
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
    
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
        
        Command1.Enabled = True
     Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
     End If
End If

    Exit Sub
    End Sub

Private Sub Form_Load()
        cmbTerm.Visible = False
        cmbP_Term.Visible = False
        cmbD_Term.Visible = False
        
        Label6.Visible = False
        cmbFO_Code.Visible = False
        cmbFO_Name.Visible = False
        lblFO_Code.Visible = False
        lblFo_Name.Visible = False
        
        cmbCenter_Code.Visible = False
        cmbCenter_Name.Visible = False
        lblCenter_Code.Visible = False
        lblCenter_Name.Visible = False
        
        cmbSamity_Code.Visible = False
        cmbSamity_Name.Visible = False
        lblSamity_Code.Visible = False
        lblSamity_Name.Visible = False

        Call ComboPTerm
        Call ComboDTerm
        Call ComboTerm
        
        Call ComboCenter_Code
        Call ComboCenter_Name
        Call ComboSamity_Code
        Call ComboSamity_Name
        Call ComboFO_Code
        Call ComboFO_Name
        Command1.Enabled = False
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
        str = "SELECT DISTINCT Branch_Code, Branch_name FROM Branch where Branch_Code like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbCenter_Name.Text = rs!Branch_Name
        rs.Close
        End If
End Sub

Private Sub cmbCenter_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3.SetFocus
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
        str = "SELECT DISTINCT Branch_Code, Branch_name FROM Branch where Branch_Name like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbCenter_Code.Text = rs!Branch_Code
        rs.Close
        End If
End Sub
Private Sub cmbSamity_Code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbSamity_Name.SetFocus
End If
End Sub

Private Sub cmbSamity_Code_LostFocus()
If cmbSamity_Code.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbSamity_Code.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Samity_Code, Samity_Name FROM Samity where Samity_code like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbSamity_Name.Text = rs!Samity_Name
        rs.Close
        End If
End Sub

Private Sub cmbSamity_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3.SetFocus
End If
End Sub

Private Sub cmbSamity_Name_LostFocus()
If cmbSamity_Name.Text = "" Then
Exit Sub
End If

Dim code As String
code = cmbSamity_Name.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Samity_Code, Samity_Name FROM Samity where Samity_Name like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbSamity_Code.Text = rs!Samity_Code
        rs.Close
        End If
End Sub
Private Sub optCenter_GotFocus()
cmbFO_Code.Visible = False
cmbFO_Name.Visible = False
lblFO_Code.Visible = False
lblFo_Name.Visible = False

cmbCenter_Code.Visible = True
cmbCenter_Name.Visible = True
lblCenter_Code.Visible = True
lblCenter_Name.Visible = True

cmbSamity_Code.Visible = False
cmbSamity_Name.Visible = False
lblSamity_Code.Visible = False
lblSamity_Name.Visible = False
End Sub

Private Sub optCenter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbCenter_Code.SetFocus
End If
End Sub

Private Sub optFO_Click()
cmbD_Term.Visible = False
cmbP_Term.Visible = False

cmbTerm.Visible = True
cmbTerm.Text = "All"
Label6.Visible = True
cmbFO_Code.Visible = True
cmbFO_Name.Visible = True
lblFO_Code.Visible = True
lblFo_Name.Visible = True
cmbSamity_Code.Visible = True
cmbSamity_Name.Visible = True
lblSamity_Code.Visible = True
lblSamity_Name.Visible = True
cmbFO_Code.SetFocus
End Sub

Private Sub optFO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbFO_Code.SetFocus
End If
End Sub
Private Sub optPreodic_Click()
cmbP_Term.Visible = True
cmbD_Term.Visible = False
cmbP_Term.Text = "All"
cmbP_Term.SetFocus


cmbTerm.Visible = False
Label6.Visible = False
cmbFO_Code.Visible = False
cmbFO_Name.Visible = False
lblFO_Code.Visible = False
lblFo_Name.Visible = False

cmbCenter_Code.Visible = False
cmbCenter_Name.Visible = False
lblCenter_Code.Visible = False
lblCenter_Name.Visible = False

cmbSamity_Code.Visible = False
cmbSamity_Name.Visible = False
lblSamity_Code.Visible = False
lblSamity_Name.Visible = False
End Sub



Private Sub optSamity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbSamity_Code.SetFocus
End If
End Sub

Private Sub optTransaction_Click()

cmbD_Term.Visible = True
cmbP_Term.Visible = False

cmbD_Term.Text = "All"
cmbD_Term.SetFocus

cmbTerm.Visible = False
Label6.Visible = False
cmbFO_Code.Visible = False
cmbFO_Name.Visible = False
lblFO_Code.Visible = False
lblFo_Name.Visible = False

cmbCenter_Code.Visible = False
cmbCenter_Name.Visible = False
lblCenter_Code.Visible = False
lblCenter_Name.Visible = False

cmbSamity_Code.Visible = False
cmbSamity_Name.Visible = False
lblSamity_Code.Visible = False
lblSamity_Name.Visible = False
End Sub

Private Sub Timer1_Timer()
If optPreodic.Value = True Then
    ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptLoan_Collection.Show 1
    End If
End If

If optTransaction.Value = True Then
    ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptLoan_Collection.Show 1
    End If
End If

If optFO.Value = True Then
    ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptFO_Loan.Show 1
    End If
End If

End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
optFO.SetFocus
End If
End Sub

Private Sub txtDay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.Enabled = True
End If
End Sub

Private Sub txtToday_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.Enabled = True
End If
End Sub


