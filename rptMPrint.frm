VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMPrint 
   BackColor       =   &H00008000&
   Caption         =   "Monthly Report Print"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5700
   Icon            =   "rptMPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   -90
      Top             =   2520
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Statement Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   135
      TabIndex        =   4
      Top             =   180
      Width           =   5370
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   420
         Left            =   1890
         TabIndex        =   16
         Top             =   2880
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   420
         Left            =   3555
         TabIndex        =   15
         Top             =   2880
         Width           =   1365
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   420
         Left            =   315
         TabIndex        =   14
         Top             =   2880
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
         Left            =   270
         TabIndex        =   13
         Top             =   3915
         Width           =   1500
      End
      Begin VB.OptionButton optFO 
         BackColor       =   &H00C0FFC0&
         Caption         =   "FO Report"
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
         Top             =   1260
         Width           =   1230
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
         Left            =   1890
         TabIndex        =   11
         Top             =   1215
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
         Left            =   1890
         TabIndex        =   10
         Text            =   "Combo2"
         Top             =   1980
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
         Left            =   3150
         TabIndex        =   9
         Text            =   "Combo2"
         Top             =   3870
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
         Left            =   1845
         TabIndex        =   8
         Top             =   3870
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
         Left            =   3195
         TabIndex        =   7
         Text            =   "Combo2"
         Top             =   1935
         Width           =   1815
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
         Left            =   1890
         TabIndex        =   6
         Text            =   "Combo2"
         Top             =   495
         Width           =   1275
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
         Left            =   3195
         TabIndex        =   5
         Top             =   1215
         Width           =   1770
      End
      Begin VB.Shape Shape2 
         Height          =   2235
         Left            =   180
         Top             =   315
         Width           =   4950
      End
      Begin VB.Shape Shape1 
         Height          =   780
         Left            =   180
         Top             =   2700
         Width           =   4950
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
         Left            =   1890
         TabIndex        =   23
         Top             =   1665
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
         Left            =   3195
         TabIndex        =   22
         Top             =   900
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
         Left            =   1890
         TabIndex        =   21
         Top             =   900
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
         Left            =   3150
         TabIndex        =   20
         Top             =   3780
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
         Left            =   1845
         TabIndex        =   19
         Top             =   3780
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
         Left            =   3195
         TabIndex        =   18
         Top             =   1665
         Width           =   1185
      End
      Begin VB.Label Label6 
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
         Left            =   315
         TabIndex        =   17
         Top             =   540
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   135
      ScaleHeight     =   180
      ScaleWidth      =   5310
      TabIndex        =   2
      Top             =   3960
      Width           =   5370
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   2130
      Left            =   135
      ScaleHeight     =   2070
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   4365
      Width           =   5415
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "rptMPrint.frx":0442
         Height          =   1785
         Left            =   135
         TabIndex        =   1
         Top             =   135
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3149
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
      Left            =   2385
      Top             =   5400
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
Attribute VB_Name = "frmMPrint"
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
        cmbTerm.AddItem "All"
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
        cmbSamity_Name.AddItem "ALL"
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
        cmbSamity_Code.AddItem "ALL"
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
cmbSamity_Code.SetFocus
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
optFO.SetFocus
End If
End Sub

Private Sub Command1_Click()
    Dim Today As Date
    Dim Mont As String
    Dim yr As String
    Dim Trm As String
    Dim Center As String
    Dim FO As String
    Dim Samity As String
    
    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
        On Error Resume Next
        Today = rs!Today
        End If
        
        Mont = MonthName(Month(CDate(Today)))
        Mont = UCase(Mont)
        yr = Year(CDate(Today))
        FO = cmbFO_Code.Text
        Samity = cmbSamity_Code.Text
        Trm = cmbTerm.Text
        Center = cmbCenter_Code.Text
        sort = 0


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
        str = "select * from Deposit_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Deposit_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
    If cmbTerm.Text = "All" Then
        If cmbSamity_Code.Text = "ALL" Then
            str = "select * FROM Deposit_Master where FO_Code like '" & FO & "' order by Sl_No"
        Else
            str = "select * FROM Deposit_Master where FO_Code like '" & FO & "' AND Samity_Code like '" & Samity & "' order by Sl_No"
        End If
            rptSBMonthly.Label38.Caption = "MONTHLY SAVINGS COLLECTION REPORT ALL FOR THE MONTH OF " & Mont & "-" & yr
    Else
        If cmbSamity_Code.Text = "ALL" Then
            str = "select * FROM Deposit_Master where FO_Code like '" & FO & "' AND Term like '" & Trm & "' order by Sl_No"
        Else
            str = "select * FROM Deposit_Master where FO_Code like '" & FO & "' AND Samity_Code like '" & Samity & "' AND Term like '" & Trm & "' order by Sl_No"
        End If
        rptSBMonthly.Label38.Caption = "MONTHLY SAVINGS COLLECTION REPORT TERM WISE FOR THE MONTH OF " & Mont & "-" & yr
    End If
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Deposit_Print", conn, adOpenDynamic, adLockOptimistic, -1
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
        rsN!Payment = rs!Payment
        rsN!Withdraw = rs!Withdraw
        rsN!Due = rs!Due
        rsN!Advance = rs!Advance
        rsN!Fine = rs!Fine
        rsN!Term_Fail = rs!Term_Fail
        
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
    If cmbTerm.Text = "All" Then
        If cmbSamity_Code.Text = "ALL" Then
            str = "select * FROM Deposit_Print where FO_Code like '" & FO & "' order by Sl_No"
        Else
            str = "select * FROM Deposit_Print where FO_Code like '" & FO & "' AND Samity_Code like '" & Samity & "' order by Sl_No"
        End If
    Else
        If cmbSamity_Code.Text = "ALL" Then
            str = "select * FROM Deposit_Print where FO_Code like '" & FO & "' AND Term like '" & Trm & "' order by Sl_No"
        Else
            str = "select * FROM Deposit_Print where FO_Code like '" & FO & "' AND Samity_Code like '" & Samity & "' AND Term like '" & Trm & "' order by Sl_No"
        End If
        
    End If
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        
    If Not rs.EOF Then
        rptSBMonthly.rsSBMonthly.ConnectionString = cnStr
        rptSBMonthly.rsSBMonthly.Source = str
        Timer1.Enabled = True
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
Timer1.Enabled = True
Command1.Enabled = False
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
   Dim Trm As String
    Dim Center As String
    Dim FO As String
    Dim Samity As String
        FO = cmbFO_Code.Text
        Samity = cmbSamity_Code.Text
        Trm = cmbTerm.Text
        Center = cmbCenter_Code.Text
    
If optCenter.Value = True Then
    
    If cmbCenter_Code.Text = "" Then
    Exit Sub
    End If
    
        
    Set rs = New ADODB.Recordset
    If cmbTerm.Text = "All" Then
    str = "select * FROM Deposit_Master where Center_Code like '" & Center & "'order by Sl_No"
    Else
    str = "select * FROM Deposit_Master where Center_Code like '" & Center & "'AND Term like '" & Trm & "'order by Sl_No"
    End If
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
   
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
    
End If

If optFO.Value = True Then
    If cmbFO_Code.Text = "" Then
    Exit Sub
    End If
    
    Set rs = New ADODB.Recordset
    If cmbTerm.Text = "All" Then
        If cmbSamity_Code.Text = "ALL" Then
            str = "select * FROM Deposit_Master where FO_Code like '" & FO & "' order by Sl_No"
        Else
            str = "select * FROM Deposit_Master where FO_Code like '" & FO & "' AND Samity_Code like '" & Samity & "' order by Sl_No"
        End If
    Else
        If cmbSamity_Code.Text = "ALL" Then
            str = "select * FROM Deposit_Master where FO_Code like '" & FO & "' AND Term like '" & Trm & "' order by Sl_No"
        Else
            str = "select * FROM Deposit_Master where FO_Code like '" & FO & "' AND Samity_Code like '" & Samity & "' AND Term like '" & Trm & "' order by Sl_No"
        End If
    End If
    rs.Open str, conn
        
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
        Command1.Enabled = False
    End If
End If

    Exit Sub
    End Sub

Private Sub Form_Load()
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

        Call ComboTerm
      
        Call ComboCenter_Code
        Call ComboCenter_Name
        Call ComboSamity_Code
        Call ComboSamity_Name
        Call ComboFO_Code
        Call ComboFO_Name
        cmbTerm.Text = "All"
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

If cmbSamity_Code.Text = "ALL" Then
cmbSamity_Name.Text = "ALL"
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

If cmbSamity_Name.Text = "ALL" Then
cmbSamity_Code.Text = "ALL"
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

Private Sub optFO_GotFocus()
cmbFO_Code.Visible = True
cmbFO_Name.Visible = True
lblFO_Code.Visible = True
lblFo_Name.Visible = True

cmbCenter_Code.Visible = True
cmbCenter_Name.Visible = True
lblCenter_Code.Visible = True
lblCenter_Name.Visible = True

cmbSamity_Code.Visible = True
cmbSamity_Name.Visible = True
lblSamity_Code.Visible = True
lblSamity_Name.Visible = True
End Sub

Private Sub optFO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbFO_Code.SetFocus
End If
End Sub

Private Sub optSamity_GotFocus()
cmbFO_Code.Visible = False
cmbFO_Name.Visible = False
lblFO_Code.Visible = False
lblFo_Name.Visible = False

cmbCenter_Code.Visible = False
cmbCenter_Name.Visible = False
lblCenter_Code.Visible = False
lblCenter_Name.Visible = False

cmbSamity_Code.Visible = True
cmbSamity_Name.Visible = True
lblSamity_Code.Visible = True
lblSamity_Name.Visible = True
End Sub

Private Sub optSamity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbSamity_Code.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
If optFO.Value = True Then
    ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptSBMonthly.Show 1
    End If
End If

End Sub
Private Sub txtTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
optFO.SetFocus
End If
End Sub

