VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStaff_Print 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Report"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "frmStaff_Print.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2130
      Left            =   135
      ScaleHeight     =   2070
      ScaleWidth      =   5535
      TabIndex        =   11
      Top             =   3015
      Width           =   5595
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmStaff_Print.frx":0442
         Height          =   1785
         Left            =   135
         TabIndex        =   12
         Top             =   135
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   3149
         _Version        =   393216
         BackColor       =   16777215
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   135
      ScaleHeight     =   180
      ScaleWidth      =   5535
      TabIndex        =   9
      Top             =   2655
      Width           =   5595
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   2415
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   5595
      Begin VB.OptionButton optAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All Center Staff Report"
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
         Left            =   360
         TabIndex        =   13
         Top             =   450
         Width           =   2310
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
         Left            =   585
         TabIndex        =   6
         Top             =   1575
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
         Left            =   1800
         TabIndex        =   5
         Text            =   "Combo2"
         Top             =   1575
         Width           =   1815
      End
      Begin VB.OptionButton optCenter 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Center wise Staff Report"
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
         Left            =   360
         TabIndex        =   4
         Top             =   945
         Width           =   2445
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   420
         Left            =   3915
         TabIndex        =   3
         Top             =   360
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   420
         Left            =   3915
         TabIndex        =   2
         Top             =   1620
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   420
         Left            =   3915
         TabIndex        =   1
         Top             =   990
         Width           =   1365
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
         Left            =   585
         TabIndex        =   8
         Top             =   1305
         Width           =   1080
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
         Left            =   1845
         TabIndex        =   7
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         Height          =   1950
         Left            =   3780
         Top             =   270
         Width           =   1665
      End
      Begin VB.Shape Shape2 
         Height          =   1965
         Left            =   180
         Top             =   270
         Width           =   3510
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   -135
      Top             =   2475
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   420
      Left            =   2295
      Top             =   4905
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
Attribute VB_Name = "frmStaff_Print"
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
Private Sub Command1_Click()
    Dim Today As Date
    Dim Mont As String
    Dim yr As String
    Dim Center As String
    
    
    
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
    
        Center = cmbCenter_Code.Text
        sort = 0
    
    Set rs = New ADODB.Recordset
    
        If optAll.Value = True Then
            str = "select * FROM Staff_Master Order by Center_Code, Staff_ID"
            rptStaff.Label11.Caption = "MONTHLY STAFF REPORT FOR THE MONTH OF " & Mont & "-" & yr
        Else
        If optCenter.Value = True Then
            If cmbCenter_Code.Text = "" Then
                Exit Sub
                MsgBox " Please Input Center code.", vbCritical, "Error"
                cmbCenter_Code.SetFocus
            Else
                str = "select * FROM Staff_Master where Center_Code like '" & Center & "' Order by Staff_ID"
                rptStaff.Label11.Caption = "MONTHLY CENTER WISE STAFF REPORT FOR THE MONTH OF " & Mont & "-" & yr
            End If
        End If
        End If
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
        If Not rs.EOF Then
            
            rptStaff.rsStaff.ConnectionString = cnStr
            rptStaff.rsStaff.Source = str
            rs.Close
            Timer1.Enabled = True
            
         Else
            MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
            rs.Close
         End If
    
Command1.Enabled = False


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    Dim Center As String
    Center = cmbCenter_Code.Text
    Set rs = New ADODB.Recordset
    
        If optAll.Value = True Then
           str = "select * FROM Staff_Master Order by Center_Code and Staff_ID"
        Else
        If optCenter.Value = True Then
            If cmbCenter_Code.Text = "" Then
                Exit Sub
                MsgBox " Please Input Center code.", vbCritical, "Error"
                cmbCenter_Code.SetFocus
            Else
                str = "select * FROM Staff_Master where Center_Code like '" & Center & "' order by Center_Code and Staff_ID"
            End If
        End If
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
            Command1.Enabled = False
         End If
    
    Exit Sub
    End Sub

Private Sub Form_Load()
        
        cmbCenter_Code.Visible = False
        cmbCenter_Name.Visible = False
        lblCenter_Code.Visible = False
        lblCenter_Name.Visible = False
        
        Call ComboCenter_Code
        Call ComboCenter_Name
        
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
Command3.SetFocus
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

Private Sub optAll_Click()
    cmbCenter_Code.Visible = False
    cmbCenter_Name.Visible = False
    lblCenter_Code.Visible = False
    lblCenter_Name.Visible = False
End Sub

Private Sub optCenter_GotFocus()
cmbCenter_Code.Visible = True
cmbCenter_Name.Visible = True
lblCenter_Code.Visible = True
lblCenter_Name.Visible = True
End Sub

Private Sub optCenter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbCenter_Code.SetFocus
End If
End Sub
Private Sub Timer1_Timer()
    ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptStaff.Show 1
    End If

End Sub

