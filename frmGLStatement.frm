VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "Mscal.ocx"
Begin VB.Form frmGLStatement 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GL Account Statement"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7755
   Icon            =   "frmGLStatement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin MSACAL.Calendar Calendar2 
      Height          =   2265
      Left            =   3420
      TabIndex        =   16
      Top             =   2430
      Width           =   2745
      _Version        =   524288
      _ExtentX        =   4842
      _ExtentY        =   3995
      _StockProps     =   1
      BackColor       =   16777215
      Year            =   2012
      Month           =   5
      Day             =   19
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   6
      GridCellEffect  =   1
      GridFontColor   =   0
      GridLinesColor  =   0
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   0
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2265
      Left            =   720
      TabIndex        =   17
      Top             =   2430
      Width           =   2745
      _Version        =   524288
      _ExtentX        =   4842
      _ExtentY        =   3995
      _StockProps     =   1
      BackColor       =   16777215
      Year            =   2012
      Month           =   5
      Day             =   19
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   6
      GridCellEffect  =   1
      GridFontColor   =   0
      GridLinesColor  =   0
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   0
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   2385
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   135
      ScaleHeight     =   270
      ScaleWidth      =   7455
      TabIndex        =   7
      Top             =   3015
      Width           =   7515
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   2730
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   7485
      Begin VB.ComboBox txtID 
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
         Left            =   1740
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   480
         Width           =   1545
      End
      Begin VB.ComboBox txtName 
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
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   1170
         Width           =   3330
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   4740
         Picture         =   "frmGLStatement.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1800
         Width           =   345
      End
      Begin VB.CommandButton Command4 
         Height          =   375
         Left            =   1440
         Picture         =   "frmGLStatement.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1800
         Width           =   345
      End
      Begin VB.TextBox txtFrom 
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
         Left            =   1800
         TabIndex        =   11
         Top             =   1815
         Width           =   1275
      End
      Begin VB.TextBox txtTo 
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
         Left            =   3375
         TabIndex        =   10
         Top             =   1815
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   420
         Left            =   5850
         TabIndex        =   3
         Top             =   1140
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   420
         Left            =   5850
         TabIndex        =   2
         Top             =   1845
         Width           =   1365
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   420
         Left            =   5850
         TabIndex        =   1
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
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
         Left            =   960
         TabIndex        =   15
         Top             =   1875
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
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
         Left            =   3135
         TabIndex        =   14
         Top             =   1875
         Width           =   210
      End
      Begin VB.Label Label5 
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
         Left            =   315
         TabIndex        =   6
         Top             =   1890
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Head Name"
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
         Left            =   525
         TabIndex        =   5
         Top             =   1260
         Width           =   1020
      End
      Begin VB.Label Label3 
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
         Left            =   585
         TabIndex        =   4
         Top             =   525
         Width           =   1080
      End
      Begin VB.Shape Shape2 
         Height          =   2220
         Left            =   135
         Top             =   270
         Width           =   5355
      End
      Begin VB.Shape Shape1 
         Height          =   2220
         Left            =   5715
         Top             =   270
         Width           =   1635
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmGLStatement.frx":0CC6
      Height          =   2775
      Left            =   135
      TabIndex        =   9
      Top             =   3435
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   4895
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   420
      Left            =   405
      Top             =   4950
      Visible         =   0   'False
      Width           =   6405
      _ExtentX        =   11298
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
Attribute VB_Name = "frmGLStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub CmbAc()
        txtID.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT AC_No FROM GL_Master"
        rs.Open str, conn
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        txtID.AddItem rs!AC_No
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub CmbName()
        txtName.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Head_Name FROM GL_Master"
        rs.Open str, conn
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        txtName.AddItem rs!Head_Name
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub Calendar1_Click()
txtFrom.Text = Calendar1.Value
Calendar1.Visible = False
End Sub

Private Sub Calendar2_Click()
txtTo.Text = Calendar2.Value
Calendar2.Visible = False
End Sub

Private Sub Command1_Click()
    If txtID.Text = "" Then
    MsgBox "Please Enter Member ID.", vbOKOnly
    Exit Sub
    End If
    If txtFrom.Text = "" Then
    MsgBox "Please Enter Start Date.", vbOKOnly
    Exit Sub
    End If
    If txtTo.Text = "" Then
    MsgBox "Please Enter End Date.", vbOKOnly
    Exit Sub
    End If

Dim ID As String
Dim Fromdate As Date
Dim Today As Date
Dim sort As Integer
    ID = txtID.Text
    Fromdate = txtFrom.Text
    Today = txtTo.Text
    sort = 0


On Error Resume Next
   

Set rs = New ADODB.Recordset
        str = "select * from GL_Tran where cdate(Date) Between cdate('" & Fromdate & "') And cdate('" & Today & "')and AC_No like '" & ID & "' order by Sl"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

    rptGL_Statement.rsGL.ConnectionString = cnStr
    rptGL_Statement.rsGL.Source = str

Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If txtID.Text = "" Then
Exit Sub
End If
If txtFrom.Text = "" Then
Exit Sub
End If
If txtTo.Text = "" Then
Exit Sub
End If

Dim ID As String
Dim Fromdate As Date
Dim Today As Date
    ID = txtID.Text
    Fromdate = txtFrom.Text
    Today = txtTo.Text
    
     Set rs = New ADODB.Recordset
        str = "select * from GL_Tran where cdate(Date) Between cdate('" & Fromdate & "') And cdate('" & Today & "')and AC_No like '" & ID & "' order by Sl"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such Member Id found, ", vbCritical + vbOKOnly
        rs.Close
    End If
    Exit Sub
    End Sub

Private Sub Command4_Click()
Calendar1.Visible = True
Calendar1.Value = Today
End Sub

Private Sub Command5_Click()
Calendar2.Visible = True
Calendar2.Value = Today
End Sub
Private Sub Form_Load()
Call CmbAc
Call CmbName
txtTo.Text = Today
txtFrom.Text = Today
Calendar1.Visible = False
Calendar1.Value = Today
Calendar2.Visible = False
Calendar2.Value = Today
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Visible = True
    ProgressBar1.Max = 100
    ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = ProgressBar1.Max Then
    ProgressBar1.Value = 0
    Timer1.Enabled = False
    ProgressBar1.Visible = False
    rptGL_Statement.Show 1
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtID.Text = "" Then
Exit Sub
End If

Dim ID As String
    ID = txtID.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Tran where AC_No like '" & ID & "'"
        rs.Open str, conn
        
    If Not rs.EOF Then
        txtID.Text = rs!AC_No
        txtName.Text = rs!Name
        txtTo.Text = Date
        txtFrom.SetFocus
    Else
        MsgBox "Transaction Transaction found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
    End If
    Exit Sub
End Sub

Private Sub txtAccount_Change()

End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTo.SetFocus
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtFrom.SetFocus
End If
End Sub

Private Sub txtName_LostFocus()
If txtName.Text = "" Then
Exit Sub
End If

Dim ID As String
    ID = txtName.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from GL_Master where Head_Name like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
    'Call clearTextboxes
    On Error Resume Next
        txtID.Text = rs!AC_No
        txtFrom.Text = Today
        txtTo.Text = Today
        txtName.Text = rs!Head_Name
        rs.Close
        
    Else
    MsgBox "There is no such Account no. found.,", vbCritical
        rs.Close
        'Call clearTextboxes
        txtID.SetFocus
    End If
    Exit Sub
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtID.Text = "" Then
Exit Sub
End If
If txtFrom.Text = "" Then
Exit Sub
End If
If txtTo.Text = "" Then
Exit Sub
End If

Dim ID As String
Dim Fromdate As Date
Dim Today As Date
    ID = txtID.Text
    Fromdate = txtFrom.Text
    Today = txtTo.Text
    
     Set rs = New ADODB.Recordset
        str = "select * from GL_Tran where cdate(Date) Between cdate('" & Fromdate & "') And cdate('" & Today & "')and AC_No like '" & ID & "' order by Sl"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such Member Id found, ", vbCritical + vbOKOnly
        rs.Close
    End If
    End If
    Exit Sub

End Sub


