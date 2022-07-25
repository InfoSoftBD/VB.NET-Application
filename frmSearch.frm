VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmCash 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Flow Statement"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   135
      ScaleHeight     =   270
      ScaleWidth      =   6825
      TabIndex        =   11
      Top             =   2295
      Width           =   6885
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   6855
      Begin VB.TextBox txtToday 
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
         Left            =   3030
         TabIndex        =   10
         Top             =   450
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date wise"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1260
         Width           =   1155
      End
      Begin VB.OptionButton optToday 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print Toadys Transaction"
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
         Left            =   240
         TabIndex        =   8
         Top             =   540
         Width           =   2505
      End
      Begin VB.TextBox txtFrom 
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
         Left            =   2040
         TabIndex        =   4
         Top             =   1215
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtTo 
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
         Left            =   3585
         TabIndex        =   3
         Top             =   1215
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   420
         Left            =   5220
         TabIndex        =   2
         Top             =   420
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   420
         Left            =   5220
         TabIndex        =   1
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Shape Shape2 
         Height          =   1545
         Left            =   120
         Top             =   270
         Width           =   4830
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Left            =   5085
         Top             =   270
         Width           =   1635
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
         Left            =   1485
         TabIndex        =   6
         Top             =   1275
         Visible         =   0   'False
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
         Left            =   3285
         TabIndex        =   5
         Top             =   1260
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSearch.frx":0442
      Height          =   2595
      Left            =   135
      TabIndex        =   7
      Top             =   2775
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   4577
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
Attribute VB_Name = "frmCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub Command1_Click()
   
If optToday.Value = True Then
Dim sort As Integer
    sort = 0
        
       If txtToday.Text = "" Then
        Exit Sub
        End If
        
    Dim Today As Date
    
    Today = txtToday.Text
  

On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Cash_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Cash_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If


Set rs = New ADODB.Recordset
    str = "select * from Cash_Book where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') order by Sl"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsU = New ADODB.Recordset
        rsU.Open "Cash_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsU.AddNew
            rsU!sl = 1
            rsU!Sl_no = sort + 1
            rsU!Date = rs!Date
            rsU!Name = rs!Name
            rsU!MR_No = rs!MR_No
            rsU!Description = rs!Description
            rsU!Cr = rs!Cr
            rsU!Dr = rs!Dr
            rsU!Balance = rs!Balance
            rsU.Update
            sort = sort + 1
            rs.MoveNext
            Loop
            rs.Close
            rsU.Close

Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
        
            Set rsU = New ADODB.Recordset
            str = "select * from Cash_Print"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
        
        If Not rsU.EOF Then
            rsU!Cash_Open = rs!Cash_Open
            rsU!Cash_Dr = rs!Cash_Dr
            rsU!Cash_Cr = rs!Cash_Cr
            rsU!Cash_Close = rs!Cash_Close
            rsU.Update
            rsU.Close
            rs.Close
        End If
    End If

str = "select * from Cash_Print order by Sl_No"
rptCash.rsCash.ConnectionString = cnStr
rptCash.rsCash.Source = str
Command1.Enabled = False

Timer1.Enabled = True
End If

If optDate.Value = True Then
        
    If txtFrom.Text = "" Then
    Exit Sub
    End If
    If txtTo.Text = "" Then
    Exit Sub
    End If

Dim Fromdate As Date
Dim toDate As Date
    Fromdate = txtFrom.Text
    toDate = txtTo.Text
  

str = "select * from Cash_Book where cdate(Date) Between cdate('" & Fromdate & "') And cdate('" & toDate & "')"
rs.Open str, conn

rptCash.rsCash.ConnectionString = cnStr
rptCash.rsCash.Source = str
Command1.Enabled = False
Timer1.Enabled = True
End If



End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Command1.Enabled = False
End Sub

Private Sub optDate_Click()
txtToday.Visible = False
txtFrom.Visible = True
txtFrom.Text = ""
txtFrom.SetFocus
txtTo.Visible = True
txtTo.Text = ""
Label1.Visible = True
Label2.Visible = True
End Sub

Private Sub optToday_Click()
txtToday.Visible = True
txtFrom.Visible = False
txtTo.Visible = False
Label1.Visible = False
Label2.Visible = False
txtToday.Text = Date
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Visible = True
    ProgressBar1.Max = 100
    ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = ProgressBar1.Max Then
    ProgressBar1.Value = 0
    Timer1.Enabled = False
    ProgressBar1.Visible = False
    rptCash.Show 1
End If
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTo.SetFocus
End If
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtFrom.Text = "" Then
Exit Sub
End If
If txtTo.Text = "" Then
Exit Sub
End If
On Error Resume Next
Dim Fromdate As Date
Dim toDate As Date
    Fromdate = txtFrom.Text
    toDate = txtTo.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Cash_Book where cdate(Date) Between cdate('" & Fromdate & "') And cdate('" & toDate & "')"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
        Command1.Enabled = True
    Else
        MsgBox "There is no such Parts Name found, ", vbCritical + vbOKOnly
        rs.Close
    End If
    End If
    Exit Sub

End Sub
Private Sub txtToday_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtToday.Text = "" Then
Exit Sub
End If

Dim Today As Date
    Today = txtToday.Text
   
    Set rs = New ADODB.Recordset
        str = "select * from Cash_Book where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') order by Sl"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
        Command1.Enabled = True
    Else
        MsgBox "There is no such Transaction found, ", vbCritical + vbOKOnly
        rs.Close
    End If
    End If
    Exit Sub

End Sub
