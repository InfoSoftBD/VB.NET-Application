VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFixed_stm 
   BackColor       =   &H00004080&
   Caption         =   "Fixed Asset statement"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   Icon            =   "frmFixed_stm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   2385
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   330
      Left            =   135
      ScaleHeight     =   270
      ScaleWidth      =   7455
      TabIndex        =   9
      Top             =   3015
      Width           =   7515
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   0
         TabIndex        =   10
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
      Begin VB.TextBox txtId 
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
         TabIndex        =   14
         Top             =   360
         Width           =   1350
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
         Left            =   1800
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   900
         Width           =   2115
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
         Left            =   1800
         TabIndex        =   12
         Text            =   "Combo2"
         Top             =   1440
         Width           =   2115
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
         Left            =   1800
         TabIndex        =   5
         Top             =   2025
         Width           =   1485
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
         Left            =   3765
         TabIndex        =   4
         Top             =   2025
         Width           =   1485
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asset Code"
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
         Left            =   555
         TabIndex        =   17
         Top             =   450
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asset Model"
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
         Left            =   480
         TabIndex        =   16
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asset Name"
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
         Left            =   495
         TabIndex        =   15
         Top             =   1035
         Width           =   1080
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
         Left            =   495
         TabIndex        =   8
         Top             =   2025
         Width           =   405
      End
      Begin VB.Shape Shape2 
         Height          =   2310
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
         Left            =   1125
         TabIndex        =   7
         Top             =   2040
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
         Left            =   3405
         TabIndex        =   6
         Top             =   2070
         Width           =   210
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFixed_stm.frx":0442
      Height          =   2775
      Left            =   135
      TabIndex        =   11
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
Attribute VB_Name = "frmFixed_stm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub Prod_Name()
        cmbProd_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name FROM Fixed_Master"
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
        str = "SELECT DISTINCT Prod_Model FROM Fixed_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbProd_Model.AddItem rsN!Prod_Model
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub cmbProd_Model_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtFrom.SelStart = 0
    txtFrom.SelLength = Len(txtFrom.Text)
    txtFrom.SetFocus
End If
End Sub
Private Sub cmbProd_Model_LostFocus()
Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Fixed_Master where Prod_Name Like '" & cmbProd_Name.Text & "' AND Prod_Model Like '" & cmbProd_Model.Text & "' order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtID.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
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
        str = "SELECT DISTINCT Prod_Model FROM Fixed_Master where Prod_Name Like '" & Prod & "'"
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
    
    Prod = txtID.Text

    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Fixed_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtID.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            rsN.Close
        Else
        Exit Sub
        End If

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
    ID = txtID.Text
    Fromdate = txtFrom.Text
    Today = txtTo.Text
   
On Error Resume Next
   
     Set rs = New ADODB.Recordset
        str = "select * from Fixed_Tran where cdate(Date) Between cdate('" & Fromdate & "') And cdate('" & Today & "')and Prod_Code like '" & txtID.Text & "' order by Sl"
        rs.Open str, conn
        
    If Not rs.EOF Then
    rptFixed_Asset.rsFixed_stm.ConnectionString = cnStr
    rptFixed_Asset.rsFixed_stm.Source = str

        rs.Close
    Else
        MsgBox "There is no such product found, ", vbCritical + vbOKOnly
        rs.Close
    End If
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
        str = "select * from Fixed_Tran where cdate(Date) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Prod_Code like '" & ID & "' order by Sl "
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

Private Sub Form_Load()
Call Prod_Name
Call Prod_Model
txtFrom.Text = Date
txtTo.Text = Date
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Visible = True
    ProgressBar1.Max = 100
    ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = ProgressBar1.Max Then
    ProgressBar1.Value = 0
    Timer1.Enabled = False
    ProgressBar1.Visible = False
    rptFixed_Asset.Show 1
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProd_Name.SetFocus
End If
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTo.SetFocus
End If
End Sub

Private Sub txtId_LostFocus()
On Error Resume Next

    Dim Prod As String
    Prod = txtID.Text
    
    
    
    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Fixed_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtID.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            rsN.Close
        Else
        MsgBox "There is no such Asset Code found, ", vbCritical + vbOKOnly
        rsN.Close
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
        str = "select * from Fixed_Tran where cdate(Date) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Prod_Code like '" & ID & "' order by sl"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such Asset Code found, ", vbCritical + vbOKOnly
        rs.Close
    End If
    End If
    Exit Sub

End Sub




