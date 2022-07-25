VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmProd_Statement 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Statement"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   Icon            =   "frmProd_Statement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin MSACAL.Calendar Calendar2 
      Height          =   2265
      Left            =   2970
      TabIndex        =   17
      Top             =   3330
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
      Left            =   270
      TabIndex        =   18
      Top             =   3330
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
      Height          =   3315
      Left            =   90
      TabIndex        =   2
      Top             =   135
      Width           =   7485
      Begin VB.ComboBox cmbAgent_Code 
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
         ItemData        =   "frmProd_Statement.frx":0442
         Left            =   4230
         List            =   "frmProd_Statement.frx":0458
         TabIndex        =   27
         Top             =   817
         Width           =   1170
      End
      Begin VB.OptionButton optSummary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Summary Report"
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
         Left            =   360
         TabIndex        =   26
         Top             =   810
         Width           =   1770
      End
      Begin VB.ComboBox txtId 
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
         Left            =   360
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   2190
         Width           =   1545
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
         Left            =   3405
         TabIndex        =   22
         Top             =   2790
         Width           =   1365
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
         Left            =   1740
         TabIndex        =   21
         Top             =   2790
         Width           =   1275
      End
      Begin VB.CommandButton Command4 
         Height          =   375
         Left            =   1320
         Picture         =   "frmProd_Statement.frx":048D
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2790
         Width           =   345
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   4800
         Picture         =   "frmProd_Statement.frx":08CF
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2790
         Width           =   345
      End
      Begin VB.ComboBox cmbType 
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
         ItemData        =   "frmProd_Statement.frx":0D11
         Left            =   2295
         List            =   "frmProd_Statement.frx":0D27
         TabIndex        =   14
         Top             =   817
         Width           =   1845
      End
      Begin VB.OptionButton optProduct 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product wise Report"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1530
         Width           =   2085
      End
      Begin VB.OptionButton optAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Periodic Report"
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
         Left            =   360
         TabIndex        =   12
         Top             =   330
         Width           =   1680
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
         Left            =   3840
         TabIndex        =   9
         Text            =   "Combo2"
         Top             =   2160
         Width           =   1485
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
         Left            =   2085
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   2160
         Width           =   1485
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   420
         Left            =   5850
         TabIndex        =   5
         Top             =   450
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   420
         Left            =   5850
         TabIndex        =   4
         Top             =   2025
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   420
         Left            =   5850
         TabIndex        =   3
         Top             =   1230
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Code"
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
         TabIndex        =   28
         Top             =   397
         Width           =   1020
      End
      Begin VB.Shape Shape3 
         Height          =   1215
         Left            =   210
         Top             =   1470
         Width           =   5265
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
         Left            =   3105
         TabIndex        =   24
         Top             =   2850
         Width           =   210
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
         Left            =   840
         TabIndex        =   23
         Top             =   2850
         Width           =   450
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
         Left            =   330
         TabIndex        =   16
         Top             =   2850
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Category"
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
         TabIndex        =   15
         Top             =   397
         Width           =   1410
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
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
         Left            =   2085
         TabIndex        =   11
         Top             =   1890
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Model"
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
         Left            =   3840
         TabIndex        =   10
         Top             =   1890
         Width           =   1260
      End
      Begin VB.Shape Shape1 
         Height          =   2415
         Left            =   5715
         Top             =   270
         Width           =   1635
      End
      Begin VB.Shape Shape2 
         Height          =   1035
         Left            =   210
         Top             =   240
         Width           =   5265
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
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
         TabIndex        =   6
         Top             =   1890
         Width           =   1185
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   90
      ScaleHeight     =   195
      ScaleWidth      =   7485
      TabIndex        =   0
      Top             =   3495
      Width           =   7515
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   -45
      Top             =   2340
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmProd_Statement.frx":0D5B
      Height          =   2730
      Left            =   90
      TabIndex        =   7
      Top             =   3780
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   4815
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
      Left            =   360
      Top             =   4905
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
Attribute VB_Name = "frmProd_Statement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub Prod_Code()
        txtId.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Code FROM Prod_Master"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtId.AddItem rs!Prod_Code
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub Prod_Name()
        cmbProd_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name FROM Prod_Master"
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
        str = "SELECT DISTINCT Prod_Model FROM Prod_Master"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        cmbProd_Model.AddItem rsN!Prod_Model
        rsN.MoveNext
        Loop
        rsN.Close
End Sub
Private Sub Agent_Id()
   cmbAgent_Code.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_Id FROM Employee"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbAgent_Code.AddItem rs!Staff_Id
        rs.MoveNext
        Loop
        rs.Close
    Else
    rs.Close
    End If
End Sub

'Private Sub Branch()
'        cmbBranch.Clear
'         Set rsN = New ADODB.Recordset
'        str = "SELECT DISTINCT Branch_Name FROM Branch"
'        rsN.Open str, conn
'        rsN.MoveFirst
'
'    Do While Not rsN.EOF
'        On Error Resume Next
'        cmbBranch.AddItem rsN!Branch_Name
'        rsN.MoveNext
'        Loop
'        rsN.Close
'End Sub

Private Sub Calendar1_Click()
txtFrom.Text = Calendar1.Value
Calendar1.Visible = False
End Sub

Private Sub Calendar2_Click()
txtTo.Text = Calendar2.Value
Calendar2.Visible = False
End Sub
Private Sub cmbBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtFrom.SelStart = 0
txtFrom.SelLength = Len(txtFrom.Text)
txtFrom.SetFocus
End If
End Sub

Private Sub cmbProd_Model_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtFrom.SetFocus
End If
End Sub
Private Sub cmbProd_Model_LostFocus()
Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Name Like '" & cmbProd_Name.Text & "' AND Prod_Model Like '" & cmbProd_Model.Text & "' order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtId.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
           
            rsN.Close
        Else
        rsN.Close
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
        str = "SELECT DISTINCT Prod_Model FROM Prod_Master where Prod_Name Like '" & Prod & "'"
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



Private Sub cmbType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmbType.Text = "Issue" Or cmbType.Text = "Return" Then
'Label4.Visible = True
'cmbBranch.Visible = True
'Call Branch
'cmbBranch.SelStart = 0
'cmbBranch.SelLength = Len(cmbBranch.Text)
'cmbBranch.SetFocus
'Else
'Label4.Visible = False
'cmbBranch.Visible = False
'cmbBranch.Clear

txtFrom.SelStart = 0
txtFrom.SelLength = Len(txtFrom.Text)
txtFrom.SetFocus

End If
End If
End Sub


Private Sub Command1_Click()
If optAll.Value = True Then

    Dim Fromdate As Date
    Dim Today As Date
    
     If txtFrom.Text = "" Then
        MsgBox "Please Input Begining Date", vbCritical
        txtFrom.Text = Today
        txtFrom.SelStart = 0
        txtFrom.SelLength = Len(txtFrom.Text)
        txtFrom.SetFocus
        Exit Sub
    End If
    
    If txtTo.Text = "" Then
        MsgBox "Please Input End Date", vbCritical
        txtTo.Text = Today
        txtTo.SelStart = 0
        txtTo.SelLength = Len(txtTo.Text)
        txtTo.SetFocus
        Exit Sub
    End If

    Fromdate = txtFrom.Text
    Today = txtTo.Text
    
If cmbType.Text = "All" Then
            Set rs = New ADODB.Recordset
                str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') order by Sl"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rptProd_Statement.rsProd_Statement.ConnectionString = cnStr
                rptProd_Statement.rsProd_Statement.Source = str
                rs.Close
            Else
                MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                rs.Close
            End If
End If
       
If cmbType.Text = "Purchase" Then
            Set rs = New ADODB.Recordset
                str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Purchase > 0 order by Sl"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rptPurchase.rsIssue.ConnectionString = cnStr
                rptPurchase.rsIssue.Source = str
                rptPurchase.Label11.Caption = "PRODUCT PURCHASE REPORT"
                rptPurchase.Field19.DataField = "Purchase"
                rptPurchase.Field16.DataField = "Prod_Price"
                rs.Close
            Else
                MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                rs.Close
            End If
End If
       
If cmbType.Text = "Sale" Then
            Set rs = New ADODB.Recordset
                str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and sale > 0 order by Prod_Code"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rptPurchase.rsIssue.ConnectionString = cnStr
                rptPurchase.rsIssue.Source = str
                rptPurchase.Label11.Caption = "PRODUCT SALES REPORT"
                rptPurchase.Field19.DataField = "Sale"
                rptPurchase.Field31.DataField = "Sale"
                rptPurchase.Field16.DataField = "Sale_Price"
                rs.Close
            Else
                MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                rs.Close
            End If
End If


If cmbType.Text = "Agent Sale" Then
    If cmbAgent_Code.Text = "" Then
        Exit Sub
        Else
            Set rs = New ADODB.Recordset
                str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and sale > 0 and Ref_no like '" & cmbAgent_Code.Text & "' order by Prod_Code"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rptPurchase.rsIssue.ConnectionString = cnStr
                rptPurchase.rsIssue.Source = str
                rptPurchase.Label11.Caption = "PRODUCT SALES REPORT"
                rptPurchase.Field19.DataField = "Sale"
                rptPurchase.Field31.DataField = "Sale"
                rptPurchase.Field16.DataField = "Sale_Price"
                rs.Close
            Else
                MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                rs.Close
            End If
    End If
End If
       
If cmbType.Text = "Issue" Then
        Set rs = New ADODB.Recordset
            str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Lift > 0 order by Sl"
            rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
            rptIssue.rsIssue.ConnectionString = cnStr
            rptIssue.rsIssue.Source = str
            rptIssue.Label11.Caption = "PRODUCT ISSUE REPORT"
            rptIssue.Field19.DataField = "Issue"
            rs.Close
        Else
            MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
            rs.Close
        End If
End If
If cmbType.Text = "Return" Then
        Set rs = New ADODB.Recordset
            str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Return > 0 order by Sl"
            rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
            rptIssue.rsIssue.ConnectionString = cnStr
            rptIssue.rsIssue.Source = str
            rptIssue.Label11.Caption = "PRODUCT RETURN REPORT"
            rptIssue.Field19.DataField = "Return"
            rs.Close
        Else
           MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
           rs.Close
        End If
End If

End If
'----------------------------------------------------------------------------------
If optSummary.Value = True Then
    
    If txtFrom.Text = "" Then
        MsgBox "Please Input Begining Date", vbCritical
        txtFrom.Text = Today
        txtFrom.SelStart = 0
        txtFrom.SelLength = Len(txtFrom.Text)
        txtFrom.SetFocus
        Exit Sub
    End If
    
    If txtTo.Text = "" Then
        MsgBox "Please Input End Date", vbCritical
        txtTo.Text = Today
        txtTo.SelStart = 0
        txtTo.SelLength = Len(txtTo.Text)
        txtTo.SetFocus
        Exit Sub
    End If

    Fromdate = txtFrom.Text
    Today = txtTo.Text
    
    If cmbType.Text = "All" Then
            Set rs = New ADODB.Recordset
                str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') order by Prod_Code"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rptProd_Statement.rsProd_Statement.ConnectionString = cnStr
                rptProd_Statement.rsProd_Statement.Source = str
                rs.Close
            Else
                MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                rs.Close
            End If
    End If
       
    If cmbType.Text = "Purchase" Then
        Set rs = New ADODB.Recordset
            str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Purchase > 0 order by Prod_Code"
            rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
            rptProd_Summary.rsIssue.ConnectionString = cnStr
            rptProd_Summary.rsIssue.Source = str
            rptProd_Summary.lblFrom.Caption = txtFrom.Text
            rptProd_Summary.lblTo.Caption = txtTo.Text
            rptProd_Summary.Label11.Caption = "PRODUCT PURCHASE SUMMARY REPORT"
            rptProd_Summary.Field31.DataField = "Purchase"
            'rptProd_Summary.Field130.DataField = "Prod_Price"
            rs.Close
        Else
            MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
            rs.Close
        End If
    End If
       
    If cmbType.Text = "Sale" Then
        Set rs = New ADODB.Recordset
            str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and sale > 0 order by Prod_Code"
            rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
            rptProd_Summary.rsIssue.ConnectionString = cnStr
            rptProd_Summary.rsIssue.Source = str
            rptProd_Summary.lblFrom.Caption = txtFrom.Text
            rptProd_Summary.lblTo.Caption = txtTo.Text
            rptProd_Summary.Label11.Caption = "PRODUCT SALES SUMMARY REPORT"
            rptProd_Summary.Field31.DataField = "Sale"
            'rptProd_Summary.Field30.DataField = "Sale_Price"
            rs.Close
        Else
            MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
            rs.Close
        End If
    End If
       
       
    If cmbType.Text = "Agent Sale" Then
    
        If cmbAgent_Code.Text = "" Then
        Exit Sub
        Else
        
        Set rs = New ADODB.Recordset
            str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and sale > 0 and Ref_no like '" & cmbAgent_Code.Text & "' order by Prod_Code"
            rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
            rptProd_Summary.rsIssue.ConnectionString = cnStr
            rptProd_Summary.rsIssue.Source = str
            rptProd_Summary.lblFrom.Caption = txtFrom.Text
            rptProd_Summary.lblTo.Caption = txtTo.Text
            rptProd_Summary.Label11.Caption = "PRODUCT SALES SUMMARY REPORT"
            rptProd_Summary.Field31.DataField = "Sale"
            'rptProd_Summary.Field30.DataField = "Sale_Price"
            rs.Close
        Else
            MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
            rs.Close
        End If
        End If
    End If
       
    If cmbType.Text = "Issue" Then
            Set rs = New ADODB.Recordset
                str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Lift > 0 order by Prod_Code"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rptProd_Summary.rsIssue.ConnectionString = cnStr
                rptProd_Summary.rsIssue.Source = str
                rptProd_Summary.lblFrom.Caption = txtFrom.Text
                rptProd_Summary.lblTo.Caption = txtTo.Text
                rptProd_Summary.Label11.Caption = "PRODUCT ISSUE SUMMARY REPORT"
                rptProd_Summary.Field31.DataField = "Lift"
                rs.Close
            Else
                MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                rs.Close
            End If
    End If
    
    If cmbType.Text = "Return" Then
            Set rs = New ADODB.Recordset
                str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Return > 0 order by Prod_Code"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rptProd_Summary.rsIssue.ConnectionString = cnStr
                rptProd_Summary.rsIssue.Source = str
                rptProd_Summary.lblFrom.Caption = txtFrom.Text
                rptProd_Summary.lblTo.Caption = txtTo.Text
                rptProd_Summary.Label11.Caption = "PRODUCT RETURN SUMMARY REPORT"
                rptProd_Summary.Field31.DataField = "Return"
                rs.Close
            Else
               MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
               rs.Close
            End If
    End If

End If

'--------------------------------------------------------------
If optProduct.Value = True Then

    Fromdate = txtFrom.Text
    Today = txtTo.Text
    
     Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "')and Prod_Code like '" & txtId.Text & "' order by Sl"
        rs.Open str, conn
        
    If Not rs.EOF Then
        rptProd_Statement.rsProd_Statement.ConnectionString = cnStr
        rptProd_Statement.rsProd_Statement.Source = str
        rs.Close
    Else
        MsgBox "There is no such product found, ", vbCritical + vbOKOnly
        rs.Close
    End If
End If
    
Timer1.Enabled = True
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If optAll.Value = True Then
    Dim Fromdate As Date
    Dim Today As Date
        
    If txtFrom.Text = "" Then
        MsgBox "Please Input Begining Date", vbCritical
        txtFrom.Text = Today
        txtFrom.SelStart = 0
        txtFrom.SelLength = Len(txtFrom.Text)
        txtFrom.SetFocus
        Exit Sub
    End If
    
    If txtTo.Text = "" Then
        MsgBox "Please Input End Date", vbCritical
        txtTo.Text = Today
        txtTo.SelStart = 0
        txtTo.SelLength = Len(txtTo.Text)
        txtTo.SetFocus
        Exit Sub
    End If
    

    Fromdate = txtFrom.Text
    Today = txtTo.Text
    
    If cmbType.Text = "All" Then
                Set rs = New ADODB.Recordset
                    str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "')order by Sl"
                    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                If Not rs.EOF Then
                    Adodc1.ConnectionString = cnStr
                    Adodc1.RecordSource = str
                    Adodc1.Refresh
                    DataGrid1.Refresh
                    rs.Close
                    Command1.Enabled = True
                Else
                    MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                    rs.Close
                    Command1.Enabled = False
                End If
    End If
       
    If cmbType.Text = "Purchase" Then
                Dim ps As String
                ps = ""
                Set rs = New ADODB.Recordset
                    str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Purchase > 0 order by Sl"
                    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                If Not rs.EOF Then
                    Adodc1.ConnectionString = cnStr
                    Adodc1.RecordSource = str
                    Adodc1.Refresh
                    DataGrid1.Refresh
                    rs.Close
                    Command1.Enabled = True
                Else
                    MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                    rs.Close
                    Command1.Enabled = False
                End If
    End If
       
    If cmbType.Text = "Sale" Then
                Set rs = New ADODB.Recordset
                    str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "')and sale > 0 order by Sl"
                    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                If Not rs.EOF Then
                    Adodc1.ConnectionString = cnStr
                    Adodc1.RecordSource = str
                    Adodc1.Refresh
                    DataGrid1.Refresh
                    rs.Close
                    Command1.Enabled = True
                Else
                    MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                    rs.Close
                    Command1.Enabled = False
                End If
    End If
    
    If cmbType.Text = "Agent Sale" Then
     If cmbAgent_Code.Text = "" Then
        Exit Sub
        Else
                Set rs = New ADODB.Recordset
                    str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and sale > 0 and Ref_no like '" & cmbAgent_Code.Text & "' order by Sl"
                    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                If Not rs.EOF Then
                    Adodc1.ConnectionString = cnStr
                    Adodc1.RecordSource = str
                    Adodc1.Refresh
                    DataGrid1.Refresh
                    rs.Close
                    Command1.Enabled = True
                Else
                    MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                    rs.Close
                    Command1.Enabled = False
                End If
        End If
    End If
       
       
    If cmbType.Text = "Issue" Then
            Set rs = New ADODB.Recordset
                str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Lift > 0 order by Sl"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                Adodc1.ConnectionString = cnStr
                Adodc1.RecordSource = str
                Adodc1.Refresh
                DataGrid1.Refresh
                rs.Close
                Command1.Enabled = True
            Else
                MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                rs.Close
                Command1.Enabled = False
            End If
    End If

    If cmbType.Text = "Return" Then
            Set rs = New ADODB.Recordset
                str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Return > 0 order by Sl"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                Adodc1.ConnectionString = cnStr
                Adodc1.RecordSource = str
                Adodc1.Refresh
                DataGrid1.Refresh
                rs.Close
                Command1.Enabled = True
            Else
               MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
               rs.Close
               Command1.Enabled = False
            End If
    End If

End If
'-------------------------------------------------------------
If optSummary.Value = True Then
            
    If txtFrom.Text = "" Then
        MsgBox "Please Input Begining Date", vbCritical
        txtFrom.Text = Today
        txtFrom.SelStart = 0
        txtFrom.SelLength = Len(txtFrom.Text)
        txtFrom.SetFocus
        Exit Sub
    End If
    
    If txtTo.Text = "" Then
        MsgBox "Please Input End Date", vbCritical
        txtTo.Text = Today
        txtTo.SelStart = 0
        txtTo.SelLength = Len(txtTo.Text)
        txtTo.SetFocus
        Exit Sub
    End If
    

    Fromdate = txtFrom.Text
    Today = txtTo.Text
    
    If cmbType.Text = "All" Then
                Set rs = New ADODB.Recordset
                    str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "')order by Sl"
                    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                If Not rs.EOF Then
                    Adodc1.ConnectionString = cnStr
                    Adodc1.RecordSource = str
                    Adodc1.Refresh
                    DataGrid1.Refresh
                    rs.Close
                    Command1.Enabled = True
                Else
                    MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                    rs.Close
                    Command1.Enabled = False
                End If
    End If
       
    If cmbType.Text = "Purchase" Then
                'Dim ps As String
                ps = ""
                Set rs = New ADODB.Recordset
                    str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Purchase > 0 order by Sl"
                    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                If Not rs.EOF Then
                    Adodc1.ConnectionString = cnStr
                    Adodc1.RecordSource = str
                    Adodc1.Refresh
                    DataGrid1.Refresh
                    rs.Close
                    Command1.Enabled = True
                Else
                    MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                    rs.Close
                    Command1.Enabled = False
                End If
    End If
       
    If cmbType.Text = "Sale" Then
                Set rs = New ADODB.Recordset
                    str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "')and sale > 0 order by Sl"
                    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                If Not rs.EOF Then
                    Adodc1.ConnectionString = cnStr
                    Adodc1.RecordSource = str
                    Adodc1.Refresh
                    DataGrid1.Refresh
                    rs.Close
                    Command1.Enabled = True
                Else
                    MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                    rs.Close
                    Command1.Enabled = False
                End If
    End If
    
    If cmbType.Text = "Agent Sale" Then
    
        If cmbAgent_Code.Text = "" Then
        Exit Sub
        Else
                Set rs = New ADODB.Recordset
                    str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "')and sale > 0 and Ref_no like '" & cmbAgent_Code.Text & "' order by Sl"
                    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                If Not rs.EOF Then
                    Adodc1.ConnectionString = cnStr
                    Adodc1.RecordSource = str
                    Adodc1.Refresh
                    DataGrid1.Refresh
                    rs.Close
                    Command1.Enabled = True
                Else
                    MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                    rs.Close
                    Command1.Enabled = False
                End If
        End If
    End If
       
       
    If cmbType.Text = "Issue" Then
            Set rs = New ADODB.Recordset
                str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Lift > 0 order by Sl"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                Adodc1.ConnectionString = cnStr
                Adodc1.RecordSource = str
                Adodc1.Refresh
                DataGrid1.Refresh
                rs.Close
                Command1.Enabled = True
            Else
                MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
                rs.Close
                Command1.Enabled = False
            End If
    End If

    If cmbType.Text = "Return" Then
            Set rs = New ADODB.Recordset
                str = "select * From Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "') and Return > 0 order by Sl"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                Adodc1.ConnectionString = cnStr
                Adodc1.RecordSource = str
                Adodc1.Refresh
                DataGrid1.Refresh
                rs.Close
                Command1.Enabled = True
            Else
               MsgBox "There is no such Product found, ", vbCritical + vbOKOnly
               rs.Close
               Command1.Enabled = False
            End If
    End If

End If

'--------------------------------------------------------------
If optProduct.Value = True Then

    Fromdate = txtFrom.Text
    Today = txtTo.Text
    
     Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "')and Prod_Code like '" & txtId.Text & "' order by Sl"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
        Command1.Enabled = True
    Else
        MsgBox "There is no such product found, ", vbCritical + vbOKOnly
        rs.Close
        Command1.Enabled = False
    End If
End If

    Exit Sub

    End Sub

Private Sub Command4_Click()
Calendar1.Visible = True
End Sub

Private Sub Command5_Click()
Calendar2.Visible = True
End Sub

Private Sub Form_Load()
Label3.Visible = False
Label6.Visible = False
Label7.Visible = False

Call Prod_Code
Call Prod_Name
Call Prod_Model
Call Agent_Id

txtId.Visible = False
cmbProd_Name.Visible = False
cmbProd_Model.Visible = False

cmbType.Visible = False
cmbAgent_Code.Visible = False
Label8.Visible = False
Label4.Visible = False
txtFrom.Text = Today
txtTo.Text = Today
Calendar1.Value = Today
Calendar2.Value = Today
 Calendar1.Visible = False
Calendar2.Visible = False
Command1.Enabled = False
End Sub

Private Sub optAll_Click()
Label3.Visible = False
Label6.Visible = False
Label7.Visible = False

txtId.Visible = False
cmbProd_Name.Visible = False
cmbProd_Model.Visible = False

cmbType.Visible = True
cmbAgent_Code.Visible = True
Label8.Visible = True
Label4.Visible = True
cmbType.SetFocus

End Sub

Private Sub optProduct_Click()
Label3.Visible = True
Label6.Visible = True
Label7.Visible = True

txtId.Visible = True
cmbProd_Name.Visible = True
cmbProd_Model.Visible = True

Label8.Visible = False
cmbType.Visible = False

txtId.SetFocus

End Sub

Private Sub optSummary_Click()
Label3.Visible = False
Label6.Visible = False
Label7.Visible = False

txtId.Visible = False
cmbProd_Name.Visible = False
cmbProd_Model.Visible = False

cmbType.Visible = True
cmbAgent_Code.Visible = True
Label8.Visible = True
Label4.Visible = True
cmbType.SetFocus
End Sub

Private Sub Timer1_Timer()
If optAll.Value = True Then
    If cmbType.Text = "All" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptProd_Statement.Show 1
        End If
    End If
    
    If cmbType.Text = "Purchase" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptPurchase.Show 1
        End If
    End If
    
    If cmbType.Text = "Sale" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptPurchase.Show 1
        End If
    End If
    
    If cmbType.Text = "Agent Sale" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptPurchase.Show 1
        End If
    End If

    If cmbType.Text = "Issue" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptIssue.Show 1
        End If
    End If
    
    If cmbType.Text = "Return" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptIssue.Show 1
        End If
    End If
End If

If optSummary.Value = True Then
    If cmbType.Text = "All" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptProd_Summary.Show 1
        End If
    End If
    
    If cmbType.Text = "Purchase" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptProd_Summary.Show 1
        End If
    End If
    
    If cmbType.Text = "Sale" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptProd_Summary.Show 1
        End If
    End If
    
    If cmbType.Text = "Agent Sale" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptProd_Summary.Show 1
        End If
    End If

    If cmbType.Text = "Issue" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptProd_Summary.Show 1
        End If
    End If
    
    If cmbType.Text = "Return" Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptProd_Summary.Show 1
        End If
    End If
End If


If optProduct.Value = True Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptProd_Statement.Show 1
        End If
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProd_Name.SetFocus
End If
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTo.SelStart = 0
txtTo.SelLength = Len(txtTo.Text)
txtTo.SetFocus
End If
End Sub

Private Sub txtId_LostFocus()
    Dim Prod As String
    Prod = txtId.Text
    
    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtId.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            rsN.Close
        Else
            MsgBox "No such Product found", vbCritical, "Error!"
        rsN.Close
        End If
    Exit Sub
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3.SetFocus
End If
End Sub

