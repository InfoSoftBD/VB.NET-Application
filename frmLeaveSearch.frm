VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLeaveSearch 
   BackColor       =   &H00004000&
   Caption         =   "Search"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   Icon            =   "frmLeaveSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Search Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   6855
      Begin VB.OptionButton OptStation 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Station"
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
         Left            =   225
         TabIndex        =   17
         Top             =   2385
         Width           =   1140
      End
      Begin VB.ComboBox txtStation 
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
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   2340
         Width           =   3075
      End
      Begin VB.OptionButton optBom 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ID"
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
         Left            =   225
         TabIndex        =   12
         Top             =   450
         Width           =   1365
      End
      Begin VB.OptionButton optModel 
         BackColor       =   &H00C0FFC0&
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
         Left            =   225
         TabIndex        =   11
         Top             =   945
         Width           =   1410
      End
      Begin VB.OptionButton optParts 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Parts Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   1440
         Width           =   1410
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H00C0FFC0&
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
         Height          =   285
         Left            =   225
         TabIndex        =   9
         Top             =   1913
         Width           =   735
      End
      Begin VB.TextBox txtBom 
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
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   3030
      End
      Begin VB.ComboBox txtModel 
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
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   855
         Width           =   3075
      End
      Begin VB.ComboBox txtParts 
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
         TabIndex        =   6
         Text            =   "Combo2"
         Top             =   1350
         Width           =   3075
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
         Text            =   "Text2"
         Top             =   1845
         Width           =   1275
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
         Left            =   3555
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   1845
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   420
         Left            =   5220
         TabIndex        =   3
         Top             =   1305
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   420
         Left            =   5220
         TabIndex        =   2
         Top             =   2160
         Width           =   1365
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   420
         Left            =   5220
         TabIndex        =   1
         Top             =   450
         Width           =   1365
      End
      Begin VB.Shape Shape2 
         Height          =   2535
         Left            =   135
         Top             =   270
         Width           =   4830
      End
      Begin VB.Shape Shape1 
         Height          =   2535
         Left            =   5085
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   14
         Top             =   1935
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
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
         Left            =   3195
         TabIndex        =   13
         Top             =   1890
         Width           =   210
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLeaveSearch.frx":0442
      Height          =   2040
      Left            =   135
      TabIndex        =   15
      Top             =   3240
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   3598
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   420
      Left            =   405
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
Attribute VB_Name = "frmLeaveSearch"
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
If optBom = True Then
    If txtBom.Text = "" Then
    Exit Sub
    End If
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from IssueSearch where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from IssueSearch where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * from Issue where Bom_code like '" & txtBom.Text & "'"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "IssueSearch", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!sl = 1
        rsN!Bom_code = rs!Bom_code
        rsN!Model_name = rs!Model_name
        rsN!Parts_name = rs!Parts_name
        rsN!Out_date = rs!Out_date
        rsN!Req_no = rs!Req_no
        rsN!Qty = rs!Qty
        rsN!Station = rs!Station
        rsN!Unit_Price = rs!Unit_Price
        rsN!Amount = rs!Amount
        rs.MoveNext
        rsN.Update
        Loop
        
        rs.Close
        rsN.Close
End If

If optModel = True Then
    If txtModel.Text = "" Then
    Exit Sub
    End If
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from IssueSearch where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from IssueSearch where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * from Issue where Model_name like '" & txtModel.Text & "'"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "IssueSearch", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!sl = 1
        rsN!Bom_code = rs!Bom_code
        rsN!Model_name = rs!Model_name
        rsN!Parts_name = rs!Parts_name
        rsN!Out_date = rs!Out_date
        rsN!Req_no = rs!Req_no
        rsN!Qty = rs!Qty
        rsN!Station = rs!Station
        rsN!Unit_Price = rs!Unit_Price
        rsN!Amount = rs!Amount
        rs.MoveNext
        rsN.Update
        Loop
        
        rs.Close
        rsN.Close
 End If
 
If optParts = True Then
    If txtParts.Text = "" Then
    Exit Sub
    End If
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from IssueSearch where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from IssueSearch where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * from Issue where Parts_name like '" & txtParts.Text & "'"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "IssueSearch", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!sl = 1
        rsN!Bom_code = rs!Bom_code
        rsN!Model_name = rs!Model_name
        rsN!Parts_name = rs!Parts_name
        rsN!Out_date = rs!Out_date
        rsN!Req_no = rs!Req_no
        rsN!Qty = rs!Qty
        rsN!Station = rs!Station
        rsN!Unit_Price = rs!Unit_Price
        rsN!Amount = rs!Amount
        rs.MoveNext
        rsN.Update
        Loop
        
        rs.Close
        rsN.Close
 End If
 
If optDate = True Then
    If txtFrom.Text = "" Then
        Exit Sub
    End If
    If txtTo.Text = "" Then
        Exit Sub
    End If
      
Dim fromDate As Date
Dim toDate As Date
    fromDate = txtFrom.Text
    toDate = txtTo.Text
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from IssueSearch where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from IssueSearch where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * from Issue where cdate(Out_Date) Between cdate('" & fromDate & "') And cdate('" & toDate & "')"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "IssueSearch", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!sl = 1
        rsN!Bom_code = rs!Bom_code
        rsN!Model_name = rs!Model_name
        rsN!Parts_name = rs!Parts_name
        rsN!Out_date = rs!Out_date
        rsN!Req_no = rs!Req_no
        rsN!Qty = rs!Qty
        rsN!Station = rs!Station
        rsN!Unit_Price = rs!Unit_Price
        rsN!Amount = rs!Amount
        rs.MoveNext
        rsN.Update
        Loop
        
        rs.Close
        rsN.Close
 End If
 
If OptStation = True Then
    If txtStation.Text = "" Then
    Exit Sub
    End If
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from IssueSearch where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from IssueSearch where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * from Issue where Station like '" & txtStation.Text & "'"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "IssueSearch", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!sl = 1
        rsN!Bom_code = rs!Bom_code
        rsN!Model_name = rs!Model_name
        rsN!Parts_name = rs!Parts_name
        rsN!Out_date = rs!Out_date
        rsN!Req_no = rs!Req_no
        rsN!Qty = rs!Qty
        rsN!Station = rs!Station
        rsN!Unit_Price = rs!Unit_Price
        rsN!Amount = rs!Amount
        rs.MoveNext
        rsN.Update
        Loop
        
        rs.Close
        rsN.Close
 End If
Load.Show 1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If optBom = True Then
    If txtBom.Text = "" Then
    Exit Sub
    End If
    Dim Bom As String
        Bom = txtBom.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Issue where Bom_Code like '" & Bom & "'"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such BOM Code found, ", vbCritical + vbOKOnly
        rs.Close
    End If
End If
    
If optModel = True Then
    If txtModel.Text = "" Then
    Exit Sub
    End If

    Dim Model As String
        Model = txtModel.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Issue where Model_name like '" & Model & "'"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such Model found, ", vbCritical + vbOKOnly
        rs.Close
    End If
End If
    
If optParts = True Then
    If txtParts.Text = "" Then
    Exit Sub
    End If

    Dim Parts As String
        Parts = txtParts.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Issue where Parts_name like '" & Parts & "'"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such Parts Name found, ", vbCritical + vbOKOnly
        rs.Close
    End If
End If
    
If optDate = True Then
    If txtFrom.Text = "" Then
    Exit Sub
    End If
    If txtTo.Text = "" Then
    Exit Sub
    End If

    Dim fromDate As Date
    Dim toDate As Date
        fromDate = txtFrom.Text
        toDate = txtTo.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Issue where cdate(Out_Date) Between cdate('" & fromDate & "') And cdate('" & toDate & "')"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such Parts Name found, ", vbCritical + vbOKOnly
        rs.Close
    End If
End If
    
If OptStation = True Then
    If txtStation.Text = "" Then
    Exit Sub
    End If

    Dim Station As String
        Station = txtStation.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Issue where Station like '" & Station & "'"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such Station found, ", vbCritical + vbOKOnly
        rs.Close
    End If
End If
    Exit Sub
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Parts FROM Stock"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtParts.AddItem rs!Parts
        rs.MoveNext
        Loop
        rs.Close
 
    Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Model FROM Stock"
        rsN.Open str, conn
        rsN.MoveFirst
  
    Do While Not rsN.EOF
        On Error Resume Next
        txtModel.AddItem rsN!Model
        rsN.MoveNext
        Loop
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Station FROM Issue"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtStation.AddItem rs!Station
        rs.MoveNext
        Loop
        rs.Close

txtBom.Visible = False
txtModel.Visible = False
txtParts.Visible = False
txtFrom.Visible = False
txtStation.Visible = False
txtTo.Visible = False
Label1.Visible = False
Label2.Visible = False
End Sub
Private Sub optBom_Click()
txtBom.Visible = True
txtBom.Text = ""
txtBom.SetFocus
txtModel.Visible = False
txtParts.Visible = False
txtFrom.Visible = False
txtStation.Visible = False
txtTo.Visible = False
Label1.Visible = False
Label2.Visible = False
End Sub
Private Sub optDate_Click()
txtBom.Visible = False
txtModel.Visible = False
txtParts.Visible = False
txtFrom.Visible = True
txtFrom.Text = ""
txtFrom.SetFocus
txtStation.Visible = False
txtTo.Visible = True
txtTo.Text = ""
Label1.Visible = True
Label2.Visible = True
End Sub
Private Sub optModel_Click()
txtBom.Visible = False
txtModel.Visible = True
txtModel.Text = ""
txtModel.SetFocus
txtParts.Visible = False
txtStation.Visible = False
txtFrom.Visible = False
txtTo.Visible = False
Label1.Visible = False
Label2.Visible = False
End Sub
Private Sub optParts_Click()
txtBom.Visible = False
txtModel.Visible = False
txtParts.Visible = True
txtParts.Text = ""
txtParts.SetFocus
txtStation.Visible = False
txtFrom.Visible = False
txtTo.Visible = False
Label1.Visible = False
Label2.Visible = False
End Sub
Private Sub OptStation_Click()
txtBom.Visible = False
txtModel.Visible = False
txtParts.Visible = False
txtStation.Visible = True
txtStation.SetFocus
txtStation.Text = ""
txtFrom.Visible = False
txtTo.Visible = False
Label1.Visible = False
Label2.Visible = False
End Sub
Private Sub txtBom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtBom.Text = "" Then
Exit Sub
End If

Dim Bom As String
    Bom = txtBom.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Issue where Bom_Code like '" & Bom & "'"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such BOM Code found, ", vbCritical + vbOKOnly
        rs.Close
    End If
    End If
    Exit Sub
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTo.SetFocus
End If
End Sub

Private Sub txtModel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtModel.Text = "" Then
Exit Sub
End If

Dim Model As String
    Model = txtModel.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Issue where Model_name like '" & Model & "'"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such Model found, ", vbCritical + vbOKOnly
        rs.Close
    End If
    End If
    Exit Sub
End Sub

Private Sub txtParts_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtParts.Text = "" Then
Exit Sub
End If

Dim Parts As String
    Parts = txtParts.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Issue where Parts_name like '" & Parts & "'"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such Parts Name found, ", vbCritical + vbOKOnly
        rs.Close
    End If
    End If
    Exit Sub
End Sub

Private Sub txtStation_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtStation.Text = "" Then
Exit Sub
End If

Dim Station As String
    Station = txtStation.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Issue where Station like '" & Station & "'"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such Station found, ", vbCritical + vbOKOnly
        rs.Close
    End If
    End If
    Exit Sub

End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtFrom.Text = "" Then
Exit Sub
End If
If txtTo.Text = "" Then
Exit Sub
End If

Dim fromDate As Date
Dim toDate As Date
    fromDate = txtFrom.Text
    toDate = txtTo.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Issue where cdate(Out_Date) Between cdate('" & fromDate & "') And cdate('" & toDate & "')"
        rs.Open str, conn
        
    If Not rs.EOF Then
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
    Else
        MsgBox "There is no such Parts Name found, ", vbCritical + vbOKOnly
        rs.Close
    End If
    End If
    Exit Sub

End Sub
