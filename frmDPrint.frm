VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmDPrint 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Report Generation"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   Icon            =   "frmDPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   -135
      Top             =   2475
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
      Height          =   3810
      Left            =   135
      TabIndex        =   4
      Top             =   135
      Width           =   6030
      Begin VB.CommandButton Command4 
         Caption         =   "Export"
         Height          =   420
         Left            =   3000
         TabIndex        =   20
         Top             =   3030
         Width           =   1245
      End
      Begin VB.ComboBox cmbSale 
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
         Left            =   3930
         TabIndex        =   19
         Top             =   990
         Width           =   1710
      End
      Begin VB.OptionButton optSales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Daily Customer Sales / Stock Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   18
         Top             =   960
         Width           =   3525
      End
      Begin VB.OptionButton optAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Daily All Customer  Report"
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
         TabIndex        =   17
         Top             =   480
         Width           =   2805
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3990
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   450
         Width           =   1545
      End
      Begin VB.TextBox txtType 
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
         Left            =   4230
         TabIndex        =   14
         Top             =   2235
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   420
         Left            =   1635
         TabIndex        =   10
         Top             =   3030
         Width           =   1245
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   420
         Left            =   4395
         TabIndex        =   9
         Top             =   3030
         Width           =   1245
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   420
         Left            =   285
         TabIndex        =   8
         Top             =   3030
         Width           =   1245
      End
      Begin VB.OptionButton optDSR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Daily Customer Wise Report"
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
         TabIndex        =   7
         Top             =   1575
         Width           =   2805
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
         Left            =   615
         TabIndex        =   6
         Top             =   2235
         Width           =   1140
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
         Left            =   1845
         TabIndex        =   5
         Top             =   2235
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today"
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
         Left            =   3345
         TabIndex        =   16
         Top             =   465
         Width           =   525
      End
      Begin VB.Shape Shape2 
         Height          =   2520
         Left            =   180
         Top             =   270
         Width           =   5625
      End
      Begin VB.Shape Shape1 
         Height          =   660
         Left            =   180
         Top             =   2910
         Width           =   5625
      End
      Begin VB.Label lblFo_Name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         Left            =   1875
         TabIndex        =   13
         Top             =   1920
         Width           =   1410
      End
      Begin VB.Label lblFO_Code 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
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
         Left            =   615
         TabIndex        =   12
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Type"
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
         Left            =   4230
         TabIndex        =   11
         Top             =   1920
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   135
      ScaleHeight     =   180
      ScaleWidth      =   6000
      TabIndex        =   2
      Top             =   4050
      Width           =   6060
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2265
      Left            =   135
      ScaleHeight     =   2205
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   4410
      Width           =   6060
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmDPrint.frx":0442
         Height          =   1920
         Left            =   135
         TabIndex        =   1
         Top             =   135
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3387
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   420
      Left            =   2340
      Top             =   5790
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
Attribute VB_Name = "frmDPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub ComboSale()
        cmbSale.Clear
        On Error Resume Next
        cmbSale.AddItem "Sales Report"
        cmbSale.AddItem "Issue Report"
        cmbSale.AddItem "Return Report"
        cmbSale.AddItem "Stock Report"
End Sub
Private Sub ComboFO_Code()
        cmbFO_Code.Clear
        cmbFO_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer_Code, Customer_Name FROM Customer_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbFO_Code.AddItem rs!Customer
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

Dim Code As String
Code = cmbFO_Code.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer, Name, Type FROM Customer_Master where Customer like '" & Code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbFO_Name.Text = rs!Name
        txtType.Text = rs!Type
        rs.Close
        End If
End Sub
Private Sub cmbFO_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtType.SetFocus
End If
End Sub
Private Sub cmbFO_Name_LostFocus()
If cmbFO_Name.Text = "" Then
Exit Sub
End If

Dim Code As String
Code = cmbFO_Name.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer, Name, Type FROM Customer_Master where Name like '" & Code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbFO_Code.Text = rs!Customer
        txtType.Text = rs!Type
        rs.Close
        End If
End Sub

Private Sub Command1_Click()
    
If optDSR.Value = True Then
    If cmbFO_Code.Text = "" Then
        Exit Sub
    MsgBox " Please Input Customer Id", vbCritical, "Error"
    cmbFO_Code.SetFocus
    End If
    
    
On Error Resume Next
   
Set rs = New ADODB.Recordset
    str = "select * FROM Customer_Master where Customer like '" & cmbFO_Code.Text & "'"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        rptDSR.rsDsr.ConnectionString = cnStr
        rptDSR.rsDsr.Source = str
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
       
        Timer1.Enabled = True
  
End If

If optAll.Value = True Then
    
On Error Resume Next
   
Set rs = New ADODB.Recordset
    str = "select * FROM Customer_Master order by Customer"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        rptDSR_All.rsDsr.ConnectionString = cnStr
        rptDSR_All.rsDsr.Source = str
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
       
        Timer1.Enabled = True
   
End If


If optSales.Value = True Then

Dim ID As Integer
Dim sort As Integer
Dim cid As String
Dim Prod As String

'---------------------------------------------------------------------------
'Sales Report
'---------------------------------------------------------------------------

If cmbSale.Text = "Sales Report" Then


On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If
'------------------------------------------------------------------------------------------------
ID = 0
sort = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")

  
        Set rs = New ADODB.Recordset
                str = "select DISTINCT Customer, Name, Type FROM Customer_Master where Customer like '" & cid & "'"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

        Do While Not rs.EOF
        On Error Resume Next

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!D_ate = Today
                    rsU!Customer = rs!Customer
                    rsU!Name = rs!Name
                    rsU!Type = rs!Type
                    rsU.Update
                    rsU.Close
                    rs.MoveNext
              Else

              Set rsN = New ADODB.Recordset
                    rsN.Open "Sales_Report", conn, adOpenDynamic, adLockOptimistic, -1
                    rsN.AddNew
                        sort = sort + 1
                        rsN!sl = 1
                        rsN!Sl_no = sort
                        rsN!D_ate = Today
                        rsN!Customer = rs!Customer
                        rsN!Name = rs!Name
                        rsN!Type = rs!Type
                        rsN.Update
                        rsN.Close
                        rs.MoveNext
                End If
                Loop
                   rs.Close
             Else
                   rs.Close
             End If

    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "101"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Load = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "BKS01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Bkash = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC10"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC10 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC20"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC20 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC50"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC50 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC100"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC100 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop


'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC300"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC300 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM01 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM02 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM03"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM03 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM04"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM04 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM05"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM05 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM06"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM06 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM07"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM07 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB01 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB02 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

    Set rs = New ADODB.Recordset
        str = "select * FROM Sales_Report order by Customer"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
    
        rptSales.rsSale.ConnectionString = cnStr
        rptSales.rsSale.Source = str
        rptSales.Label11.Caption = "DAILY CUSTOMER SALES REPORT"
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
       
        Timer1.Enabled = True
End If

'---------------------------------------------------------------------------
'Stock Report
'---------------------------------------------------------------------------

If cmbSale.Text = "Stock Report" Then


On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If
'------------------------------------------------------------------------------------------------
ID = 0
sort = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")

  
        Set rs = New ADODB.Recordset
                str = "select DISTINCT Customer, Name, Type FROM Customer_Master where Customer like '" & cid & "'"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

        Do While Not rs.EOF
        On Error Resume Next

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!D_ate = Today
                    rsU!Customer = rs!Customer
                    rsU!Name = rs!Name
                    rsU!Type = rs!Type
                    rsU.Update
                    rsU.Close
                    rs.MoveNext
              Else

              Set rsN = New ADODB.Recordset
                    rsN.Open "Sales_Report", conn, adOpenDynamic, adLockOptimistic, -1
                    rsN.AddNew
                        sort = sort + 1
                        rsN!sl = 1
                        rsN!Sl_no = sort
                        rsN!D_ate = Today
                        rsN!Customer = rs!Customer
                        rsN!Name = rs!Name
                        rsN!Type = rs!Type
                        rsN.Update
                        rsN.Close
                        rs.MoveNext
                End If
                Loop
                   rs.Close
             Else
                   rs.Close
             End If

    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "101"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Load = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "BKS01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Bkash = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC10"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC10 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC20"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC20 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC50"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC50 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC100"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC100 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop


'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC300"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC300 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM01 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM02 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM03"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM03 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM04"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM04 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM05"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM05 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM06"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM06 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM07"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM07 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB01 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB02 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

    Set rs = New ADODB.Recordset
        str = "select * FROM Sales_Report order by Customer"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
    
        rptSales.rsSale.ConnectionString = cnStr
        rptSales.rsSale.Source = str
        rptSales.Label11.Caption = "DAILY CUSTOMER STOCK REPORT"
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
       
        Timer1.Enabled = True
End If


'---------------------------------------------------------------------------
'Issue Report
'---------------------------------------------------------------------------

If cmbSale.Text = "Issue Report" Then


On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If
'------------------------------------------------------------------------------------------------
ID = 0
sort = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")

  
        Set rs = New ADODB.Recordset
                str = "select DISTINCT Customer, Name, Type FROM Customer_Master where Customer like '" & cid & "'"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

        Do While Not rs.EOF
        On Error Resume Next

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!D_ate = Today
                    rsU!Customer = rs!Customer
                    rsU!Name = rs!Name
                    rsU!Type = rs!Type
                    rsU.Update
                    rsU.Close
                    rs.MoveNext
              Else

              Set rsN = New ADODB.Recordset
                    rsN.Open "Sales_Report", conn, adOpenDynamic, adLockOptimistic, -1
                    rsN.AddNew
                        sort = sort + 1
                        rsN!sl = 1
                        rsN!Sl_no = sort
                        rsN!D_ate = Today
                        rsN!Customer = rs!Customer
                        rsN!Name = rs!Name
                        rsN!Type = rs!Type
                        rsN.Update
                        rsN.Close
                        rs.MoveNext
                End If
                Loop
                   rs.Close
             Else
                   rs.Close
             End If

    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "101"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Load = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "BKS01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Bkash = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC10"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC10 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC20"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC20 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC50"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC50 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC100"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC100 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop


'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC300"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC300 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM01 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM02 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM03"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM03 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM04"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM04 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM05"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM05 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM06"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM06 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM07"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM07 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB01 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB02 = rs!Receive
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

    Set rs = New ADODB.Recordset
        str = "select * FROM Sales_Report order by Customer"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
    
        rptSales.rsSale.ConnectionString = cnStr
        rptSales.rsSale.Source = str
        rptSales.Label11.Caption = "DAILY CUSTOMER ISSUE REPORT"
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
       
        Timer1.Enabled = True
End If


'---------------------------------------------------------------------------
'Return Report
'---------------------------------------------------------------------------

If cmbSale.Text = "Return Report" Then


On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If
'------------------------------------------------------------------------------------------------
ID = 0
sort = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")

  
        Set rs = New ADODB.Recordset
                str = "select DISTINCT Customer, Name, Type FROM Customer_Master where Customer like '" & cid & "'"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

        Do While Not rs.EOF
        On Error Resume Next

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!D_ate = Today
                    rsU!Customer = rs!Customer
                    rsU!Name = rs!Name
                    rsU!Type = rs!Type
                    rsU.Update
                    rsU.Close
                    rs.MoveNext
              Else

              Set rsN = New ADODB.Recordset
                    rsN.Open "Sales_Report", conn, adOpenDynamic, adLockOptimistic, -1
                    rsN.AddNew
                        sort = sort + 1
                        rsN!sl = 1
                        rsN!Sl_no = sort
                        rsN!D_ate = Today
                        rsN!Customer = rs!Customer
                        rsN!Name = rs!Name
                        rsN!Type = rs!Type
                        rsN.Update
                        rsN.Close
                        rs.MoveNext
                End If
                Loop
                   rs.Close
             Else
                   rs.Close
             End If

    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "101"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Load = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "BKS01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Bkash = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC10"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC10 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC20"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC20 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC50"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC50 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC100"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC100 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop


'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC300"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC300 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM01 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM02 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM03"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM03 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM04"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM04 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM05"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM05 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM06"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM06 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM07"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM07 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB01 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB02 = rs!Return
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

    Set rs = New ADODB.Recordset
        str = "select * FROM Sales_Report order by Customer"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
    
        rptSales.rsSale.ConnectionString = cnStr
        rptSales.rsSale.Source = str
        rptSales.Label11.Caption = "DAILY CUSTOMER RETURN REPORT"
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
       
        Timer1.Enabled = True
End If


'-----------------------------------------------------------------------------------------
'

End If

Command1.Enabled = False

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
   
If optDSR.Value = True Then
     
    If cmbFO_Code.Text = "" Then
        Exit Sub
    MsgBox " Please Input Customer ID", vbCritical, "Error"
    cmbFO_Code.SetFocus
    End If

Set rs = New ADODB.Recordset
        str = "select * FROM Customer_Master where Customer_Code like '" & cmbFO_Code.Text & "'"
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

If optAll.Value = True Then
    
Set rs = New ADODB.Recordset
        str = "select * FROM Customer_Master order by Customer_Code"
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


If optSales.Value = True Then
   '----------------------------------------------------------------------
   'Sales Report
   '----------------------------------------------------------------------
         Dim ID As Integer
        Dim sort As Integer
        Dim cid As String
        Dim Prod As String
   
   If cmbSale.Text = "Sales Report" Then
        
        
        
        Set rs = New ADODB.Recordset
                str = "select * FROM Sales_Report"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
            
                Adodc1.ConnectionString = cnStr
                Adodc1.RecordSource = str
                Adodc1.Refresh
                DataGrid1.Refresh
                rs.Close
                
                Command1.Enabled = True
                Command4.Enabled = True
            End If
    End If
'--------------------------------------------------------------------------------------------
'Stock
'--------------------------------------------------------------------------------------------

 If cmbSale.Text = "Stock Report" Then
        
End If
'-------------------------------------------------------------------------------------
'Issue
'-------------------------------------------------------------------------------------
If cmbSale.Text = "Issue Report" Then
        
        'Dim ID As Integer
       ' Dim sort As Integer
       ' Dim cid As String
       ' Dim Prod As String
        
        On Error Resume Next
            Set rsU = New ADODB.Recordset
                str = "select * from Sales_Report where sl like '" & 1 & "'"
                rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
            
            If Not rsU.EOF Then
            On Error Resume Next
                rsU.Close
                str = "delete from Sales_Report where sl like '" & 1 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                rsU.Update
                rsU.Close
           End If
        '------------------------------------------------------------------------------------------------
        ID = 0
        sort = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "101"
        sort = sort + 1
                Set rs = New ADODB.Recordset
                        str = "select DISTINCT Customer, Name, Type FROM Customer_Master where Customer like '" & cid & "'"
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                Do While Not rs.EOF
                On Error Resume Next
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!D_ate = Today
                            rsU!Customer = rs!Customer
                            rsU!Name = rs!Name
                            rsU!Type = rs!Type
                            'rsU!Load = rs!Sale
                            rsU.Update
                            rsU.Close
                            rs.MoveNext
                      Else
        
                      Set rsN = New ADODB.Recordset
                            rsN.Open "Sales_Report", conn, adOpenDynamic, adLockOptimistic, -1
                            rsN.AddNew
                                rsN!sl = 1
                                rsN!Sl_no = sort
                                rsN!D_ate = Today
                                rsN!Customer = rs!Customer
                                rsN!Name = rs!Name
                                rsN!Type = rs!Type
                                'rsN!Load = rs!Sale
                                rsN.Update
                                rsN.Close
                                rs.MoveNext
                        End If
                        Loop
                           rs.Close
                     Else
                           rs.Close
                     End If
        
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "101"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!Load = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "BKS01"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!Bkash = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SC10"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SC10 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SC20"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SC20 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SC50"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SC50 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SC100"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SC100 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SC300"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SC300 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM01"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM01 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM02"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM02 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM03"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM03 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM04"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM04 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM05"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM05 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM06"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM06 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM07"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM07 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "MOB01"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!MOB01 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "MOB02"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Receive >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!MOB02 = rs!Receive
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        
        
        Set rs = New ADODB.Recordset
                str = "select * FROM Sales_Report"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
            
                Adodc1.ConnectionString = cnStr
                Adodc1.RecordSource = str
                Adodc1.Refresh
                DataGrid1.Refresh
                rs.Close
                
                Command1.Enabled = True
            End If
    End If

'---------------------------------------------------------------------------------------------
'Return
'---------------------------------------------------------------------------------------------

If cmbSale.Text = "Return Report" Then
        
        'Dim ID As Integer
       ' Dim sort As Integer
       ' Dim cid As String
       ' Dim Prod As String
        
        On Error Resume Next
            Set rsU = New ADODB.Recordset
                str = "select * from Sales_Report where sl like '" & 1 & "'"
                rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
            
            If Not rsU.EOF Then
            On Error Resume Next
                rsU.Close
                str = "delete from Sales_Report where sl like '" & 1 & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                rsU.Update
                rsU.Close
           End If
        '------------------------------------------------------------------------------------------------
        ID = 0
        sort = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "101"
        sort = sort + 1
                Set rs = New ADODB.Recordset
                        str = "select DISTINCT Customer, Name, Type FROM Customer_Master where Customer like '" & cid & "'"
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                Do While Not rs.EOF
                On Error Resume Next
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!D_ate = Today
                            rsU!Customer = rs!Customer
                            rsU!Name = rs!Name
                            rsU!Type = rs!Type
                            'rsU!Load = rs!Sale
                            rsU.Update
                            rsU.Close
                            rs.MoveNext
                      Else
        
                      Set rsN = New ADODB.Recordset
                            rsN.Open "Sales_Report", conn, adOpenDynamic, adLockOptimistic, -1
                            rsN.AddNew
                                rsN!sl = 1
                                rsN!Sl_no = sort
                                rsN!D_ate = Today
                                rsN!Customer = rs!Customer
                                rsN!Name = rs!Name
                                rsN!Type = rs!Type
                                'rsN!Load = rs!Sale
                                rsN.Update
                                rsN.Close
                                rs.MoveNext
                        End If
                        Loop
                           rs.Close
                     Else
                           rs.Close
                     End If
        
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "101"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!Load = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "BKS01"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!Bkash = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SC10"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SC10 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SC20"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SC20 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SC50"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SC50 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SC100"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SC100 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SC300"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SC300 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM01"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM01 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM02"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM02 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM03"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM03 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM04"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM04 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM05"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM05 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM06"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM06 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "SM07"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!SM07 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "MOB01"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!MOB01 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        '----------------------------------------------------------------------------------------------------
        ID = 0
        Do While Not ID = "0040"
        ID = ID + 1
        cid = Format$(Val(ID), "00#0")
        Prod = "MOB02"
          
                Set rs = New ADODB.Recordset
                        str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Return >0 "
                        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        rs.MoveFirst
        
                    Set rsU = New ADODB.Recordset
                        str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                    
                      If Not rsU.EOF Then
                            rsU!MOB02 = rs!Return
                            rsU.Update
                            rsU.Close
                        Else
                            rsU.Close
                        End If
                           rs.Close
                     Else
                           rs.Close
                     End If
            Loop
        
        
        
        Set rs = New ADODB.Recordset
                str = "select * FROM Sales_Report"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
            
                Adodc1.ConnectionString = cnStr
                Adodc1.RecordSource = str
                Adodc1.Refresh
                DataGrid1.Refresh
                rs.Close
                
                Command1.Enabled = True
            End If
    End If
'--------------------------------------------------------------------
'--------------------------------------------------------------------

End If


    Exit Sub
    End Sub

Private Sub Command4_Click()
If optSales.Value = True Then

Dim ID As Integer
Dim sort As Integer
Dim cid As String
Dim Prod As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Set oSheet = oBook.Worksheets(1)

'---------------------------------------------------------------------------
'Sales Report
'---------------------------------------------------------------------------

If cmbSale.Text = "Sales Report" Then

On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If
'------------------------------------------------------------------------------------------------
ID = 0
sort = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")

  
        Set rs = New ADODB.Recordset
                str = "select DISTINCT Customer, Name, Type FROM Customer_Master where Customer like '" & cid & "'"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

        Do While Not rs.EOF
        On Error Resume Next

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!D_ate = Today
                    rsU!Customer = rs!Customer
                    rsU!Name = rs!Name
                    rsU!Type = rs!Type
                    rsU.Update
                    rsU.Close
                    rs.MoveNext
              Else

              Set rsN = New ADODB.Recordset
                    rsN.Open "Sales_Report", conn, adOpenDynamic, adLockOptimistic, -1
                    rsN.AddNew
                        sort = sort + 1
                        rsN!sl = 1
                        rsN!Sl_no = sort
                        rsN!D_ate = Today
                        rsN!Customer = rs!Customer
                        rsN!Name = rs!Name
                        rsN!Type = rs!Type
                        rsN.Update
                        rsN.Close
                        rs.MoveNext
                End If
                Loop
                   rs.Close
             Else
                   rs.Close
             End If

    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "101"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Load = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "BKS01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Bkash = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC10"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC10 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC20"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC20 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC50"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC50 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC100"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC100 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop


'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC300"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC300 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM01 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM02 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM03"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM03 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM04"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM04 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM05"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM05 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM06"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM06 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM07"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM07 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB01 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Sale >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB02 = rs!Sale
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

    Set rs = New ADODB.Recordset
        str = "select Sl_No, Customer, Name, Type, Bkash, Load, SC10, SC20, SC50, SC100, SC300, SM03, SM01, SM02, SM06, MOB01 FROM Sales_Report order by Customer"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
    
   'Transfer the data to Excel
   oSheet.Range("A1").Value = "Sl"
   oSheet.Range("B1").Value = "Customer Id"
   oSheet.Range("C1").Value = "Customer Name"
   oSheet.Range("D1").Value = "Type"
   oSheet.Range("E1").Value = "BKash"
   oSheet.Range("F1").Value = "Load"
   oSheet.Range("G1").Value = "SC10"
   oSheet.Range("H1").Value = "SC20"
   oSheet.Range("I1").Value = "SC50"
   oSheet.Range("J1").Value = "SC100"
   oSheet.Range("K1").Value = "SC300"
   oSheet.Range("L1").Value = "Uday"
   oSheet.Range("M1").Value = "Ecush"
   oSheet.Range("N1").Value = "Uddakta"
   oSheet.Range("O1").Value = "HLR"
   oSheet.Range("P1").Value = "B6Set"
   
   
   oSheet.Range("A2").CopyFromRecordset rs
   oBook.SaveAs "D:\Sales_Report" & Today & ".xls"
   oExcel.Quit
   rs.Close
   MsgBox "Export completed to D:\Sales_Report" & Today & ".xls", vbInformation, "Export!"
    
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
       
   Command4.Enabled = False
  Command1.Enabled = False
End If

'---------------------------------------------------------------------------
'Stock Report
'---------------------------------------------------------------------------

If cmbSale.Text = "Stock Report" Then


On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Sales_Report where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If
'------------------------------------------------------------------------------------------------
ID = 0
sort = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")

  
        Set rs = New ADODB.Recordset
                str = "select DISTINCT Customer, Name, Type FROM Customer_Master where Customer like '" & cid & "'"
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

        Do While Not rs.EOF
        On Error Resume Next

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!D_ate = Today
                    rsU!Customer = rs!Customer
                    rsU!Name = rs!Name
                    rsU!Type = rs!Type
                    rsU.Update
                    rsU.Close
                    rs.MoveNext
              Else

              Set rsN = New ADODB.Recordset
                    rsN.Open "Sales_Report", conn, adOpenDynamic, adLockOptimistic, -1
                    rsN.AddNew
                        sort = sort + 1
                        rsN!sl = 1
                        rsN!Sl_no = sort
                        rsN!D_ate = Today
                        rsN!Customer = rs!Customer
                        rsN!Name = rs!Name
                        rsN!Type = rs!Type
                        rsN.Update
                        rsN.Close
                        rs.MoveNext
                End If
                Loop
                   rs.Close
             Else
                   rs.Close
             End If

    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "101"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Load = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "BKS01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!Bkash = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC10"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC10 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC20"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC20 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC50"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC50 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC100"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC100 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop


'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SC300"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SC300 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM01 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM02 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM03"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM03 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM04"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM04 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop
'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM05"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM05 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM06"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM06 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "SM07"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!SM07 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB01"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB01 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

'----------------------------------------------------------------------------------------------------
ID = 0
Do While Not ID = "0040"
ID = ID + 1
cid = Format$(Val(ID), "00#0")
Prod = "MOB02"
  
        Set rs = New ADODB.Recordset
                str = "select * FROM Customer_Master where Customer like '" & cid & "' and Prod_Code like '" & Prod & "' and Close_Bal >0 "
                rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                rs.MoveFirst

            Set rsU = New ADODB.Recordset
                str = "select * FROM Sales_Report where Customer like '" & cid & "'"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
              If Not rsU.EOF Then
                    rsU!MOB02 = rs!Close_Bal
                    rsU.Update
                    rsU.Close
                Else
                    rsU.Close
                End If
                   rs.Close
             Else
                   rs.Close
             End If
    Loop

    Set rs = New ADODB.Recordset
        str = "select Sl_No, Customer, Name, Type, Bkash, Load, SC10, SC20, SC50, SC100, SC300, SM03, SM01, SM02, SM06, MOB01 FROM Sales_Report order by Customer"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
    
   'Transfer the data to Excel
   oSheet.Range("A1").Value = "Sl"
   oSheet.Range("B1").Value = "Customer Id"
   oSheet.Range("C1").Value = "Customer Name"
   oSheet.Range("D1").Value = "Type"
   oSheet.Range("E1").Value = "BKash"
   oSheet.Range("F1").Value = "Load"
   oSheet.Range("G1").Value = "SC10"
   oSheet.Range("H1").Value = "SC20"
   oSheet.Range("I1").Value = "SC50"
   oSheet.Range("J1").Value = "SC100"
   oSheet.Range("K1").Value = "SC300"
   oSheet.Range("L1").Value = "Uday"
   oSheet.Range("M1").Value = "Ecush"
   oSheet.Range("N1").Value = "Uddakta"
   oSheet.Range("O1").Value = "HLR"
   oSheet.Range("P1").Value = "B6Set"
   
   oSheet.Range("A2").CopyFromRecordset rs
   oBook.SaveAs "D:\Stock_Report" & Today & ".xls"
   oExcel.Quit
  
   rs.Close
   MsgBox "Export completed to D:\Stock_Report" & Today & ".xls", vbInformation, "Export!"
          
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
       
  Command4.Enabled = False
  Command1.Enabled = False
End If
End If
   
End Sub

Private Sub Form_Load()
    lblType.Visible = False
    txtType.Visible = False
    cmbFO_Code.Visible = False
    cmbFO_Name.Visible = False
    lblFO_Code.Visible = False
    lblFo_Name.Visible = False
    txtDate.Text = Today
    Call ComboFO_Code
    Call ComboSale
    Command1.Enabled = False
    Command4.Enabled = False
End Sub

Private Sub optDate_Click()
lblType.Visible = True
txtType.Visible = True
cmbFO_Code.Visible = True
cmbFO_Name.Visible = True
lblFO_Code.Visible = True
lblFo_Name.Visible = True
cmbFO_Code.SetFocus
End Sub

Private Sub optDSR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbFO_Code.SetFocus
End If
End Sub
Private Sub optDSR_Click()
lblType.Visible = True
txtType.Visible = True
cmbFO_Code.Visible = True
cmbFO_Name.Visible = True
lblFO_Code.Visible = True
lblFo_Name.Visible = True
cmbFO_Code.SetFocus
End Sub
Private Sub Timer1_Timer()
If optDSR.Value = True Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptDSR.Show 1
    End If
End If
If optAll.Value = True Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptDSR_All.Show 1
    End If
End If
If optSales.Value = True Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = 100
        ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = 0
        Timer1.Enabled = False
        ProgressBar1.Visible = False
        rptSales.Show 1
    End If
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
Private Sub txtType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3.SetFocus
End If
End Sub
