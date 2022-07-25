VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmProd_Search 
   BackColor       =   &H00004080&
   Caption         =   "Product Search"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14835
   Icon            =   "frmProdSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   14835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   915
      Left            =   180
      TabIndex        =   1
      Top             =   135
      Width           =   14505
      Begin VB.ComboBox cmbProduct_Name 
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
         Left            =   5040
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   270
         Width           =   2490
      End
      Begin VB.ComboBox cmbProduct_Sl 
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
         Left            =   1575
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   270
         Width           =   1950
      End
      Begin VB.ComboBox cmbProd_model 
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
         Left            =   9045
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   270
         Width           =   3030
      End
      Begin VB.Label Label2 
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
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label1 
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
         Left            =   7695
         TabIndex        =   3
         Top             =   315
         Width           =   1260
      End
      Begin VB.Label Label4 
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
         Left            =   3690
         TabIndex        =   2
         Top             =   315
         Width           =   1245
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmProdSearch.frx":0442
      Height          =   7230
      Left            =   135
      TabIndex        =   0
      Top             =   1170
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   12753
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
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
      Left            =   225
      Top             =   4635
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
Attribute VB_Name = "frmProd_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String

Private Sub Prod_Name()
        cmbProduct_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name FROM Prod_Master"
        rs.Open str, conn
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbProduct_Name.AddItem rs!Prod_Name
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub Prod_Model()

        cmbProd_Model.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Model FROM Prod_Master"
        rs.Open str, conn
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbProd_Model.AddItem rs!Prod_Model
        rs.MoveNext
        Loop
        rs.Close
End Sub
Private Sub Prod_Sl()
        cmbProduct_Sl.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Code FROM Prod_Master "
        rs.Open str, conn
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbProduct_Sl.AddItem rs!Prod_Code
        rs.MoveNext
        Loop
        rs.Close
End Sub

Private Sub cmbProd_Model_KeyPress(KeyAscii As Integer)
On Error Resume Next
If cmbProd_Model.Text = "" Then
Exit Sub
End If
If KeyAscii = 13 Then
'Call Prod_Sl
cmbProduct_Sl.SetFocus
End If
    Dim search As String
        search = cmbProd_Model.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Model Like '" & "%" & search & "%" & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
     End If
End Sub

Private Sub cmbProd_Model_LostFocus()
Set rs = New ADODB.Recordset
        str = "SELECT * FROM Prod_Tran where Prod_Model like '" & cmbProduct_Name.Text & "'"
        rs.Open str, conn
        
        frmProd_Search.Adodc1.ConnectionString = cnStr
        frmProd_Search.Adodc1.RecordSource = str
        frmProd_Search.Adodc1.Refresh
        rs.Close
End Sub

Private Sub cmbProduct_Name_KeyPress(KeyAscii As Integer)

On Error Resume Next
If cmbProduct_Name.Text = "" Then
Exit Sub
End If
If KeyAscii = 13 Then
cmbProd_Model.SetFocus
cmbProd_Model.SetFocus
End If


    Dim search As String
        search = cmbProduct_Name.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Name Like '" & "%" & search & "%" & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
     End If
End Sub

Private Sub cmbProduct_Name_LostFocus()
Set rs = New ADODB.Recordset
        str = "SELECT * FROM Prod_Tran where Prod_Name like '" & cmbProduct_Name.Text & "'"
        rs.Open str, conn
        
        frmProd_Search.Adodc1.ConnectionString = cnStr
        frmProd_Search.Adodc1.RecordSource = str
        frmProd_Search.Adodc1.Refresh
        rs.Close

End Sub

Private Sub cmbProduct_Sl_KeyPress(KeyAscii As Integer)
On Error Resume Next
If cmbProduct_Sl.Text = "" Then
Exit Sub
End If

'If KeyAscii = 13 Then
'cmbProduct_Name.SetFocus
'End If



    Dim search As String
        search = cmbProduct_Sl.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Code Like '" & "%" & search & "%" & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
     End If

If KeyAscii = 13 Then

Dim ser As String
ser = cmbProduct_Sl.Text

Set rs = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Code Like '" & ser & "'"
        rs.Open str, conn
        If Not rs.EOF Then

        frmSales.txtProd_Sl.Text = rs!Prod_Code
        frmSales.cmbProd_Name.Text = rs!Prod_Name
        frmSales.cmbProd_Model.Text = rs!Prod_Model
        frmSales.txtProd_Cost.Text = Format$(Val(rs!Sale_Price), "###0.00")
        frmSales.txtQty.SetFocus
        
        rs.Close
        Else
        rs.Close
        End If

Unload Me
frmSales.Show
End If
End Sub

Private Sub DataGrid1_DblClick()
cmbProduct_Sl.Text = DataGrid1.Text
End Sub

Private Sub Form_Load()
On Error Resume Next
Call Prod_Name
Call Prod_Model
Call Prod_Sl

End Sub
