VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Merchandiser"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   -3165
   ClientWidth     =   14550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   14550
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6810
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   10605
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   5490
      Top             =   270
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4815
      Top             =   270
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   6120
      Top             =   270
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4275
      Top             =   270
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   12645
      TabIndex        =   0
      Top             =   10185
      Width           =   12645
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "LICENCE TO: M/S ASIYA Trading"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   360
         Left            =   495
         TabIndex        =   1
         Top             =   30
         Width           =   4770
      End
   End
   Begin VB.Image Image10 
      Height          =   1440
      Left            =   5100
      Stretch         =   -1  'True
      Top             =   8490
      Width           =   4965
   End
   Begin VB.Image Image1 
      Height          =   10575
      Left            =   0
      Picture         =   "frmMain.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14670
   End
   Begin VB.Menu GL 
      Caption         =   "&General Ledger"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu GLAC 
         Caption         =   "GL Account Opening"
      End
      Begin VB.Menu pp 
         Caption         =   "-"
      End
      Begin VB.Menu ctp 
         Caption         =   "Cash Transaction Posting"
      End
      Begin VB.Menu gf54 
         Caption         =   "-"
      End
      Begin VB.Menu GLTP 
         Caption         =   "Transfer Transaction Posting"
      End
      Begin VB.Menu po 
         Caption         =   "-"
      End
      Begin VB.Menu BTP 
         Caption         =   "Bank Transaction Posting"
      End
      Begin VB.Menu e 
         Caption         =   "-"
      End
      Begin VB.Menu GLST 
         Caption         =   "GL Account Statement"
      End
      Begin VB.Menu pim 
         Caption         =   "-"
      End
      Begin VB.Menu GLDT 
         Caption         =   "Daily Journal Statement"
      End
      Begin VB.Menu poiu 
         Caption         =   "-"
      End
      Begin VB.Menu Jsum 
         Caption         =   "Daily Journal Summary"
      End
      Begin VB.Menu loiut 
         Caption         =   "-"
      End
      Begin VB.Menu CTR 
         Caption         =   "Daily Cash Flow Statement"
      End
      Begin VB.Menu ghgfdsdg 
         Caption         =   "-"
      End
      Begin VB.Menu dwcfstm 
         Caption         =   "Date wise Cash Flow Statement"
      End
      Begin VB.Menu dwcfstm1 
         Caption         =   "-"
      End
      Begin VB.Menu ndr 
         Caption         =   "Note Denomination Statement"
      End
      Begin VB.Menu qwe 
         Caption         =   "-"
      End
      Begin VB.Menu bac 
         Caption         =   "Bank Account Statement"
      End
      Begin VB.Menu pu 
         Caption         =   "-"
      End
      Begin VB.Menu bp 
         Caption         =   "Daily Bank Balance Statement"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu tb 
         Caption         =   "Daily Statement of Affairs"
      End
      Begin VB.Menu py 
         Caption         =   "-"
      End
      Begin VB.Menu SOA 
         Caption         =   "Daily Trial Balance"
      End
      Begin VB.Menu prt 
         Caption         =   "-"
      End
      Begin VB.Menu dplac 
         Caption         =   "Daily Income/Expense"
      End
      Begin VB.Menu loiufr 
         Caption         =   "-"
      End
      Begin VB.Menu dateinc 
         Caption         =   "Date Wise Income/Expense"
      End
      Begin VB.Menu dateinc1 
         Caption         =   "-"
      End
      Begin VB.Menu pl 
         Caption         =   "Profit / Loss Account"
      End
      Begin VB.Menu pre 
         Caption         =   "-"
      End
      Begin VB.Menu CHA 
         Caption         =   "Chart of Accounts"
      End
   End
   Begin VB.Menu Savings 
      Caption         =   "&Customer Management"
      Begin VB.Menu SBAC 
         Caption         =   "Customer Account Opening"
      End
      Begin VB.Menu lk 
         Caption         =   "-"
      End
      Begin VB.Menu cdu 
         Caption         =   "Customer Due Collection"
      End
      Begin VB.Menu cde1 
         Caption         =   "-"
      End
      Begin VB.Menu SBST 
         Caption         =   "Customer Account Statement"
         Index           =   0
      End
      Begin VB.Menu dsfdsfa 
         Caption         =   "-"
      End
      Begin VB.Menu agent 
         Caption         =   "Agent wise Customer Summary"
      End
      Begin VB.Menu agent1 
         Caption         =   "-"
      End
      Begin VB.Menu dcwsr 
         Caption         =   "Daily Customer wise Sales Report"
      End
      Begin VB.Menu dcwsr1 
         Caption         =   "-"
      End
      Begin VB.Menu SBDT 
         Caption         =   "Daily All Customer Report"
      End
      Begin VB.Menu loac123 
         Caption         =   "-"
      End
      Begin VB.Menu loac 
         Caption         =   "List of All Customer"
      End
      Begin VB.Menu cdu1 
         Caption         =   "-"
      End
      Begin VB.Menu cdu2 
         Caption         =   "Customer Data Update"
      End
   End
   Begin VB.Menu PM 
      Caption         =   "&Product Management"
      Begin VB.Menu PE 
         Caption         =   "Product Receive"
      End
      Begin VB.Menu mn 
         Caption         =   "-"
      End
      Begin VB.Menu pr 
         Caption         =   "Product Sales"
      End
      Begin VB.Menu mb 
         Caption         =   "-"
      End
      Begin VB.Menu prtn 
         Caption         =   "Purchase Return"
      End
      Begin VB.Menu prtn1 
         Caption         =   "-"
      End
      Begin VB.Menu pigdn 
         Caption         =   "Product Issue To Godown"
      End
      Begin VB.Menu pigdn1 
         Caption         =   "-"
      End
      Begin VB.Menu prtngd 
         Caption         =   "Product Return From Godown"
      End
      Begin VB.Menu prtngd1 
         Caption         =   "-"
      End
      Begin VB.Menu dsrpt12 
         Caption         =   "Daily Sales Invoice Report"
      End
      Begin VB.Menu dsrpt1 
         Caption         =   "-"
      End
      Begin VB.Menu datsalinv 
         Caption         =   "Date Wise Sales Invoice Report"
      End
      Begin VB.Menu datsalinv1 
         Caption         =   "-"
      End
      Begin VB.Menu ps 
         Caption         =   "Daily Product Report"
      End
      Begin VB.Menu mv 
         Caption         =   "-"
      End
      Begin VB.Menu SP 
         Caption         =   "Summary Stock of Inventory"
      End
      Begin VB.Menu dfdfdh 
         Caption         =   "-"
      End
      Begin VB.Menu dsr123654 
         Caption         =   "Details Stock of Inventory"
      End
      Begin VB.Menu hfg876 
         Caption         =   "-"
      End
      Begin VB.Menu gdnSum 
         Caption         =   "Summary Stock of Godown"
      End
      Begin VB.Menu gdnSum1 
         Caption         =   "-"
      End
      Begin VB.Menu gdnstm 
         Caption         =   "Daily Godown Statement"
      End
      Begin VB.Menu gdnstm1 
         Caption         =   "-"
      End
      Begin VB.Menu stkrptexp 
         Caption         =   "Stock Report Export"
      End
      Begin VB.Menu stkexp1 
         Caption         =   "-"
      End
      Begin VB.Menu sdu 
         Caption         =   "Stock Data Update"
      End
      Begin VB.Menu sdu1 
         Caption         =   "-"
      End
      Begin VB.Menu pser 
         Caption         =   "Product Search"
      End
      Begin VB.Menu pser1 
         Caption         =   "-"
      End
      Begin VB.Menu bcp12 
         Caption         =   "Bar Code Print"
      End
      Begin VB.Menu lcent1 
         Caption         =   "-"
      End
      Begin VB.Menu lc 
         Caption         =   "Letter of Credit Entry"
      End
      Begin VB.Menu lcrg1 
         Caption         =   "-"
      End
      Begin VB.Menu lcr 
         Caption         =   "LC Registrar"
      End
   End
   Begin VB.Menu c 
      Caption         =   "&Vendor Management"
      Begin VB.Menu vao 
         Caption         =   "Vendor Account Opening"
      End
      Begin VB.Menu vao1 
         Caption         =   "-"
      End
      Begin VB.Menu VDP 
         Caption         =   "Vendor Due Payment"
      End
      Begin VB.Menu mc 
         Caption         =   "-"
      End
      Begin VB.Menu vs 
         Caption         =   "Vendor Account Statement"
      End
      Begin VB.Menu mx 
         Caption         =   "-"
      End
      Begin VB.Menu DVP 
         Caption         =   "Daily Vendor Position"
      End
   End
   Begin VB.Menu STb 
      Caption         =   "&Office Management"
      Begin VB.Menu Employee 
         Caption         =   "Staff Information Entry"
      End
      Begin VB.Menu mnb 
         Caption         =   "-"
      End
      Begin VB.Menu ssp 
         Caption         =   "Staff Payment"
      End
      Begin VB.Menu ssp1 
         Caption         =   "-"
      End
      Begin VB.Menu ssd 
         Caption         =   "Staff Due Collection"
      End
      Begin VB.Menu fghdf 
         Caption         =   "-"
      End
      Begin VB.Menu Stold 
         Caption         =   "Staff Old Due Posting"
      End
      Begin VB.Menu stold123 
         Caption         =   "-"
      End
      Begin VB.Menu sas 
         Caption         =   "Staff Account Statement"
      End
      Begin VB.Menu ssd1 
         Caption         =   "-"
      End
      Begin VB.Menu SR 
         Caption         =   "Staff Report"
      End
      Begin VB.Menu lkjh 
         Caption         =   "-"
      End
      Begin VB.Menu rteter 
         Caption         =   "Staff Transfer Posting"
      End
      Begin VB.Menu cm 
         Caption         =   "-"
      End
      Begin VB.Menu FAS 
         Caption         =   "Fixed Asset Schedule"
      End
      Begin VB.Menu nb 
         Caption         =   "-"
      End
      Begin VB.Menu FA 
         Caption         =   "Fixed Asset Statement"
      End
      Begin VB.Menu nv 
         Caption         =   "-"
      End
      Begin VB.Menu DFP 
         Caption         =   "Daily Fixed Position"
      End
      Begin VB.Menu bcp987 
         Caption         =   "-"
      End
      Begin VB.Menu FASUM 
         Caption         =   "Fixed Asset Summary"
      End
      Begin VB.Menu nc 
         Caption         =   "-"
      End
      Begin VB.Menu BM 
         Caption         =   "Branch Mainatance"
      End
      Begin VB.Menu xm 
         Caption         =   "-"
      End
      Begin VB.Menu bl 
         Caption         =   "Branch List"
      End
      Begin VB.Menu bn 
         Caption         =   "-"
      End
      Begin VB.Menu sm 
         Caption         =   "Samaty Maintanance"
      End
      Begin VB.Menu vn 
         Caption         =   "-"
      End
      Begin VB.Menu smr 
         Caption         =   "Samity Report"
      End
   End
   Begin VB.Menu tl 
      Caption         =   "&Tools"
      Begin VB.Menu DClose 
         Caption         =   "Day Close Process"
      End
      Begin VB.Menu qp 
         Caption         =   "-"
      End
      Begin VB.Menu WClose 
         Caption         =   "Week Close Process"
      End
      Begin VB.Menu wo 
         Caption         =   "-"
      End
      Begin VB.Menu MClose 
         Caption         =   "Month Close Process"
      End
      Begin VB.Menu fd 
         Caption         =   "-"
      End
      Begin VB.Menu Vm 
         Caption         =   "View Parameter Table"
      End
      Begin VB.Menu ei 
         Caption         =   "-"
      End
      Begin VB.Menu UM 
         Caption         =   "User Maintanance"
      End
      Begin VB.Menu al 
         Caption         =   "-"
      End
      Begin VB.Menu cal 
         Caption         =   "Calculator"
      End
      Begin VB.Menu sk 
         Caption         =   "-"
      End
      Begin VB.Menu bk 
         Caption         =   "Backup"
      End
      Begin VB.Menu fg 
         Caption         =   "-"
      End
      Begin VB.Menu rs 
         Caption         =   "Restore"
      End
      Begin VB.Menu HGF 
         Caption         =   "-"
      End
      Begin VB.Menu su25 
         Caption         =   "Sales Update"
      End
      Begin VB.Menu hjuy00 
         Caption         =   "-"
      End
      Begin VB.Menu cupdate 
         Caption         =   "Cash Update"
      End
   End
   Begin VB.Menu abt 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Super()
GL.Enabled = True
PM.Enabled = True
STb.Enabled = True
tl.Enabled = True
End Sub
Private Sub Ope()
GL.Enabled = False
PM.Enabled = False
STb.Enabled = False
tl.Enabled = False
End Sub
Private Sub abt_Click()
frmAbout.Show 1
End Sub

Private Sub agent_Click()
Dim rs As ADODB.Recordset
Dim str As String

Set rs = New ADODB.Recordset
        str = "select * FROM Customer_Master order by Agent_Code"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        rs.Close
        rpt_Customer_Summary.rsTran.ConnectionString = cnStr
        rpt_Customer_Summary.rsTran.Source = str
End If
rpt_Customer_Summary.Show 1

End Sub

Private Sub Bac_Click()
frmBank_Stm.Show 1
End Sub

Private Sub baccount_Click()
frmBank_Stm.Show 1
End Sub

Private Sub bcp12_Click()
frmBar_Print.Show 1
'frmBar_Code.Show 1
End Sub

Private Sub bk_Click()
frmBackup.Show 1
End Sub

Private Sub bl_Click()
Dim str As String
str = "select * from Branch order by Branch_Code"
rptBranch.rsBranch.ConnectionString = cnStr
rptBranch.rsBranch.Source = str
rptBranch.Show 1
End Sub

Private Sub BM_Click()
frmBranch.Show 1
End Sub

Private Sub bp_Click()
Dim sort As Integer
        sort = 0

Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Bank_Master"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Bank_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Bank_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * FROM Bank_Master order by Id"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Bank_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!AC_No = rs!AC_No
        rsN!Bank_Name = rs!Bank_Name
        rsN!Branch_Name = rs!Branch_Name
        rsN!Open_Bal = rs!Open_Bal
        rsN!Deposit = rs!Deposit
        rsN!Withdraw = rs!Withdraw
        rsN!Balance = rs!Balance
        
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
        
            Set rsU = New ADODB.Recordset
            str = "select * from Bank_Print"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
        
        If Not rsU.EOF Then
            rsU!Bank_Open = rs!Bank_Open
            rsU!Bank_Dr = rs!Bank_Dr
            rsU!Bank_Cr = rs!Bank_Cr
            rsU!Bank_Close = rs!Bank_Close
            rsU.Update
            rsU.Close
            rs.Close
        End If
    End If
    
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Bank_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptBank_Position.rsTran.ConnectionString = cnStr
        rptBank_Position.rsTran.Source = str
    
    Else
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Bank_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptBank_Position.rsTran.ConnectionString = cnStr
        rptBank_Position.rsTran.Source = str
End If
rptBank_Position.Show 1
End Sub

Private Sub brstatem_Click()
frmBranch_stm.Show 1
End Sub

Private Sub BTP_Click()
frmBank_Tran.Show 1
End Sub

Private Sub cal_Click()
Shell App.Path & "\calc.exe"
End Sub

Private Sub CDU_Click()
frmCustomer_Collection.Show 1
End Sub

Private Sub cdu2_Click()
frmCustomer_Update.Show 1
End Sub

Private Sub CHA_Click()
Dim str As String
str = "select * from GL_Master order by H_Code"
rptAffairs.rsAffairs.ConnectionString = cnStr
rptAffairs.rsAffairs.Source = str
rptAffairs.Show 1
End Sub

Private Sub ctp_Click()
frmCash_Posting.Show 1
End Sub

Private Sub CTR_Click()
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Dim sort As Integer
Dim Today As Date
    sort = 0
        
    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
           On Error Resume Next
           Today = rs!Today
           rs.Close
        End If

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
If Not rs.EOF Then
    
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
Else
        rs.Close
        rsU.Close
        Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
            Set rsU = New ADODB.Recordset
            rsU.Open "Cash_Print", conn, adOpenDynamic, adLockOptimistic, -1
            rsU.AddNew
            rsU!sl = 1
            rsU!Cash_Open = rs!Cash_Open
            rsU!Cash_Dr = rs!Cash_Dr
            rsU!Cash_Cr = rs!Cash_Cr
            rsU!Cash_Close = rs!Cash_Close
            rsU.Update
            rsU.Close
            rs.Close
        End If
End If
Set rs = New ADODB.Recordset
        str = "select * from Cash_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    rs.Close
  
rptCash.rsCash.ConnectionString = cnStr
rptCash.rsCash.Source = str
rptCash.Show 1

End Sub

Private Sub cupdate_Click()
On Error Resume Next
Dim rs As ADODB.Recordset
Dim str As String
    
    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.MoveFirst
    
        On Error Resume Next
        rs!Cash_Open = 0
        rs!Cash_Dr = 0
        rs!Cash_Cr = 0
        rs!Cash_Close = 0
        rs.Update
        rs.Close
End Sub

Private Sub cust258_Click()
frmStatement.Show 1
End Sub

Private Sub dbbrprt_Click()
Dim sort As Integer
        sort = 0

Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Bank_Master"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Bank_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Bank_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * FROM Bank_Master order by Id"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Bank_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!AC_No = rs!AC_No
        rsN!Bank_Name = rs!Bank_Name
        rsN!Branch_Name = rs!Branch_Name
        rsN!Open_Bal = rs!Open_Bal
        rsN!Deposit = rs!Deposit
        rsN!Withdraw = rs!Withdraw
        rsN!Balance = rs!Balance
        
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
        
            Set rsU = New ADODB.Recordset
            str = "select * from Bank_Print"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
        
        If Not rsU.EOF Then
            rsU!Bank_Open = rs!Bank_Open
            rsU!Bank_Dr = rs!Bank_Dr
            rsU!Bank_Cr = rs!Bank_Cr
            rsU!Bank_Close = rs!Bank_Close
            rsU.Update
            rsU.Close
            rs.Close
        End If
    End If
    
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Bank_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptBank_Position.rsTran.ConnectionString = cnStr
        rptBank_Position.rsTran.Source = str
    
    Else
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Bank_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptBank_Position.rsTran.ConnectionString = cnStr
        rptBank_Position.rsTran.Source = str
End If
rptBank_Position.Show 1
End Sub

Private Sub dbp_Click()
Dim sort As Integer
        sort = 0

Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Branch"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Branch_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Branch_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * FROM Branch order by Id"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Branch_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!Branch_Code = rs!Branch_Code
        rsN!Branch_Name = rs!Branch_Name
        rsN!Branch_Adress = rs!Branch_Address
        rsN!Open_Bal = rs!Open_Bal
        rsN!Dr = rs!Dr
        rsN!Cr = rs!Cr
        rsN!Balance = rs!Balance
        
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Branch_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptBranch_Position.rsBranch.ConnectionString = cnStr
        rptBranch_Position.rsBranch.Source = str
    
    Else
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Branch_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptBranch_Position.rsBranch.ConnectionString = cnStr
        rptBranch_Position.rsBranch.Source = str
End If
rptBranch_Position.Show 1
End Sub

Private Sub dcash_Click()
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Dim sort As Integer
Dim Today As Date
    sort = 0
        
    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
           On Error Resume Next
           Today = rs!Today
           rs.Close
        End If

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
If Not rs.EOF Then
    
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
Else
        rs.Close
        rsU.Close
        Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
            Set rsU = New ADODB.Recordset
            rsU.Open "Cash_Print", conn, adOpenDynamic, adLockOptimistic, -1
            rsU.AddNew
            rsU!sl = 1
            rsU!Cash_Open = rs!Cash_Open
            rsU!Cash_Dr = rs!Cash_Dr
            rsU!Cash_Cr = rs!Cash_Cr
            rsU!Cash_Close = rs!Cash_Close
            rsU.Update
            rsU.Close
            rs.Close
        End If
End If
Set rs = New ADODB.Recordset
        str = "select * from Cash_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    rs.Close
  
rptCash.rsCash.ConnectionString = cnStr
rptCash.rsCash.Source = str
rptCash.Show 1
End Sub

Private Sub dateinc_Click()
frmGL_Tran.Show 1
End Sub

Private Sub datsalinv_Click()
frmDInvo.Show 1
End Sub

Private Sub DClose_Click()
frmDayclose.Show 1
End Sub

Private Sub dcwsr_Click()
Dim rs As ADODB.Recordset
Dim str As String

Set rs = New ADODB.Recordset
        str = "select * FROM Customer_Master Where Dr > 0 order by Customer_Code"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        rs.Close
        rptCustomer_Position.rsTran.ConnectionString = cnStr
        rptCustomer_Position.rsTran.Source = str
End If
rptCustomer_Position.Show 1

End Sub

Private Sub DFP_Click()
Dim sort As Integer
        sort = 0

Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Fixed_Master"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Fixed_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Fixed_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * FROM Fixed_Master order by Prod_Code"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Fixed_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!Prod_Code = rs!Prod_Code
        rsN!Prod_Group = rs!Prod_Group
        rsN!Prod_Name = rs!Prod_Name
        rsN!Prod_Model = rs!Prod_Model
        rsN!Open_Bal = rs!Open_Bal
        rsN!Purchase = rs!Purchase
        rsN!Disposal = rs!Disposal
        rsN!Stock = rs!Stock
        rsN!Prod_Price = rs!Prod_Price
        rsN!Amount = rs!Amount
        
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Fixed_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptFixed_Position.rsTran.ConnectionString = cnStr
        rptFixed_Position.rsTran.Source = str
    
    Else
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Fixed_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptFixed_Position.rsTran.ConnectionString = cnStr
        rptFixed_Position.rsTran.Source = str
End If
rptFixed_Position.Show 1
End Sub

Private Sub drg_Click()
frmDPrint_Loan.Show 1
End Sub

Private Sub dgltrn_Click()
Dim rs As ADODB.Recordset
Dim str As String
Dim Today As Date
        
    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
           On Error Resume Next
           Today = rs!Today
           rs.Close
        End If
  
str = "select * from GL_Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') order by Sl"
rptGL_Tran.rsGL.ConnectionString = cnStr
rptGL_Tran.rsGL.Source = str
If MsgBox("Do you want to Print Todays Trasacion Listing?", vbInformation + vbYesNo, "Print Option") = vbYes Then
rptGL_Tran.Show 1
Else
frmGL_Tran.Show 1
End If
End Sub

Private Sub dplac_Click()
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Dim sort As Integer
Dim srt As Integer
Dim ast As String
Dim Lb As String
Dim Inc As String
Dim Exp As String
    srt = 1
    sort = 0
    
    ast = "ASSET"
    Lb = "LIABILITY"
    Inc = "INCOME"
    Exp = "EXPENSE"

On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Income where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Income where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If


Set rs = New ADODB.Recordset
    str = "select * from GL_Master where Head_Type like '" & Inc & "' order by Sl"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
    str = "select * from GL_Tran where AC_No like '" & rs!AC_No & "' and Date like '" & Today & "' order by Sl"
    rsN.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rsN.MoveFirst
    
    Do While Not rsN.EOF
    On Error Resume Next

        Set rsU = New ADODB.Recordset
            rsU.Open "Income", conn, adOpenDynamic, adLockOptimistic, -1
            rsU.AddNew
                rsU!sl = 1
                rsU!Lb_Sl = sort + 1
                'rsU!Lb_Type = rs!Head_Type
                rsU!Lb_No = rsN!AC_No
                rsU!Lb_Name = rsN!Name
                rsU!Lb_Balance = rsN!Cr
                rsU.Update
                sort = sort + 1
        rsN.MoveNext
      Loop
        rsN.Close
    rs.MoveNext
Loop
    rsU.Close
'------------------------------------------------------------------------------------------
Set rs = New ADODB.Recordset
        str = "select * from GL_Master where Head_Type like '" & Exp & "' order by Sl"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
    str = "select * from GL_Tran where AC_No like '" & rs!AC_No & "' and Date like '" & Today & "' order by Sl"
    rsN.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rsN.MoveFirst

 Do While Not rsN.EOF
    On Error Resume Next

        Set rsU = New ADODB.Recordset
            str = "select * from Income where Lb_Sl like '" & srt & "' order by Sl"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            
        If Not rsU.EOF Then
            
            rsU!Ast_Sl = srt
            'rsN!Ast_Type = rs!Head_Type
            rsU!Ast_No = rsN!AC_No
            rsU!Ast_Name = rsN!Name
            rsU!Ast_Balance = rsN!Dr
            srt = srt + 1
            rsU.Update
        rsN.MoveNext
            
        Else
            rsU.Close
            
        Set rsU = New ADODB.Recordset
            rsU.Open "Income", conn, adOpenDynamic, adLockOptimistic, -1
            rsU.AddNew
            
            rsU!sl = 1
            rsU!Ast_Sl = srt
            'rsU!Ast_Type = rs!Head_Type
            rsU!Ast_No = rsN!AC_No
            rsU!Ast_Name = rsN!Name
            rsU!Ast_Balance = rsN!Dr
            srt = srt + 1
            rsU.Update
        rsN.MoveNext
        End If
    Loop
         rsN.Close
    rs.MoveNext
Loop
       rsU.Close
       rs.Close
         
   
Set rs = New ADODB.Recordset
    str = "select * from Income"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.RecordCount > 0 Then
    rs.Update
    rs.Close
    
        rptDaily_Income.rsIncome.ConnectionString = cnStr
        rptDaily_Income.rsIncome.Source = str
    End If
        rptDaily_Income.Show 1
End Sub

Private Sub dsorm_Click()
Dim sort As Integer
        sort = 0

Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Prod_Master"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Prod_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Prod_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * FROM Raw_Master order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Prod_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!Prod_Code = rs!Prod_Code
        rsN!Prod_Name = rs!Prod_Name
        rsN!Prod_Model = rs!Prod_Model
        rsN!Open_Bal = rs!Open_Bal
        rsN!Purchase = rs!Purchase
        rsN!Sale = rs!Sale
        rsN!Lift = rs!Lift
        rsN!Return = rs!Return
        rsN!Stock = rs!Stock
        rsN!Prod_Price = rs!Prod_Price
        rsN!Com = rs!Com
        rsN!Amount = rs!Amount
      
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Prod_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptStock.rsStock.ConnectionString = cnStr
        rptStock.rsStock.Source = str
    
    Else
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Prod_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptStock.rsStock.ConnectionString = cnStr
        rptStock.rsStock.Source = str
End If
rptStock.Show 1

End Sub

Private Sub dspr_Click()
Dim sort As Integer
        sort = 0

Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Vendor_Master"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Vendor_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Vendor_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * FROM Vendor_Master order by Id"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Vendor_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!Vendor_Code = rs!Vendor_Code
        rsN!Vendor_Name = rs!Vendor_Name
        rsN!Vendor_Adress = rs!Vendor_Address
        rsN!Open_Bal = rs!Open_Bal
        rsN!Dr = rs!Dr
        rsN!Cr = rs!Cr
        rsN!Balance = rs!Balance
        
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Vendor_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptVendor_Position.rsTran.ConnectionString = cnStr
        rptVendor_Position.rsTran.Source = str
    
    Else
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Vendor_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptVendor_Position.rsTran.ConnectionString = cnStr
        rptVendor_Position.rsTran.Source = str
End If
rptVendor_Position.Show 1
End Sub

Private Sub dsr_Click()
Dim sort As Integer
        sort = 0

Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Prod_Master"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Prod_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Prod_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * FROM Prod_Master order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Prod_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!Prod_Code = rs!Prod_Code
        rsN!Prod_Name = rs!Prod_Name
        rsN!Prod_Model = rs!Prod_Model
        rsN!Open_Bal = rs!Open_Bal
        rsN!Purchase = rs!Purchase
        rsN!Sale = rs!Sale
        rsN!Lift = rs!Lift
        rsN!Return = rs!Return
        rsN!Stock = rs!Stock
        rsN!Prod_Price = rs!Prod_Price
        rsN!Com = rs!Com
        rsN!Amount = rs!Amount
      
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Prod_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptStock.rsStock.ConnectionString = cnStr
        rptStock.rsStock.Source = str
    
    Else
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Prod_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptStock.rsStock.ConnectionString = cnStr
        rptStock.rsStock.Source = str
End If
rptStock.Show 1
End Sub

Private Sub dsr123654_Click()
Dim sort As Integer
        sort = 0

Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Prod_Master"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Prod_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Prod_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * FROM Prod_Master order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Prod_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!Prod_Code = rs!Prod_Code
        rsN!Prod_Name = rs!Prod_Name
        rsN!Prod_Model = rs!Prod_Model
        rsN!Open_Bal = rs!Open_Bal
        rsN!Purchase = rs!Purchase
        rsN!Sale = rs!Sale
        rsN!Lift = rs!Lift
        rsN!Return = rs!Return
        rsN!Damage = rs!Damage
        rsN!Stock = rs!Stock
        rsN!Prod_Price = rs!Prod_Price
        rsN!Com = rs!Com
        rsN!Amount = rs!Amount
      
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Prod_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptStock.rsStock.ConnectionString = cnStr
        rptStock.rsStock.Source = str
    
    Else
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Prod_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptStock.rsStock.ConnectionString = cnStr
        rptStock.rsStock.Source = str
End If
rptStock.Show 1
End Sub

Private Sub dsrrpt_Click()
frmDPrint.Show 1
End Sub

Private Sub dsrpt12_Click()
Dim rs As ADODB.Recordset
Dim str As String
    
Set rs = New ADODB.Recordset
        str = "select * FROM Sales_Invoice where cdate(D_ate) Between cdate('" & Today & "') And cdate('" & Today & "') And Sale_Type Like '" & "Cash Sale" & "' order by D_ate, Invo_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        rs.Close
        rptServ_Report.rsTran.ConnectionString = cnStr
        rptServ_Report.rsTran.Source = str
        rptServ_Report.Label11.Caption = "DAILY INVOICE WISE SALES REPORT"
End If
rptServ_Report.Show 1

End Sub

Private Sub DVP_Click()
Dim sort As Integer
        sort = 0

Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Vendor_Master"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Vendor_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Vendor_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * FROM Vendor_Master order by Id"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Vendor_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!Vendor_Code = rs!Vendor_Code
        rsN!Vendor_Name = rs!Vendor_Name
        rsN!Vendor_Adress = rs!Vendor_Address
        rsN!Open_Bal = rs!Open_Bal
        rsN!Dr = rs!Dr
        rsN!Cr = rs!Cr
        rsN!Balance = rs!Balance
        
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Vendor_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptVendor_Position.rsTran.ConnectionString = cnStr
        rptVendor_Position.rsTran.Source = str
    
    Else
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Vendor_Print order by Sl_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptVendor_Position.rsTran.ConnectionString = cnStr
        rptVendor_Position.rsTran.Source = str
End If
rptVendor_Position.Show 1
End Sub

Private Sub DWP_Click()
frmSBWithdraw.Show 1
End Sub

Private Sub e09_Click()
frmLoan_Old.Show 1
End Sub

Private Sub dwcfstm_Click()
frmCash.Show 1
End Sub

Private Sub Employee_Click()
frmEmployee.Show 1
End Sub

Private Sub FA_Click()
frmFixed_stm.Show 1
End Sub

Private Sub FAS_Click()
frmFixed_Asset.Show 1
End Sub

Private Sub Form_Activate()

 Dim pnlStatusIcon As Panel
    wdt = 100
    ht = frmMain.Height
    
    ImageList1.ListImages.Clear
    ImageList1.ImageHeight = 16
    ImageList1.ImageWidth = 16
    'You must place paths to your own icons here
    ImageList1.ListImages.Add , "Calender", LoadPicture(App.Path + "\Photo\" + "NOTE03.ICO")
    ImageList1.ListImages.Add , "Clock", LoadPicture(App.Path + "\Photo\" + "CLOCK02.ICO")
    ImageList1.ListImages.Add , "Code", LoadPicture(App.Path + "\Photo\" + "EYE.ICO")
    ImageList1.ListImages.Add , "Branch", LoadPicture(App.Path + "\Photo\" + "HOUSE.ICO")
    ImageList1.ListImages.Add , "User", LoadPicture(App.Path + "\Photo\" + "FACE03.ICO")
    ImageList1.ListImages.Add , "Name", LoadPicture(App.Path + "\Photo\" + "MISC27.ICO")
    
    StatusBar1.Panels.Clear
    StatusBar1.Height = ScaleY(20, vbPixels, ScaleMode)
    Set pnlStatusIcon = StatusBar1.Panels.Add(1, "Date", "Business Date: " & Today)
    Set pnlStatusIcon = StatusBar1.Panels.Add(2, "Time", "Business Time: " & Time)
    Set pnlStatusIcon = StatusBar1.Panels.Add(3, "Code", "Branch Code: " & Branch_Code)
    Set pnlStatusIcon = StatusBar1.Panels.Add(4, "Branch", "Branch Name: " & Branch_Name)
    Set pnlStatusIcon = StatusBar1.Panels.Add(5, "User", "User ID: " & User_Id)
    Set pnlStatusIcon = StatusBar1.Panels.Add(6, "Name", "User Name: " & User_Name)
    
    pnlStatusIcon.Width = ScaleX(18, vbPixels, ScaleMode)
    Set pnlStatusIcon = Nothing
    
    
    StatusBar1.Panels("Date").Picture = ImageList1.ListImages("Calender").ExtractIcon
    StatusBar1.Panels.Item(1).MinWidth = 3000
    StatusBar1.Panels.Item(1).Alignment = sbrLeft
    StatusBar1.Panels.Item(1).Bevel = sbrRaised

    StatusBar1.Panels("Time").Picture = ImageList1.ListImages("Clock").ExtractIcon
    StatusBar1.Panels.Item(2).MinWidth = 3000
    StatusBar1.Panels.Item(2).Alignment = sbrLeft
    StatusBar1.Panels.Item(2).Bevel = sbrRaised
    
    StatusBar1.Panels("Code").Picture = ImageList1.ListImages("Code").ExtractIcon
    StatusBar1.Panels.Item(3).MinWidth = 2500
    StatusBar1.Panels.Item(3).Alignment = sbrLeft
    StatusBar1.Panels.Item(3).Bevel = sbrRaised
    
    StatusBar1.Panels("Branch").Picture = ImageList1.ListImages("Branch").ExtractIcon
    StatusBar1.Panels.Item(4).MinWidth = 3000
    StatusBar1.Panels.Item(4).Alignment = sbrLeft
    StatusBar1.Panels.Item(4).Bevel = sbrRaised
    
    StatusBar1.Panels("User").Picture = ImageList1.ListImages("User").ExtractIcon
    StatusBar1.Panels.Item(5).MinWidth = 2500
    StatusBar1.Panels.Item(5).Alignment = sbrLeft
    StatusBar1.Panels.Item(5).Bevel = sbrRaised
    
    StatusBar1.Panels("Name").Picture = ImageList1.ListImages("Name").ExtractIcon
    StatusBar1.Panels.Item(6).MinWidth = 3000
    StatusBar1.Panels.Item(6).Alignment = sbrLeft
    StatusBar1.Panels.Item(6).Bevel = sbrRaised
    

If U_Type = "Operator" Then
Call Ope
Else
If U_Type = "Supervisor" Then
Call Super
End If
End If

End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
'Timer4.Enabled = True


End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub gacstm_Click()
frmGLStatement.Show 1
End Sub

Private Sub gdnstm_Click()
frmFactory_Stm.Show 1
End Sub

Private Sub gdnSum_Click()
Dim sort As Integer
        sort = 0

Dim rs As ADODB.Recordset
Dim str As String
 
    Set rs = New ADODB.Recordset
        str = "select * FROM Godown_Master order by Vendor_Name"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    rs.Close
        rptFactory_Position.rsStock.ConnectionString = cnStr
        rptFactory_Position.rsStock.Source = str
    
        rptFactory_Position.rsStock.ConnectionString = cnStr
        rptFactory_Position.rsStock.Source = str
        rptFactory_Position.Label48.Caption = "Issue"
        rptFactory_Position.Label11.Caption = "SUMMARY STOCK REPORT OF GODOWN"
        rptFactory_Position.Label19.Caption = "Receive"

rptFactory_Position.Show 1

End Sub

Private Sub GLAC_Click()
frmGLAccount.Show 1
End Sub

Private Sub GLDT_Click()
Dim rs As ADODB.Recordset
Dim str As String
  
str = "select * from GL_Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') order by Sl"
rptGL_Tran.rsGL.ConnectionString = cnStr
rptGL_Tran.rsGL.Source = str
If MsgBox("Do you want to Print Todays Trasacion Listing?", vbInformation + vbYesNo, "Print Option") = vbYes Then
rptGL_Tran.Show 1
Else
frmGL_Tran.Show 1
End If
End Sub

Private Sub GLST_Click()
frmGLStatement.Show 1
End Sub

Private Sub GLTP_Click()
frmGLPosting.Show 1
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
frmClock.Show
'lblExit.Visible = True
'Image3.Visible = False
'Image2.Visible = True
'lblDate.Visible = False


'Image5.Visible = False
'Image4.Visible = True
'lblTime.Visible = False

'Image7.Visible = False
'Image6.Visible = True
'lblExit.Visible = False
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.Visible = True
Image2.Visible = False
End Sub
Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.Visible = True
Image3.Visible = False
frmProd_Receive.Show 1
End Sub




Private Sub itl_Click()


On Error Resume Next
conn.Execute ("ALTER TABLE Loan_info ADD Form NUMERIC(5) DEFAULT 0 NOT NULL")
conn.Execute ("ALTER TABLE Loan_info ADD Insurance NUMERIC(5) DEFAULT 0 NOT NULL")
conn.Execute ("ALTER TABLE Loan_info ADD Security NUMERIC(5) DEFAULT 0 NOT NULL")
conn.Execute ("ALTER TABLE Loan_info ADD Down_Payment NUMERIC(5) DEFAULT 0 NOT NULL")
conn.Execute ("ALTER TABLE Loan_info ADD Photo TEXT(100)")
conn.Execute ("ALTER TABLE Loan_info ADD Thumb TEXT(100)")
conn.Execute ("ALTER TABLE Loan_info ADD G_ID5 TEXT(10)")
conn.Execute ("ALTER TABLE Loan_info ADD G_Name5 TEXT(25)")
conn.Execute ("ALTER TABLE Loan_info ADD G_ID6 TEXT(10)")
conn.Execute ("ALTER TABLE Loan_info ADD G_Name6 TEXT(25)")
conn.Execute ("ALTER TABLE Loan_info ADD G_ID7 TEXT(10)")
conn.Execute ("ALTER TABLE Loan_info ADD G_Name7 TEXT(25)")
conn.Execute ("ALTER TABLE Loan_info ADD G_ID8 TEXT(10)")
conn.Execute ("ALTER TABLE Loan_info ADD G_Name8 TEXT(25)")
conn.Execute ("ALTER TABLE Loan_info ADD G_ID9 TEXT(10)")
conn.Execute ("ALTER TABLE Loan_info ADD G_Name9 TEXT(25)")
conn.Execute ("ALTER TABLE Loan_info ADD G_ID10 TEXT(10)")
conn.Execute ("ALTER TABLE Loan_info ADD G_Name10 TEXT(25)")
MsgBox "Indexing Completed!", vbInformation, "Index!"
'Dim Today As Date
'Dim rs As ADODB.Recordset
'Dim rsU As ADODB.Recordset
'Dim str As String
'Dim Due As Integer
'On Error Resume Next
'
'Set rs = New ADODB.Recordset
'        str = "select * from Others"
'        rs.Open str, conn
'        rs.MoveFirst
'        If Not rs.EOF Then
'           On Error Resume Next
'           Today = rs!Today
'           rs.Close
'        End If
'
'    Set rsU = New ADODB.Recordset
'        str = "select * from Loan_Master Order by Sl_No"
'        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
'        rsU.MoveFirst
'
'    Do While Not rsU.EOF
'        'rsU!Due = 0
'       ' rsU!Term_Fail = 0
'        If rsU!C_lose = "Yes" Then
'        rsU.MoveNext
'        Else
'        rsU!C_lose = "No"
'        rsU.MoveNext
'        End If
'     Loop
'        rsU.Update
'        rsU.Close
'        MsgBox "Indexing Completed!", vbInformation + vbOKOnly, "Successfull!"
End Sub

Private Sub lac_Click()
frmLoan_Close.Show 1
End Sub

Private Sub LDT_Click()
Dim str As String
str = "select * from Loan_Master order by Sl_no"
rptLoan_Collection.rsLoan.ConnectionString = cnStr
rptLoan_Collection.rsLoan.Source = str
rptLoan_Collection.Show 1
End Sub

Private Sub lodp_Click()
frmBack_Loan.Show 1
End Sub

Private Sub LOPEN_Click()
frmLoanAccount.Show 1
End Sub

Private Sub LST_Click()
frmLoanStatement.Show 1
End Sub

Private Sub LTP_Click()
frmLoan_Posting.Show 1
End Sub

Private Sub ipr_Click()
frmProd_Statement.Show 1
End Sub

Private Sub ips_Click()

End Sub

Private Sub isprs_Click()
frmVendor_Stm.Show 1
End Sub

Private Sub Label6_Click()

End Sub

Private Sub Jsum_Click()
Dim rs As ADODB.Recordset
Dim str As String
 
str = "select * from GL_Tran where cdate(Date) Between cdate('" & Today & "') And cdate('" & Today & "') order by Name"
rptJournal_Summary.rsGL.ConnectionString = cnStr
rptJournal_Summary.rsGL.Source = str
If MsgBox("Do you want to Print Todays Trasacion Listing?", vbInformation + vbYesNo, "Print Option") = vbYes Then
rptJournal_Summary.lblFrom.Caption = Today
rptJournal_Summary.lblTo.Caption = Today
rptJournal_Summary.Show 1
Else
frmGL_Tran.Show 1
End If
End Sub

Private Sub lc_Click()
frmLC.Show 1
End Sub

Private Sub lcr_Click()
frmLC_Statement.Show 1
End Sub

Private Sub loac_Click()
Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * FROM Customer_Master order by Customer_Code"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
If Not rs.EOF Then
    rs.Close
  
    rptAll_Customer.rsDsr.ConnectionString = cnStr
    rptAll_Customer.rsDsr.Source = str
    rptAll_Customer.Show 1
End If
End Sub

Private Sub loac758_Click()
Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * FROM Customer_Info order by Customer"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
If Not rs.EOF Then
    rs.Close
  
    rptAll_Customer.rsDsr.ConnectionString = cnStr
    rptAll_Customer.rsDsr.Source = str
    rptAll_Customer.Show 1
End If
End Sub

Private Sub MClose_Click()
frmM_Close.Show 1
End Sub

Private Sub MRG_Click()
frmMPrint_Loan.Show 1
End Sub

Private Sub mrgn_Click()
frmMPrint.Show 1
End Sub

Private Sub ndr_Click()
Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Cash_Master order by Sl"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
If Not rs.EOF Then
    rs.Close
  
    rptNote.rsCash.ConnectionString = cnStr
    rptNote.rsCash.Source = str
    rptNote.Show 1
End If
End Sub

Private Sub note258_Click()
Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Cash_Master order by Sl"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
If Not rs.EOF Then
    rs.Close
  
    rptNote.rsCash.ConnectionString = cnStr
    rptNote.rsCash.Source = str
    rptNote.Show 1
End If
End Sub

Private Sub PE_Click()
frmProd_Receive.Show 1
End Sub

Private Sub PI_Click()
frmRaw_Receive.Show 1
End Sub

Private Sub pigdn_Click()
frmConsumption.Show 1
End Sub

Private Sub PL_Click()
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Dim sort As Integer
Dim srt As Integer
Dim ast As String
Dim Lb As String
Dim Inc As String
Dim Exp As String
    srt = 1
    sort = 0
    
    ast = "ASSET"
    Lb = "LIABILITY"
    Inc = "INCOME"
    Exp = "EXPENSE"

On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Income where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Income where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If


Set rs = New ADODB.Recordset
    str = "select * from GL_Master where Head_Type like '" & Inc & "' order by Sl"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsU = New ADODB.Recordset
        rsU.Open "Income", conn, adOpenDynamic, adLockOptimistic, -1
        rsU.AddNew
            rsU!sl = 1
            rsU!Lb_Sl = sort + 1
            rsU!Lb_Type = rs!Head_Type
            rsU!Lb_No = rs!AC_No
            rsU!Lb_Name = rs!Head_Name
            rsU!Lb_Balance = rs!Balance
            rsU.Update
            sort = sort + 1
            rs.MoveNext
            Loop
            rs.Close
            rsU.Close

Set rs = New ADODB.Recordset
        str = "select * from GL_Master where Head_Type like '" & Exp & "' order by Sl"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
       str = "select * from Income where Lb_Sl like '" & srt & "' order by Sl"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
            
            If Not rsN.EOF Then
            
            
            rsN!Ast_Sl = srt
            rsN!Ast_Type = rs!Head_Type
            rsN!Ast_No = rs!AC_No
            rsN!Ast_Name = rs!Head_Name
            rsN!Ast_Balance = rs!Balance
            srt = srt + 1
            rsN.Update
            rs.MoveNext
            
        Else
            rsN.Close
            
        Set rsU = New ADODB.Recordset
            rsU.Open "Income", conn, adOpenDynamic, adLockOptimistic, -1
            rsU.AddNew
            
            rsU!sl = 1
            rsU!Ast_Sl = srt
            rsU!Ast_Type = rs!Head_Type
            rsU!Ast_No = rs!AC_No
            rsU!Ast_Name = rs!Head_Name
            rsU!Ast_Balance = rs!Balance
            srt = srt + 1
            rsU.Update
            rs.MoveNext
        End If
        Loop
            
            rs.Close
            rsN.Close
            rsU.Close
   
Set rs = New ADODB.Recordset
    str = "select * from Income"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.RecordCount > 0 Then
    rs.Update
    rs.Close
    
        rptIncome.rsIncome.ConnectionString = cnStr
        rptIncome.rsIncome.Source = str
    End If
        rptIncome.Show 1
End Sub

Private Sub plac_Click()
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Dim sort As Integer
Dim srt As Integer
Dim ast As String
Dim Lb As String
Dim Inc As String
Dim Exp As String
    srt = 1
    sort = 0
    
    ast = "ASSET"
    Lb = "LIABILITY"
    Inc = "INCOME"
    Exp = "EXPENSE"

On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Income where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Income where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If


Set rs = New ADODB.Recordset
    str = "select * from GL_Master where Head_Type like '" & Inc & "' order by Sl"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsU = New ADODB.Recordset
        rsU.Open "Income", conn, adOpenDynamic, adLockOptimistic, -1
        rsU.AddNew
            rsU!sl = 1
            rsU!Lb_Sl = sort + 1
            rsU!Lb_Type = rs!Head_Type
            rsU!Lb_No = rs!AC_No
            rsU!Lb_Name = rs!Head_Name
            rsU!Lb_Balance = rs!Balance
            rsU.Update
            sort = sort + 1
            rs.MoveNext
            Loop
            rs.Close
            rsU.Close

Set rs = New ADODB.Recordset
        str = "select * from GL_Master where Head_Type like '" & Exp & "' order by Sl"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
       str = "select * from Income where Lb_Sl like '" & srt & "' order by Sl"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
            
            If Not rsN.EOF Then
            
            
            rsN!Ast_Sl = srt
            rsN!Ast_Type = rs!Head_Type
            rsN!Ast_No = rs!AC_No
            rsN!Ast_Name = rs!Head_Name
            rsN!Ast_Balance = rs!Balance
            srt = srt + 1
            rsN.Update
            rs.MoveNext
            
        Else
            rsN.Close
            
        Set rsU = New ADODB.Recordset
            rsU.Open "Income", conn, adOpenDynamic, adLockOptimistic, -1
            rsU.AddNew
            
            rsU!sl = 1
            rsU!Ast_Sl = srt
            rsU!Ast_Type = rs!Head_Type
            rsU!Ast_No = rs!AC_No
            rsU!Ast_Name = rs!Head_Name
            rsU!Ast_Balance = rs!Balance
            srt = srt + 1
            rsU.Update
            rs.MoveNext
        End If
        Loop
            
            rs.Close
            rsN.Close
            rsU.Close
   
Set rs = New ADODB.Recordset
    str = "select * from Income"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.RecordCount > 0 Then
    rs.Update
    rs.Close
    
        rptIncome.rsIncome.ConnectionString = cnStr
        rptIncome.rsIncome.Source = str
    End If
        rptIncome.Show 1
End Sub

Private Sub pr_Click()
frmSales.Label2.Caption = "GOODS SALES ENTRY"
frmSales.cmbType.Text = "SALE"
frmSales.Show 1
End Sub

Private Sub prb_Click()
frmSales.Label2.Caption = "GOODS RETURN ENTRY"
frmSales.cmbType.Text = "RETURN"
frmSales.Show 1
End Sub

Private Sub prtn_Click()
frmPReturn.Show 1
End Sub

Private Sub prtngd_Click()
frmGD_Rtn.Show 1
End Sub

Private Sub ps_Click()
frmProd_Statement.Show 1
End Sub

Private Sub RPT_Click()
frmRtprint.Show 1
End Sub

Private Sub pser_Click()
frmProd_Search.Show 1
End Sub

Private Sub rs_Click()
frmRestore.Show 1
End Sub

Private Sub rteter_Click()
frmStaff_Transfer.Show 1
End Sub

Private Sub SAC_Click()
frmSBClose.Show 1
End Sub

Private Sub sas_Click()
frmStaff_stm.Show 1
End Sub

Private Sub SBAC_Click()
frmCustomer_Info.Show 1
End Sub

Private Sub SBDT_Click()
Dim rs As ADODB.Recordset
Dim str As String

Set rs = New ADODB.Recordset
        str = "select * FROM Customer_Master order by Customer_Code"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        rs.Close
        rptCustomer_Position.rsTran.ConnectionString = cnStr
        rptCustomer_Position.rsTran.Source = str
End If
rptCustomer_Position.Show 1
End Sub

Private Sub SBST_Click(Index As Integer)
frmStatement.Show 1
End Sub
Private Sub SBTP_Click()
frmSavings.Show 1
End Sub
Private Sub sl_Click()
frmStatement.Show 1
End Sub

Private Sub sdu_Click()
frmStock_Update.Show 1
End Sub

Private Sub Sm_Click()
frmSamity.Show 1
End Sub

Private Sub smr_Click()
Dim str As String
str = "select * from Samity order by Samity_Code"
rptSamity.rsSamity.ConnectionString = cnStr
rptSamity.rsSamity.Source = str
rptSamity.Show 1
End Sub

Private Sub SOA_Click()
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String

Dim sort As Integer
Dim srt As Integer
Dim ast As String
Dim Lb As String
Dim Inc As String
Dim Exp As String


    srt = 1
    sort = 0
    
    ast = "ASSET"
    Lb = "LIABILITY"
    Inc = "INCOME"
    Exp = "EXPENSE"

On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Trial_Balance where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Trial_Balance where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If


Set rs = New ADODB.Recordset
    str = "select * from GL_Master where Head_Type like '" & Lb & "' or Head_Type like '" & Inc & "' order by AC_No"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsU = New ADODB.Recordset
        rsU.Open "Trial_Balance", conn, adOpenDynamic, adLockOptimistic, -1
        rsU.AddNew
            rsU!sl = 1
            rsU!Lb_Sl = sort + 1
            rsU!Lb_Type = rs!Head_Type
            rsU!Lb_No = rs!AC_No
            rsU!Lb_Name = rs!Head_Name
            rsU!Lb_Balance = rs!Balance
            rsU.Update
            sort = sort + 1
            ProgressBar1.Value = ProgressBar1.Value + 1
            rs.MoveNext
            Loop
            rs.Close
            rsU.Close

Set rs = New ADODB.Recordset
        str = "select * from GL_Master where Head_Type like '" & ast & "' or Head_Type like '" & Exp & "' order by AC_No"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
       str = "select * from Trial_Balance where Lb_Sl like '" & srt & "' order by Sl"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
            
            If Not rsN.EOF Then
            
            
            rsN!Ast_Sl = srt
            rsN!Ast_Type = rs!Head_Type
            rsN!Ast_No = rs!AC_No
            rsN!Ast_Name = rs!Head_Name
            rsN!Ast_Balance = rs!Balance
            srt = srt + 1
            
            ProgressBar1.Value = ProgressBar1.Value + 1
            rsN.Update
            rs.MoveNext
            
        Else
            rsN.Close
            
        Set rsU = New ADODB.Recordset
            rsU.Open "Trial_Balance", conn, adOpenDynamic, adLockOptimistic, -1
            rsU.AddNew
            
            rsU!sl = 1
            rsU!Ast_Sl = srt
            rsU!Ast_Type = rs!Head_Type
            rsU!Ast_No = rs!AC_No
            rsU!Ast_Name = rs!Head_Name
            rsU!Ast_Balance = rs!Balance
            srt = srt + 1
            ProgressBar1.Value = ProgressBar1.Value + 1
            rsU.Update
            rs.MoveNext
        End If
        Loop
            
            rs.Close
            rsN.Close
            rsU.Close
   
   
   Set rs = New ADODB.Recordset
   str = "select * from Trial_Balance Order by Ast_Sl"
   rs.Open str, conn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        
        rptTrial_Balance.rsTrial_Balance.ConnectionString = cnStr
        rptTrial_Balance.rsTrial_Balance.Source = str
        
        MsgBox "Report Generation Completed!", vbInformation, "Complete!"
        rptTrial_Balance.Show 1
    End If
End Sub

Private Sub SOD_Click()
frmSBBack.Show 1
End Sub

Private Sub SP_Click()
Dim sort As Integer
        sort = 0

Dim rs As ADODB.Recordset
Dim str As String
Set rs = New ADODB.Recordset
    str = "select * from Prod_Master"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Prod_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Prod_Print where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
    End If

Set rs = New ADODB.Recordset
        str = "select * FROM Prod_Master order by Prod_Name"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
        rsN.Open "Prod_Print", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        
        rsN!sl = 1
        rsN!Sl_no = sort + 1
        
        rsN!Prod_Code = rs!Prod_Code
        rsN!Prod_Name = rs!Prod_Name
        rsN!Prod_Model = rs!Prod_Model
        rsN!Open_Bal = rs!Open_Bal
        rsN!Purchase = rs!Purchase
        rsN!Sale = rs!Sale
        rsN!Lift = rs!Lift
        rsN!Return = rs!Return
        rsN!Stock = rs!Stock
        rsN!Prod_Price = rs!Prod_Price
        rsN!Com = rs!Com
        rsN!Amount = rs!Amount
      
        sort = sort + 1
        rs.MoveNext
        rsN.Update
        
        Loop
        rs.Close
        rsN.Close
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Prod_Print order by Prod_Name"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptStock_Summary.rsStock.ConnectionString = cnStr
        rptStock_Summary.rsStock.Source = str
    
    Else
    
    Set rs = New ADODB.Recordset
        str = "select * FROM Prod_Print order by Prod_Name"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    
    rs.Close
        rptStock_Summary.rsStock.ConnectionString = cnStr
        rptStock_Summary.rsStock.Source = str
End If
rptStock_Summary.Show 1
End Sub

Private Sub SR_Click()
frmStaff_Print.Show 1
End Sub

Private Sub ssd_Click()
frmSalary_Draw.Show 1
End Sub

Private Sub ssp_Click()
frmSalary.Show 1
End Sub

Private Sub sti_Click()
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim str As String
Dim Due As Integer
On Error Resume Next

    Set rsN = New ADODB.Recordset
        str = "select * from Deposit_Master Order by Sl_No"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        rsNU.MoveFirst
        
    Do While Not rsN.EOF
        'If rsN!C_lose = "Yes" Then
        rsN!C_lose = "No"
        
        rsN!Amount = 0
        rsN!Balance = 0
        rsN!Payment = 0
        rsN!Withdraw = 0
        rsN!Due = 0
        rsN!Advance = 0
        rsN!Security = 0
        rsN!Fine = 0
        rsN!Term_Fail = 0
       
        rsN!Daily_Pay = 0
        rsN!Weekly_Pay = 0
        rsN!Monthly_Pay = 0
        rsN!Yearly_Pay = 0
        
        rsN!Daily_Draw = 0
        rsN!Weekly_Draw = 0
        rsN!Monthly_Draw = 0
        rsN!Yearly_Draw = 0
        
        rsN!Daily_Bal = 0
        rsN!Weekly_Bal = 0
        rsN!Monthly_Bal = 0
        rsN!Yearly_Bal = 0
        
        rsN!Week_1 = 0
        rsN!Week_2 = 0
        rsN!Week_3 = 0
        rsN!Week_4 = 0
        rsN!Week_5 = 0
        
        rsN.MoveNext
       ' Else
        'rsN!C_lose = "No"
       ' rsN.MoveNext
        'End If
     Loop
        rsN.Update
        rsN.Close
        MsgBox "Indexing Completed!", vbInformation + vbOKOnly, "Successfull!"
    End Sub

Private Sub stkrptexp_Click()

Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim str As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Set oSheet = oBook.Worksheets(1)

Set rs = New ADODB.Recordset
        str = "select Prod_Code, Prod_Name, Prod_Model, Stock, Prod_Price, Sale_Price FROM Prod_Master order by Prod_Code"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
    
   'Transfer the data to Excel
   oSheet.Range("A1").Value = "Product Code"
   oSheet.Range("B1").Value = "Product Name"
   oSheet.Range("C1").Value = "Product Model"
   oSheet.Range("D1").Value = "Stock"
   oSheet.Range("E1").Value = "Product Price"
   oSheet.Range("F1").Value = "Sale Price"
   
   oSheet.Range("A2").CopyFromRecordset rs
   oBook.SaveAs "D:\Stock_Report" & Today & ".xls"
   oExcel.Quit
   rs.Close
   MsgBox "Export completed to D:\Sales_Report" & Today & ".xls", vbInformation, "Export!"
    
    Else
        MsgBox "There is no such criteria found!, ", vbCritical + vbOKOnly
        rs.Close
    End If
End Sub

Private Sub Stold_Click()
frmOld_due.Show 1
End Sub

Private Sub su25_Click()
On Error Resume Next
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim str As String
Dim Prod As String

    Set rs = New ADODB.Recordset
        str = "select * From Prod_Master"
        rs.Open str, conn
        rs.MoveFirst
        
    Do While Not rs.EOF
    If rs!Prod_Code = "101" Or rs!Prod_Code = "BKS01" Then
            Prod = rs!Prod_Code

            Set rsU = New ADODB.Recordset
                str = "select * from Prod_Tran where Prod_Code like '" & Prod & "' and sale > 0"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                rsU.MoveFirst
    
                Do While Not rsU.EOF
                    rsU!Prod_Price = 0
                    rsU!Sale_Price = 0
                    rsU.Update
                    rsU.MoveNext
                Loop
                rsU.Close
            rs.MoveNext
    Else
        Prod = rs!Prod_Code

            Set rsU = New ADODB.Recordset
                str = "select * from Prod_Tran where Prod_Code like '" & Prod & "' and sale > 0"
                rsU.Open str, conn, adOpenDynamic, adLockOptimistic
                rsU.MoveFirst
    
                Do While Not rsU.EOF
                    rsU!Prod_Price = rs!Prod_Price
                    rsU.Update
                    rsU.MoveNext
                Loop
                rsU.Close
            rs.MoveNext
        End If
    Loop
    rs.Close
    MsgBox "Product Update Successfull!", vbInformation, "Update!"
End Sub

Private Sub TB_Click()
Dim str As String
str = "select * from GL_Master order by H_Code"
rptTrialbalance.rsTB.ConnectionString = cnStr
rptTrialbalance.rsTB.Source = str
rptTrialbalance.Show 1
End Sub

Private Sub tbsn_Click()
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String

Dim sort As Integer
Dim srt As Integer
Dim ast As String
Dim Lb As String
Dim Inc As String
Dim Exp As String
    
    srt = 1
    sort = 0
    
    ast = "ASSET"
    Lb = "LIABILITY"
    Inc = "INCOME"
    Exp = "EXPENSE"

On Error Resume Next
    Set rsU = New ADODB.Recordset
        str = "select * from Trial_Balance where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsU.EOF Then
    On Error Resume Next
        rsU.Close
        str = "delete from Trial_Balance where sl like '" & 1 & "'"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.Update
        rsU.Close
   End If


Set rs = New ADODB.Recordset
    str = "select * from GL_Master where Head_Type like '" & Lb & "' or Head_Type like '" & Inc & "' order by Sl"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsU = New ADODB.Recordset
        rsU.Open "Trial_Balance", conn, adOpenDynamic, adLockOptimistic, -1
        rsU.AddNew
            rsU!sl = 1
            rsU!Lb_Sl = sort + 1
            rsU!Lb_Type = rs!Head_Type
            rsU!Lb_No = rs!AC_No
            rsU!Lb_Name = rs!Head_Name
            rsU!Lb_Balance = rs!Balance
            rsU.Update
            sort = sort + 1
            ProgressBar1.Value = ProgressBar1.Value + 1
            rs.MoveNext
            Loop
            rs.Close
            rsU.Close

Set rs = New ADODB.Recordset
        str = "select * from GL_Master where Head_Type like '" & ast & "' or Head_Type like '" & Exp & "' order by Sl"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

Do While Not rs.EOF
On Error Resume Next

Set rsN = New ADODB.Recordset
       str = "select * from Trial_Balance where Lb_Sl like '" & srt & "' order by Sl"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
            
            If Not rsN.EOF Then
            
            
            rsN!Ast_Sl = srt
            rsN!Ast_Type = rs!Head_Type
            rsN!Ast_No = rs!AC_No
            rsN!Ast_Name = rs!Head_Name
            rsN!Ast_Balance = rs!Balance
            srt = srt + 1
            ProgressBar1.Value = ProgressBar1.Value + 1
            rsN.Update
            rs.MoveNext
            
        Else
            rsN.Close
            
        Set rsU = New ADODB.Recordset
            rsU.Open "Trial_Balance", conn, adOpenDynamic, adLockOptimistic, -1
            rsU.AddNew
            
            rsU!sl = 1
            rsU!Ast_Sl = srt
            rsU!Ast_Type = rs!Head_Type
            rsU!Ast_No = rs!AC_No
            rsU!Ast_Name = rs!Head_Name
            rsU!Ast_Balance = rs!Balance
            srt = srt + 1
            ProgressBar1.Value = ProgressBar1.Value + 1
            rsU.Update
            rs.MoveNext
        End If
        Loop
            
            rs.Close
            rsN.Close
            rsU.Close
   
   Set rs = New ADODB.Recordset
   str = "select * from Trial_Balance"
   rs.Open str, conn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        rptTrial_Balance.rsTrial_Balance.ConnectionString = cnStr
        rptTrial_Balance.rsTrial_Balance.Source = str
        
        MsgBox "Report Generation Completed!", vbInformation, "Complete!"
        rptTrial_Balance.Show 1
    End If
End Sub

Private Sub Timer1_Timer()
Label5.Caption = "LICENCE TO: " & UCase(Org_Name)
Image1.Width = frmMain.ScaleWidth
Image1.Height = frmMain.ScaleHeight - 300
Image1.Stretch = True


Image10.Top = frmMain.ScaleHeight - 3150
Image10.Left = frmMain.ScaleWidth - 5000

Picture1.Top = frmMain.ScaleHeight - 800
Picture1.Width = frmMain.ScaleWidth
Picture1.Left = 0

'If getVersion = "Windows XP" Then
    'WebBrowser1.Navigate App.Path + "\Vedio\globe.gif"
'End If
'Picture4.Width = frmMain.ScaleWidth - 700
'Picture5.Width = frmMain.ScaleWidth - 975
Label5.Left = frmMain.ScaleWidth - 975
'BeginPlaySound 101
Timer1.Enabled = False

End Sub
Private Sub Transaction_Click()
frmSearch.Show 1
End Sub

Private Sub Timer2_Timer()
StatusBar1.Panels("Time").Text = "Business Time: " & Time
End Sub

Private Sub Timer3_Timer()
Label5.Left = Label5.Left - 50
If Label5.Left < -15550 Then
Label5.Left = frmMain.ScaleWidth - 950
End If
End Sub
Private Sub Timer4_Timer()
'BeginPlaySound 101
End Sub

Private Sub UM_Click()
frmUser.Show 1
End Sub

Private Sub vao_Click()
frmVendor_info.Show 1
End Sub

Private Sub VDP_Click()
frmVendor_Pay.Show 1
End Sub

Private Sub Vm_Click()
Dim str As String
str = "select * from Parameter order by Sl_No"
rptParameter.rsParameter.ConnectionString = cnStr
rptParameter.rsParameter.Source = str
rptParameter.Show 1
End Sub

Private Sub vs_Click()
frmVendor_Stm.Show 1
End Sub

Private Sub WClose_Click()
frmW_Close.Show 1
End Sub

Private Sub WindowsMediaPlayer1_OpenStateChange(ByVal NewState As Long)

End Sub

Private Sub wrg_Click()
frmWPrint_Loan.Show 1
End Sub
