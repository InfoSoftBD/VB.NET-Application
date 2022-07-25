VERSION 5.00
Begin VB.Form frmCustomer_Info 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Information"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   Icon            =   "frmCustomer_Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   135
      ScaleHeight     =   660
      ScaleWidth      =   5325
      TabIndex        =   18
      Top             =   90
      Width           =   5355
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER INFO"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   825
         TabIndex        =   19
         Top             =   60
         Width           =   3660
      End
      Begin VB.Image Image1 
         Height          =   1200
         Left            =   0
         Picture         =   "frmCustomer_Info.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5325
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   5370
      TabIndex        =   13
      Top             =   5025
      Width           =   5400
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         Height          =   555
         Left            =   2670
         TabIndex        =   17
         Top             =   135
         Width           =   1305
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         Height          =   555
         Left            =   90
         TabIndex        =   16
         Top             =   135
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   555
         Left            =   4080
         TabIndex        =   15
         Top             =   135
         Width           =   1170
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update"
         Height          =   555
         Left            =   1410
         TabIndex        =   14
         Top             =   135
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Account Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3945
      Left            =   120
      TabIndex        =   0
      Top             =   930
      Width           =   5340
      Begin VB.ComboBox cmbP_Type 
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
         Left            =   3690
         TabIndex        =   24
         Text            =   " "
         Top             =   3285
         Width           =   1500
      End
      Begin VB.ComboBox cmbAgent_Name 
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
         Left            =   1620
         TabIndex        =   21
         Text            =   " "
         Top             =   3285
         Width           =   2040
      End
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
         Left            =   180
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   3285
         Width           =   1365
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
         Left            =   1965
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   1530
         Width           =   3210
      End
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
         Left            =   180
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   690
         Width           =   1605
      End
      Begin VB.ComboBox cmbType 
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
         Height          =   360
         ItemData        =   "frmCustomer_Info.frx":A4D0
         Left            =   180
         List            =   "frmCustomer_Info.frx":A4D2
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1545
         Width           =   1635
      End
      Begin VB.TextBox txtDate 
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
         Height          =   405
         Left            =   1965
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   675
         Width           =   1350
      End
      Begin VB.TextBox txtPresent 
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
         Height          =   405
         Left            =   180
         TabIndex        =   2
         Text            =   "Text15"
         Top             =   2385
         Width           =   3435
      End
      Begin VB.TextBox txtMobile 
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
         Height          =   405
         Left            =   3690
         TabIndex        =   1
         Text            =   "Text17"
         Top             =   2385
         Width           =   1500
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Type"
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
         TabIndex        =   25
         Top             =   2925
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Name"
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
         Left            =   1665
         TabIndex        =   23
         Top             =   2925
         Width           =   1080
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent ID"
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
         TabIndex        =   22
         Top             =   2925
         Width           =   750
      End
      Begin VB.Label Label4 
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
         Left            =   1965
         TabIndex        =   12
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label3 
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
         Left            =   180
         TabIndex        =   11
         Top             =   1170
         Width           =   1320
      End
      Begin VB.Label Label31 
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
         Left            =   180
         TabIndex        =   10
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Customer"
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
         Left            =   1980
         TabIndex        =   9
         Top             =   1170
         Width           =   1620
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         TabIndex        =   8
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No."
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
         TabIndex        =   7
         Top             =   2025
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmCustomer_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Dim strSQL As String
Private Sub Cust_Id()
        txtID.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer_Code FROM Customer_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtID.AddItem rs!Customer_Code
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
End Sub
Private Sub Cust_Name()
        txtName.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Customer_Name FROM Customer_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtName.AddItem rs!Customer_Name
        rs.MoveNext
        Loop
        rs.Close
         Else
        rs.Close
        End If
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
Private Sub Agent_Name()
   cmbAgent_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name FROM Employee"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbAgent_Name.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
    Else
    rs.Close
    End If
End Sub
Private Sub Prod_Type()
        cmbP_Type.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name FROM Prod_Master"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        cmbP_Type.AddItem rs!Prod_Name
        rs.MoveNext
        Loop
        rs.Close
    Else
    rs.Close
    End If
End Sub
Private Sub ComboType()
        cmbType.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT Type FROM Parameter GROUP BY Type ORDER BY Type ASC"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
  
        Do While Not rs.EOF
        On Error Resume Next
        cmbType.AddItem rs!Type
        rs.MoveNext
        Loop
        rs.Close
        Else
        rs.Close
        End If
End Sub
Private Sub Type_Name()
If cmbType.Text = "" Then
Exit Sub
End If
On Error Resume Next
 Dim DPS_Type As String
 DPS_Type = cmbType.Text
   Set rs = New ADODB.Recordset
        str = "select DISTINCT Type, Description from Parameter where Type like '" & DPS_Type & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
        On Error Resume Next
        lblDescription.Caption = rs!Description
        rs.Close
    End If
End Sub
Private Sub C_om()

If cmbType.Text = "" Then
Exit Sub
End If
On Error Resume Next
 Dim DPS_Type As String
 DPS_Type = cmbType.Text
   Set rs = New ADODB.Recordset
        str = "select DISTINCT Type, Com from Parameter where Type like '" & DPS_Type & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
        On Error Resume Next
        txtCommission.Text = rs!Com
        rs.Close
    End If
End Sub


Private Sub clearTextboxes()
        txtID.Text = ""
        cmbType.Text = ""
        txtDate.Text = ""
        
        txtName.Text = ""
        txtPresent.Text = ""
        txtMobile.Text = ""
        cmbAgent_Code.Text = ""
        cmbAgent_Name.Text = ""
        cmbP_Type.Text = ""
        
End Sub

Private Sub cmbTerm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCommission.SelStart = 0
    txtCommission.SelLength = Len(txtCommission.Text)
    txtCommission.SetFocus
End If
End Sub

Private Sub cmbTerm_LostFocus()
Call C_om
End Sub

Private Sub cmbAgent_Code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbAgent_Name.SetFocus
End If
End Sub

Private Sub cmbAgent_Code_LostFocus()
Dim ID As String
    ID = cmbAgent_Code.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Employee where Staff_Id like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
'        Call clearTextboxes
        cmbAgent_Code.Text = rs!Staff_Id
        cmbAgent_Name.Text = rs!Name
        rs.Close
        cmbAgent_Name.SetFocus
    Else
    
    Exit Sub
    End If

End Sub

Private Sub cmbAgent_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdAdd.Enabled = True Then
    cmdAdd.SetFocus
    Else
    If cmdUpdate.Enabled = True Then
    cmdUpdate.SetFocus
    End If
    End If
End If
End Sub

Private Sub cmbAgent_Name_LostFocus()
Dim ID As String
    ID = cmbAgent_Name.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Employee where Name like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
'        Call clearTextboxes
        cmbAgent_Code.Text = rs!Staff_Id
        cmbAgent_Name.Text = rs!Name
        rs.Close
        'cmbAgent_Name.SetFocus
    Else
    
    Exit Sub
    End If
End Sub

Private Sub cmbType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtName.SetFocus
End If
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
    Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtID.Text & "'"
        rs.Open str, conn
        
     If Not rs.EOF Then
        MsgBox "Customer Already Exist!", vbCritical, "Customer Info!"
        rs.Close
        Exit Sub
    
    cmdAdd.Enabled = False
   
    Else
        
        rs.Close
    
        Set rsN = New ADODB.Recordset
            rsN.Open "Customer_Master", conn, adOpenDynamic, adLockOptimistic, -1
            rsN.AddNew
        
            rsN!Date = txtDate.Text
            rsN!Customer_Code = txtID.Text
            rsN!Customer_Name = txtName.Text
            rsN!Customer_Address = txtPresent.Text
            rsN!Customer_Type = cmbType.Text
            rsN!Mobile = txtMobile.Text
            rsN!Open_Bal = 0
            rsN!Dr = 0
            rsN!Cr = 0
            rsN!Balance = 0
            rsN!Agent_Code = cmbAgent_Code.Text
            rsN!Agent_Name = cmbAgent_Name.Text
            rsN!Prod_Type = cmbP_Type.Text
            rsN.Update
            rsN.Close
    
    cmdAdd.Enabled = False
    
    MsgBox "Account Open Successfull! Your Account no.: " & txtID.Text, vbInformation, "Customer Info"
    Call Cust_Id
    Call Cust_Name
    Call clearTextboxes
    Exit Sub
End If
End Sub

Private Sub cmdPrint_Click()
On Error Resume Next
If txtID.Text = "" Then
Exit Sub
Else

Set rs = New ADODB.Recordset
    str = "select * from Customer_Info where Customer like '" & txtID.Text & "'"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
    rs.Close
    rptCustomer_Info.lblBranch.Caption = Branch_Name & " Branch, " & Branch_Address
    rptCustomer_Info.lblUser_Id.Caption = User_Id
    rptCustomer_Info.lblUser_Name.Caption = User_Name
    rptCustomer_Info.rsPeronal_info.ConnectionString = cnStr
    rptCustomer_Info.rsPeronal_info.Source = str
    End If
    
    rptCustomer_Info.Show 1
    cmdPrint.Enabled = False
    cmdUpdate.Enabled = False
End If
End Sub

Private Sub cmdUpdate_Click()
 On Error Resume Next
    
    Set rsN = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & txtID.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
    
    If Not rsN.EOF Then
           
        rsN!Customer_Code = txtID.Text
        rsN!Customer_Type = cmbType.Text
        rsN!Date = txtDate.Text
        
        rsN!Customer_Name = txtName.Text
        rsN!Customer_Address = txtPresent.Text
        rsN!Mobile = txtMobile.Text
        rsN!Agent_Code = cmbAgent_Code.Text
        rsN!Agent_Name = cmbAgent_Name.Text
        rsN!Prod_Type = cmbP_Type.Text
        rsN.Update
        rsN.Close
        
        MsgBox "Account Update Successfull! ", vbInformation, "Customer Info"
    
    Call clearTextboxes
    cmdUpdate.Enabled = False
    cmdPrint.Enabled = False
    Else
        rsN.Close
        MsgBox "Invalid Customer ID", vbCritical, "Error"
    
        cmdUpdate.Enabled = False
        cmdPrint.Enabled = False
    End If
    
    Exit Sub
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Cust_Id
Call Cust_Name
Call Agent_Id
Call Agent_Name
Call Prod_Type
Call ComboType
Call clearTextboxes
txtDate.Text = Today
End Sub

Private Sub txtCommission_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtName.SetFocus
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbType.SetFocus
End If
End Sub

Private Sub txtId_LostFocus()
On Error Resume Next
Dim mid As Integer
    mid = 0
Dim ID As String
    ID = txtID.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Customer_Master where Customer_Code like '" & ID & "'"
        rs.Open str, conn
       
    If Not rs.EOF Then
    
    MsgBox "Customer exist!", vbInformation, "Customer Info!"
    
    
        txtID.Text = rs!Customer_Code
        txtDate.Text = Today
        cmbType.Text = rs!Customer_Type
        'txtCommission.Text = rs!Com
        
        txtName.Text = rs!Customer_Name
        txtPresent.Text = rs!Customer_Address
        txtMobile.Text = rs!Mobile
        cmbAgent_Code.Text = rs!Agent_Code
        cmbAgent_Name.Text = rs!Agent_Name
        cmbP_Type.Text = rs!Prod_Type
        rs.Close
        
        cmdUpdate.Enabled = True
        cmdAdd.Enabled = False
        cmdPrint.Enabled = True
    
Else
    rs.Close
    
    If MsgBox("Do you want add new Customer?", vbInformation + vbYesNo, "Add New") = vbYes Then
        
        Set rsU = New ADODB.Recordset
        str = "select * from Customer_Master order by Customer_Code"
        rsU.Open str, conn
    
            If Not rsU.EOF Then
                rsU.MoveFirst
                
                Do While Not rsU.EOF
                    mid = mid + 1
                    rsU.MoveNext
                Loop
                    rsU.Close
                    mid = mid + 1
            Else
                    mid = "0001"
            End If
    
        Call clearTextboxes
        
        txtID.Text = Format$(Val(mid), "000#")
        txtDate.Text = Today
        cmdAdd.Enabled = True
    Else
        Call clearTextboxes
        cmbType.SetFocus
    End If
End If

End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbAgent_Code.SetFocus
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPresent.SetFocus
End If
End Sub

Private Sub txtPresent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMobile.SetFocus
End If
End Sub

Private Sub txtTelephone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMobile.SetFocus
End If
End Sub




