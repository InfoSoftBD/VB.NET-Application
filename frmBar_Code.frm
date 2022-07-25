VERSION 5.00
Begin VB.Form frmBar_Code 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bar Code Print"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   Icon            =   "frmBar_Code.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   150
      ScaleHeight     =   2490
      ScaleWidth      =   6225
      TabIndex        =   2
      Top             =   990
      Width           =   6255
      Begin VB.TextBox txtPrice 
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
         Left            =   1665
         TabIndex        =   11
         Top             =   1755
         Width           =   2130
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   555
         Left            =   4320
         TabIndex        =   10
         Top             =   1455
         Width           =   1605
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print Bar Code"
         Height          =   555
         Left            =   4290
         TabIndex        =   9
         Top             =   375
         Width           =   1605
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
         Left            =   1665
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   765
         Width           =   2130
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
         Left            =   1665
         TabIndex        =   4
         Text            =   "Combo2"
         Top             =   1245
         Width           =   2130
      End
      Begin VB.ComboBox txtProd_Sl 
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
         Left            =   1665
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   285
         Width           =   2130
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Price"
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
         Left            =   270
         TabIndex        =   12
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         Height          =   2115
         Left            =   4140
         Top             =   135
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   270
         TabIndex        =   8
         Top             =   1305
         Width           =   975
      End
      Begin VB.Label Label5 
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
         Left            =   270
         TabIndex        =   7
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label1 
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
         Left            =   270
         TabIndex        =   6
         Top             =   315
         Width           =   1185
      End
      Begin VB.Shape Shape3 
         Height          =   2115
         Left            =   150
         Top             =   150
         Width           =   3915
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   150
      ScaleHeight     =   705
      ScaleWidth      =   6225
      TabIndex        =   0
      Top             =   150
      Width           =   6255
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BAR CODE PRINT"
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
         Height          =   465
         Left            =   1290
         TabIndex        =   1
         Top             =   135
         Width           =   3750
      End
      Begin VB.Image Image3 
         Height          =   690
         Left            =   0
         Picture         =   "frmBar_Code.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frmBar_Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub Prod_Code()
        txtProd_Sl.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Code FROM Prod_Master"
        rs.Open str, conn
        rs.MoveFirst
  
    Do While Not rs.EOF
        On Error Resume Next
        txtProd_Sl.AddItem rs!Prod_Code
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

Private Sub cmbProd_Model_GotFocus()
Dim Prod As String
Prod = cmbProd_Name.Text
        
        cmbProd_Model.Clear
        Set rsN = New ADODB.Recordset
        str = "SELECT DISTINCT Prod_Name, Prod_Model FROM Prod_Master where Prod_Name Like '" & Prod & "'"
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

Private Sub cmbProd_Model_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub cmbProd_Model_LostFocus()
Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Name Like '" & cmbProd_Name.Text & "' AND Prod_Model Like '" & cmbProd_Model.Text & "' order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
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

Private Sub Command1_Click()
rptBar_Code.lblName1.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode1.Caption = txtProd_Sl.Text
rptBar_Code.lblCode1.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice1.Caption = Format(Val(txtPrice.Text), "0.00")

'rptBar_Code.lblName2.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
'rptBar_Code.Barcode2.Caption = txtProd_Sl.Text
'rptBar_Code.lblCode2.Caption = txtProd_Sl.Text
'rptBar_Code.lblPrice2.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName3.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode3.Caption = txtProd_Sl.Text
rptBar_Code.lblCode3.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice3.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName4.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode4.Caption = txtProd_Sl.Text
rptBar_Code.lblCode4.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice4.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName5.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode5.Caption = txtProd_Sl.Text
rptBar_Code.lblCode5.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice5.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName6.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode6.Caption = txtProd_Sl.Text
rptBar_Code.lblCode6.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice6.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName7.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode7.Caption = txtProd_Sl.Text
rptBar_Code.lblCode7.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice7.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName8.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode8.Caption = txtProd_Sl.Text
rptBar_Code.lblCode8.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice8.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName9.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode9.Caption = txtProd_Sl.Text
rptBar_Code.lblCode9.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice9.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName10.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode10.Caption = txtProd_Sl.Text
rptBar_Code.lblCode10.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice10.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName11.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode11.Caption = txtProd_Sl.Text
rptBar_Code.lblCode11.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice11.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName12.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode12.Caption = txtProd_Sl.Text
rptBar_Code.lblCode12.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice12.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName13.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode13.Caption = txtProd_Sl.Text
rptBar_Code.lblCode13.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice13.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName14.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode14.Caption = txtProd_Sl.Text
rptBar_Code.lblCode14.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice14.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName15.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode15.Caption = txtProd_Sl.Text
rptBar_Code.lblCode15.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice15.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName16.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode16.Caption = txtProd_Sl.Text
rptBar_Code.lblCode16.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice16.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName17.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode17.Caption = txtProd_Sl.Text
rptBar_Code.lblCode17.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice17.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.lblName18.Caption = cmbProd_Name.Text & ", " & cmbProd_Model.Text
rptBar_Code.Barcode18.Caption = txtProd_Sl.Text
rptBar_Code.lblCode18.Caption = txtProd_Sl.Text
rptBar_Code.lblPrice18.Caption = Format(Val(txtPrice.Text), "0.00")

rptBar_Code.Show 1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call Prod_Name
    Call Prod_Code
    Call Prod_Model
End Sub
Private Sub txtProd_Sl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProd_Name.SetFocus
End If
End Sub

Private Sub txtProd_Sl_LostFocus()
If txtProd_Sl.Text = "" Then
Exit Sub
End If
    Dim Prod As String
    Prod = txtProd_Sl.Text
    
    Set rsN = New ADODB.Recordset
        str = "SELECT * FROM Prod_Master where Prod_Code Like '" & Prod & "'order by Prod_code"
        rsN.Open str, conn
        
        If Not rsN.EOF Then
            txtProd_Sl.Text = rsN!Prod_Code
            cmbProd_Name.Text = rsN!Prod_Name
            cmbProd_Model.Text = rsN!Prod_Model
            txtPrice.Text = rsN!Sale_Price
            rsN.Close
            Command1.SetFocus
        Else
        MsgBox "Invalid Product Code!", vbCritical, "Error"
        cmbProd_Name.SetFocus
        End If
End Sub
