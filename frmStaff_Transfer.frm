VERSION 5.00
Begin VB.Form frmStaff_Transfer 
   BackColor       =   &H00008000&
   Caption         =   "Staff Transer"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   Icon            =   "frmStaff_Transfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   690
      Left            =   135
      ScaleHeight     =   630
      ScaleWidth      =   7230
      TabIndex        =   31
      Top             =   90
      Width           =   7290
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STAFF TRANSFER POSTING"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1005
         TabIndex        =   32
         Top             =   90
         Width           =   4800
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   180
      ScaleHeight     =   555
      ScaleWidth      =   7200
      TabIndex        =   26
      Top             =   5040
      Width           =   7260
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   3780
         TabIndex        =   30
         Top             =   90
         Width           =   1305
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Replace"
         Height          =   375
         Left            =   450
         TabIndex        =   29
         Top             =   90
         Width           =   1305
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   375
         Left            =   5490
         TabIndex        =   28
         Top             =   90
         Width           =   1305
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   2115
         TabIndex        =   27
         Top             =   90
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Replace with Staff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   135
      TabIndex        =   13
      Top             =   2970
      Width           =   7305
      Begin VB.ComboBox rAM_Code 
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
         Left            =   1425
         TabIndex        =   19
         Text            =   "Combo2"
         Top             =   1305
         Width           =   1455
      End
      Begin VB.ComboBox rAM_Name 
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
         Left            =   4470
         TabIndex        =   18
         Text            =   "Combo2"
         Top             =   1305
         Width           =   2670
      End
      Begin VB.ComboBox rFO_Code 
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
         Left            =   1440
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   315
         Width           =   1455
      End
      Begin VB.ComboBox rDPO_Name 
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
         Left            =   4485
         TabIndex        =   16
         Text            =   "Combo2"
         Top             =   810
         Width           =   2670
      End
      Begin VB.ComboBox rFO_Name 
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
         Left            =   4485
         TabIndex        =   15
         Text            =   "Combo2"
         Top             =   315
         Width           =   2670
      End
      Begin VB.ComboBox rDPO_Code 
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
         Left            =   1440
         TabIndex        =   14
         Text            =   "Combo2"
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AM Name"
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
         Left            =   3495
         TabIndex        =   25
         Top             =   1350
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AM Code"
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
         TabIndex        =   24
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DPO/BM Code"
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
         Left            =   90
         TabIndex        =   23
         Top             =   855
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DPO/BM Name"
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
         Left            =   3015
         TabIndex        =   22
         Top             =   855
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.O. Name"
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
         TabIndex        =   21
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.O. Code"
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
         Left            =   435
         TabIndex        =   20
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Replaceing Staff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   135
      TabIndex        =   0
      Top             =   900
      Width           =   7305
      Begin VB.ComboBox cmbDPO_Code 
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
         Left            =   1440
         TabIndex        =   6
         Text            =   "Combo2"
         Top             =   810
         Width           =   1455
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
         Left            =   4485
         TabIndex        =   5
         Text            =   "Combo2"
         Top             =   315
         Width           =   2670
      End
      Begin VB.ComboBox cmbDPO_Name 
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
         Left            =   4485
         TabIndex        =   4
         Text            =   "Combo2"
         Top             =   810
         Width           =   2670
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
         Left            =   1440
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   315
         Width           =   1455
      End
      Begin VB.ComboBox cmbAM_Name 
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
         Left            =   4470
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   1305
         Width           =   2670
      End
      Begin VB.ComboBox cmbAM_Code 
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
         Left            =   1425
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   1305
         Width           =   1455
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.O. Code"
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
         Left            =   435
         TabIndex        =   12
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.O. Name"
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
         TabIndex        =   11
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DPO/BM Name"
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
         Left            =   3015
         TabIndex        =   10
         Top             =   855
         Width           =   1350
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DPO/BM Code"
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
         Left            =   90
         TabIndex        =   9
         Top             =   855
         Width           =   1290
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AM Code"
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
         TabIndex        =   8
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AM Name"
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
         Left            =   3495
         TabIndex        =   7
         Top             =   1350
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmStaff_Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub clearText()
cmbFO_Code.Text = ""
cmbFO_Name.Text = ""
cmbDPO_Code.Text = ""
cmbDPO_Name.Text = ""
cmbAM_Code.Text = ""
cmbAM_Name.Text = ""

rFO_Code.Text = ""
rFO_Name.Text = ""
rDPO_Code.Text = ""
rDPO_Name.Text = ""
rAM_Code.Text = ""
rAM_Name.Text = ""
End Sub

Private Sub ComboFO_Code()
Dim FO As String
FO = "FO"
        cmbFO_Code.Clear
        rFO_Code.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_ID, Designation FROM Employee Where Designation like '" & FO & "' and center_code like '" & Branch_Code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbFO_Code.AddItem rs!Staff_Id
        rFO_Code.AddItem rs!Staff_Id
        rs.MoveNext
        Loop
        rs.Close
        Else
        cmbFO_Code.Text = ""
        rFO_Code.Text = ""
        rs.Close
        End If
End Sub
Private Sub ComboFO_Name()
Dim FO As String
FO = "FO"
        cmbFO_Name.Clear
        rFO_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name, Designation FROM Employee Where Designation like '" & FO & "' and center_code like '" & Branch_Code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbFO_Name.AddItem rs!Name
        rFO_Name.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        cmbFO_Name.Text = ""
        rFO_Name.Text = ""
        rs.Close
        End If
End Sub
Private Sub ComboDPO_Code()
Dim DPO As String
DPO = "DPO"
        
        cmbDPO_Code.Clear
        rDPO_Code.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_Id, Designation FROM Employee Where Designation like '" & DPO & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbDPO_Code.AddItem rs!Staff_Id
        rDPO_Code.AddItem rs!Staff_Id
        rs.MoveNext
        Loop
        rs.Close
        Else
        cmbDPO_Code.Text = ""
        rDPO_Code.Text = ""
        rs.Close
        End If
End Sub
Private Sub ComboAM_Code()
Dim AM As String
AM = "AM"
        
        cmbAM_Code.Clear
        rAM_Code.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_Id, Designation FROM Employee Where Designation like '" & AM & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbAM_Code.AddItem rs!Staff_Id
        rAM_Code.AddItem rs!Staff_Id
        rs.MoveNext
        Loop
        rs.Close
        Else
        cmbAM_Code.Text = ""
        rAM_Code.Text = ""
        rs.Close
        End If
End Sub


Private Sub ComboDPO_Name()
Dim DPO As String
DPO = "DPO"
        
        cmbDPO_Name.Clear
        rDPO_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name, Designation FROM Employee Where Designation like '" & DPO & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbDPO_Name.AddItem rs!Name
        rDPO_Name.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        cmbDPO_Name.Text = ""
        rDPO_Name.Text = ""
        rs.Close
        End If
End Sub
Private Sub ComboAM_Name()
Dim AM As String
AM = "AM"
        
        cmbDPO_Name.Clear
        rDPO_Name.Clear
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Name, Designation FROM Employee Where Designation like '" & AM & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
        On Error Resume Next
        cmbAM_Name.AddItem rs!Name
        rAM_Name.AddItem rs!Name
        rs.MoveNext
        Loop
        rs.Close
        Else
        cmbAM_Name.Text = ""
        rAM_Name.Text = ""
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

Dim code As String
code = cmbFO_Code.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_ID, Name, S_Code, S_Name FROM Employee where Staff_Id like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        cmbFO_Name.Text = rs!Name
        cmbDPO_Code.Text = rs!S_Code
        cmbDPO_Name.Text = rs!S_Name
        rs.Close
        End If
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
conn.Execute ("UPDATE Personal_Master SET FO_Code = '" & rFO_Code.Text & "', FO_Name = '" & rFO_Name.Text & "', DPO_Code = '" & rDPO_Code.Text & "', DPO_Name = '" & rDPO_Name.Text & "', AM_Code = '" & rAM_Code.Text & "', AM_Name = '" & rAM_Name.Text & "' WHERE FO_Code LIKE '" & cmbFO_Code.Text & "'")
conn.Execute ("UPDATE Deposit_Master SET FO_Code = '" & rFO_Code.Text & "', FO_Name = '" & rFO_Name.Text & "', DPO_Code = '" & rDPO_Code.Text & "', DPO_Name = '" & rDPO_Name.Text & "', AM_Code = '" & rAM_Code.Text & "', AM_Name = '" & rAM_Name.Text & "' WHERE FO_Code LIKE '" & cmbFO_Code.Text & "'")
conn.Execute ("UPDATE Loan_Info SET FO_Code = '" & rFO_Code.Text & "', FO_Name = '" & rFO_Name.Text & "', DPO_Code = '" & rDPO_Code.Text & "', DPO_Name = '" & rDPO_Name.Text & "', AM_Code = '" & rAM_Code.Text & "', AM_Name = '" & rAM_Name.Text & "' WHERE FO_Code LIKE '" & cmbFO_Code.Text & "'")
conn.Execute ("UPDATE Loan_Master SET FO_Code = '" & rFO_Code.Text & "', FO_Name = '" & rFO_Name.Text & "', DPO_Code = '" & rDPO_Code.Text & "', DPO_Name = '" & rDPO_Name.Text & "', AM_Code = '" & rAM_Code.Text & "', AM_Name = '" & rAM_Name.Text & "' WHERE FO_Code LIKE '" & cmbFO_Code.Text & "'")
MsgBox "Transfer Completed!", vbInformation, "Transfer!"
clearText
cmbFO_Code.SetFocus
End Sub
Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
        Call ComboFO_Code
        Call ComboFO_Name
        Call ComboDPO_Code
        Call ComboDPO_Name
        Call ComboAM_Code
        Call ComboAM_Name
        clearText
End Sub

Private Sub rFO_Code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rFO_Name.SetFocus
End If
End Sub
Private Sub rFO_Code_LostFocus()
If rFO_Code.Text = "" Then
Exit Sub
End If

Dim code As String
code = rFO_Code.Text
On Error Resume Next
        Set rs = New ADODB.Recordset
        str = "SELECT DISTINCT Staff_ID, Name, S_Code, S_Name FROM Employee where Staff_Id like '" & code & "'"
        rs.Open str, conn
        If Not rs.EOF Then
        rFO_Name.Text = rs!Name
        rDPO_Code.Text = rs!S_Code
        rDPO_Name.Text = rs!S_Name
        rs.Close
        End If
End Sub
