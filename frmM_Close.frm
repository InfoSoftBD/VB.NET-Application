VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmM_Close 
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Month Close"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   Icon            =   "frmM_Close.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   180
      ScaleHeight     =   705
      ScaleWidth      =   5115
      TabIndex        =   10
      Top             =   180
      Width           =   5145
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTH CLOSE PROCESS"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   135
         TabIndex        =   11
         Top             =   45
         Width           =   4635
      End
      Begin VB.Image Image3 
         Height          =   735
         Left            =   0
         Picture         =   "frmM_Close.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5205
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   180
      ScaleHeight     =   705
      ScaleWidth      =   5115
      TabIndex        =   7
      Top             =   3240
      Width           =   5145
      Begin VB.CommandButton Command1 
         Caption         =   "Start Processing"
         Height          =   510
         Left            =   360
         TabIndex        =   9
         Top             =   90
         Width           =   1545
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   510
         Left            =   3150
         TabIndex        =   8
         Top             =   90
         Width           =   1545
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   180
      ScaleHeight     =   1515
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   1080
      Width           =   5145
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2430
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   810
         Width           =   2400
      End
      Begin VB.TextBox txtLast 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2430
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   135
         Width           =   2400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Todays Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   6
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Closed on"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   5
         Top             =   225
         Width           =   2085
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   90
      Top             =   90
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   180
      ScaleHeight     =   315
      ScaleWidth      =   5085
      TabIndex        =   0
      Top             =   2790
      Width           =   5145
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "frmM_Close"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub Command1_Click()
On Error Resume Next
Dim Inc As String
Dim Exp As String
    Inc = "INCOME"
    Exp = "EXPENSE"

If MsgBox("Are you sure you want start Day close process?", vbInformation + vbYesNo, "Week Close") = vbYes Then
If MsgBox("Have you taken Backup?", vbInformation + vbYesNo, "Day Close") = vbYes Then

    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.MoveFirst
        If Not rs.EOF Then
        On Error Resume Next
        txtLast.Text = Today
        rs!Month_Day = Today
        rs.Update
        rs.Close
        End If
'-----------------------------------------------------------------------


On Error Resume Next
Set rsU = New ADODB.Recordset
    str = "select * from GL_Master where Head_Type like '" & Inc & "' order by Sl"
    rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    rsU.MoveFirst

Do While Not rsU.EOF
'Dr------------------------------------------------
   Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100152 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance + rsU!Balance
        rs!Date = txtDate.Text
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = rsU!Head_Name & " " & "Realized"
        rsN!Dr = 0
        rsN!Cr = rsU!Balance
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
'Cr------------------------------------------------
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rsU!Head_Name
        rsN!Description = "Income Realized"
        rsN!Dr = rsU!Balance
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        
     rsU!Balance = rsU!Balance - rsU!Balance
        rsU!Date = txtDate.Text
    
'---------------------------------------------
'---------------------------------------------
        rsU.MoveNext
 Loop
    rsU.Update
    rsU.Close

'-----------------------------------------------------------------------------------------
Set rsU = New ADODB.Recordset
    str = "select * from GL_Master where Head_Type like '" & Exp & "' order by Sl"
    rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    rsU.MoveFirst

Do While Not rsU.EOF
'Dr------------------------------------------------
   Set rs = New ADODB.Recordset
        str = "select * from GL_Master where AC_No like '" & 100152 & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs!Balance = rs!Balance - rsU!Balance
        rs!Date = txtDate.Text
        rs.Update
    
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rs!Head_Name
        rsN!Description = rsU!Head_Name & " " & "Realized"
        rsN!Dr = rsU!Balance
        rsN!Cr = 0
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        rs.Close
    
'Cr------------------------------------------------
        
    Set rsN = New ADODB.Recordset
        rsN.Open "GL_Tran", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!AC_No = rs!AC_No
        rsN!Name = rsU!Head_Name
        rsN!Description = "Income Realized"
        rsN!Dr = 0
        rsN!Cr = rsU!Balance
        rsN!Balance = rs!Balance
        rsN.Update
        rsN.Close
        
     rsU!Balance = rsU!Balance - rsU!Balance
        rsU!Date = txtDate.Text
    
'---------------------------------------------
'---------------------------------------------
        rsU.MoveNext
 Loop
    rsU.Update
    rsU.Close
Timer1.Enabled = True
'------------------------------------------------------------------------------------------
    Else
        Exit Sub
    End If
Else
    Exit Sub
End If
   
   Resume Next
End Sub
     
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        
        rs.MoveFirst
        If Not rs.EOF Then
        txtDate.Text = rs!Today
        txtLast.Text = rs!Month_Day
        rs.Close
        End If
End Sub
Private Sub Timer1_Timer()
ProgressBar1.Visible = True
    ProgressBar1.Max = 100
    ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = ProgressBar1.Max Then
    ProgressBar1.Value = 0
    Timer1.Enabled = False
    ProgressBar1.Visible = False
    MsgBox "Week close processing successfully completed!", vbInformation + vbOKOnly, "Day Close"
End If
End Sub

