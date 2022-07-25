VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmW_Close 
   BackColor       =   &H00008000&
   Caption         =   "Week Close"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   Icon            =   "frmW_Close.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   135
      ScaleHeight     =   675
      ScaleWidth      =   5085
      TabIndex        =   10
      Top             =   135
      Width           =   5145
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WEEK CLOSE PROCESS"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   135
         TabIndex        =   11
         Top             =   45
         Width           =   4635
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   135
      ScaleHeight     =   675
      ScaleWidth      =   5085
      TabIndex        =   7
      Top             =   3195
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
      BackColor       =   &H00C0FFC0&
      Height          =   1545
      Left            =   135
      ScaleHeight     =   1485
      ScaleWidth      =   5085
      TabIndex        =   2
      Top             =   1035
      Width           =   5145
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
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
      Left            =   45
      Top             =   45
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000003&
      Height          =   375
      Left            =   135
      ScaleHeight     =   315
      ScaleWidth      =   5085
      TabIndex        =   0
      Top             =   2745
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
Attribute VB_Name = "frmW_Close"
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
Dim Day As Date
  
Day = CDate(txtDate.Text)

If MsgBox("Are you sure you want start Day close process?", vbInformation + vbYesNo, "Week Close") = vbYes Then
If MsgBox("Have you taken Backup?", vbInformation + vbYesNo, "Day Close") = vbYes Then

    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.MoveFirst
        If Not rs.EOF Then
        On Error Resume Next
        txtLast.Text = Day
        rs!Week_Day = Day
        rs.Update
        rs.Close
'-----------------------------------------------------------------------
Set rsU = New ADODB.Recordset
    str = "select * from Loan_Master Order by Sl_No"
    rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    rsU.MoveFirst

Do While Not rsU.EOF
    rsU!Weekly_Bal = rsU!Balance
    rsU!Weekly_Pay = 0
    rsU!Weekly_Draw = 0
    rsU!Week_Close = CDate(txtDate.Text)
    rsU.MoveNext
 Loop
    rsU.Update
    rsU.Close
'---------------------------------------------------------------------------
    
Set rsU = New ADODB.Recordset
        str = "select * from Deposit_Master Order by Sl_No"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        
Do While Not rsU.EOF
    rsU!Weekly_Bal = rsU!Amount
    rsU!Weekly_Pay = 0
    rsU!Weekly_Draw = 0
    rsU!Week_Close = CDate(txtDate.Text)
    rsU.MoveNext
 Loop
    rsU.Update
    rsU.Close
        
    
'--------------------------------------------------------------
 
    Timer1.Enabled = True
End If

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
        txtLast.Text = rs!Week_Day
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

