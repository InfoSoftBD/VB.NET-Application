VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDayclose 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Day Close"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   Icon            =   "frmDayclose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   135
      ScaleHeight     =   315
      ScaleWidth      =   5085
      TabIndex        =   8
      Top             =   2745
      Width           =   5145
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   0
         TabIndex        =   9
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   45
      Top             =   45
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   135
      ScaleHeight     =   1515
      ScaleWidth      =   5115
      TabIndex        =   5
      Top             =   1035
      Width           =   5145
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
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   135
         Width           =   2400
      End
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
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   810
         Width           =   2400
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
         TabIndex        =   10
         Top             =   225
         Width           =   2085
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
         TabIndex        =   7
         Top             =   900
         Width           =   1755
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   135
      ScaleHeight     =   705
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   3195
      Width           =   5145
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   510
         Left            =   3150
         TabIndex        =   4
         Top             =   90
         Width           =   1545
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start Processing"
         Height          =   510
         Left            =   360
         TabIndex        =   3
         Top             =   90
         Width           =   1545
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   135
      ScaleHeight     =   705
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   135
      Width           =   5145
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DAY CLOSE PROCESS"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   405
         TabIndex        =   1
         Top             =   135
         Width           =   4320
      End
      Begin VB.Image Image3 
         Height          =   735
         Left            =   0
         Picture         =   "frmDayclose.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5205
      End
   End
End
Attribute VB_Name = "frmDayclose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Today As Date
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub Command1_Click()
On Error Resume Next
Dim Day As Date
Day = DateAdd("D", 1, CDate(Today))

If MsgBox("Are you sure you want start Day close process?", vbInformation + vbYesNo, "Week Close") = vbYes Then
If MsgBox("Have you taken Backup?", vbInformation + vbYesNo, "Day Close") = vbYes Then
    
    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.MoveFirst
    
        On Error Resume Next
        txtLast.Text = Day
        rs!Today = Day
        rs!Cash_Open = rs!Cash_Close
        rs!Cash_Dr = 0
        rs!Cash_Cr = 0
        rs!Cash_Close = rs!Cash_Close
        
        rs!Bank_Open = rs!Bank_Close
        rs!Bank_Dr = 0
        rs!Bank_Cr = 0
        rs!Bank_Close = rs!Bank_Close
        rs.Update
        rs.Close
    
        
   Set rsU = New ADODB.Recordset
        str = "select * from Cash_Master Order by Sl"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst

    Do While Not rsU.EOF
        rsU!Date = Day
        rsU!Open_Bal = 0
        rsU!Cr = 0
        rsU!Dr = 0
        rsU!Balance = 0
        
        rsU!Cash_Open = 0
        rsU!Cash_Dr = 0
        rsU!Cash_Cr = 0
        rsU!Cash_Close = 0
        rsU.MoveNext
    Loop
    rsU.Update
    rsU.Close
         
    Set rsU = New ADODB.Recordset
        str = "select * from Bank_Master Order by Id"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
    
    Do While Not rsU.EOF
        rsU!Open_Bal = rsU!Balance
        rsU!Deposit = 0
        rsU!Withdraw = 0
        rsU!Balance = rsU!Balance
        rsU.MoveNext
    Loop
    rsU.Update
    rsU.Close
        
'Customer Close------------------------------------------------------------------

    Set rsU = New ADODB.Recordset
            str = "select * from Customer_Master Order by Customer_Code"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
            
    Do While Not rsU.EOF
    
        rsU!Open_Bal = rsU!Balance
        rsU!Dr = 0
        rsU!Cr = 0
        rsU!Balance = rsU!Balance
        rsU.MoveNext
    Loop
    
    rsU.Update
    rsU.Close
'--------------------------------------------------------------
     Set rsU = New ADODB.Recordset
            str = "select * from Prod_Master Order by Sl_No"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
    
    Do While Not rsU.EOF
        rsU!Open_Bal = rsU!Stock
        rsU!Purchase = 0
        rsU!Sale = 0
        rsU!Lift = 0
        rsU!Return = 0
        rsU!Stock = rsU!Stock
        rsU.MoveNext
    Loop
        rsU.Update
        rsU.Close


    Set rsU = New ADODB.Recordset
            str = "select * from Staff_Master Order by Id"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
    
    Do While Not rsU.EOF
        rsU!Salary = 0
        rsU!Commission = 0
        rsU!Mobile = 0
        rsU!Balance = (rsU!Salary + rsU!Commission + rsU!Mobile) - (rsU!Advance + rsU!Sales_Due + rsU!Draw)
        
        rsU.MoveNext
    Loop
        rsU.Update
        rsU.Close

    Set rsU = New ADODB.Recordset
            str = "select * from Fixed_Master Order by Sl_No"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
    
    Do While Not rsU.EOF
        rsU!Open_Bal = rsU!Stock
        rsU!Purchase = 0
        rsU!Disposal = 0
        rsU!Stock = rsU!Stock
        rsU.MoveNext
    Loop
    rsU.Update
    rsU.Close

    Set rsU = New ADODB.Recordset
            str = "select * from Vendor_Master Order by Id"
            rsU.Open str, conn, adOpenDynamic, adLockOptimistic
            rsU.MoveFirst
    
    Do While Not rsU.EOF
        rsU!Open_Bal = rsU!Balance
        rsU!Dr = 0
        rsU!Cr = 0
        rsU!Balance = rsU!Balance
        rsU.MoveNext
    Loop
    rsU.Update
    rsU.Close

    Set rsU = New ADODB.Recordset
        str = "select * from Branch Order by Id"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst

    Do While Not rsU.EOF
        rsU!Open_Bal = rsU!Balance
        rsU!Dr = 0
        rsU!Cr = 0
        rsU!Balance = rsU!Balance
        rsU.MoveNext
    Loop
    rsU.Update
    rsU.Close
    Today = Day
    Timer1.Enabled = True
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
        Today = rs!Today
        txtDate.Text = Today
        txtLast.Text = DateAdd("D", -1, CDate(Today))
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
