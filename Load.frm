VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Load 
   BackColor       =   &H00008000&
   Caption         =   "Loading......"
   ClientHeight    =   735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      Height          =   345
      Left            =   135
      ScaleHeight     =   285
      ScaleWidth      =   5130
      TabIndex        =   0
      Top             =   180
      Width           =   5190
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   300
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Activate()
Dim progMax As Integer
Set rs = New ADODB.Recordset
    str = "select * from GL_Master"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly

    rs.MoveFirst
    Do While Not rs.EOF
    progMax = progMax + 1
    rs.MoveNext
    Loop
    rs.Close

ProgressBar1.Visible = True
ProgressBar1.Min = 0
ProgressBar1.Max = progMax
ProgressBar1.Value = 0

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
        Unload Me
        rptTrial_Balance.Show 1
    End If
End Sub

