VERSION 5.00
Begin VB.Form frmBranch 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Branch Entry"
   ClientHeight    =   4935
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   6705
   Icon            =   "frmBranch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   150
      ScaleHeight     =   2595
      ScaleWidth      =   6360
      TabIndex        =   7
      Top             =   1035
      Width           =   6390
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Branch Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   120
         TabIndex        =   8
         Top             =   165
         Width           =   6060
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   240
            Left            =   4320
            TabIndex        =   17
            Top             =   1035
            Width           =   240
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
            Height          =   435
            Left            =   4140
            TabIndex        =   12
            Top             =   345
            Width           =   1695
         End
         Begin VB.TextBox txtId 
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
            Height          =   435
            Left            =   1710
            TabIndex        =   11
            Top             =   330
            Width           =   1695
         End
         Begin VB.TextBox txtName 
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
            Height          =   435
            Left            =   1710
            TabIndex        =   10
            Top             =   945
            Width           =   2385
         End
         Begin VB.TextBox txtAddress 
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
            Height          =   435
            Left            =   1710
            TabIndex        =   9
            Top             =   1530
            Width           =   4140
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Host Branch"
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
            Left            =   4725
            TabIndex        =   18
            Top             =   1035
            Width           =   1080
         End
         Begin VB.Label Label5 
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
            Left            =   855
            TabIndex        =   16
            Top             =   1635
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Name"
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
            Left            =   390
            TabIndex        =   15
            Top             =   1005
            Width           =   1185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Date"
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
            Left            =   3555
            TabIndex        =   14
            Top             =   420
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Code"
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
            Left            =   450
            TabIndex        =   13
            Top             =   435
            Width           =   1125
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   135
      ScaleHeight     =   840
      ScaleWidth      =   6405
      TabIndex        =   2
      Top             =   3840
      Width           =   6435
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Save"
         Height          =   435
         Left            =   120
         TabIndex        =   6
         Top             =   200
         Width           =   1155
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   435
         Left            =   1785
         TabIndex        =   5
         Top             =   200
         Width           =   1155
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   435
         Left            =   3405
         TabIndex        =   4
         Top             =   200
         Width           =   1155
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   435
         Left            =   4995
         TabIndex        =   3
         Top             =   200
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   135
      Picture         =   "frmBranch.frx":0442
      ScaleHeight     =   615
      ScaleWidth      =   6405
      TabIndex        =   0
      Top             =   180
      Width           =   6435
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BRANCH ENTRY FORM"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1410
         TabIndex        =   1
         Top             =   120
         Width           =   3780
      End
      Begin VB.Image Image3 
         Height          =   645
         Left            =   0
         Picture         =   "frmBranch.frx":5CF0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6990
      End
   End
End
Attribute VB_Name = "frmBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Today As Date
Dim str As String
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Private Sub clearTextboxes()
        txtID.Text = ""
        txtDate.Text = ""
        txtName.Text = ""
        txtAddress.Text = ""
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdReceive_Click()
If Check1.Value = 1 Then

    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        
        rsU!Branch_Code = txtID.Text
        rsU!Branch_Name = txtName.Text
        rsU!Branch_Address = txtAddress.Text
        rsU.Update
        rsU.Close
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Branch", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!Branch_Code = txtID.Text
        rsN!Branch_Name = txtName.Text
        rsN!Branch_Address = txtAddress.Text
        
        rsN.Update
        rsN.Close
Else
        
    Set rsN = New ADODB.Recordset
        rsN.Open "Branch", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!Branch_Code = txtID.Text
        rsN!Branch_Name = txtName.Text
        rsN!Branch_Address = txtAddress.Text
        
        rsN.Update
        rsN.Close
End If
Call clearTextboxes
txtDate.Text = Date
txtID.SetFocus
End Sub

Private Sub cmdUpdate_Click()
If Check1.Value = 1 Then

    Set rsU = New ADODB.Recordset
        str = "select * from Others"
        rsU.Open str, conn, adOpenDynamic, adLockOptimistic
        rsU.MoveFirst
        
        rsU!Branch_Code = txtID.Text
        rsU!Branch_Name = txtName.Text
        rsU!Branch_Address = txtAddress.Text
        rsU.Update
        rsU.Close
        
    Set rsN = New ADODB.Recordset
        str = "select * from Branch where Branch_Code like '" & txtID.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
        If Not rsN.EOF Then
        
            rsN!Date = txtDate.Text
            rsN!Branch_Code = txtID.Text
            rsN!Branch_Name = txtName.Text
            rsN!Branch_Address = txtAddress.Text
            rsN.Update
            rsN.Close
        Else
        rsN.Close
        MsgBox "Invalid Branch Code!"
        txtID.SetFocus
        End If
Else
        
    Set rsN = New ADODB.Recordset
        str = "select * from Branch where Branch_Code like '" & txtID.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
        If Not rsN.EOF Then
        
            rsN!Date = txtDate.Text
            rsN!Branch_Code = txtID.Text
            rsN!Branch_Name = txtName.Text
            rsN!Branch_Address = txtAddress.Text
            rsN.Update
            rsN.Close
        Else
        rsN.Close
        MsgBox "Invalid Branch Code!"
        txtID.SetFocus
        End If
End If
Call clearTextboxes
txtDate.Text = Date
txtID.SetFocus
cmdUpdate.Enabled = False
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call clearTextboxes
    Set rs = New ADODB.Recordset
        str = "select * from Others"
        rs.Open str, conn
        rs.MoveFirst
        If Not rs.EOF Then
           On Error Resume Next
           Today = rs!Today
           rs.Close
        End If
    txtDate.Text = Today
    cmdReceive.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdReceive.SetFocus
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtName.SetFocus
End If
End Sub
Private Sub txtId_LostFocus()
Dim mid As Integer
    Search_Id = txtID.Text
    mid = 0
    
    Set rs = New ADODB.Recordset
        str = "select * from Branch where Branch_Code like '" & Search_Id & "' order by Branch_Code"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        txtID.Text = rs!Branch_Code
        txtDate.Text = rs!Date
        txtName.Text = rs!Branch_Name
        txtAddress.Text = rs!Branch_Address
        rs.Close
               
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = True
        'cmdPrint.Enabled = True
    Else
        If MsgBox("Do you want add new Branch?", vbInformation + vbYesNo, "Add New") = vbYes Then
            
            Set rsU = New ADODB.Recordset
            str = "select * from Branch order by Branch_Code"
            rsU.Open str, conn
        
                If Not rsU.EOF Then
                    
                    rsU.MoveFirst
                    
                    Do While Not rsU.EOF = True
                        mid = Val(rsU!Branch_Code)
                        rsU.MoveNext
                    Loop
                        rsU.Close
                        mid = mid + 1
                Else
                        mid = "1001"
                End If
                
            Call clearTextboxes
            txtDate.Text = Today
            txtID.Text = Format$(Val(mid), "000#")
            cmdReceive.Enabled = True
        Else
            Call clearTextboxes
            txtName.SetFocus
        End If
    End If
    Exit Sub
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAddress.SetFocus
End If
End Sub



