VERSION 5.00
Begin VB.Form frmSamity 
   BackColor       =   &H00008000&
   Caption         =   "Samity Entry"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6705
   Icon            =   "frmSamity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0FFC0&
      Height          =   645
      Left            =   90
      ScaleHeight     =   585
      ScaleWidth      =   6375
      TabIndex        =   15
      Top             =   180
      Width           =   6435
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SAMITY ENTRY FORM"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   105
         TabIndex        =   16
         Top             =   0
         Width           =   5670
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   870
      Left            =   90
      ScaleHeight     =   810
      ScaleWidth      =   6375
      TabIndex        =   10
      Top             =   4200
      Width           =   6435
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   435
         Left            =   4995
         TabIndex        =   14
         Top             =   200
         Width           =   1155
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   435
         Left            =   3405
         TabIndex        =   13
         Top             =   200
         Width           =   1155
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   435
         Left            =   1785
         TabIndex        =   12
         Top             =   200
         Width           =   1155
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Add New"
         Height          =   435
         Left            =   255
         TabIndex        =   11
         Top             =   200
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   3030
      Left            =   105
      ScaleHeight     =   2970
      ScaleWidth      =   6330
      TabIndex        =   0
      Top             =   1020
      Width           =   6390
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Samity Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6060
         Begin VB.TextBox txtMember_no 
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
            TabIndex        =   18
            Top             =   2070
            Width           =   1695
         End
         Begin VB.TextBox txtAddress 
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
            TabIndex        =   17
            Top             =   1440
            Width           =   4140
         End
         Begin VB.TextBox txtName 
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
            TabIndex        =   4
            Top             =   840
            Width           =   4140
         End
         Begin VB.TextBox txtId 
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
            TabIndex        =   3
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtDate 
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
            TabIndex        =   2
            Top             =   255
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Samity Code"
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
            TabIndex        =   9
            Top             =   345
            Width           =   1125
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
            TabIndex        =   8
            Top             =   330
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Samity Name"
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
            Left            =   330
            TabIndex        =   7
            Top             =   915
            Width           =   1185
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
            Left            =   795
            TabIndex        =   6
            Top             =   1545
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Member"
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
            Left            =   240
            TabIndex        =   5
            Top             =   2160
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frmSamity"
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
        txtId.Text = ""
        txtDate.Text = ""
        txtName.Text = ""
        txtAddress.Text = ""
        txtMember_no.Text = ""
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdReceive_Click()
Set rsN = New ADODB.Recordset
        rsN.Open "Samity", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!Date = txtDate.Text
        rsN!Samity_Code = txtId.Text
        rsN!Samity_Name = txtName.Text
        rsN!Samity_Address = txtAddress.Text
        rsN!Member_no = txtMember_no.Text
        rsN.Update
        rsN.Close
Call clearTextboxes
txtDate.Text = Today
txtId.SetFocus
End Sub

Private Sub cmdUpdate_Click()
Set rsN = New ADODB.Recordset
        str = "select * from Samity where Samity_Code like '" & txtId.Text & "'"
        rsN.Open str, conn, adOpenDynamic, adLockOptimistic
        
    If Not rsN.EOF Then
        rsN!Date = txtDate.Text
        rsN!Samity_Code = txtId.Text
        rsN!Samity_Name = txtName.Text
        rsN!Samity_Address = txtAddress.Text
        rsN!Member_no = txtMember_no.Text
        rsN.Update
        rsN.Close
   Call clearTextboxes
   txtDate.Text = Today
        Else
        rsN.Close
        MsgBox "Invalid Samity Code!"
        txtId.SetFocus
        End If
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
    
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMember_no.SetFocus
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtName.SetFocus
End If
End Sub
Private Sub txtId_LostFocus()
Dim mid As Integer
    Search_Id = txtId.Text
    mid = 0
    
    Set rs = New ADODB.Recordset
        str = "select * from Samity where Samity_Code like '" & Search_Id & "' order by Samity_Code"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        txtId.Text = rs!Samity_Code
        txtDate.Text = rs!Date
        txtName.Text = rs!Samity_Name
        txtAddress.Text = rs!Samity_Address
        txtMember_no.Text = rs!Member_no
        rs.Close
               
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = True
        'cmdPrint.Enabled = True
    Else
        If MsgBox("Do you want add new Samity?", vbInformation + vbYesNo, "Add New") = vbYes Then
            
            Set rsU = New ADODB.Recordset
            str = "select * from Samity order by Samity_Code"
            rsU.Open str, conn
        
                If Not rsU.EOF Then
                    
                    rsU.MoveFirst
                    
                    Do While Not rsU.EOF = True
                        mid = Val(rsU!Samity_Code)
                        rsU.MoveNext
                    Loop
                        rsU.Close
                        mid = mid + 1
                Else
                        mid = "2001"
                End If
                
            Call clearTextboxes
            txtDate.Text = Today
            txtId.Text = Format$(Val(mid), "000#")
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


