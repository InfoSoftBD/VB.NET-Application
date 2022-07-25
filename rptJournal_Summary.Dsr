VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptJournal_Summary 
   Caption         =   "Journal Summary"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19288
   SectionData     =   "rptJournal_Summary.dsx":0000
End
Attribute VB_Name = "rptJournal_Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
On Error Resume Next
Dim rs As ADODB.Recordset
Dim str As String

    Set rs = New ADODB.Recordset
        str = "select * from GL_Tran where cdate(Date) Between cdate('" & lblFrom.Caption & "') And cdate('" & lblTo.Caption & "') and Name like '" & Link & "' order by Sl"
        rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
        rs.MoveFirst

    rptGL_Statement.rsGL.ConnectionString = cnStr
    rptGL_Statement.rsGL.Source = str
    rptGL_Statement.Show 1
    
End Sub
Private Sub GroupFooter1_Format()
Field13.Hyperlink = Field13.Text

Field7.Text = Val(Field9.DataValue) - Val(Field8.DataValue)
Field7.Text = Format(Field7.Text, "#,##0.00")

Static iRow As Integer
    If iRow Mod 2 = 0 Then
       GroupFooter1.BackStyle = ddBKTransparent
        GroupFooter1.BackColor = vbWhite
    Else
        GroupFooter1.BackStyle = ddBKNormal
        GroupFooter1.BackColor = &HE0E0E0
    End If
    iRow = iRow + 1

End Sub

Private Sub PageFooter_BeforePrint()
Me.Label47.Caption = Power
End Sub

Private Sub PageHeader_BeforePrint()
Me.Label1.Caption = UCase(Org_Name)
Me.lblBranch.Caption = UCase(Branch_Name & " Branch, " & Branch_Address)
Me.lblUser_Id.Caption = User_Id
Me.lblUser_Name.Caption = User_Name
End Sub

