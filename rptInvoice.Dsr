VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptInvoice 
   Caption         =   "Sales Invoice"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   19288
   SectionData     =   "rptInvoice.dsx":0000
End
Attribute VB_Name = "rptInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub GroupFooter1_Format()
Static iRow As Integer
    If iRow Mod 2 = 0 Then
       GroupFooter1.BackStyle = ddBKTransparent
        GroupFooter1.BackColor = vbWhite
    Else
        GroupFooter1.BackStyle = ddBKTransparent
        GroupFooter1.BackColor = vbWhite
    End If
    iRow = iRow + 1
    lblRank.Caption = str(iRow)

    
    'Field15.Text = (Val(Field19.DataValue) * Val(Field16.DataValue)) - Val(Field28.DataValue)
    'Field15.Text = Format(Field15.Text, "#,##0.00")

End Sub

Private Sub PageFooter_BeforePrint()
Me.Label47.Caption = Power
End Sub

Private Sub PageHeader_BeforePrint()
Me.Label1.Caption = UCase(Org_Name)
Me.lblBranch.Caption = (Branch_Address)
Image1.Picture = LoadPicture(App.Path + "\Photo\Mono.jpg")
'Me.lblUser_Id.Caption = User_Id
'Me.lblUser_Name.Caption = User_Name
End Sub

Private Sub ReportFooter_Format()
Me.Label82.Caption = UCase(Org_Name)
End Sub
