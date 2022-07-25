VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptServ_Report 
   Caption         =   "Daily Clinical Service Report"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   19288
   SectionData     =   "rptServ_Report.dsx":0000
End
Attribute VB_Name = "rptServ_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
frmSales.txtInvoice.Text = Link
frmSales.Show 1
End Sub

Private Sub Detail_Format()
Static iRow As Integer
    If iRow Mod 2 = 0 Then
       Detail.BackStyle = ddBKTransparent
        Detail.BackColor = vbWhite
    Else
        Detail.BackStyle = ddBKTransparent
        Detail.BackColor = vbWhite
    End If
    iRow = iRow + 1
    lblRank.Caption = str(iRow)
    
    If Val(Bill.Text) = 0 Then
        Detail.BackStyle = ddBKNormal
        Detail.BackColor = &HC0C0FF
    End If
Field16.Hyperlink = Field16.Text
   Me.NBalance.Text = Format$(Val(Bill.Text) - (Val(Discount.Text) + Val(Payment.Text)), "###0.00")
End Sub

Private Sub PageFooter_BeforePrint()
Me.Label47.Caption = Power
End Sub

Private Sub PageHeader_BeforePrint()
Me.Label1.Caption = UCase(Org_Name)
Me.lblBranch.Caption = (Branch_Address)
Image1.Picture = LoadPicture(App.Path + "\Photo\Mono.jpg")
Me.lblUser_Id.Caption = User_Id
Me.lblUser_Name.Caption = User_Name
End Sub


Private Sub ReportFooter_Format()
Nbalance_Total.Text = Format$(Val(Bill_Total.Text) - Val(Disc_Total.Text) - Val(Payment_Total.Text), "###0.00")
End Sub
