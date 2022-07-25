VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptDaily_Income 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptDaily_Income.dsx":0000
End
Attribute VB_Name = "rptDaily_Income"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PageFooter_BeforePrint()
Me.Label47.Caption = Power
End Sub

Private Sub GroupFooter1_Format()
Field19.Hyperlink = Field19.Text
Field16.Hyperlink = Field16.Text
Static iRow As Integer
    If iRow Mod 2 = 0 Then
       GroupFooter1.BackStyle = ddBKTransparent
        GroupFooter1.BackColor = vbWhite
    Else
        GroupFooter1.BackStyle = ddBKNormal
        GroupFooter1.BackColor = &HE0E0E0
    End If
    iRow = iRow + 1
    lblRank.Caption = str(iRow)
    lblRank1.Caption = str(iRow)

End Sub

Private Sub PageHeader_Format()
Me.Label1.Caption = UCase(Org_Name)
Me.lblBranch.Caption = (Branch_Address)
Image1.Picture = LoadPicture(App.Path + "\Photo\Mono.jpg")
End Sub

Private Sub ReportFooter_Format()
Field29.Text = Val(Field27.DataValue) - Val(Field28.DataValue)
Field29.Text = Format(Field29.Text, "#,##0.00")
End Sub
