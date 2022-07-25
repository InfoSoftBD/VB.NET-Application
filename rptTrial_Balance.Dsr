VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptTrial_Balance 
   Caption         =   "Trial Balance"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19288
   SectionData     =   "rptTrial_Balance.dsx":0000
End
Attribute VB_Name = "rptTrial_Balance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
Static iRow As Integer
    If iRow Mod 2 = 0 Then
       Detail.BackStyle = ddBKTransparent
        Detail.BackColor = vbWhite
    Else
        Detail.BackStyle = ddBKNormal
        Detail.BackColor = &HE0E0E0
    End If
    iRow = iRow + 1
    'lblRank.Caption = str(iRow)

End Sub

Private Sub GroupFooter1_Format()
Field29.Text = Val(Field28.DataValue) - Val(Field27.DataValue)
Field29.Text = Format(Field29.Text, "#,##0.00")
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
