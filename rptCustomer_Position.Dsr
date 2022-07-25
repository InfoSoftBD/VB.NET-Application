VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCustomer_Position 
   Caption         =   "ActiveReport1"
   ClientHeight    =   12915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   22781
   SectionData     =   "rptCustomer_Position.dsx":0000
End
Attribute VB_Name = "rptCustomer_Position"
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
        Detail.BackStyle = ddBKTransparent
        Detail.BackColor = vbWhite
    End If
    iRow = iRow + 1
    lblRank.Caption = str(iRow)
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

