VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptStatement 
   Caption         =   "Customer Account Statement"
   ClientHeight    =   11010
   ClientLeft      =   840
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptStatement.dsx":0000
End
Attribute VB_Name = "rptStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Customer_BeforePrint()
Me.lblUser_Id.Caption = User_Id
Me.lblUser_Name.Caption = User_Name
End Sub


Private Sub PageFooter_BeforePrint()
Me.Label47.Caption = Power
End Sub

Private Sub PageHeader_BeforePrint()
Me.Label1.Caption = UCase(Org_Name)
Me.lblBranch.Caption = (Branch_Address)
Image1.Picture = LoadPicture(App.Path + "\Photo\Mono.jpg")
End Sub
