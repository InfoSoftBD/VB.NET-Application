VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptLoan_Info 
   Caption         =   "Loan Account Opening"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   26194
   SectionData     =   "rptLoan_Info.dsx":0000
End
Attribute VB_Name = "rptLoan_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
On Error Resume Next
Image1.Picture = LoadPicture(App.Path + "\Photo\" & Photo_Source.Text)
Image2.Picture = LoadPicture(App.Path + "\Thumb\" & Thumb_Source.Text)
Me.lblUser_Id.Caption = User_Id
Me.lblUser_Name.Caption = User_Name
End Sub
Private Sub PageHeader_BeforePrint()
Me.lblBranch.Caption = UCase(Branch_Name & " Branch, " & Branch_Address)
End Sub

