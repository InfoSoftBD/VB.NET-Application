VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptLoanMonthly 
   Caption         =   "ActiveReport1"
   ClientHeight    =   14820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18960
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   33443
   _ExtentY        =   26141
   SectionData     =   "rptLoanMonthly.dsx":0000
End
Attribute VB_Name = "rptLoanMonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PageHeader_BeforePrint()
Me.lblBranch.Caption = UCase(Branch_Name & " Branch, " & Branch_Address)
Me.lblUser_Id.Caption = User_Id
Me.lblUser_Name.Caption = User_Name
End Sub
