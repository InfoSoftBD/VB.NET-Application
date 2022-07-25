VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptTrialbalance 
   Caption         =   "Trial Balance"
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptTralbalance.dsx":0000
End
Attribute VB_Name = "rptTrialbalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
MsgBox Link
End Sub
Private Sub Detail_Format()
Field5.Hyperlink = Field5.Text
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
