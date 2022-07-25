VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptMoney_Receipt 
   Caption         =   "Money Receipt"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptMoney_Receipt.dsx":0000
End
Attribute VB_Name = "rptMoney_Receipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GroupHeader1_Format()
Me.lblUser_Id.Caption = User_Id
Me.lblUser_Name.Caption = User_Name
End Sub

Private Sub PageFooter_BeforePrint()
Me.Label47.Caption = Power
End Sub

Private Sub PageHeader_BeforePrint()
On Error Resume Next
If Language = "Bangla" Then
       
    Me.Label1.Caption = Org_Name
    Me.Label1.Font.Name = "SutonnyMJ"
    Me.Label1.Font.Size = 20
    Me.Label1.Font.Bold = True
   ' Me.Label2.Caption = Reg_No
    'Me.Label2.Font.Name = "SutonnyMJ"
   ' Me.Label2.Font.Size = 12
    Me.lblBranch.Caption = Branch_Address
    Me.lblBranch.Font.Name = "SutonnyMJ"
    Me.lblBranch.Font.Size = 12
    'Me.lblUser_Id.Caption = User_Id
    'Me.lblUser_Name.Caption = User_Name
Else

    Me.Label1.Caption = UCase(Org_Name)
    Me.Label1.Font.Name = "Cooper Black"
    Me.Label1.Font.Size = 16
    Me.Label1.Font.Bold = True
    'Me.Label2.Caption = Reg_No
    'Me.Label2.Font.Name = "Arial"
    'Me.Label2.Font.Size = 10
    Me.lblBranch.Caption = UCase(Branch_Name & ", " & Branch_Address)
    Me.lblBranch.Font.Name = "Arial"
    Me.lblBranch.Font.Size = 10
'    Me.lblUser_Id.Caption = User_Id
'    Me.lblUser_Name.Caption = User_Name

End If
'Me.lblUser_Id.Caption = User_Id
'Me.lblUser_Name.Caption = User_Name

End Sub

