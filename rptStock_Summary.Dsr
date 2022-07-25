VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptStock_Summary 
   Caption         =   "Summary Stock of Inventory"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   26194
   SectionData     =   "rptStock_Summary.dsx":0000
End
Attribute VB_Name = "rptStock_Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
    Dim rs As ADODB.Recordset
    Dim str As String

   ' Fromdate = lblFrom.Caption
   ' Today = lblTo.Caption
    
     Set rs = New ADODB.Recordset
        str = "select * from Prod_Master where Prod_Name like '" & Link & "'"
        rs.Open str, conn
        
    If Not rs.EOF Then
        rptStock.rsStock.ConnectionString = cnStr
        rptStock.rsStock.Source = str
        rs.Close
        rptStock.Show 1
    Else
        MsgBox "There is no such product found, ", vbCritical + vbOKOnly
        rs.Close
    End If
    
End Sub
Private Sub GroupFooter1_Format()
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
    Field54.Hyperlink = Field54.Text

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


