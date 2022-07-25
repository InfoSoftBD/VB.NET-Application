VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptFact_Cat 
   Caption         =   "Factory Product Category"
   ClientHeight    =   12915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   22781
   SectionData     =   "rptFact_Cat.dsx":0000
End
Attribute VB_Name = "rptFact_Cat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
'    Dim rs As ADODB.Recordset
'    Dim str As String
'    Dim f_code As String
'    f_code = Me.Field56.Text
'
'   ' Fromdate = lblFrom.Caption
'   ' Today = lblTo.Caption
'
'     Set rs = New ADODB.Recordset
'        str = "select * from Factory_Master where Prod_Type like '" & Link & "' And Vendor_Code like '" & f_code & "' Order by Prod_Name"
'        rs.Open str, conn
'
'    If Not rs.EOF Then
'        rptFact_Prod.rsStock.ConnectionString = cnStr
'        rptFact_Prod.rsStock.Source = str
'        rs.Close
'        rptFact_Prod.Show 1
'    Else
'        MsgBox "There is no such product found, ", vbCritical + vbOKOnly
'        rs.Close
'    End If
    
End Sub

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
    lblRank.Caption = str(iRow)
'    Field54.Hyperlink = Field54.Text
End Sub

Private Sub PageFooter_BeforePrint()
Me.Label47.Caption = Power
End Sub

Private Sub PageHeader_BeforePrint()
Image1.Picture = LoadPicture(App.Path + "\Photo\Mono.jpg")
Me.lblOrg_Name.Caption = UCase(Org_Name)
Me.lblAddress.Caption = Branch_Address
Me.lblUser_Id.Caption = User_Id
Me.lblUser_Name.Caption = User_Name
End Sub



