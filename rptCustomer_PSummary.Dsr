VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCustomer_PSummary 
   Caption         =   "Product Wise Customer Summary"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   26194
   SectionData     =   "rptCustomer_PSummary.dsx":0000
End
Attribute VB_Name = "rptCustomer_PSummary"
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
        str = "select * from Customer_Master where Prod_Type like '" & Link & "'"
        rs.Open str, conn
        
    If Not rs.EOF Then
        rptCustomer_Position.rsTran.ConnectionString = cnStr
        rptCustomer_Position.rsTran.Source = str
        rs.Close
        rptCustomer_Position.Show 1
    Else
        MsgBox "There is no such product found, ", vbCritical + vbOKOnly
        rs.Close
    End If
   
End Sub

Private Sub GroupFooter1_Format()
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
    Field5.Hyperlink = Field5.Text
End Sub

Private Sub PageHeader_Format()
Me.Label1.Caption = UCase(Org_Name)
Me.lblBranch.Caption = (Branch_Address)
Image1.Picture = LoadPicture(App.Path + "\Photo\Mono.jpg")
Me.lblUser_Id.Caption = User_Id
Me.lblUser_Name.Caption = User_Name
End Sub

