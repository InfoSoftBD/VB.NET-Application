VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptProd_Summary 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19288
   SectionData     =   "rptProd_Summary.dsx":0000
End
Attribute VB_Name = "rptProd_Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim amnt As Double
Dim profit As Double
Private Sub PageHeader_BeforePrint()
Me.Label1.Caption = UCase(Org_Name)
Me.lblBranch.Caption = (Branch_Address)
Image1.Picture = LoadPicture(App.Path + "\Photo\Mono.jpg")
Me.lblUser_Id.Caption = User_Id
Me.lblUser_Name.Caption = User_Name
End Sub
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
    Dim rs As ADODB.Recordset
    Dim str As String

    Fromdate = lblFrom.Caption
    Today = lblTo.Caption
    
     Set rs = New ADODB.Recordset
        str = "select * from Prod_Tran where cdate(D_ate) Between cdate('" & Fromdate & "') And cdate('" & Today & "')and Prod_Code like '" & Link & "' order by Sl"
        rs.Open str, conn
        
    If Not rs.EOF Then
        rptProd_Statement.rsProd_Statement.ConnectionString = cnStr
        rptProd_Statement.rsProd_Statement.Source = str
        rs.Close
        rptProd_Statement.Show 1
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



Field32.Hyperlink = Field32.Text

If Field32.Text = "101" Then
    Field30.Text = 0
    Field30.Text = Format(Field30.Text, "#,##0.00")
    
    Field36.Text = 0
    Field36.Text = Format(Field36.Text, "#,##0.00")
    
    Field37.Text = Val(Field31.DataValue) * 0.0127
    Field37.Text = Format(Field37.Text, "#,##0.00")
    
    Field27.Text = Val(Field31.DataValue)
    Field27.Text = Format(Field27.Text, "#,##0.00")
    
    profit = profit + Field37.Text
Else
    If Field32.Text = "BKS01" Then
        Field30.Text = 0
        Field30.Text = Format(Field30.Text, "#,##0.00")
        
        Field36.Text = 0
        Field36.Text = Format(Field36.Text, "#,##0.00")
        
        Field37.Text = 0
        Field37.Text = Format(Field37.Text, "#,##0.00")
        
        Field27.Text = Val(Field31.DataValue)
        Field27.Text = Format(Field27.Text, "#,##0.00")
        
        profit = profit + Field37.Text
    Else

            Field37.Text = (Val(Field31.DataValue) * Val(Field36.DataValue)) - (Val(Field31.DataValue) * Val(Field30.DataValue))
            Field37.Text = Format(Field37.Text, "#,##0.00")
            
            Field27.Text = Val(Field31.DataValue) * Val(Field36.DataValue)
            Field27.Text = Format(Field27.Text, "#,##0.00")
            
            amnt = amnt + Field27.Text
            profit = profit + Field37.Text
    End If
    End If
End Sub

Private Sub ReportFooter_Format()
Field38.Text = profit
Field38.Text = Format(Field38.Text, "#,##0.00")
Field39.Text = amnt
Field39.Text = Format(Field39.Text, "#,##0.00")
End Sub

