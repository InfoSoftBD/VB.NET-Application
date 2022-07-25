Attribute VB_Name = "Module1"
Global conn As ADODB.Connection
Global Today As Date
Global Power As String
Global Org_Name As String
Global Branch_Code As String
Global Branch_Name As String
Global Branch_Address As String
Global User_Id As String
Global User_Name As String
Global U_Type As String
Global prt As String
Global cnStr As String
Public Sub main()
Dim dbpath As String
Dim dbName As String
Dim rs As ADODB.Recordset
Dim str As String

dbName = "Master.mdb"
dbpath = App.Path + "\" + dbName
cnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath & " ;Persist Security Info=False;Jet OLEDB:Database Password=136873;Jet OLEDB:Encrypt Database=True;Jet OLEDB:Compact Without Replica Repair=True"
Set conn = New ADODB.Connection
conn.Open cnStr
frmSplash.Show 1
End Sub
