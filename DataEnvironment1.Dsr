VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DataEnvironment1 
   ClientHeight    =   13545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15000
   _ExtentX        =   26458
   _ExtentY        =   23892
   FolderFlags     =   1
   TypeLibGuid     =   "{7AE996D3-AD31-4214-8354-1E664F57C7CF}"
   TypeInfoGuid    =   "{5504A3DF-0952-455D-8C5B-3F24C325F430}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Connection1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   $"DataEnvironment1.dsx":0000
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   2
   BeginProperty Recordset1 
      CommandName     =   "Command1"
      CommDispId      =   1096
      RsDispId        =   1107
      CommandText     =   "Select * From Statement order by Sl_No"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      Grouping        =   -1  'True
      GroupingName    =   "Command1_Grouping"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Sl_no"
         Caption         =   "Sl_no"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Sl"
         Caption         =   "Sl"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "Date"
         Caption         =   "Date"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Name"
         Caption         =   "Name"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "AC_No"
         Caption         =   "AC_No"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Cr"
         Caption         =   "Cr"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Dr"
         Caption         =   "Dr"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Balance"
         Caption         =   "Balance"
      EndProperty
      NumGroups       =   2
      BeginProperty Grouping1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Name"
         Caption         =   "Name"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "AC_No"
         Caption         =   "AC_No"
      EndProperty
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "Command2"
      CommDispId      =   1103
      RsDispId        =   1108
      CommandText     =   "Select * From Statement order by Sl_No"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Sl_no"
         Caption         =   "Sl_no"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Sl"
         Caption         =   "Sl"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "Date"
         Caption         =   "Date"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Name"
         Caption         =   "Name"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "AC_No"
         Caption         =   "AC_No"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Cr"
         Caption         =   "Cr"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Dr"
         Caption         =   "Dr"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Balance"
         Caption         =   "Balance"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DataEnvironment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub DataEnvironment_Initialize()
Dim conn As ADODB.Connection
Dim cnStr As String
Dim dbpath As String
Dim dbName As String
Dim str As String
dbName = "Master.mdb"
dbpath = App.Path + "\" + dbName
cnStr = "DSN=MS Access Database;DBQ=" & dbpath & ";DefaultDir=App.Path;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=136873;UID=admin;"

'Set conn = New ADODB.Connection
'conn.Open cnStr

Connection1.ConnectionString = cnStr
End Sub

