VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public db As ADODB.Connection
Public rs As ADODB.Recordset
Public Function Database(ByVal database2 As String, ByVal query As String)
Set db = New ADODB.Connection
Source = database2
db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Source
db.Open
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open query, db, adOpenKeyset, adLockOptimistic
End Function



