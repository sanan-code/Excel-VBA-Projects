VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsActiveXDataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_DBFileAddress As String
Private m_ConnectionString As String

Private conn As ADODB.Connection



Private Sub Class_Initialize()
m_DBFileAddress = ThisWorkbook.Sheets(Sheet2.Name).Range("B2").Value
m_ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & m_DBFileAddress & ";Persist Security Info=False;"

Set conn = New ADODB.Connection
conn.Open m_ConnectionString
End Sub

Private Sub Class_Terminate()
conn.Close
Set conn = Nothing
End Sub

Public Sub ExecuteQuery(query As String)
conn.Execute query
End Sub

Public Sub ExecuteQueries(queries() As String)
For i = LBound(queries) To UBound(queries)
  conn.Execute queries(i)
Next i
End Sub

Public Function FetchDatas(table As String) As Variant
'default returns Empty

Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM " & table, conn, adOpenStatic, adLockReadOnly

If Not (rs.EOF And rs.BOF) Then FetchDatas = rs.GetRows()

rs.Close
Set rs = Nothing
End Function





