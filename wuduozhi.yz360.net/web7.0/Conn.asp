<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Session.Timeout=600
'Option Explicit
'Response.Buffer = True
Dim StarTime
Dim Conn,Rs,Sql,ConnStr,MyDbPath,Db,SqlNow
StarTime=Timer()

'MsSQL数据库设置
Const IsSqlDataBase = 0                          '是否使用MsSQL数据库，1是 0否
Const SqlDatabaseName = "miloc"
Const SqlUsername = "sa"
Const SqlPassword = "123456"
Const SqlLocalName = "(local)"

'Access数据库路径
MyDbPath=""
Db="xgdb/db.mdb"

Sub ConnectionDatabase()
	If IsSqlDataBase = 1 Then
		SqlNow = "GetDate()"
	'	ConnStr = "Provider = SQLOLEDB; Data Source = " & SqlLocalName & "; UID = " & SqlUsername & "; PWD = " & SqlPassword & "; DATABASE = " & SqlDatabaseName & ";"
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	Else
		SqlNow = "Now()"
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(MyDbPath & db)
	End If

	On Error Resume Next
	Set Conn = Server.CreateObject("ADODB.Connection")
	Set Rs=Server.CreateObject("ADODB.Recordset")
	Conn.Open ConnStr
	If Err Then
		err.Clear
		Set Conn = Nothing
		Response.Write("数据库连接出错，请检查连接字串")
		Response.End
	End If
End Sub

Sub CloseConn()
'	Rs.Close
	Set Rs=Nothing
	Conn.Close
	Set Conn=Nothing
End Sub
%>
