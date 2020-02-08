<%	

	Set cnnCheckServiceMemosRedispatch = Server.CreateObject("ADODB.Connection")
	cnnCheckServiceMemosRedispatch.open (Session("ClientCnnString"))
	Set rsCheckServiceMemosRedispatch = Server.CreateObject("ADODB.Recordset")
	rsCheckServiceMemosRedispatch.CursorLocation = 3 

	SQL_CheckServiceMemosRedispatch = "SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE "
	SQL_CheckServiceMemosRedispatch = SQL_CheckServiceMemosRedispatch & " TABLE_NAME = 'FS_ServiceMemosRedispatch' AND "
	SQL_CheckServiceMemosRedispatch = SQL_CheckServiceMemosRedispatch & " COLUMN_NAME = 'MemoNumber'"
	
	Set rsCheckServiceMemosRedispatch = cnnCheckServiceMemosRedispatch.Execute(SQL_CheckServiceMemosRedispatch)
	If rsCheckServiceMemosRedispatch("DATA_TYPE") = "int" Then
		SQL_CheckServiceMemosRedispatch  = "ALTER TABLE FS_ServiceMemosRedispatch ALTER COLUMN MemoNumber varchar(255) NULL"
		Set rsCheckServiceMemosRedispatch = cnnCheckServiceMemosRedispatch.Execute(SQL_CheckServiceMemosRedispatch)
	End If
	
	Set rsCheckServiceMemosRedispatch = Nothing
	cnnCheckServiceMemosRedispatch.Close
	Set cnnCheckServiceMemosRedispatch = Nothing
%>