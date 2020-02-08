<%	

	Set cnnCheckSCAlertsSent = Server.CreateObject("ADODB.Connection")
	cnnCheckSCAlertsSent.open (Session("ClientCnnString"))
	Set rsCheckSCAlertsSent = Server.CreateObject("ADODB.Recordset")
	rsCheckSCAlertsSent.CursorLocation = 3 
	
	SQL_CheckSCAlertsSent = "SELECT COL_LENGTH('SC_AlertsSent', 'ServiceMemoDetailRecIfApplicable') AS IsItThere"
	Set rsCheckSCAlertsSent = cnnCheckSCAlertsSent.Execute(SQL_CheckSCAlertsSent)
	If IsNull(rsCheckSCAlertsSent("IsItThere")) Then
		SQL_CheckSCAlertsSent = "ALTER TABLE SC_AlertsSent ADD ServiceMemoDetailRecIfApplicable int NULL"
		Set rsCheckSCAlertsSent = cnnCheckSCAlertsSent.Execute(SQL_CheckSCAlertsSent)
	End If
	
	SQL_CheckSCAlertsSent = "SELECT COL_LENGTH('SC_AlertsSent', 'UserNameSentToIfApplicable') AS IsItThere"
	Set rsCheckSCAlertsSent = cnnCheckSCAlertsSent.Execute(SQL_CheckSCAlertsSent)
	If IsNull(rsCheckSCAlertsSent("IsItThere")) Then
		SQL_CheckSCAlertsSent = "ALTER TABLE SC_AlertsSent ADD UserNameSentToIfApplicable varchar(255) NULL"
		Set rsCheckSCAlertsSent = cnnCheckSCAlertsSent.Execute(SQL_CheckSCAlertsSent)
	End If
	
	
	SQL_CheckSCAlertsSent = "SELECT COL_LENGTH('SC_AlertsSent', 'OrderIDIfApplicable') AS IsItThere"
	Set rsCheckSCAlertsSent = cnnCheckSCAlertsSent.Execute(SQL_CheckSCAlertsSent)
	If IsNull(rsCheckSCAlertsSent("IsItThere")) Then
		SQL_CheckSCAlertsSent = "ALTER TABLE SC_AlertsSent ADD OrderIDIfApplicable varchar(50) NULL"
		Set rsCheckSCAlertsSent = cnnCheckSCAlertsSent.Execute(SQL_CheckSCAlertsSent)
	End If
	

	Set rsCheckSCAlertsSent = Nothing
	cnnCheckSCAlertsSent.Close
	Set cnnCheckSCAlertsSent = Nothing
				
%>