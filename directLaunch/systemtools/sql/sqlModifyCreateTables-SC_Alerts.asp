<%	

	Response.Write("sqlModifyCreateTables-SC_Alerts.asp" & "<br>")
	On Error Goto 0

	Set cnnCheckSCAlerts = Server.CreateObject("ADODB.Connection")
	cnnCheckSCAlerts.open (Session("ClientCnnString"))
	Set rsCheckSCAlerts = Server.CreateObject("ADODB.Recordset")
	rsCheckSCAlerts.CursorLocation = 3 
	
	SQL_CheckSCAlerts = "SELECT COL_LENGTH('SC_Alerts', 'DayOfMonth') AS IsItThere"
	Set rsCheckSCAlerts = cnnCheckSCAlerts.Execute(SQL_CheckSCAlerts)
	If IsNull(rsCheckSCAlerts("IsItThere")) Then
		SQL_CheckSCAlerts = "ALTER TABLE SC_Alerts ADD DayOfMonth int NULL"
		Set rsCheckSCAlerts = cnnCheckSCAlerts.Execute(SQL_CheckSCAlerts)
	End If

	SQL_CheckSCAlerts = "SELECT COL_LENGTH('SC_Alerts', 'EmailPrimarySls') AS IsItThere"
	Set rsCheckSCAlerts = cnnCheckSCAlerts.Execute(SQL_CheckSCAlerts)
	If IsNull(rsCheckSCAlerts("IsItThere")) Then
		SQL_CheckSCAlerts = "ALTER TABLE SC_Alerts ADD EmailPrimarySls int NULL"
		Set rsCheckSCAlerts = cnnCheckSCAlerts.Execute(SQL_CheckSCAlerts)
	End If

	SQL_CheckSCAlerts = "SELECT COL_LENGTH('SC_Alerts', 'EmailSecondarySls') AS IsItThere"
	Set rsCheckSCAlerts = cnnCheckSCAlerts.Execute(SQL_CheckSCAlerts)
	If IsNull(rsCheckSCAlerts("IsItThere")) Then
		SQL_CheckSCAlerts = "ALTER TABLE SC_Alerts ADD EmailSecondarySls int NULL"
		Set rsCheckSCAlerts = cnnCheckSCAlerts.Execute(SQL_CheckSCAlerts)
	End If

	Set rsCheckSCAlerts = Nothing
	cnnCheckSCAlerts.Close
	Set cnnCheckSCAlerts = Nothing
				
%>