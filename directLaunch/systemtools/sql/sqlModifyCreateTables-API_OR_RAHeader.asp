<%	

	Set cnnOR_RAHeader = Server.CreateObject("ADODB.Connection")
	cnnOR_RAHeader.open (Session("ClientCnnString"))
	Set rsOR_RAHeader = Server.CreateObject("ADODB.Recordset")
	rsOR_RAHeader.CursorLocation = 3 
			
	'See if these fields are in the table& add them if not there
	
	SQL_OR_RAHeader = "SELECT COL_LENGTH('API_OR_RAHeader', 'Orig_RAID') AS IsItThere"
	Set rsOR_RAHeader = cnnOR_RAHeader.Execute(SQL_OR_RAHeader)
	If IsNull(rsOR_RAHeader("IsItThere")) Then
		SQL_OR_RAHeader = "ALTER TABLE API_OR_RAHeader ADD Orig_RAID varchar(50) NULL"
		Set rsOR_RAHeader = cnnOR_RAHeader.Execute(SQL_OR_RAHeader)
	End If
	
	Set rsOR_RAHeader = Nothing
	cnnOR_RAHeader.Close
	Set cnnOR_RAHeader = Nothing

%>