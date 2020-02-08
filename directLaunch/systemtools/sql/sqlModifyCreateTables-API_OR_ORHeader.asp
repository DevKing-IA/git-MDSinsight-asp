<%	

	Set cnnOR_ORHeader = Server.CreateObject("ADODB.Connection")
	cnnOR_ORHeader.open (Session("ClientCnnString"))
	Set rsOR_ORHeader = Server.CreateObject("ADODB.Recordset")
	rsOR_ORHeader.CursorLocation = 3 
			
	'See if these fields are in the table& add them if not there
	
	SQL_OR_ORHeader = "SELECT COL_LENGTH('API_OR_OrderHeader', 'Orig_CustID') AS IsItThere"
	Set rsOR_ORHeader = cnnOR_ORHeader.Execute(SQL_OR_ORHeader)
	If IsNull(rsOR_ORHeader("IsItThere")) Then
		SQL_OR_ORHeader = "ALTER TABLE API_OR_OrderHeader ADD Orig_CustID varchar(50) NULL"
		Set rsOR_ORHeader = cnnOR_ORHeader.Execute(SQL_OR_ORHeader)
	End If

	SQL_OR_ORHeader = "SELECT COL_LENGTH('API_OR_OrderHeader', 'Source') AS IsItThere"
	Set rsOR_ORHeader = cnnOR_ORHeader.Execute(SQL_OR_ORHeader)
	If IsNull(rsOR_ORHeader("IsItThere")) Then
		SQL_OR_ORHeader = "ALTER TABLE API_OR_OrderHeader ADD Source varchar(255) NULL"
		Set rsOR_ORHeader = cnnOR_ORHeader.Execute(SQL_OR_ORHeader)
	End If
	
	Set rsOR_ORHeader = Nothing
	cnnOR_ORHeader.Close
	Set cnnOR_ORHeader = Nothing

%>