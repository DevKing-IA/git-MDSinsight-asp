<%	

	Response.Write("sqlModifyCreateTables-AR_CustomerMapping.asp" & "<br>")
	On Error Goto 0
	
	Set cnnARCustomerMapping = Server.CreateObject("ADODB.Connection")
	cnnARCustomerMapping.open (Session("ClientCnnString"))
	Set rsARCustomerMapping = Server.CreateObject("ADODB.Recordset")
	rsARCustomerMapping.CursorLocation = 3 
			
	
	SQL_ARCustomerMapping = "SELECT COL_LENGTH('AR_CustomerMapping', 'RecordSource') AS IsItThere"
	Set rsARCustomerMapping = cnnARCustomerMapping.Execute(SQL_ARCustomerMapping)
	If IsNull(rsARCustomerMapping("IsItThere")) Then
		SQL_ARCustomerMapping = "ALTER TABLE AR_CustomerMapping ADD RecordSource varchar(255) NULL"
		Set rsARCustomerMapping = cnnARCustomerMapping.Execute(SQL_ARCustomerMapping)
		SQL_ARCustomerMapping = "UPDATE AR_CustomerMapping SET RecordSource = 'Insight'"
		Set rsARCustomerMapping = cnnARCustomerMapping.Execute(SQL_ARCustomerMapping)
	End If

	'This one is needed too until the code cathes up
	SQL_ARCustomerMapping = "UPDATE AR_CustomerMapping SET RecordSource = 'Insight' WHERE RecordSource IS NULL"
	Set rsARCustomerMapping = cnnARCustomerMapping.Execute(SQL_ARCustomerMapping)

	SQL_ARCustomerMapping = "SELECT COL_LENGTH('AR_CustomerMapping', 'partnerShipToID') AS IsItThere"
	Set rsARCustomerMapping = cnnARCustomerMapping.Execute(SQL_ARCustomerMapping)
	If IsNull(rsARCustomerMapping("IsItThere")) Then
		SQL_ARCustomerMapping = "ALTER TABLE AR_CustomerMapping ADD partnerShipToID varchar(255) NULL"
		Set rsARCustomerMapping = cnnARCustomerMapping.Execute(SQL_ARCustomerMapping)
	End If

	Set rsARCustomerMapping = Nothing
	cnnARCustomerMapping.Close
	Set cnnARCustomerMapping = Nothing
				
%>