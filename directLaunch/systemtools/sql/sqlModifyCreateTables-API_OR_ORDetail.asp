<%	

	Set cnnOR_ORDetail = Server.CreateObject("ADODB.Connection")
	cnnOR_ORDetail.open (Session("ClientCnnString"))
	Set rsOR_ORDetail = Server.CreateObject("ADODB.Recordset")
	rsOR_ORDetail.CursorLocation = 3 
			
	'See if these fields are in the table& add them if not there

	SQL_OR_ORDetail = "SELECT COL_LENGTH('API_OR_OrderDetail', 'Source') AS IsItThere"
	Set rsOR_ORDetail = cnnOR_ORDetail.Execute(SQL_OR_ORDetail)
	If IsNull(rsOR_ORDetail("IsItThere")) Then
		SQL_OR_ORDetail = "ALTER TABLE API_OR_OrderDetail ADD Source varchar(255) NULL"
		Set rsOR_ORDetail = cnnOR_ORDetail.Execute(SQL_OR_ORDetail)
	End If
	
	Set rsOR_ORDetail = Nothing
	cnnOR_ORDetail.Close
	Set cnnOR_ORDetail = Nothing

%>