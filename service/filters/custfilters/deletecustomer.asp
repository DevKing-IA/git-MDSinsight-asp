
		<%
		customerID=REQUEST("customerID")
		Set cnnCustFiltersList = Server.CreateObject("ADODB.Connection")
		cnnCustFiltersList.open (Session("ClientCnnString"))
		SQl="DELETE FROM FS_CustomerFilters WHERE CustID=" & customerID
		cnnCustFiltersList.Execute(SQL)
		
		cnnCustFiltersList.Close
		
		%>
		