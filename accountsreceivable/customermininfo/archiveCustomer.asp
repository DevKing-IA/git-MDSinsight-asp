<%

	customerID=REQUEST("customerID")
	Set cnnCustFiltersList = Server.CreateObject("ADODB.Connection")
	cnnCustFiltersList.open (Session("ClientCnnString"))
	SQL = "UPDATE AR_Customer SET AcctStatus = 'I' WHERE CustNum = '" & customerID & "'"
	cnnCustFiltersList.Execute(SQL)
	cnnCustFiltersList.Close

%>
		