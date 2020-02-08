<%	
	Set cnnAPI_IN_InvoiceHeader = Server.CreateObject("ADODB.Connection")
	cnnAPI_IN_InvoiceHeader.open (Session("ClientCnnString"))
	Set rsAPI_IN_InvoiceHeader = Server.CreateObject("ADODB.Recordset")
	rsAPI_IN_InvoiceHeader.CursorLocation = 3 

	Err.Clear

	SQLAPI_IN_InvoiceHeader = "SELECT COL_LENGTH('API_IN_InvoiceHeader', 'NO_ORDER') AS IsItThere"
	Set rsAPI_IN_InvoiceHeader = cnnAPI_IN_InvoiceHeader.Execute(SQLAPI_IN_InvoiceHeader)
	If IsNull(rsAPI_IN_InvoiceHeader("IsItThere")) Then
		SQLAPI_IN_InvoiceHeader  = "ALTER TABLE API_IN_InvoiceHeader ADD NO_ORDER varchar(255) NULL"
		Set rsAPI_IN_InvoiceHeader = cnnAPI_IN_InvoiceHeader.Execute(SQLAPI_IN_InvoiceHeader)
	End If

	SQLAPI_IN_InvoiceHeader = "SELECT COL_LENGTH('API_IN_InvoiceHeader', 'POSTING_STATUS') AS IsItThere"
	Set rsAPI_IN_InvoiceHeader = cnnAPI_IN_InvoiceHeader.Execute(SQLAPI_IN_InvoiceHeader)
	If IsNull(rsAPI_IN_InvoiceHeader("IsItThere")) Then
		SQLAPI_IN_InvoiceHeader  = "ALTER TABLE API_IN_InvoiceHeader ADD POSTING_STATUS varchar(255) NULL"
		Set rsAPI_IN_InvoiceHeader = cnnAPI_IN_InvoiceHeader.Execute(SQLAPI_IN_InvoiceHeader)
	End If

	set rsAPI_IN_InvoiceHeader = nothing
	cnnAPI_IN_InvoiceHeader.close
	set cnnAPI_IN_InvoiceHeader = nothing
%>