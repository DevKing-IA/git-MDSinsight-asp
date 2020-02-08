<%	

	Set cnnCheckPRLeadSources = Server.CreateObject("ADODB.Connection")
	cnnCheckPRLeadSources.open (Session("ClientCnnString"))
	Set rsCheckPRLeadSources = Server.CreateObject("ADODB.Recordset")
	rsCheckPRLeadSources.CursorLocation = 3 

	Err.Clear
	
	On Error Goto 0

	'Make sure code 0 is there
	SQLCheckPRLeadSources  = "SELECT * FROM PR_LeadSources WHERE InternalRecordIdentifier = 0"
	Set rsCheckPRLeadSources = cnnCheckPRLeadSources.Execute(SQLCheckPRLeadSources)
	If rsCheckPRLeadSources.EOF Then 
	
		SQLCheckPRLeadSources = "SET IDENTITY_INSERT PR_LeadSources ON;"
		Set rsCheckPRLeadSources = cnnCheckPRLeadSources.Execute(SQLCheckPRLeadSources)

		SQLCheckPRLeadSources = SQLCheckPRLeadSources & "INSERT INTO PR_LeadSources (InternalRecordIdentifier,LeadSource) "
		SQLCheckPRLeadSources = SQLCheckPRLeadSources & " VALUES (0,'Undefined')"
		Response.Write(SQLCheckPRLeadSources)
		Set rsCheckPRLeadSources = cnnCheckPRLeadSources.Execute(SQLCheckPRLeadSources)
		
		SQLCheckPRLeadSources = "SET IDENTITY_INSERT PR_LeadSources OFF;"
		Set rsCheckPRLeadSources = cnnCheckPRLeadSources.Execute(SQLCheckPRLeadSources)
		
	End If
		
	set rsCheckPRLeadSources = nothing
	cnnCheckPRLeadSources.close
	set cnnCheckPRLeadSources = nothing
%>