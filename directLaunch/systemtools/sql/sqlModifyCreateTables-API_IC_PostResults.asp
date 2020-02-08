<%	
	Set cnnAPI_IC_PostResults = Server.CreateObject("ADODB.Connection")
	cnnAPI_IC_PostResults.open (Session("ClientCnnString"))
	Set rsAPI_IC_PostResults = Server.CreateObject("ADODB.Recordset")
	rsAPI_IC_PostResults.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsAPI_IC_PostResults = cnnAPI_IC_PostResults.Execute("SELECT TOP 1 * FROM API_IC_PostResults ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLAPI_IC_PostResults = "CREATE TABLE [API_IC_PostResults]( "
			SQLAPI_IC_PostResults = SQLAPI_IC_PostResults & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLAPI_IC_PostResults = SQLAPI_IC_PostResults & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_api_tblIC_PostResults_RecordCreated]  DEFAULT (getdate()), "
			SQLAPI_IC_PostResults = SQLAPI_IC_PostResults & " [APIKey] [varchar](50) NULL, "
			SQLAPI_IC_PostResults = SQLAPI_IC_PostResults & " 	[LogEntryThread] [int] NULL, "
			SQLAPI_IC_PostResults = SQLAPI_IC_PostResults & " 	[PostResults] [varchar](50) NULL, "
			SQLAPI_IC_PostResults = SQLAPI_IC_PostResults & " 	[PostStatusMessage] [varchar](8000) NULL "
			SQLAPI_IC_PostResults = SQLAPI_IC_PostResults & ") ON [PRIMARY]"
		
			Set rsAPI_IC_PostResults = cnnAPI_IC_PostResults.Execute(SQLAPI_IC_PostResults)
		End If
	End If

	'Drop unused column
	SQLAPI_IC_PostResults = "SELECT COL_LENGTH('API_IC_PostResults', 'Comment') AS IsItThere"
	Set rsAPI_IC_PostResults = cnnAPI_IC_PostResults.Execute(SQLAPI_IC_PostResults)
	If NOT IsNull(rsAPI_IC_PostResults("IsItThere")) Then
		SQLAPI_IC_PostResults = "ALTER TABLE API_IC_PostResults DROP COLUMN Comment"
		Set rsAPI_IC_PostResults = cnnAPI_IC_PostResults.Execute(SQLAPI_IC_PostResults)
	End If

	SQLAPI_IC_PostResults = "SELECT COL_LENGTH('API_IC_PostResults', 'prodSKU') AS IsItThere"
	Set rsAPI_IC_PostResults  = cnnAPI_IC_PostResults.Execute(SQLAPI_IC_PostResults)
	If IsNull(rsAPI_IC_PostResults("IsItThere")) Then
		SQLAPI_IC_PostResults = "ALTER TABLE API_IC_PostResults  ADD prodSKU varchar(255) NULL"
		Set rsAPI_IC_PostResults = cnnAPI_IC_PostResults.Execute(SQLAPI_IC_PostResults)
	End If
	
	SQLAPI_IC_PostResults = "SELECT COL_LENGTH('API_IC_PostResults', 'UM') AS IsItThere"
	Set rsAPI_IC_PostResults  = cnnAPI_IC_PostResults.Execute(SQLAPI_IC_PostResults)
	If IsNull(rsAPI_IC_PostResults("IsItThere")) Then
		SQLAPI_IC_PostResults = "ALTER TABLE API_IC_PostResults  ADD UM varchar(255) NULL"
		Set rsAPI_IC_PostResults = cnnAPI_IC_PostResults.Execute(SQLAPI_IC_PostResults)
	End If
	
	SQLAPI_IC_PostResults = "SELECT COL_LENGTH('API_IC_PostResults', 'Qty') AS IsItThere"
	Set rsAPI_IC_PostResults  = cnnAPI_IC_PostResults.Execute(SQLAPI_IC_PostResults)
	If IsNull(rsAPI_IC_PostResults("IsItThere")) Then
		SQLAPI_IC_PostResults = "ALTER TABLE API_IC_PostResults  ADD Qty int NULL"
		Set rsAPI_IC_PostResults = cnnAPI_IC_PostResults.Execute(SQLAPI_IC_PostResults)
	End If

	
	set rsAPI_IC_PostResults = nothing
	cnnAPI_IC_PostResults.close
	set cnnAPI_IC_PostResults = nothing
				
%>