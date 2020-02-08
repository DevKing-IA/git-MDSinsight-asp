<%	
	Set cnnAPI_IC_AdjustOnHand = Server.CreateObject("ADODB.Connection")
	cnnAPI_IC_AdjustOnHand.open (Session("ClientCnnString"))
	Set rsAPI_IC_AdjustOnHand = Server.CreateObject("ADODB.Recordset")
	rsAPI_IC_AdjustOnHand.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsAPI_IC_AdjustOnHand = cnnAPI_IC_AdjustOnHand.Execute("SELECT TOP 1 * FROM API_IC_AdjustOnHand ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLAPI_IC_AdjustOnHand = "CREATE TABLE [API_IC_AdjustOnHand]( "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_api_tblIC_AdjustOnHand_RecordCreated]  DEFAULT (getdate()), "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [TempUnqiueRecordIdentifier] [varchar](255) NULL, "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [Thread] [int] NULL, "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [prodSKU] [varchar](255) NULL, "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [orig_prodSKU] [varchar](255) NULL, "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [prodUM] [varchar](255) NULL, "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [orig_prodUM] [varchar](255) NULL, "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [prodDesc] [varchar](255) NULL, "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [Qty] [int] NULL, "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [Comment] [varchar](8000) NULL, "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & " [IPAddress] [varchar](255) NULL  "
			SQLAPI_IC_AdjustOnHand = SQLAPI_IC_AdjustOnHand & ") ON [PRIMARY]"
		
			Set rsAPI_IC_AdjustOnHand = cnnAPI_IC_AdjustOnHand.Execute(SQLAPI_IC_AdjustOnHand)
		End If
	End If
	
	'Drop Mis-spelled column
	SQLAPI_IC_AdjustOnHand = "SELECT COL_LENGTH('API_IC_AdjustOnHand', '[[Thread]') AS IsItThere"
	Set rsAPI_IC_AdjustOnHand = cnnAPI_IC_AdjustOnHand.Execute(SQLAPI_IC_AdjustOnHand)
	If NOT IsNull(rsAPI_IC_AdjustOnHand("IsItThere")) Then
		SQLAPI_IC_AdjustOnHand = "ALTER TABLE API_IC_AdjustOnHand DROP COLUMN [[Thread]"
		Set rsAPI_IC_AdjustOnHand = cnnAPI_IC_AdjustOnHand.Execute(SQLAPI_IC_AdjustOnHand)
	End If

	'Add in with correct spelling
	SQLAPI_IC_AdjustOnHand = "SELECT COL_LENGTH('API_IC_AdjustOnHand', 'Thread') AS IsItThere"
	Set rsAPI_IC_AdjustOnHand  = cnnAPI_IC_AdjustOnHand.Execute(SQLAPI_IC_AdjustOnHand )
	If IsNull(rsAPI_IC_AdjustOnHand ("IsItThere")) Then
		SQLAPI_IC_AdjustOnHand = "ALTER TABLE API_IC_AdjustOnHand  ADD Thread INT NULL"
		Set rsAPI_IC_AdjustOnHand = cnnAPI_IC_AdjustOnHand.Execute(SQLAPI_IC_AdjustOnHand)
	End If

	'Drop unused column
	SQLAPI_IC_AdjustOnHand = "SELECT COL_LENGTH('API_IC_AdjustOnHand', 'prodID') AS IsItThere"
	Set rsAPI_IC_AdjustOnHand = cnnAPI_IC_AdjustOnHand.Execute(SQLAPI_IC_AdjustOnHand)
	If NOT IsNull(rsAPI_IC_AdjustOnHand("IsItThere")) Then
		SQLAPI_IC_AdjustOnHand = "ALTER TABLE API_IC_AdjustOnHand DROP COLUMN prodID"
		Set rsAPI_IC_AdjustOnHand = cnnAPI_IC_AdjustOnHand.Execute(SQLAPI_IC_AdjustOnHand)
	End If

	'Drop unused column
	SQLAPI_IC_AdjustOnHand = "SELECT COL_LENGTH('API_IC_AdjustOnHand', 'UM') AS IsItThere"
	Set rsAPI_IC_AdjustOnHand = cnnAPI_IC_AdjustOnHand.Execute(SQLAPI_IC_AdjustOnHand)
	If NOT IsNull(rsAPI_IC_AdjustOnHand("IsItThere")) Then
		SQLAPI_IC_AdjustOnHand = "ALTER TABLE API_IC_AdjustOnHand DROP COLUMN UM"
		Set rsAPI_IC_AdjustOnHand = cnnAPI_IC_AdjustOnHand.Execute(SQLAPI_IC_AdjustOnHand)
	End If
	
	set rsAPI_IC_AdjustOnHand = nothing
	cnnAPI_IC_AdjustOnHand.close
	set cnnAPI_IC_AdjustOnHand = nothing
				
%>