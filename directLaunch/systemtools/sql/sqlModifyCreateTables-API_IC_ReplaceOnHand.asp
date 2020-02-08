<%	
	Set cnnAPI_IC_ReplaceOnHand = Server.CreateObject("ADODB.Connection")
	cnnAPI_IC_ReplaceOnHand.open (Session("ClientCnnString"))
	Set rsAPI_IC_ReplaceOnHand = Server.CreateObject("ADODB.Recordset")
	rsAPI_IC_ReplaceOnHand.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute("SELECT TOP 1 * FROM API_IC_ReplaceOnHand ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLAPI_IC_ReplaceOnHand = "CREATE TABLE [API_IC_ReplaceOnHand]( "
			SQLAPI_IC_ReplaceOnHand = SQLAPI_IC_ReplaceOnHand & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLAPI_IC_ReplaceOnHand = SQLAPI_IC_ReplaceOnHand & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_api_tblIC_ReplaceOnHand_RecordCreated]  DEFAULT (getdate()), "
			SQLAPI_IC_ReplaceOnHand = SQLAPI_IC_ReplaceOnHand & " [TempUnqiueRecordIdentifier] [varchar](255) NULL, "
			SQLAPI_IC_ReplaceOnHand = SQLAPI_IC_ReplaceOnHand & " [Thread] [int] NULL, "
			SQLAPI_IC_ReplaceOnHand = SQLAPI_IC_ReplaceOnHand & " [prodSKU] [varchar](255) NULL, "
			SQLAPI_IC_ReplaceOnHand = SQLAPI_IC_ReplaceOnHand & " [QtyCases] [int] NULL, "
			SQLAPI_IC_ReplaceOnHand = SQLAPI_IC_ReplaceOnHand & " [QtyUnits] [int] NULL, "
			SQLAPI_IC_ReplaceOnHand = SQLAPI_IC_ReplaceOnHand & " [UnitBin] [varchar](255) NULL, "
			SQLAPI_IC_ReplaceOnHand = SQLAPI_IC_ReplaceOnHand & " [CaseBin] [varchar](255) NULL, "			
			SQLAPI_IC_ReplaceOnHand = SQLAPI_IC_ReplaceOnHand & " [IPAddress] [varchar](255) NULL  "
			SQLAPI_IC_ReplaceOnHand = SQLAPI_IC_ReplaceOnHand & ") ON [PRIMARY]"
		
			Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
		End If
	End If
	
	'Drop unused column
	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'orig_prodSKU') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If NOT IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand = "ALTER TABLE API_IC_ReplaceOnHand DROP COLUMN orig_prodSKU"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	'Drop unused column
	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'orig_prodUM') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If NOT IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand = "ALTER TABLE API_IC_ReplaceOnHand DROP COLUMN orig_prodUM"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	'Drop unused column
	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'prodDesc') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If NOT IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand = "ALTER TABLE API_IC_ReplaceOnHand DROP COLUMN prodDesc"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	'Drop unused column
	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'prodUM') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If NOT IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand = "ALTER TABLE API_IC_ReplaceOnHand DROP COLUMN prodUM"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	'Drop unused column
	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'Qty') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If NOT IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand = "ALTER TABLE API_IC_ReplaceOnHand DROP COLUMN Qty"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	'Drop unused column
	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'Comment') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If NOT IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand = "ALTER TABLE API_IC_ReplaceOnHand DROP COLUMN Comment"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	'Drop unused column
	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'Bin') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If NOT IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand = "ALTER TABLE API_IC_ReplaceOnHand DROP COLUMN Bin"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If


	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'QtyCases') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand  = "ALTER TABLE API_IC_ReplaceOnHand ADD QtyCases int NULL"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'QtyUnits') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand  = "ALTER TABLE API_IC_ReplaceOnHand ADD QtyUnits int NULL"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'UnitBin') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand  = "ALTER TABLE API_IC_ReplaceOnHand ADD UnitBin varchar(255) NULL"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'CaseBin') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand  = "ALTER TABLE API_IC_ReplaceOnHand ADD CaseBin varchar(255) NULL"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'UnitUPC') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand  = "ALTER TABLE API_IC_ReplaceOnHand ADD UnitUPC varchar(255) NULL"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	SQLAPI_IC_ReplaceOnHand = "SELECT COL_LENGTH('API_IC_ReplaceOnHand', 'CaseUPC') AS IsItThere"
	Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	If IsNull(rsAPI_IC_ReplaceOnHand("IsItThere")) Then
		SQLAPI_IC_ReplaceOnHand  = "ALTER TABLE API_IC_ReplaceOnHand ADD CaseUPC varchar(255) NULL"
		Set rsAPI_IC_ReplaceOnHand = cnnAPI_IC_ReplaceOnHand.Execute(SQLAPI_IC_ReplaceOnHand)
	End If

	set rsAPI_IC_ReplaceOnHand = nothing
	cnnAPI_IC_ReplaceOnHand.close
	set cnnAPI_IC_ReplaceOnHand = nothing
				
%>