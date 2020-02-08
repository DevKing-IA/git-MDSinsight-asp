<%	
	Set cnnBI_MCSData = Server.CreateObject("ADODB.Connection")
	cnnBI_MCSData.open (Session("ClientCnnString"))
	Set rsBI_MCSData = Server.CreateObject("ADODB.Recordset")
	rsBI_MCSData.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsBI_MCSData = cnnBI_MCSData.Execute("SELECT TOP 1 * FROM BI_MCSData ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBI_MCSData = "CREATE TABLE [BI_MCSData]( "
			SQLBI_MCSData = SQLBI_MCSData & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLBI_MCSData = SQLBI_MCSData & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_biMCSData]  DEFAULT (getdate()), "
			SQLBI_MCSData = SQLBI_MCSData & " [CustID] [varchar](255) NULL, "
			SQLBI_MCSData = SQLBI_MCSData & " [Month1Sales_NoRent] [money] NULL, "
			SQLBI_MCSData = SQLBI_MCSData & " [Month2Sales_NoRent] [money] NULL, "
			SQLBI_MCSData = SQLBI_MCSData & " [Month3Sales_NoRent] [money] NULL, "
			SQLBI_MCSData = SQLBI_MCSData & " [Month3Cost_NoRent] [money] NULL, "
			SQLBI_MCSData = SQLBI_MCSData & " [LVFHolder] [money] NULL, "
			SQLBI_MCSData = SQLBI_MCSData & " [LVFHolderCurrent] [money] NULL, "
			SQLBI_MCSData = SQLBI_MCSData & " [TotalEquipmentValue] [money] NULL, "
			SQLBI_MCSData = SQLBI_MCSData & " [CurrentHolder] [money] NULL, "
			SQLBI_MCSData = SQLBI_MCSData & " [RentalHolder] [money] NULL "
			SQLBI_MCSData = SQLBI_MCSData & ") ON [PRIMARY]"
		
			Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
		End If
	End If

' This one is a DROP
	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'Cat21Sales') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If NOT IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData DROP COLUMN Cat21Sales"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If

' This one is a DROP
	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'Month12SF') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If NOT IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData DROP COLUMN Month12SF"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If

' This one is a DROP
	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'Month2SF') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If NOT IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData DROP COLUMN Month2SF"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If


	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'Month1Cat21Sales') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData  ADD Month1Cat21Sales money NULL"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If

	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'Month2Cat21Sales') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData  ADD Month2Cat21Sales money NULL"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If

	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'Month3Cat21Sales') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData  ADD Month3Cat21Sales money NULL"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If

	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'Month1XSF') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData  ADD Month1XSF money NULL"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If

	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'Month2XSF') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData  ADD Month2XSF money NULL"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If

	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'Month3XSF') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData  ADD Month3XSF money NULL"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If

	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'CurrentXSF') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData  ADD CurrentXSF money NULL"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If

	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'PendingLVF') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData  ADD PendingLVF money NULL"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If

	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'ChainID') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData  ADD ChainID varchar(255) NULL"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If

	SQLBI_MCSData = "SELECT COL_LENGTH('BI_MCSData', 'ChainName') AS IsItThere"
	Set rsBI_MCSData  = cnnBI_MCSData.Execute(SQLBI_MCSData)
	If IsNull(rsBI_MCSData("IsItThere")) Then
		SQLBI_MCSData = "ALTER TABLE BI_MCSData  ADD ChainName varchar(255) NULL"
		Set rsBI_MCSData = cnnBI_MCSData.Execute(SQLBI_MCSData)
	End If



	set rsBI_MCSData = nothing
	cnnBI_MCSData.close
	set cnnBI_MCSData = nothing
				
%>