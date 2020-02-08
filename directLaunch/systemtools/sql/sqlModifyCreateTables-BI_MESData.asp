<%	
	Set cnnBI_MESData = Server.CreateObject("ADODB.Connection")
	cnnBI_MESData.open (Session("ClientCnnString"))
	Set rsBI_MESData = Server.CreateObject("ADODB.Recordset")
	rsBI_MESData.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsBI_MESData = cnnBI_MESData.Execute("SELECT TOP 1 * FROM BI_MESData ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBI_MESData = "CREATE TABLE [BI_MESData]( "
			SQLBI_MESData = SQLBI_MESData & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLBI_MESData = SQLBI_MESData & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_biMESData]  DEFAULT (getdate()), "
			SQLBI_MESData = SQLBI_MESData & " [CustID] [varchar](255) NULL, "
			SQLBI_MESData = SQLBI_MESData & " [Month1Sales_NoRent] [money] NULL, "
			SQLBI_MESData = SQLBI_MESData & " [Month2Sales_NoRent] [money] NULL, "
			SQLBI_MESData = SQLBI_MESData & " [Month3Sales_NoRent] [money] NULL, "
			SQLBI_MESData = SQLBI_MESData & " [Month3Cost_NoRent] [money] NULL, "
			SQLBI_MESData = SQLBI_MESData & " [TotalEquipmentValue] [money] NULL, "
			SQLBI_MESData = SQLBI_MESData & " [CurrentHolder] [money] NULL, "
			SQLBI_MESData = SQLBI_MESData & " [RentalHolder] [money] NULL "
			SQLBI_MESData = SQLBI_MESData & ") ON [PRIMARY]"
		
			Set rsBI_MESData = cnnBI_MESData.Execute(SQLBI_MESData)
		End If
	End If

' This one is a DROP
	SQLBI_MESData = "SELECT COL_LENGTH('BI_MESData', 'Cat21Sales') AS IsItThere"
	Set rsBI_MESData  = cnnBI_MESData.Execute(SQLBI_MESData)
	If NOT IsNull(rsBI_MESData("IsItThere")) Then
		SQLBI_MESData = "ALTER TABLE BI_MESData DROP COLUMN Cat21Sales"
		Set rsBI_MESData = cnnBI_MESData.Execute(SQLBI_MESData)
	End If

' This one is a DROP
	SQLBI_MESData = "SELECT COL_LENGTH('BI_MESData', 'LVFHolder') AS IsItThere"
	Set rsBI_MESData  = cnnBI_MESData.Execute(SQLBI_MESData)
	If NOT IsNull(rsBI_MESData("IsItThere")) Then
		SQLBI_MESData = "ALTER TABLE BI_MESData DROP COLUMN LVFHolder"
		Set rsBI_MESData = cnnBI_MESData.Execute(SQLBI_MESData)
	End If

' This one is a DROP
	SQLBI_MESData = "SELECT COL_LENGTH('BI_MESData', 'LVFHolderCurrent') AS IsItThere"
	Set rsBI_MESData  = cnnBI_MESData.Execute(SQLBI_MESData)
	If NOT IsNull(rsBI_MESData("IsItThere")) Then
		SQLBI_MESData = "ALTER TABLE BI_MESData DROP COLUMN LVFHolderCurrent"
		Set rsBI_MESData = cnnBI_MESData.Execute(SQLBI_MESData)
	End If


	SQLBI_MESData = "SELECT COL_LENGTH('BI_MESData', 'Month1Cat21Sales') AS IsItThere"
	Set rsBI_MESData  = cnnBI_MESData.Execute(SQLBI_MESData)
	If IsNull(rsBI_MESData("IsItThere")) Then
		SQLBI_MESData = "ALTER TABLE BI_MESData  ADD Month1Cat21Sales money NULL"
		Set rsBI_MESData = cnnBI_MESData.Execute(SQLBI_MESData)
	End If

	SQLBI_MESData = "SELECT COL_LENGTH('BI_MESData', 'Month2Cat21Sales') AS IsItThere"
	Set rsBI_MESData  = cnnBI_MESData.Execute(SQLBI_MESData)
	If IsNull(rsBI_MESData("IsItThere")) Then
		SQLBI_MESData = "ALTER TABLE BI_MESData  ADD Month2Cat21Sales money NULL"
		Set rsBI_MESData = cnnBI_MESData.Execute(SQLBI_MESData)
	End If

	SQLBI_MESData = "SELECT COL_LENGTH('BI_MESData', 'Month3Cat21Sales') AS IsItThere"
	Set rsBI_MESData  = cnnBI_MESData.Execute(SQLBI_MESData)
	If IsNull(rsBI_MESData("IsItThere")) Then
		SQLBI_MESData = "ALTER TABLE BI_MESData  ADD Month3Cat21Sales money NULL"
		Set rsBI_MESData = cnnBI_MESData.Execute(SQLBI_MESData)
	End If

	
	set rsBI_MESData = nothing
	cnnBI_MESData.close
	set cnnBI_MESData = nothing
				
%>