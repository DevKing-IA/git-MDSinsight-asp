<%	

	Set cnnICFilters = Server.CreateObject("ADODB.Connection")
	cnnICFilters.open (Session("ClientCnnString"))
	Set rsICFilters = Server.CreateObject("ADODB.Recordset")
	rsICFilters.CursorLocation = 3 


	Err.Clear
	on error resume next
	
	Set rsICFilters = cnnICFilters.Execute("SELECT TOP 1 * FROM IC_Filters")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLICFilters = "CREATE TABLE [IC_Filters]( "
			SQLICFilters = SQLICFilters & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLICFilters = SQLICFilters & "[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_IC_Filters_RecordCreationDateTime]  DEFAULT (getdate()),"	
			SQLICFilters = SQLICFilters & "[RecordSource] [varchar](255) NULL CONSTRAINT [DF_IC_Filters_RecordSource]  DEFAULT ('Insight'), "
			SQLICFilters = SQLICFilters & "[FilterID] [varchar](255) NULL, "
			SQLICFilters = SQLICFilters & "[Description] [varchar](8000) NULL, "
			SQLICFilters = SQLICFilters & "[DefaultCost] [money] NULL, "
			SQLICFilters = SQLICFilters & "[ListPrice] [money] NULL, "			
			SQLICFilters = SQLICFilters & "[InventoriedItem] [int] NULL CONSTRAINT [DF_IC_Filters_InventoriedItem]  DEFAULT (0), "
			SQLICFilters = SQLICFilters & "[PickableItem] [int] NULL CONSTRAINT [DF_IC_Filters_PickableItem]  DEFAULT (0), "
			SQLICFilters = SQLICFilters & "[Taxable] [int] NULL CONSTRAINT [DF_IC_Filters_Taxable]  DEFAULT (0), "
			SQLICFilters = SQLICFilters & "[UPCCode] [varchar](255) NULL "
			SQLICFilters = SQLICFilters & ") ON [PRIMARY] "
	
			Set rsICFilters = cnnICFilters.Execute(SQLICFilters)
			
		End If
	End If

	SQLICFilters = "SELECT COL_LENGTH('IC_Filters', 'prodSKU') AS IsItThere"
	Set rsICFilters = cnnICFilters.Execute(SQLICFilters )
	If IsNull(rsICFilters("IsItThere")) Then
		SQLICFilters = "ALTER TABLE IC_Filters ADD prodSKU varchar(1000) NULL"
		Set rsICFilters = cnnICFilters.Execute(SQLICFilters)
	End If

	SQLICFilters = "SELECT COL_LENGTH('IC_Filters', 'displayOrder') AS IsItThere"
	Set rsICFilters = cnnICFilters.Execute(SQLICFilters )
	If IsNull(rsICFilters("IsItThere")) Then
		SQLICFilters = "ALTER TABLE IC_Filters ADD displayOrder int NULL"
		Set rsICFilters = cnnICFilters.Execute(SQLICFilters)
		SQLICFilters = "UPDATE IC_Filters SET displayOrder = 0"
		Set rsICFilters = cnnICFilters.Execute(SQLICFilters)
	End If
	
	set rsICFilters = nothing
	cnnICFilters.close
	set cnnICFilters = nothing
%>