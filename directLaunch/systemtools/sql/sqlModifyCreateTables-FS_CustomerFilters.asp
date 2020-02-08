<%	
	Set cnnCheckFSCustomerFilters = Server.CreateObject("ADODB.Connection")
	cnnCheckFSCustomerFilters.open (Session("ClientCnnString"))
	Set rsCheckFSCustomerFilters = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckFSCustomerFilters = cnnCheckFSCustomerFilters.Execute("SELECT TOP 1 * FROM FS_CustomerFilters")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckFSCustomerFilters = "CREATE TABLE [FS_CustomerFilters]("
			SQLCheckFSCustomerFilters = SQLCheckFSCustomerFilters & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckFSCustomerFilters = SQLCheckFSCustomerFilters & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_FS_CustomerFilters_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckFSCustomerFilters = SQLCheckFSCustomerFilters & " [CustID] [varchar](255) NULL,"
			SQLCheckFSCustomerFilters = SQLCheckFSCustomerFilters & " [prodSKU] [varchar](255) NULL,"
			SQLCheckFSCustomerFilters = SQLCheckFSCustomerFilters & " [FrequencyType] [varchar](255) NULL,"
			SQLCheckFSCustomerFilters = SQLCheckFSCustomerFilters & " [FrequencyTime] [int] NULL,"			
			SQLCheckFSCustomerFilters = SQLCheckFSCustomerFilters & " [Price] [money] NULL,"						
			SQLCheckFSCustomerFilters = SQLCheckFSCustomerFilters & " [Location] [varchar](1000) NULL,"						
			SQLCheckFSCustomerFilters = SQLCheckFSCustomerFilters & " [LastChangeDateTime] [datetime]"						
			SQLCheckFSCustomerFilters = SQLCheckFSCustomerFilters & " ) ON [PRIMARY]"      

		   Set rsCheckFSCustomerFilters = cnnCheckFSCustomerFilters.Execute(SQLCheckFSCustomerFilters)
		   
		End If
	End If

	SQLCheckFSCustomerFilters  = "SELECT COL_LENGTH('FS_CustomerFilters]', 'Qty') AS IsItThere"
	Set rsCheckFSCustomerFilters = cnnCheckFSCustomerFilters.Execute(SQLCheckFSCustomerFilters )
	
	If IsNull(rsCheckFSCustomerFilters("IsItThere")) Then
		SQLCheckFSCustomerFilters = "ALTER TABLE FS_CustomerFilters ADD Qty INT NULL"
		Set rsCheckFSCustomerFilters= cnnCheckFSCustomerFilters.Execute(SQLCheckFSCustomerFilters)
		SQLCheckFSCustomerFilters = "UPDATE FS_CustomerFilters SET Qty = 1"
		Set rsCheckFSCustomerFilters = cnnCheckFSCustomerFilters.Execute(SQLCheckFSCustomerFilters)
	End If

on error goto 0	
	'This code will rename a column
	SQLCheckFSCustomerFilters = "SELECT COL_LENGTH('FS_CustomerFilters', 'Location') AS IsItThere"
	Set rsCheckFSCustomerFilters = cnnCheckFSCustomerFilters.Execute(SQLCheckFSCustomerFilters )
	If NOT IsNull(rsCheckFSCustomerFilters("IsItThere")) Then
		SQLCheckFSCustomerFilters = "sp_rename 'FS_CustomerFilters.Location','Notes','COLUMN'"
		Set rsCheckFSCustomerFilters= cnnCheckFSCustomerFilters.Execute(SQLCheckFSCustomerFilters )
	End If



	'Special code to handle the fact that we changed tables structure after data had
	'already been created in the table
	SQL_CheckFSCustomerFilters = "SELECT COL_LENGTH('FS_CustomerFilters', 'FilterIntRecID') AS IsItThere"
	Set rsCheckFSCustomerFilters = cnnCheckFSCustomerFilters.Execute(SQL_CheckFSCustomerFilters)
	
	If IsNull(rsCheckFSCustomerFilters("IsItThere")) Then
	
		SQL_CheckFSCustomerFilters = "ALTER TABLE FS_CustomerFilters ADD FilterIntRecID int NULL"
		Set rsCheckFSCustomerFilters = cnnCheckFSCustomerFilters.Execute(SQL_CheckFSCustomerFilters)
		
		'Only if the prodsku is there
		SQL_CheckFSCustomerFilters = "SELECT COL_LENGTH('FS_CustomerFilters', 'prodSKU') AS IsItThere"
		Set rsCheckFSCustomerFilters = cnnCheckFSCustomerFilters.Execute(SQL_CheckFSCustomerFilters)
		If NOT IsNull(rsCheckFSCustomerFilters("IsItThere")) Then
	
			'Now lookup & update FilterIntRecID based on the filter SKU
			SQL_CheckFSCustomerFilters = "UPDATE FS_CustomerFilters "
			SQL_CheckFSCustomerFilters = SQL_CheckFSCustomerFilters & "SET FS_CustomerFilters.FilterIntRecID = IC_Filters.InternalRecordIdentifier "
			SQL_CheckFSCustomerFilters = SQL_CheckFSCustomerFilters & "FROM FS_CustomerFilters INNER JOIN "
			SQL_CheckFSCustomerFilters = SQL_CheckFSCustomerFilters & "IC_Filters ON FS_CustomerFilters.prodSKU = IC_Filters.FilterID "
	
			Set rsCheckFSCustomerFilters = cnnCheckFSCustomerFilters.Execute(SQL_CheckFSCustomerFilters)
	
			
			'Now drop the prodSKU field fro the table
			SQL_CheckFSCustomerFilters  = "ALTER TABLE FS_CustomerFilters DROP COLUMN prodSKU"
			Set rsCheckFSCustomerFilters = cnnCheckFSCustomerFilters.Execute(SQL_CheckFSCustomerFilters)
			
			'Cleanup orphan records
			SQL_CheckFSCustomerFilters  = "DELETE FROM FS_CustomerFilters WHERE FilterIntRecID IS NULL"
'''''		Set rsCheckFSCustomerFilters = cnnCheckFSCustomerFilters.Execute(SQL_CheckFSCustomerFilters)
			
		End If

	End If

	
	set rsCheckFSCustomerFilters = nothing
	cnnCheckFSCustomerFilters.close
	set cnnCheckFSCustomerFilters = nothing
				
%>