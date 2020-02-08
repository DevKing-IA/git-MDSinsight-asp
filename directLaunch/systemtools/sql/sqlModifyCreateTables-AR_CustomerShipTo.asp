<%	
	Set cnnCheckARCustShipTo = Server.CreateObject("ADODB.Connection")
	cnnCheckARCustShipTo.open (Session("ClientCnnString"))
	Set rsCheckARCustShipTo = Server.CreateObject("ADODB.Recordset")


	Err.Clear
	on error resume next
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute("SELECT TOP 1 * FROM AR_CustomerShipTo")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckARCustShipTo = "CREATE TABLE [AR_CustomerShipTo]("
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_CustomerShipTo_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [CustNum] [varchar](255) NULL,"
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [ShipName] [varchar](255) NULL,"
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [Addr1] [varchar](255) NULL,"
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [Addr2] [varchar](255) NULL,"
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [City] [varchar](255) NULL,"
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [State] [varchar](50) NULL,"
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [Zip] [varchar](50) NULL,"	
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [Phone] [varchar](255) NULL,"
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [Fax] [varchar](50) NULL,"
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [Email] [varchar](255) NULL,"
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " [Contact] [varchar](255) NULL" 
			SQLCheckARCustShipTo = SQLCheckARCustShipTo & " ) ON [PRIMARY]"      

		   Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
		   
		End If
	End If

on error goto 0

	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'Description') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD [Description] [varchar] (255) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If

	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'ExtraDescription1') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD ExtraDescription1 [varchar] (255) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If

	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'ExtraDescription2') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD ExtraDescription2 [varchar] (255) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If

	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'ExtraDescription3') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD ExtraDescription3 [varchar] (255) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If
		
	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'DriverNotes') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD DriverNotes [varchar] (8000) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If

	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'Country') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD Country [varchar] (255) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If

	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'TaxIDNumber') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD TaxIDNumber [varchar] (255) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If

	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'PrimarySalesperson') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD PrimarySalesperson [int] NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If

	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'SecondarySalesPerson') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD SecondarySalesPerson [int] NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If


	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'BackendShipToIDIfApplicable') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD BackendShipToIDIfApplicable varchar(255) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If
	
	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'DefaultShipTo') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD [DefaultShipTo] [int] NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
		SQLCheckARCustShipTo = "UPDATE AR_CustomerShipTo SET DefaultShipTo= 0"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If
	
	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'ContactFirstName') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD ContactFirstName [varchar] (255) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If
	
	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'ContactLastName') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD ContactLastName [varchar] (255) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If
	
	'This one is a DROP
	SQLCheckARCustShipTo = "SELECT COL_LENGTH('AR_CustomerShipTo', 'DefaultBillToIntRecID') AS IsItThere"
	Set rsCheckARCustShipTo  = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If NOT IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo DROP COLUMN DefaultBillToIntRecID"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If
	
	'This one is a DROP
	SQLCheckARCustShipTo = "SELECT COL_LENGTH('AR_CustomerShipTo', 'Enabled') AS IsItThere"
	Set rsCheckARCustShipTo  = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If NOT IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo DROP COLUMN Enabled"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If
	
	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'Longitude') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD Longitude [varchar] (255) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If
	
	SQLCheckARCustShipTo  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'Latitude') AS IsItThere"
	Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo ADD Latitude [varchar] (255) NULL"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If

	'This one is a DROP
	SQLCheckARCustShipTo = "SELECT COL_LENGTH('AR_CustomerShipTo', 'ShipToID') AS IsItThere"
	Set rsCheckARCustShipTo  = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	If NOT IsNull(rsCheckARCustShipTo("IsItThere")) Then
		SQLCheckARCustShipTo = "ALTER TABLE AR_CustomerShipTo DROP COLUMN ShipToID"
		Set rsCheckARCustShipTo = cnnCheckARCustShipTo.Execute(SQLCheckARCustShipTo)
	End If

	set rsCheckARCustShipTo = nothing
	cnnCheckARCustShipTo.close
	set cnnCheckARCustShipTo = nothing
				
%>