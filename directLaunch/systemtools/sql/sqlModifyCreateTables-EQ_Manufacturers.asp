<%	
	Set cnnCheckEQManufacturers = Server.CreateObject("ADODB.Connection")
	cnnCheckEQManufacturers.open (Session("ClientCnnString"))
	Set rsCheckEQManufacturers = Server.CreateObject("ADODB.Recordset")

	on error goto 0

	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'Description') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD [Description] [varchar] (255) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If

	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'ExtraDescription1') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD ExtraDescription1 [varchar] (255) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If

	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'ExtraDescription2') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD ExtraDescription2 [varchar] (255) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If

	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'ExtraDescription3') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD ExtraDescription3 [varchar] (255) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If
		
	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'DriverNotes') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD DriverNotes [varchar] (8000) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If

	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'Country') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD Country [varchar] (255) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If

	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'TaxIDNumber') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD TaxIDNumber [varchar] (255) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If

	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'PrimarySalesperson') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD PrimarySalesperson [int] NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If

	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'SecondarySalesPerson') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD SecondarySalesPerson [int] NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If

	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'DefaultBillToIntRecID') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD DefaultBillToIntRecID [int] NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If

	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'BackendShipToIDIfApplicable') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD BackendShipToIDIfApplicable varchar(255) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If
	
	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'DefaultShipTo') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD [DefaultShipTo] [int] NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
		SQLCheckEQManufacturers = "UPDATE AR_CustomerShipTo SET DefaultShipTo= 0"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If
	
	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'ContactFirstName') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD ContactFirstName [varchar] (255) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If
	
	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'ContactLastName') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD ContactLastName [varchar] (255) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If
	
	'This one is a DROP
	SQLCheckEQManufacturers = "SELECT COL_LENGTH('AR_CustomerShipTo', 'DefaultBillToIntRecID') AS IsItThere"
	Set rsCheckEQManufacturers  = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If NOT IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo DROP COLUMN DefaultBillToIntRecID"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If
	
	'This one is a DROP
	SQLCheckEQManufacturers = "SELECT COL_LENGTH('AR_CustomerShipTo', 'Enabled') AS IsItThere"
	Set rsCheckEQManufacturers  = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If NOT IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo DROP COLUMN Enabled"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If
	
	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'Longitude') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD Longitude [varchar] (255) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If
	
	SQLCheckEQManufacturers  = "SELECT COL_LENGTH('AR_CustomerShipTo', 'Latitude') AS IsItThere"
	Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	If IsNull(rsCheckEQManufacturers("IsItThere")) Then
		SQLCheckEQManufacturers = "ALTER TABLE AR_CustomerShipTo ADD Latitude [varchar] (255) NULL"
		Set rsCheckEQManufacturers = cnnCheckEQManufacturers.Execute(SQLCheckEQManufacturers)
	End If
	
	
	set rsCheckEQManufacturers = nothing
	cnnCheckEQManufacturers.close
	set cnnCheckEQManufacturers = nothing
				
%>