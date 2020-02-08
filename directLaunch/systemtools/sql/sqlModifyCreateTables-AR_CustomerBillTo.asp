<%	

	Set cnnCheckARCustBillTo = Server.CreateObject("ADODB.Connection")
	cnnCheckARCustBillTo.open (Session("ClientCnnString"))
	Set rsCheckARCustBillTo = Server.CreateObject("ADODB.Recordset")
	rsCheckARCustBillTo.CursorLocation = 3 
			
	'See if these fields are in the table& add them if not there

	SQL_CheckARCustBillTo = "SELECT COL_LENGTH('AR_CustomerBillTo', 'Phone') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQL_CheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD Phone varchar(255) NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	End If

	SQL_CheckARCustBillTo = "SELECT COL_LENGTH('AR_CustomerBillTo', 'Fax') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQL_CheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD Fax varchar(255) NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	End If
		
	SQL_CheckARCustBillTo = "SELECT COL_LENGTH('AR_CustomerBillTo', 'Email') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQL_CheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD Email varchar(255) NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	End If
	
	SQL_CheckARCustBillTo = "SELECT COL_LENGTH('AR_CustomerBillTo', 'Contact') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQL_CheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD Contact varchar(255) NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	End If

	SQL_CheckARCustBillTo = "SELECT COL_LENGTH('AR_CustomerBillTo', 'BackendBillToIDIfApplicable') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQL_CheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD BackendBillToIDIfApplicable varchar(255) NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	End If

	SQL_CheckARCustBillTo = "SELECT COL_LENGTH('AR_CustomerBillTo', 'Country') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQL_CheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD Country varchar(255) NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	End If

	SQL_CheckARCustBillTo = "SELECT COL_LENGTH('AR_CustomerBillTo', 'Description') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQL_CheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD Description varchar(255) NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	End If

	SQL_CheckARCustBillTo = "SELECT COL_LENGTH('AR_CustomerBillTo', 'DefaultBillTo') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQL_CheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD [DefaultBillTo] [int] NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
		SQL_CheckARCustBillTo = "UPDATE AR_CustomerBillTo SET DefaultBillTo = 0"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQL_CheckARCustBillTo)
	End If
	
	SQLCheckARCustBillTo  = "SELECT COL_LENGTH('AR_CustomerBillTo', 'ContactFirstName') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQLCheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQLCheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD ContactFirstName [varchar] (255) NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQLCheckARCustBillTo)
	End If
	
	SQLCheckARCustBillTo  = "SELECT COL_LENGTH('AR_CustomerBillTo', 'ContactLastName') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQLCheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQLCheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD ContactLastName [varchar] (255) NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQLCheckARCustBillTo)
	End If
	
	'This one is a DROP
	SQLCheckARCustBillTo = "SELECT COL_LENGTH('AR_CustomerBillTo', 'Enabled') AS IsItThere"
	Set rsCheckARCustBillTo  = cnnCheckARCustBillTo.Execute(SQLCheckARCustBillTo)
	If NOT IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQLCheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo DROP COLUMN Enabled"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQLCheckARCustBillTo)
	End If

	SQLCheckARCustBillTo  = "SELECT COL_LENGTH('AR_CustomerBillTo', 'Longitude') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQLCheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQLCheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD Longitude [varchar] (255) NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQLCheckARCustBillTo)
	End If
	SQLCheckARCustBillTo  = "SELECT COL_LENGTH('AR_CustomerBillTo', 'Latitude') AS IsItThere"
	Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQLCheckARCustBillTo)
	If IsNull(rsCheckARCustBillTo("IsItThere")) Then
		SQLCheckARCustBillTo = "ALTER TABLE AR_CustomerBillTo ADD Latitude [varchar] (255) NULL"
		Set rsCheckARCustBillTo = cnnCheckARCustBillTo.Execute(SQLCheckARCustBillTo)
	End If
	
	Set rsCheckARCustBillTo = Nothing
	cnnCheckARCustBillTo.Close
	Set cnnCheckARCustBillTo = Nothing
				
%>