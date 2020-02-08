<%	

	Set cnnCheckARCust = Server.CreateObject("ADODB.Connection")
	cnnCheckARCust.open (Session("ClientCnnString"))
	Set rsCheckARCust = Server.CreateObject("ADODB.Recordset")
	rsCheckARCust.CursorLocation = 3 
			
	'See if these fields are in the table& add them if not there
	
	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'MonthlyExpectedSalesDollars') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD MonthlyExpectedSalesDollars money NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
		SQL_CheckARCust = "UPDATE AR_Customer SET MonthlyExpectedSalesDollars = 0"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If
	
	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'MonthlyContractedSalesDollars') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD MonthlyContractedSalesDollars money NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
		SQL_CheckARCust = "UPDATE AR_Customer SET MonthlyContractedSalesDollars = 0"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If
	
	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'MaxMCSCharge') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD MaxMCSCharge money  NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If


	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'City') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD City varchar(255) NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'State') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD State varchar(255) NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'Zip') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD Zip varchar(255) NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'NewAcctSalesman') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD NewAcctSalesman INT NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'ContactFirstName') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD ContactFirstName varchar(255) NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'ContactLastName') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD ContactLastName varchar(255) NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'Longitude') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD Longitude varchar(255) NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'Latitude') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD Latitude varchar(255) NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'MCSEnrollmentDate') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD MCSEnrollmentDate datetime NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
		SQL_CheckARCust = "UPDATE AR_Customer SET MCSEnrollmentDate = '2018-01-01 08:00:00.000'"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'MESEnrollmentDate') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD MESEnrollmentDate datetime NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
		SQL_CheckARCust = "UPDATE AR_Customer SET MESEnrollmentDate = '2018-01-01 09:00:00.000'"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If
	
	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'PriceLevel') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD PriceLevel INT NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If
	
	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'PORequired') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD PORequired INT NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If
	
	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'BlanketPONumber') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD BlanketPONumber varchar(255) NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If
	
	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'TermsIntRecID') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD TermsIntRecID INT NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'RecordCreationDateTime') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD RecordCreationDateTime datetime NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
		SQL_CheckARCust = "UPDATE AR_Customer SET RecordCreationDateTime = getdate()"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'InternalRecordIdentifier') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD InternalRecordIdentifier INT IDENTITY(1,1) NOT NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If
	
	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'Country') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD Country varchar(255) NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If
	
	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'LastPriceChangeDate') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer ADD LastPriceChangeDate datetime NULL"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	'This one is a drop
	SQL_CheckARCust = "SELECT COL_LENGTH('AR_Customer', 'Archived') AS IsItThere"
	Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	If NOT IsNull(rsCheckARCust("IsItThere")) Then
		SQL_CheckARCust = "ALTER TABLE AR_Customer DROP COLUMN Archived"
		Set rsCheckARCust = cnnCheckARCust.Execute(SQL_CheckARCust)
	End If

	Set rsCheckARCust = Nothing
	cnnCheckARCust.Close
	Set cnnCheckARCust = Nothing
				
%>