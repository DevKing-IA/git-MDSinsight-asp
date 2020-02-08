<%	

	Set cnnCheckCompany = Server.CreateObject("ADODB.Connection")
	cnnCheckCompany.open (Session("ClientCnnString"))
	Set rsCheckCompany = Server.CreateObject("ADODB.Recordset")
	rsCheckCompany.CursorLocation = 3 
			
	'See if these fields are in the table& add them if not there
	
	SQL_CheckCompany = "SELECT COL_LENGTH('Settings_CompanyID', 'Timezone') AS IsItThere"
	Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	If IsNull(rsCheckCompany("IsItThere")) Then
		SQL_CheckCompany = "ALTER TABLE Settings_CompanyID ADD Timezone varchar(255) NULL"
		Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
		SQL_CheckCompany = "UPDATE Settings_CompanyID SET Timezone = 'Eastern'"
		Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	End If
	
	SQL_CheckCompany = "SELECT COL_LENGTH('Settings_CompanyID', 'BusinessDayStart') AS IsItThere"
	Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	If IsNull(rsCheckCompany("IsItThere")) Then
		SQL_CheckCompany = "ALTER TABLE Settings_CompanyID ADD BusinessDayStart varchar(255) NULL"
		Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
		SQL_CheckCompany = "UPDATE Settings_CompanyID SET BusinessDayStart = '8:00'"
		Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	End If
	
	SQL_CheckCompany = "SELECT COL_LENGTH('Settings_CompanyID', 'BusinessDayEnd') AS IsItThere"
	Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	If IsNull(rsCheckCompany("IsItThere")) Then
		SQL_CheckCompany = "ALTER TABLE Settings_CompanyID ADD BusinessDayEnd varchar(255) NULL"
		Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
		SQL_CheckCompany = "UPDATE Settings_CompanyID SET BusinessDayEnd = '17:00'"
		Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	End If

	SQL_CheckCompany = "SELECT COL_LENGTH('Settings_CompanyID', 'Stmt_Country') AS IsItThere"
	Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	If IsNull(rsCheckCompany("IsItThere")) Then
		SQL_CheckCompany = "ALTER TABLE Settings_CompanyID ADD Stmt_Country varchar(255) NULL"
		Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	End If

	SQL_CheckCompany = "SELECT COL_LENGTH('Settings_CompanyID', 'PeriodsOrMonths') AS IsItThere"
	Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	If IsNull(rsCheckCompany("IsItThere")) Then
		SQL_CheckCompany = "ALTER TABLE Settings_CompanyID ADD PeriodsOrMonths varchar(1) NULL"
		Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
		SQL_CheckCompany = "UPDATE Settings_CompanyID SET PeriodsOrMonths = 'M'"
		Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	End If

	SQL_CheckCompany = "SELECT COL_LENGTH('Settings_CompanyID', 'PointOfServiceLogicOnOff') AS IsItThere"
	Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	If IsNull(rsCheckCompany("IsItThere")) Then
		SQL_CheckCompany = "ALTER TABLE Settings_CompanyID ADD PointOfServiceLogicOnOff int NULL"
		Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
		SQL_CheckCompany = "UPDATE Settings_CompanyID SET PointOfServiceLogicOnOff = 0"
		Set rsCheckCompany = cnnCheckCompany.Execute(SQL_CheckCompany)
	End If
	
	Set rsCheckCompany = Nothing
	cnnCheckCompany.Close
	Set cnnCheckCompany = Nothing
				
%>