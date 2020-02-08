<%	

	Set cnnSettings_CompanyCalendar = Server.CreateObject("ADODB.Connection")
	cnnSettings_CompanyCalendar.open (Session("ClientCnnString"))
	Set rsSettings_CompanyCalendar = Server.CreateObject("ADODB.Recordset")
	rsSettings_CompanyCalendar.CursorLocation = 3 
			
	'See if these fields are in the table& add them if not there
	
	SQL_Settings_CompanyCalendar = "SELECT COL_LENGTH('Settings_CompanyCalendar', 'AlternateDeliveryDate') AS IsItThere"
	Set rsSettings_CompanyCalendar = cnnSettings_CompanyCalendar.Execute(SQL_Settings_CompanyCalendar)
	If IsNull(rsSettings_CompanyCalendar("IsItThere")) Then
		SQL_Settings_CompanyCalendar = "ALTER TABLE Settings_CompanyCalendar ADD AlternateDeliveryDate datetime NULL"
		Set rsSettings_CompanyCalendar = cnnSettings_CompanyCalendar.Execute(SQL_Settings_CompanyCalendar)
	End If

	' If there are norecords in here at all, put in Christmas of the
	' current year so there is at least one entry
	SQL_Settings_CompanyCalendar = "SELECT COUNT (*) AS GlobalCount FROM Settings_CompanyCalendar"
	Set rsSettings_CompanyCalendar = cnnSettings_CompanyCalendar.Execute(SQL_Settings_CompanyCalendar)

	If rsSettings_CompanyCalendar("GlobalCount") < 1 Then
		SQL_Settings_CompanyCalendar = "INSERT INTO Settings_CompanyCalendar (MonthNam, MonthNum, DayNum, YearNum, OpenClosedCloseEarly,  Description) "
		SQL_Settings_CompanyCalendar = SQL_Settings_CompanyCalendar & " VALUES "
		SQL_Settings_CompanyCalendar = SQL_Settings_CompanyCalendar & " ("
		SQL_Settings_CompanyCalendar = SQL_Settings_CompanyCalendar & "'Dec'"
		SQL_Settings_CompanyCalendar = SQL_Settings_CompanyCalendar & ", 12"
		SQL_Settings_CompanyCalendar = SQL_Settings_CompanyCalendar & ", 25"
		SQL_Settings_CompanyCalendar = SQL_Settings_CompanyCalendar & ", " & Year(Now())
		SQL_Settings_CompanyCalendar = SQL_Settings_CompanyCalendar & ", 'Closed'"
		SQL_Settings_CompanyCalendar = SQL_Settings_CompanyCalendar & ", 'Christmas Day'"
		SQL_Settings_CompanyCalendar = SQL_Settings_CompanyCalendar & " )"
		
		Set rsSettings_CompanyCalendar = cnnSettings_CompanyCalendar.Execute(SQL_Settings_CompanyCalendar)
		
	End If

	
	Set rsSettings_CompanyCalendar = Nothing
	cnnSettings_CompanyCalendar.Close
	Set cnnSettings_CompanyCalendar = Nothing
				
%>