<%	

	Set cnnCheckICPartners = Server.CreateObject("ADODB.Connection")
	cnnCheckICPartners.open (Session("ClientCnnString"))
	Set rsCheckICPartners = Server.CreateObject("ADODB.Recordset")
	rsCheckICPartners.CursorLocation = 3 
	
	SQL_CheckICPartners = "SELECT COL_LENGTH('IC_Partners', 'partnerRejectsBlankProdDescs') AS IsItThere"
	Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
	If IsNull(rsCheckICPartners("IsItThere")) Then
		SQL_CheckICPartners = "ALTER TABLE IC_Partners ADD partnerRejectsBlankProdDescs bit NOT NULL DEFAULT 1"
		Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
		SQL_CheckICPartners = "UPDATE IC_Partners SET partnerRejectsBlankProdDescs = 1"
		Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
	End If
	
	SQL_CheckICPartners = "SELECT COL_LENGTH('IC_Partners', 'partnerRejectsBlankProdUOMS') AS IsItThere"
	Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
	If IsNull(rsCheckICPartners("IsItThere")) Then
		SQL_CheckICPartners = "ALTER TABLE IC_Partners ADD partnerRejectsBlankProdUOMS bit NOT NULL DEFAULT 1"
		Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
		SQL_CheckICPartners = "UPDATE IC_Partners SET partnerRejectsBlankProdUOMS = 1"
		Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
	End If

	SQL_CheckICPartners = "SELECT COL_LENGTH('IC_Partners', 'partnerUseShipIDForMappingLookup') AS IsItThere"
	Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
	If IsNull(rsCheckICPartners("IsItThere")) Then
		SQL_CheckICPartners = "ALTER TABLE IC_Partners ADD partnerUseShipIDForMappingLookup int NULL"
		Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
		SQL_CheckICPartners = "UPDATE IC_Partners SET partnerUseShipIDForMappingLookup = 0"
		Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
	End If

	SQL_CheckICPartners = "SELECT COL_LENGTH('IC_Partners', 'Longitude') AS IsItThere"
	Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
	If IsNull(rsCheckICPartners("IsItThere")) Then
		SQL_CheckICPartners = "ALTER TABLE IC_Partners ADD Longitude varchar (255) NULL"
		Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
	End If

	SQL_CheckICPartners = "SELECT COL_LENGTH('IC_Partners', 'Latitude') AS IsItThere"
	Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
	If IsNull(rsCheckICPartners("IsItThere")) Then
		SQL_CheckICPartners = "ALTER TABLE IC_Partners ADD Latitude varchar (255) NULL"
		Set rsCheckICPartners = cnnCheckICPartners.Execute(SQL_CheckICPartners)
	End If
	
	Set rsCheckICPartners = Nothing
	cnnCheckICPartners.Close
	Set cnnCheckICPartners = Nothing
				
%>