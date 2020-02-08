<%	

	Set cnnCheckAPVendor = Server.CreateObject("ADODB.Connection")
	cnnCheckAPVendor.open (Session("ClientCnnString"))
	Set rsCheckAPVendor = Server.CreateObject("ADODB.Recordset")
	rsCheckAPVendor.CursorLocation = 3 
			
	'See if these fields are in the table& add them if not there

	SQL_CheckAPVendor = "SELECT COL_LENGTH('AP_Vendor', 'Website') AS IsItThere"
	Set rsCheckAPVendor = cnnCheckAPVendor.Execute(SQL_CheckAPVendor)
	If IsNull(rsCheckAPVendor("IsItThere")) Then
		SQL_CheckAPVendor = "ALTER TABLE AP_Vendor ADD Website varchar(255) NULL"
		Set rsCheckAPVendor = cnnCheckAPVendor.Execute(SQL_CheckAPVendor)
	End If

	SQL_CheckAPVendor = "SELECT COL_LENGTH('AP_Vendor', 'AccountNumber') AS IsItThere"
	Set rsCheckAPVendor = cnnCheckAPVendor.Execute(SQL_CheckAPVendor)
	If IsNull(rsCheckAPVendor("IsItThere")) Then
		SQL_CheckAPVendor = "ALTER TABLE AP_Vendor ADD AccountNumber varchar(255) NULL"
		Set rsCheckAPVendor = cnnCheckAPVendor.Execute(SQL_CheckAPVendor)
	End If
		
	SQL_CheckAPVendor = "SELECT COL_LENGTH('AP_Vendor', 'Notes') AS IsItThere"
	Set rsCheckAPVendor = cnnCheckAPVendor.Execute(SQL_CheckAPVendor)
	If IsNull(rsCheckAPVendor("IsItThere")) Then
		SQL_CheckAPVendor = "ALTER TABLE AP_Vendor ADD Notes varchar(8000) NULL"
		Set rsCheckAPVendor = cnnCheckAPVendor.Execute(SQL_CheckAPVendor)
	End If
	
	Set rsCheckAPVendor = Nothing
	cnnCheckAPVendor.Close
	Set cnnCheckAPVendor = Nothing
				
%>