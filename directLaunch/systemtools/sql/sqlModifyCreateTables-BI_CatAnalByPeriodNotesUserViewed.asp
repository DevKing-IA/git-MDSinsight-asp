<%	
	'*******************************************
	'A little different from other pages
	'this one actually renames an existing table
	'*******************************************

	Set cnnCompanyLeakage = Server.CreateObject("ADODB.Connection")
	cnnCompanyLeakage.open (Session("ClientCnnString"))
	Set rsCompanyLeakage = Server.CreateObject("ADODB.Recordset")
	rsCompanyLeakage.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsCompanyLeakage = cnnCompanyLeakage.Execute("SELECT TOP 1 * FROM AR_CustomerNotesUserViewed")
	
	If Err.Description = "" Then
	
		On error goto 0		

		'The table is there, so rename it
		
		SQLCompanyLeakage = "EXEC sp_rename  'AR_CustomerNotesUserViewed','AR_CustomerNotesUserViewed'"
		
		Set rsCompanyLeakage = cnnCompanyLeakage.Execute(SQLCompanyLeakage)
		
	End If
				
	set rsCompanyLeakage = nothing
	cnnCompanyLeakage.close
	set cnnCompanyLeakage = nothing
				
%>