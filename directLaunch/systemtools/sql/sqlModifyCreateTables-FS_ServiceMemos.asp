<%	

	Set cnnCheckServiceMemos = Server.CreateObject("ADODB.Connection")
	cnnCheckServiceMemos.open (Session("ClientCnnString"))
	Set rsCheckServiceMemos = Server.CreateObject("ADODB.Recordset")
	rsCheckServiceMemos.CursorLocation = 3 
			

	SQL_CheckServiceMemos = "SELECT COL_LENGTH('FS_ServiceMemos', 'ServiceNotesFromTech') AS IsItThere"
	Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos)
	If IsNull(rsCheckServiceMemos("IsItThere")) Then
		SQL_CheckServiceMemos  = "ALTER TABLE FS_ServiceMemos ADD ServiceNotesFromTech varchar(8000) NULL"
		Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos )
	End If
	
	
	SQL_CheckServiceMemos = "SELECT COL_LENGTH('FS_ServiceMemos', 'ServiceNotesFromTech') AS IsItThere"
	Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos)
	If IsNull(rsCheckServiceMemos("IsItThere")) Then
		SQL_CheckServiceMemos = "ALTER TABLE FS_ServiceMemos ADD ServiceNotesFromTech varchar(8000) NULL"
		Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos)
	End If
	
	SQL_CheckServiceMemos = "SELECT COL_LENGTH('FS_ServiceMemos', 'CancellationNotes') AS IsItThere"
	Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos)
	If IsNull(rsCheckServiceMemos("IsItThere")) Then
		SQL_CheckServiceMemos  = "ALTER TABLE FS_ServiceMemos ADD CancellationNotes varchar(8000) NULL"
		Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos )
	End If

	SQL_CheckServiceMemos = "SELECT COL_LENGTH('FS_ServiceMemos', 'ProblemCode') AS IsItThere"
	Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos)
	If IsNull(rsCheckServiceMemos("IsItThere")) Then
		SQL_CheckServiceMemos  = "ALTER TABLE FS_ServiceMemos ADD ProblemCode int NULL"
		Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos )
		' Set to 0 for Other
		SQL_CheckServiceMemos  = "UPDATE FS_ServiceMemos SET ProblemCode = 0"
		Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos )
	End If

	SQL_CheckServiceMemos = "SELECT COL_LENGTH('FS_ServiceMemos', 'HoldReason') AS IsItThere"
	Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos)
	If IsNull(rsCheckServiceMemos("IsItThere")) Then
		SQL_CheckServiceMemos  = "ALTER TABLE FS_ServiceMemos ADD HoldReason varchar(255) NULL"
		Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos )
	End If

	SQL_CheckServiceMemos = "SELECT COL_LENGTH('FS_ServiceMemos', 'SymptomCode') AS IsItThere"
	Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos)
	If IsNull(rsCheckServiceMemos("IsItThere")) Then
		SQL_CheckServiceMemos  = "ALTER TABLE FS_ServiceMemos ADD SymptomCode int NULL"
		Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos )
		' Set to 0 for Other
		SQL_CheckServiceMemos  = "UPDATE FS_ServiceMemos SET SymptomCode = 0"
		Set rsCheckServiceMemos = cnnCheckServiceMemos.Execute(SQL_CheckServiceMemos )
	End If

	Set rsCheckServiceMemos = Nothing
	cnnCheckServiceMemos.Close
	Set cnnCheckServiceMemos = Nothing
				
%>