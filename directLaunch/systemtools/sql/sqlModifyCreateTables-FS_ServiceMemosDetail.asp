<%	

	Set cnnCheckServiceMemosDetail = Server.CreateObject("ADODB.Connection")
	cnnCheckServiceMemosDetail.open (Session("ClientCnnString"))
	Set rsCheckServiceMemosDetail = Server.CreateObject("ADODB.Recordset")
	rsCheckServiceMemosDetail.CursorLocation = 3 
			

	SQL_CheckServiceMemosDetail = "SELECT COL_LENGTH('FS_ServiceMemosDetail', 'ProblemCode') AS IsItThere"
	Set rsCheckServiceMemosDetail = cnnCheckServiceMemosDetail.Execute(SQL_CheckServiceMemosDetail)
	If IsNull(rsCheckServiceMemosDetail("IsItThere")) Then
		SQL_CheckServiceMemosDetail  = "ALTER TABLE FS_ServiceMemosDetail ADD ProblemCode int NULL"
		Set rsCheckServiceMemosDetail = cnnCheckServiceMemosDetail.Execute(SQL_CheckServiceMemosDetail )
		' Set to 0 for Other
		SQL_CheckServiceMemosDetail  = "UPDATE FS_ServiceMemosDetail SET ProblemCode = 0"
		Set rsCheckServiceMemosDetail = cnnCheckServiceMemosDetail.Execute(SQL_CheckServiceMemosDetail )
	End If

	SQL_CheckServiceMemosDetail = "SELECT COL_LENGTH('FS_ServiceMemosDetail', 'ResolutionCode') AS IsItThere"
	Set rsCheckServiceMemosDetail = cnnCheckServiceMemosDetail.Execute(SQL_CheckServiceMemosDetail)
	If IsNull(rsCheckServiceMemosDetail("IsItThere")) Then
		SQL_CheckServiceMemosDetail  = "ALTER TABLE FS_ServiceMemosDetail ADD ResolutionCode int NULL"
		Set rsCheckServiceMemosDetail = cnnCheckServiceMemosDetail.Execute(SQL_CheckServiceMemosDetail )
		' Set to 0 for Other
		SQL_CheckServiceMemosDetail  = "UPDATE FS_ServiceMemosDetail SET ResolutionCode = 0"
		Set rsCheckServiceMemosDetail = cnnCheckServiceMemosDetail.Execute(SQL_CheckServiceMemosDetail )
	End If

	Set rsCheckServiceMemosDetail = Nothing
	cnnCheckServiceMemosDetail.Close
	Set cnnCheckServiceMemosDetail = Nothing
				
%>