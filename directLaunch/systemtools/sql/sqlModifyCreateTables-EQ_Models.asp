<%	
	Set cnnCheckEQModels = Server.CreateObject("ADODB.Connection")
	cnnCheckEQModels.open (Session("ClientCnnString"))
	Set rsCheckEQModels = Server.CreateObject("ADODB.Recordset")


	on error goto 0

	SQLCheckEQModels  = "SELECT COL_LENGTH('EQ_Models', 'MightUseAFilter') AS IsItThere"
	Set rsCheckEQModels = cnnCheckEQModels.Execute(SQLCheckEQModels)
	If IsNull(rsCheckEQModels("IsItThere")) Then
		SQLCheckEQModels = "ALTER TABLE EQ_Models ADD MightUseAFilter int NULL"
		Set rsCheckEQModels = cnnCheckEQModels.Execute(SQLCheckEQModels)
		SQLCheckEQModels = "UPDATE EQ_Models SET MightUseAFilter = 0"
		Set rsCheckEQModels = cnnCheckEQModels.Execute(SQLCheckEQModels)
	End If

		
	set rsCheckEQModels = nothing
	cnnCheckEQModels.close
	set cnnCheckEQModels = nothing
				
%>