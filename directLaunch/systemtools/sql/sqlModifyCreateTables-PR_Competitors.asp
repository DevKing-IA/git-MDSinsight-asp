<%	

	Set cnnCheckPR_Competitors = Server.CreateObject("ADODB.Connection")
	cnnCheckPR_Competitors.open (Session("ClientCnnString"))
	Set rsCheckPR_Competitors = Server.CreateObject("ADODB.Recordset")
	rsCheckPR_Competitors.CursorLocation = 3 

	SQL_PR_Competitors = "SELECT COL_LENGTH('PR_Competitors', 'CompetitorWebsite') AS IsItThere"
	Set rsCheckPR_Competitors = cnnCheckPR_Competitors.Execute(SQL_PR_Competitors)
	If IsNull(rsCheckPR_Competitors("IsItThere")) Then
		SQL_PR_Competitors = "ALTER TABLE PR_Competitors ADD CompetitorWebsite [varchar](255) NULL "
		Set rsCheckPR_Competitors = cnnCheckPR_Competitors.Execute(SQL_PR_Competitors)
	End If

	SQL_PR_Competitors = "SELECT COL_LENGTH('PR_Competitors', 'AdditionalNotes') AS IsItThere"
	Set rsCheckPR_Competitors = cnnCheckPR_Competitors.Execute(SQL_PR_Competitors)
	If IsNull(rsCheckPR_Competitors("IsItThere")) Then
		SQL_PR_Competitors = "ALTER TABLE PR_Competitors ADD AdditionalNotes [varchar](8000) NULL "
		Set rsCheckPR_Competitors = cnnCheckPR_Competitors.Execute(SQL_PR_Competitors)
	End If

	Set rsCheckPR_Competitors = Nothing
	cnnCheckPR_Competitors.Close
	Set cnnCheckPR_Competitors = Nothing
%>