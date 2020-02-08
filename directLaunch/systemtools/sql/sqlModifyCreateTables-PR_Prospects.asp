<%	

	Set cnnCheckPR_Prospects = Server.CreateObject("ADODB.Connection")
	cnnCheckPR_Prospects.open (Session("ClientCnnString"))
	Set rsCheckPR_Prospects = Server.CreateObject("ADODB.Recordset")
	rsCheckPR_Prospects.CursorLocation = 3 

	' To change the size of the field
	SQL_CheckPR_Prospects = "SELECT COL_LENGTH ('PR_Prospects','Company') AS ColLen "
	Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	If rsCheckPR_Prospects("ColLen") <> 255 Then
		SQL_CheckPR_Prospects  = "ALTER TABLE PR_Prospects ALTER COLUMN Company varchar(255) NULL"
		Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	End If

	' To change the size of the field
	SQL_CheckPR_Prospects = "SELECT COL_LENGTH ('PR_Prospects','Street') AS ColLen "
	Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	If rsCheckPR_Prospects("ColLen") <> 255 Then
		SQL_CheckPR_Prospects  = "ALTER TABLE PR_Prospects ALTER COLUMN Street varchar(255) NULL"
		Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	End If

	' To change the size of the field
	SQL_CheckPR_Prospects = "SELECT COL_LENGTH ('PR_Prospects','City') AS ColLen "
	Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	If rsCheckPR_Prospects("ColLen") <> 255 Then
		SQL_CheckPR_Prospects  = "ALTER TABLE PR_Prospects ALTER COLUMN City varchar(255) NULL"
		Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	End If

	' To change the size of the field
	SQL_CheckPR_Prospects = "SELECT COL_LENGTH ('PR_Prospects','Website') AS ColLen "
	Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	If rsCheckPR_Prospects("ColLen") <> 255 Then
		SQL_CheckPR_Prospects  = "ALTER TABLE PR_Prospects ALTER COLUMN Website varchar(255) NULL"
		Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	End If

	' To change the size of the field
	SQL_CheckPR_Prospects = "SELECT COL_LENGTH ('PR_Prospects','Floor_Suite_Room__c') AS ColLen "
	Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	If rsCheckPR_Prospects("ColLen") <> 255 Then
		SQL_CheckPR_Prospects  = "ALTER TABLE PR_Prospects ALTER COLUMN Floor_Suite_Room__c varchar(255) NULL"
		Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	End If


	SQL_CheckPR_Prospects = "SELECT COL_LENGTH ('PR_Prospects','LastVerifiedDate') AS IsItThere"
	Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	If IsNull(rsCheckPR_Prospects("IsItThere")) Then
		SQL_CheckPR_Prospects  = "ALTER TABLE Pr_Prospects ADD LastVerifiedDate datetime NULL"
		Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	End If

	SQL_CheckPR_Prospects = "SELECT COL_LENGTH ('PR_Prospects','LastVerifiedDate') AS IsItThere"
	Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	If IsNull(rsCheckPR_Prospects("IsItThere")) Then
		SQL_CheckPR_Prospects  = "ALTER TABLE Pr_Prospects ADD LastVerifiedDate datetime NULL"
		Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	End If

	SQL_CheckPR_Prospects = "SELECT COL_LENGTH ('PR_Prospects','Longitude') AS IsItThere"
	Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	If IsNull(rsCheckPR_Prospects("IsItThere")) Then
		SQL_CheckPR_Prospects  = "ALTER TABLE Pr_Prospects ADD Longitude varchar(255) NULL"
		Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	End If

	SQL_CheckPR_Prospects = "SELECT COL_LENGTH ('PR_Prospects','Latitude') AS IsItThere"
	Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	If IsNull(rsCheckPR_Prospects("IsItThere")) Then
		SQL_CheckPR_Prospects  = "ALTER TABLE Pr_Prospects ADD Latitude varchar(255) NULL"
		Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)
	End If

	' Handle records where this didn't get set upon saving
	SQL_CheckPR_Prospects  = "UPDATE Pr_Prospects SET LastVerifiedDate = RecordCreationDateTime WHERE LastVerifiedDate IS NULL"
	Set rsCheckPR_Prospects = cnnCheckPR_Prospects.Execute(SQL_CheckPR_Prospects)

	Set rsCheckPR_Prospects = Nothing
	cnnCheckPR_Prospects.Close
	Set cnnCheckPR_Prospects = Nothing
%>