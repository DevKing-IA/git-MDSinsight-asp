<%	

	Set cnnPRProspectContactSearch = Server.CreateObject("ADODB.Connection")
	cnnPRProspectContactSearch.open (Session("ClientCnnString"))
	Set rsPRProspectContactSearch = Server.CreateObject("ADODB.Recordset")
	rsPRProspectContactSearch.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsPRProspectContactSearch = cnnPRProspectContactSearch.Execute("SELECT TOP 1 * FROM PR_ProspectContactSearch")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLPRProspectContactSearch = "CREATE TABLE [PR_ProspectContactSearch]( "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & "ProspectIntRecID int NOT NULL, "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & "Company varchar(255) NULL, "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & "City varchar(255) NULL, "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & "State varchar(20) NULL, "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & "FirstName varchar(255) NULL, "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & "LastName varchar(255) NULL "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & ") ON [PRIMARY]"
			
			Set rsPRProspectContactSearch = cnnPRProspectContactSearch.Execute(SQLPRProspectContactSearch)
			
			SQLPRProspectContactSearch = "INSERT INTO PR_ProspectContactSearch (ProspectIntRecID, Company, City, State, FirstName, LastName) "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & "SELECT PR_Prospects.InternalRecordIdentifier, PR_Prospects.Company, PR_Prospects.City, PR_Prospects.State, "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & "PR_ProspectContacts.FirstName, PR_ProspectContacts.LastName "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & "FROM PR_Prospects LEFT OUTER JOIN "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & "PR_ProspectContacts ON PR_ProspectContacts.ProspectIntRecID = PR_Prospects.InternalRecordIdentifier "
			SQLPRProspectContactSearch = SQLPRProspectContactSearch & "WHERE PR_Prospects.Pool = 'Live'"

			Set rsPRProspectContactSearch = cnnPRProspectContactSearch.Execute(SQLPRProspectContactSearch)
			
		End If
	End If
	
	
	set rsPRProspectContactSearch = nothing
	cnnPRProspectContactSearch.close
	set cnnPRProspectContactSearch = nothing
%>