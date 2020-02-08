<%	

	Set cnnCheckProspectContacts = Server.CreateObject("ADODB.Connection")
	cnnCheckProspectContacts.open (Session("ClientCnnString"))
	Set rsCheckProspectContacts = Server.CreateObject("ADODB.Recordset")
	rsCheckProspectContacts.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute("SELECT TOP 1 * FROM PR_ProspectContacts")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLCheckProspectContacts = "CREATE TABLE [PR_ProspectContacts]( "
			SQLCheckProspectContacts = SQLCheckProspectContacts & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[ProspectIntRecID] [int] NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[Suffix] [varchar](50) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[FirstName] [varchar](255) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[LastName] [varchar](255) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[ContactTitleNumber] [int] NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[Email] [varchar](255) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[Phone] [varchar](255) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[PhoneExt] [varchar](255) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[Cell] [varchar](255) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[Fax] [varchar](255) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[DecisionMaker] [bit] NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[PrimaryContact] [bit] NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[Notes] [varchar](8000) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[DoNotEmail] [bit] NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[RecordCreationDate] [datetime] NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[Address1] [varchar](255) NULL, " 
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[Address2] [varchar](255) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[City] [varchar](255) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[State] [varchar](255) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[PostalCode] [varchar](255) NULL, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & "[Country] [varchar](255) NULL "
			SQLCheckProspectContacts = SQLCheckProspectContacts & ") ON [PRIMARY]"
			Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)

			'Create indexes
			SQLCheckProspectContacts = "CREATE NONCLUSTERED INDEX [IX_PR_ProspectContacts] ON [PR_ProspectContacts]"
			SQLCheckProspectContacts = SQLCheckProspectContacts & " ([InternalRecordIdentifier] ASC "
			SQLCheckProspectContacts = SQLCheckProspectContacts & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & " DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
			Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)

			
			SQLCheckProspectContacts = "CREATE NONCLUSTERED INDEX [IX_PR_ProspectContacts_1] ON [PR_ProspectContacts]"
			SQLCheckProspectContacts = SQLCheckProspectContacts & " (	[InternalRecordIdentifier] ASC,	[ProspectIntRecID] ASC "
			SQLCheckProspectContacts = SQLCheckProspectContacts & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF"
			SQLCheckProspectContacts = SQLCheckProspectContacts & ", ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
			Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)


			SQLCheckProspectContacts = "CREATE NONCLUSTERED INDEX [IX_PR_ProspectContacts_2] ON [PR_ProspectContacts]"
			SQLCheckProspectContacts = SQLCheckProspectContacts & " ("
			SQLCheckProspectContacts = SQLCheckProspectContacts & " 	[LastName] ASC"
			SQLCheckProspectContacts = SQLCheckProspectContacts & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & " SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
			Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)


			SQLCheckProspectContacts = "CREATE NONCLUSTERED INDEX [IX_PR_ProspectContacts_3] ON [PR_ProspectContacts]"
			SQLCheckProspectContacts = SQLCheckProspectContacts & " ("
			SQLCheckProspectContacts = SQLCheckProspectContacts & " 	[Phone] ASC"
			SQLCheckProspectContacts = SQLCheckProspectContacts & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF,"
			SQLCheckProspectContacts = SQLCheckProspectContacts & "  SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
			Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)


			SQLCheckProspectContacts = "CREATE NONCLUSTERED INDEX [IX_PR_ProspectContacts_4] ON [PR_ProspectContacts]"
			SQLCheckProspectContacts = SQLCheckProspectContacts & " ("
			SQLCheckProspectContacts = SQLCheckProspectContacts & " 	[FirstName] ASC"
			SQLCheckProspectContacts = SQLCheckProspectContacts & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, "
			SQLCheckProspectContacts = SQLCheckProspectContacts & " DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
			Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
			
		End If
	End If
	
	On Error Goto 0
	
	Set cnnCheckProspectContacts = Server.CreateObject("ADODB.Connection")
	cnnCheckProspectContacts.open (Session("ClientCnnString"))
	Set rsCheckProspectContacts = Server.CreateObject("ADODB.Recordset")

	SQLCheckProspectContacts  = "SELECT COL_LENGTH('Pr_ProspectContacts', 'Address1') AS IsItThere"
	Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	If IsNull(rsCheckProspectContacts("IsItThere")) Then
		SQLCheckProspectContacts = "ALTER TABLE Pr_ProspectContacts ADD Address1 varchar(255) NULL"
		Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	End If

	SQLCheckProspectContacts  = "SELECT COL_LENGTH('Pr_ProspectContacts', 'Address2') AS IsItThere"
	Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	If IsNull(rsCheckProspectContacts("IsItThere")) Then
		SQLCheckProspectContacts = "ALTER TABLE Pr_ProspectContacts ADD Address2 varchar(255) NULL"
		Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	End If

	SQLCheckProspectContacts  = "SELECT COL_LENGTH('Pr_ProspectContacts', 'City') AS IsItThere"
	Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	If IsNull(rsCheckProspectContacts("IsItThere")) Then
		SQLCheckProspectContacts = "ALTER TABLE Pr_ProspectContacts ADD City varchar(255) NULL"
		Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	End If

	SQLCheckProspectContacts  = "SELECT COL_LENGTH('Pr_ProspectContacts', 'State') AS IsItThere"
	Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	If IsNull(rsCheckProspectContacts("IsItThere")) Then
		SQLCheckProspectContacts = "ALTER TABLE Pr_ProspectContacts ADD [State] varchar(255) NULL"
		Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	End If

	SQLCheckProspectContacts  = "SELECT COL_LENGTH('Pr_ProspectContacts', 'PostalCode') AS IsItThere"
	Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	If IsNull(rsCheckProspectContacts("IsItThere")) Then
		SQLCheckProspectContacts = "ALTER TABLE Pr_ProspectContacts ADD PostalCode varchar(255) NULL"
		Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	End If

	SQLCheckProspectContacts  = "SELECT COL_LENGTH('Pr_ProspectContacts', 'Country') AS IsItThere"
	Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	If IsNull(rsCheckProspectContacts("IsItThere")) Then
		SQLCheckProspectContacts = "ALTER TABLE Pr_ProspectContacts ADD Country varchar(255) NULL"
		Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	End If

	SQLCheckProspectContacts  = "SELECT COL_LENGTH('Pr_ProspectContacts', 'Longitude') AS IsItThere"
	Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	If IsNull(rsCheckProspectContacts("IsItThere")) Then
		SQLCheckProspectContacts = "ALTER TABLE Pr_ProspectContacts ADD Longitude varchar(255) NULL"
		Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	End If

	SQLCheckProspectContacts  = "SELECT COL_LENGTH('Pr_ProspectContacts', 'Latitude') AS IsItThere"
	Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	If IsNull(rsCheckProspectContacts("IsItThere")) Then
		SQLCheckProspectContacts = "ALTER TABLE Pr_ProspectContacts ADD Latitude varchar(255) NULL"
		Set rsCheckProspectContacts = cnnCheckProspectContacts.Execute(SQLCheckProspectContacts)
	End If
		
	set rsCheckProspectContacts = nothing
	cnnCheckProspectContacts.close
	set cnnCheckProspectContacts = nothing
%>