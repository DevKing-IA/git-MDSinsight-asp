<%	
Response.Write("sql/sqlModifyCreateTables-PR_ProspectSocialMedia.asp<br>")

	Set cnnProspectsSocialMedia = Server.CreateObject("ADODB.Connection")
	cnnProspectsSocialMedia.open (Session("ClientCnnString"))
	Set rsProspectsSocialMedia = Server.CreateObject("ADODB.Recordset")
	rsProspectsSocialMedia.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsProspectsSocialMedia = cnnProspectsSocialMedia.Execute("SELECT TOP 1 * FROM PR_ProspectSocialMedia")
	
	If Err.Description <> "" Then

		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then

			On error goto 0		

			'The table is not there, we need to create it
			SQLProspectsSocialMedia = "CREATE TABLE [PR_ProspectSocialMedia]( "
			
			SQLProspectsSocialMedia = SQLProspectsSocialMedia & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLProspectsSocialMedia = SQLProspectsSocialMedia & "[RecordCreationDateTime] [datetime] NULL, "
			SQLProspectsSocialMedia = SQLProspectsSocialMedia & "[ProspectIntRecID] [int] NULL, "
			SQLProspectsSocialMedia = SQLProspectsSocialMedia & "[SocialMediaPlatform] [varchar](255) NULL, "
			SQLProspectsSocialMedia = SQLProspectsSocialMedia & "[SocialMediaLink] [varchar](1000) NULL "
			SQLProspectsSocialMedia = SQLProspectsSocialMedia& ") ON [PRIMARY]"
			
			Set rsProspectsSocialMedia= cnnProspectsSocialMedia.Execute(SQLProspectsSocialMedia)
			


			SQLProspectsSocialMedia = "CREATE CLUSTERED INDEX [IX_PR_ProspectSocialMedia] ON [PR_ProspectSocialMedia] "
			SQLProspectsSocialMedia = SQLProspectsSocialMedia& "( [ProspectIntRecID] ASC "
			SQLProspectsSocialMedia = SQLProspectsSocialMedia& ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, "
			SQLProspectsSocialMedia = SQLProspectsSocialMedia& "DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"

			Set rsProspectsSocialMedia= cnnProspectsSocialMedia.Execute(SQLProspectsSocialMedia)
			
			

			SQLProspectsSocialMedia = "ALTER TABLE PR_ProspectSocialMedia ADD CONSTRAINT [DF_PR_ProspectSocialMedia_RecordCreationDateTime]  DEFAULT (getdate()) FOR [RecordCreationDateTime]"

			Set rsProspectsSocialMedia= cnnProspectsSocialMedia.Execute(SQLProspectsSocialMedia)
		
		End If
	End If
	
	
	set rsProspectsSocialMedia= nothing
	cnnProspectsSocialMedia.close
	set cnnProspectsSocialMedia= nothing
%>