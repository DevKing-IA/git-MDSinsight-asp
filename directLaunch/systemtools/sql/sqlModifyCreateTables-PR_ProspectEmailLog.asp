<%	
	On Error Goto 0
	If lcase(Request.ServerVariables("HTTPS")) = "on" Then 
	    strProtocol = "https" 
	Else
	    strProtocol = "http" 
	End If
	strDomain= Request.ServerVariables("SERVER_NAME")
	strPath= Request.ServerVariables("SCRIPT_NAME") 
	strQueryString= Request.ServerVariables("QUERY_STRING")
	strFullUrl = strProtocol & "://" & strDomain & strPath
	If Len(strQueryString) > 0 Then
	   strFullUrl = strFullUrl & "?" & strQueryString
	End If
	Response.Write "</br>" & strFullUrl & "</br>"

	Set cnnProspectsEmailLog = Server.CreateObject("ADODB.Connection")
	cnnProspectsEmailLog.open (Session("ClientCnnString"))
	Set rsProspectsEmailLog = Server.CreateObject("ADODB.Recordset")
	rsProspectsEmailLog.CursorLocation = 3 

	If Year(now()) = 2019 AND DAY(now()) = 18 AND month(now()) = 5 Then
		SQLProspectsEmailLog = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PR_ProspectEmailLog')"
	    SQLProspectsEmailLog = SQLProspectsEmailLog & "DROP TABLE PR_ProspectEmailLog;"
	    Response.Write(SQLProspectsEmailLog & "<br>")
		Set rsProspectsEmailLog = cnnProspectsEmailLog.Execute(SQLProspectsEmailLog )
	End If


	Err.Clear
	on error resume next
	Set rsProspectsEmailLog = cnnProspectsEmailLog.Execute("SELECT TOP 1 * FROM PR_ProspectEmailLog")
	
	If Err.Description <> "" Then

		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then

			On error goto 0		

			'The table is not there, we need to create it
			SQLProspectsEmailLog = "CREATE TABLE [PR_ProspectEmailLog]( "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[RecordCreationDateTime] [datetime] NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[ProspectRecID] [int] NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[msg_id] [varchar](200) NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[emaildatetime] [datetime] NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[attach_count] [int] NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[msg_foldername] [varchar](50) NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[from_addr] [varchar](200) NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[to_addr] [varchar](4000) NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[cc_addr] [varchar](4000) NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[bcc_addr] [varchar](4000) NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[sub] [varchar](4000) NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[body_html] [varchar](8000) NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[body_text] [varchar](8000) NULL, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "[sticky] [bit] NULL "
			SQLProspectsEmailLog = SQLProspectsEmailLog& ") ON [PRIMARY]"
			
			Set rsProspectsEmailLog= cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)
			

			SQLProspectsEmailLog = "ALTER TABLE PR_ProspectEmailLog ADD CONSTRAINT [DF_PR_ProspectEmailLog_RecordCreationDateTime]  DEFAULT (getdate()) FOR [RecordCreationDateTime]"

			Set rsProspectsEmailLog= cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)

			SQLProspectsEmailLog = "ALTER TABLE PR_ProspectEmailLog ADD CONSTRAINT [DF_PR_ProspectEmailLog_sticky]  DEFAULT ((0)) FOR [sticky]"

			Set rsProspectsEmailLog= cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)


			SQLProspectsEmailLog = "CREATE NONCLUSTERED INDEX [IX_PR_ProspectsEmailLog_1] ON [PR_ProspectEmailLog] "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "( [msg_id] ASC "
			SQLProspectsEmailLog = SQLProspectsEmailLog & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"

			Set rsProspectsEmailLog= cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)
			

			SQLProspectsEmailLog = "CREATE CLUSTERED INDEX [IX_PR_ProspectsEmailLog_2] ON [PR_ProspectEmailLog] "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "( [ProspectRecID] ASC, [emaildatetime] DESC "
			SQLProspectsEmailLog = SQLProspectsEmailLog & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, "
			SQLProspectsEmailLog = SQLProspectsEmailLog & "DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"

			Set rsProspectsEmailLog= cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)
		
		End If
	End If
	

	SQLProspectsEmailLog  = "SELECT COL_LENGTH ('PR_ProspectEmailLog','from_name') AS IsItThere"
	Set rsProspectsEmailLog = cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)
	If IsNull(rsProspectsEmailLog("IsItThere")) Then
		SQLProspectsEmailLog = "ALTER TABLE Pr_ProspectEmailLog ADD from_name varchar(1000) NULL"
		Set rsProspectsEmailLog = cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)
	End If

	SQLProspectsEmailLog  = "SELECT COL_LENGTH ('PR_ProspectEmailLog','to_name') AS IsItThere"
	Set rsProspectsEmailLog = cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)
	If IsNull(rsProspectsEmailLog("IsItThere")) Then
		SQLProspectsEmailLog = "ALTER TABLE Pr_ProspectEmailLog ADD to_name varchar(1000) NULL"
		Set rsProspectsEmailLog = cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)
	End If

	SQLProspectsEmailLog  = "SELECT COL_LENGTH ('PR_ProspectEmailLog','cc_name') AS IsItThere"
	Set rsProspectsEmailLog = cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)
	If IsNull(rsProspectsEmailLog("IsItThere")) Then
		SQLProspectsEmailLog = "ALTER TABLE Pr_ProspectEmailLog ADD cc_name varchar(1000) NULL"
		Set rsProspectsEmailLog = cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)
	End If

	SQLProspectsEmailLog  = "SELECT COL_LENGTH ('PR_ProspectEmailLog','HideFromUserNos') AS IsItThere"
	Set rsProspectsEmailLog = cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)
	If IsNull(rsProspectsEmailLog("IsItThere")) Then
		SQLProspectsEmailLog = "ALTER TABLE Pr_ProspectEmailLog ADD HideFromUserNos varchar(1000) NULL"
		Set rsProspectsEmailLog = cnnProspectsEmailLog.Execute(SQLProspectsEmailLog)
	End If

	
	set rsProspectsEmailLog= nothing
	cnnProspectsEmailLog.close
	set cnnProspectsEmailLog= nothing
%>