<%	
	Response.Write("sqlModifyCreateTables-SC_SchedulerLog.asp" & "<br>")
	On Error Goto 0

	Set cnnCheckSCSchedulerLog = Server.CreateObject("ADODB.Connection")
	cnnCheckSCSchedulerLog.open (Session("ClientCnnString"))
	Set rsCheckSCSchedulerLog = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckSCSchedulerLog = cnnCheckSCSchedulerLog.Execute("SELECT TOP 1 * FROM SC_SchedulerLog")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckSCSchedulerLog = "CREATE TABLE [SC_SchedulerLog]("
			SQLCheckSCSchedulerLog = SQLCheckSCSchedulerLog & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckSCSchedulerLog = SQLCheckSCSchedulerLog & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_SC_SchedulerLog_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckSCSchedulerLog = SQLCheckSCSchedulerLog & " [pageName] [varchar](8000) NULL "
			SQLCheckSCSchedulerLog = SQLCheckSCSchedulerLog & " ) ON [PRIMARY]"      
		   Set rsCheckSCSchedulerLog = cnnCheckSCSchedulerLog.Execute(SQLCheckSCSchedulerLog)
		   
		End If
	End If


	set rsCheckSCSchedulerLog = nothing
	cnnCheckSCSchedulerLog.close
	set cnnCheckSCSchedulerLog = nothing
				
%>