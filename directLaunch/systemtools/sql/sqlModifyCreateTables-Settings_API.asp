<%	
	Set cnnSettings_API = Server.CreateObject("ADODB.Connection")
	cnnSettings_API.open (Session("ClientCnnString"))
	Set rsSettings_API = Server.CreateObject("ADODB.Recordset")
	rsSettings_API.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsSettings_API = cnnSettings_API.Execute("SELECT * FROM Settings_API")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLSettings_API = "CREATE TABLE [Settings_API]( "
			SQLSettings_API = SQLSettings_API & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLSettings_API = SQLSettings_API & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_Settings_API_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLSettings_API = SQLSettings_API & " [Schedule_DailyAPIActivityByPartnerReportGeneration] [varchar](1000) NULL"
			SQLSettings_API = SQLSettings_API & ") ON [PRIMARY]"
						
			Set rsSettings_API = cnnSettings_API.Execute(SQLSettings_API)
			
			SQLSettings_API = "INSERT INTO Settings_API (Schedule_DailyAPIActivityByPartnerReportGeneration) "
			SQLSettings_API = SQLSettings_API & " VALUES "
			SQLSettings_API = SQLSettings_API & " ('0,0,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0')"
			Set rsSettings_API = cnnSettings_API.Execute(SQLSettings_API)
		End If
	End If


	set rsSettings_API = nothing
	cnnSettings_API.close
	set cnnSettings_API = nothing
%>