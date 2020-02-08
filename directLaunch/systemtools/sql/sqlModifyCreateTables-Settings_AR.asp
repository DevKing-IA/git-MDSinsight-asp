<%	
	Set cnnSettings_AR = Server.CreateObject("ADODB.Connection")
	cnnSettings_AR.open (Session("ClientCnnString"))
	Set rsSettings_AR = Server.CreateObject("ADODB.Recordset")
	rsSettings_AR.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsSettings_AR = cnnSettings_AR.Execute("SELECT * FROM Settings_AR")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLSettings_AR = "CREATE TABLE [Settings_AR]( "
			SQLSettings_AR = SQLSettings_AR & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLSettings_AR = SQLSettings_AR & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_Settings_AR_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLSettings_AR = SQLSettings_AR & " [POST_Serno] [varchar](50) NULL,"
			SQLSettings_AR = SQLSettings_AR & " [POST_Mode] [varchar](50) NULL,"
			SQLSettings_AR = SQLSettings_AR & " [EmailForNon200Responses] [varchar](1000) NULL,"
			SQLSettings_AR = SQLSettings_AR & " [POST_CustomerURL1] [varchar](1000) NULL,"
			SQLSettings_AR = SQLSettings_AR & " [POST_CustomerURL2] [varchar](1000) NULL,"
			SQLSettings_AR = SQLSettings_AR & " [POST_CustomerURL1ONOFF] [int] NOT NULL DEFAULT 0, "
			SQLSettings_AR = SQLSettings_AR & " [POST_CustomerURL2ONOFF] [int] NOT NULL DEFAULT 0 "
			SQLSettings_AR = SQLSettings_AR & ") ON [PRIMARY]"
						
			Set rsSettings_AR = cnnSettings_AR.Execute(SQLSettings_AR)
			
			SQLSettings_AR = "INSERT INTO Settings_AR (POST_Serno,POST_Mode,EmailForNon200Responses,POST_CustomerURL1,POST_CustomerURL2,POST_CustomerURL1ONOFF,POST_CustomerURL2ONOFF) "
			SQLSettings_AR = SQLSettings_AR & " VALUES "
			SQLSettings_AR = SQLSettings_AR & " ('','TEST','insight@ocsaccess.com','http://apidev.mdsinsight.com/apiIn/receive_customer_xml.asp','',0,0)"
			Set rsSettings_AR = cnnSettings_AR.Execute(SQLSettings_AR)
		End If
	End If




	set rsSettings_AR = nothing
	cnnSettings_AR.close
	set cnnSettings_AR = nothing
%>