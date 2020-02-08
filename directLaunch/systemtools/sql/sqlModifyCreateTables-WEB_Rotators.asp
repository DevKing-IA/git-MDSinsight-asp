<%	

	Set cnnWEB_Rotators = Server.CreateObject("ADODB.Connection")
	cnnWEB_Rotators.open (Session("ClientCnnString"))
	Set rsWEB_Rotators = Server.CreateObject("ADODB.Recordset")
	rsWEB_Rotators.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_Rotators = cnnWEB_Rotators.Execute("SELECT TOP 1 * FROM WEB_Rotators")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLWEB_Rotators = "CREATE TABLE [WEB_Rotators]( "
			SQLWEB_Rotators = SQLWEB_Rotators & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLWEB_Rotators = SQLWEB_Rotators & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_Rotators]  DEFAULT (getdate()), "
			SQLWEB_Rotators = SQLWEB_Rotators & " [RotatorID] [int] NOT NULL, "
			SQLWEB_Rotators = SQLWEB_Rotators & " [RotatorFileName] [varchar](255) NULL, "
			SQLWEB_Rotators = SQLWEB_Rotators & " [RotatorLink] [varchar](255) NULL, "
			SQLWEB_Rotators = SQLWEB_Rotators & " [RotatorSequence] [int] NULL, "
			SQLWEB_Rotators = SQLWEB_Rotators & " [RotatorNewWindow] [bit] NOT NULL DEFAULT 0, "
			SQLWEB_Rotators = SQLWEB_Rotators & " [RotatorStart] [datetime] NULL, "
			SQLWEB_Rotators = SQLWEB_Rotators & " [RotatorEnd] [datetime] NULL, "
			SQLWEB_Rotators = SQLWEB_Rotators & " [RotatorDisplaySeconds] [int] NULL "
			SQLWEB_Rotators = SQLWEB_Rotators & ") ON [PRIMARY]"
			Set rsWEB_Rotators = cnnWEB_Rotators.Execute(SQLWEB_Rotators)
		End If
	End If
	
	
	set rsWEB_Rotators = nothing
	cnnWEB_Rotators.close
	set cnnWEB_Rotators = nothing


%>