<%	

	Set cnnWEB_Tracking = Server.CreateObject("ADODB.Connection")
	cnnWEB_Tracking.open (Session("ClientCnnString"))
	Set rsWEB_Tracking = Server.CreateObject("ADODB.Recordset")
	rsWEB_Tracking.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_Tracking = cnnWEB_Tracking.Execute("SELECT TOP 1 * FROM WEB_Tracking")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLWEB_Tracking = "CREATE TABLE [WEB_Tracking]( "
			SQLWEB_Tracking = SQLWEB_Tracking & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLWEB_Tracking = SQLWEB_Tracking & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_Tracking]  DEFAULT (getdate()), "
			SQLWEB_Tracking = SQLWEB_Tracking & " [trkUserNo] [int] NULL "
			SQLWEB_Tracking = SQLWEB_Tracking & ") ON [PRIMARY]"
			Set rsWEB_Tracking = cnnWEB_Tracking.Execute(SQLWEB_Tracking)
		End If
	End If
	
	
	set rsWEB_Tracking = nothing
	cnnWEB_Tracking.close
	set cnnWEB_Tracking = nothing


%>