<%	
	Set cnnFS_ServiceMemosNotes = Server.CreateObject("ADODB.Connection")
	cnnFS_ServiceMemosNotes.open (Session("ClientCnnString"))
	Set rsFS_ServiceMemosNotes = Server.CreateObject("ADODB.Recordset")
	rsFS_ServiceMemosNotes.CursorLocation = 3 

	Set cnnCheckFS_ServiceMemos = Server.CreateObject("ADODB.Connection")
	cnnCheckFS_ServiceMemos.open (Session("ClientCnnString"))
	Set rsCheckFS_ServiceMemos = Server.CreateObject("ADODB.Recordset")
	rsCheckFS_ServiceMemos.CursorLocation = 3 

	Set cnnUpdateFS_ServiceMemos = Server.CreateObject("ADODB.Connection")
	cnnUpdateFS_ServiceMemos.open (Session("ClientCnnString"))
	Set rsUpdateFS_ServiceMemos = Server.CreateObject("ADODB.Recordset")
	rsUpdateFS_ServiceMemos.CursorLocation = 3 

	'*******************************************************************************************************************************************
	'CREATE THE FS_ServiceMemosNotes TABLE
	'*******************************************************************************************************************************************
	Err.Clear
	on error resume next
	Set rsFS_ServiceMemosNotes = cnnFS_ServiceMemosNotes.Execute("SELECT TOP 1 * FROM FS_ServiceMemosNotes ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLFS_ServiceMemosNotes = "CREATE TABLE [FS_ServiceMemosNotes]( "
			SQLFS_ServiceMemosNotes = SQLFS_ServiceMemosNotes & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLFS_ServiceMemosNotes = SQLFS_ServiceMemosNotes & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_FS_ServiceMemosNotes]  DEFAULT (getdate()), "
			SQLFS_ServiceMemosNotes = SQLFS_ServiceMemosNotes & " [ServiceTicketID] [varchar](255) NULL, "
			SQLFS_ServiceMemosNotes = SQLFS_ServiceMemosNotes & " [EnteredByUserNo] [int] NULL, "
			SQLFS_ServiceMemosNotes = SQLFS_ServiceMemosNotes & " [Note] [varchar](8000) NULL "
			SQLFS_ServiceMemosNotes = SQLFS_ServiceMemosNotes & ") ON [PRIMARY]"
		
			Set rsFS_ServiceMemosNotes = cnnFS_ServiceMemosNotes.Execute(SQLFS_ServiceMemosNotes)
		End If
	End If
	
	
	'*******************************************************************************************************************************************
	'GET ALL THE ServiceNotesFromTech THAT WERE STORED IN FS_ServiceMemos AND MOVE IT INTO FS_ServiceMemosNotes
	'MAKE SURE TO COPY THE DATE THAT THE NOTE WAS ENTERED INTO FS_ServiceMemosNotes AS WELL
	'THEN SET THE ORIGINAL SERVICE NOTES (ServiceNotesFromTech) IN FS_ServiceMemos TO EMPTY
	'*******************************************************************************************************************************************
	
	Set rsCheckFS_ServiceMemos = cnnFS_ServiceMemosNotes.Execute("SELECT * FROM FS_ServiceMemos WHERE ServiceNotesFromTech <> '' ORDER BY RecordCreatedateTime DESC")
	
	If NOT rsCheckFS_ServiceMemos.EOF Then
	
		DO WHILE NOT rsCheckFS_ServiceMemos.EOF
		
		
			ServiceMemoRecNumber = rsCheckFS_ServiceMemos("ServiceMemoRecNumber")
			origServiceNotesFromFSServiceMemos = rsCheckFS_ServiceMemos("ServiceNotesFromTech")
			origDateEnteredFromFSServiceMemos = rsCheckFS_ServiceMemos("RecordCreatedateTime")
			origServiceTicketIDFromFSServiceMemos = rsCheckFS_ServiceMemos("MemoNumber")
			origUserNoFromFSServiceMemos = rsCheckFS_ServiceMemos("UserNoOfServiceTech")
			
			SQLUpdateFS_ServiceMemosNotes = "INSERT INTO FS_ServiceMemosNotes(RecordCreationDateTime, ServiceTicketID, EnteredByUserNo, Note) "
			SQLUpdateFS_ServiceMemosNotes = SQLUpdateFS_ServiceMemosNotes & " VALUES "
			SQLUpdateFS_ServiceMemosNotes = SQLUpdateFS_ServiceMemosNotes & " ('" & origDateEnteredFromFSServiceMemos & "', "
			SQLUpdateFS_ServiceMemosNotes = SQLUpdateFS_ServiceMemosNotes & " '" & origServiceTicketIDFromFSServiceMemos & "', "
			SQLUpdateFS_ServiceMemosNotes = SQLUpdateFS_ServiceMemosNotes & " " & origUserNoFromFSServiceMemos & ", "
			SQLUpdateFS_ServiceMemosNotes = SQLUpdateFS_ServiceMemosNotes & " '" & origServiceNotesFromFSServiceMemos & "') "

			'Response.Write("SQLUpdateFS_ServiceMemosNotes (" & ClientKey & ") : " & SQLUpdateFS_ServiceMemosNotes & "<br>")
					
			'********************************************************************************************
			'INSERT SERVICE TICKET NOTES INTO FS_ServiceMemosNotes
			'********************************************************************************************
			Set rsFS_ServiceMemosNotes = cnnFS_ServiceMemosNotes.Execute(SQLUpdateFS_ServiceMemosNotes)
	
			'********************************************************************************************
			'UPDATE SERVICE TICKET NOTES FROM FS_ServiceMemos TO EMPTY
			'********************************************************************************************
			Set rsUpdateFS_ServiceMemos = cnnUpdateFS_ServiceMemos.Execute("UPDATE FS_ServiceMemos SET ServiceNotesFromTech = '' WHERE ServiceMemoRecNumber = " & ServiceMemoRecNumber)
			
		
			rsCheckFS_ServiceMemos.MoveNext
			
		LOOP
		
	End If


	Set rsCheckFS_ServiceMemos = Nothing
	cnnCheckFS_ServiceMemos.Close
	Set cnnCheckFS_ServiceMemos = Nothing

	Set rsUpdateFS_ServiceMemos = Nothing
	cnnUpdateFS_ServiceMemos.Close
	Set cnnUpdateFS_ServiceMemos = Nothing
	
	set rsFS_ServiceMemosNotes = nothing
	cnnFS_ServiceMemosNotes.close
	set cnnFS_ServiceMemosNotes = nothing
				
%>