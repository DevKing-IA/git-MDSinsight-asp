<%	
	Set cnnFS_ServiceMemosNotesUserViewed = Server.CreateObject("ADODB.Connection")
	cnnFS_ServiceMemosNotesUserViewed.open (Session("ClientCnnString"))
	Set rsFS_ServiceMemosNotesUserViewed = Server.CreateObject("ADODB.Recordset")
	rsFS_ServiceMemosNotesUserViewed.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsFS_ServiceMemosNotesUserViewed = cnnFS_ServiceMemosNotesUserViewed.Execute("SELECT TOP 1 * FROM FS_ServiceMemosNotesUserViewed ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLFS_ServiceMemosNotesUserViewed = "CREATE TABLE [FS_ServiceMemosNotesUserViewed]( "
			SQLFS_ServiceMemosNotesUserViewed = SQLFS_ServiceMemosNotesUserViewed & " [UserNo] [int] NULL, "
			SQLFS_ServiceMemosNotesUserViewed = SQLFS_ServiceMemosNotesUserViewed & " [ServiceTicketID] [varchar](255) NULL, "
			SQLFS_ServiceMemosNotesUserViewed = SQLFS_ServiceMemosNotesUserViewed & " [DateLastViewed] [datetime] NULL "
			SQLFS_ServiceMemosNotesUserViewed = SQLFS_ServiceMemosNotesUserViewed & ") ON [PRIMARY]"
		
			Set rsFS_ServiceMemosNotesUserViewed = cnnFS_ServiceMemosNotesUserViewed.Execute(SQLFS_ServiceMemosNotesUserViewed)
		End If
	End If

	
	set rsFS_ServiceMemosNotesUserViewed = nothing
	cnnFS_ServiceMemosNotesUserViewed.close
	set cnnFS_ServiceMemosNotesUserViewed = nothing
				
%>