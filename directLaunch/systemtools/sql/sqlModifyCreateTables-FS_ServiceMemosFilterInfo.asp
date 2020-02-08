<%	
	Set cnnCheckFSServiceMemosFilterInfo = Server.CreateObject("ADODB.Connection")
	cnnCheckFSServiceMemosFilterInfo.open (Session("ClientCnnString"))
	Set rsCheckFSServiceMemosFilterInfo = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckFSServiceMemosFilterInfo = cnnCheckFSServiceMemosFilterInfo.Execute("SELECT TOP 1 * FROM FS_ServiceMemosFilterInfo")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckFSServiceMemosFilterInfo = "CREATE TABLE [FS_ServiceMemosFilterInfo]("
			SQLCheckFSServiceMemosFilterInfo = SQLCheckFSServiceMemosFilterInfo & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckFSServiceMemosFilterInfo = SQLCheckFSServiceMemosFilterInfo & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_FS_ServiceMemosFilterInfo_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckFSServiceMemosFilterInfo = SQLCheckFSServiceMemosFilterInfo & " [CustID] [varchar](255) NULL,"
			SQLCheckFSServiceMemosFilterInfo = SQLCheckFSServiceMemosFilterInfo & " [ServiceTicketID] [varchar](255) NULL,"
			SQLCheckFSServiceMemosFilterInfo = SQLCheckFSServiceMemosFilterInfo & " [CustFilterIntRecID] [int] NULL,"
			SQLCheckFSServiceMemosFilterInfo = SQLCheckFSServiceMemosFilterInfo & " [ICFilterIntRecID] [int] NULL,"			
			SQLCheckFSServiceMemosFilterInfo = SQLCheckFSServiceMemosFilterInfo & " [Completed] [int] NULL,"						
			SQLCheckFSServiceMemosFilterInfo = SQLCheckFSServiceMemosFilterInfo & " [CompletedDate] [datetime] NULL,"						
			SQLCheckFSServiceMemosFilterInfo = SQLCheckFSServiceMemosFilterInfo & " [CompletedByUserNo] [int]"						
			SQLCheckFSServiceMemosFilterInfo = SQLCheckFSServiceMemosFilterInfo & " ) ON [PRIMARY]"      

		   Set rsCheckFSServiceMemosFilterInfo = cnnCheckFSServiceMemosFilterInfo.Execute(SQLCheckFSServiceMemosFilterInfo)
		   
		End If
	End If
	
	set rsCheckFSServiceMemosFilterInfo = nothing
	cnnCheckFSServiceMemosFilterInfo.close
	set cnnCheckFSServiceMemosFilterInfo = nothing
		
%>