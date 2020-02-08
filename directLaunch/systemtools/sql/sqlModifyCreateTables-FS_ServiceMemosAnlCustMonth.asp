<%	
	Set cnnCheckFS_ServiceMemosAnlCustMonth = Server.CreateObject("ADODB.Connection")
	cnnCheckFS_ServiceMemosAnlCustMonth.open (Session("ClientCnnString"))
	Set rsCheckFS_ServiceMemosAnlCustMonth = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckFS_ServiceMemosAnlCustMonth = cnnCheckFS_ServiceMemosAnlCustMonth.Execute("SELECT TOP 1 * FROM FS_ServiceMemosAnlCustMonth")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckFS_ServiceMemosAnlCustMonth = "CREATE TABLE [FS_ServiceMemosAnlCustMonth]("
			SQLCheckFS_ServiceMemosAnlCustMonth = SQLCheckFS_ServiceMemosAnlCustMonth & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckFS_ServiceMemosAnlCustMonth = SQLCheckFS_ServiceMemosAnlCustMonth & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_FS_ServiceMemosAnlCustMonth_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckFS_ServiceMemosAnlCustMonth = SQLCheckFS_ServiceMemosAnlCustMonth & " [CustID] [varchar](255) NULL,"
			SQLCheckFS_ServiceMemosAnlCustMonth = SQLCheckFS_ServiceMemosAnlCustMonth & " [TicketMonth] [int] NULL,"
			SQLCheckFS_ServiceMemosAnlCustMonth = SQLCheckFS_ServiceMemosAnlCustMonth & " [TicketYear] [int] NULL,"
			SQLCheckFS_ServiceMemosAnlCustMonth = SQLCheckFS_ServiceMemosAnlCustMonth & " [Period] [int] NULL,"			
			SQLCheckFS_ServiceMemosAnlCustMonth = SQLCheckFS_ServiceMemosAnlCustMonth & " [PeriodYear] [int] NULL,"						
			SQLCheckFS_ServiceMemosAnlCustMonth = SQLCheckFS_ServiceMemosAnlCustMonth & " [NumberOfServiceTickets] [int] NULL,"						
			SQLCheckFS_ServiceMemosAnlCustMonth = SQLCheckFS_ServiceMemosAnlCustMonth & " ) ON [PRIMARY]"      

		   Set rsCheckFS_ServiceMemosAnlCustMonth = cnnCheckFS_ServiceMemosAnlCustMonth.Execute(SQLCheckFS_ServiceMemosAnlCustMonth)
		   
		End If
	End If
	
	set rsCheckFS_ServiceMemosAnlCustMonth = nothing
	cnnCheckFS_ServiceMemosAnlCustMonth.close
	set cnnCheckFS_ServiceMemosAnlCustMonth = nothing
		
%>