<%	
	Set cnnIN_WebFulfillment = Server.CreateObject("ADODB.Connection")
	cnnIN_WebFulfillment.open (Session("ClientCnnString"))
	Set rsIN_WebFulfillment = Server.CreateObject("ADODB.Recordset")
	rsIN_WebFulfillment.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsIN_WebFulfillment = cnnIN_WebFulfillment.Execute("SELECT * FROM IN_WebFulfillment")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLIN_WebFulfillment = "CREATE TABLE [IN_WebFulfillment]( "
			SQLWEB_WebFulfillment = SQLWEB_WebFulfillment & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLWEB_WebFulfillment = SQLWEB_WebFulfillment & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_Tracking]  DEFAULT (getdate()), "

			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [OCSAccessOrderID] [varchar](50) NULL, "
			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [OCSAccessOrderDate] [datetime] NULL, "
			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [CustID] [varchar](50) NULL, "
			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [CustClassCode] [varchar](50) NULL, "
			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [CustTypeNum] [int] NULL, "
			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [MDSInvoiceID] [varchar](50) NULL, "
			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [MDSInvoiceDate] [datetime] NULL, "
			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [OCSAccessMerchTotal] [money] NULL, "
			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [MDSInvoiceTotal] [money] NULL, "
			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [DontIncludeOnReport] [int] NOT NULL DEFAULT 0, "
			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [Remarks] [varchar](8000) NULL,"
			SQLIN_WebFulfillment = SQLIN_WebFulfillment & " [OCSAccessOrderComments] [varchar](8000) NULL "

			SQLIN_WebFulfillment = SQLIN_WebFulfillment & ") ON [PRIMARY]"
						
			Set rsIN_WebFulfillment = cnnIN_WebFulfillment.Execute(SQLIN_WebFulfillment)
						
		End If
	End If

on error goto 0	
	SQL_IN_WebFulfillment = "SELECT COL_LENGTH('IN_WebFulfillment', 'CustTypeNum') AS IsItThere"
	Set rsIN_WebFulfillment = cnnIN_WebFulfillment.Execute(SQL_IN_WebFulfillment)
	If IsNull(rsIN_WebFulfillment("IsItThere")) Then
		SQL_IN_WebFulfillment = "ALTER TABLE IN_WebFulfillment ADD CustTypeNum Int NULL"
		Set rsIN_WebFulfillment = cnnIN_WebFulfillment.Execute(SQL_IN_WebFulfillment)
	End If
	
	set rsIN_WebFulfillment = nothing
	cnnIN_WebFulfillment.close
	set cnnIN_WebFulfillment = nothing
%>