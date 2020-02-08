<%	

	Set cnnINUnpostedInvoiceRpt = Server.CreateObject("ADODB.Connection")
	cnnINUnpostedInvoiceRpt.open (Session("ClientCnnString"))
	Set rsINUnpostedInvoiceRpt = Server.CreateObject("ADODB.Recordset")
	rsINUnpostedInvoiceRpt.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsINUnpostedInvoiceRpt = cnnINUnpostedInvoiceRpt.Execute("SELECT TOP 1 * FROM IN_UnpostedInvoiceRpt")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLINUnpostedInvoiceRpt = "CREATE TABLE [IN_UnpostedInvoiceRpt]( "
			SQLINUnpostedInvoiceRpt = SQLINUnpostedInvoiceRpt & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLINUnpostedInvoiceRpt = SQLINUnpostedInvoiceRpt & "[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_IN_UnpostedInvoiceRpt_RecordCreationDateTime]  DEFAULT (getdate()),"	
			SQLINUnpostedInvoiceRpt = SQLINUnpostedInvoiceRpt & "[CustID] [varchar](255) NULL, "
			SQLINUnpostedInvoiceRpt = SQLINUnpostedInvoiceRpt & "[CustName1] [varchar](255) NULL, "
			SQLINUnpostedInvoiceRpt = SQLINUnpostedInvoiceRpt & "[CustName2] [varchar](255) NULL, "
			SQLINUnpostedInvoiceRpt = SQLINUnpostedInvoiceRpt & "[InvoiceID] [varchar](255) NULL, "
			SQLINUnpostedInvoiceRpt = SQLINUnpostedInvoiceRpt & "[InvoiceDate] [date] NULL, "
			SQLINUnpostedInvoiceRpt = SQLINUnpostedInvoiceRpt & "[InvoiceAmount] [money] NULL, "
			SQLINUnpostedInvoiceRpt = SQLINUnpostedInvoiceRpt & "[Route] [varchar](255) NULL, "
			SQLINUnpostedInvoiceRpt = SQLINUnpostedInvoiceRpt & "[SalesMan] [varchar](255) NULL "
			SQLINUnpostedInvoiceRpt = SQLINUnpostedInvoiceRpt & ") ON [PRIMARY] "
	
	
			Set rsINUnpostedInvoiceRpt = cnnINUnpostedInvoiceRpt.Execute(SQLINUnpostedInvoiceRpt)
			
			
		End If
	End If
	
	
	set rsINUnpostedInvoiceRpt = nothing
	cnnINUnpostedInvoiceRpt.close
	set cnnINUnpostedInvoiceRpt = nothing
%>