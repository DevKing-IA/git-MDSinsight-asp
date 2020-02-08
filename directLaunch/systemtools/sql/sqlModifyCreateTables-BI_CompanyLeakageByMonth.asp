<%	
	Set cnnCompanyLeakageByMonth = Server.CreateObject("ADODB.Connection")
	cnnCompanyLeakageByMonth.open (Session("ClientCnnString"))
	Set rsCompanyLeakageByMonth = Server.CreateObject("ADODB.Recordset")
	rsCompanyLeakageByMonth.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsCompanyLeakageByMonth = cnnCompanyLeakageByMonth.Execute("SELECT TOP 1 * FROM BI_CompanyLeakageByMonth")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLCompanyLeakageByMonth = "CREATE TABLE [BI_CompanyLeakageByMonth]( "
			SQLCompanyLeakageByMonth = SQLCompanyLeakageByMonth & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLCompanyLeakageByMonth = SQLCompanyLeakageByMonth & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_bi_CompanyLeakageByMonth_RecordCreated]  DEFAULT (getdate()), "
			
			SQLCompanyLeakageByMonth = SQLCompanyLeakageByMonth & " [PeriodSequenceNumber] [int] NULL, "
			SQLCompanyLeakageByMonth = SQLCompanyLeakageByMonth & " [TotalActiveCustomers] [int] NULL, "
			SQLCompanyLeakageByMonth = SQLCompanyLeakageByMonth & " [TotalLeakingCustomer] [int] NULL, "
			SQLCompanyLeakageByMonth = SQLCompanyLeakageByMonth & " [TotalLCPvs3PPAvgVariance] [money] NULL, "
			SQLCompanyLeakageByMonth = SQLCompanyLeakageByMonth & ") ON [PRIMARY]"
		
			Set rsCompanyLeakageByMonth = cnnCompanyLeakageByMonth.Execute(SQLCompanyLeakageByMonth)
		End If
	End If
				
	set rsCompanyLeakageByMonth = nothing
	cnnCompanyLeakageByMonth.close
	set cnnCompanyLeakageByMonth = nothing
				
%>