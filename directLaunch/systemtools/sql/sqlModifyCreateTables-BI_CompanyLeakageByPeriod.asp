<%	
	Set cnnCompanyLeakage = Server.CreateObject("ADODB.Connection")
	cnnCompanyLeakage.open (Session("ClientCnnString"))
	Set rsCompanyLeakage = Server.CreateObject("ADODB.Recordset")
	rsCompanyLeakage.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsCompanyLeakage = cnnCompanyLeakage.Execute("SELECT TOP 1 * FROM BI_CompanyLeakageByPeriod")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLCompanyLeakage = "CREATE TABLE [BI_CompanyLeakageByPeriod]( "
			SQLCompanyLeakage = SQLCompanyLeakage & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLCompanyLeakage = SQLCompanyLeakage & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_bi_companyleakage_RecordCreated]  DEFAULT (getdate()), "
			
			SQLCompanyLeakage = SQLCompanyLeakage & " [PeriodSequenceNumber] [int] NULL, "
			SQLCompanyLeakage = SQLCompanyLeakage & " [TotalActiveCustomers] [int] NULL, "
			SQLCompanyLeakage = SQLCompanyLeakage & " [TotalLeakingCustomer] [int] NULL, "
			SQLCompanyLeakage = SQLCompanyLeakage & " [TotalLCPvs3PPAvgVariance] [money] NULL, "
			SQLCompanyLeakage = SQLCompanyLeakage & ") ON [PRIMARY]"
		
			Set rsCompanyLeakage = cnnCompanyLeakage.Execute(SQLCompanyLeakage)
		End If
	End If
				
	set rsCompanyLeakage = nothing
	cnnCompanyLeakage.close
	set cnnCompanyLeakage = nothing
				
%>