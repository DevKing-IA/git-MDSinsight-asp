<%	
	Set cnnBI_PostedUnpostedByCustCatPeriod = Server.CreateObject("ADODB.Connection")
	cnnBI_PostedUnpostedByCustCatPeriod.open (Session("ClientCnnString"))
	Set rsBI_PostedUnpostedByCustCatPeriod = Server.CreateObject("ADODB.Recordset")
	rsBI_PostedUnpostedByCustCatPeriod.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsBI_PostedUnpostedByCustCatPeriod = cnnBI_PostedUnpostedByCustCatPeriod.Execute("SELECT TOP 1 * FROM BI_PostedUnpostedByCustCatPeriod ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBI_PostedUnpostedByCustCatPeriod = "CREATE TABLE [BI_PostedUnpostedByCustCatPeriod]( "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_BI_PostedUnpostedByCustCatPeriod_RecordCreationDateTime]  DEFAULT (getdate()), "			
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[CustID] [varchar](255) NULL, "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[CategoryID] [int] NULL, "			
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[ThisPeriodSeqNumber] [int] NULL, "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[ThisPeriodNumber] [int] NULL, "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[ThisPeriodYear] [int] NULL, "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[TotalSales] [money] NULL, "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[TotalSalesWithTax] [money] NULL, "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[TotalCost] [money] NULL, "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[PostedOrUnposted] [varchar](1) NULL "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & ") ON [PRIMARY]"
		
			Set rsBI_PostedUnpostedByCustCatPeriod = cnnBI_PostedUnpostedByCustCatPeriod.Execute(SQLBI_PostedUnpostedByCustCatPeriod)
			
			
			SQLBI_PostedUnpostedByCustCatPeriod = "CREATE CLUSTERED INDEX [IX_BI_PostedUnpostedByCustCatPeriod] ON [BI_PostedUnpostedByCustCatPeriod] "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "( "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[CustID] ASC, "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "[CategoryID] ASC "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, "
			SQLBI_PostedUnpostedByCustCatPeriod = SQLBI_PostedUnpostedByCustCatPeriod & "DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] "

			Set rsBI_PostedUnpostedByCustCatPeriod = cnnBI_PostedUnpostedByCustCatPeriod.Execute(SQLBI_PostedUnpostedByCustCatPeriod)
			
		End If
	End If
	
	set rsBI_PostedUnpostedByCustCatPeriod = nothing
	cnnBI_PostedUnpostedByCustCatPeriod.close
	set cnnBI_PostedUnpostedByCustCatPeriod = nothing
				
%>