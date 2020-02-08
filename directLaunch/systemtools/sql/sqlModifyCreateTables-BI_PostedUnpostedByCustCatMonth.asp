<%	
	Set cnnBI_PostedUnpostedByCustCatMonth = Server.CreateObject("ADODB.Connection")
	cnnBI_PostedUnpostedByCustCatMonth.open (Session("ClientCnnString"))
	Set rsBI_PostedUnpostedByCustCatMonth = Server.CreateObject("ADODB.Recordset")
	rsBI_PostedUnpostedByCustCatMonth.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsBI_PostedUnpostedByCustCatMonth = cnnBI_PostedUnpostedByCustCatMonth.Execute("SELECT TOP 1 * FROM BI_PostedUnpostedByCustCatMonth ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBI_PostedUnpostedByCustCatMonth = "CREATE TABLE [BI_PostedUnpostedByCustCatMonth]( "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_BI_PostedUnpostedByCustCatMonth_RecordCreationDateTime]  DEFAULT (getdate()), "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "[CustID] [varchar](255) NULL, "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "[CategoryID] [int] NULL, "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "[ThisPeriodSeqNumber] [int] NULL, "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "[ThisPeriodNumber] [int] NULL, "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "[ThisPeriodYear] [int] NULL, "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "[TotalSales] [money] NULL, "	
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "[TotalSalesWithTax] [money] NULL, "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "[TotalCost] [money] NULL, "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "[PostedOrUnposted] [varchar](1) NULL  "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & ") ON [PRIMARY]"
		
			Set rsBI_PostedUnpostedByCustCatMonth = cnnBI_PostedUnpostedByCustCatMonth.Execute(SQLBI_PostedUnpostedByCustCatMonth)
			
			
			
			SQLBI_PostedUnpostedByCustCatMonth = "CREATE CLUSTERED INDEX [IX_BI_PostedUnpostedByCustCatMonth2] ON [BI_PostedUnpostedByCustCatMonth] "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "( "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "	[CustID] ASC, "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "	[CategoryID] ASC "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & ")WITH (PAD_INDEX = OFF, "
			SQLBI_PostedUnpostedByCustCatMonth = SQLBI_PostedUnpostedByCustCatMonth & "STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
			

			Set rsBI_PostedUnpostedByCustCatMonth = cnnBI_PostedUnpostedByCustCatMonth.Execute(SQLBI_PostedUnpostedByCustCatMonth)
						
		End If
	End If
	
	set rsBI_PostedUnpostedByCustCatMonth = nothing
	cnnBI_PostedUnpostedByCustCatMonth.close
	set cnnBI_PostedUnpostedByCustCatMonth = nothing
				
%>