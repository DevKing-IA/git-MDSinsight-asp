<%	
	Set cnnCustCatPeriodSales2 = Server.CreateObject("ADODB.Connection")
	cnnCustCatPeriodSales2.open (Session("ClientCnnString"))
	Set rsCustCatPeriodSales2 = Server.CreateObject("ADODB.Recordset")
	rsCustCatPeriodSales2.CursorLocation = 3 

	Err.Clear
	on error resume next ' In case it is already there

	'Create indexes
	SQLCustCatPeriodSales2 = "CREATE NONCLUSTERED INDEX [IX_CatPeriodSales_2_1] ON [BI_PostedUnpostedByCustCat]"
	SQLCustCatPeriodSales2 = SQLCustCatPeriodSales2 & " ([ThisPeriodSeqNumber],[PostedOrUnposted])"
	SQLCustCatPeriodSales2 = SQLCustCatPeriodSales2 & " INCLUDE ([CustID],[TotalSales])"

	Set rsCustCatPeriodSales2 = cnnCustCatPeriodSales2.Execute(SQLCustCatPeriodSales2 )
				
	set rsCustCatPeriodSales2 = nothing
	cnnCustCatPeriodSales2.close
	set cnnCustCatPeriodSales2 = nothing
	
	On Error Goto 0
				
%>