<%	
	Set cnnBI_DailySalesByTypeByClass = Server.CreateObject("ADODB.Connection")
	cnnBI_DailySalesByTypeByClass.open (Session("ClientCnnString"))
	Set rsBI_DailySalesByTypeByClass = Server.CreateObject("ADODB.Recordset")
	rsBI_DailySalesByTypeByClass.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsBI_DailySalesByTypeByClass = cnnBI_DailySalesByTypeByClass.Execute("SELECT TOP 1 * FROM BI_DailySalesByTypeByClass ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBI_DailySalesByTypeByClass = "CREATE TABLE [BI_DailySalesByTypeByClass]( "
			SQLBI_DailySalesByTypeByClass = SQLBI_DailySalesByTypeByClass & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLBI_DailySalesByTypeByClass = SQLBI_DailySalesByTypeByClass & " [TotNumOrders] [int] NULL, "
			SQLBI_DailySalesByTypeByClass = SQLBI_DailySalesByTypeByClass & " [TotSales] [money] NULL, "
			SQLBI_DailySalesByTypeByClass = SQLBI_DailySalesByTypeByClass & " [TotCost] [money] NULL, "
			SQLBI_DailySalesByTypeByClass = SQLBI_DailySalesByTypeByClass & " [IvsDate] [date] NULL, "
			SQLBI_DailySalesByTypeByClass = SQLBI_DailySalesByTypeByClass & " [IvsType] [varchar](255) NULL, "
			SQLBI_DailySalesByTypeByClass = SQLBI_DailySalesByTypeByClass & " [ClassCode] [varchar](255) NULL, "
			SQLBI_DailySalesByTypeByClass = SQLBI_DailySalesByTypeByClass & " [PeriodYear] [int] NULL, "
			SQLBI_DailySalesByTypeByClass = SQLBI_DailySalesByTypeByClass & " [Period] [int] NULL "
			SQLBI_DailySalesByTypeByClass = SQLBI_DailySalesByTypeByClass & ") ON [PRIMARY]"
			Set rsBI_DailySalesByTypeByClass = cnnBI_DailySalesByTypeByClass.Execute(SQLBI_DailySalesByTypeByClass)
		End If
	End If
	
	set rsBI_DailySalesByTypeByClass = nothing
	cnnBI_DailySalesByTypeByClass.close
	set cnnBI_DailySalesByTypeByClass = nothing
				
%>