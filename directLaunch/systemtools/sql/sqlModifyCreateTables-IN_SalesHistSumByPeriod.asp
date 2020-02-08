<%	
	Set cnnCheckINSalesHistSumByPeriod = Server.CreateObject("ADODB.Connection")
	cnnCheckINSalesHistSumByPeriod.open (Session("ClientCnnString"))
	Set rsCheckINSalesHistSumByPeriod = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckINSalesHistSumByPeriod = cnnCheckINSalesHistSumByPeriod.Execute("SELECT TOP 1 * FROM IN_SalesHistSumByPeriod")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckINSalesHistSumByPeriod = "CREATE TABLE [IN_SalesHistSumByPeriod]("
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[RecordCreationDateTime] [datetime] NULL DEFAULT (getdate()), "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[CustNum] [varchar](255) NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[Period] [int] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[PeriodYear] [int] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[GrossSales] [money] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[ProductSales] [money] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[ProductCost] [money] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[ProductTax] [money] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[ProductOther] [money] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[GP] [money] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[Margin] [decimal](8, 2) NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[Rent] [money] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[Interest] [money] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[NonproductSales] [money] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[EquipmentValue] [money] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[NumberOfInvoices] [int] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[AccountBalance] [money] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[Salesman] [int] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[SecondarySalesman] [int] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[CustType] [int] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[ReferralCode] [int] NULL, "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & "	[ReferralDesc2] [varchar](255) NULL "
			SQLCheckINSalesHistSumByPeriod = SQLCheckINSalesHistSumByPeriod & " ) ON [PRIMARY]"      

		   	Set rsCheckINSalesHistSumByPeriod = cnnCheckINSalesHistSumByPeriod.Execute(SQLCheckINSalesHistSumByPeriod)
		   
		End If
	End If


	set rsCheckINSalesHistSumByPeriod = nothing
	cnnCheckINSalesHistSumByPeriod.close
	set cnnCheckINSalesHistSumByPeriod = nothing

%>