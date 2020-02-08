<%	

	Set cnnAR_CustomerCategoryDiscount = Server.CreateObject("ADODB.Connection")
	cnnAR_CustomerCategoryDiscount.open (Session("ClientCnnString"))
	Set rsAR_CustomerCategoryDiscount = Server.CreateObject("ADODB.Recordset")
	rsAR_CustomerCategoryDiscount.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsAR_CustomerCategoryDiscount = cnnAR_CustomerCategoryDiscount.Execute("SELECT TOP 1 * FROM AR_CustomerCategoryDiscount")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLAR_CustomerCategoryDiscount = "CREATE TABLE [AR_CustomerCategoryDiscount]( "
			SQLAR_CustomerCategoryDiscount = SQLAR_CustomerCategoryDiscount & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLAR_CustomerCategoryDiscount = SQLAR_CustomerCategoryDiscount & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_CustomerCategoryDiscount]  DEFAULT (getdate()), "
			SQLAR_CustomerCategoryDiscount = SQLAR_CustomerCategoryDiscount & " [CustID] [varchar](255) NULL, "
			SQLAR_CustomerCategoryDiscount = SQLAR_CustomerCategoryDiscount & " [Category] [varchar](255) NULL, "
			SQLAR_CustomerCategoryDiscount = SQLAR_CustomerCategoryDiscount & " [Discount] [float] NULL, "
			SQLAR_CustomerCategoryDiscount = SQLAR_CustomerCategoryDiscount & " [DiscountType] [varchar](255) NULL "
			SQLAR_CustomerCategoryDiscount = SQLAR_CustomerCategoryDiscount & ") ON [PRIMARY]"
			Set rsAR_CustomerCategoryDiscount = cnnAR_CustomerCategoryDiscount.Execute(SQLAR_CustomerCategoryDiscount)
		End If
	End If
	
	
	set rsAR_CustomerCategoryDiscount = nothing
	cnnAR_CustomerCategoryDiscount.close
	set cnnAR_CustomerCategoryDiscount = nothing


%>