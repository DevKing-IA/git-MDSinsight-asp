<%	

	Set cnnWEB_Order = Server.CreateObject("ADODB.Connection")
	cnnWEB_Order.open (Session("ClientCnnString"))
	Set rsWEB_Order = Server.CreateObject("ADODB.Recordset")
	rsWEB_Order.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_Order = cnnWEB_Order.Execute("SELECT TOP 1 * FROM WEB_Order")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		   SQLWEB_Order = "CREATE TABLE [WEB_Order]("
	       SQLWEB_Order = SQLWEB_Order & " [OrderID] [int] IDENTITY(2,1) NOT NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [OrderDate] [datetime] NULL CONSTRAINT [DF_WEB_Order]  DEFAULT (getdate()),"
	       SQLWEB_Order = SQLWEB_Order & " [CustID] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [PromoCode] [varchar](20) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [merchTotal] [money] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [OrderTotal] [money] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [tax] [money] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [freight] [money] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Comments] [varchar](3000) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [trackingNum] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [giftMessage] [varchar](255) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [shipMethod] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [OrderType] [char](10) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [UserNo] [int] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Dept] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [CCDescription] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [PO] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [RequestedDate] [smalldatetime] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Name] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Address1] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Address2] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [City] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [State] [varchar](10) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Zip] [varchar](10) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [RushRequest] [bit] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [OrderPending] [smallmoney] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [OrderShipLoc] [int] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [CostCenter] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billMethod] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billCheckName] [varchar](255) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billCheckNumber] [int] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billMICRNumber] [varchar](100) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billCheckType] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billStateID] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billCheckDLSID] [varchar](100) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [StreamW_ShipToLocation] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [StreamW_POS] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [MPlex_ChargeCCOnFile] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Exported] [bit] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [TelSelOrder] [bit] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Email] [varchar](150) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [TransactionID] [varchar](100) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [transOrderDate] [datetime] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [transOrderTime] [datetime] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [merchSubtotal] [money] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [shipAddrType] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Company] [varchar](100) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billCompany] [varchar](100) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billName] [varchar](100) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billAddress1] [varchar](100) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billAddress2] [varchar](100) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billCity] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billState] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billZip] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [billPhone] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [ccType] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [ccNumber] [varchar](100) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [ccExpYear] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [ccNameOnCard] [varchar](255) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [ccExpMonth] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [CVV2] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Phone] [varchar](100) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [ccLast4Digits] [varchar](50) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [RouteSelOrder] [bit] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [PrintedName] [varchar](255) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Signature] [text] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [SignaturePNG] [text] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [RouteSelUserNo] [int] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [RouteSelUserName] [varchar](255) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [RouteSelOrderType] [varchar](255) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [MobileOrder] [bit] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [RouteSelDeviceGUID] [varchar](100) NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [Deposit] [money] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [DepositReturn] [money] NULL,"
	       SQLWEB_Order = SQLWEB_Order & " [TaxRate] [float] NULL"
	       SQLWEB_Order = SQLWEB_Order & " ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

		   Set rsWEB_Order = cnnWEB_Order.Execute(SQLWEB_Order)
		   
		End If
	End If
	

' This one is a DROP
	SQLWEB_Order = "SELECT COL_LENGTH('WEB_Order', 'InternalRecordIdentifier') AS IsItThere"
	Set rsWEB_Order  = cnnWEB_Order.Execute(SQLWEB_Order)
	If NOT IsNull(rsWEB_Order("IsItThere")) Then
		SQLWEB_Order = "ALTER TABLE WEB_Order DROP COLUMN InternalRecordIdentifier"
		Set rsWEB_Order = cnnWEB_Order.Execute(SQLWEB_Order)
	End If

' This one is a DROP
	SQLWEB_Order = "SELECT COL_LENGTH('WEB_Order', 'RecordCreationDateTime') AS IsItThere"
	Set rsWEB_Order  = cnnWEB_Order.Execute(SQLWEB_Order)
	If NOT IsNull(rsWEB_Order("IsItThere")) Then
		SQLWEB_Order = "ALTER TABLE WEB_Order DROP COLUMN RecordCreationDateTime"
		Set rsWEB_Order = cnnWEB_Order.Execute(SQLWEB_Order)
	End If
	
	set rsWEB_Order = nothing
	cnnWEB_Order.close
	set cnnWEB_Order = nothing


%>