<%	
	Set cnnShopifyAPI = Server.CreateObject("ADODB.Connection")
	cnnShopifyAPI.open (Session("ClientCnnString"))
	Set rsShopifyAPI = Server.CreateObject("ADODB.Recordset")
	rsShopifyAPI.CursorLocation = 3 


		
		
	Err.Clear
	on error resume next
	Set rsShopifyAPI = cnnShopifyAPI.Execute("SELECT TOP 1 * FROM API_IN_Shopify_FulfillmentHeader ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE API_IN_Shopify_FulfillmentHeader ("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_API_IN_Shopify_FulfillmentHeader_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "insightClientSerno [varchar](50) NULL, "
			SQLBuild = SQLBuild & "HTTP_X_SHOPIFY_TOPIC [varchar](255) NULL, "
			SQLBuild = SQLBuild & "HTTP_X_SHOPIFY_SHOP_DOMAIN [varchar](255) NULL, "
			SQLBuild = SQLBuild & "HTTP_X_SHOPIFY_ORDER_ID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "HTTP_X_SHOPIFY_TEST [varchar](255) NULL, "
			SQLBuild = SQLBuild & "HTTP_X_SHOPIFY_HMAC_SHA256 [varchar](255) NULL, "
			SQLBuild = SQLBuild & "internalShopifyFulfillmentID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "internalShopifyOrderID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "fulfillmentCreatedAt [datetime] NULL, "
			SQLBuild = SQLBuild & "fulfillmentUpdatedAt [datetime] NULL, "
			SQLBuild = SQLBuild & "fulfillmentCreatedAtDate [date] NULL, "
			SQLBuild = SQLBuild & "fulfillmentCreatedAtTime [time](7) NULL, "
			SQLBuild = SQLBuild & "fulfillmentEmail [varchar](255) NULL, "
			SQLBuild = SQLBuild & "fulfillmentShipmentStatus [varchar](255) NULL, "
			SQLBuild = SQLBuild & "fulfillmentService [varchar](255) NULL, "
			SQLBuild = SQLBuild & "fulfillmentTrackingCompany [varchar](255) NULL, "
			SQLBuild = SQLBuild & "fulfillmentTrackingNumbers [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "fulfillmentTrackingURLs [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "destinationFirstName [varchar](255) NULL, "
			SQLBuild = SQLBuild & "destinationLastName [varchar](255) NULL, "
			SQLBuild = SQLBuild & "destinationCompany [varchar](255) NULL, "
			SQLBuild = SQLBuild & "destinationPhone [varchar](80) NULL, "
			SQLBuild = SQLBuild & "destinationAddress1 [varchar](255) NULL, "
			SQLBuild = SQLBuild & "destinationAddress2 [varchar](255) NULL, "
			SQLBuild = SQLBuild & "destinationCity [varchar](255) NULL, "
			SQLBuild = SQLBuild & "destinationZip [varchar](50) NULL, "
			SQLBuild = SQLBuild & "destinationProvince [varchar](50) NULL, "
			SQLBuild = SQLBuild & "destinationProvinceCode [varchar](10) NULL, "
			SQLBuild = SQLBuild & "destinationCountry [varchar](50) NULL, "
			SQLBuild = SQLBuild & "destinationCountryCode [varchar](10) NULL, "
			SQLBuild = SQLBuild & "destinationLatitude [varchar](50) NULL, "
			SQLBuild = SQLBuild & "destinationLongitude [varchar](50) NULL, "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			
			Set rsShopifyAPI = cnnShopifyAPI.Execute(SQLBuild)
		End If
	End If
	
		
	
	Err.Clear
	on error resume next
	Set rsShopifyAPI = cnnShopifyAPI.Execute("SELECT TOP 1 * FROM API_IN_Shopify_FulfillmentDetail ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE API_IN_Shopify_FulfillmentDetail ("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_API_IN_Shopify_FulfillmentDetail_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "internalShopifyFulfillmentID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "internalShopifyOrderID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shopifyInternalLineItemID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "variantID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "variantTitle [varchar](255) NULL, "
			SQLBuild = SQLBuild & "sku [varchar](255) NULL, "
			SQLBuild = SQLBuild & "title [varchar](255) NULL, "
			SQLBuild = SQLBuild & "quantity [int] NULL, "
			SQLBuild = SQLBuild & "um [varchar](50) NULL, "
			SQLBuild = SQLBuild & "price [float] NULL, "
			SQLBuild = SQLBuild & "vendor [varchar](255) NULL, "
			SQLBuild = SQLBuild & "fulfillableQuantity [int] NULL, "
			SQLBuild = SQLBuild & "fulfillmentService [varchar](255) NULL, "
			SQLBuild = SQLBuild & "fulfillmentStatus [varchar](255) NULL, "
			SQLBuild = SQLBuild & "requiresShipping [bit] NULL, "
			SQLBuild = SQLBuild & "taxable [bit] NULL, "
			SQLBuild = SQLBuild & "giftCard [bit] NULL, "
			SQLBuild = SQLBuild & "variantInventoryManagement [varchar](255) NULL, "
			SQLBuild = SQLBuild & "productExists [bit] NULL, "
			SQLBuild = SQLBuild & "properties [varchar](255) NULL, "
			SQLBuild = SQLBuild & "grams [float] NULL, "
			SQLBuild = SQLBuild & "totalDiscount [float] NULL, "
			SQLBuild = SQLBuild & "taxLines [varchar](255) NULL, "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			
			Set rsShopifyAPI = cnnShopifyAPI.Execute(SQLBuild)
		End If
	End If
	
	
	Err.Clear
	on error resume next
	Set rsShopifyAPI = cnnShopifyAPI.Execute("SELECT TOP 1 * FROM API_IN_Shopify_OrderHeader ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE API_IN_Shopify_OrderHeader ("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_API_IN_Shopify_OrderHeader_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "insightClientSerno [varchar](50) NULL, "
			SQLBuild = SQLBuild & "HTTP_X_SHOPIFY_TOPIC [varchar](255) NULL, "
			SQLBuild = SQLBuild & "HTTP_X_SHOPIFY_SHOP_DOMAIN [varchar](255) NULL, "
			SQLBuild = SQLBuild & "HTTP_X_SHOPIFY_ORDER_ID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "HTTP_X_SHOPIFY_TEST [varchar](255) NULL, "
			SQLBuild = SQLBuild & "HTTP_X_SHOPIFY_HMAC_SHA256 [varchar](255) NULL, "
			SQLBuild = SQLBuild & "internalShopifyOrderID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shopifyOfficialOrderNumber [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shopifyAlternateOrderNumber1 [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shopifyAlternateOrderNumber2 [varchar](255) NULL, "
			SQLBuild = SQLBuild & "orderClosedAt [datetime] NULL, "
			SQLBuild = SQLBuild & "orderCreatedAt [datetime] NULL, "
			SQLBuild = SQLBuild & "orderUpdatedAt [datetime] NULL, "
			SQLBuild = SQLBuild & "orderProcessedAt [datetime] NULL, "
			SQLBuild = SQLBuild & "orderProcessedAtDate [date] NULL, "
			SQLBuild = SQLBuild & "orderProcessedAtTime [time](7) NULL, "
			SQLBuild = SQLBuild & "orderNumber [varchar](255) NULL, "
			SQLBuild = SQLBuild & "orderEmail [varchar](255) NULL, "
			SQLBuild = SQLBuild & "orderTestType [bit] NULL, "
			SQLBuild = SQLBuild & "orderTotalWeight [float] NULL, "
			SQLBuild = SQLBuild & "orderTotalTax [float] NULL, "
			SQLBuild = SQLBuild & "orderTaxesIncluded [bit] NULL, "
			SQLBuild = SQLBuild & "orderTaxLines [varchar](50) NULL, "
			SQLBuild = SQLBuild & "orderTotalDiscounts [float] NULL, "
			SQLBuild = SQLBuild & "orderTotalLineItemsPrice [float] NULL, "
			SQLBuild = SQLBuild & "orderSubtotalPrice [float] NULL, "
			SQLBuild = SQLBuild & "orderTotalPrice [float] NULL, "
			SQLBuild = SQLBuild & "orderTotalPriceUSD [float] NULL, "
			SQLBuild = SQLBuild & "orderDiscountCodes [varchar](255) NULL, "
			SQLBuild = SQLBuild & "orderCurrency [varchar](10) NULL, "
			SQLBuild = SQLBuild & "orderFinancialStatus [varchar](50) NULL, "
			SQLBuild = SQLBuild & "orderProcessingMethod [varchar](50) NULL, "
			SQLBuild = SQLBuild & "orderFulfillmentStatus [varchar](50) NULL, "
			SQLBuild = SQLBuild & "orderStatusURL [varchar](255) NULL, "
			SQLBuild = SQLBuild & "orderConfirmed [bit] NULL, "
			SQLBuild = SQLBuild & "note [varchar](1000) NULL, "
			SQLBuild = SQLBuild & "noteAttributes [varchar](255) NULL, "
			SQLBuild = SQLBuild & "token [varchar](255) NULL, "
			SQLBuild = SQLBuild & "gateway [varchar](255) NULL, "
			SQLBuild = SQLBuild & "cartToken [varchar](50) NULL, "
			SQLBuild = SQLBuild & "checkoutToken [varchar](50) NULL, "
			SQLBuild = SQLBuild & "buyerAcceptsMarketing [bit] NULL, "
			SQLBuild = SQLBuild & "referringSite [varchar](255) NULL, "
			SQLBuild = SQLBuild & "landingSite [varchar](255) NULL, "
			SQLBuild = SQLBuild & "landingSiteRef [varchar](255) NULL, "
			SQLBuild = SQLBuild & "cancelledAt [datetime] NULL, "
			SQLBuild = SQLBuild & "cancelReason [varchar](255) NULL, "
			SQLBuild = SQLBuild & "reference [varchar](255) NULL, "
			SQLBuild = SQLBuild & "userID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "locationID [int] NULL, "
			SQLBuild = SQLBuild & "sourceIdentifier [varchar](50) NULL, "
			SQLBuild = SQLBuild & "sourceName [varchar](50) NULL, "
			SQLBuild = SQLBuild & "sourceURL [varchar](255) NULL, "
			SQLBuild = SQLBuild & "deviceID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "checkoutID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "appID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "customerLocale [varchar](50) NULL, "
			SQLBuild = SQLBuild & "browserIP [varchar](50) NULL, "
			SQLBuild = SQLBuild & "paymentGatewayNames [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shippingMethod [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shippingTitle [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shippingPrice [float] NULL, "
			SQLBuild = SQLBuild & "shippingCode [varchar](50) NULL, "
			SQLBuild = SQLBuild & "shippingSource [varchar](50) NULL, "
			SQLBuild = SQLBuild & "shippingPhone [varchar](50) NULL, "
			SQLBuild = SQLBuild & "shippingReqFulfillmentSvcID [int] NULL, "
			SQLBuild = SQLBuild & "shippingDeliveryCategory [varchar](50) NULL, "
			SQLBuild = SQLBuild & "shippingCarrierIdentifier [varchar](50) NULL, "
			SQLBuild = SQLBuild & "shippingDiscountedPrice [float] NULL, "
			SQLBuild = SQLBuild & "shippingTaxLines [varchar](50) NULL, "
			SQLBuild = SQLBuild & "shippingAddressFirstName [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shippingAddressLastName [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shippingAddressFullName [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shippingAddressCompany [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shippingAddressAddress1 [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shippingAddressAddress2 [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shippingAddressCity [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shippingAddressZip [varchar](50) NULL, "
			SQLBuild = SQLBuild & "shippingAddressProvince [varchar](50) NULL, "
			SQLBuild = SQLBuild & "shippingAddressProvinceCode [varchar](10) NULL, "
			SQLBuild = SQLBuild & "shippingAddressCountry [varchar](50) NULL, "
			SQLBuild = SQLBuild & "shippingAddressCountryCode[varchar](10) NULL, "
			SQLBuild = SQLBuild & "shippingAddressLatitude [varchar](50) NULL, "
			SQLBuild = SQLBuild & "shippingAddressLongitude [varchar](50) NULL, "
			SQLBuild = SQLBuild & "billingAddressFirstName [varchar](255) NULL, "
			SQLBuild = SQLBuild & "billingAddressLastName [varchar](255) NULL, "
			SQLBuild = SQLBuild & "billingAddressFullName [varchar](255) NULL, "
			SQLBuild = SQLBuild & "billingAddressCompany [varchar](255) NULL, "
			SQLBuild = SQLBuild & "billingAddressAddress1 [varchar](255) NULL, "
			SQLBuild = SQLBuild & "billingAddressAddress2 [varchar](255) NULL, "
			SQLBuild = SQLBuild & "billingAddressCity [varchar](255) NULL, "
			SQLBuild = SQLBuild & "billingAddressZip [varchar](50) NULL, "
			SQLBuild = SQLBuild & "billingAddressProvince [varchar](50) NULL, "
			SQLBuild = SQLBuild & "billingAddressProvinceCode [varchar](10) NULL, "
			SQLBuild = SQLBuild & "billingAddressCountry [varchar](50) NULL, "
			SQLBuild = SQLBuild & "billingAddressCountryCode[varchar](10) NULL, "
			SQLBuild = SQLBuild & "billingAddressLatitude [varchar](50) NULL, "
			SQLBuild = SQLBuild & "billingAddressLongitude [varchar](50) NULL, "
			SQLBuild = SQLBuild & "customerIDShopify [varchar](255) NULL, "
			SQLBuild = SQLBuild & "customerFirstName [varchar](255) NULL, "
			SQLBuild = SQLBuild & "customerLastName [varchar](255) NULL, "
			SQLBuild = SQLBuild & "customerEmail [varchar](255) NULL, "
			SQLBuild = SQLBuild & "customerPhone [varchar](50) NULL, "
			SQLBuild = SQLBuild & "customerCreated [datetime] NULL, "
			SQLBuild = SQLBuild & "customerUpdated [datetime] NULL, "
			SQLBuild = SQLBuild & "customerOrdersCount [int] NULL, "
			SQLBuild = SQLBuild & "customerTotalSpent [float] NULL, "
			SQLBuild = SQLBuild & "customerLastOrderID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "customerTaxExempt [bit] NULL, "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			
			Set rsShopifyAPI = cnnShopifyAPI.Execute(SQLBuild)
		End If
	End If
	
	
	Err.Clear
	on error resume next
	Set rsShopifyAPI = cnnShopifyAPI.Execute("SELECT TOP 1 * FROM API_IN_Shopify_OrderDetail ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE API_IN_Shopify_OrderDetail ("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_API_IN_Shopify_OrderDetail_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "shopifyOrderID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "internalShopifyOrderID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "shopifyInternalLineItemID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "variantID [varchar](255) NULL, "
			SQLBuild = SQLBuild & "variantTitle [varchar](255) NULL, "
			SQLBuild = SQLBuild & "sku [varchar](255) NULL, "
			SQLBuild = SQLBuild & "title [varchar](255) NULL, "
			SQLBuild = SQLBuild & "quantity [int] NULL, "
			SQLBuild = SQLBuild & "um [varchar](50) NULL, "
			SQLBuild = SQLBuild & "price [float] NULL, "
			SQLBuild = SQLBuild & "vendor [varchar](255) NULL, "
			SQLBuild = SQLBuild & "fulfillableQuantity [int] NULL, "
			SQLBuild = SQLBuild & "fulfillmentService [varchar](255) NULL, "
			SQLBuild = SQLBuild & "fulfillmentStatus [varchar](255) NULL, "
			SQLBuild = SQLBuild & "requiresShipping [bit] NULL, "
			SQLBuild = SQLBuild & "taxable [bit] NULL, "
			SQLBuild = SQLBuild & "giftCard [bit] NULL, "
			SQLBuild = SQLBuild & "variantInventoryManagement [varchar](255) NULL, "
			SQLBuild = SQLBuild & "productExists [bit] NULL, "
			SQLBuild = SQLBuild & "properties [varchar](255) NULL, "
			SQLBuild = SQLBuild & "grams [float] NULL, "
			SQLBuild = SQLBuild & "totalDiscount [float] NULL, "
			SQLBuild = SQLBuild & "taxLines [varchar](255) NULL, "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			
			Set rsShopifyAPI = cnnShopifyAPI.Execute(SQLBuild)
		End If
	End If
	
	
	
	Err.Clear
	on error resume next
	Set rsShopifyAPI = cnnShopifyAPI.Execute("SELECT TOP 1 * FROM API_IN_Shopify_Log ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE API_IN_Shopify_Log ("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_API_IN_Shopify_Log_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "Thread [int] NULL, "
			SQLBuild = SQLBuild & "Event [varchar](1000) NULL, "
			SQLBuild = SQLBuild & "Data [varchar](1000) NULL, "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			
			Set rsShopifyAPI = cnnShopifyAPI.Execute(SQLBuild)
		End If
	End If
				
	set rsShopifyAPI = nothing
	cnnShopifyAPI.close
	set cnnShopifyAPI = nothing
	
				
%>