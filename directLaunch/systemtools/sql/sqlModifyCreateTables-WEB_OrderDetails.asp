<%	

	Set cnnWEB_OrderDetails = Server.CreateObject("ADODB.Connection")
	cnnWEB_OrderDetails.open (Session("ClientCnnString"))
	Set rsWEB_OrderDetails = Server.CreateObject("ADODB.Recordset")
	rsWEB_OrderDetails.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_OrderDetails = cnnWEB_OrderDetails.Execute("SELECT TOP 1 * FROM WEB_OrderDetails")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		   SQLWEB_OrderDetails = "CREATE TABLE [WEB_OrderDetails]("
		   SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
		   SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_OrderDetails]  DEFAULT (getdate()), "
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [OrderID] [int] NULL,"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [seq] [char](10) NULL,"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [ProdSKU] [varchar](50) NULL,"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [Qty] [int] NULL,"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [SellPrice] [money] NULL,"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [UM] [varchar](20) NULL,"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [UMQty] [int] NULL,"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [UM_MPLEX] [varchar](50) NULL,"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [isSuggestedItem] [bit] NULL,"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [ShortDesc] [varchar](1000) NULL,"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [prodGroup] [varchar](500) NULL,"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & " [Tax] [money] NULL"
	       SQLWEB_OrderDetails = SQLWEB_OrderDetails & ") ON [PRIMARY]"

		   Set rsWEB_OrderDetails = cnnWEB_OrderDetails.Execute(SQLWEB_OrderDetails)
		   
		End If
	End If
	
' This one is a DROP
	SQLWEB_OrderDetails = "SELECT COL_LENGTH('WEB_OrderDetails', 'DetailID') AS IsItThere"
	Set rsWEB_OrderDetails  = cnnWEB_OrderDetails.Execute(SQLWEB_OrderDetails)
	If NOT IsNull(rsWEB_OrderDetails("IsItThere")) Then
		SQLWEB_OrderDetails = "ALTER TABLE WEB_OrderDetails DROP COLUMN DetailID"
		Set rsWEB_OrderDetails = cnnWEB_OrderDetails.Execute(SQLWEB_OrderDetails)
	End If
	
	set rsWEB_OrderDetails = nothing
	cnnWEB_OrderDetails.close
	set cnnWEB_OrderDetails = nothing


%>