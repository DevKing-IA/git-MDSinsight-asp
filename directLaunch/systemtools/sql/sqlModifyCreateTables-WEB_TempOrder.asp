<%	

	Set cnnWEB_TempOrder = Server.CreateObject("ADODB.Connection")
	cnnWEB_TempOrder.open (Session("ClientCnnString"))
	Set rsWEB_TempOrder = Server.CreateObject("ADODB.Recordset")
	rsWEB_TempOrder.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_TempOrder = cnnWEB_TempOrder.Execute("SELECT TOP 1 * FROM WEB_TempOrder")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		   SQLWEB_TempOrder = "CREATE TABLE [WEB_TempOrder]("
		   SQLWEB_TempOrder = SQLWEB_TempOrder & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
		   SQLWEB_TempOrder = SQLWEB_TempOrder & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_TempOrder]  DEFAULT (getdate()), "
	       SQLWEB_TempOrder = SQLWEB_TempOrder & " [tmpOrderID] [int] NULL,"
	       SQLWEB_TempOrder = SQLWEB_TempOrder & " [tmpUserNo] [int] NULL,"
	       SQLWEB_TempOrder = SQLWEB_TempOrder & " [tmpCustID] [varchar](50) NULL,"
	       SQLWEB_TempOrder = SQLWEB_TempOrder & " [tmpProdSKU] [varchar](50) NULL,"
	       SQLWEB_TempOrder = SQLWEB_TempOrder & " [tmpQty] [int] NULL,"
	       SQLWEB_TempOrder = SQLWEB_TempOrder & " [tmpUMQty] [int] NULL,"
	       SQLWEB_TempOrder = SQLWEB_TempOrder & " [tmpUM] [varchar](50) NULL"
	       SQLWEB_TempOrder = SQLWEB_TempOrder & " ) ON [PRIMARY]"

		   Set rsWEB_TempOrder = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
		   
		End If
	End If
	
' This one is a DROP
	SQLWEB_TempOrder = "SELECT COL_LENGTH('WEB_TempOrder', 'tmpId') AS IsItThere"
	Set rsWEB_TempOrder  = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	If NOT IsNull(rsWEB_TempOrder("IsItThere")) Then
		SQLWEB_TempOrder = "ALTER TABLE WEB_TempOrder DROP COLUMN tmpId"
		Set rsWEB_TempOrder = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	End If

' This one is a DROP
	SQLWEB_TempOrder = "SELECT COL_LENGTH('WEB_TempOrder', 'tmpSuggestedItem') AS IsItThere"
	Set rsWEB_TempOrder  = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	If NOT IsNull(rsWEB_TempOrder("IsItThere")) Then
		SQLWEB_TempOrder = "ALTER TABLE WEB_TempOrder DROP COLUMN tmpSuggestedItem"
		Set rsWEB_TempOrder = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	End If
	
' This one is a DROP
	SQLWEB_TempOrder = "SELECT COL_LENGTH('WEB_TempOrder', 'tmpSuggestedPrice') AS IsItThere"
	Set rsWEB_TempOrder  = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	If NOT IsNull(rsWEB_TempOrder("IsItThere")) Then
		SQLWEB_TempOrder = "ALTER TABLE WEB_TempOrder DROP COLUMN tmpSuggestedPrice"
		Set rsWEB_TempOrder = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	End If
	
' This one is a DROP
	SQLWEB_TempOrder = "SELECT COL_LENGTH('WEB_TempOrder', 'tmpIsSuggestedItem') AS IsItThere"
	Set rsWEB_TempOrder  = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	If NOT IsNull(rsWEB_TempOrder("IsItThere")) Then
		SQLWEB_TempOrder = "ALTER TABLE WEB_TempOrder DROP COLUMN tmpIsSuggestedItem"
		Set rsWEB_TempOrder = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	End If
	
' This one is a DROP
	SQLWEB_TempOrder = "SELECT COL_LENGTH('WEB_TempOrder', 'tmpFeaturedItem') AS IsItThere"
	Set rsWEB_TempOrder  = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	If NOT IsNull(rsWEB_TempOrder("IsItThere")) Then
		SQLWEB_TempOrder = "ALTER TABLE WEB_TempOrder DROP COLUMN tmpFeaturedItem"
		Set rsWEB_TempOrder = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	End If

' This one is a DROP
	SQLWEB_TempOrder = "SELECT COL_LENGTH('WEB_TempOrder', 'tmpFeaturedPrice') AS IsItThere"
	Set rsWEB_TempOrder  = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	If NOT IsNull(rsWEB_TempOrder("IsItThere")) Then
		SQLWEB_TempOrder = "ALTER TABLE WEB_TempOrder DROP COLUMN tmpFeaturedPrice"
		Set rsWEB_TempOrder = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	End If

' This one is a DROP
	SQLWEB_TempOrder = "SELECT COL_LENGTH('WEB_TempOrder', 'tmpDate') AS IsItThere"
	Set rsWEB_TempOrder  = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	If NOT IsNull(rsWEB_TempOrder("IsItThere")) Then
		SQLWEB_TempOrder = "ALTER TABLE WEB_TempOrder DROP COLUMN tmpDate"
		Set rsWEB_TempOrder = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	End If
	
' This one is a DROP
	SQLWEB_TempOrder = "SELECT COL_LENGTH('WEB_TempOrder', 'tmpProdGroup') AS IsItThere"
	Set rsWEB_TempOrder  = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	If NOT IsNull(rsWEB_TempOrder("IsItThere")) Then
		SQLWEB_TempOrder = "ALTER TABLE WEB_TempOrder DROP COLUMN tmpProdGroup"
		Set rsWEB_TempOrder = cnnWEB_TempOrder.Execute(SQLWEB_TempOrder)
	End If

	
	set rsWEB_TempOrder = nothing
	cnnWEB_TempOrder.close
	set cnnWEB_TempOrder = nothing


%>