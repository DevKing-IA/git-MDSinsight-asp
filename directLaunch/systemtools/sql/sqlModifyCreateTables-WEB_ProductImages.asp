<%	

	Set cnnWEB_ProductImages = Server.CreateObject("ADODB.Connection")
	cnnWEB_ProductImages.open (Session("ClientCnnString"))
	Set rsWEB_ProductImages = Server.CreateObject("ADODB.Recordset")
	rsWEB_ProductImages.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_ProductImages = cnnWEB_ProductImages.Execute("SELECT TOP 1 * FROM WEB_ProductImages")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLWEB_ProductImages = "CREATE TABLE [WEB_ProductImages]( "
			SQLWEB_ProductImages = SQLWEB_ProductImages & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLWEB_ProductImages = SQLWEB_ProductImages & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_ProductImages]  DEFAULT (getdate()), "
			SQLWEB_ProductImages = SQLWEB_ProductImages & " [prodSKU] [varchar](255) NULL, "
			SQLWEB_ProductImages = SQLWEB_ProductImages & " [imgType] [varchar](50) NULL, "
			SQLWEB_ProductImages = SQLWEB_ProductImages & " [imgFilename] [varchar](255) NULL "
			SQLWEB_ProductImages = SQLWEB_ProductImages & ") ON [PRIMARY]"
			Set rsWEB_ProductImages = cnnWEB_ProductImages.Execute(SQLWEB_ProductImages)
		End If
	End If
	
	
	set rsWEB_ProductImages = nothing
	cnnWEB_ProductImages.close
	set cnnWEB_ProductImages = nothing


%>