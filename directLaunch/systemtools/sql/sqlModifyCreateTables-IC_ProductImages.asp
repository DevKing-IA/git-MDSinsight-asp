<%	
	
	Set cnnProductImages = Server.CreateObject("ADODB.Connection")
	cnnProductImages.open (Session("ClientCnnString"))
	Set rsProductImages = Server.CreateObject("ADODB.Recordset")
	rsProductImages.CursorLocation = 3 
		
	'There was a pre-existing incorrect version to the IC_ProductImages table
	'First will check to see if it has the old partno field & kill the
	'table if it exists that way

	Err.Clear
	on error resume next
	Set rsProductImages = cnnProductImages.Execute("SELECT TOP 1 * FROM IC_ProductImages")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			SQL_ProductImages = "SELECT COL_LENGTH('IC_ProductImages', 'partno') AS IsItThere"
			Set rsProductImages = cnnProductImages.Execute(SQL_ProductImages)
			If NOT IsNull(rsProductImages("IsItThere")) Then
				SQL_ProductImages = "DROP TABLE IC_ProductImages"
				Set rsProductImages = cnnProductImages.Execute(SQL_ProductImages)
			End If
		End If
	End If
	
	
	Err.Clear
	on error resume next
	Set rsProductImages = cnnProductImages.Execute("SELECT TOP 1 * FROM IC_ProductImages")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE IC_ProductImages ("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_IC_ProductImages_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "[prodSKU] [varchar](255) NULL, "
			SQLBuild = SQLBuild & "[imgType] [varchar](255) NULL, "
			SQLBuild = SQLBuild & "[imgFilename] [varchar](255) NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			
			Set rsProductImages = cnnProductImages.Execute(SQLBuild)
		End If
	End If
	
			
	set rsProductImages = nothing
	cnnProductImages.close
	set cnnProductImages = nothing
	
				
%>