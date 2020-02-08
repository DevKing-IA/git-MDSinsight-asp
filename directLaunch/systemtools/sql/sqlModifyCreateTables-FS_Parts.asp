<%	
	Set cnnCheckFSParts = Server.CreateObject("ADODB.Connection")
	cnnCheckFSParts.open (Session("ClientCnnString"))
	Set rsCheckFSParts = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	
	Set rsCheckFSParts = cnnCheckFSParts.Execute("SELECT TOP 1 * FROM FS_Parts")

	
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			'The table is not there, we need to create it
			On error goto 0	
		    SQLCheckFSParts = "CREATE TABLE [FS_Parts]("
			SQLCheckFSParts = SQLCheckFSParts & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckFSParts = SQLCheckFSParts & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_FS_Parts_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckFSParts = SQLCheckFSParts & " [PartNumber] [varchar](255) NULL,"
			SQLCheckFSParts = SQLCheckFSParts & " [PartDescription] [varchar](8000) NULL, "
			SQLCheckFSParts = SQLCheckFSParts & " [DisplayOrder] [int] NULL, "
			SQLCheckFSParts = SQLCheckFSParts & " ) ON [PRIMARY]"      

			Set rsCheckFSParts = cnnCheckFSParts.Execute(SQLCheckFSParts)

		End If
	End If
	
	'Special for the parts  file
	'Make sure code 0 is there
	SQLCheckFSParts = "SELECT * FROM FS_Parts WHERE InternalRecordIdentifier = 0"
	Set rsCheckFSParts = cnnCheckFSParts.Execute(SQLCheckFSParts)
	If rsCheckFSParts.EOF Then 
	
		SQLCheckFSParts = "SET IDENTITY_INSERT FS_Parts ON;"
		Set rsCheckFSParts = cnnCheckFSParts.Execute(SQLCheckFSParts)

		SQLCheckFSParts = SQLCheckFSParts & "INSERT INTO FS_Parts (InternalRecordIdentifier,PartNumber,DisplayOrder) "
		SQLCheckFSParts = SQLCheckFSParts & " VALUES (0,'Other',0)"
		Response.Write(SQLCheckFSParts)
		Set rsCheckFSParts = cnnCheckFSParts.Execute(SQLCheckFSParts)
		
		SQLCheckFSParts = "SET IDENTITY_INSERT FS_Parts OFF;"
		Set rsCheckFSParts = cnnCheckFSParts.Execute(SQLCheckFSParts)
		
	End If

	SQLCheckFSParts = "SELECT COL_LENGTH('FS_Parts', 'SearchKeywords') AS IsItThere"
	Set rsCheckFSParts = cnnCheckFSParts.Execute(SQLCheckFSParts)
	If IsNull(rsCheckFSParts("IsItThere")) Then
		SQLCheckFSParts = "ALTER TABLE FS_Parts ADD SearchKeywords varchar(8000) NULL"
		Set rsCheckFSParts = cnnCheckFSParts.Execute(SQLCheckFSParts)
	End If
	
	set rsCheckFSParts = nothing
	cnnCheckFSParts.close
	set cnnCheckFSParts = nothing

		
%>