<%	
	Set cnnCheckFSResolutionCodes = Server.CreateObject("ADODB.Connection")
	cnnCheckFSResolutionCodes.open (Session("ClientCnnString"))
	Set rsCheckFSResolutionCodes = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckFSResolutionCodes = cnnCheckFSResolutionCodes.Execute("SELECT TOP 1 * FROM FS_ResolutionCodes")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckFSResolutionCodes = "CREATE TABLE [FS_ResolutionCodes]("
			SQLCheckFSResolutionCodes = SQLCheckFSResolutionCodes & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckFSResolutionCodes = SQLCheckFSResolutionCodes & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_FS_ResolutionCodes_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckFSResolutionCodes = SQLCheckFSResolutionCodes & " [ResolutionDescription] [varchar](255) NULL"
			SQLCheckFSResolutionCodes = SQLCheckFSResolutionCodes & " ) ON [PRIMARY]"      

		   Set rsCheckFSResolutionCodes = cnnCheckFSResolutionCodes.Execute(SQLCheckFSResolutionCodes)
		   
		End If
	End If

	on error goto 0

	'Special for the Resolution code file
	'Make sure code 0 is there
	SQLCheckFSResolutionCodes  = "SELECT * FROM FS_ResolutionCodes WHERE InternalRecordIdentifier = 0"
	Set rsCheckFSResolutionCodes = cnnCheckFSResolutionCodes.Execute(SQLCheckFSResolutionCodes)
	If rsCheckFSResolutionCodes.EOF Then 
	
		on error resume next
		SQLCheckFSResolutionCodes = "SET IDENTITY_INSERT FS_ResolutionCodes ON;"
		Set rsCheckFSResolutionCodes = cnnCheckFSResolutionCodes.Execute(SQLCheckFSResolutionCodes)
		On Error Goto 0

		SQLCheckFSResolutionCodes = "INSERT INTO FS_ResolutionCodes (InternalRecordIdentifier,ResolutionDescription) "
		SQLCheckFSResolutionCodes = SQLCheckFSResolutionCodes & " VALUES (0,'Other');"
		'Set rsCheckFSResolutionCodes = cnnCheckFSResolutionCodes.Execute(SQLCheckFSResolutionCodes)
		
		SQLCheckFSResolutionCodes = "SET IDENTITY_INSERT FS_ResolutionCodes OFF;"
		Set rsCheckFSResolutionCodes = cnnCheckFSResolutionCodes.Execute(SQLCheckFSResolutionCodes)
		
	End If
	
	' This one is a DROP
	SQLCheckFSResolutionCodes = "SELECT COL_LENGTH('FS_ResolutionCodes', 'ShowOnWebsite') AS IsItThere"
	Set rsCheckFSResolutionCodes   = cnnCheckFSResolutionCodes.Execute(SQLCheckFSResolutionCodes)
	If NOT IsNull(rsCheckFSResolutionCodes("IsItThere")) Then
		SQLCheckFSResolutionCodes = "ALTER TABLE FS_ResolutionCodes DROP COLUMN ShowOnWebsite"
		Set rsCheckFSResolutionCodes = cnnCheckFSResolutionCodes.Execute(SQLCheckFSResolutionCodes)
	End If

		
	set rsCheckFSResolutionCodes = nothing
	cnnCheckFSResolutionCodes.close
	set cnnCheckFSResolutionCodes = nothing
				
%>