<%	
	Set cnnCheckFSSymptomCodes = Server.CreateObject("ADODB.Connection")
	cnnCheckFSSymptomCodes.open (Session("ClientCnnString"))
	Set rsCheckFSSymptomCodes = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckFSSymptomCodes = cnnCheckFSSymptomCodes.Execute("SELECT TOP 1 * FROM FS_SymptomCodes")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckFSSymptomCodes = "CREATE TABLE [FS_SymptomCodes]("
			SQLCheckFSSymptomCodes = SQLCheckFSSymptomCodes & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckFSSymptomCodes = SQLCheckFSSymptomCodes & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_FS_SymptomCodes_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckFSSymptomCodes = SQLCheckFSSymptomCodes & " [SymptomDescription] [varchar](255) NULL,"
			SQLCheckFSSymptomCodes = SQLCheckFSSymptomCodes & " [ShowOnWebsite] [int] NULL"
			SQLCheckFSSymptomCodes = SQLCheckFSSymptomCodes & " ) ON [PRIMARY]"      

		   Set rsCheckFSSymptomCodes = cnnCheckFSSymptomCodes.Execute(SQLCheckFSSymptomCodes)
		   
		End If
	End If

	on error goto 0

	SQLCheckFSSymptomCodes  = "SELECT COL_LENGTH('FS_SymptomCodes', 'ShowOnWebsite') AS IsItThere"
	Set rsCheckFSSymptomCodes = cnnCheckFSSymptomCodes.Execute(SQLCheckFSSymptomCodes)
	If IsNull(rsCheckFSSymptomCodes("IsItThere")) Then
		SQLCheckFSSymptomCodes = "ALTER TABLE FS_SymptomCodes ADD ShowOnWebsite int NULL"
		Set rsCheckFSSymptomCodes = cnnCheckFSSymptomCodes.Execute(SQLCheckFSSymptomCodes)
		SQLCheckFSSymptomCodes = "UPDATE FS_SymptomCodes SET ShowOnWebsite = 1"
		Set rsCheckFSSymptomCodes = cnnCheckFSSymptomCodes.Execute(SQLCheckFSSymptomCodes)
	End If

	'Special for the Symptom code file
	'Make sure code 0 is there
	SQLCheckFSSymptomCodes  = "SELECT * FROM FS_SymptomCodes WHERE InternalRecordIdentifier = 0"
	Set rsCheckFSSymptomCodes = cnnCheckFSSymptomCodes.Execute(SQLCheckFSSymptomCodes)
	If rsCheckFSSymptomCodes.EOF Then 
	
		SQLCheckFSSymptomCodes = "SET IDENTITY_INSERT FS_SymptomCodes ON;"
		'Set rsCheckFSSymptomCodes = cnnCheckFSSymptomCodes.Execute(SQLCheckFSSymptomCodes)

		SQLCheckFSSymptomCodes = SQLCheckFSSymptomCodes & "INSERT INTO FS_SymptomCodes (InternalRecordIdentifier,SymptomDescription,ShowOnWebsite) "
		SQLCheckFSSymptomCodes = SQLCheckFSSymptomCodes & " VALUES (0,'Other',1);"
		'Set rsCheckFSSymptomCodes = cnnCheckFSSymptomCodes.Execute(SQLCheckFSSymptomCodes)
		
		SQLCheckFSSymptomCodes = SQLCheckFSSymptomCodes & "SET IDENTITY_INSERT FS_SymptomCodes OFF;"
		Set rsCheckFSSymptomCodes = cnnCheckFSSymptomCodes.Execute(SQLCheckFSSymptomCodes)
		
	End If
		
	set rsCheckFSSymptomCodes = nothing
	cnnCheckFSSymptomCodes.close
	set cnnCheckFSSymptomCodes = nothing
				
%>