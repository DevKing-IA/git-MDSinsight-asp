<%	
	Set cnnCheckFSProblemCodes = Server.CreateObject("ADODB.Connection")
	cnnCheckFSProblemCodes.open (Session("ClientCnnString"))
	Set rsCheckFSProblemCodes = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckFSProblemCodes = cnnCheckFSProblemCodes.Execute("SELECT TOP 1 * FROM FS_ProblemCodes")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckFSProblemCodes = "CREATE TABLE [FS_ProblemCodes]("
			SQLCheckFSProblemCodes = SQLCheckFSProblemCodes & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckFSProblemCodes = SQLCheckFSProblemCodes & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_FS_ProblemCodes_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckFSProblemCodes = SQLCheckFSProblemCodes & " [ProblemDescription] [varchar](255) NULL,"
			SQLCheckFSProblemCodes = SQLCheckFSProblemCodes & " [ShowOnWebsite] [int] NULL"
			SQLCheckFSProblemCodes = SQLCheckFSProblemCodes & " ) ON [PRIMARY]"      

		   Set rsCheckFSProblemCodes = cnnCheckFSProblemCodes.Execute(SQLCheckFSProblemCodes)
		   
		End If
	End If

	on error goto 0

	SQLCheckFSProblemCodes  = "SELECT COL_LENGTH('FS_ProblemCodes', 'ShowOnWebsite') AS IsItThere"
	Set rsCheckFSProblemCodes = cnnCheckFSProblemCodes.Execute(SQLCheckFSProblemCodes)
	If IsNull(rsCheckFSProblemCodes("IsItThere")) Then
		SQLCheckFSProblemCodes = "ALTER TABLE FS_ProblemCodes ADD ShowOnWebsite int NULL"
		Set rsCheckFSProblemCodes = cnnCheckFSProblemCodes.Execute(SQLCheckFSProblemCodes)
		SQLCheckFSProblemCodes = "UPDATE FS_ProblemCodes SET ShowOnWebsite = 1"
		Set rsCheckFSProblemCodes = cnnCheckFSProblemCodes.Execute(SQLCheckFSProblemCodes)
	End If

	'Special for the problem code file
	'Make sure code 0 is there
	SQLCheckFSProblemCodes  = "SELECT * FROM FS_ProblemCodes WHERE InternalRecordIdentifier = 0"
	Set rsCheckFSProblemCodes = cnnCheckFSProblemCodes.Execute(SQLCheckFSProblemCodes)
	If rsCheckFSProblemCodes.EOF Then 
	
		SQLCheckFSProblemCodes = "SET IDENTITY_INSERT FS_ProblemCodes ON;"
		'Set rsCheckFSProblemCodes = cnnCheckFSProblemCodes.Execute(SQLCheckFSProblemCodes)

		SQLCheckFSProblemCodes = SQLCheckFSProblemCodes & "INSERT INTO FS_ProblemCodes (InternalRecordIdentifier,ProblemDescription,ShowOnWebsite) "
		SQLCheckFSProblemCodes = SQLCheckFSProblemCodes & " VALUES (0,'Other',1);"
		'Set rsCheckFSProblemCodes = cnnCheckFSProblemCodes.Execute(SQLCheckFSProblemCodes)
		
		SQLCheckFSProblemCodes = SQLCheckFSProblemCodes & "SET IDENTITY_INSERT FS_ProblemCodes OFF;"
		Set rsCheckFSProblemCodes = cnnCheckFSProblemCodes.Execute(SQLCheckFSProblemCodes)
		
	End If
		
	set rsCheckFSProblemCodes = nothing
	cnnCheckFSProblemCodes.close
	set cnnCheckFSProblemCodes = nothing
				
%>