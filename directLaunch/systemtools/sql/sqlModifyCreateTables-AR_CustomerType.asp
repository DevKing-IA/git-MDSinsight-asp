<%	
	Response.Write("sqlModifyCreateTables-AR_CustomerType.asp" & "<br>")
	On Error Goto 0

	Set cnnCheckARCustomerType = Server.CreateObject("ADODB.Connection")
	cnnCheckARCustomerType.open (Session("ClientCnnString"))
	Set rsCheckARCustomerType = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute("SELECT TOP 1 * FROM AR_CustomerType")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckARCustomerType = "CREATE TABLE [AR_CustomerType]("
			SQLCheckARCustomerType = SQLCheckARCustomerType & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckARCustomerType = SQLCheckARCustomerType & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_CustomerType_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckARCustomerType = SQLCheckARCustomerType & " [TypeDescription] [varchar](8000) NULL"
			SQLCheckARCustomerType = SQLCheckARCustomerType & " ) ON [PRIMARY]"      
		   Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
		   
		End If
	End If

	'Special for the parts  file
	'Make sure code 0 is there
	SQLCheckARCustomerType = "SELECT * FROM AR_CustomerType WHERE InternalRecordIdentifier = 0"
	Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	If rsCheckARCustomerType.EOF Then 
	
		SQLCheckARCustomerType = "SET IDENTITY_INSERT AR_CustomerType ON;"
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)

		SQLCheckARCustomerType = SQLCheckARCustomerType & "INSERT INTO AR_CustomerType (InternalRecordIdentifier,TypeDescription) "
		SQLCheckARCustomerType = SQLCheckARCustomerType & " VALUES (0,'Undefined')"
		Response.Write(SQLCheckARCustomerType)
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
		
		SQLCheckARCustomerType = "SET IDENTITY_INSERT AR_CustomerType OFF;"
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
		
	End If

	SQLCheckARCustomerType = "SELECT COL_LENGTH('AR_CustomerType', 'IvsComment1') AS IsItThere"
	Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	If IsNull(rsCheckARCustomerType("IsItThere")) Then
		SQLCheckARCustomerType = "ALTER TABLE AR_CustomerType ADD IvsComment1 varchar(1000) NULL"
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	End If

	SQLCheckARCustomerType = "SELECT COL_LENGTH('AR_CustomerType', 'IvsComment2') AS IsItThere"
	Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	If IsNull(rsCheckARCustomerType("IsItThere")) Then
		SQLCheckARCustomerType = "ALTER TABLE AR_CustomerType ADD IvsComment2 varchar(1000) NULL"
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	End If
	
	SQLCheckARCustomerType = "SELECT COL_LENGTH('AR_CustomerType', 'IvsComment3') AS IsItThere"
	Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	If IsNull(rsCheckARCustomerType("IsItThere")) Then
		SQLCheckARCustomerType = "ALTER TABLE AR_CustomerType ADD IvsComment3 varchar(1000) NULL"
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	End If
	
	SQLCheckARCustomerType = "SELECT COL_LENGTH('AR_CustomerType', 'IvsComment4') AS IsItThere"
	Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	If IsNull(rsCheckARCustomerType("IsItThere")) Then
		SQLCheckARCustomerType = "ALTER TABLE AR_CustomerType ADD IvsComment4 varchar(1000) NULL"
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	End If
	
	SQLCheckARCustomerType = "SELECT COL_LENGTH('AR_CustomerType', 'IvsComment5') AS IsItThere"
	Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	If IsNull(rsCheckARCustomerType("IsItThere")) Then
		SQLCheckARCustomerType = "ALTER TABLE AR_CustomerType ADD IvsComment5 varchar(1000) NULL"
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	End If
	
	SQLCheckARCustomerType = "SELECT COL_LENGTH('AR_CustomerType', 'HoldDays') AS IsItThere"
	Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	If IsNull(rsCheckARCustomerType("IsItThere")) Then
		SQLCheckARCustomerType = "ALTER TABLE AR_CustomerType ADD HoldDays int NULL"
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	End If
	
	SQLCheckARCustomerType = "SELECT COL_LENGTH('AR_CustomerType', 'HoldAmt') AS IsItThere"
	Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	If IsNull(rsCheckARCustomerType("IsItThere")) Then
		SQLCheckARCustomerType = "ALTER TABLE AR_CustomerType ADD HoldAmt money NULL"
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	End If
	
	SQLCheckARCustomerType = "SELECT COL_LENGTH('AR_CustomerType', 'WholesaleFlag') AS IsItThere"
	Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	If IsNull(rsCheckARCustomerType("IsItThere")) Then
		SQLCheckARCustomerType = "ALTER TABLE AR_CustomerType ADD WholesaleFlag varchar(255) NULL"
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	End If
	
	SQLCheckARCustomerType = "SELECT COL_LENGTH('AR_CustomerType', 'MemoMessagingFlag') AS IsItThere"
	Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	If IsNull(rsCheckARCustomerType("IsItThere")) Then
		SQLCheckARCustomerType = "ALTER TABLE AR_CustomerType ADD MemoMessagingFlag varchar(255) NULL"
		Set rsCheckARCustomerType = cnnCheckARCustomerType.Execute(SQLCheckARCustomerType)
	End If
	
	set rsCheckARCustomerType = nothing
	cnnCheckARCustomerType.close
	set cnnCheckARCustomerType = nothing
				
%>