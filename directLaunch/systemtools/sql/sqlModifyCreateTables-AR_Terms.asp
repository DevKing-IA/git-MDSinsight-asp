<%	
	Response.Write("sqlModifyCreateTables-AR_Terms.asp" & "<br>")
	On Error Goto 0

	Set cnnCheckARTerms = Server.CreateObject("ADODB.Connection")
	cnnCheckARTerms.open (Session("ClientCnnString"))
	Set rsCheckARTerms = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckARTerms = cnnCheckARTerms.Execute("SELECT TOP 1 * FROM AR_Terms")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckARTerms = "CREATE TABLE [AR_Terms]("
			SQLCheckARTerms = SQLCheckARTerms & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckARTerms = SQLCheckARTerms & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_Terms_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckARTerms = SQLCheckARTerms & " [Description] [varchar](1000) NULL, "
			SQLCheckARTerms = SQLCheckARTerms & " [firstTermsPercent] [float] NULL, "
			SQLCheckARTerms = SQLCheckARTerms & " [firstTermsPeriod] [int] NULL, "
			SQLCheckARTerms = SQLCheckARTerms & " [secondTermsPeriod] [int] NULL, "
			SQLCheckARTerms = SQLCheckARTerms & " [TermsType] [varchar](255) NULL, "
			SQLCheckARTerms = SQLCheckARTerms & " [CreditCardBill] [varchar](255) NULL "
			SQLCheckARTerms = SQLCheckARTerms & " ) ON [PRIMARY]"      

		   Set rsCheckARTerms = cnnCheckARTerms.Execute(SQLCheckARTerms)
		   
		End If
	End If


	'Make sure code 0 is there
	SQLCheckARTerms = "SELECT * FROM AR_Terms WHERE InternalRecordIdentifier = 0"
	Set rsCheckARTerms = cnnCheckARTerms.Execute(SQLCheckARTerms)
	If rsCheckARTerms.EOF Then 
	
		SQLCheckARTerms = "SET IDENTITY_INSERT AR_Terms ON;"
		Set rsCheckARTerms = cnnCheckARTerms.Execute(SQLCheckARTerms)

		SQLCheckARTerms = SQLCheckARTerms & "INSERT INTO AR_Terms (InternalRecordIdentifier,Description) "
		SQLCheckARTerms = SQLCheckARTerms & " VALUES (0,'Undefined')"
		Response.Write(SQLCheckARTerms)
		Set rsCheckARTerms = cnnCheckARTerms.Execute(SQLCheckARTerms)
		
		SQLCheckARTerms = "SET IDENTITY_INSERT AR_Terms OFF;"
		Set rsCheckARTerms = cnnCheckARTerms.Execute(SQLCheckARTerms)
		
	End If

	set rsCheckARTerms = nothing
	cnnCheckARTerms.close
	set cnnCheckARTerms = nothing
%>