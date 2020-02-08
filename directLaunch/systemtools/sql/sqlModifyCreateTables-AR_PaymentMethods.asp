<%	
	Response.Write("sqlModifyCreateTables-AR_PaymentMethods.asp" & "<br>")
	On Error Goto 0

	Set cnnCheckARPaymentMethods = Server.CreateObject("ADODB.Connection")
	cnnCheckARPaymentMethods.open (Session("ClientCnnString"))
	Set rsCheckARPaymentMethods = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckARPaymentMethods = cnnCheckARPaymentMethods.Execute("SELECT TOP 1 * FROM AR_PaymentMethods")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckARPaymentMethods = "CREATE TABLE [AR_PaymentMethods]("
			SQLCheckARPaymentMethods = SQLCheckARPaymentMethods & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckARPaymentMethods = SQLCheckARPaymentMethods & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_PaymentMethods_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckARPaymentMethods = SQLCheckARPaymentMethods & " [PayMethDescription] [varchar](255) NULL,"
			SQLCheckARPaymentMethods = SQLCheckARPaymentMethods & " [DaysToClear] [int] NULL"
			SQLCheckARPaymentMethods = SQLCheckARPaymentMethods & " ) ON [PRIMARY]"      
		   Set rsCheckARPaymentMethods = cnnCheckARPaymentMethods.Execute(SQLCheckARPaymentMethods)
		   
		End If
	End If


	'Make sure code 0 is there
	SQLCheckARPaymentMethods = "SELECT * FROM AR_PaymentMethods WHERE InternalRecordIdentifier = 0"
	Set rsCheckARPaymentMethods = cnnCheckARPaymentMethods.Execute(SQLCheckARPaymentMethods)
	If rsCheckARPaymentMethods.EOF Then 
	
		SQLCheckARPaymentMethods = "SET IDENTITY_INSERT AR_PaymentMethods ON;"
		Set rsCheckARPaymentMethods = cnnCheckARPaymentMethods.Execute(SQLCheckARPaymentMethods)

		SQLCheckARPaymentMethods = SQLCheckARPaymentMethods & "INSERT INTO AR_PaymentMethods (InternalRecordIdentifier,PayMethDescription,DaysToClear) "
		SQLCheckARPaymentMethods = SQLCheckARPaymentMethods & " VALUES (0,'Undefined',999)"
		Response.Write(SQLCheckARPaymentMethods)
		Set rsCheckARPaymentMethods = cnnCheckARPaymentMethods.Execute(SQLCheckARPaymentMethods)
		
		SQLCheckARPaymentMethods = "SET IDENTITY_INSERT AR_PaymentMethods OFF;"
		Set rsCheckARPaymentMethods = cnnCheckARPaymentMethods.Execute(SQLCheckARPaymentMethods)
		
	End If

	set rsCheckARPaymentMethods = nothing
	cnnCheckARPaymentMethods.close
	set cnnCheckARPaymentMethods = nothing
				
%>