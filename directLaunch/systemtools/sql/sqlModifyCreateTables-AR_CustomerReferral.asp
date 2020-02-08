<%	
	Response.Write("sqlModifyCreateTables-AR_CustomerReferral.asp" & "<br>")
	On Error Goto 0

	Set cnnCheckARCustomerReferral = Server.CreateObject("ADODB.Connection")
	cnnCheckARCustomerReferral.open (Session("ClientCnnString"))
	Set rsCheckARCustomerReferral = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckARCustomerReferral = cnnCheckARCustomerReferral.Execute("SELECT TOP 1 * FROM AR_CustomerReferral")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckARCustomerReferral = "CREATE TABLE [AR_CustomerReferral]("
			SQLCheckARCustomerReferral = SQLCheckARCustomerReferral & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckARCustomerReferral = SQLCheckARCustomerReferral & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_CustomerReferral_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckARCustomerReferral = SQLCheckARCustomerReferral & " [ReferralName] [varchar](8000) NULL, "
			SQLCheckARCustomerReferral = SQLCheckARCustomerReferral & " [Description] [varchar](8000) NULL, "
			SQLCheckARCustomerReferral = SQLCheckARCustomerReferral & " [Description2] [varchar](8000) NULL "
			SQLCheckARCustomerReferral = SQLCheckARCustomerReferral & " ) ON [PRIMARY]"      
		   Set rsCheckARCustomerReferral = cnnCheckARCustomerReferral.Execute(SQLCheckARCustomerReferral)
		   
		End If
	End If

	'Special for the parts  file
	'Make sure code 0 is there
	SQLCheckARCustomerReferral = "SELECT * FROM AR_CustomerReferral WHERE InternalRecordIdentifier = 0"
	Set rsCheckARCustomerReferral = cnnCheckARCustomerReferral.Execute(SQLCheckARCustomerReferral)
	If rsCheckARCustomerReferral.EOF Then 
	
		SQLCheckARCustomerReferral = "SET IDENTITY_INSERT AR_CustomerReferral ON;"
		Set rsCheckARCustomerReferral = cnnCheckARCustomerReferral.Execute(SQLCheckARCustomerReferral)

		SQLCheckARCustomerReferral = SQLCheckARCustomerReferral & "INSERT INTO AR_CustomerReferral (InternalRecordIdentifier,ReferralName,Description,Description2) "
		SQLCheckARCustomerReferral = SQLCheckARCustomerReferral & " VALUES (0,'Undefined','Undefined','Undefined')"
		Response.Write(SQLCheckARCustomerReferral)
		Set rsCheckARCustomerReferral = cnnCheckARCustomerReferral.Execute(SQLCheckARCustomerReferral)
		
		SQLCheckARCustomerReferral = "SET IDENTITY_INSERT AR_CustomerReferral OFF;"
		Set rsCheckARCustomerReferral = cnnCheckARCustomerReferral.Execute(SQLCheckARCustomerReferral)
		
	End If

	
	set rsCheckARCustomerReferral = nothing
	cnnCheckARCustomerReferral.close
	set cnnCheckARCustomerReferral = nothing
				
%>