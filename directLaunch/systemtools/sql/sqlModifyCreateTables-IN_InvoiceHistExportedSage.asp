<%	

	Set cnnCheckINInvoiceHistExportedSage = Server.CreateObject("ADODB.Connection")
	cnnCheckINInvoiceHistExportedSage.open (Session("ClientCnnString"))
	Set rsCheckINInvoiceHistExportedSage = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckINInvoiceHistExportedSage = cnnCheckINInvoiceHistExportedSage.Execute("SELECT TOP 1 * FROM IN_InvoicesExportedSage")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckINInvoiceHistExportedSage = "CREATE TABLE [IN_InvoicesExportedSage]("
			SQLCheckINInvoiceHistExportedSage = SQLCheckINInvoiceHistExportedSage & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckINInvoiceHistExportedSage = SQLCheckINInvoiceHistExportedSage & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_IN_InvoicesExportedSage_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckINInvoiceHistExportedSage = SQLCheckINInvoiceHistExportedSage & " [InvoiceID] [varchar](255) NULL,"
			SQLCheckINInvoiceHistExportedSage = SQLCheckINInvoiceHistExportedSage & " ) ON [PRIMARY]"      

		   Set rsCheckINInvoiceHistExportedSage = cnnCheckINInvoiceHistExportedSage.Execute(SQLCheckINInvoiceHistExportedSage)
		   
		End If
	End If
	
	set rsCheckINInvoiceHistExportedSage = nothing
	cnnCheckINInvoiceHistExportedSage.close
	set cnnCheckINInvoiceHistExportedSage = nothing
		
%>