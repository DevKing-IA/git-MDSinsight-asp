<%	
	Set cnnTableDrop = Server.CreateObject("ADODB.Connection")
	cnnTableDrop.open (Session("ClientCnnString"))
	Set rsTableDrop = Server.CreateObject("ADODB.Recordset")
	rsTableDrop.CursorLocation = 3 

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'EQ_ScheduledServiceDates')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE EQ_ScheduledServiceDates;"

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'CustCatGroupPeriodSales')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE CustCatGroupPeriodSales;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)
	
	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'GL_ChartOfAccounts')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE GL_ChartOfAccounts;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'GL_AccountTypes')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE GL_AccountTypes;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'tblUser')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE tblUser;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)
	
	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Settings_FServCalendar')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE Settings_FServCalendar;"
		
	'Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'API_IC_AuditLog')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE API_IC_AuditLog;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Email_Attachments')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE Email_Attachments;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Email_Master')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE Email_Master;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'SC_Email_Master')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE SC_Email_Master;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'SC_Email_Attachments')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE SC_Email_Attachments;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'IN_InvoiceExportSage')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE IN_InvoiceExportSage;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'tblDebugLog')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE tblDebugLog;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'ACC_BudgetIncomeCategories')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE ACC_BudgetIncomeCategories;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)
	
	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'SC_ExternalEmailMessages')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE SC_ExternalEmailMessages;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'SC_ExternalEmailMessagesBody')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE SC_ExternalEmailMessagesBody;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PR_CustomerContacts')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE PR_CustomerContacts;"
		
	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)
			
	SQLTableDrop = "IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'tblUsers_Teams')"
    SQLTableDrop = SQLTableDrop  & "DROP TABLE tblUsers_Teams;"

	Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)

	set rsTableDrop = nothing
	cnnTableDrop.close
	set cnnTableDrop = nothing
	
				
%>