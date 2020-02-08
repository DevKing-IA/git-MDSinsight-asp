<%	
	Set cnnSettings_Quickbooks = Server.CreateObject("ADODB.Connection")
	cnnSettings_Quickbooks.open (Session("ClientCnnString"))
	Set rsSettings_Quickbooks = Server.CreateObject("ADODB.Recordset")
	rsSettings_Quickbooks.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsSettings_Quickbooks = cnnSettings_Quickbooks.Execute("SELECT TOP 1 * FROM Settings_Quickbooks")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLSettings_Quickbooks = "CREATE TABLE [Settings_Quickbooks]( "
			SQLSettings_Quickbooks = SQLSettings_Quickbooks & " [ImportCustomersFromQB] [int] NULL, "
			SQLSettings_Quickbooks = SQLSettings_Quickbooks & " [ImportCustomersUpdateOrReplace] [varchar](1) NULL "
			SQLSettings_Quickbooks = SQLSettings_Quickbooks & ") ON [PRIMARY]"
			Set rsSettings_Quickbooks = cnnSettings_Quickbooks.Execute(SQLSettings_Quickbooks)
			
			SSQLSettings_Quickbooks = "ALTER TABLE Settings_Quickbooks ADD CONSTRAINT [DF_Settings_Quickbooks_ImportCustomersFromQB]  DEFAULT ((0)) FOR [ImportCustomersFromQB]"
			Set rsSettings_Quickbooks = cnnSettings_Quickbooks.Execute(SQLSettings_Quickbooks)
			
			SQLSettings_Quickbooks = "ALTER TABLE Settings_Quickbooks ADD CONSTRAINT [DF_Settings_Quickbooks_ImportCustomersUpdateOrReplace]  DEFAULT ('R') FOR [ImportCustomersUpdateOrReplace]"
			Set rsSettings_Quickbooks = cnnSettings_Quickbooks.Execute(SQLSettings_Quickbooks)
			
		End If
	End If
	
	
	set rsSettings_Quickbooks = nothing
	cnnSettings_Quickbooks.close
	set cnnSettings_Quickbooks = nothing
				
%>