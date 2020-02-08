<%	
	Set cnnARCustomerCounts = Server.CreateObject("ADODB.Connection")
	cnnARCustomerCounts.open (Session("ClientCnnString"))
	Set rsARCustomerCounts = Server.CreateObject("ADODB.Recordset")
	rsARCustomerCounts.CursorLocation = 3 


		
		
	Err.Clear
	on error resume next
	Set rsARCustomerCounts = cnnARCustomerCounts.Execute("SELECT TOP 1 * FROM AR_CustomerCounts ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE AR_CustomerCounts ("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_AR_CustomerCounts_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "numTotalAccounts [int] NULL, "
			SQLBuild = SQLBuild & "numActiveAccounts [int] NULL, "
			SQLBuild = SQLBuild & "numInactiveAccounts [int] NULL, "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			
			Set rsARCustomerCounts = cnnARCustomerCounts.Execute(SQLBuild)
		End If
	End If
	
	
	set rsARCustomerCounts = nothing
	cnnARCustomerCounts.close
	set cnnARCustomerCounts = nothing
	
				
%>