<%	
	Set cnnARCustomerInactive = Server.CreateObject("ADODB.Connection")
	cnnARCustomerInactive.open (Session("ClientCnnString"))
	Set rsARCustomerInactive = Server.CreateObject("ADODB.Recordset")
	rsARCustomerInactive.CursorLocation = 3 


		
		
	Err.Clear
	on error resume next
	Set rsARCustomerInactive = cnnARCustomerInactive.Execute("SELECT TOP 1 * FROM AR_CustomerInactive ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE AR_CustomerInactive ("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_AR_CustomerInactive_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "CustID [varchar](255) NULL, "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			
			Set rsARCustomerInactive = cnnARCustomerInactive.Execute(SQLBuild)
		End If
	End If
	
	
	set rsARCustomerInactive = nothing
	cnnARCustomerInactive.close
	set cnnARCustomerInactive = nothing
	
				
%>