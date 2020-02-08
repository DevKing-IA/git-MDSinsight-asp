<%	

	' Special case
	' This table is an Mplex table, not needed by Insight for 
	' other backends. If it does not exist, then don't worry about it
	
	on error goto 0

	If MUV_READ("BackendSystem") = "Metroplex" Then 
			
		Set cnnCheckCustomerType = Server.CreateObject("ADODB.Connection")
		cnnCheckCustomerType.open (Session("ClientCnnString"))
		Set rsCheckCustomerType = Server.CreateObject("ADODB.Recordset")
		rsCheckCustomerType.CursorLocation = 3 
				
	
		SQL_CheckCustomerType = "SELECT COL_LENGTH('CustomerType', 'MemoMessagingFlag') AS IsItThere"
		Set rsCheckCustomerType = cnnCheckCustomerType.Execute(SQL_CheckCustomerType)
		If IsNull(rsCheckCustomerType("IsItThere")) Then
			SQL_CheckCustomerType = "ALTER TABLE CustomerType ADD MemoMessagingFlag char(1) NOT NULL DEFAULT 'N'"
			Set rsCheckCustomerType = cnnCheckCustomerType.Execute(SQL_CheckCustomerType)
		End If
			
		'*****************************************************************
		'If there are no 0, undefined records in the Customer Type table
		'create the 0 record
		'*****************************************************************
		On Error Resume Next
		rsCheckCustomerType.Close 
		On Error Goto 0
		rsCheckCustomerType.CursorLocation = 3
		SQL_CheckCustomerType = "SELECT * FROM CustomerType WHERE CustTypeSequence = 0"
		Set rsCheckCustomerType = cnnCheckCustomerType.Execute(SQL_CheckCustomerType)
		If rsCheckCustomerType.EOF Then
			SQL_CheckCustomerType = "INSERT INTO CustomerType ("
			SQL_CheckCustomerType = SQL_CheckCustomerType & "CustTypeSequence "
			SQL_CheckCustomerType = SQL_CheckCustomerType & ", Description"								
			SQL_CheckCustomerType = SQL_CheckCustomerType & ") VALUES ("
			SQL_CheckCustomerType = SQL_CheckCustomerType & "0"			
			SQL_CheckCustomerType = SQL_CheckCustomerType & ",'Undefined'"							
			SQL_CheckCustomerType = SQL_CheckCustomerType & ")"
			Set rsCheckCustomerType = cnnCheckCustomerType.Execute(SQL_CheckCustomerType)
		Else
			If rsCheckCustomerType("Description") = "" Then
				SQL_CheckCustomerType = "UPDATE CustomerType SET Description = 'Undefined' WHERE CustTypeSequence = 0"
				Set rsCheckCustomerType = cnnCheckCustomerType.Execute(SQL_CheckCustomerType)
			Elseif IsNull(rsCheckCustomerType("Description")) Then
				SQL_CheckCustomerType = "UPDATE CustomerType SET Description = 'Undefined' WHERE CustTypeSequence = 0"
				Set rsCheckCustomerType = cnnCheckCustomerType.Execute(SQL_CheckCustomerType)
			End If
		End If
				
		
		Set rsCheckCustomerType = Nothing
		cnnCheckCustomerType.Close
		Set cnnCheckCustomerType = Nothing
		
	End If				
%>