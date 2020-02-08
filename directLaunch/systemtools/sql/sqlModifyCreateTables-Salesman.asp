<%	

	' Special case
	' This table is an Mplex table, not needed by Insight for 
	' other backends. If it does not exist, then don't worry about it
	
	on error goto 0

	If MUV_READ("BackendSystem") = "Metroplex" Then 
			
		Set cnnCheckSalesman = Server.CreateObject("ADODB.Connection")
		cnnCheckSalesman.open (Session("ClientCnnString"))
		Set rsCheckSalesman = Server.CreateObject("ADODB.Recordset")
		rsCheckSalesman.CursorLocation = 3 
				
	
		SQL_CheckSalesman = "SELECT COL_LENGTH('Salesman', 'emailAuthId') AS IsItThere"
		Set rsCheckSalesman = cnnCheckSalesman.Execute(SQL_CheckSalesman)
		If IsNull(rsCheckSalesman("IsItThere")) Then
			SQL_CheckSalesman = "ALTER TABLE Salesman ADD emailAuthId varchar(50) NULL"
			Set rsCheckSalesman = cnnCheckSalesman.Execute(SQL_CheckSalesman)
		End If
			
		'***********************************************************
		'If there are no 0, undefined records in the Salesman table
		'create the 0 record
		'***********************************************************
		On Error Resume Next
		rsCheckSalesman.Close 
		On Error Goto 0
		rsCheckSalesman.CursorLocation = 3
		SQL_CheckSalesman = "SELECT * FROM Salesman WHERE SalesmanSequence= 0"
		Set rsCheckSalesman = cnnCheckSalesman.Execute(SQL_CheckSalesman)
		If rsCheckSalesman.EOF Then
			SQL_CheckSalesman = "INSERT INTO Salesman ("
			SQL_CheckSalesman = SQL_CheckSalesman & "SalesmanSequence "
			SQL_CheckSalesman = SQL_CheckSalesman & ", Name "								
			SQL_CheckSalesman = SQL_CheckSalesman & ") VALUES ("
			SQL_CheckSalesman = SQL_CheckSalesman & "0"			
			SQL_CheckSalesman = SQL_CheckSalesman & ",'Undefined'"							
			SQL_CheckSalesman = SQL_CheckSalesman & ")"
			Set rsCheckSalesman = cnnCheckSalesman.Execute(SQL_CheckSalesman)
		Else
			If rsCheckSalesman("Name") = "" Then
				SQL_CheckSalesman = "UPDATE Salesman SET [Name] = 'Undefined' WHERE SalesmanSequence= 0"
				Set rsCheckSalesman = cnnCheckSalesman.Execute(SQL_CheckSalesman)
			Elseif IsNull(rsCheckSalesman("Name")) Then
				SQL_CheckSalesman = "UPDATE Salesman SET [Name] = 'Undefined' WHERE SalesmanSequence= 0"
				Set rsCheckSalesman = cnnCheckSalesman.Execute(SQL_CheckSalesman)
			End If
		End If
		
		Set rsCheckSalesman = Nothing
		cnnCheckSalesman.Close
		Set cnnCheckSalesman = Nothing
		
End If					
%>
