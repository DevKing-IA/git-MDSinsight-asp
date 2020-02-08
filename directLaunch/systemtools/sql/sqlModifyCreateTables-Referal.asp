<%	

	' Special case
	' The referal table is an Mplex table, not needed by Insight for 
	' other backends. If it does not exist, then don't worry about it
	
	on error goto 0

	If MUV_READ("BackendSystem") = "Metroplex" Then 
			
		Set cnnCheckReferal = Server.CreateObject("ADODB.Connection")
		cnnCheckReferal.open (Session("ClientCnnString"))
		Set rsCheckReferal = Server.CreateObject("ADODB.Recordset")

		SQL_CheckReferal = "SELECT COL_LENGTH('Referal', 'Description2') AS IsItThere"
		Set rsCheckReferal = cnnCheckReferal.Execute(SQL_CheckReferal)
		If IsNull(rsCheckReferal("IsItThere")) Then
			SQL_CheckReferal = "ALTER TABLE Referal ADD Description2 varchar(255) NULL"
			Set rsCheckReferal = cnnCheckReferal.Execute(SQL_CheckReferal)
		End If
		
		'***********************************************************
		'If there are no 0, undefined records in the Referal table
		'create the 0 record
		'***********************************************************
		On Error Resume Next
		rsCheckReferal.Close 
		On Error Goto 0
		rsCheckReferal.CursorLocation = 3
		SQL_CheckReferal = "SELECT * FROM Referal WHERE ReferalCode = 0"
		Set rsCheckReferal = cnnCheckReferal.Execute(SQL_CheckReferal)
		If rsCheckReferal.EOF Then
			SQL_CheckReferal = "INSERT INTO Referal ("
			SQL_CheckReferal = SQL_CheckReferal & "ReferalCode "
			SQL_CheckReferal = SQL_CheckReferal & ", Name"				
			SQL_CheckReferal = SQL_CheckReferal & ", Description "	
			SQL_CheckReferal = SQL_CheckReferal & ", Description2 "					
			SQL_CheckReferal = SQL_CheckReferal & ") VALUES ("
			SQL_CheckReferal = SQL_CheckReferal & "0"			
			SQL_CheckReferal = SQL_CheckReferal & ",'Undefined'"
			SQL_CheckReferal = SQL_CheckReferal & ",'Undefined'"
			SQL_CheckReferal = SQL_CheckReferal & ",'Undefined'"							
			SQL_CheckReferal = SQL_CheckReferal & ")"
			Set rsCheckReferal = cnnCheckReferal.Execute(SQL_CheckReferal)
		Else
			If rsCheckReferal("Description2") = "" Then
				SQL_CheckReferal = "UPDATE Referal SET Description2 = 'Undefined' WHERE ReferalCode = 0"
				Set rsCheckReferal= cnnCheckReferal.Execute(SQL_CheckReferal)
			Elseif IsNull(rsCheckReferal("Description2")) Then
				SQL_CheckReferal = "UPDATE Referal SET Description2 = 'Undefined' WHERE ReferalCode = 0"
				Set rsCheckReferal= cnnCheckReferal.Execute(SQL_CheckReferal)
			End If
		End If

		Set rsCheckReferal = Nothing
		cnnCheckReferal.Close
		Set cnnCheckReferal = Nothing

End If		
				
%>