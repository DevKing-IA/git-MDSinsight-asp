<%
'********************************
'List of all the functions & subs
'********************************
'Sub CheckIfTeamNameAlreadyExists()

action = Request("action")

Select Case action
	Case "CheckIfTeamNameAlreadyExists" 
		CheckIfTeamNameAlreadyExists()	
End Select




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub CheckIfTeamNameAlreadyExists()

	passedNewTeamName = Request.Form("passedNewTeamName")
	passedCurrTeamName = Request.Form("passedCurrTeamName")
	
	If passedCurrTeamName <> passedNewTeamName Then
		
		SQLCheckForDuplicateTeamName = "SELECT * FROM USER_TEAMS WHERE TeamName = '" & passedNewTeamName & "'"
			
		Set cnnCheckForDuplicateTeamName = Server.CreateObject("ADODB.Connection")
		cnnCheckForDuplicateTeamName.open(Session("ClientCnnString"))
		Set rsCheckForDuplicateTeamName = Server.CreateObject("ADODB.Recordset")
		rsCheckForDuplicateTeamName.CursorLocation = 3 
		
		Set rsCheckForDuplicateTeamName = cnnCheckForDuplicateTeamName.Execute(SQLCheckForDuplicateTeamName)
	
		If NOT rsCheckForDuplicateTeamName.EOF Then
			Response.Write("TEAMNAMEALREADYEXISTS")
		Else
			Response.Write("TEAMNAMENOTINUSE")
		End If
			
		Set rsCheckForDuplicateTeamName = Nothing
		cnnCheckForDuplicateTeamName.Close
		Set cnnCheckForDuplicateTeamName = Nothing
		
	Else
	
		Response.Write("TEAMNAMENOTCHANGED")
		
	End If
	

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


%>
