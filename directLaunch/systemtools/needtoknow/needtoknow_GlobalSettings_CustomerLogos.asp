<%

	'**********************************************************************************
	 Response.Write("Creating Logo File Entries for Global Customer Settings<br><br>")
	'**********************************************************************************
	
	Set cnnNeedToKnowLogo = Server.CreateObject("ADODB.Connection")
	cnnNeedToKnowLogo.open (Session("ClientCnnString"))
	Set rsNeedToKnowLogo = Server.CreateObject("ADODB.Recordset")
	rsNeedToKnowLogo.CursorLocation = 3 	


	'********************************************************************************
	'FIRST CLEAR OUT THE ENTRIES IN THE SC_NEEDTOKNOW TABLE FOR CLIENT LOGOS
	'********************************************************************************
	SQL_NeedToKnowLogo = "DELETE FROM SC_NeedToKnow WHERE Module = 'Global Settings' AND SummaryDescription ='Missing Client Logo File'"
	Set rsNeedToKnowLogo = cnnNeedToKnowLogo.Execute(SQL_NeedToKnowLogo)
	'********************************************************************************


	'*******************************************************************
	'Begin Client Logo File SC_NeedToKnow Analysis
	'*******************************************************************
	
	Response.Write("<br>******** BEGIN Processing Client Logo File SC_NeedToKnow For " & ClientKey & "************<br>")

	'******************************************

	serverName = Request.ServerVariables("SERVER_NAME")
	
	If serverName = "www.mdsinsight.com" Then serverName = "mdsinsight.com"
	
	'Response.Write("serverName: " & serverName & "<br>")
	
	pathForLogoFile = "C:\home\" & serverName & "\wwwroot\clientfiles\" & ClientKey & "\logos\logo.png"
	
	'Response.Write("path to check: " & pathForLogoFile & "<br>")
	
	
	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")
	
	If NOT fso.FileExists(pathForLogoFile) Then
	
		SCNeedToKnow_Module = "Global Settings"
		SCNeedToKnow_SummaryDescription = "Missing Client Logo File"
		SCNeedToKnow_DetailedDescription1 = "The logo file for " & ClientKey & " (" & CompanyName & ") is empty."
		SCNeedToKnow_InsightStaffOnly = 1
	
		'*****************************************************************************************************************
		'Check to see if record already exists in SC_NeedToKnow
		'*****************************************************************************************************************
		
		SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' "
		SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
		SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
		
		Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
		
		If rsSCNeedToKnowCheckIfExists.EOF Then
					
			SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
			SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SummaryDescription & "', "
			SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
			
			Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
			
			If QuietMode = False Then Response.Write("<strong>" & SCNeedToKnow_Module & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "</strong><br>")
		
		End If
		'*****************************************************************************************************************

	End If	


	Response.Write("<br>******** DONE Processing Client Logo File SC_NeedToKnow For " & ClientKey & "************<br>")
	
	

%>