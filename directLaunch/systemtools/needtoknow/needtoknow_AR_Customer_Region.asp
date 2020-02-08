<%	

	'****************************************************************************************
	 Response.Write("Begin Entries for AR_Customer Region<br><br>")
	'****************************************************************************************
	
	Set cnnARCustomer = Server.CreateObject("ADODB.Connection")
	cnnARCustomer.open (Session("ClientCnnString"))
	Set rsARCustomer = Server.CreateObject("ADODB.Recordset")
	rsARCustomer.CursorLocation = 3 	


	'**************************************************************************************************************
	'FIRST CLEAR OUT THE ENTRIES IN THE SC_NEEDTOKNOW TABLE FOR ACCOUNTS RECEIVABLE
	'**************************************************************************************************************
	SQL_ARCustomer = "DELETE FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SubModule ='Region'"
	Set rsARCustomer = cnnARCustomer.Execute(SQL_ARCustomer)
	'**************************************************************************************************************
	

	SQL_ARCustomer = "SELECT * FROM AR_Customer WHERE AcctStatus = 'A' ORDER BY CustNum ASC"
	
	Response.Write(SQL_ARCustomer & "<br>")
	
	Set rsARCustomer = cnnARCustomer.Execute(SQL_ARCustomer)
	
	If NOT rsARCustomer.EOF Then
	
		Response.Write("111111111111111111111111111111111111111111111111111111<br>")
	
		Do While NOT rsARCustomer.EOF
	
			CustNum = rsARCustomer("CustNum")

			If GetCustRegionByCustID(CustNum) = "" Then
			
				Response.Write("2222222222222222222222222222222222222222222222222<br>")
			
				SCNeedToKnow_Module = "Accounts Receivable"
				SCNeedToKnow_SubModule = "Region"
				SCNeedToKnow_SummaryDescription = "No region for this customer"
				SCNeedToKnow_DetailedDescription1 = "The region  for customer " & CustName & " (" & CustNum & ") located at " & Addr1 & ", " & CityStateZip & " could not be determined."
				If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
				SCNeedToKnow_InsightStaffOnly = 0
				
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
				SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
				SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND CustIDIfApplicable = '" & SCNeedToKnow_CustID & "' "
				SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
				
				Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
				
				If rsSCNeedToKnowCheckIfExists.EOF Then
							
					SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, CustIDIfApplicable, InsightStaffOnly) VALUES "
					SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
					SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', '" & SCNeedToKnow_CustID & "', " & SCNeedToKnow_InsightStaffOnly & ")"
					
					Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
					
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				
				End If
				'*****************************************************************************************************************
			
			End If
			
		rsARCustomer.MoveNext
		Loop	
	
	End If	
	Set rsARCustomer = Nothing
	cnnARCustomer.Close
	Set cnnARCustomer = Nothing

	'****************************************************************************************
	 Response.Write("End Entries for AR_Customer Region<br><br>")
	'****************************************************************************************
	
							
%>