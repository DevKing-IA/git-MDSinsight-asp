<%	

	'****************************************************************************************
	 Response.Write("Begin Entries for AR_Customer Customer Type<br><br>")
	'****************************************************************************************
	
	Set cnnARCustomer = Server.CreateObject("ADODB.Connection")
	cnnARCustomer.open (Session("ClientCnnString"))
	Set rsARCustomer = Server.CreateObject("ADODB.Recordset")
	rsARCustomer.CursorLocation = 3 	


	'**************************************************************************************************************
	'FIRST CLEAR OUT THE ENTRIES IN THE SC_NEEDTOKNOW TABLE FOR ACCOUNTS RECEIVABLE
	'**************************************************************************************************************
	SQL_ARCustomer = "DELETE FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SubModule ='Customer Type'"
	Set rsARCustomer = cnnARCustomer.Execute(SQL_ARCustomer)
	'**************************************************************************************************************
	
	
	SQL_ARCustomer = "SELECT * FROM AR_Customer WHERE AcctStatus = 'A' ORDER BY CustNum ASC"
	Set rsARCustomer = cnnARCustomer.Execute(SQL_ARCustomer)
	
	If NOT rsARCustomer.EOF Then
	
		Do While NOT rsARCustomer.EOF
	
			CustNum = rsARCustomer("CustNum")
			CustName = rsARCustomer("Name")
			
			If CustName = "" OR IsNull(CustName) OR IsEmpty(CustName) OR Len(CustName) < 1 Then
				CustName = ""
			Else
				CustName = Replace(CustName,"'","''")
			End If
			
			Addr1 = rsARCustomer("Addr1")
			
			If Addr1 = "" OR IsNull(Addr1) OR IsEmpty(Addr1) OR Len(Addr1) < 1 Then
				Addr1 = ""
			Else
				Addr1 = Replace(Addr1,"'","''")
			End If
			
			Addr2 = rsARCustomer("Addr2")
			
			If Addr2 = "" OR IsNull(Addr2) OR IsEmpty(Addr2) OR Len(Addr2) < 1 Then
				Addr2 = ""
			Else
				Addr2 = Replace(Addr2,"'","''")
			End If

			CityStateZip = rsARCustomer("CityStateZip")
		
			If CityStateZip = "" OR IsNull(CityStateZip) OR IsEmpty(CityStateZip) OR Len(CityStateZip) < 1 Then
				CityStateZip = ""
			Else
				CityStateZip = Replace(CityStateZip,"'","''")
			End If

		
			
			SCNeedToKnow_Module = "Accounts Receivable"
			SCNeedToKnow_SubModule = "Customer Type"
			SCNeedToKnow_SummaryDescription = "Missing customer type"
			SCNeedToKnow_DetailedDescription1 = "The customer type for customer " & CustName & " (" & CustNum & ") located at " & Addr1 & ", " & CityStateZip & " is missing. Every account must have a valid customer type."
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
			
		rsARCustomer.MoveNext
		Loop	
	
	End If	
	Set rsARCustomer = Nothing
	cnnARCustomer.Close
	Set cnnARCustomer = Nothing

	'****************************************************************************************
	 Response.Write("End Entries for AR_Customer Customer Type<br><br>")
	'****************************************************************************************
	
							
%>