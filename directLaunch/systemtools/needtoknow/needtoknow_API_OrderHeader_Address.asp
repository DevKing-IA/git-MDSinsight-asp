<%	

	'************************************************************************************
	 Response.Write("Begin Creating Rules/Record Entries for API Order Address<br><br>")
	'************************************************************************************
	
	Set cnnAPIOrderHeader = Server.CreateObject("ADODB.Connection")
	cnnAPIOrderHeader.open (Session("ClientCnnString"))
	Set rsAPIOrderHeader = Server.CreateObject("ADODB.Recordset")
	rsAPIOrderHeader.CursorLocation = 3 	


	'**************************************************************************************************************
	'FIRST CLEAR OUT THE ENTRIES IN THE SC_NEEDTOKNOW TABLE FOR API
	'**************************************************************************************************************
	SQL_APIOrderHeader = "DELETE FROM SC_NeedToKnow WHERE Module = 'API' AND SubModule ='Orders'"
	Set rsAPIOrderHeader = cnnAPIOrderHeader.Execute(SQL_APIOrderHeader)
	'**************************************************************************************************************
	
	
	SQL_APIOrderHeader = "SELECT DISTINCT Orig_CustID, BillToCompany, BillToAddressLine1, BillToAddressLine2, BillToCity, BillToState, "
	SQL_APIOrderHeader = SQL_APIOrderHeader & " BillToZip, ShipToCompany, ShipToAddressLine1, ShipToAddressLine2, ShipToCity, ShipToState, ShipToZip "
	SQL_APIOrderHeader = SQL_APIOrderHeader & " FROM API_OR_OrderHeader "
	SQL_APIOrderHeader = SQL_APIOrderHeader & " WHERE (Orig_CustID IN "
	SQL_APIOrderHeader = SQL_APIOrderHeader & " (SELECT DISTINCT Orig_CustID FROM API_OR_OrderHeader AS API_OR_OrderHeader_1 "
	SQL_APIOrderHeader = SQL_APIOrderHeader & " WHERE (RecordCreationDateTime > DATEADD(d, - 9, GETDATE())))) AND (RecordCreationDateTime > DATEADD(d, - 9, GETDATE())) "
	SQL_APIOrderHeader = SQL_APIOrderHeader & " ORDER BY Orig_CustID "
		
	Set rsAPIOrderHeader = cnnAPIOrderHeader.Execute(SQL_APIOrderHeader)
	
	If NOT rsAPIOrderHeader.EOF Then
	
		Do While NOT rsAPIOrderHeader.EOF
	
			CustNum = rsAPIOrderHeader("Orig_CustID")
			CustName = rsAPIOrderHeader("BillToCompany")
			
			If CustName = "" OR IsNull(CustName) OR IsEmpty(CustName) OR Len(CustName) < 1 Then
				CustName = ""
			Else
				CustName = Replace(CustName,"'","''")
			End If
			
			BillAddr1 = rsAPIOrderHeader("BillToAddressLine1")

			If BillAddr1 = "" OR IsNull(BillAddr1) OR IsEmpty(BillAddr1) OR Len(BillAddr1) < 1 Then
				BillAddr1 = ""
			Else
				BillAddr1 = Replace(BillAddr1,"'","''")
			End If
						
			BillCity = rsAPIOrderHeader("BillToCity")
			
			If BillCity = "" OR IsNull(BillCity) OR IsEmpty(BillCity) OR Len(BillCity) < 1 Then
				BillCity = ""
			Else
				BillCity = Replace(BillCity,"'","''")
			End If
			
			BillState = rsAPIOrderHeader("BillToState")
			BillZip = rsAPIOrderHeader("BillToZip")			
	
			ShipAddr1 = rsAPIOrderHeader("ShipToAddressLine1")
			
			If ShipAddr1 = "" OR IsNull(ShipAddr1) OR IsEmpty(ShipAddr1) OR Len(ShipAddr1) < 1 Then
				ShipAddr1 = ""
			Else
				ShipAddr1 = Replace(ShipAddr1,"'","''")
			End If
			
			ShipCity = rsAPIOrderHeader("ShipToCity")
			
			If ShipCity = "" OR IsNull(ShipCity) OR IsEmpty(ShipCity) OR Len(ShipCity) < 1 Then
				ShipCity = ""
			Else
				ShipCity = Replace(ShipCity,"'","''")
			End If
			
			ShipState = rsAPIOrderHeader("ShipToState")
			ShipZip = rsAPIOrderHeader("ShipToZip")	
						
			'*****************************************************************************************************************
			'Begin Validate API Order Customer Number
			'*****************************************************************************************************************
			
			'Response.Write("Checking API Order Customer Numbers.......<br>")
			
			 If CustNum = "" OR IsNull(CustNum) OR IsEmpty(CustNum) OR Len(CustNum) < 1 Then
				
				SCNeedToKnow_Module = "API"
				SCNeedToKnow_SubModule = "Order"
				SCNeedToKnow_SummaryDescription = "Empty Billing Customer/Account Number"
				SCNeedToKnow_DetailedDescription1 = "The customer/account number for the customer " & CustName & " located at " & Addr1 & ", " & CityStateZip & " is empty. Every account must have a customer/account number."
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
				
				End If
				'*****************************************************************************************************************
				
		
			 End If
			'*****************************************************************************************************************
			'End Validate API Order Customer Number
			'*****************************************************************************************************************
			
			
			'*****************************************************************************************************************
			'Begin Validate Customer Billing Name
			'*****************************************************************************************************************
			
			'Response.Write("Checking API Order Customer Billing Name.......<br>")
			
			 If CustName = "" OR IsNull(CustName) OR IsEmpty(CustName) OR Len(CustName) < 1 Then
				
				SCNeedToKnow_Module = "API"
				SCNeedToKnow_SubModule = "Order"
				SCNeedToKnow_SummaryDescription = "Empty Billing Customer Name"
				SCNeedToKnow_DetailedDescription1 = "The name field for customer " & CustNum & " is empty. Every account must specify a customer name."
				If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
				SCNeedToKnow_InsightStaffOnly = 0
				
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
				
				End If
				'*****************************************************************************************************************
				
		
			 End If
			'*****************************************************************************************************************
			'End Validate API Order Customer Billing Name
			'*****************************************************************************************************************
			
			
			'*****************************************************************************************************************
			'Begin Validate API Order Customer Billing Address 1
			'*****************************************************************************************************************
			
			'Response.Write("Checking API Customer Billing Address 1.......<br>")
		
		
			 If BillAddr1 = "" OR IsNull(BillAddr1) OR IsEmpty(BillAddr1) OR Len(BillAddr1) < 1 Then
				
				SCNeedToKnow_Module = "API"
				SCNeedToKnow_SubModule = "Order"
				SCNeedToKnow_SummaryDescription = "Empty Billing Address 1"
				SCNeedToKnow_DetailedDescription1 = "The billing first address field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify an address."
				If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
				SCNeedToKnow_InsightStaffOnly = 0
				
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
				
				End If
				'*****************************************************************************************************************
		
			 End If
			'*****************************************************************************************************************
			'End Validate API Order Customer Billing Address 1
			'*****************************************************************************************************************
			

	

			'*****************************************************************************************************************
			'Begin Validate API Order Customer Billing City
			'*****************************************************************************************************************
			
			'Response.Write("Checking API Customer City.......<br>")
			
			 If BillCity = "" OR IsNull(BillCity) OR IsEmpty(BillCity) OR Len(BillCity) < 1 Then
				
				SCNeedToKnow_Module = "API"
				SCNeedToKnow_SubModule = "Order"
				SCNeedToKnow_SummaryDescription = "Empty Billing City"
				SCNeedToKnow_DetailedDescription1 = "The billing city field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify a city."
				If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
				SCNeedToKnow_InsightStaffOnly = 0
				
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
				
				End If
				'*****************************************************************************************************************
				
			Else
			
			  strid = BillCity
			  Set re = New RegExp
			  With re
			      .Pattern    = "[a-zA-Z]+"
			      .IgnoreCase = False
			      .Global     = False
			  End With
			  
			  ' Test method returns TRUE if a match is found
			  
			  If re.Test( strid ) Then
			        'Response.write(strid & " is a valid city name")
			  Else
					SCNeedToKnow_Module = "API"
					SCNeedToKnow_SubModule = "Order"
					SCNeedToKnow_SummaryDescription = "Invalid Billing City"
					SCNeedToKnow_DetailedDescription1 = "The billing city field (" & BillCity & ") for customer " & CustNum & " - " & CustName & " is invalid. A city can only contain the letters a-z."
					If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
					SCNeedToKnow_InsightStaffOnly = 0
	
					'*****************************************************************************************************************
					'Check to see if record already exists in SC_NeedToKnow
					'*****************************************************************************************************************
					
					SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
						
					End If
					
			  End If
			  
			  Set re = Nothing
				
		
			End If
			'*****************************************************************************************************************
			'End Validate API Order Customer City
			'*****************************************************************************************************************
			
			


			
			
	

			'*****************************************************************************************************************
			'Begin Validate API Order Customer Billing State
			'*****************************************************************************************************************
			
			'Response.Write("Checking API Customer Billing State.......<br>")
			
			 If BillState = "" OR IsNull(BillState) OR IsEmpty(BillState) OR Len(BillState) < 1 Then
				
				SCNeedToKnow_Module = "API"
				SCNeedToKnow_SubModule = "Order"
				SCNeedToKnow_SummaryDescription = "Empty Billing State"
				SCNeedToKnow_DetailedDescription1 = "The billing state field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify a state."
				If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
				SCNeedToKnow_InsightStaffOnly = 0
				
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
				
				End If
				'*****************************************************************************************************************
				
			Else
			
				If ClientCountry = "Canada" Then
				
					  strid = BillState
					  Set re = New RegExp
					  With re
					      .Pattern    = "^[A-Z]{2}$"
					      .IgnoreCase = False
					      .Global     = False
					  End With
					  
					  ' Test method returns TRUE if a match is found
					  
					  If re.Test( strid ) Then
					        'Response.write(strid & " is a valid canadian province abbreviation")
					  Else
							SCNeedToKnow_Module = "API"
							SCNeedToKnow_SubModule = "Order"
							SCNeedToKnow_SummaryDescription = "Invalid Billing State"
							SCNeedToKnow_DetailedDescription1 = "The billing province field (" & BillState & ") for customer " & CustNum & " - " & CustName & " is invalid. A province must be a capitalized 2 character abbreviation."
							If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
							SCNeedToKnow_InsightStaffOnly = 0
		
							'*****************************************************************************************************************
							'Check to see if record already exists in SC_NeedToKnow
							'*****************************************************************************************************************
							
							SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
								
							End If
						End If
							
					ElseIf ClientCountry = "United States" Then
					

					  strid = BillState
					  Set re = New RegExp
					  With re
					      .Pattern    = "^((A[LKSZR])|(C[AOT])|(D[EC])|(F[ML])|(G[AU])|(HI)|(I[DLNA])|(K[SY])|(LA)|(M[EHDAINSOT])|(N[EVHJMYCD])|(MP)|(O[HKR])|(P[WAR])|(RI)|(S[CD])|(T[NX])|(UT)|(V[TIA])|(W[AVIY]))$"
					      .IgnoreCase = False
					      .Global     = False
					  End With
					  
					  ' Test method returns TRUE if a match is found
					  
					  If re.Test( strid ) Then
					        'Response.write(strid & " is a valid unites states state abbreviation")
					  Else
							SCNeedToKnow_Module = "API"
							SCNeedToKnow_SubModule = "Order"
							SCNeedToKnow_SummaryDescription = "Invalid Billing State"
							SCNeedToKnow_DetailedDescription1 = "The billing state field (" & BillState & ") for customer " & CustNum & " - " & CustName & " is invalid. A state must be a capitalized 2 character abbreviation."
							If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
							SCNeedToKnow_InsightStaffOnly = 0
		
							'*****************************************************************************************************************
							'Check to see if record already exists in SC_NeedToKnow
							'*****************************************************************************************************************
							
							SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
								
							End If
							
						End If
					
			  	End If
			  
			  Set re = Nothing
		
			 End If
			'*****************************************************************************************************************
			'End Validate API Order Customer Billing State
			'*****************************************************************************************************************
			
		
			
			
	

			'*****************************************************************************************************************
			'Begin Validate API Order Customer Billing Zip
			'*****************************************************************************************************************
			
			'Response.Write("Checking API Order Customer Billing Zip.......<br>")
			
			 If BillZip = "" OR IsNull(BillZip) OR IsEmpty(BillZip) OR Len(BillZip) < 1 Then
				
				SCNeedToKnow_Module = "API"
				SCNeedToKnow_SubModule = "Order"
				SCNeedToKnow_SummaryDescription = "Empty Billing Zip"
				SCNeedToKnow_DetailedDescription1 = "The billing zip code field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify a zip code."
				If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
				SCNeedToKnow_InsightStaffOnly = 0
				
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
				
				End If
				'*****************************************************************************************************************
				
			Else
			
				If ClientCountry = "Canada" Then
				
					strid = BillZip
					Set re = New RegExp
					With re
					  .Pattern    = "^(?!.*[DFIOQU])[A-VXY][0-9][A-Z] ?[0-9][A-Z][0-9]$"
					  .IgnoreCase = False
					  .Global     = False
					End With
					
					' Test method returns TRUE if a match is found
					
					If re.Test( strid ) Then
					    'Response.write(strid & " is a valid canadian zip code")
					Else
						SCNeedToKnow_Module = "API"
						SCNeedToKnow_SubModule = "Order"
						SCNeedToKnow_SummaryDescription = "Invalid Billing Zip Code"
						SCNeedToKnow_DetailedDescription1 = "The billing zip code field (" & BillZip & ") for customer " & CustNum & " - " & CustName & " is invalid. A valid Canadian zip code is the format " & chr(34) & "T2X 1V4" & chr(34) & " or " & chr(34) & "T2X1V4" & chr(34) & "."
						If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
						SCNeedToKnow_InsightStaffOnly = 0

						'*****************************************************************************************************************
						'Check to see if record already exists in SC_NeedToKnow
						'*****************************************************************************************************************
						
						SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
							
						End If
						
					End If
					
					Set re = Nothing
					
				ElseIf ClientCountry = "United States" Then
				
					strid = BillZip
					Set re = New RegExp
					With re
					  .Pattern    = "^\d{5}(-\d{4})?$"
					  .IgnoreCase = False
					  .Global     = False
					End With
					
					' Test method returns TRUE if a match is found
					
					If re.Test( strid ) Then
					    'Response.write(strid & " is a valid united states zip code")
					Else
						SCNeedToKnow_Module = "API"
						SCNeedToKnow_SubModule = "Order"
						SCNeedToKnow_SummaryDescription = "Invalid Billing Zip Code"
						SCNeedToKnow_DetailedDescription1 = "The billing zip code field (" & BillZip & ") for customer " & CustNum & " - " & CustName & " is invalid. A valid US format zip code format is " & chr(34) & "94105-0011" & chr(34) & "or " & chr(34) & "94105" & chr(34) & "."
						If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
						SCNeedToKnow_InsightStaffOnly = 0
							
						'*****************************************************************************************************************
						'Check to see if record already exists in SC_NeedToKnow
						'*****************************************************************************************************************
						
						SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
							
						End If
						
					End If
					
					Set re = Nothing
				
				
				End If
		
			 End If
			'*****************************************************************************************************************
			'End Validate API Order Customer Billing Zip
			'*****************************************************************************************************************






			
			
			'*****************************************************************************************************************
			'Begin Validate API Order Customer Shipping Address 1
			'*****************************************************************************************************************
			
			'Response.Write("Checking API Customer Shipping Address 1.......<br>")
		
		
			 If ShipAddr1 = "" OR IsNull(ShipAddr1) OR IsEmpty(ShipAddr1) OR Len(ShipAddr1) < 1 Then
				
				SCNeedToKnow_Module = "API"
				SCNeedToKnow_SubModule = "Order"
				SCNeedToKnow_SummaryDescription = "Empty Shipping Address 1"
				SCNeedToKnow_DetailedDescription1 = "The shipping first address field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify an address."
				If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
				SCNeedToKnow_InsightStaffOnly = 0
				
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
				
				End If
				'*****************************************************************************************************************
		
			 End If
			'*****************************************************************************************************************
			'End Validate API Order Customer Shipping Address 1
			'*****************************************************************************************************************
			

	

			'*****************************************************************************************************************
			'Begin Validate API Order Customer Shipping City
			'*****************************************************************************************************************
			
			'Response.Write("Checking API Customer City.......<br>")
			
			 If ShipCity = "" OR IsNull(ShipCity) OR IsEmpty(ShipCity) OR Len(ShipCity) < 1 Then
				
				SCNeedToKnow_Module = "API"
				SCNeedToKnow_SubModule = "Order"
				SCNeedToKnow_SummaryDescription = "Empty Shipping City"
				SCNeedToKnow_DetailedDescription1 = "The shipping city field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify a city."
				If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
				SCNeedToKnow_InsightStaffOnly = 0
				
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
				
				End If
				'*****************************************************************************************************************
				
			Else
			
			  strid = ShipCity
			  Set re = New RegExp
			  With re
			      .Pattern    = "[a-zA-Z]+"
			      .IgnoreCase = False
			      .Global     = False
			  End With
			  
			  ' Test method returns TRUE if a match is found
			  
			  If re.Test( strid ) Then
			        'Response.write(strid & " is a valid city name")
			  Else
					SCNeedToKnow_Module = "API"
					SCNeedToKnow_SubModule = "Order"
					SCNeedToKnow_SummaryDescription = "Invalid Shipping City"
					SCNeedToKnow_DetailedDescription1 = "The shipping city field (" & ShipCity & ") for customer " & CustNum & " - " & CustName & " is invalid. A city can only contain the letters a-z."
					If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
					SCNeedToKnow_InsightStaffOnly = 0
	
					'*****************************************************************************************************************
					'Check to see if record already exists in SC_NeedToKnow
					'*****************************************************************************************************************
					
					SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
						
					End If
					
			  End If
			  
			  Set re = Nothing
				
		
			End If
			'*****************************************************************************************************************
			'End Validate API Order Customer City
			'*****************************************************************************************************************
			
			


			
			
	

			'*****************************************************************************************************************
			'Begin Validate API Order Customer Shipping State
			'*****************************************************************************************************************
			
			'Response.Write("Checking API Customer Shipping State.......<br>")
			
			 If ShipState = "" OR IsNull(ShipState) OR IsEmpty(ShipState) OR Len(ShipState) < 1 Then
				
				SCNeedToKnow_Module = "API"
				SCNeedToKnow_SubModule = "Order"
				SCNeedToKnow_SummaryDescription = "Empty Shipping State"
				SCNeedToKnow_DetailedDescription1 = "The shipping state field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify a state."
				If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
				SCNeedToKnow_InsightStaffOnly = 0
				
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
				
				End If
				'*****************************************************************************************************************
				
			Else
			
				If ClientCountry = "Canada" Then
				
					  strid = ShipState
					  Set re = New RegExp
					  With re
					      .Pattern    = "^[A-Z]{2}$"
					      .IgnoreCase = False
					      .Global     = False
					  End With
					  
					  ' Test method returns TRUE if a match is found
					  
					  If re.Test( strid ) Then
					        'Response.write(strid & " is a valid canadian province abbreviation")
					  Else
							SCNeedToKnow_Module = "API"
							SCNeedToKnow_SubModule = "Order"
							SCNeedToKnow_SummaryDescription = "Invalid Shipping State"
							SCNeedToKnow_DetailedDescription1 = "The shipping province field (" & ShipState & ") for customer " & CustNum & " - " & CustName & " is invalid. A province must be a capitalized 2 character abbreviation."
							If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
							SCNeedToKnow_InsightStaffOnly = 0
		
							'*****************************************************************************************************************
							'Check to see if record already exists in SC_NeedToKnow
							'*****************************************************************************************************************
							
							SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
								
							End If
						End If
							
					ElseIf ClientCountry = "United States" Then
					

					  strid = ShipState
					  Set re = New RegExp
					  With re
					      .Pattern    = "^((A[LKSZR])|(C[AOT])|(D[EC])|(F[ML])|(G[AU])|(HI)|(I[DLNA])|(K[SY])|(LA)|(M[EHDAINSOT])|(N[EVHJMYCD])|(MP)|(O[HKR])|(P[WAR])|(RI)|(S[CD])|(T[NX])|(UT)|(V[TIA])|(W[AVIY]))$"
					      .IgnoreCase = False
					      .Global     = False
					  End With
					  
					  ' Test method returns TRUE if a match is found
					  
					  If re.Test( strid ) Then
					        'Response.write(strid & " is a valid unites states state abbreviation")
					  Else
							SCNeedToKnow_Module = "API"
							SCNeedToKnow_SubModule = "Order"
							SCNeedToKnow_SummaryDescription = "Invalid Shipping State"
							SCNeedToKnow_DetailedDescription1 = "The shipping state field (" & ShipState & ") for customer " & CustNum & " - " & CustName & " is invalid. A state must be a capitalized 2 character abbreviation."
							If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
							SCNeedToKnow_InsightStaffOnly = 0
		
							'*****************************************************************************************************************
							'Check to see if record already exists in SC_NeedToKnow
							'*****************************************************************************************************************
							
							SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
								
							End If
							
						End If
					
			  	End If
			  
			  Set re = Nothing
		
			 End If
			'*****************************************************************************************************************
			'End Validate API Order Customer Shipping State
			'*****************************************************************************************************************
			
		
			
			
	

			'*****************************************************************************************************************
			'Begin Validate API Order Customer Shipping Zip
			'*****************************************************************************************************************
			
			'Response.Write("Checking API Order Customer Shipping Zip.......<br>")
			
			 If ShipZip = "" OR IsNull(ShipZip) OR IsEmpty(ShipZip) OR Len(ShipZip) < 1 Then
				
				SCNeedToKnow_Module = "API"
				SCNeedToKnow_SubModule = "Order"
				SCNeedToKnow_SummaryDescription = "Empty Shipping Zip"
				SCNeedToKnow_DetailedDescription1 = "The shipping zip code field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify a zip code."
				If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
				SCNeedToKnow_InsightStaffOnly = 0
				
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
				
				End If
				'*****************************************************************************************************************
				
			Else
			
				If ClientCountry = "Canada" Then
				
					strid = ShipZip
					Set re = New RegExp
					With re
					  .Pattern    = "^(?!.*[DFIOQU])[A-VXY][0-9][A-Z] ?[0-9][A-Z][0-9]$"
					  .IgnoreCase = False
					  .Global     = False
					End With
					
					' Test method returns TRUE if a match is found
					
					If re.Test( strid ) Then
					    'Response.write(strid & " is a valid canadian zip code")
					Else
						SCNeedToKnow_Module = "API"
						SCNeedToKnow_SubModule = "Order"
						SCNeedToKnow_SummaryDescription = "Invalid Shipping Zip Code"
						SCNeedToKnow_DetailedDescription1 = "The shipping zip code field (" & ShipZip & ") for customer " & CustNum & " - " & CustName & " is invalid. A valid Canadian zip code is the format " & chr(34) & "T2X 1V4" & chr(34) & " or " & chr(34) & "T2X1V4" & chr(34) & "."
						If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
						SCNeedToKnow_InsightStaffOnly = 0

						'*****************************************************************************************************************
						'Check to see if record already exists in SC_NeedToKnow
						'*****************************************************************************************************************
						
						SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
							
						End If
						
					End If
					
					Set re = Nothing
					
				ElseIf ClientCountry = "United States" Then
				
					strid = ShipZip
					Set re = New RegExp
					With re
					  .Pattern    = "^\d{5}(-\d{4})?$"
					  .IgnoreCase = False
					  .Global     = False
					End With
					
					' Test method returns TRUE if a match is found
					
					If re.Test( strid ) Then
					    'Response.write(strid & " is a valid united states zip code")
					Else
						SCNeedToKnow_Module = "API"
						SCNeedToKnow_SubModule = "Order"
						SCNeedToKnow_SummaryDescription = "Invalid Shipping Zip Code"
						SCNeedToKnow_DetailedDescription1 = "The shipping zip code field (" & ShipZip & ") for customer " & CustNum & " - " & CustName & " is invalid. A valid US format zip code format is " & chr(34) & "94105-0011" & chr(34) & "or " & chr(34) & "94105" & chr(34) & "."
						If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
						SCNeedToKnow_InsightStaffOnly = 0
							
						'*****************************************************************************************************************
						'Check to see if record already exists in SC_NeedToKnow
						'*****************************************************************************************************************
						
						SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "'"
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
							
						End If
						
					End If
					
					Set re = Nothing
				
				
				End If
		
			 End If
			'*****************************************************************************************************************
			'End Validate API Order Customer Shipping Zip
			'*****************************************************************************************************************




			
			Response.Write("<hr>")
			
			
		rsAPIOrderHeader.MoveNext
		Loop	
	
	End If	
	
	'************************************************************************************
	 Response.Write("End Creating Rules/Record Entries for API Order Address<br><br>")
	'************************************************************************************
	
	Set rsAPIOrderHeader = Nothing
	cnnAPIOrderHeader.Close
	Set cnnAPIOrderHeader = Nothing
	
							
%>