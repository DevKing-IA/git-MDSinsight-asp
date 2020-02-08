<%	

	'***************************************************************************
	 Response.Write("Creating Rules/Record Entries for AR_Customer<br><br>")
	'***************************************************************************
	
	Set cnnARCustomer = Server.CreateObject("ADODB.Connection")
	cnnARCustomer.open (Session("ClientCnnString"))
	Set rsARCustomer = Server.CreateObject("ADODB.Recordset")
	rsARCustomer.CursorLocation = 3 	


	'**************************************************************************************************************
	'FIRST CLEAR OUT THE ENTRIES IN THE SC_NEEDTOKNOW TABLE FOR ACCOUNTS RECEIVABLE
	'**************************************************************************************************************
	SQL_ARCustomer = "DELETE FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SubModule ='Address'"
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
			
			City = rsARCustomer("City")
			
			If City = "" OR IsNull(City) OR IsEmpty(City) OR Len(City) < 1 Then
				City = ""
			Else
				City = Replace(City,"'","''")
			End If

			State = rsARCustomer("State")
			Zip = rsARCustomer("Zip")			
			Phone = rsARCustomer("Phone")
		
			'*****************************************************************************************************************
			'Begin Validate Customer Number
			'*****************************************************************************************************************
			
			'Response.Write("Checking Accounts Receivable Customer Numbers.......<br>")
			
			 If CustNum = "" OR IsNull(CustNum) OR IsEmpty(CustNum) OR Len(CustNum) < 1 Then
				
				SCNeedToKnow_Module = "Accounts Receivable"
				SCNeedToKnow_SubModule = "Address"
				SCNeedToKnow_SummaryDescription = "Empty Customer/Account Number"
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
			'End Validate Customer Number
			'*****************************************************************************************************************
			
			
			'*****************************************************************************************************************
			'Begin Validate Customer Name
			'*****************************************************************************************************************
			
			'Response.Write("Checking Accounts Receivable Customer Names.......<br>")
			
			 If CustName = "" OR IsNull(CustName) OR IsEmpty(CustName) OR Len(CustName) < 1 Then
				
				SCNeedToKnow_Module = "Accounts Receivable"
				SCNeedToKnow_SubModule = "Address"
				SCNeedToKnow_SummaryDescription = "Empty Customer Name"
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
			'End Validate Customer Name
			'*****************************************************************************************************************
			
			
			'*****************************************************************************************************************
			'Begin Validate Customer Address 1
			'*****************************************************************************************************************
			
			'Response.Write("Checking Accounts Receivable Customer Address 1.......<br>")
			
			If ClientKey <> "1071" AND ClientKey <> "1071d" Then
			
				 If Addr1 = "" OR IsNull(Addr1) OR IsEmpty(Addr1) OR Len(Addr1) < 1 Then
					
					SCNeedToKnow_Module = "Accounts Receivable"
					SCNeedToKnow_SubModule = "Address"
					SCNeedToKnow_SummaryDescription = "Empty Address 1"
					SCNeedToKnow_DetailedDescription1 = "The first address field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify an address."
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
			End If
			'*****************************************************************************************************************
			'End Validate Customer Address 1
			'*****************************************************************************************************************



			
			'*****************************************************************************************************************
			'Begin Validate Customer Address 2 - CORPORATE COFFEE SYSTEMS ONLY!
			'*****************************************************************************************************************
			
			'Response.Write("Checking Accounts Receivable Customer Address 1.......<br>")
			If ClientKey = "1071" OR ClientKey = "1071d" Then
			
				 If Addr2 = "" OR IsNull(Addr2) OR IsEmpty(Addr2) OR Len(Addr2) < 1 Then
					
					SCNeedToKnow_Module = "Accounts Receivable"
					SCNeedToKnow_SubModule = "Address"
					SCNeedToKnow_SummaryDescription = "Empty Address 2"
					SCNeedToKnow_DetailedDescription1 = "The address field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify an address."
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
			 End If
			'*****************************************************************************************************************
			'End Validate Customer Address 2
			'*****************************************************************************************************************
	

			'*****************************************************************************************************************
			'Begin Validate Customer CityStateZip
			'*****************************************************************************************************************
			
			'Response.Write("Checking Accounts Receivable Customer CityStateZip.......<br>")
			
			 If CityStateZip = "" OR IsNull(CityStateZip) OR IsEmpty(CityStateZip) OR Len(CityStateZip) < 1 Then
				
				SCNeedToKnow_Module = "Accounts Receivable"
				SCNeedToKnow_SubModule = "Address"
				SCNeedToKnow_SummaryDescription = "Empty CityStateZip"
				SCNeedToKnow_DetailedDescription1 = "The CityStateZip field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify a CityStateZip."
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
				
					strid = CityStateZip
					Set reCityStateZip = New RegExp
					
					'***********************************************************
										
					'	^           # anchor to the start of the string
					'	([^,]+)     # match everything except a comma one or more times
					'	,           # match the comma itself
					'	\s          # match a single whitespace character
					'	([A-Z]{2})  # now match a two letter state code 
					'	\s        # match a single whitespace character
					'	([A-Z][0-9][A-VXY][\s][0-9][A-Z][0-9])   # match a Canadian zip code

					'***********************************************************

					With reCityStateZip					  
					  .Pattern = "^([^,]+),\s([A-Z]{2})\s([A-Z][0-9][A-Z][\s][0-9][A-Z][0-9])"
					  .IgnoreCase = False
					  .Global     = False
					End With
					
					' Test method returns TRUE if a match is found				
					If reCityStateZip.Test( strid ) Then
					    'Response.write(strid & " is a valid canadian CityStateZip")
					Else
						SCNeedToKnow_Module = "Accounts Receivable"
						SCNeedToKnow_SubModule = "Address"
						SCNeedToKnow_SummaryDescription = "Invalid CityStateZip"
						SCNeedToKnow_DetailedDescription1 = "The CityStateZip field (" & CityStateZip & ") for customer " & CustNum & " - " & CustName & " is invalid. A valid Canadian CityStateZip format is " & chr(34) & "MISSISSAUGA, ON L5M 1Y7" & chr(34) & "."
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
									
					Set reCityStateZip = Nothing
					
				ElseIf ClientCountry = "United States" Then
				
					strid = CityStateZip
					Set re = New RegExp
					
					
					'***********************************************************
										
					'	^           # anchor to the start of the string
					'	([^,]+)     # match everything except a comma one or more times
					'	,           # match the comma itself
					'	\s          # match a single whitespace character
					'	([A-Z]{2})  # now match a two letter state code 
					'	\s        # match a single whitespace character
					'	\d{5} - Ends in a five digit zip code.
					'	-?\d{4}? - Optionally matches the zip+4 format. It is not required
					
					'***********************************************************
					
					With re
					  .Pattern = "^([^,]+),\s([A-Z]{2})\s(\d{5}(?:[-\s]\d{4})?)"
					  .IgnoreCase = False
					  .Global     = False
					End With
					
					' Test method returns TRUE if a match is found
					
					If re.Test( strid ) Then
					    'Response.write(strid & " is a valid canadian CityStateZip")
					Else
						SCNeedToKnow_Module = "Accounts Receivable"
						SCNeedToKnow_SubModule = "Address"
						SCNeedToKnow_SummaryDescription = "Invalid CityStateZip"
						SCNeedToKnow_DetailedDescription1 = "The CityStateZip field (" & CityStateZip & ") for customer " & CustNum & " - " & CustName & " is invalid. A valid US CityStateZip format is " & chr(34) & "Beverly Hills, CA 90210" & chr(34) & "."
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
			'End Validate Customer CityStateZip
			'*****************************************************************************************************************
			
			
			
			
	

			'*****************************************************************************************************************
			'Begin Validate Customer City
			'*****************************************************************************************************************
			
			'Response.Write("Checking Accounts Receivable Customer City.......<br>")
			
			 If City = "" OR IsNull(City) OR IsEmpty(City) OR Len(City) < 1 Then
				
				SCNeedToKnow_Module = "Accounts Receivable"
				SCNeedToKnow_SubModule = "Address"
				SCNeedToKnow_SummaryDescription = "Empty City"
				SCNeedToKnow_DetailedDescription1 = "The city field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify a city."
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
			
			  strid = City
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
					SCNeedToKnow_Module = "Accounts Receivable"
					SCNeedToKnow_SubModule = "Address"
					SCNeedToKnow_SummaryDescription = "Invalid City"
					SCNeedToKnow_DetailedDescription1 = "The city field (" & City & ") for customer " & CustNum & " - " & CustName & " is invalid. A city can only contain the letters a-z."
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
			'End Validate Customer City
			'*****************************************************************************************************************
			
			


			
			
	

			'*****************************************************************************************************************
			'Begin Validate Customer State
			'*****************************************************************************************************************
			
			'Response.Write("Checking Accounts Receivable Customer State.......<br>")
			
			 If State = "" OR IsNull(State) OR IsEmpty(State) OR Len(State) < 1 Then
				
				SCNeedToKnow_Module = "Accounts Receivable"
				SCNeedToKnow_SubModule = "Address"
				SCNeedToKnow_SummaryDescription = "Empty State"
				SCNeedToKnow_DetailedDescription1 = "The state field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify a state."
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
				
					  strid = State
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
							SCNeedToKnow_Module = "Accounts Receivable"
							SCNeedToKnow_SubModule = "Address"
							SCNeedToKnow_SummaryDescription = "Invalid State"
							SCNeedToKnow_DetailedDescription1 = "The province field (" & State & ") for customer " & CustNum & " - " & CustName & " is invalid. A province must be a capitalized 2 character abbreviation."
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
					

					  strid = State
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
							SCNeedToKnow_Module = "Accounts Receivable"
							SCNeedToKnow_SubModule = "Address"
							SCNeedToKnow_SummaryDescription = "Invalid State"
							SCNeedToKnow_DetailedDescription1 = "The state field (" & State & ") for customer " & CustNum & " - " & CustName & " is invalid. A state must be a capitalized 2 character abbreviation."
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
			'End Validate Customer State
			'*****************************************************************************************************************
			
		
			
			
	

			'*****************************************************************************************************************
			'Begin Validate Customer Zip
			'*****************************************************************************************************************
			
			'Response.Write("Checking Accounts Receivable Customer Zip.......<br>")
			
			 If Zip = "" OR IsNull(Zip) OR IsEmpty(Zip) OR Len(Zip) < 1 Then
				
				SCNeedToKnow_Module = "Accounts Receivable"
				SCNeedToKnow_SubModule = "Address"
				SCNeedToKnow_SummaryDescription = "Empty Zip"
				SCNeedToKnow_DetailedDescription1 = "The zip code field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify a zip code."
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
				
					strid = Zip
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
						SCNeedToKnow_Module = "Accounts Receivable"
						SCNeedToKnow_SubModule = "Address"
						SCNeedToKnow_SummaryDescription = "Invalid Zip Code"
						SCNeedToKnow_DetailedDescription1 = "The zip code field (" & Zip & ") for customer " & CustNum & " - " & CustName & " is invalid. A valid Canadian zip code is the format " & chr(34) & "T2X 1V4" & chr(34) & " or " & chr(34) & "T2X1V4" & chr(34) & "."
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
				
					strid = Zip
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
						SCNeedToKnow_Module = "Accounts Receivable"
						SCNeedToKnow_SubModule = "Address"
						SCNeedToKnow_SummaryDescription = "Invalid Zip Code"
						SCNeedToKnow_DetailedDescription1 = "The zip code field (" & Zip & ") for customer " & CustNum & " - " & CustName & " is invalid. A valid US format zip code format is " & chr(34) & "94105-0011" & chr(34) & "or " & chr(34) & "94105" & chr(34) & "."
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
			'End Validate Customer Zip
			'*****************************************************************************************************************
			
			
			
			
	
			
			'*****************************************************************************************************************
			'Begin Validate Customer Phone
			'*****************************************************************************************************************
			
			'Response.Write("Checking Accounts Receivable Customer Phone Number.......<br>")
			
			 If Phone = "" OR IsNull(Phone) OR IsEmpty(Phone) OR Len(Phone) < 1 Then
				
				SCNeedToKnow_Module = "Accounts Receivable"
				SCNeedToKnow_SubModule = "Address"
				SCNeedToKnow_SummaryDescription = "Empty Phone Number"
				SCNeedToKnow_DetailedDescription1 = "The phone number field for customer " & CustNum & " - " & CustName & " is empty. Every account must specify a phone number."
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
				
				strid = Phone
				Set re = New RegExp
				With re
				  .Pattern    = "^(\+\d{1,2}\s)?\(?\d{3}\)?[\s.-]\d{3}[\s.-]\d{4}$"
				  .IgnoreCase = False
				  .Global     = False
				End With
				
				' Test method returns TRUE if a match is found
				
				If re.Test( strid ) Then
				    'Response.write(strid & " is a valid phone number")
				Else
				
					strid2 = Phone
					Set re2 = New RegExp
					With re2
					  .Pattern    = "\d{10}$"
					  .IgnoreCase = False
					  .Global     = False
					End With
					
					If re2.Test( strid2 ) Then
					    'Response.write(strid2 & " is a valid phone number")
					Else
						SCNeedToKnow_Module = "Accounts Receivable"
						SCNeedToKnow_SubModule = "Address"
						SCNeedToKnow_SummaryDescription = "Invalid Phone Number"
						SCNeedToKnow_DetailedDescription1 = "The phone number field (" & Phone & ") for customer " & CustNum & " - " & CustName & " is invalid. A valid phone number format is 123-456-7890 or (123) 456-7890 or 123 456 7890 or 123.456.7890 or 1234567890."
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
			'End Validate Customer Phone
			'*****************************************************************************************************************
			
			Response.Write("<hr>")
			
			
		rsARCustomer.MoveNext
		Loop	
	
	End If	
	Set rsARCustomer = Nothing
	cnnARCustomer.Close
	Set cnnARCustomer = Nothing
	
							
%>