<%	

	'****************************************************************************************
	 Response.Write("Begin Entries for IC_Product UPC Codes<br><br>")
	'****************************************************************************************
	
	Set cnnICProduct = Server.CreateObject("ADODB.Connection")
	cnnICProduct.open (Session("ClientCnnString"))
	Set rsICProduct = Server.CreateObject("ADODB.Recordset")
	rsICProduct.CursorLocation = 3 	


	'**************************************************************************************************************
	'FIRST CLEAR OUT THE ENTRIES IN THE SC_NEEDTOKNOW TABLE FOR PRODUCT UPC CODES
	'**************************************************************************************************************
	SQL_ICProduct = "DELETE FROM SC_NeedToKnow WHERE Module = 'Inventory Control' AND SubModule ='Product UPC Codes'"
	Set rsICProduct = cnnICProduct.Execute(SQL_ICProduct)
	'**************************************************************************************************************

	'**************************************************************************************************************
	'CHECK ALL INVENTORIED AND PICKABLE PRODUCTS FOR UPC CODES
	'**************************************************************************************************************
	
	SQL_ICProduct = "SELECT * FROM IC_Product WHERE prodInventoriedItem = 'True' OR prodPickableItem = 'True' OR prodInventoriedItem = 1 OR prodPickableItem = 1 ORDER BY prodSKU"
	If QuietMode = False Then
		Response.Write("<br><br><br>" & SQL_ICProduct & "<br>")
	End If

	Set rsICProduct = cnnICProduct.Execute(SQL_ICProduct)
	
	If NOT rsICProduct.EOF Then
	
		Do While NOT rsICProduct.EOF
	
			prodSKU = rsICProduct("prodSKU")
			
		
			If prodCasePricing = "U" or prodCasePricing = "N" Then
				prodDescription= rsICProduct("prodDescription")
			ElseIf prodCasePricing = "U" Then
				prodDescription= rsICProduct("prodCaseDescription")
			End If
			
			If prodSKU <> "" Then
				prodSKU = replace(prodSKU, "'", "")
			End If
			
			If prodDescription <> "" Then
				prodDescription = replace(prodDescription, "'", "")
			End If
			
			prodCasePricing = rsICProduct("prodCasePricing")
			
			If prodCasePricing <> "" Then
				prodCasePricing = replace(prodCasePricing, "'", "")
			End If
			
			
			If prodCasePricing = "U" Then
			
				prodUnitUPC = rsICProduct("prodUnitUPC")
				
				If prodUnitUPC <> "" Then prodUnitUPC = replace(prodUnitUPC, "'", "")
				
				If prodUnitUPC = "" OR IsNull(prodUnitUPC) OR IsEmpty(prodUnitUPC) OR Len(prodUnitUPC) < 1 Then
				
					SCNeedToKnow_Module = "Inventory Control"
					SCNeedToKnow_SubModule = "Product UPC Codes"
					SCNeedToKnow_SummaryDescription = "Blank Unit UPC Code"
					SCNeedToKnow_DetailedDescription1 = "The Unit UPC Code for product " & prodSKU & " (" & prodDescription & ") is blank. Every unit must have a UPC Code."
					SCNeedToKnow_InsightStaffOnly = 0
					
			
					'*****************************************************************************************************************
					'Check to see if record already exists in SC_NeedToKnow
					'*****************************************************************************************************************
					
					SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
					SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
					SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
					
					Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
					
					If rsSCNeedToKnowCheckIfExists.EOF Then
								
						SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly, prodSKUIfApplicable) VALUES "
						SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
						SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ",'" & prodSKU  & "')"
						
						Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
									
						If QuietMode = False Then
							Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
						End If
					
					End If
					'*****************************************************************************************************************
					
				End If
				
			End If
		
			
			If prodCasePricing = "C" Then
			
				prodCaseUPC = rsICProduct("prodCaseUPC")
				
				If prodCaseUPC <> "" Then prodCaseUPC = replace(prodCaseUPC, "'", "")

				If prodCaseUPC = "" OR IsNull(prodCaseUPC) OR IsEmpty(prodCaseUPC) OR Len(prodCaseUPC) < 1 Then
				
					SCNeedToKnow_Module = "Inventory Control"
					SCNeedToKnow_SubModule = "Product UPC Codes"
					SCNeedToKnow_SummaryDescription = "Blank Case UPC Code"
					SCNeedToKnow_DetailedDescription1 = "The Case UPC Code for product " & prodSKU & " (" & prodDescription & ") is blank. Every case must have a UPC Code."
					SCNeedToKnow_InsightStaffOnly = 0
			
					'*****************************************************************************************************************
					'Check to see if record already exists in SC_NeedToKnow
					'*****************************************************************************************************************
					
					SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
					SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
					SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
					
					Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
					
					If rsSCNeedToKnowCheckIfExists.EOF Then
								
						SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly, prodSKUIfApplicable) VALUES "
						SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
						SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly& ",'" & prodSKU  & "')"
						
						Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
						
						If QuietMode = False Then
							Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
							Response.Write("<hr>")
						End If
							
					
					End If
					'*****************************************************************************************************************
					
				End If
				
			End If
			
		
			
			If prodCasePricing = "N" Then
			
				prodUnitUPC = rsICProduct("prodUnitUPC")
				prodCaseUPC = rsICProduct("prodCaseUPC")
				
				If prodUnitUPC <> "" Then prodUnitUPC = replace(prodUnitUPC, "'", "")
				
				If prodCaseUPC <> "" Then prodCaseUPC = replace(prodCaseUPC, "'", "")
			
				If (prodUnitUPC = "" OR IsNull(prodUnitUPC) OR IsEmpty(prodUnitUPC) OR Len(prodUnitUPC) < 1) AND (prodCaseUPC = "" OR IsNull(prodCaseUPC) OR IsEmpty(prodCaseUPC) OR Len(prodCaseUPC) < 1) Then
				
					SCNeedToKnow_Module = "Inventory Control"
					SCNeedToKnow_SubModule = "Product UPC Codes"
					SCNeedToKnow_SummaryDescription = "Blank Unit and Case UPC Code"
					SCNeedToKnow_DetailedDescription1 = "The Unit and Case UPC Code for the product " & prodSKU & " (" & prodDescription & ") is blank. Every product must have a UPC Code."
					SCNeedToKnow_InsightStaffOnly = 0
					
			
					'*****************************************************************************************************************
					'Check to see if record already exists in SC_NeedToKnow
					'*****************************************************************************************************************
					
					SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
					SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
					SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
					
					Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
					
					If rsSCNeedToKnowCheckIfExists.EOF Then
								
						SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly, prodSKUIfApplicable) VALUES "
						SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
						SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly& ",'" & prodSKU  & "')"
						
						Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
						
						If QuietMode = False Then
							Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
							Response.Write("<hr>")
						End If
					
					End If
					'*****************************************************************************************************************
					
				End If

				
			End If
						
					
			'***********************************************************************************
			'CHECK TO MAKE SURE THIS PRODUCT UPC IS NOT ALREADY ASSIGNED TO ANOTHER PRODUCT CASE
			'***********************************************************************************
			If prodCaseUPC <> "" Then	
			
				Set cnnICProductDuplicateUPC = Server.CreateObject("ADODB.Connection")
				cnnICProductDuplicateUPC.open (Session("ClientCnnString"))
				Set rsICProductDuplicateUPC = Server.CreateObject("ADODB.Recordset")
				rsICProductDuplicateUPC.CursorLocation = 3 		
						
				SQL_ICProductDuplicateUPC = "SELECT * FROM IC_Product WHERE (prodUnitUPC = '" & prodCaseUPC & "'  OR prodCaseUPC = '" & prodCaseUPC & "'"
				SQL_ICProductDuplicateUPC = SQL_ICProductDuplicateUPC & ") AND prodSKU <> '" & prodSKU & "'"
				
				Set rsICProductDuplicateUPC = cnnICProductDuplicateUPC.Execute(SQL_ICProductDuplicateUPC)	
	
				If NOT rsICProductDuplicateUPC.EOF Then
						
					duplicateSKUList = ""
					
					Do While NOT rsICProductDuplicateUPC.EOF
					
						currentProdSKU = rsICProductDuplicateUPC("prodSKU")
						currentProdDescription = rsICProductDuplicateUPC("prodDescription")
										
						If currentProdSKU <> "" Then
							currentProdSKU = replace(currentProdSKU, "'", "")
						End If
						
						If currentProdDescription <> "" Then
							currentProdDescription = replace(currentProdDescription, "'", "")
						End If
										
						duplicateSKUList = duplicateSKUList & " " & currentProdSKU  & ", "
						rsICProductDuplicateUPC.MoveNext
					Loop
							
					duplicateSKUList = Left(duplicateSKUList, Len(duplicateSKUList) - 2)
					
					If Len(duplicateSKUList) >= 8000 Then
						duplicateSKUList = Left(duplicateSKUList, 7500) & "......."
					End If						
					
					SCNeedToKnow_Module = "Inventory Control"
					SCNeedToKnow_SubModule = "Product UPC Codes"
					SCNeedToKnow_SummaryDescription = "Duplicate UPC Code"
					SCNeedToKnow_DetailedDescription1 =  "The Case UPC Code " & " (" & prodCaseUPC & ") is also assigned to the following products: " & duplicateSKUList  ' Every UPC Code must be unique."
					SCNeedToKnow_InsightStaffOnly = 0
								
					'*****************************************************************************************************************
					'Check to see if record already exists in SC_NeedToKnow
					'*****************************************************************************************************************
					
					SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
					SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
					SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
					
					Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
					
					If rsSCNeedToKnowCheckIfExists.EOF Then
								
					SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly, prodSKUIfApplicable, prodUPCIfApplicable) VALUES "
					SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
					SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly& ",'" & prodSKU  & "','" & prodCaseUPC & "')"
					
					Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
					
					If QuietMode = False Then
						Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
						Response.Write("<hr>")
					End If
				
				End If
				'*****************************************************************************************************************
			
				Set rsICProductDuplicateUPC = Nothing
				cnnICProductDuplicateUPC.Close
				Set cnnICProductDuplicateUPC = Nothing
		
			End If
			
		End If	

			'***********************************************************************************
			'CHECK TO MAKE SURE THIS PRODUCT UPC IS NOT ALREADY ASSIGNED TO ANOTHER PRODUCT UNIT
			'***********************************************************************************
			If prodUnitUPC <> "" Then	
			
				Set cnnICProductDuplicateUPC = Server.CreateObject("ADODB.Connection")
				cnnICProductDuplicateUPC.open (Session("ClientCnnString"))
				Set rsICProductDuplicateUPC = Server.CreateObject("ADODB.Recordset")
				rsICProductDuplicateUPC.CursorLocation = 3 		
						
				SQL_ICProductDuplicateUPC = "SELECT * FROM IC_Product WHERE (prodUnitUPC = '" & prodUnitUPC & "'  OR prodCaseUPC = '" & prodUnitUPC & "'"
				SQL_ICProductDuplicateUPC = SQL_ICProductDuplicateUPC & ") AND prodSKU <> '" & prodSKU & "'"
				
				Set rsICProductDuplicateUPC = cnnICProductDuplicateUPC.Execute(SQL_ICProductDuplicateUPC)	
	
				If NOT rsICProductDuplicateUPC.EOF Then
						
					duplicateSKUList = ""
					
					Do While NOT rsICProductDuplicateUPC.EOF
					
						currentProdSKU = rsICProductDuplicateUPC("prodSKU")
						currentProdDescription = rsICProductDuplicateUPC("prodDescription")
										
						If currentProdSKU <> "" Then
							currentProdSKU = replace(currentProdSKU, "'", "")
						End If
						
						If currentProdDescription <> "" Then
							currentProdDescription = replace(currentProdDescription, "'", "")
						End If
										
						duplicateSKUList = duplicateSKUList & " " & currentProdSKU  & ", "
						rsICProductDuplicateUPC.MoveNext
					Loop
							
					duplicateSKUList = Left(duplicateSKUList, Len(duplicateSKUList) - 2)
					
					If Len(duplicateSKUList) >= 8000 Then
						duplicateSKUList = Left(duplicateSKUList, 7500) & "......."
					End If						
					
					SCNeedToKnow_Module = "Inventory Control"
					SCNeedToKnow_SubModule = "Product UPC Codes"
					SCNeedToKnow_SummaryDescription = "Duplicate UPC Code"
					SCNeedToKnow_DetailedDescription1 =  "The Unit UPC Code " & " (" & prodUnitUPC & ") is also assigned to the following products: " & duplicateSKUList ' Every UPC Code must be unique."
					SCNeedToKnow_InsightStaffOnly = 0
								
					'*****************************************************************************************************************
					'Check to see if record already exists in SC_NeedToKnow
					'*****************************************************************************************************************
					
					SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
					SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
					SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
					
					Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
					
					If rsSCNeedToKnowCheckIfExists.EOF Then
								
					SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly, prodSKUIfApplicable,prodUPCIfApplicable) VALUES "
					SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
					SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly& ",'" & prodSKU  & "','" & prodUnitUPC & "')"
					
					Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
					
					If QuietMode = False Then
						Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
						Response.Write("<hr>")
					End If
				
				End If
				'*****************************************************************************************************************
			
				Set rsICProductDuplicateUPC = Nothing
				cnnICProductDuplicateUPC.Close
				Set cnnICProductDuplicateUPC = Nothing
		
			End If
			
		End If	
				
				
		rsICProduct.MoveNext
		Loop	
	
	End If
	
		
	Set rsICProduct = Nothing
	cnnICProduct.Close
	Set cnnICProduct = Nothing
								


	'****************************************************************************************
	 Response.Write("End Entries for IC_Product UPC Codes<br><br>")
	'****************************************************************************************
	
							
%>