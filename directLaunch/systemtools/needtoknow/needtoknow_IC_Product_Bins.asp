<%	

	'****************************************************************************************
	 Response.Write("Begin Entries for IC_Product Bins<br><br>")
	'****************************************************************************************
	
	Set cnnICProduct = Server.CreateObject("ADODB.Connection")
	cnnICProduct.open (Session("ClientCnnString"))
	Set rsICProduct = Server.CreateObject("ADODB.Recordset")
	rsICProduct.CursorLocation = 3 	


	'**************************************************************************************************************
	'FIRST CLEAR OUT THE ENTRIES IN THE SC_NEEDTOKNOW TABLE FOR Product Bins
	'**************************************************************************************************************
	SQL_ICProduct = "DELETE FROM SC_NeedToKnow WHERE Module = 'Inventory Control' AND SubModule ='Bins'"
	Set rsICProduct = cnnICProduct.Execute(SQL_ICProduct)
	'**************************************************************************************************************
	
	
	'**************************************************************************************************************
	'NOW OBTAIN ALLOWED DUPLICATE BIN LOCATIONS FROM SETTINGS_NEEDTOKNOW TABLE FOR Product Bins
	'**************************************************************************************************************
	SQL = "SELECT * FROM Settings_NeedToKnow"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		N2KInventoryAllowedDuplicateBins = rs("N2KInventoryAllowedDuplicateBins")
		If Not IsNull(N2KInventoryAllowedDuplicateBins) Then
			N2KInventoryAllowedDuplicateBinsArray = Split(N2KInventoryAllowedDuplicateBins,",")
		End If
		
		'TRIM ANY SPACES AROUND THE BIN NAMES IF THERE WERE NAMES FOUND
		If IsArray(N2KInventoryAllowedDuplicateBinsArray) Then
			For i = 0 to UBound(N2KInventoryAllowedDuplicateBinsArray)
				N2KInventoryAllowedDuplicateBinsArray(i) = trim(N2KInventoryAllowedDuplicateBinsArray(i))
			Next 
		End If
		
		
	End If
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	'**************************************************************************************************************
	

	'**************************************************************************************************************
	'CHECK ALL INVENTORIED AND PICKABLE PRODUCTS FOR BINS
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
			
			
			If prodCasePricing = "U" OR prodCasePricing = "N" Then
			
				prodUnitBin = rsICProduct("prodUnitBin")
				
				If prodUnitBin <> "" Then prodUnitBin = replace(prodUnitBin, "'", "")
				
				If prodUnitBin = "" OR IsNull(prodUnitBin) OR IsEmpty(prodUnitBin) OR Len(prodUnitBin) < 1 Then
				
					SCNeedToKnow_Module = "Inventory Control"
					SCNeedToKnow_SubModule = "Bins"
					SCNeedToKnow_SummaryDescription = "Blank Unit Bin"
					SCNeedToKnow_DetailedDescription1 = "The Unit Bin for product " & prodSKU & " (" & prodDescription & ") is blank. Every unit must have a Bin."
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
			
				prodCaseBin = rsICProduct("prodCaseBin")
				
				If prodCaseBin <> "" Then prodCaseBin = replace(prodCaseBin, "'", "")

				If prodCaseBin = "" OR IsNull(prodCaseBin) OR IsEmpty(prodCaseBin) OR Len(prodCaseBin) < 1 Then
				
					SCNeedToKnow_Module = "Inventory Control"
					SCNeedToKnow_SubModule = "Bins"
					SCNeedToKnow_SummaryDescription = "Blank Case Bin"
					SCNeedToKnow_DetailedDescription1 = "The Case Bin for product " & prodSKU & " (" & prodDescription & ") is blank. Every case must have a Bin."
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
			
	
					
			'***************************************************************************************
			'CHECK TO MAKE SURE THE CASE PRODUCT BIN IS NOT ALREADY ASSIGNED TO ANOTHER PRODUCT BIN
			'***************************************************************************************
			If prodCaseBin <> "" Then	
			
				ThisBinIsAnAllowedDuplciateBin = false
				
				If IsArray(N2KInventoryAllowedDuplicateBinsArray) Then
					For i = 0 to UBound(N2KInventoryAllowedDuplicateBinsArray)
						If prodCaseBin = N2KInventoryAllowedDuplicateBinsArray(i) Then
							ThisBinIsAnAllowedDuplciateBin = true
						End If	
					Next 
				End If
					
				If ThisBinIsAnAllowedDuplciateBin = false Then
					
					Set cnnICProductDuplicateBin = Server.CreateObject("ADODB.Connection")
					cnnICProductDuplicateBin.open (Session("ClientCnnString"))
					Set rsICProductDuplicateBin = Server.CreateObject("ADODB.Recordset")
					rsICProductDuplicateBin.CursorLocation = 3 		
							
					SQL_ICProductDuplicateBin = "SELECT * FROM IC_Product WHERE (prodUnitBin = '" & prodCaseBin & "'  OR prodCaseBin = '" & prodCaseBin & "'"
					SQL_ICProductDuplicateBin = SQL_ICProductDuplicateBin & ") AND prodSKU <> '" & prodSKU & "'"
					
					Set rsICProductDuplicateBin = cnnICProductDuplicateBin.Execute(SQL_ICProductDuplicateBin)	
		
					If NOT rsICProductDuplicateBin.EOF Then
							
						duplicateSKUList = ""
						
						Do While NOT rsICProductDuplicateBin.EOF
						
							currentProdSKU = rsICProductDuplicateBin("prodSKU")
							currentProdDescription = rsICProductDuplicateBin("prodDescription")
											
							If currentProdSKU <> "" Then
								currentProdSKU = replace(currentProdSKU, "'", "")
							End If
							
							If currentProdDescription <> "" Then
								currentProdDescription = replace(currentProdDescription, "'", "")
							End If
											
							duplicateSKUList = duplicateSKUList & " " & currentProdSKU  & ", "
							rsICProductDuplicateBin.MoveNext
						Loop
								
						duplicateSKUList = Left(duplicateSKUList, Len(duplicateSKUList) - 2)
						
						If Len(duplicateSKUList) >= 8000 Then
							duplicateSKUList = Left(duplicateSKUList, 7500) & "......."
						End If						
						
						SCNeedToKnow_Module = "Inventory Control"
						SCNeedToKnow_SubModule = "Bins"
						SCNeedToKnow_SummaryDescription = "Duplicate Unit or Case Bin"
						SCNeedToKnow_DetailedDescription1 =  "The Case Bin " & " (" & prodCaseBin & ") is also assigned to the following products: " & duplicateSKUList  ' Every Bin must be unique."
						SCNeedToKnow_InsightStaffOnly = 0
									
						'*****************************************************************************************************************
						'Check to see if record already exists in SC_NeedToKnow
						'*****************************************************************************************************************
						
						SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
						SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
						SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
						
						Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
						
						If rsSCNeedToKnowCheckIfExists.EOF Then
									
						SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly, prodSKUIfApplicable, prodBinIfApplicable) VALUES "
						SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
						SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly& ",'" & prodSKU  & "','" & prodCaseBin & "')"
						
						Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
						
						If QuietMode = False Then
							Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
							Response.Write("<hr>")
						End If
					
					End If
					'*****************************************************************************************************************
				
					Set rsICProductDuplicateBin = Nothing
					cnnICProductDuplicateBin.Close
					Set cnnICProductDuplicateBin = Nothing
					
				End If
		
			End If
			
			
			
		End If	

			'**************************************************************************************
			'CHECK TO MAKE SURE THE UNIT PRODUCT BIN IS NOT ALREADY ASSIGNED TO ANOTHER PRODUCT BIN
			'**************************************************************************************
			If prodUnitBin <> "" Then	
			
				ThisBinIsAnAllowedDuplciateBin = false
				
				If IsArray(N2KInventoryAllowedDuplicateBinsArray) Then
					For i = 0 to UBound(N2KInventoryAllowedDuplicateBinsArray)
						If prodUnitBin = N2KInventoryAllowedDuplicateBinsArray(i) Then
							ThisBinIsAnAllowedDuplciateBin = true
						End If	
					Next 
				End If
					
				If ThisBinIsAnAllowedDuplciateBin = false then
				
					Set cnnICProductDuplicateBin = Server.CreateObject("ADODB.Connection")
					cnnICProductDuplicateBin.open (Session("ClientCnnString"))
					Set rsICProductDuplicateBin = Server.CreateObject("ADODB.Recordset")
					rsICProductDuplicateBin.CursorLocation = 3 		
							
					SQL_ICProductDuplicateBin = "SELECT * FROM IC_Product WHERE (prodUnitBin = '" & prodUnitBin & "' OR prodCaseBin = '" & prodUnitBin & "'"
					SQL_ICProductDuplicateBin = SQL_ICProductDuplicateBin & ") AND prodSKU <> '" & prodSKU & "'"
					
					Set rsICProductDuplicateBin = cnnICProductDuplicateBin.Execute(SQL_ICProductDuplicateBin)	
		
					If NOT rsICProductDuplicateBin.EOF Then
							
						duplicateSKUList = ""
						
						Do While NOT rsICProductDuplicateBin.EOF
						
							currentProdSKU = rsICProductDuplicateBin("prodSKU")
							currentProdDescription = rsICProductDuplicateBin("prodDescription")
											
							If currentProdSKU <> "" Then
								currentProdSKU = replace(currentProdSKU, "'", "")
							End If
							
							If currentProdDescription <> "" Then
								currentProdDescription = replace(currentProdDescription, "'", "")
							End If
											
							duplicateSKUList = duplicateSKUList & " " & currentProdSKU  & ", "
							rsICProductDuplicateBin.MoveNext
						Loop
								
						duplicateSKUList = Left(duplicateSKUList, Len(duplicateSKUList) - 2)
						
						If Len(duplicateSKUList) >= 8000 Then
							duplicateSKUList = Left(duplicateSKUList, 7500) & "......."
						End If						
						
						SCNeedToKnow_Module = "Inventory Control"
						SCNeedToKnow_SubModule = "Bins"
						SCNeedToKnow_SummaryDescription = "Duplicate Unit or Case Bin"
						SCNeedToKnow_DetailedDescription1 =  "The Unit Bin " & " (" & prodUnitBin & ") is also assigned to the following products: " & duplicateSKUList ' Every Bin must be unique."
						SCNeedToKnow_InsightStaffOnly = 0
									
						'*****************************************************************************************************************
						'Check to see if record already exists in SC_NeedToKnow
						'*****************************************************************************************************************
						
						SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
						SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
						SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
						
						Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
						
						If rsSCNeedToKnowCheckIfExists.EOF Then
									
						SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly, prodSKUIfApplicable,prodBinIfApplicable) VALUES "
						SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
						SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly& ",'" & prodSKU  & "','" & prodUnitBin & "')"
						
						Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
						
						If QuietMode = False Then
							Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
							Response.Write("<hr>")
						End If
					
					End If
					'*****************************************************************************************************************
				
					Set rsICProductDuplicateBin = Nothing
					cnnICProductDuplicateBin.Close
					Set cnnICProductDuplicateBin = Nothing
		
				End If
				
			End If
			
		End If	
				
				
		rsICProduct.MoveNext
		Loop	
	
	End If
	
		
	Set rsICProduct = Nothing
	cnnICProduct.Close
	Set cnnICProduct = Nothing
								


	'****************************************************************************************
	 Response.Write("End Entries for IC_Product Bins<br><br>")
	'****************************************************************************************
	
							
%>