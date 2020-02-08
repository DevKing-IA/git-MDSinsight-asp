<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<!--#include file="InsightFuncs_InventoryControl.asp"-->

<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
 
'Sub SaveEquivalentSKUandUM()
'Sub ReturnProductEquivalentSKUs()
'Sub ReturnPartnerEquivalentSKUs()
'Func GenerateInventoryReportCSV()

'***************************************************
'End List of all the AJAX functions & subs
'***************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'ALL AJAX MODAL SUBROUTINES AND FUNCTIONS BELOW THIS AREA

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

action = Request("action")

Select Case action
	Case "SaveEquivalentSKUandUM" 
		SaveEquivalentSKUandUM()	
	Case "ReturnProductEquivalentSKUs"
		ReturnProductEquivalentSKUs()
	Case "ReturnPartnerEquivalentSKUs"
		ReturnPartnerEquivalentSKUs()
	Case "GenerateInventoryReportCSV"
		Response.Write(GenerateInventoryReportCSV())
End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub ReturnPartnerEquivalentSKUs()

	companySKUEnteredByUser = Request.Form("companySKU") 
	
	Set rsLookupPartnerEquviSKU= Server.CreateObject("ADODB.Recordset")
	rsLookupPartnerEquviSKU.CursorLocation = 3 
	
	If companySKUEnteredByUser <> "" Then	
	
		SQLLookupPartnerEquviSKU = "SELECT * FROM IC_ProductMapping WHERE SKU= '" & companySKUEnteredByUser & "' ORDER BY partnerIntRecID DESC"

		Set cnnLookupPartnerEquviSKU = Server.CreateObject("ADODB.Connection")
		cnnLookupPartnerEquviSKU.open (Session("ClientCnnString"))
		Set rsLookupPartnerEquviSKU = cnnLookupPartnerEquviSKU.Execute(SQLLookupPartnerEquviSKU)
		
		If NOT rsLookupPartnerEquviSKU.EOF Then
			%>
			<table>
			  <thead>
			    <tr>
			      <th>Partner</th>
			      <th>Partner Code</th>
			      <th>UoM</th>
			    </tr>
			  </thead>
			  <tbody>			
			<%
		
			Do While Not rsLookupPartnerEquviSKU.EOF
			
				
				partnerIntRecID = rsLookupPartnerEquviSKU("partnerIntRecID")
				partnerName = GetPartnerNameByIntRecID(partnerIntRecID)
				
				partnerEquivalentSKU1 = rsLookupPartnerEquviSKU("partnerEquivalentSKU1")
				partnerEquivalentSKU2 = rsLookupPartnerEquviSKU("partnerEquivalentSKU2")
				partnerEquivalentSKU3 = rsLookupPartnerEquviSKU("partnerEquivalentSKU3")
				partnerEquivalentSKU4 = rsLookupPartnerEquviSKU("partnerEquivalentSKU4")
				partnerEquivalentSKU5 = rsLookupPartnerEquviSKU("partnerEquivalentSKU5")
				partnerEquivalentSKU6 = rsLookupPartnerEquviSKU("partnerEquivalentSKU6")
				
				partnerUM = rsLookupPartnerEquviSKU("UM")
				%>
					<% If partnerEquivalentSKU1 <> "" Then %>
				    <tr>
				      <td data-label="Partner"><%= partnerName %> Code 1</td>
				      <td data-label="Partner Code"><%= partnerEquivalentSKU1 %></td>
				      <td data-label="UoM"><%= partnerUM %></td>
				    </tr>
				    <% End If %>	
					<% If partnerEquivalentSKU2 <> "" Then %>
				    <tr>
				      <td data-label="Partner"><%= partnerName %> Code 2</td>
				      <td data-label="Partner Code"><%= partnerEquivalentSKU2 %></td>
				      <td data-label="UoM"><%= partnerUM %></td>
				    </tr>
				    <% End If %>
					<% If partnerEquivalentSKU3 <> "" Then %>
				    <tr>
				      <td data-label="Partner"><%= partnerName %> Code 3</td>
				      <td data-label="Partner Code"><%= partnerEquivalentSKU3 %></td>
				      <td data-label="UoM"><%= partnerUM %></td>
				    </tr>
				    <% End If %>
					<% If partnerEquivalentSKU4 <> "" Then %>
				    <tr>
				      <td data-label="Partner"><%= partnerName %> Code 4</td>
				      <td data-label="Partner Code"><%= partnerEquivalentSKU4 %></td>
				      <td data-label="UoM"><%= partnerUM %></td>
				    </tr>
				    <% End If %>
					<% If partnerEquivalentSKU5 <> "" Then %>
				    <tr>
				      <td data-label="Partner"><%= partnerName %> Code 5</td>
				      <td data-label="Partner Code"><%= partnerEquivalentSKU5 %></td>
				      <td data-label="UoM"><%= partnerUM %></td>
				    </tr>
				    <% End If %>
					<% If partnerEquivalentSKU6 <> "" Then %>
				    <tr>
				      <td data-label="Partner"><%= partnerName %> Code 6</td>
				      <td data-label="Partner Code"><%= partnerEquivalentSKU6 %></td>
				      <td data-label="UoM"><%= partnerUM %></td>
				    </tr>
				    <% End If %>				    				    				    				    				    	
				<%
			rsLookupPartnerEquviSKU.MoveNext
			Loop
			%>
			  </tbody>
			</table>			
			<%
		Else
			Response.Write("No Partner Entry SKU was Found For " & companySKUEnteredByUser)
		End If
		
		
		set rsLookupPartnerEquviSKU = Nothing
		cnnLookupPartnerEquviSKU.close
		set cnnLookupPartnerEquviSKU = Nothing
		
	Else
		Response.Write("Cannot Lookup, Invalid Data")
		
	End If

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub ReturnProductEquivalentSKUs()

	partnerSKUEnteredByUser = Request.Form("partnerSKU")
	shortCompanyName = Request.Form("shortCompanyName") 
	
	Set rsLookupProductSKU= Server.CreateObject("ADODB.Recordset")
	rsLookupProductSKU.CursorLocation = 3 
	
	If partnerSKUEnteredByUser <> "" Then	

		SQLLookupProductSKU = "SELECT * FROM IC_ProductMapping WHERE partnerEquivalentSKU1 = '" & partnerSKUEnteredByUser & "' OR "
		SQLLookupProductSKU = SQLLookupProductSKU & "partnerEquivalentSKU2 = '" & partnerSKUEnteredByUser & "' OR "
		SQLLookupProductSKU = SQLLookupProductSKU & "partnerEquivalentSKU3 = '" & partnerSKUEnteredByUser & "' OR "
		SQLLookupProductSKU = SQLLookupProductSKU & "partnerEquivalentSKU4 = '" & partnerSKUEnteredByUser & "' OR "
		SQLLookupProductSKU = SQLLookupProductSKU & "partnerEquivalentSKU5 = '" & partnerSKUEnteredByUser & "' OR "
		SQLLookupProductSKU = SQLLookupProductSKU & "partnerEquivalentSKU6 = '" & partnerSKUEnteredByUser & "' ORDER BY partnerIntRecID DESC"
		
		Set cnnLookupProductSKU = Server.CreateObject("ADODB.Connection")
		cnnLookupProductSKU.open (Session("ClientCnnString"))
		Set rsLookupProductSKU = cnnLookupProductSKU.Execute(SQLLookupProductSKU)
		
		If NOT rsLookupProductSKU.EOF Then
			%>
			<table>
			  <thead>
			    <tr>
			      <th><%= shortCompanyName %> Code For PARTNER</th>
			      <th><%= shortCompanyName %> UoM For PARTNER</th>
			    </tr>
			  </thead>
			  <tbody>			
			<%
		
			Do While Not rsLookupProductSKU.EOF
					
				%>
			    <tr>
			      <td data-label="<%= shortCompanyName %> Code"><%= rsLookupProductSKU("SKU") %></td>
			      <td data-label="<%= shortCompanyName %> UoM"><%= rsLookupProductSKU("UM") %></td>
			    </tr>			
				<%
			rsLookupProductSKU.MoveNext
			Loop
			%>
			  </tbody>
			</table>			
			<%
		Else
			Response.Write("No Product Entry SKU was Found For " & partnerSKUEnteredByUser)
		End If
		
		
		set rsLookupProductSKU = Nothing
		cnnLookupProductSKU.close
		set cnnLookupProductSKU = Nothing
		
	Else
		Response.Write("Cannot Lookup, Invalid Data")
		
	End If

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub SaveEquivalentSKUandUM()

	equivSKUEnteredByUser = Request.Form("sku") 
	equivSKUEnteredByUser = Replace(equivSKUEnteredByUser, "'", "''")
	
	skuIdentifyingInfoForSQL = Request.Form("id")

	
	skuIdentifyingInfoForSQLArray = Split(skuIdentifyingInfoForSQL,"*")
	
	EquivSKUNum = Right(skuIdentifyingInfoForSQLArray(0), 1)
	productsTableUM = skuIdentifyingInfoForSQLArray(1)
	productsTableSKU = skuIdentifyingInfoForSQLArray(2)
	partnerID = skuIdentifyingInfoForSQLArray(3)
	productsTableCatID = skuIdentifyingInfoForSQLArray(4)
	
	Set rsSaveEquivSKU = Server.CreateObject("ADODB.Recordset")
	rsSaveEquivSKU.CursorLocation = 3 
	
	If productsTableSKU <> "" AND equivSKUEnteredByUser <> "" AND EquivSKUNum <> "" AND productsTableUM <> "" AND partnerID <> "" Then
	
		SQLSaveEquivSKU = "SELECT * FROM IC_ProductMapping WHERE "
		SQLSaveEquivSKU = SQLSaveEquivSKU & " partnerIntRecID =" & partnerID & " AND "
		SQLSaveEquivSKU = SQLSaveEquivSKU & "SKU = '" & productsTableSKU & "' AND "
		SQLSaveEquivSKU = SQLSaveEquivSKU & "UM = '" & productsTableUM & "' AND "
		SQLSaveEquivSKU = SQLSaveEquivSKU & "CategoryID = " & productsTableCatID
		
		Set cnnSaveEquivSKU = Server.CreateObject("ADODB.Connection")
		cnnSaveEquivSKU.open (Session("ClientCnnString"))
		Set rsSaveEquivSKU = cnnSaveEquivSKU.Execute(SQLSaveEquivSKU)
		
		If NOT rsSaveEquivSKU.EOF Then
		
			If cInt(EquivSKUNum) = 1 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU1 = '" & equivSKUEnteredByUser & "' WHERE "
			ElseIf cInt(EquivSKUNum) = 2 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU2 = '" & equivSKUEnteredByUser & "' WHERE "
			ElseIf cInt(EquivSKUNum) = 3 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU3 = '" & equivSKUEnteredByUser & "' WHERE "
			ElseIf cInt(EquivSKUNum) = 4 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU4 = '" & equivSKUEnteredByUser & "' WHERE "
			ElseIf cInt(EquivSKUNum) = 5 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU5 = '" & equivSKUEnteredByUser & "' WHERE "
			ElseIf cInt(EquivSKUNum) = 6 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU6 = '" & equivSKUEnteredByUser & "' WHERE "
			End If
				
			SQLUpdate = SQLUpdate & " partnerIntRecID =" & partnerID & " AND "
			SQLUpdate = SQLUpdate & "SKU = '" & productsTableSKU & "' AND "
			SQLUpdate = SQLUpdate & "CategoryID = '" & productsTableCatID & "' AND "
			SQLUpdate = SQLUpdate & "UM = '" & productsTableUM & "'"
			
			Response.Write(SQLUpdate)
			
			Set cnnUpdate = Server.CreateObject("ADODB.Connection")
			cnnUpdate.open (Session("ClientCnnString"))
			Set rsUpdate = Server.CreateObject("ADODB.Recordset")
			rsUpdate.CursorLocation = 3 
			Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
			cnnUpdate.close
			
			'Response.Write("Success")
				
		Else

			If cInt(EquivSKUNum) = 1 Then
				SQLInsert = "INSERT INTO IC_ProductMapping (partnerIntRecID, CategoryID, SKU, UM, partnerEquivalentSKU1) VALUES "
				SQLInsert = SQLInsert & " (" & partnerID & "," & productsTableCatID & ",'" & productsTableSKU & "','" & productsTableUM & "','" & equivSKUEnteredByUser & "')"
			ElseIf cInt(EquivSKUNum) = 2 Then
				SQLInsert = "INSERT INTO IC_ProductMapping (partnerIntRecID, CategoryID, SKU, UM, partnerEquivalentSKU2) VALUES "
				SQLInsert = SQLInsert & " (" & partnerID & "," & productsTableCatID & ",'" & productsTableSKU & "','" & productsTableUM & "','" & equivSKUEnteredByUser & "')"
			ElseIf cInt(EquivSKUNum) = 3 Then
				SQLInsert = "INSERT INTO IC_ProductMapping (partnerIntRecID, CategoryID, SKU, UM, partnerEquivalentSKU3) VALUES "
				SQLInsert = SQLInsert & " (" & partnerID & "," & productsTableCatID & ",'" & productsTableSKU & "','" & productsTableUM & "','" & equivSKUEnteredByUser & "')"
			ElseIf cInt(EquivSKUNum) = 4 Then
				SQLInsert = "INSERT INTO IC_ProductMapping (partnerIntRecID, CategoryID, SKU, UM, partnerEquivalentSKU4) VALUES "
				SQLInsert = SQLInsert & " (" & partnerID & "," & productsTableCatID & ",'" & productsTableSKU & "','" & productsTableUM & "','" & equivSKUEnteredByUser & "')"
			ElseIf cInt(EquivSKUNum) = 5 Then
				SQLInsert = "INSERT INTO IC_ProductMapping (partnerIntRecID, CategoryID, SKU, UM, partnerEquivalentSKU5) VALUES "
				SQLInsert = SQLInsert & " (" & partnerID & "," & productsTableCatID & ",'" & productsTableSKU & "','" & productsTableUM & "','" & equivSKUEnteredByUser & "')"
			ElseIf cInt(EquivSKUNum) = 6 Then
				SQLInsert = "INSERT INTO IC_ProductMapping (partnerIntRecID, CategoryID, SKU, UM, partnerEquivalentSKU6) VALUES "
				SQLInsert = SQLInsert & " (" & partnerID & "," & productsTableCatID & ",'" & productsTableSKU & "','" & productsTableUM & "','" & equivSKUEnteredByUser & "')"
			End If
		
			Set cnnInsert = Server.CreateObject("ADODB.Connection")
			cnnInsert.open (Session("ClientCnnString"))
			Set rsInsert = Server.CreateObject("ADODB.Recordset")
			rsInsert.CursorLocation = 3 
			Set rsInsert = cnnInsert.Execute(SQLInsert)
			cnnInsert.close
			Response.Write("Success")
			
		End If
		
		set rsSaveEquivSKU = Nothing
		cnnSaveEquivSKU.close
		set cnnSaveEquivSKU = Nothing
		
	ElseIf productsTableSKU <> "" AND equivSKUEnteredByUser = "" AND EquivSKUNum <> "" AND productsTableUM <> "" AND partnerID <> "" Then
	
		SQLSaveEquivSKU = "SELECT * FROM IC_ProductMapping WHERE "
		SQLSaveEquivSKU = SQLSaveEquivSKU & " partnerIntRecID =" & partnerID & " AND "
		SQLSaveEquivSKU = SQLSaveEquivSKU & "SKU = '" & productsTableSKU & "' AND "
		SQLSaveEquivSKU = SQLSaveEquivSKU & "UM = '" & productsTableUM & "' AND "
		SQLSaveEquivSKU = SQLSaveEquivSKU & "CategoryID = " & productsTableCatID
		
		Set cnnSaveEquivSKU = Server.CreateObject("ADODB.Connection")
		cnnSaveEquivSKU.open (Session("ClientCnnString"))
		Set rsSaveEquivSKU = cnnSaveEquivSKU.Execute(SQLSaveEquivSKU)
		
		If NOT rsSaveEquivSKU.EOF Then
		
			If cInt(EquivSKUNum) = 1 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU1 = '" & equivSKUEnteredByUser & "' WHERE "
			ElseIf cInt(EquivSKUNum) = 2 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU2 = '" & equivSKUEnteredByUser & "' WHERE "
			ElseIf cInt(EquivSKUNum) = 3 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU3 = '" & equivSKUEnteredByUser & "' WHERE "
			ElseIf cInt(EquivSKUNum) = 4 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU4 = '" & equivSKUEnteredByUser & "' WHERE "
			ElseIf cInt(EquivSKUNum) = 5 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU5 = '" & equivSKUEnteredByUser & "' WHERE "
			ElseIf cInt(EquivSKUNum) = 6 Then
				SQLUpdate = "UPDATE IC_ProductMapping SET partnerEquivalentSKU6 = '" & equivSKUEnteredByUser & "' WHERE "
			End If
				
			SQLUpdate = SQLUpdate & " partnerIntRecID =" & partnerID & " AND "
			SQLUpdate = SQLUpdate & "SKU = '" & productsTableSKU & "' AND "
			SQLUpdate = SQLUpdate & "CategoryID = '" & productsTableCatID & "' AND "
			
			SQLUpdate = SQLUpdate & "UM = '" & productsTableUM & "'"
			
			Set cnnUpdate = Server.CreateObject("ADODB.Connection")
			cnnUpdate.open (Session("ClientCnnString"))
			Set rsUpdate = Server.CreateObject("ADODB.Recordset")
			rsUpdate.CursorLocation = 3 
			Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
			cnnUpdate.close
			Response.Write("Success")
		
		End If
	
		
	Else
		Response.Write("Cannot Save, Invalid Data")
		
	End If

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Function GenerateInventoryReportCSV()

	baseURL = Request.Form("baseURL")

	'************************
	'Read Settings_Reports
	'************************
	SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1600 AND UserNo = " & Session("userNo")
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs= cnn8.Execute(SQL)
	If NOT rs.EOF Then
		UnitUPCData = rs("ReportSpecificData1")
		CaseUPCData = rs("ReportSpecificData2")
		InventoriedItem = rs("ReportSpecificData3")
		PickableItem = rs("ReportSpecificData4")
		ProductCategoriesForInventoryReport = rs("ReportSpecificData5")
		If IsNull(UnitUPCData) Then UnitUPCData = ""
		If IsNull(CaseUPCData) Then CaseUPCData  = ""
		If IsNull(InventoriedItem) Then InventoriedItem = ""
		If IsNull(PickableItem) Then PickableItem = ""
		If IsNull(ProductCategoriesForInventoryReport) Then ProductCategoriesForInventoryReport = ""
	Else
		UnitUPCData = ""
		CaseUPCData = ""
		InventoriedItem = ""
		PickableItem = ""
		ProductCategoriesForInventoryReport = ""
	End If										
	'****************************
	'End Read Settings_Reports
	'****************************
	


	'**************************************************************************************
	'Build WHERE Clause For Unit UPC, Case UPC, Inventoried and Pickable Items
	'**************************************************************************************
	
	WHERE_CLAUSE_UNITUPC = ""
	WHERE_CLAUSE_CASEUPC = ""
	WHERE_CLAUSE_INVENTORIEDITEM = ""
	WHERE_CLAUSE_PICKABLEITEM = ""
	WHERE_CLAUSE_CATEGORY = ""
	
	If UnitUPCData = "NOTEMPTY" Then
		WHERE_CLAUSE_UNITUPC = " OR (prodUnitUPC <> '') "
	ElseIf  UnitUPCData = "EMPTY" Then
		WHERE_CLAUSE_UNITUPC = " OR (prodUnitUPC = '' OR prodUnitUPC IS NULL) "
	Else
		WHERE_CLAUSE_UNITUPC = ""
	End If
	
	
	If CaseUPCData = "NOTEMPTY" Then
		WHERE_CLAUSE_CASEUPC = " AND (prodCaseUPC <> '') "
	ElseIf  CaseUPCData = "EMPTY" Then
		WHERE_CLAUSE_CASEUPC = " AND (prodCaseUPC = '' OR prodCaseUPC IS NULL) "
	Else
		WHERE_CLAUSE_CASEUPC = ""
	End If
				
			
	If InventoriedItem = "YES" Then
		WHERE_CLAUSE_INVENTORIEDITEM = " AND (prodInventoriedItem = 1) "
	ElseIf InventoriedItem = "NO" Then
		WHERE_CLAUSE_INVENTORIEDITEM = " AND (prodInventoriedItem = 0) "
	Else
		WHERE_CLAUSE_INVENTORIEDITEM = " "
	End If
				
	If PickableItem = "YES" Then
		WHERE_CLAUSE_PICKABLEITEM = " AND (prodPickableItem = 1) "
	ElseIf InventoriedItem = "NO" Then
		WHERE_CLAUSE_PICKABLEITEM = " AND (prodPickableItem = 0) "
	Else
		WHERE_CLAUSE_PICKABLEITEM = ""
	End If
	
	CategoryArray = ""
	CategoryArray = Split(ProductCategoriesForInventoryReport,",")
	
	'**************************************************************************************
	'Build WHERE Clause For Product Category Array
	'**************************************************************************************
	
	For z = 0 to UBound(CategoryArray)
		If z = 0 AND UBound(CategoryArray) = 0 Then
			WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & " AND (prodCategory = '" & CategoryArray(z) & "' "
		ElseIf z = 0 AND UBound(CategoryArray) > 0 Then
			WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & " AND ((prodCategory = '" & CategoryArray(z) & "') "
		Else
			WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & " OR (prodCategory = '" & CategoryArray(z) & "')"
		End If
	Next	
		
	If WHERE_CLAUSE_CATEGORY <> "" Then
		WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & ")"
	End If



	SQL = " SELECT prodUnitUPC, prodCaseUPC, prodSKU, prodDescription, prodCategory, "
	SQL = SQL & "prodCasePricing, prodCaseDescription, prodInventoriedItem, prodPickableItem "
	SQL = SQL & "FROM IC_Product "
	SQL = SQL & " WHERE prodSKU <> '' " & WHERE_CLAUSE_UNITUPC
	SQL = SQL & " " & WHERE_CLAUSE_CASEUPC
	SQL = SQL & " " & WHERE_CLAUSE_INVENTORIEDITEM
	SQL = SQL & " " & WHERE_CLAUSE_PICKABLEITEM
	SQL = SQL & " " & WHERE_CLAUSE_CATEGORY
	SQL = SQL & " ORDER BY prodCategory, prodSKU ASC "


	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open SQL, Session("ClientCnnString")
	
	dim fs,tfile
	
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	
	set tfile=fs.CreateTextFile(Server.MapPath("..\clientfiles\" & trim(MUV_Read("SERNO")) & "\csv\product_inventory_report_" & trim(MUV_Read("SERNO")) & ".csv"))
	
	tfile.WriteLine("SKU, Desc, Category, Inventoried, Pickable, Case Pricing, Case Desc, Unit UPC, Case UPC")
	
	Do While Not rs.EOF

		prodSKU = rs("prodSKU")
		prodDescription = rs("prodDescription")
		prodCategoryID = rs("prodCategory")
		
		If prodCategoryID <> "" Then
			prodCategoryName = GetCategoryByID(prodCategoryID)
		End If
		
		prodInventoriedItem = rs("prodInventoriedItem")

		If prodInventoriedItem = 1 OR prodInventoriedItem = vbtrue Then
			prodInventoriedItemDisplay = "YES"
		Else
			prodInventoriedItemDisplay = "NO"																		
		End If
		
		prodPickableItem = rs("prodPickableItem")

		If prodPickableItem = 1 OR prodPickableItem = vbtrue Then
			prodPickableItemDisplay = "YES"
		Else
			prodPickableItemDisplay = "NO"																		
		End If
		
		prodCasePricing = rs("prodCasePricing")
		prodCaseDescription = rs("prodCaseDescription")
		prodUnitUPC = rs("prodUnitUPC")
		prodCaseUPC = rs("prodCaseUPC")
				
		tfile.WriteLine(prodSKU & "," & prodDescription & "," & prodCategoryName & "," & prodInventoriedItemDisplay & "," & prodPickableItemDisplay & "," & prodCasePricing & "," & prodCaseDescription & "," & prodUnitUPC & "," & prodCaseUPC)
		
		rs.MoveNext

	Loop


	tfile.close
	set tfile=nothing
	set fs=nothing

	rs.Close

	GenerateInventoryReportCSV = BaseURL & "clientfiles/" & trim(MUV_Read("ClientID")) & "/z_pdfs/product_inventory_report_" & trim(MUV_Read("ClientID")) & ".csv"
	
End Function


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'END ALL AJAX MODAL SUBROUTINES AND FUNCTIONS

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

%>