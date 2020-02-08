<!--#include file="SubsAndFuncs.asp"-->
<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<!--#include file="InSightFuncs_Routing.asp"-->
<!--#include file="InSightFuncs_Prospecting.asp"-->
<!--#include file="InsightFuncs_BizIntel.asp"-->
<!--#include file="InsightFuncs_Equipment.asp"-->
<!--#include file="InsightFuncs_AR_AP.asp"-->
<!--#include file="mail.asp"-->
<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
 
'Sub DeleteQuotedItemFromCustomer()
'Sub UndoQuotedItemChangesForCustomer()
'Sub UndoSingleQuotedItemChangeForCustomer()
'Sub GetProductInformationForAddQuotedItemModal()
'Sub GetCategoryInformationForAddQuotedItemModal()
'Sub UpdateExpireDateSingleQuotedItem()
'Sub UpdateExpireDateAllQuotedItems()
'Sub UpdateNewPriceSingleQuotedItem()
'Sub UpdateNewGPPercentSingleQuotedItem()
'Sub AutoQuoteAllAlternateUMSsForCustomer()
'Sub AutoQuoteSingleUMForCustomer()
'Sub WritePeriodsInUseDropdownForReportYearAdd()
'Sub GetReportPeriodDeleteInformationForModal()
'Sub ValidateAndAddReportPeriod()
'Sub UpdateReportPeriod()
'Sub GetTitleForCategoryVPCModal()
'Sub GetContentForCategoryVPCModal()
'Sub GetTitleForEquipmentVPCModal()
'Sub GetContentForEquipmentVPCModal()
'Sub SaveGeneralNotesGroupM()
'Sub DeleteMCSClientbyCustID()
'Sub DeleteMESClientbyCustID()
'Sub AddMCSClientbyCustID()
'Sub AddMESClientbyCustID()
'Sub GetTRofNewMCSClientbyCustID()
'Sub getCCUsers()
'Sub getSelectUsers()
'Sub getUsersBySalesperson()
'Sub WritePeriodsInUseDropdownForAccountingYearAdd()
'Sub GetAccountingPeriodDeleteInformationForModal()
'Sub ValidateAndAddAccountingPeriod()
'Sub UpdateAccountingPeriod()
'Sub GenerateMCSPendingChargesPDF()

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
	Case "DeleteQuotedItemFromCustomer" 
		DeleteQuotedItemFromCustomer()	
	Case "UndoQuotedItemChangesForCustomer"
		UndoQuotedItemChangesForCustomer()	
	Case "UndoSingleQuotedItemChangeForCustomer"
		UndoSingleQuotedItemChangeForCustomer()
	Case "GetProductInformationForAddQuotedItemModal" 
		GetProductInformationForAddQuotedItemModal()
	Case "GetCategoryInformationForAddQuotedItemModal" 
		GetCategoryInformationForAddQuotedItemModal()
	Case "UpdateExpireDateSingleQuotedItem"
		UpdateExpireDateSingleQuotedItem()
	Case "UpdateExpireDateAllQuotedItems"
		UpdateExpireDateAllQuotedItems()	
	Case "UpdateNewPriceSingleQuotedItem"
		UpdateNewPriceSingleQuotedItem()	
	Case "UpdateNewGPPercentSingleQuotedItem"
		UpdateNewGPPercentSingleQuotedItem()
	Case "AutoQuoteAllAlternateUMSsForCustomer"
		AutoQuoteAllAlternateUMSsForCustomer()
	Case "AutoQuoteSingleUMForCustomer"
		AutoQuoteSingleUMForCustomer()
	Case "WritePeriodsInUseDropdownForReportYearAdd"
		WritePeriodsInUseDropdownForReportYearAdd()
	Case "GetReportPeriodDeleteInformationForModal"
		GetReportPeriodDeleteInformationForModal()
	Case "ValidateAndAddReportPeriod"
		ValidateAndAddReportPeriod()
	Case "UpdateReportPeriod"
		UpdateReportPeriod()
	Case "GetContentForCategoryAnalysisByPeriodNotesModal"
		GetContentForCategoryAnalysisByPeriodNotesModal()
	Case "GetContentForCustomerNotesModal"
		GetContentForCustomerNotesModal()
	Case "GetTitleForCategoryVPCModal"
		GetTitleForCategoryVPCModal()
	Case "GetContentForCategoryVPCModal"
		GetContentForCategoryVPCModal()
	Case "GetTitleForEquipmentVPCModal"
		GetTitleForEquipmentVPCModal()
	Case "GetContentForEquipmentVPCModal"
		GetContentForEquipmentVPCModal()
	Case "SaveGeneralNotesGroupM"
		SaveGeneralNotesGroupM()
	Case "DeleteMCSClientbyCustID"
		DeleteMCSClientbyCustID()
	Case "DeleteMESClientbyCustID"
		DeleteMESClientbyCustID()
	Case "AddMCSClientbyCustID"
		AddMCSClientbyCustID()
	Case "AddMESClientbyCustID"
		AddMESClientbyCustID()
	Case "GetTRofNewMCSClientbyCustID"
		GetTRofNewMCSClientbyCustID()
	Case "getSelectUsers"
		getSelectUsers()
	Case "getCCUsers"
		getCCUsers()
	Case "getUsersBySalesperson"
		getUsersBySalesperson()
	Case "WritePeriodsInUseDropdownForAccountingYearAdd"
		WritePeriodsInUseDropdownForAccountingYearAdd()
	Case "WriteStartDateForAccountingYearAdd"
		WriteStartDateForAccountingYearAdd()
	Case "EditStartDateForAccountingYearAdd"
		EditStartDateForAccountingYearAdd()
	Case "GetAccountingPeriodDeleteInformationForModal"
		GetAccountingPeriodDeleteInformationForModal()
	Case "ValidateAndAddAccountingPeriod"
		ValidateAndAddAccountingPeriod()
	Case "UpdateAccountingPeriod"
		UpdateAccountingPeriod()
	Case "GenerateMCSPendingChargesPDF"
		GenerateMCSPendingChargesPDF()
End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub DeleteQuotedItemFromCustomer()

	IntRecID = Request.Form("recid") 
	
	Set rsCheckForChain = Server.CreateObject("ADODB.Recordset")
	rsCheckForChain.CursorLocation = 3 

	SQLCheckForChain = "SELECT * FROM zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " WHERE InternalRecordIdentifier=" & IntRecID
	
	Set cnnCheckForChain = Server.CreateObject("ADODB.Connection")
	cnnCheckForChain.open (Session("ClientCnnString"))
	Set rsCheckForChain = cnnCheckForChain.Execute(SQLCheckForChain)
	
	If NOT rsCheckForChain.EOF Then
	
		QuotedSKU = rsCheckForChain("ProdSKU")
		QuotedUM = rsCheckForChain("QuoteType")
		QuotedToChainOrAccount = rsCheckForChain("QuotedToChainOrAccount")
	
		If QuotedToChainOrAccount <> "C" Then
	
			SQL = "UPDATE zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " SET DeleteFlag = 1 WHERE InternalRecordIdentifier=" & IntRecID
			
			Set cnn8 = Server.CreateObject("ADODB.Connection")
			cnn8.open (Session("ClientCnnString"))
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3 
			Set rs = cnn8.Execute(SQL)
			cnn8.close
			Response.Write("ACCOUNT," & QuotedSKU & "," & QuotedUM & "," & IntRecID)
			
		Else
			Response.Write("CHAIN," & QuotedSKU & "," & QuotedUM & "," & IntRecID)
		End If
		
	End If
	
	set rsCheckForChain = Nothing
	cnnCheckForChain.close
	set cnnCheckForChain = Nothing

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub UndoQuotedItemChangesForCustomer()

	SQL = "UPDATE zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " SET DeleteFlag = 0, NewPrice = Null, NewExpireDate = Null WHERE QuotedToChainOrAccount <> 'C'"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	SQL = "DELETE FROM zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " WHERE AutoGenerated = 1"
	
	Set rs = cnn8.Execute(SQL)
	cnn8.close

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub UndoSingleQuotedItemChangeForCustomer()

	IntRecID = Request.Form("recid")
	
	SQL = "UPDATE zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " SET DeleteFlag = 0, NewPrice = Null, NewExpireDate = Null WHERE InternalRecordIdentifier=" & IntRecID 
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	cnn8.close

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetCategoryInformationForAddQuotedItemModal() 

	custID = Request.Form("custID")
	%>
	<!--#include file="../bizintel/tools/customer/quotes/buildProdList.asp"-->
	<p><strong>Company:</strong> <%= custID %>, <%= GetCustNameByCustNum(custID) %></p>
	
	<div class="col-lg-12" style="padding-left:0px;">
		<label class="control-label" style="padding-left:0px;">Choose a category that contains the product to add:</label>
	</div>
	
	<div class="col-lg-12" style="padding-left:0px;">	
	  	<select class="form-control" name="selAddQuotedItemCategories" id="selAddQuotedItemCategories" onchange="quotedItemCategorySelected();">
			<% 
				'Get all unquoted items for this account
			  	SQLQuotedItemsCategory = "SELECT Distinct(Category) FROM zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno")) & " ORDER BY Category ASC"
			
				Set cnnQuotedItemsCategory = Server.CreateObject("ADODB.Connection")
				cnnQuotedItemsCategory.open (Session("ClientCnnString"))
				Set rsQuotedItemsCategory = Server.CreateObject("ADODB.Recordset")
				rsQuotedItemsCategory.CursorLocation = 3 
				Set rsQuotedItemsCategory = cnnQuotedItemsCategory.Execute(SQLQuotedItemsCategory)
					
				If not rsQuotedItemsCategory.EOF Then
										
					Do
						%><option value="<%= rsQuotedItemsCategory("Category") %>"><%= GetTerm(GetCategoryByID(rsQuotedItemsCategory("Category"))) %></option><%
						rsQuotedItemsCategory.movenext
					Loop until rsQuotedItemsCategory.EOF
				End If
				set rsQuotedItemsCategory = Nothing
				cnnQuotedItemsCategory.close
				set cnnQuotedItemsCategory = Nothing
				
			%>									
		</select>
	</div>
<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetProductInformationForAddQuotedItemModal() 

	custID = Request.Form("custID")
	category = Request.Form("categoryID")
	
	%>
	
	<div class="col-lg-12" style="padding-left:0px; margin-top:15px;">
		<label class="control-label" style="padding-left:0px;">Choose an item to add to <%=GetTerm("Customer")%> quoted items list:</label>
	</div>
	
	<div class="col-lg-12" style="padding-left:0px;">	
	  	<select class="form-control" name="selAddQuotedItemQuotedItemsSKUs" id="selAddQuotedItemQuotedItemsSKUs" onchange="quotedItemSelected();">
			<% 
				'Get all unquoted items for this account
			  	SQLQuotedItems = "SELECT * FROM zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno")) & " WHERE Category = " & category & " ORDER BY ProdSKU, UM ASC"
				Set cnnQuotedItems = Server.CreateObject("ADODB.Connection")
				cnnQuotedItems.open (Session("ClientCnnString"))
				Set rsQuotedItems = Server.CreateObject("ADODB.Recordset")
				rsQuotedItems.CursorLocation = 3 
				Set rsQuotedItems = cnnQuotedItems.Execute(SQLQuotedItems)
					
				If not rsQuotedItems.EOF Then
					
					Do
						%><option value="<%= rsQuotedItems("prodSKU") %>*<%= rsQuotedItems("UM") %>"><%= rsQuotedItems("prodSKU") %>---<%= rsQuotedItems("UM") %>---<%= rsQuotedItems("Description") %></option><%
						rsQuotedItems.movenext	
					Loop until rsQuotedItems.eof
				End If
				set rsQuotedItems = Nothing
				cnnQuotedItems.close
				set cnnQuotedItems = Nothing
				
			%>									
		</select>
	</div>
<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub UpdateExpireDateSingleQuotedItem()

	IntRecID = Request.Form("recid")
	ExpDate = Request.Form("expdate")

	SQLUpdateExpireDateSingle = "UPDATE zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " SET NewExpireDate = '" & expdate & "' WHERE QuotedToChainOrAccount <> 'C' AND InternalRecordIdentifier=" & IntRecID
	
	Set cnnUpdateExpireDateSingle = Server.CreateObject("ADODB.Connection")
	cnnUpdateExpireDateSingle.open (Session("ClientCnnString"))
	Set rsUpdateExpireDateSingle = Server.CreateObject("ADODB.Recordset")
	rsUpdateExpireDateSingle.CursorLocation = 3 
	Set rsUpdateExpireDateSingle = cnnUpdateExpireDateSingle.Execute(SQLUpdateExpireDateSingle)
	cnnUpdateExpireDateSingle.close

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub UpdateExpireDateAllQuotedItems()

	ExpDate = Request.Form("expdate")

	SQLUpdateExpireDateAll = "UPDATE zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " SET NewExpireDate = '" & expdate & "' WHERE QuotedToChainOrAccount <> 'C'"
	
	Set cnnUpdateExpireDateAll = Server.CreateObject("ADODB.Connection")
	cnnUpdateExpireDateAll.open (Session("ClientCnnString"))
	Set rsUpdateExpireDateAll = Server.CreateObject("ADODB.Recordset")
	rsUpdateExpireDateAll.CursorLocation = 3 
	Set rsUpdateExpireDateAll = cnnUpdateExpireDateAll.Execute(SQLUpdateExpireDateAll)
	cnnUpdateExpireDateAll.close

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub UpdateNewPriceSingleQuotedItem()

	IntRecID = Request.Form("recid")
	NewPrice = Request.Form("newprice")

	Set cnnUpdateNewPriceSingle = Server.CreateObject("ADODB.Connection")
	cnnUpdateNewPriceSingle.open (Session("ClientCnnString"))
	Set rsUpdateNewPriceSingle = Server.CreateObject("ADODB.Recordset")
	rsUpdateNewPriceSingle.CursorLocation = 3 

	SQLUpdateNewPriceSingle ="Select * From zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " WHERE InternalRecordIdentifier = " & IntRecID

	Set rsUpdateNewPriceSingle = cnnUpdateNewPriceSingle.Execute(SQLUpdateNewPriceSingle)

	If Not rsUpdateNewPriceSingle.EOF Then 
		CurrentCost = rsUpdateNewPriceSingle("Cost")
		QuotedPrice = rsUpdateNewPriceSingle("Price") 
	Else 
		CurrentCost = 0
		QuotedPrice = 0
	End If
	
	SQLUpdateNewPriceSingle = "UPDATE zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " SET NewPrice = " & NewPrice & ",NewGPPercent= NULL WHERE InternalRecordIdentifier=" & IntRecID
	
	Set rsUpdateNewPriceSingle = cnnUpdateNewPriceSingle.Execute(SQLUpdateNewPriceSingle)
	cnnUpdateNewPriceSingle.close
	
	If cDbl(NewPrice) > cDbl(QuotedPrice) Then
		ChangeMessage = "INCREASE"
	ElseIf cDbl(NewPrice) < cDbl(QuotedPrice) Then
		ChangeMessage = "DECREASE"
	ElseIf cDbl(NewPrice) = cDbl(QuotedPrice) Then
		ChangeMessage = "NOCHANGE"
	Else
		ChangeMessage = ""
	End If
		
	Response.write(CurrentCost & "*" & QuotedPrice & "*" & ChangeMessage)

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub UpdateNewGPPercentSingleQuotedItem()

	IntRecID = Request.Form("recid")
	NewGPPercent = Request.Form("gpp")

	Set cnnUpdateNewPriceSingle = Server.CreateObject("ADODB.Connection")
	cnnUpdateNewPriceSingle.open (Session("ClientCnnString"))
	Set rsUpdateNewPriceSingle = Server.CreateObject("ADODB.Recordset")
	rsUpdateNewPriceSingle.CursorLocation = 3 

	SQLUpdateNewPriceSingle ="Select * From zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " WHERE InternalRecordIdentifier = " & IntRecID

	Set rsUpdateNewPriceSingle = cnnUpdateNewPriceSingle.Execute(SQLUpdateNewPriceSingle)

	If Not rsUpdateNewPriceSingle.EOF Then 
		CurrentCost = rsUpdateNewPriceSingle("Cost")
		QuotedPrice = rsUpdateNewPriceSingle("Price") 
	Else 
		CurrentCost = 0
		QuotedPrice = 0
	End If
				
	SQLUpdateNewPriceSingle = "UPDATE zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " SET NewGPPercent = " & NewGPPercent & ", NewPrice = NULL WHERE InternalRecordIdentifier=" & IntRecID
	
	Set rsUpdateNewPriceSingle = cnnUpdateNewPriceSingle.Execute(SQLUpdateNewPriceSingle)
	cnnUpdateNewPriceSingle.close

	
	If cDbl(QuotedPrice) > 0 Then
		OldGPPercent = Round((((QuotedPrice - CurrentCost)/QuotedPrice) * 100),2)
	Else
		OldGPPercent = 0
	End If
	
	If cDbl(NewGPPercent) > cDbl(OldGPPercent) Then
		ChangeMessage = "INCREASE"
	ElseIf cDbl(NewGPPercent) < cDbl(OldGPPercent) Then
		ChangeMessage = "DECREASE"
	ElseIf cDbl(NewGPPercent) = cDbl(OldGPPercent) Then
		ChangeMessage = "NOCHANGE"
	Else
		ChangeMessage = ""
	End If
		
	Response.write(CurrentCost & "*" & QuotedPrice & "*" & ChangeMessage)


End Sub



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub AutoQuoteAllAlternateUMSsForCustomer()

	SQL = "SELECT * FROM  zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " WHERE DeleteFlag <> 1 AND QuotedToChainOrAccount <> 'C' AND QuoteType <> 'N' ORDER BY Category, ProdSKU" 
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rsAltUMButtonCheck = Server.CreateObject("ADODB.Recordset")
	rsAltUMButtonCheck.CursorLocation = 3 
	Set rsProduct = Server.CreateObject("ADODB.Recordset")
	rsProduct.CursorLocation = 3 
	Set rsAltUMInsert = Server.CreateObject("ADODB.Recordset")
	rsAltUMInsert.CursorLocation = 3 


	Set rs = cnn8.Execute(SQL)
	
	If NOT rs.EOF Then
	
		Do While NOT rs.EOF
	
	
			'See if we need to make another UM

			SQLAltUMButtonCount = "SELECT COUNT (prodSKU) AS skuCount FROM zPRC_AccountQuotedItems_" & trim(Session("Userno"))
			SQLAltUMButtonCount = SQLAltUMButtonCount & " WHERE prodSKU='" & rs("ProdSKU") & "'"

			Set rsAltUMButtonCheck = cnn8.Execute(SQLAltUMButtonCount)

			If NOT rsAltUMButtonCheck.EOF Then
					If rsAltUMButtonCheck("skuCount") < 2 Then
					
						MakeAltProd = True
						NewPrice = ""
						NewGPPercent = ""
						OriginalRecordNewGPPercent = ""
						OriginalRecordNewPrice = ""
						
						prodSku = rs("prodSku")
						QuoteType = rs("QuoteType")
						
						If QuoteType= "U" Then
							SQLProduct = "SELECT * FROM Product WHERE PartNo = '" & prodSku & "' COLLATE Latin1_General_CS_AS" ' AND casePricing = 'C'"
							Set rsProduct = cnn8.Execute(SQLProduct)
							
							If rsProduct.EOF Then 
								MakeAltProd = False
							Else
								Description = rsProduct("CaseDescription")
								If Description = "" Then Description = rs("Description")
								Category = rs("Category")
								QuoteType = "C"
								If rs("SuggestedQty") <> 0 AND rsProduct("CaseConversionFactor") <> 0 Then 
									SuggestedQty = Round(rs("SuggestedQty") / rsProduct("CaseConversionFactor") ,0)
								Else 
									SuggestedQty = 0
								End If
								If rsProduct("CaseConversionFactor") <> 0 AND rsProduct("CaseConversionFactor") <> 1 Then
								
									Price = rs("Price") * rsProduct("CaseConversionFactor")
									
									Cost = rsProduct("UnitCost") * rsProduct("CaseConversionFactor")
									
									OriginalRecordNewPrice = rs("NewPrice")
									
									If OriginalRecordNewPrice <> "" AND NOT ISNULL(OriginalRecordNewPrice) Then
										NewPrice = OriginalRecordNewPrice * rsProduct("CaseConversionFactor")
									Else
										NewPrice = NULL
									End If
									
									OriginalRecordNewGPPercent = rs("NewGPPercent")
									
									If OriginalRecordNewGPPercent <> 0 AND OriginalRecordNewGPPercent <> "" AND NOT ISNULL(OriginalRecordNewGPPercent) Then
										NewGPPercent = OriginalRecordNewGPPercent 
									Else 
										NewGPPercent = NULL
									End If
									
									
									NewExpireDate = rs("NewExpireDate")
									ExpireDate = rs("ExpireDate") 
									ListFlag = rs("ListFlag")
									QuotedToChainOrAccount = "A"
								Else
									MakeAltProd = False
								End If
							End IF							
							
						ElseIf  QuoteType= "C" Then
						
							SQLProduct = "SELECT * FROM Product WHERE PartNo = '" & prodSku & "' COLLATE Latin1_General_CS_AS" ' AND casePricing = 'U'"
							Set rsProduct = cnn8.Execute(SQLProduct)
							If rsProduct.EOF Then 
								MakeAltProd = False
							Else
								Description = rsProduct("Description")
								If Description = "" Then Description = rs("Description")
								Category = rs("Category")
								QuoteType = "U"
								If rs("SuggestedQty") <> 0 AND rsProduct("CaseConversionFactor") <> 0 Then 
									SuggestedQty = Round(rs("SuggestedQty") * rsProduct("CaseConversionFactor"),0)
								Else
									SuggestedQty = 0
								End If
								If rsProduct("CaseConversionFactor") <> 0 AND rsProduct("CaseConversionFactor") <> 1 Then
								
									Price = rs("Price") / rsProduct("CaseConversionFactor")
									
									Cost = rsProduct("UnitCost")
									
									OriginalRecordNewPrice = rs("NewPrice")
									
									If OriginalRecordNewPrice <> "" AND NOT ISNULL(OriginalRecordNewPrice) Then
										NewPrice = OriginalRecordNewPrice / rsProduct("CaseConversionFactor")
									Else
										NewPrice = NULL
									End If
									
									
									OriginalRecordNewGPPercent = rs("NewGPPercent")
									
									If OriginalRecordNewGPPercent <> 0 AND OriginalRecordNewGPPercent <> "" AND NOT ISNULL(OriginalRecordNewGPPercent) Then
										NewGPPercent = OriginalRecordNewGPPercent 
									Else 
										NewGPPercent = NULL
									End If
									
									NewExpireDate = rs("NewExpireDate")
									ExpireDate = rs("ExpireDate") 
									ListFlag = rs("ListFlag")
									QuotedToChainOrAccount = "A"
								Else
									MakeAltProd = False
								End If
							End IF							
						End If
					
					
						'OK to make the new entry
						If MakeAltProd = True Then 
						
							Description = Replace(Description,"'","''")

							SQLAltUMInsert = "INSERT INTO zPRC_AccountQuotedItems_" & trim(Session("Userno"))
							SQLAltUMInsert = SQLAltUMInsert & " (prodSKU, Description, Category, QuoteType, SuggestedQty, Price, ListFlag, Cost, DateQuoted, "
							SQLAltUMInsert = SQLAltUMInsert & " QuotedToChainOrAccount, "
							
							If NewPrice <> "" AND NOT ISNULL(NewPrice) Then SQLAltUMInsert = SQLAltUMInsert & "	NewPrice, "
							If NewGPPercent <> "" AND NOT ISNULL(NewGPPercent) Then SQLAltUMInsert = SQLAltUMInsert & "	NewGPPercent, "
							If NewExpireDate <> "" AND NOT ISNULL(NewExpireDate) Then SQLAltUMInsert = SQLAltUMInsert & "	NewExpireDate, "
							If ExpireDate <> "" AND NOT ISNULL(ExpireDate) Then SQLAltUMInsert = SQLAltUMInsert & "	ExpireDate, "
							
							SQLAltUMInsert = SQLAltUMInsert &  "DeleteFlag, AutoGenerated)  VALUES ("
							SQLAltUMInsert = SQLAltUMInsert & "'" & prodSKU & "', " 
							SQLAltUMInsert = SQLAltUMInsert & "'" & Description & "', " 
							SQLAltUMInsert = SQLAltUMInsert & "'" & Category & "', " 
							SQLAltUMInsert = SQLAltUMInsert & "'" & QuoteType & "', " 														
							SQLAltUMInsert = SQLAltUMInsert & "'" & SuggestedQty & "', " 							
							SQLAltUMInsert = SQLAltUMInsert & Price & ", " 						
							SQLAltUMInsert = SQLAltUMInsert & "'" & ListFlag & "', " 			
							SQLAltUMInsert = SQLAltUMInsert & Cost & ", " 
							SQLAltUMInsert = SQLAltUMInsert & "'" & FormatDateTime(Now(),2) & "', "
							SQLAltUMInsert = SQLAltUMInsert & "'" & QuotedToChainOrAccount & "', "							
							
							If NewPrice <> "" AND NOT ISNULL(NewPrice) Then SQLAltUMInsert = SQLAltUMInsert & NewPrice & ","	
							If NewGPPercent <> "" AND NOT ISNULL(NewGPPercent) Then SQLAltUMInsert = SQLAltUMInsert & NewGPPercent & ","							
							If NewExpireDate <> "" AND NOT ISNULL(NewExpireDate) Then SQLAltUMInsert = SQLAltUMInsert & "'" & NewExpireDate & "',"
							If ExpireDate <> "" AND NOT ISNULL(ExpireDate) Then SQLAltUMInsert = SQLAltUMInsert & "'" & ExpireDate & "',"
							
							SQLAltUMInsert = SQLAltUMInsert & " 0, 1 )"
								
							Response.Write("<br><br>" & SQLAltUMInsert & "<br><br>")
							
							Set rsAltUMInsert = cnn8.Execute(SQLAltUMInsert)
							
						End IF
					
					End IF
			End If
			
			SET rsAltUMButtonCheck = Nothing


	
	
			rs.MoveNext
		Loop
		
	
	End If
	
	
	cnn8.close
	Set rs = Nothing
	Set cnn8 = Nothing

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub AutoQuoteSingleUMForCustomer()

	IntRecID = Request.Form("recid")

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 

	SQLAltUM = "SELECT * FROM zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " WHERE InternalRecordIdentifier=" & IntRecID

	Set rs = cnn8.Execute(SQLAltUM)

	If NOT rs.EOF Then
	
	
		MakeAltProd = True
		NewPrice = ""
		NewGPPercent = ""
		OriginalRecordNewGPPercent = ""
		OriginalRecordNewPrice = ""

		prodSku = rs("prodSku")
		QuoteType = rs("QuoteType")
				
		If QuoteType= "U" Then
			SQLProduct = "SELECT * FROM Product WHERE PartNo = '" & prodSku & "' COLLATE Latin1_General_CS_AS" ' AND casePricing = 'C'"
			Set rsProduct = cnn8.Execute(SQLProduct)
			
			If rsProduct.EOF Then 
				MakeAltProd = False
			Else
				Description = rsProduct("CaseDescription")
				If Description = "" Then Description = rs("Description")
				Category = rs("Category")
				QuoteType = "C"
				If rs("SuggestedQty") <> 0 AND rsProduct("CaseConversionFactor") <> 0 Then 
					SuggestedQty = Round(rs("SuggestedQty") / rsProduct("CaseConversionFactor") ,0)
				Else 
					SuggestedQty = 0
				End If
				If rsProduct("CaseConversionFactor") <> 0 Then
				
					Price = rs("Price") * rsProduct("CaseConversionFactor")
					
					Cost = rsProduct("UnitCost") * rsProduct("CaseConversionFactor")
					
					OriginalRecordNewPrice = rs("NewPrice")
					
					If OriginalRecordNewPrice <> "" AND NOT ISNULL(OriginalRecordNewPrice) Then
						NewPrice = OriginalRecordNewPrice * rsProduct("CaseConversionFactor")
					Else
						NewPrice = NULL
					End If
					
					OriginalRecordNewGPPercent = rs("NewGPPercent")
					
					If OriginalRecordNewGPPercent <> 0 AND OriginalRecordNewGPPercent <> "" AND NOT ISNULL(OriginalRecordNewGPPercent) Then
						NewGPPercent = OriginalRecordNewGPPercent 
					Else 
						NewGPPercent = NULL
					End If
							
							
					NewExpireDate = rs("NewExpireDate")
					ExpireDate = rs("ExpireDate") 
					ListFlag = rs("ListFlag")
					QuotedToChainOrAccount = "A"
				Else
					MakeAltProd = False
				End If
			End IF							
			
		ElseIf  QuoteType= "C" Then
		
			SQLProduct = "SELECT * FROM Product WHERE PartNo = '" & prodSku & "' COLLATE Latin1_General_CS_AS" ' AND casePricing = 'U'"
			Set rsProduct = cnn8.Execute(SQLProduct)
			If rsProduct.EOF Then 
				MakeAltProd = False
			Else
				Description = rsProduct("Description")
				If Description = "" Then Description = rs("Description")
				Category = rs("Category")
				QuoteType = "U"
				If rs("SuggestedQty") <> 0 AND rsProduct("CaseConversionFactor") <> 0 Then 
					SuggestedQty = Round(rs("SuggestedQty") * rsProduct("CaseConversionFactor"),0)
				Else
					SuggestedQty = 0
				End If
				If rsProduct("CaseConversionFactor") <> 0 Then
				
					Price = rs("Price") / rsProduct("CaseConversionFactor")
					
					Cost = rsProduct("UnitCost")
					
					OriginalRecordNewPrice = rs("NewPrice")
					
					If OriginalRecordNewPrice <> "" AND NOT ISNULL(OriginalRecordNewPrice) Then
						NewPrice = OriginalRecordNewPrice / rsProduct("CaseConversionFactor")
					Else
						NewPrice = NULL
					End If
							
							
					OriginalRecordNewGPPercent = rs("NewGPPercent")
					
					If OriginalRecordNewGPPercent <> 0 AND OriginalRecordNewGPPercent <> "" AND NOT ISNULL(OriginalRecordNewGPPercent) Then
						NewGPPercent = OriginalRecordNewGPPercent 
					Else 
						NewGPPercent = NULL
					End If
					
					NewExpireDate = rs("NewExpireDate")
					ExpireDate = rs("ExpireDate") 
					ListFlag = rs("ListFlag")
					QuotedToChainOrAccount = "A"
				Else
					MakeAltProd = False
				End If
			End IF							
		End If
			
			
		'OK to make the new entry
		If MakeAltProd = True Then 
		
			Description = Replace(Description,"'","''")

			SQLAltUMInsert = "INSERT INTO zPRC_AccountQuotedItems_" & trim(Session("Userno"))
			SQLAltUMInsert = SQLAltUMInsert & " (prodSKU, Description, Category, QuoteType, SuggestedQty, Price, ListFlag, Cost, DateQuoted, "
			SQLAltUMInsert = SQLAltUMInsert & " QuotedToChainOrAccount, "
			
			If NewPrice <> "" AND NOT ISNULL(NewPrice) Then SQLAltUMInsert = SQLAltUMInsert & "	NewPrice, "
			If NewGPPercent <> "" AND NOT ISNULL(NewGPPercent) Then SQLAltUMInsert = SQLAltUMInsert & "	NewGPPercent, "
			If NewExpireDate <> "" AND NOT ISNULL(NewExpireDate) Then SQLAltUMInsert = SQLAltUMInsert & "	NewExpireDate, "
			If ExpireDate <> "" AND NOT ISNULL(ExpireDate) Then SQLAltUMInsert = SQLAltUMInsert & "	ExpireDate, "
			
			SQLAltUMInsert = SQLAltUMInsert &  "DeleteFlag, AutoGenerated)  VALUES ("
			SQLAltUMInsert = SQLAltUMInsert & "'" & prodSKU & "', " 
			SQLAltUMInsert = SQLAltUMInsert & "'" & Description & "', " 
			SQLAltUMInsert = SQLAltUMInsert & "'" & Category & "', " 
			SQLAltUMInsert = SQLAltUMInsert & "'" & QuoteType & "', " 														
			SQLAltUMInsert = SQLAltUMInsert & "'" & SuggestedQty & "', " 							
			SQLAltUMInsert = SQLAltUMInsert & Price & ", " 						
			SQLAltUMInsert = SQLAltUMInsert & "'" & ListFlag & "', " 			
			SQLAltUMInsert = SQLAltUMInsert & Cost & ", " 
			SQLAltUMInsert = SQLAltUMInsert & "'" & FormatDateTime(Now(),2) & "', "
			SQLAltUMInsert = SQLAltUMInsert & "'" & QuotedToChainOrAccount & "', "							
			
			If NewPrice <> "" AND NOT ISNULL(NewPrice) Then SQLAltUMInsert = SQLAltUMInsert & NewPrice & ","	
			If NewGPPercent <> "" AND NOT ISNULL(NewGPPercent) Then SQLAltUMInsert = SQLAltUMInsert & NewGPPercent & ","							
			If NewExpireDate <> "" AND NOT ISNULL(NewExpireDate) Then SQLAltUMInsert = SQLAltUMInsert & "'" & NewExpireDate & "',"
			If ExpireDate <> "" AND NOT ISNULL(ExpireDate) Then SQLAltUMInsert = SQLAltUMInsert & "'" & ExpireDate & "',"
			
			SQLAltUMInsert = SQLAltUMInsert & " 0, 1 )"
				
			Response.Write("<br><br>" & SQLAltUMInsert & "<br><br>")
			
			Set rsAltUMInsert = cnn8.Execute(SQLAltUMInsert)
			
		End If
	End If
	

	
	cnn8.close
	Set rs = Nothing
	Set cnn8 = Nothing

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub WritePeriodsInUseDropdownForReportYearAdd()
	
	periodYear = cInt(Request.Form("periodYear"))
	periodNum = cInt(Request.Form("periodNum"))
	periodsInUseThisYear = ""
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 

	SQLCheckPeriodInUseForYear = "SELECT * FROM Settings_CompanyPeriods WHERE Year = " & periodYear

	Set rs = cnn8.Execute(SQLCheckPeriodInUseForYear)
		
	If NOT rs.EOF Then
		Do While Not rs.EOF
			periodsInUseThisYear = periodsInUseThisYear & "---" & rs("Period")
			rs.MoveNext
		Loop							
	End If
						
	cnn8.close
	Set rs = Nothing
	Set cnn8 = Nothing
	
	%>
	<label for="selPeriodNumAdd">Period</label>
	<select class="form-control" id="selPeriodNumAdd" name="selPeriodNumAdd">				
		<%
	
		For i = 1 To 100
		
		  currentPeriod = cStr(i)
		  
		  If InStr(periodsInUseThisYear, currentPeriod) Then
		  
		  	If cInt(periodNum) = cInt(i) Then
		  		%><option value="<%= i %>" disabled selected="selected"><%= i %> (currently in use, please delete first)</option><%
		  	Else
		  		%><option value="<%= i %>" disabled><%= i %> (currently in use, please delete first)</option><%
		  	End If
		  	
		  Else
		  
		  	If cInt(periodNum) = cInt(i) Then
		  		%><option value="<%= i %>" selected="selected"><%= i %></option><%
		  	Else
		  		%><option value="<%= i %>"><%= i %></option><%
		  	End If
		  	
		  End If
		Next
		%>				
	</select>	
	<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetReportPeriodDeleteInformationForModal()

	reportPeriodsArray = Split(Request.Form("reportPeriodsArray"),",")
	
	%>

	<input type="hidden" name="periodArray" id="periodArray" value="<%= Request.Form("reportPeriodsArray") %>">


	<%
	For i = 0 to uBound(reportPeriodsArray)

		IntRecID = cInt(reportPeriodsArray(i))
		
		Set rsDelete = Server.CreateObject("ADODB.Recordset")
		rsDelete.CursorLocation = 3 
	
		SQLDelete = "SELECT * FROM Settings_CompanyPeriods WHERE InternalRecordIdentifier = " & IntRecID		
		
		Set cnnDelete = Server.CreateObject("ADODB.Connection")
		cnnDelete.open (Session("ClientCnnString"))
		Set rsDelete = cnnDelete.Execute(SQLDelete)
		
		If NOT rsDelete.EOF Then
			PeriodYear = rsDelete("Year")
			Period = rsDelete("Period")
			PeriodBeginDate = formatDateTime(rsDelete("BeginDate"),2)
			PeriodEndDate = formatDateTime(rsDelete("EndDate"),2)				
			%><strong><%= PeriodYear %></strong>,&nbsp;Period <%= Period %>,&nbsp;<%= PeriodBeginDate %> - <%= PeriodEndDate %><br><%
		End If
		
	Next
		
	Set rsDelete = Nothing
	cnnDelete.Close
	Set cnnDelete = Nothing

%>


	<label class="control-label" style="padding-left:0px; margin-top:20px;">Click the delete button below to PERMANENTLY DELETE report period(s). This cannot be undone.</label>


<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub ValidateAndAddReportPeriod()

	periodYear = Request.Form("periodYear")
	periodNum = Request.Form("periodNum")
	periodStartDate = Request.Form("periodStartDate")
	periodEndDate = Request.Form("periodEndDate")
	
	Set rsValidatePeriodToAdd = Server.CreateObject("ADODB.Recordset")
	rsValidatePeriodToAdd.CursorLocation = 3 

	SQLValidatePeriodToAdd = "SELECT * FROM Settings_CompanyPeriods WHERE Year = " & periodYear & " AND Period = " & periodNum	
	
	Set cnnValidatePeriodToAdd = Server.CreateObject("ADODB.Connection")
	cnnValidatePeriodToAdd.open (Session("ClientCnnString"))
	Set rsValidatePeriodToAdd = cnnValidatePeriodToAdd.Execute(SQLValidatePeriodToAdd)
	
	If NOT rsValidatePeriodToAdd.EOF Then
	
		Response.write("Period " & periodNum & " already exists in " & periodYear & ".")
		
	Else
		SQLAddPeriod = "INSERT INTO Settings_CompanyPeriods (Year, Period, BeginDate, EndDate) "
		SQLAddPeriod = SQLAddPeriod & " VALUES (" & periodYear & "," & periodNum & ",'" & periodStartDate & "','" & periodEndDate & "') "
		
		Set cnnAddPeriod = Server.CreateObject("ADODB.Connection")
		cnnAddPeriod.open (Session("ClientCnnString"))
		Set rsAddPeriod = Server.CreateObject("ADODB.Recordset")
		rsAddPeriod.CursorLocation = 3 
		Set rsAddPeriod = cnnAddPeriod.Execute(SQLAddPeriod)
		
		set rsAddPeriod = Nothing
		cnnAddPeriod.close
		set cnnAddPeriod = Nothing
		
		Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added period " & periodNum & " in " & periodYear & " ranging from " & periodStartDate & " to " & periodEndDate & "."	 			
		CreateAuditLogEntry "Company Reporting Period Added", "Company Report Perioding Added", "Major", 1, Description		

		Response.write("Success")
			
	End If
	
	Set rsValidatePeriodToAdd = Nothing
	cnnValidatePeriodToAdd.Close
	Set cnnValidatePeriodToAdd = Nothing
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub UpdateReportPeriod()

	periodYear = Request.Form("periodYear")
	periodNum = Request.Form("periodNum")
	periodStartDate = Request.Form("periodStartDate")
	periodEndDate = Request.Form("periodEndDate")
	periodIntRecID = Request.Form("periodIntRecID")
	
	Set rsValidatePeriodToUpdate = Server.CreateObject("ADODB.Recordset")
	rsValidatePeriodToUpdate.CursorLocation = 3 

	SQLValidatePeriodToUpdate = "SELECT * FROM Settings_CompanyPeriods WHERE InternalRecordIdentifier = " & periodIntRecID	
	
	Set cnnValidatePeriodToUpdate = Server.CreateObject("ADODB.Connection")
	cnnValidatePeriodToUpdate.open (Session("ClientCnnString"))
	Set rsValidatePeriodToUpdate = cnnValidatePeriodToUpdate.Execute(SQLValidatePeriodToUpdate)
	
	If NOT rsValidatePeriodToUpdate.EOF Then
		orig_periodStartDate = rsValidatePeriodToUpdate("BeginDate")
		orig_periodEndDate = rsValidatePeriodToUpdate("EndDate")			
	End If
	
	SQLValidatePeriodToUpdate = "UPDATE Settings_CompanyPeriods SET BeginDate = '" & periodStartDate & "', EndDate = '" & periodEndDate & "' WHERE InternalRecordIdentifier = " & periodIntRecID
	Set rsValidatePeriodToUpdate = cnnValidatePeriodToUpdate.Execute(SQLValidatePeriodToUpdate)
		
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " edited period " & periodNum & " in " & periodYear & ", changing the date range from (" & orig_periodStartDate & " - " & orig_periodEndDate & ") to (" & periodStartDate & " - " & periodEndDate & ")."	 			
	CreateAuditLogEntry "Company Reporting Period Edited", "Company Report Perioding Edited", "Major", 1, Description		
	
	Set rsValidatePeriodToUpdate = Nothing
	cnnValidatePeriodToUpdate.Close
	Set cnnValidatePeriodToUpdate = Nothing
	
	Response.write("Success")
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetTitleForCategoryVPCModal() 

	CategoryIDPassed = Request.Form("CategoryID")
	PeriodSeqPassed = Request.Form("PeriodSeq")
	VarianceBasisPassed = Request.Form("VarBasis")
	CustIDPassed = Request.Form("CustID")
	CustName = GetCustNameByCustNum(CustIDPassed)
%>
	<h3><%= CustName %>&nbsp;-
		<%
		If CategoryIDPassed = -1 Then 'Note for entire customer
			If VarianceBasisPassed = "3Periods" Then
				Response.Write("3 Period Avg (" & GetPeriodBySeq(PeriodSeqPassed-1) & "-" &  GetPeriodBySeq(PeriodSeqPassed-3) & ") vs Period " & GetPeriodBySeq(PeriodSeqPassed))
			Else
				Response.Write("12 Period Avg (" & GetPeriodBySeq(PeriodSeqPassed-1) & "-" &  GetPeriodBySeq(PeriodSeqPassed-12) & ") vs Period " & GetPeriodBySeq(PeriodSeqPassed))
			End If
		Else
			If VarianceBasisPassed = "3Periods" Then
				Response.Write(GetTerm(GetCategoryByID(CategoryIDPassed)) & " - 3 Period Avg (" & GetPeriodBySeq(PeriodSeqPassed-1) & "-" &  GetPeriodBySeq(PeriodSeqPassed-3) & ") vs Period " & GetPeriodBySeq(PeriodSeqPassed))
			Else
				Response.Write(GetTerm(GetCategoryByID(CategoryIDPassed)) & " - 12 Period Avg (" & GetPeriodBySeq(PeriodSeqPassed-1) & "-" &  GetPeriodBySeq(PeriodSeqPassed-12) & ") vs Period " & GetPeriodBySeq(PeriodSeqPassed))
			End If
		End If
		%>
	</h3>
<%

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetContentForCategoryVPCModal() 

	CategoryIDPassed = Request.Form("CategoryID")
	CustIDPassed = Request.Form("CustID")
	PeriodSeqPassed = Request.Form("PeriodSeq")
	VarianceBasisPassed = Request.Form("VarBasis")

	CustForDetail = CustIDPassed
	SelectedCategoryID = CategoryIDPassed 
	PeriodSeq = PeriodSeqPassed 
	
	VarianceBasis = VarianceBasisPassed 
	If VarianceBasis <> "3Periods" Then VarianceBasis ="12Periods"
	
	FirstPeriodBeingEvaluated = PeriodSeq -1
	SecondPeriodBeingEvaluated = PeriodSeq 
	
	CurrentPeriodNumber = GetPeriodBySeq(PeriodSeq + 1)
	
	' Zero out all the total variables
	VarianceGrandTot_Sales = 0 : VarianceGrandTot_Cases = 0 : VarianceGrandTot_PriceChange = 0 : VarianceGrandTot_VolumeChange = 0
	VarianceBasisGrandTot_Sales = 0 : VarianceBasisGrandTot_Cases = 0 
	PeriodBeingEvalGrandTot_Sales = 0 : PeriodBeingEvalGrandTot_Cases = 0 
	CurrentPeriodGrandTot_Sales = 0 : CurrentPeriodGrandTot_Cases = 0	

	If VarianceBasis = "3Periods" Then
		WorkDaysInPeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeq-3), GetPeriodEndDateBySeq(PeriodSeq-1))
	Else
		WorkDaysInPeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeq-12), GetPeriodEndDateBySeq(PeriodSeq-1))
	End If
	
	WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeq), GetPeriodEndDateBySeq(PeriodSeq))
	WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeq+1), GetPeriodEndDateBySeq(PeriodSeq+1))
	WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeq+1),Date())

%>
	
	<style>
		table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
		    content: " \25B4\25BE" 
		}
		table.sortable thead {
		    color:#222;
		    font-weight: bold;
		    cursor: pointer;
		}
		.top-grey-table{
			background-color:#f8f9fa;
			border-top: 1px solid #ccc;
		}
		.top-grey-table tr,td{
			background-color: transparent;
		 }
		 .top-grey-table>tbody>tr>td{
			 border: 0px;
		 }
		.small-date{
			margin-left: 20px;
		} 
		.price-volume-net-proof{
			width: 33%;
		}
		.period-difference{
			width: 26%
		}
		.td-align{
			text-align: right;
		}
		.td-align1{
			text-align: center;
		}
		.table-size{
			width: 60%;
		}	
		
		.positive{
			font-weight:bold;
			color:blue;
		}
		
		.negative{
			font-weight:bold;
			color:red;	
		}
		
		.table > tfoot > tr > td, .table > tfoot > tr > th {
		    border: none !important;
		}	
		
		
		.table > tbody > tr > td {

		    line-height: 1.42857143;
		    vertical-align: top;
		    border: 1px solid #ddd;	
		 }
		 
		 
		 
		#tableSuperSum2.table > tbody > tr > td:nth-child(3) {
			border-right: 2px solid #555;
		}
		#tableSuperSum2.table > tbody > tr > th:nth-child(3) {
			border-right: 2px solid #555;
		}	


		#tableSuperSum2.table > tbody > tr > td:nth-child(7) {
			border-right: 2px solid #555;
		}
		#tableSuperSum2.table > tbody > tr > th:nth-child(7) {
			border-right: 2px solid #555;
		}	

		#tableSuperSum2.table > tbody > tr > td:nth-child(10) {
			border-right: 2px solid #555;
		}
		#tableSuperSum2.table > tbody > tr > th:nth-child(10) {
			border-right: 2px solid #555;
		}	


		#tableSuperSum2.table > tbody > tr > td:nth-child(13) {
			border-right: 2px solid #555;
		}
		#tableSuperSum2.table > tbody > tr > th:nth-child(13) {
			border-right: 2px solid #555;
		}	


		#tableSuperSum2.table > tbody > tr > td:nth-child(16) {
			border-right: 2px solid #555;
		}
		#tableSuperSum2.table > tbody > tr > th:nth-child(16) {
			border-right: 2px solid #555;
		}	
		
		.smaller-header{
			font-size: 0.8em;
			vertical-align: top !important;
			text-align: center;
		}	
		
		.vpc-variance-header{
			background: #D43F3A;
			color:#fff;
			text-align:center;
			font-weight:bold;
		}

		.vpc-3pavg-header{
			background: #F0AD4E;
			color:#fff;
			text-align:center;
			font-weight:bold;
		}

		.vpc-lcp-header{
			background: #337AB7;
			color:#fff;
			text-align:center;
			font-weight:bold;
		}

		.vpc-current-header{
			background: #5CB85C;
			color:#fff;
			text-align:center;
			font-weight:bold;
		}
		
	
		.vpc-avgdailydales-header{
			background: #D43F3A;
			color:#fff;
			text-align:center;
			font-weight:bold;
		}

		.vpc-totalsales-header{
			background: #F0AD4E;
			color:#fff;
			text-align:center;
			font-weight:bold;
		}

		.vpc-totalcases-header{
			background: #337AB7;
			color:#fff;
			text-align:center;
			font-weight:bold;
		}

		.vpc-avgsellprice-header{
			background: #5CB85C;
			color:#fff;
			text-align:center;
			font-weight:bold;
		}
	
		
</style>



	 <div class="row">
	 
		<%
		NotEnoughFound = False
		'****************************************
		'Get info for first period being reported
		'****************************************
		FirstPeriod_TotalSales = 0
		FirstPeriod_TotalCases = 0
		SQL = "SELECT prodCategory, SUM(itemQuantity * itemPrice) AS TotSales, SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
		If VarianceBasis = "3Periods" Then
			SQL = SQL & " WHERE ivsDate BETWEEN '" & GetPeriodBeginDateBySeq(PeriodSeq-3) & "' AND '" & GetPeriodEndDateBySeq(PeriodSeq-1) & "' "
		Else
			SQL = SQL & " WHERE ivsDate BETWEEN '" & GetPeriodBeginDateBySeq(PeriodSeq-12) & "' AND '" & GetPeriodEndDateBySeq(PeriodSeq-1) & "' "
		End If
		SQL = SQL & " AND CustNum = " & CustForDetail & " "
		SQL = SQL & " AND prodCategory = " & SelectedCategoryID & " "
		SQL = SQL & " GROUP BY prodCategory "

		'Response.write(SQL)
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3
		rs.Open SQL, Session("ClientCnnString")
		If not rs.eof Then

		%>

		<table id="tableSuperSum" class="table table-striped table-condensed table-hover table-bordered standard-font">
		
	  	  	<%	
			Do
			%>

				<tr>
			   		<td width="20%">&nbsp;</td>
				    <td width="40%" colspan="2" align="center" class="vpc-avgdailydales-header">Average Daily Sales</td>                  
				    <td width="20%" class="vpc-totalsales-header">Total Sales</td>
				    <td width="20%" class="vpc-totalcases-header">Total Cases</td>
				    <td width="20%" class="vpc-avgsellprice-header">Avg Sell Price</td>
			    </tr>


			<%
		
				If VarianceBasis = "3Periods" Then		
					FirstPeriod_TotalSales = rs("TotSales")/3
					FirstPeriod_TotalCases = rs("TotCases")/3
				Else
					FirstPeriod_TotalSales = rs("TotSales")/12
					FirstPeriod_TotalCases = rs("TotCases")/12
				End If

				'*****************************************
				'Get info for second period being reported
				'*****************************************
				SecondPeriod_TotalSales = 0
				SecondPeriod_TotalCases = 0
				SQL = "SELECT SUM(itemQuantity * itemPrice) AS TotSales, SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
				SQL = SQL & " WHERE ivsDate BETWEEN '" & GetPeriodBeginDateBySeq(PeriodSeq) & "' AND '" & GetPeriodEndDateBySeq(PeriodSeq) & "' "
				SQL = SQL & " AND prodCategory ='" & rs("prodCategory") & "'"
				SQL = SQL & " AND CustNum = " & CustForDetail & " "
				'Response.write(SQL)
				Set rs2 = Server.CreateObject("ADODB.Recordset")
				rs2.CursorLocation = 3
				rs2.Open SQL, Session("ClientCnnString")
		
				If not rs2.eof Then
					SecondPeriod_TotalSales = rs2("TotSales")
					SecondPeriod_TotalCases = rs2("TotCases")
				Else
					NotEnoughFound = True	
				End If
				rs2.Close
		
				If FirstPeriod_TotalSales = "" or IsNull(FirstPeriod_TotalSales) or FirstPeriod_TotalCases = ""  or IsNull(FirstPeriod_TotalCases)_
				or SecondPeriod_TotalSales = ""  or IsNull(SecondPeriod_TotalSales) or SecondPeriod_TotalCases = ""  or IsNull(SecondPeriod_TotalCases) then NotEnoughFound = True
		
		
				If NotEnoughFound <> True Then
		
					'**************************************
					' Do all the calcs we will need up here
					'**************************************
					FirstPeriod_AVGSellPrice = FirstPeriod_TotalSales /  FirstPeriod_TotalCases
					SecondPeriod_AVGSellPrice = SecondPeriod_TotalSales /  SecondPeriod_TotalCases
					AVGSellPriceDifference =  SecondPeriod_AVGSellPrice - FirstPeriod_AVGSellPrice
					PriceChange = SecondPeriod_TotalCases * AVGSellPriceDifference
					VolumeChange = ((SecondPeriod_TotalCases - FirstPeriod_TotalCases) * FirstPeriod_AVGSellPrice)

					GoalForCurrentPeriod_Dollars = FirstPeriod_TotalSales - ((SecondPeriod_TotalSales-FirstPeriod_TotalSales))
					CurrentPeriodSalesForGoalDisplay_Dollars = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeq,SelectedCategoryID) + GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeq,SelectedCategoryID)

					GoalForCurrentPeriod_Cases = Round(FirstPeriod_TotalCases,0) - ((SecondPeriod_TotalCases - FirstPeriod_TotalCases ))
					
					' Do a quick lookup to get the total number of cases sold in the Current period
					SQLtmp = "SELECT SUM(NumberOfCases) AS CurrentCaseTotal FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"			
					SQLtmp = SQLtmp & " WHERE  CustNum = " & CustForDetail & " AND ivsDate BETWEEN '" & GetPeriodBeginDateBySeq(PeriodSeq+1) & "' AND '" & GetPeriodEndDateBySeq(PeriodSeq+1) & "'"
					SQLtmp = SQLtmp & " AND prodCategory = " & SelectedCategoryID
 
					Set rstmp = Server.CreateObject("ADODB.Recordset")
					rs2.CursorLocation = 3
					rstmp.Open SQLtmp, Session("ClientCnnString")

					If Not rstmp.EOF Then
						If Not Isnull(rstmp("CurrentCaseTotal")) Then CurrentPeriodCasesForGoalDisplay_Cases = GoalForCurrentPeriod_Cases - rstmp("CurrentCaseTotal") Else CurrentPeriodCasesForGoalDisplay_Cases = GoalForCurrentPeriod_Cases
					Else
						CurrentPeriodCasesForGoalDisplay_Cases = GoalForCurrentPeriod_Cases				
					End If
					Set rstmp = Nothing					
					
					%> 
			              <tr>
			              
								<%If VarianceBasis = "3Periods" Then
									Response.Write("<td><strong>Three Period Avg (" & GetPeriodBySeq(PeriodSeq-1) & "-" &  GetPeriodBySeq(PeriodSeq-3) & ")</strong></td>")
									
									'Show the avg number of work days
									If (WorkDaysInPeriodBasis/3) - cint(WorkDaysInPeriodBasis/3) <> 0 Then RoundingAsterik ="*" Else RoundingAsterik = ""
									Response.Write("<td><small class='small-date'>(" & WorkDaysInPeriodBasis & "/3)= " & Round(WorkDaysInPeriodBasis/3,1)  & RoundingAsterik & " biz days avg</small></td>")
								Else
									Response.Write("<td><strong>Twelve Period Avg (" & GetPeriodBySeq(PeriodSeq-1) & "-" &  GetPeriodBySeq(PeriodSeq-12) & ")</strong></td>")
									'Show the avg number of work days
									If (WorkDaysInPeriodBasis/12) - cint(WorkDaysInPeriodBasis/12) <> 0 Then RoundingAsterik ="*" Else RoundingAsterik = ""
									Response.Write("<td><small class='small-date'>(" & WorkDaysInPeriodBasis & "/12)= " & Round(WorkDaysInPeriodBasis/12,1)  & RoundingAsterik  & " biz days avg</small></td>")
								End If
								%>
								
								
								<%If VarianceBasis = "3Periods" Then %>
								      <td><%= FormatCurrency(FirstPeriod_TotalSales/(WorkDaysInPeriodBasis/3),0,-2,-1)%></td>
								<% Else %>
								      <td><%= FormatCurrency(FirstPeriod_TotalSales/(WorkDaysInPeriodBasis/12),0,-2,-1)%></td>
								<% End If%>
								<td><%= FormatCurrency(FirstPeriod_TotalSales,0,-2,-1)%></td>  
								<% If (FirstPeriod_TotalCases) - cint(FirstPeriod_TotalCases) <> 0 Then RoundingAsterik ="*" Else RoundingAsterik = ""%>
								<td><%= FormatNumber(FirstPeriod_TotalCases,0)%><%=RoundingAsterik%></td>
								<td><%= FormatCurrency(FirstPeriod_AVGSellPrice,2,-2,-1) %></td>
								
			              </tr>
			              
			              <tr>
				              <%
				              	Response.Write("<td><strong>Period " & GetPeriodAndYearBySeq(PeriodSeq) & "</strong>")
				              	Response.Write("<td><small class='small-date'>" & WorkDaysInLastClosedPeriod & " biz days</small></td>")
				              %>

				              <td><%= FormatCurrency(SecondPeriod_TotalSales/WorkDaysInLastClosedPeriod ,0,-2,-1)%></td>
				              <td><%= FormatCurrency(SecondPeriod_TotalSales,0,-2,-1)%></td> 
							  <% If (SecondPeriod_TotalCases) - cint(SecondPeriod_TotalCases) <> 0 Then RoundingAsterik ="*" Else RoundingAsterik = ""%>			                           
				              <td><%= FormatNumber(SecondPeriod_TotalCases,0)%><%=RoundingAsterik%></td>
				              <td><%= FormatCurrency(SecondPeriod_AVGSellPrice,2,-2,-1) %></td>
				          </tr>
				              
				              
				          <tr>
				          
				              <td><b>Variance</b></td>
				              <td><small class='small-date'>LCP .vs Prior&nbsp;<% If VarianceBasis="3Periods" Then Response.Write("Three") Else Response.Write("Twelve")%></small></td>
				              <td>&nbsp;</td>
				              
				              
	   						  <% If (SecondPeriod_TotalSales - FirstPeriod_TotalSales) - cint(SecondPeriod_TotalSales - FirstPeriod_TotalSales) <> 0 Then RoundingAsterik ="*" Else RoundingAsterik = ""
			              		If SecondPeriod_TotalSales - FirstPeriod_TotalSales < 0 Then %>
									<td class="negative"><%= FormatCurrency(SecondPeriod_TotalSales - FirstPeriod_TotalSales,0,-2,-1)%><%=RoundingAsterik%></td>
								<% ElseIF SecondPeriod_TotalSales - FirstPeriod_TotalSales = 0 Then %>
									<td><strong>---</strong></td>
								<% Else %>
									<td class="positive"><%= FormatCurrency(SecondPeriod_TotalSales - FirstPeriod_TotalSales,0,-2,-1)%><%=RoundingAsterik%></td>
								<%End If %>
				              
				              
	   						  <% If (SecondPeriod_TotalCases - FirstPeriod_TotalCases) - cint(SecondPeriod_TotalCases - FirstPeriod_TotalCases) <> 0 Then RoundingAsterik ="*" Else RoundingAsterik = ""
			              		If SecondPeriod_TotalCases - FirstPeriod_TotalCases < 0 Then %>
									<td class="negative"><%= FormatNumber(SecondPeriod_TotalCases - FirstPeriod_TotalCases,0)%><%=RoundingAsterik%></td>
								<% ElseIF SecondPeriod_TotalCases - FirstPeriod_TotalCases = 0 Then %>
									<td><strong>---</strong></td>
								<% Else %>
									<td class="positive"><%= FormatNumber(SecondPeriod_TotalCases - FirstPeriod_TotalCases,0)%><%=RoundingAsterik%></td>
								<%End If %>
				              <td><%= FormatCurrency(AVGSellPriceDifference,2,-2,-1) %></td>
				              
			              </tr>
			              
			              
			              
			              <tr>
				              <td><strong>Goal for P<%= CurrentPeriodNumber %></strong></td>
				              <td><small class='small-date'><%= WorkDaysInCurrentPeriod%>&nbsp;biz days</small></td>
				              <td><%= FormatCurrency(GoalForCurrentPeriod_Dollars/WorkDaysInCurrentPeriod ,0,-2,-1)%></td> 
				              <td><%= FormatCurrency(GoalForCurrentPeriod_Dollars,0)%></td>   
				              <td><%= FormatNumber(GoalForCurrentPeriod_Cases,0)%></td>           
			              </tr>
			              
			              
			              <tr>
				              <td><strong>Current&nbsp;P<%= CurrentPeriodNumber %></strong></td>
				              <td><small class='small-date'><%= WorkDaysSoFar %>&nbsp;biz days so far</small></td>
				              <td><%= FormatCurrency(CurrentPeriodSalesForGoalDisplay_Dollars / WorkDaysSoFar  ,0)%></td>
				              <td><%= FormatCurrency(CurrentPeriodSalesForGoalDisplay_Dollars,0)%></td>
	   			              <td><%= FormatNumber(CurrentPeriodCasesForGoalDisplay_Cases,0)%></td> 
			              </tr>
			              
			              
			              <tr>
				              <td><strong>Still need in P<%= CurrentPeriodNumber %></strong></td>
				              <td><small class='small-date'><%=WorkDaysInCurrentPeriod - WorkDaysSoFar + 1 %> biz days left</small></td>
				              <td>&nbsp;</td>
				              <td><%= FormatCurrency(GoalForCurrentPeriod_Dollars-CurrentPeriodSalesForGoalDisplay_Dollars,0)%></td>	
				              <td><%= FormatNumber(GoalForCurrentPeriod_Cases-CurrentPeriodCasesForGoalDisplay_Cases,0)%></td>		              
			              </tr>
			              
					<% End If 
			rs.movenext
		Loop until rs.eof
		rs.Close
	   %>	       
	</table>

</div>
<% End If %>


<%''''''''''''''''''''''''''''''''''''''''
'This is where all the VPC2 stuff starts
'''''''''''''''''''''''''''''''''''''''' 
If VarianceBasis = "3Periods" Then
	FirstRangeStartDate = GetPeriodBeginDateBySeq(PeriodSeq-3)
	FirstRangeEndDate = GetPeriodEndDateBySeq(PeriodSeq-1)
Else
	FirstRangeStartDate = GetPeriodBeginDateBySeq(PeriodSeq-12)
	FirstRangeEndDate = GetPeriodEndDateBySeq(PeriodSeq-1)
End If

SecondRangeStartDate = GetPeriodBeginDateBySeq(PeriodSeq)
SecondRangeEndDate = GetPeriodEndDateBySeq(PeriodSeq)

'Just in case a dash got in there, strip it & trim also
FirstRangeStartDate = trim(Replace(FirstRangeStartDate,"-"," "))
FirstRangeEndDate = trim(Replace(FirstRangeEndDate,"-"," "))
SecondRangeStartDate = trim(Replace(SecondRangeStartDate,"-"," "))
SecondRangeEndDate = trim(Replace(SecondRangeEndDate,"-"," "))

FirstRangeStartDate = FormatDateTime(FirstRangeStartDate)
FirstRangeEndDate = FormatDateTime(FirstRangeEndDate)
SecondRangeStartDate = FormatDateTime(SecondRangeStartDate)
SecondRangeEndDate = FormatDateTime(SecondRangeEndDate)
%>


 <div class="row">
<%
'Because we might have products sold in one period that were not in the other, we need to get a list of all skus from both periods
'Create Sku list work table
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rsSKUList = Server.CreateObject("ADODB.Recordset")
rsSKUList.CursorLocation = 3 

' Drop & create temporary table
on error resume next
SQL = "DROP TABLE zReportSKUList_" & Trim(Session("userNo"))
Set rsSKUList = cnn8.Execute(SQL)
on error goto 0


'Get first list of SKUs
SQL = "SELECT Distinct partNum INTO zReportSKUList_" & Trim(Session("userNo")) & " FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
SQL = SQL & " WHERE CustNum = " & CustForDetail & " AND partNum Is Not Null AND prodCategory = " & SelectedCategoryID & " AND ((ivsDate BETWEEN '" & FirstRangeStartDate & "' AND '" & FirstRangeEndDate & "') OR"
SQL = SQL & " (ivsDate BETWEEN '" & SecondRangeStartDate & "' AND '" & SecondRangeEndDate & "')) "

'response.write(SQL & "<br><br>")

Set rsSKUList = cnn8.Execute(SQL)




NotEnoughFound = False


'Try this, preprocess to get the proper sort order
'*************************************************
'*************************************************
'*************************************************
'*************************************************
'*************************************************
'*************************************************
SQL = "SELECT DISTINCT partNum as prodSKU FROM zReportSKUList_" & Session("UserNo") 
Set rsOuter = Server.CreateObject("ADODB.Recordset")
rsOuter.CursorLocation = 3
rsOuter.Open SQL, Session("ClientCnnString")



If not rsOuter.eof Then

		ReDim ProdSortArray	(rsOuter.RecordCount,2)
		ArrayCounter =0
		
		Do While Not rsOuter.EOF
		

			'****************************************
			'Get info for first period being reported
			'****************************************

			FirstPeriod_TotalSales = 0
						
			SQL = "SELECT partnum, SUM(itemQuantity * itemPrice) AS TotSales FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
			SQL = SQL & " WHERE CustNum = " & CustForDetail & " AND ivsDate BETWEEN '" & FirstRangeStartDate  & "' AND '" & FirstRangeEndDate   & "'"
			SQL = SQL & " AND partnum = '" & rsOuter("prodSKU") & "' "
			SQL = SQL & " Group By partnum"

			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3
			rs.Open SQL, Session("ClientCnnString")
			
			If not rs.eof Then
				If VarianceBasis = "3Periods" Then
					FirstPeriod_TotalSales = rs("TotSales") / 3
				Else
					FirstPeriod_TotalSales = rs("TotSales") / 12
				End If
			Else
				FirstPeriod_TotalSales = 0
			End If

			If NOT IsNumeric(FirstPeriod_TotalSales) Then FirstPeriod_TotalSales = 0

	
		'*****************************************
		'Get info for second period being reported
		'*****************************************
		SecondPeriod_TotalSales = 0
		
		SQL = "SELECT  SUM(itemQuantity * itemPrice) AS TotSales FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
		SQL = SQL & " WHERE  CustNum = " & CustForDetail & " AND ivsDate BETWEEN '" & SecondRangeStartDate & "' AND '" & SecondRangeEndDate & "'"
		SQL = SQL & " AND partnum = '" & rsOuter("prodSKU") & "'"
	

		Set rs2 = Server.CreateObject("ADODB.Recordset")
		rs2.CursorLocation = 3
		rs2.Open SQL, Session("ClientCnnString")
		If not rs2.eof Then
			SecondPeriod_TotalSales = rs2("TotSales")
		Else
			SecondPeriod_TotalSales = 0
		End If
		rs2.Close
	
		If NOT IsNumeric(SecondPeriod_TotalSales) Then SecondPeriod_TotalSales = 0
	
		ProdSortArray(ArrayCounter,0) = rsOuter("prodSKU")
		ProdSortArray(ArrayCounter,1) = (SecondPeriod_TotalSales - FirstPeriod_TotalSales)
		
		ArrayCounter = ArrayCounter + 1
				
		rsouter.MoveNext
	Loop
	
	'Now sort the array
    for i = UBound(ProdSortArray) - 1 To 0 Step -1
	    for j= 0 to i
	        if ProdSortArray(j,1)>ProdSortArray(j+1,1) then
	            temp=ProdSortArray(j+1,0)
	            temp2=ProdSortArray(j+1,1)
	            ProdSortArray(j+1,0)=ProdSortArray(j,0)
   	            ProdSortArray(j+1,1)=ProdSortArray(j,1)
	            ProdSortArray(j,0)=temp
   	            ProdSortArray(j,1)=temp2
	        end if
	    next
	next
	sortArray = ProdSortArray

	'Now build the order by clause based on the sorted array
	CLAUSE = " ORDER BY CASE partnum "
	
	ClauseCount = 0 
	for i = 0 to UBound(sortArray) 
		CLAUSE = CLAUSE & " WHEN '" & sortArray(i,0) & "' THEN " & ClauseCount
		ClauseCount = ClauseCount + 1
	next
	CLAUSE = CLAUSE & " END "
	'Response.Write(CLAUSE & "<br>")
End If



'*************************************************
'*************************************************
'*************************************************
'*************************************************
'*************************************************
'*************************************************
' eof Try this, preprocess to get the proper sort order

SQL = "SELECT partNum as prodSKU FROM "
SQL = SQL & " zReportSKUList_" & Session("UserNo") 
SQL = SQL & CLAUSE
'Response.Write(SQL & "<br>")
Set rsOuter = Server.CreateObject("ADODB.Recordset")
rsOuter.CursorLocation = 3
rsOuter.Open SQL, Session("ClientCnnString")
	
If not rsOuter.eof Then%>

	<!-- sort table script !-->
	<script src="../../../js/sorttable.js"></script>
	<script src="../../../js/sorttable1.js"></script>
	<!-- eof sort table script !-->
	
	
	<script>
		$(window).load(function() 
		{
		   // executes when complete page is fully loaded, including all frames, objects and images
		   //alert("(window).load was called - window is loaded!");
		   sorttable.innerSortFunction.call(document.getElementById('salesColumn'));
		});  
		
	</script>

	<div class="table-responsive">


		<table id="tableSuperSum2" class="table table-striped table-condensed table-hover sortable">
		
		<thead>
	
			<tr>
				<th class="sorttable_nosort" width="8%">Item #</th>
				<th class="sorttable_nosort" width="20%">Description</th>
				<th class="sorttable_nosort td-align smaller-header">Price Chg</th>  
				<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">Sales</th>  
				<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;">Cases</th>  
				<th class="sorttable_nosort td-align smaller-header" style="border-top: 2px solid #555 !important;">Price Impact</th>  
				<th class="sorttable_nosort td-align smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;">Volume Impact</th>
				<th class="sorttable_nosort td-align smaller-header" style="border-top: 2px solid #555 !important;">Sales</th>
				<th class="sorttable_nosort td-align smaller-header" style="border-top: 2px solid #555 !important;">Cases</th>
				<th class="sorttable_nosort td-align smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;">Avg Price</th>  
				<th class="sorttable_nosort td-align smaller-header" style="border-top: 2px solid #555 !important;">Sales</th>
				<th class="sorttable_nosort td-align smaller-header" style="border-top: 2px solid #555 !important;">Cases</th>
				<th class="sorttable_nosort td-align smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;">Avg Price</th> 
				<th class="sorttable_nosort td-align smaller-header" style="border-top: 2px solid #555 !important;">Sales</th>
				<th class="sorttable_nosort td-align smaller-header" style="border-top: 2px solid #555 !important;">Cases</th>
				<th class="sorttable_nosort td-align smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;">Avg Price</th> 
			</tr>
		</thead>
		<thead>
			<tr>	
				<td colspan="3" style="border-right: 2px solid #555 !important;">&nbsp;</td>
				<td class="td-align1 vpc-variance-header"" colspan="4" style="border-right: 2px solid #555 !important;">Variance</td>
				<%If VarianceBasis = "3Periods" Then %>
					<td class="td-align1 vpc-3pavg-header" colspan="3" style="border-right: 2px solid #555 !important;">3 Period Avg</td>
				<%Else %>
					<td class="td-align1 vpc-3pavg-header" colspan="3" style="border-right: 2px solid #555 !important;">12 Period Avg</td>
				<%End If%>
				<td class="td-align1 vpc-lcp-header" colspan="3" style="border-right: 2px solid #555 !important;">Last Closed Period</td>
				<td class="td-align1 vpc-current-header" colspan="3" style="border-right: 2px solid #555 !important;">Current Period</td>
			</tr>
		</thead>
		
		<tbody>

		<%

		Do While Not rsOuter.EOF

			'****************************************
			'Get info for first period being reported
			'****************************************

			FirstPeriod_TotalSales = 0
			FirstPeriod_TotalCases = 0
			FirstPeriod_AVGSellPrice = 0
						
			SQL = "SELECT partnum, SUM(itemQuantity * itemPrice) AS TotSales, SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
			SQL = SQL & " WHERE CustNum = " & CustForDetail & " AND ivsDate BETWEEN '" & FirstRangeStartDate  & "' AND '" & FirstRangeEndDate   & "'"
			SQL = SQL & " AND partnum = '" & rsOuter("prodSKU") & "' "
			SQL = SQL & " Group By partnum"
'Response.Write(SQL & "<br>")
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3
			rs.Open SQL, Session("ClientCnnString")
			
			If not rs.eof Then
				If VarianceBasis = "3Periods" Then
					FirstPeriod_TotalSales = rs("TotSales") / 3
					FirstPeriod_TotalCases = rs("TotCases") / 3
				Else
					FirstPeriod_TotalSales = rs("TotSales") / 12
					FirstPeriod_TotalCases = rs("TotCases") / 12
				End If
			Else
				FirstPeriod_TotalSales = 0
				FirstPeriod_TotalCases = 0
				FirstPeriod_AVGSellPrice = 0
			End If

			
			If NOT IsNumeric(FirstPeriod_TotalSales) Then FirstPeriod_TotalSales = 0
			If NOT IsNumeric(FirstPeriod_TotalCases) Then FirstPeriod_TotalCases = 0

	
		'*****************************************
		'Get info for second period being reported
		'*****************************************
		SecondPeriod_TotalSales = 0
		SecondPeriod_TotalCases = 0
		SecondPeriod_AVGSellPrice = 0
		
		SQL = "SELECT  SUM(itemQuantity * itemPrice) AS TotSales, SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
		SQL = SQL & " WHERE  CustNum = " & CustForDetail & " AND ivsDate BETWEEN '" & SecondRangeStartDate & "' AND '" & SecondRangeEndDate & "'"
		SQL = SQL & " AND partnum = '" & rsOuter("prodSKU") & "'"
	
	
		'Response.write(SQL)
		Set rs2 = Server.CreateObject("ADODB.Recordset")
		rs2.CursorLocation = 3
		rs2.Open SQL, Session("ClientCnnString")
		If not rs2.eof Then
			SecondPeriod_TotalSales = rs2("TotSales")
			SecondPeriod_TotalCases = rs2("TotCases")
		Else
			SecondPeriod_TotalSales = 0
			SecondPeriod_TotalCases = 0
		End If
		rs2.Close
	
		If NOT IsNumeric(SecondPeriod_TotalSales) Then SecondPeriod_TotalSales = 0
		If NOT IsNumeric(SecondPeriod_TotalCases) Then SecondPeriod_TotalCases = 0
	
	
		'**************************************
		' Do all the calcs we will need up here
		'**************************************
		If FirstPeriod_TotalCases <> 0 Then 
			FirstPeriod_AVGSellPrice = FirstPeriod_TotalSales /  FirstPeriod_TotalCases
		Else
			FirstPeriod_AVGSellPrice = 0
		End if
		If SecondPeriod_TotalCases <> 0 Then 
			SecondPeriod_AVGSellPrice = SecondPeriod_TotalSales /  SecondPeriod_TotalCases
		Else
			SecondPeriod_AVGSellPrice = 0
		End IF
		
		AVGSellPriceDifference =  SecondPeriod_AVGSellPrice - FirstPeriod_AVGSellPrice

		'Only if there is an actual price difference & there we sales in both periods
		If FirstPeriod_TotalCases <> 0 And SecondPeriod_TotalCases <> 0 Then
			If Round(FirstPeriod_AVGSellPrice,2) <> Round(SecondPeriod_AVGSellPrice,2) Then
				PriceChange = ((FirstPeriod_AVGSellPrice - SecondPeriod_AVGSellPrice) * (SecondPeriod_TotalCases)) * -1
			Else
				PriceChange = 0
			End If
		Else
			PriceChange = 0
		End If

		VolumeChange = ((SecondPeriod_TotalSales - FirstPeriod_TotalSales)) - PriceChange 
		
		If FirstPeriod_TotalCases = 0 Or SecondPeriod_TotalCases = 0 Then
			AVGSellPriceDifference = 0
		End If

		
		'**************
		' Grand totals 
		'**************
		VarianceGrandTot_Sales = VarianceGrandTot_Sales + (SecondPeriod_TotalSales - FirstPeriod_TotalSales)
		VarianceGrandTot_Cases = VarianceGrandTot_Cases + (SecondPeriod_TotalCases - FirstPeriod_TotalCases)
		VarianceGrandTot_PriceChange = VarianceGrandTot_PriceChange + PriceChange
		VarianceGrandTot_VolumeChange = VarianceGrandTot_VolumeChange + VolumeChange 
		
		VarianceBasisGrandTot_Sales = VarianceBasisGrandTot_Sales  + FirstPeriod_TotalSales
		VarianceBasisGrandTot_Cases = VarianceBasisGrandTot_Cases + FirstPeriod_TotalCases
		
		PeriodBeingEvalGrandTot_Sales = PeriodBeingEvalGrandTot_Sales  + SecondPeriod_TotalSales
		PeriodBeingEvalGrandTot_Cases = PeriodBeingEvalGrandTot_Cases + SecondPeriod_TotalCases

		%> 
	
		<tr>
			<td><small><%= rsOuter("prodSKU")%></small></td>
			<td><small><%= GetProdDescriptionFromInvDetsByPartnum(rsOuter("prodSKU")) %></small></td>
			<% 
			' Price change only valid if not 0 and if there are sales in
			' both periods being evaluates
			If FirstPeriod_TotalCases <> 0 And SecondPeriod_TotalCases <> 0 Then
				If Round(FirstPeriod_AVGSellPrice,2) <> Round(SecondPeriod_AVGSellPrice,2) Then %>
					<td class="td-align"><%= FormatCurrency(SecondPeriod_AVGSellPrice - FirstPeriod_AVGSellPrice,2,-2,0)%></td> 
				<% Else %>
					<td class="td-align">---</td> 
				<% End If
			Else %>
				<td class="td-align">---</td> 
			<%End If%>
	
			<% If (SecondPeriod_TotalSales - FirstPeriod_TotalSales) > 0 Then %>
				<td class="td-align positive"><%= FormatCurrency(SecondPeriod_TotalSales - FirstPeriod_TotalSales,0)%></td> 
			<% ElseIf (SecondPeriod_TotalSales - FirstPeriod_TotalSales) < 0 Then %>
				<td class="td-align negative"><%= FormatCurrency(SecondPeriod_TotalSales - FirstPeriod_TotalSales,0)%></td>
			<% ElseIf (SecondPeriod_TotalSales - FirstPeriod_TotalSales) = 0 Then %>
				<td class="td-align"><%= FormatCurrency(SecondPeriod_TotalSales - FirstPeriod_TotalSales,0)%></td>
			<% End If %>
			<%
			'If the rounded result is 0 then just print 0, otherwise
			'it sometimes prints a negative 0 due to fractional results
			If Round(SecondPeriod_TotalCases - FirstPeriod_TotalCases) = 0 Then %>
				<td class="td-align"><strong><%= FormatNumber(0,0)%></strong></td>
			<%Else
				If (SecondPeriod_TotalCases - FirstPeriod_TotalCases) - cint(SecondPeriod_TotalCases - FirstPeriod_TotalCases) <> 0 Then RoundingAsterik ="*" Else RoundingAsterik = ""
				If SecondPeriod_TotalCases - FirstPeriod_TotalCases < 0 Then %>
					<td class="td-align negative">(<%= FormatNumber(SecondPeriod_TotalCases - FirstPeriod_TotalCases,0)%>)<%=RoundingAsterik%></td>
				<% Else %>
					<td class="td-align positive"><%= FormatNumber(SecondPeriod_TotalCases - FirstPeriod_TotalCases,0)%><%=RoundingAsterik%></td>
				<%End If
			End If
			If PriceChange = 0 Then %>
				<td class="td-align">---</td>
			<% Else %>
				<td class="td-align"><%= FormatCurrency(PriceChange,2,-2,0) %></td>			
			<% End If %>
			<td class="td-align"><%= FormatCurrency(VolumeChange,0) %></td> 
			
			<%
			' Print --- if everything is 0
			If FirstPeriod_TotalCases = 0 Then %>
				<td class="td-align">---</td>              
				<td class="td-align">---</td>
				<td class="td-align">---</td>
			<% Else %>
				<td class="td-align"><%= FormatCurrency(FirstPeriod_TotalSales,0)%></td>   
				<%If FirstPeriod_TotalCases - cint(FirstPeriod_TotalCases) <> 0 Then RoundingAsterik ="*" Else RoundingAsterik = ""%>        
				<td class="td-align"><%= FormatNumber(FirstPeriod_TotalCases,0)%><%=RoundingAsterik%></td>
				<% If FirstPeriod_AVGSellPrice = 0 Then %>
					<td class="td-align">---</td>
				<% Else %>
					<td class="td-align"><%= FormatCurrency(FirstPeriod_AVGSellPrice,2,-2,0) %></td>
				<% End If
			End If 
			
			If SecondPeriod_TotalCases = 0 Then %>
				<td class="td-align">---</td>              
				<td class="td-align">---</td>
				<td class="td-align">---</td>
			<% Else %>
				<td class="td-align"><%= FormatCurrency(SecondPeriod_TotalSales,0)%></td>      
				<%If SecondPeriod_TotalCases - cint(SecondPeriod_TotalCases) <> 0 Then RoundingAsterik ="*" Else RoundingAsterik = ""%>                
				<td class="td-align"><%= FormatNumber(SecondPeriod_TotalCases,0)%><%=RoundingAsterik%></td>
				<% If SecondPeriod_AVGSellPrice = 0 Then %>
					<td class="td-align">&nbsp;</td>
				<% Else %>
					<td class="td-align"><%= FormatCurrency(SecondPeriod_AVGSellPrice,2,-2,0) %></td>
				<% End If 
			End If 
			
			CurrentPeriod_TotalSales = 0
			CurrentPeriod_TotalCases = 0

			If Session("CalcTax") = True Then
				SQL = "SELECT SUM(CASE WHEN prodTaxable = 'Y' THEN (itemPrice*itemQuantity) + ((itemPrice*itemQuantity) * (prodTaxPercent / 100)) WHEN prodTaxable <> 'Y' THEN (itemPrice*itemQuantity) END ) AS TotSales, SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
			Else
				SQL = "SELECT SUM(itemPrice*itemQuantity) AS TotSales, SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"			
			End If				

			SQL = SQL & " WHERE  CustNum = " & CustForDetail & " AND ivsDate BETWEEN '" & GetPeriodBeginDateBySeq(PeriodSeq+1) & "' AND '" & GetPeriodEndDateBySeq(PeriodSeq+1) & "'"
			SQL = SQL & " AND partnum = '" & rsOuter("prodSKU") & "'"

			Set rs2 = Server.CreateObject("ADODB.Recordset")
			rs2.CursorLocation = 3
			rs2.Open SQL, Session("ClientCnnString")
			If not rs2.eof Then
				CurrentPeriod_TotalSales = rs2("TotSales")
				CurrentPeriod_TotalCases = rs2("TotCases")
			Else
				CurrentPeriod_TotalSales = 0
				CurrentPeriod_TotalCases = 0
			End If
			rs2.Close
	
			If NOT IsNumeric(CurrentPeriod_TotalSales) Then CurrentPeriod_TotalSales = 0
			If NOT IsNumeric(CurrentPeriod_TotalCases) Then CurrentPeriod_TotalCases = 0
			
			
			'**************
			' Grand totals 
			'**************
			CurrentPostedUnPostedSales = CurrentPeriod_TotalSales + GetUnposedSalesByCustByProd(CustForDetail ,rsOuter("prodSKU"),PeriodSeq+1)
			CurrentPostedUnPostedCases = CurrentPeriod_TotalCases + GetUnposedCasesByCustByProd(CustForDetail ,rsOuter("prodSKU"),PeriodSeq+1)

			CurrentPeriodGrandTot_Sales = CurrentPeriodGrandTot_Sales  + CurrentPostedUnPostedSales 
			CurrentPeriodGrandTot_Cases = CurrentPeriodGrandTot_Cases + CurrentPostedUnPostedCases 

			
			If CurrentPostedUnPostedCases = 0 Then %>
				<td class="td-align">---</td>              
				<td class="td-align">---</td>
				<td class="td-align">---</td>
			<% Else %>
				<td class="td-align"><%= FormatCurrency(CurrentPostedUnPostedSales ,0)%></td>
				<%If CurrentPostedUnPostedCases  - cint(CurrentPostedUnPostedCases ) <> 0 Then RoundingAsterik ="*" Else RoundingAsterik = "" %>
				<td class="td-align"><%= FormatNumber(CurrentPostedUnPostedCases ,0)%><%=RoundingAsterik%></td>
				<%If CurrentPostedUnPostedCases  <> 0 Then 
					CurrentPeriod_AVGSellPrice = CurrentPostedUnPostedSales  /  CurrentPostedUnPostedCases 
				Else
					CurrentPeriod_AVGSellPrice = 0
				End IF %>
				<% If CurrentPeriod_AVGSellPrice = 0 Then %>
					<td class="td-align">&nbsp;</td>
				<% Else %>
					<td class="td-align"><%= FormatCurrency(CurrentPeriod_AVGSellPrice,2) %></td>
				<% End If
			End If %>
		</tr>
		              
	   	              	              
		<%
			rsOuter.MoveNext
		
		Loop
	
	    
	    'Grand Totals%>
	    </tbody>
	    	    
	    <tfoot>
	   		<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td class="td-align"><strong>Totals</strong></td> 
				<% If VarianceGrandTot_Sales < 0 Then %>
					<td class="td-align negative"><%= FormatCurrency(VarianceGrandTot_Sales,0)%></td> 
				<% Else %>
					<td class="td-align positive"><%= FormatCurrency(VarianceGrandTot_Sales,0)%></td> 			
				<%End If%>
				<% If VarianceGrandTot_Cases < 0 Then %>
					<td class="td-align negative">(<%= FormatNumber(VarianceGrandTot_Cases,0)%>)</td>
				<% Else %>
					<td class="td-align positive"><%= FormatNumber(VarianceGrandTot_Cases,0)%></td>			
				<% End If %>
				<% If VarianceGrandTot_PriceChange = 0 Then %>
					<td>&nbsp;</td>
				<% ElseIf VarianceGrandTot_PriceChange > 0 Then %>
					<td class="td-align"><strong><%= FormatCurrency(VarianceGrandTot_PriceChange,0)%></strong></td> 
				<% Else %>
					<td class="td-align negative">(<%= FormatCurrency(VarianceGrandTot_PriceChange,0)%>)</td> 			
				<% End If %>
				<td class="td-align"><strong><%= FormatCurrency(VarianceGrandTot_VolumeChange,0)%></strong></td> 
				<td class="td-align"><strong><%= FormatCurrency(VarianceBasisGrandTot_Sales,0)%></strong></td> 
				<td class="td-align"><strong><%= FormatNumber(VarianceBasisGrandTot_Cases,0)%></strong></td>
				<td>&nbsp;</td>
				<td class="td-align"><strong><%= FormatCurrency(PeriodBeingEvalGrandTot_Sales,0)%></strong></td> 
				<td class="td-align"><strong><%= FormatNumber(PeriodBeingEvalGrandTot_Cases,0)%></strong></td>
				<td>&nbsp;</td>
				<td class="td-align"><strong><%= FormatCurrency(CurrentPeriodGrandTot_Sales,0)%></strong></td> 
				<td class="td-align"><strong><%= FormatNumber(CurrentPeriodGrandTot_Cases,0)%></strong></td>
				<td>&nbsp;</td>
			</tr>
		</tfoot>
	    
	    <tr><small>* An asterik indicates that rounding has occurred</small></tr>
	    

	    </table>
		</div>
		
	</div>
<%End If


End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************





'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetTitleForEquipmentVPCModal() 

	CustIDPassed = Request.Form("CustID")
	CustName = GetCustNameByCustNum(CustIDPassed)
	LCPGP = Request.Form("LCPGP")
	TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(CustIDPassed)
	%>
	<h3><%= CustName %></h3>
	
	<% If LCPGP <> 0 Then
		If (cInt(TotalEquipmentValue/LCPGP)) < 10 Then %>
			<div class="tile blue">
				<h4 class="title">ROI <%= Round(TotalEquipmentValue/LCPGP,1) %></h4>
			</div>	
		<% Else %>
			<div class="tile red">
				<h4 class="title">ROI <%= Round(TotalEquipmentValue/LCPGP,1) %></h4>
			</div>	
		<% End If
	End If

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetContentForEquipmentVPCModal() 
%>
	<%
	CustIDPassed = Request.Form("CustID")
	CustName = GetCustNameByCustNum(CustIDPassed)
	
	TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(CustIDPassed)
	
	%>
	<div align="right">
		<h3 style="margin-top:0px;">Total Value <span style="color:green; font-weight:bold"><%= FormatCurrency(TotalEquipmentValue,2) %></span></h3>
	</div>
	<%
	
	Set rsCustomerEquipmentByClass = Server.CreateObject("ADODB.Recordset")
	rsCustomerEquipmentByClass.CursorLocation = 3 

	Set rsCustomerEquipment = Server.CreateObject("ADODB.Recordset")
	rsCustomerEquipment.CursorLocation = 3 
	
	
	Set rsEquipStatusCode = Server.CreateObject("ADODB.Recordset")
	rsEquipStatusCode.CursorLocation = 3 
	
		
	SQLCustomerEquipmentByClass = "SELECT EQ_Classes.Class, EQ_Classes.InternalRecordIdentifier, SUM(EQ_Equipment.PurchaseCost) AS Expr1 "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " FROM EQ_CustomerEquipment INNER JOIN "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " EQ_Equipment ON EQ_CustomerEquipment.EquipIntRecID = EQ_Equipment.InternalRecordIdentifier INNER JOIN "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier INNER JOIN "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " EQ_Classes ON EQ_Models.ClassIntRecID = EQ_Classes.InternalRecordIdentifier "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " WHERE        (EQ_CustomerEquipment.CustID = " & CustIDPassed & ") "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " GROUP BY EQ_Classes.Class, EQ_Classes.InternalRecordIdentifier "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " ORDER BY Expr1 DESC"
		
	Set cnnCustomerEquipmentByClass = Server.CreateObject("ADODB.Connection")
	cnnCustomerEquipmentByClass.open (Session("ClientCnnString"))
	Set rsCustomerEquipmentByClass = cnnCustomerEquipmentByClass.Execute(SQLCustomerEquipmentByClass)
	
	If NOT rsCustomerEquipmentByClass.EOF Then
	
		Do While NOT rsCustomerEquipmentByClass.EOF
		
			ClassName = rsCustomerEquipmentByClass("Class")
			ClassIntRecID = rsCustomerEquipmentByClass("InternalRecordIdentifier")
			ClassTotalEquipValue = rsCustomerEquipmentByClass("Expr1")
	
	
			%>	
			<h3><%= ClassName %>&nbsp;<span style="color:green;"><%= FormatCurrency(ClassTotalEquipValue,2) %></span></h3>
			<table class="table table-condensed table-hover large-table">			
				<thead>
				  <tr style="background-color: #EEE;">
				  	<th style="width: 3%;">+</th>
				  	<th style="width: 25%;">Description/Type</th>
				  	<th>Status</th>
				  	<th>Frequency</th>
				  	<th style="text-align: center;">Rent $</th>
				  	<th style="text-align: center;">Install Date</th>
				  	<th style="text-align: center;">Equip. Value</th>
				  	<th style="text-align: center;">Serial #</th>
				  	<th style="text-align: center;">Asset #</th>
				  </tr>
				</thead>
				<tbody>
				
				<%	
				TotalPurchaseCost = 0 
				
				SQLCustomerEquipment = " SELECT        EQ_Equipment.ModelIntRecID, MAX(EQ_Equipment.PurchaseCost) AS purchsum "
				SQLCustomerEquipment = SQLCustomerEquipment & " FROM            EQ_CustomerEquipment INNER JOIN "
				SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID INNER JOIN "
				SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
				SQLCustomerEquipment = SQLCustomerEquipment & " WHERE        (EQ_CustomerEquipment.CustID = " & CustIDPassed & ") AND (EQ_Models.ClassIntRecID = " & ClassIntRecID & ") "
				SQLCustomerEquipment = SQLCustomerEquipment & " GROUP BY EQ_Equipment.ModelIntRecID "
				SQLCustomerEquipment = SQLCustomerEquipment & " ORDER BY purchsum DESC "		
										
				Set cnnCustomerEquipment = Server.CreateObject("ADODB.Connection")
				cnnCustomerEquipment.open (Session("ClientCnnString"))
				Set rsCustomerEquipment = cnnCustomerEquipment.Execute(SQLCustomerEquipment)
				
				'***************************************************************************************
				'BUILD THE MASTER ORDER BY CLAUSE HERE
				'****************************************************************************************
				If Not rsCustomerEquipment.EOF Then
				
					EqpOrderByClauseCustom = " ORDER BY CASE ModelIntRecID "
					SortCount = 0
				
					Do While NOT rsCustomerEquipment.EOF
				
						EqpOrderByClauseCustom = EqpOrderByClauseCustom & " WHEN " & rsCustomerEquipment("ModelIntRecID") & " THEN " & Trim(SortCount) & " "
						SortCount = SortCount + 1
				
						rsCustomerEquipment.MoveNext
					Loop
					
					EqpOrderByClauseCustom = EqpOrderByClauseCustom & " END "
				
				End If
				
				'Response.write(EqpOrderByClauseCustom & "<br>")

				SQLCustomerEquipment = "SELECT * FROM EQ_CustomerEquipment INNER JOIN "
				SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
				SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN "
				SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
				SQLCustomerEquipment = SQLCustomerEquipment & " WHERE "
				SQLCustomerEquipment = SQLCustomerEquipment & " (EQ_CustomerEquipment.CustID = " & CustIDPassed & ") AND (EQ_Models.ClassIntRecID = " & ClassIntRecID & ") "
				SQLCustomerEquipment = SQLCustomerEquipment & EqpOrderByClauseCustom 
				
				'Response.write(SQLCustomerEquipment & "<br>")

				Set rsCustomerEquipment = cnnCustomerEquipment.Execute(SQLCustomerEquipment)

				If NOT rsCustomerEquipment.EOF Then
				
					FirstPassOnModel = True
					ModelLoopCounter = 1
					TotalRentalAmount = 0
					TotalPurchaseCost  = 0
				
					Do While NOT rsCustomerEquipment.EOF
					
						InstallDate = rsCustomerEquipment("InstallDate")
						StatusCodeIntRecID = rsCustomerEquipment("StatusCodeIntRecID")
						
						SQLEquipStatusCode = "SELECT * FROM EQ_StatusCodes WHERE InternalRecordIdentifier = " & StatusCodeIntRecID
							
						Set cnnEquipStatusCode = Server.CreateObject("ADODB.Connection")
						cnnEquipStatusCode.open (Session("ClientCnnString"))
						Set rsEquipStatusCode = cnnEquipStatusCode.Execute(SQLEquipStatusCode)
						
						If NOT rsEquipStatusCode.EOF Then
							InstallType = rsEquipStatusCode("statusBackendSystemCode")
							InstallTypeFullName = rsEquipStatusCode("statusDesc")
						Else
							InstallType = ""
							InstallTypeFullName = ""
						End If
												
						
						If InstallType = "R" then
						
							RentalFrequencyType = rsCustomerEquipment("RentalFrequencyType")
							
							Select Case RentalFrequencyType
							Case "D"
								RentalFrequencyFullName = "DAYS"
							Case "M"
								RentalFrequencyFullName = "MONTH(S)"
							Case "Y"
								RentalFrequencyFullName = "YEAR(S)"
							End Select
							
							RentalFrequencyNumber = rsCustomerEquipment("RentalFrequencyNumber")
							RentAmt = rsCustomerEquipment("RentAmt")
							
							If RentAmt <> "" Then
								TotalRentalAmount = TotalRentalAmount + RentAmt
								RentAmt = FormatCurrency(RentAmt,0)
							Else
								RentAmt = 0
								RentAmt = FormatCurrency(RentAmt,0)
							End If
							
						Else
							RentalFrequencyFullName = ""
							RentalFrequencyType = ""
							RentalFrequencyNumber = ""
							RentAmt = 0
							RentAmt = FormatCurrency(RentAmt,0)
						End If
												
						SerialNumber = rsCustomerEquipment("SerialNumber")
						PurchaseCost = rsCustomerEquipment("PurchaseCost")
						
						If PurchaseCost <> "" then
							TotalPurchaseCost = TotalPurchaseCost + PurchaseCost
							PurchaseCost = FormatCurrency(PurchaseCost,2)
						End If
						
						ModelIntRecID = rsCustomerEquipment("ModelIntRecID")
						
						If ModelIntRecID <> 0 Then
							BrandName = GetBrandNameByModelIntRecID(ModelIntRecID)
						Else
							BrandName = ""
						End If
						
						AssetTag1 = rsCustomerEquipment("AssetTag1")
						Description = "DESC NEEDED"
						Description  = GetModelNameByIntRecID(rsCustomerEquipment("ModelIntRecID"))
						
						ModelCount = GetTotalNumberOfModelsForCustomer(CustIDPassed,ModelIntRecID)
						
						%>
					
						<% If cInt(ModelCount) = 1 Then %>
						
							<tr>
								<td>&nbsp;</td>
								<% If BrandName <> "" Then %>
									<td><%= UCASE(BrandName) %>&nbsp;<%= Description %></td>
								<% Else %>
									<td><%= Description %></td>
								<% End If %>
								<td><%= InstallTypeFullName %></td>
								<% If InstallType = "R" Then %>
									<td><%= RentalFrequencyNumber %>&nbsp;<%= RentalFrequencyFullName %></td>
									<td align="center"><%= RentAmt %></td>
								<% Else %>
									<td>&nbsp;</td>
									<td align="center"><%= RentAmt %></td>
								<% End If %>
								<td align="right"><%= InstallDate %></td>
								<td align="right"><%= FormatCurrency(PurchaseCost,0) %></td>
								<td align="center"><%= SerialNumber %></td>
								<td align="center"><%= AssetTag1 %></td>
							</tr>
							
						<% ElseIf (cInt(ModelCount) > 1) AND (cInt(ModelLoopCounter) <= cInt(ModelCount)) Then %>
						
							<% If FirstPassOnModel = True Then %>
							
								<% ModelLoopCounter = 1 %>
															
								<tr class="accordion-toggle">
									<% If BrandName <> "" Then %>
										<td data-toggle="collapse" data-target=".equip<%= ModelIntRecID %>"><i class="fa fa-plus-circle fa-lg" aria-hidden="true" style="color:#009800"></i></td>
										<td colspan="3"><%= UCASE(BrandName) %>&nbsp;<%= Description %>&nbsp;<span class="equip_qty">(<%= ModelCount %>)</span></td>
										<td align="center"><%= FormatCurrency(GetTotalValueOfRentalModelsForCustomer(CustIDPassed,ModelIntRecID),0) %></td>
										<td>&nbsp;</td>
										<td align="right"><%= FormatCurrency(GetTotalValueOfModelsForCustomer(CustIDPassed,ModelIntRecID),0) %></td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
									<% Else %>
										<td data-toggle="collapse" data-target=".equip<%= ModelIntRecID %>" colspan="3"><%= Description %>&nbsp;(<%= ModelCount %>)&nbsp;<i class="fa fa-plus-circle" aria-hidden="true" style="color:#009800"></i></td>
										<td align="center"><%= FormatCurrency(GetTotalValueOfRentalModelsForCustomer(CustIDPassed,ModelIntRecID),0) %></td>
										<td>&nbsp;</td>
										<td align="right"><%= FormatCurrency(GetTotalValueOfModelsForCustomer(CustIDPassed,ModelIntRecID),0) %></td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>										
									<% End If %>
								</tr>		  
								<tr class="collapse equip<%= ModelIntRecID %>" style="background-color:#e5ffe5">
									<td>&nbsp;</td>
									<% If BrandName <> "" Then %>
										<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<%= UCASE(BrandName) %>&nbsp;<%= Description %></td>
									<% Else %>
										<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<%= Description %></td>
									<% End If %>
									<td><%= InstallTypeFullName %></td>
									<% If InstallType = "R" Then %>
										<td><%= RentalFrequencyNumber %>&nbsp;<%= RentalFrequencyFullName %></td>
										<td align="center"><%= RentAmt %></td>
									<% Else %>
										<td>&nbsp;</td>
										<td align="center"><%= RentAmt %></td>
									<% End If %>
									<td align="right"><%= InstallDate %></td>
									<td align="right"><%= FormatCurrency(PurchaseCost,0) %></td>
									<td align="center"><%= SerialNumber %></td>
									<td align="center"><%= AssetTag1 %></td>		  	
								</tr>
								
								<% FirstPassOnModel = False %>
								<% ModelLoopCounter = ModelLoopCounter + 1 %>
								
							<% Else %>
							
								<tr class="collapse equip<%= ModelIntRecID %>" style="background-color:#e5ffe5">
									<td>&nbsp;</td>
									<% If BrandName <> "" Then %>
										<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<%= UCASE(BrandName) %>&nbsp;<%= Description %></td>
									<% Else %>
										<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<%= Description %></td>
									<% End If %>
									<td><%= InstallTypeFullName %></td>
									<% If InstallType = "R" Then %>
										<td><%= RentalFrequencyNumber %>&nbsp;<%= RentalFrequencyFullName %></td>
										<td align="center"><%= RentAmt %></td>
									<% Else %>
										<td>&nbsp;</td>
										<td align="center"><%= RentAmt %></td>
									<% End If %>
									<td align="right"><%= InstallDate %></td>
									<td align="right"><%= FormatCurrency(PurchaseCost,0) %></td>
									<td align="center"><%= SerialNumber %></td>
									<td align="center"><%= AssetTag1 %></td>
								</tr>
								
								<% 
									ModelLoopCounter = ModelLoopCounter + 1
								
									If cInt(ModelLoopCounter) > cInt(ModelCount) Then
										FirstPassOnModel = True
										ModelLoopCounter = 1
									End If
								 %>
								
							<% End If %>
							
						<% End If %>
						<%
					
						rsCustomerEquipment.MoveNext
					
					Loop			
				End If
				
				%>
						  	
			</tbody>
			
			<tfoot>
			  <tr>
			  	<td colspan="2">TOTAL</td>
			  	<td>---</td>
			  	<td>---</td>
			  	<td align="center"><%= FormatCurrency(TotalRentalAmount,0) %></td>
			  	<td align="right">---</td>
			  	<td align="right"><%= FormatCurrency(TotalPurchaseCost,0) %></td>
			  	<td align="center">---</td>
			  	<td align="center">---</td>			  	
			  </tr>
			</tfoot>
		</table>
	<%
	
		rsCustomerEquipmentByClass.MoveNext
		Loop
		
		Set rsCustomerEquipment = Nothing
		cnnCustomerEquipment.Close
		Set cnnCustomerEquipment = Nothing
		
	End If

	Set rsEquipStatusCode = Nothing
	cnnEquipStatusCode.Close
	Set cnnEquipStatusCode = Nothing

	Set rsCustomerEquipmentByClass = Nothing
	cnnCustomerEquipmentByClass.Close
	Set cnnCustomerEquipmentByClass = Nothing


End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub SaveGeneralNotesGroupM() 

	CustIDPassed = Request.Form("CustID")
	GNGMuserno = Request.Form("GNGMuserno")
	GNGMMSCMonth = Request.Form("GNGMMSCMonth")
	groupm = Request.Form("groupm")
	GNGMMsg = Request.Form("GNGMMsg")
	GNGMsalesperson = Request.Form("GNGMsalesperson")
	GNGMinvoice_amount = Request.Form("GNGMinvoice_amount")
	GNGMchange_lvf = Request.Form("GNGMchange_lvf")
	GNGMchange_mcs = Request.Form("GNGMchange_mcs")
	GNGMReasons = Request.Form("GNGMReasons")
	lstSelectedUserIDs = Request.Form("lstSelectedUserIDs")
	SUMsg = Request.Form("SUMsg")

	GNGMReasons1 = 0

	Set rsGroupM = Server.CreateObject("ADODB.Recordset")
	rsGroupM.CursorLocation = 3 

	Set cnnGroupM = Server.CreateObject("ADODB.Connection")
	cnnGroupM.open (Session("ClientCnnString"))

	if groupm = "no_action_necessary" then
		ActionNote = "Action selected: No action necessary at this time"
		GNGMReasons1 = GNGMReasons
	Elseif groupm = "remove_client" then
		ActionNote = "Action selected: Remove this client from the MCS Program."
	Elseif groupm = "send_invoice" then
		ActionNote = "Action selected: Send invoice to client for the amount of $" & GNGMinvoice_amount
	Elseif groupm = "notify_selected_sales_person" then
		SQLGroupM = "SELECT userEmail, userFirstName, userLastName FROM tblUsers WHERE userSalesPersonNumber='" & GNGMsalesperson & "' OR userSalesPersonNumber2='" & GNGMsalesperson & "'"
		'response.write SQLGroupM

		set rsGroupM = cnnGroupM.Execute(SQLGroupM)
		If NOT rsGroupM.EOF Then 
			userEmail = rsGroupM("userEmail")
			userName = userName & sep & rsGroupM("userFirstName")
			userNameForAud = userNameForAud & sep & rsGroupM("userFirstName") & " " & rsGroupM("userLastName")
		Else
			' Otherwise, is if they are in the salesperson file
			response.write "Error: Not able to find Sales Person " & GNGMsalesperson & vbcrlf
			exit sub
		End If 
		rsGroupM.Close	
		SQLGroupM = "SELECT userEmail, userFirstName, userLastName FROM tblUsers WHERE UserNo IN ( " & lstSelectedUserIDs & " )"
		'response.write SQLGroupM

		set rsGroupM = cnnGroupM.Execute(SQLGroupM)
		ExtraEmails = ""
		ExtrauserName = ""
		If NOT rsGroupM.EOF Then 
			sep = ""
			Do While NOT rsGroupM.EOF 				
				if userEmail = rsGroupM("userEmail") Then 
					'userEmail = rsGroupM("userEmail")
				Else
					ExtraEmails = ExtraEmails & sep & rsGroupM("userEmail")
					ExtrauserName = ExtrauserName & rsGroupM("userFirstName") & " " & rsGroupM("userLastName")
				End If

				sep = ","
				rsGroupM.MoveNext
			Loop
		Else
			response.write "Error: Not able to find Users" & vbcrlf
			exit sub
		End If 
		rsGroupM.Close			
		ActionNote = "Action selected: Notify " & userName & " (sales person) to followup on the MCS shortage" & GNGMinvoice_amount & vbcrlf & "Also sent to: " & ExtrauserName & " Message" & SUMsg
		MCSSubject = CustIDPassed & "  " & GetCustNameByCustNum(CustIDPassed) & " Followup on the MCS shortage"
	Elseif groupm = "change_mcs" then
		SQLGroupM = "SELECT MonthlyContractedSalesDollars FROM AR_Customer WHERE CustNum = '" & CustIDPassed & "'"
		set rsGroupM = cnnGroupM.Execute(SQLGroupM)
		If NOT rsGroupM.EOF Then 
			MonthlyContractedSalesDollars = rsGroupM("MonthlyContractedSalesDollars")
		Else
			response.write "Error: Not able to find MCS Amount for Client " & CustIDPassed & vbcrlf
		End If
		rsGroupM.Close	
		If Not IsNumeric(MonthlyContractedSalesDollars) Then MonthlyContractedSalesDollars = 0
		ActionNote = "Changed MCS from " & FormatCurrency(MonthlyContractedSalesDollars) & " to " & FormatCurrency(GNGMchange_mcs)
		SQLGroupM = "UPDATE AR_Customer SET MonthlyContractedSalesDollars ='" & GNGMchange_mcs & "' WHERE CustNum = '" & CustIDPassed & "'"
		dummy = MUV_WRITE("MCSFLAG","1")
		set rsGroupM = cnnGroupM.Execute(SQLGroupM)
	Elseif groupm = "change_lvf" then
		SQLGroupM = "SELECT MaxMCSCharge FROM AR_Customer WHERE CustNum = '" & CustIDPassed & "'"
		set rsGroupM = cnnGroupM.Execute(SQLGroupM)
		If NOT rsGroupM.EOF Then 
			MaxMCSCharge = rsGroupM("MaxMCSCharge")
		Else
			response.write "Error: Not able to find Max MCS fee for Client " & CustIDPassed & vbcrlf
		End If	
		rsGroupM.Close	
		If Not IsNumeric(MaxMCSCharge) Then MaxMCSCharge = 0
		ActionNote = "Changed maximum LVF from " & FormatCurrency(MaxMCSCharge) & " to " & FormatCurrency(GNGMchange_lvf)
		SQLGroupM = "UPDATE AR_Customer SET MaxMCSCharge ='" & GNGMchange_lvf & "' WHERE CUstNum = '" & CustIDPassed & "'"
		set rsGroupM = cnnGroupM.Execute(SQLGroupM)
	Elseif groupm = "send_message_to_someone" then
		SQLGroupM = "SELECT userEmail, userFirstName, userLastName, userDisplayName FROM tblUsers WHERE UserNo IN ( " & lstSelectedUserIDs & " )"
		'response.write SQLGroupM

		set rsGroupM = cnnGroupM.Execute(SQLGroupM)
		i = 0
		ExtraEmails = ""
		userName = ""
		userNameForAud = ""
		If NOT rsGroupM.EOF Then 
			sep = ""
			Do While NOT rsGroupM.EOF 
				userName = userName & sep & rsGroupM("userFirstName")
				userNameForAud = userNameForAud & sep & rsGroupM("userFirstName") & " " & rsGroupM("userLastName")
				if i = 0 Then 
					userEmail = rsGroupM("userEmail")
				else
					ExtraEmails = ExtraEmails & sep & rsGroupM("userEmail")
				End If
				i = 1				
				sep = ", "
				rsGroupM.MoveNext
			Loop
		Else
			response.write "Error: Not able to find Users" & vbcrlf
			exit sub
		End If 
		rsGroupM.Close	
		ActionNote = "Sent message to " & userNameForAud & ". Message: " & SUMsg
		MCSSubject = "MCS Message - " &  CustIDPassed & "  " & GetCustNameByCustNum(CustIDPassed)
	End if

	SQLMCSActions = "INSERT INTO BI_MCSActions (RecordCreationDateTime, CustID, MCSMonth, Action, ActionNotes, MCSReasonIntRecID) VALUES (GetDate(), '" & CustIDPassed & "','" & GNGMMSCMonth & "','" & groupm & "','" & trim(replace(ActionNote&"","'","''")) & "'," & GNGMReasons1 & ")"
	set rsGroupM = cnnGroupM.Execute(SQLMCSActions)
	SQLMCSActions = "INSERT INTO AR_CustomerNotes (RecordCreationDateTime, CustID, Category, EnteredByUserNo, Note, NoteType, MCSReasonIntRecID, NoteTypeIntRecID) VALUES (GetDate(), '" & CustIDPassed & "','" & "-2" & "','" & GNGMuserno & "','" & trim(replace(ActionNote&"","'","''"))  & "','MCS', " & GNGMReasons1 & ",4)"  
	set rsGroupM = cnnGroupM.Execute(SQLMCSActions)
	
	email_mcsvar = 0
	email_salesvar = 0
	
	SQLMCSActions = "SELECT Month3Sales_NoRent FROM BI_MCSData WHERE CustID = '" & CustIDPassed & "'"
	set rsGroupM = cnnGroupM.Execute(SQLMCSActions)
	email_salesvar = FormatCurrency(rsGroupM("Month3Sales_NoRent"),0)
	
	SQLMCSActions = "SELECT MonthlyContractedSalesDollars FROM AR_Customer WHERE CustNum = '" & CustIDPassed & "'"
	set rsGroupM = cnnGroupM.Execute(SQLMCSActions)
	email_mcsvar = FormatCurrency(rsGroupM("MonthlyContractedSalesDollars"),0)
	
	set rsGroupM = Nothing
	cnnGroupM.Close
	set cnnGroupM = Nothing
	
	if groupm = "notify_selected_sales_person" or groupm = "send_message_to_someone" then
	
			
		GNGMsalePersonEmail = userEmail		
		MCSEmailBody = "Hi " & userName & "<br><br>"
		MCSEmailBody = MCSEmailBody  & "<table>"
		MCSEmailBody = MCSEmailBody  & "<tr><td>MCS:" & "</td><td align='right'>" & email_mcsvar & "</td></tr>"
		MCSEmailBody = MCSEmailBody  & "<tr><td>Sales:" & "</td><td align='right'>" & email_salesvar & "</td></tr>"
		If email_salesvar-email_mcsvar > 0 Then
			MCSEmailBody = MCSEmailBody  & "<tr><td>Variance:&nbsp;" & "</td><td align='right'>" & FormatCurrency(email_salesvar-email_mcsvar,0) & "</td></tr>"
		Else
			MCSEmailBody = MCSEmailBody  & "<tr><td>Variance:&nbsp;" & "</td><td align='right'><font color='red'>" & FormatCurrency(email_salesvar-email_mcsvar,0) & "</font></td></tr>"		
		End If
		MCSEmailBody = MCSEmailBody  & "</table>" & "<br><br>"
		MCSEmailBody = MCSEmailBody  & SUMsg & "<br>"
		
		emailTo = userEmail
		If Left(ExtraEmails,2) = ", " Then ExtraEmails = Right(ExtraEmails,len(ExtraEmails)-2)
		emailCCs = ExtraEmails
		emailBCCs = ""
		
	'Response.Write("emailTo:" & emailto)
	'Response.Write("emailCCs:" & emailCCs)
		
	SendMailWithCCs emailFrom,emailTo,MCSSubject,MCSEmailBody,emailCCs,emailBCCs,"MCS","MCS"


	End if
	
	response.write "Changes Saved"
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub DeleteMCSClientbyCustID() 
	
	CustIDPassed = Request("CustID")

	if (CustIDPassed <> "") then
		Set rsMCS = Server.CreateObject("ADODB.Recordset")
		rsMCS.CursorLocation = 3 

		Set cnnMCS = Server.CreateObject("ADODB.Connection")
		cnnMCS.open (Session("ClientCnnString"))

		SQLMCS = "SELECT Name, MonthlyContractedSalesDollars FROM AR_Customer WHERE CustNum=" & CustIDPassed
	
		set rsMCS = cnnMCS.Execute(SQLMCS)
		If NOT rsMCS.EOF Then 
			Old_MonthlyContractedSalesDollars = rsMCS("MonthlyContractedSalesDollars")
			rsMCS.Close
			SQLRemoveCust = "UPDATE AR_Customer SET MonthlyContractedSalesDollars=0 WHERE CustNum=" & CustIDPassed
			dummy = MUV_WRITE("MCSFLAG","1")
			set rsMCS = cnnMCS.Execute(SQLRemoveCust)
			MCS = "MCS Client Removed"
			MSCMonth = ""
			ActionNote = "Client Removed from MCS Program. Prior to removal the MCS was $" & Old_MonthlyContractedSalesDollars
			SQLMCSActions = "INSERT INTO BI_MCSActions (RecordCreationDateTime, CustID, MCSMonth, Action, ActionNotes) VALUES "
			SQLMCSActions = SQLMCSActions & "(GetDate(), '" & CustIDPassed & "','" & Monthname(Month(Now() -1)) & "','" & MCS & "','" & trim(replace(ActionNote&"","'","''")) & "')"
			set rsMCS = cnnMCS.Execute(SQLMCSActions)				
			
			SQLMCSActions = "INSERT INTO AR_CustomerNotes (CustID, Category, EnteredByUserNo, Note, NoteType) VALUES "
			SQLMCSActions = SQLMCSActions & "('" & CustIDPassed & "','" & "-2" & "','" & GNGMuserno & "','" & trim(replace(ActionNote&"","'","''"))  & "','MCS')"  
			set rsMCS = cnnMCS.Execute(SQLMCSActions)

			
			
			response.write "Client " & CustIDPassed & " is Removed"
		Else
			rsMCS.Close
			response.write "Error: Client with Customer ID " & CustIDPassed & " is not found in database " & vbcrlf
		End If
		
		set rsMCS = Nothing
		cnnMCS.Close
		set cnnMCS = Nothing	
	else 
		response.write "Error: No Customer ID passed"
	end if
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub AddMCSClientbyCustID() 
	
	CustIDPassed = Request("CustID")
	MCSDollars = Request("MCSDollars")
	MaxMCSCharge = Request("MaxMCSCharge")
	MSCMonth = ""

	if (CustIDPassed <> "") then
		Set rsGroupM = Server.CreateObject("ADODB.Recordset")
		rsGroupM.CursorLocation = 3 

		Set cnnGroupM = Server.CreateObject("ADODB.Connection")
		cnnGroupM.open (Session("ClientCnnString"))

		SQLGroupM = "SELECT Name, MonthlyContractedSalesDollars FROM AR_Customer WHERE CustNum=" & CustIDPassed
	
		set rsGroupM = cnnGroupM.Execute(SQLGroupM)
		If NOT rsGroupM.EOF Then 
			if rsGroupM("MonthlyContractedSalesDollars") > 0 then
				response.write "Error: Client " & CustIDPassed & " already exists in MCS Program"
			else 
				rsGroupM.Close
				If NOT IsNumeric(MaxMCSCharge) Then MaxMCSCharge = 0
				SQLRemoveCust = "UPDATE AR_Customer SET MonthlyContractedSalesDollars=" & MCSDollars & ",MaxMCSCharge=" & MaxMCSCharge & ", MCSEnrollmentDate = getdate() WHERE CustNum=" & CustIDPassed
				dummy = MUV_WRITE("MCSFLAG","1")
				set rsGroupM = cnnGroupM.Execute(SQLRemoveCust)
				groupm = "MCS Client Added"
				ActionNote = "Client added to MCS Program. The MCS amount was set to " & FormatCurrency(MCSDollars,0) & ". The max LVF was set to " & FormatCurrency(MaxMCSCharge ,0) & "."
				SQLMCSActions = "INSERT INTO BI_MCSActions (RecordCreationDateTime, CustID, MCSMonth, Action, ActionNotes) VALUES (GetDate(), '" & CustIDPassed & "','" & Monthname(Month(Now() - 1)) & "','" & groupm & "','" & trim(replace(ActionNote&"","'","''")) & "')"
				set rsGroupM = cnnGroupM.Execute(SQLMCSActions)		

				SQLMCSActions = "INSERT INTO AR_CustomerNotes (RecordCreationDateTime, CustID, Category, EnteredByUserNo, Note, NoteType) VALUES "
				SQLMCSActions = SQLMCSActions & "(GetDate(), '" & CustIDPassed & "','" & "-2" & "','" & GNGMuserno & "','" & trim(replace(ActionNote&"","'","''"))  & "','MCS')"  
				set rsGroupM = cnnGroupM.Execute(SQLMCSActions)
					
		
				response.write "Client " & CustIDPassed & " is added to MCS."
			end if
		Else
			rsGroupM.Close
			response.write "Error: Client with Customer ID " & CustIDPassed & " is not found in database " & vbcrlf
		End If 
		
		set rsGroupM = Nothing
		cnnGroupM.Close
		set cnnGroupM = Nothing	
	else 
		response.write "Error: No Customer ID passed"
	end if
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetTRofNewMCSClientbyCustID() 
	
	CustIDPassed = Request("CustID")

	if (CustIDPassed <> "") then
		Set rsGroupM = Server.CreateObject("ADODB.Recordset")
		rsGroupM.CursorLocation = 3 

		Set cnnGroupM = Server.CreateObject("ADODB.Connection")
		cnnGroupM.open (Session("ClientCnnString"))


	SQL = "SELECT * FROM AR_Customer WHERE CustNum=" & CustIDPassed 
	
	
	Set rsGroupM = cnnGroupM.Execute(SQL)

	If Not rsGroupM.Eof Then
			
		Do While Not rsGroupM.EOF

			ShowThisRecord = True

				
			If ShowThisRecord <> False Then			
			
				PrimarySalesMan =  ""
				SecondarySalesMan =  ""
				ReferralCode =  ""
				CustomerType =  ""
				SelectedCustomerID = rsGroupM("CustNum")
				CustName = rsGroupM("Name")
				CustMonthlyContractedSalesDollars = 0
				InstallDate = ""

				PrimarySalesMan = rsGroupM("Salesman")
				SecondarySalesMan = rsGroupM("SecondarySalesman")
				ReferralCode = rsGroupM("ReferalCode")
				CustomerType = rsGroupM("CustType")
				CustMonthlyContractedSalesDollars = rsGroupM("MonthlyContractedSalesDollars")
				InstallDate = rsGroupM("InstallDate")
				MaxMCSCharge = rsGroupM("MaxMCSCharge")
				
				'Decide if this record meets the filter criteria
				If FilterSlsmn1 <> "" And FilterSlsmn1 <> "All" Then
					If CInt(FilterSlsmn1) <> Cint(rsGroupM("Salesman")) Then ShowThisRecord = False
				End If
				If FilterSlsmn2 <> "" And FilterSlsmn2 <> "All" Then
					If CInt(FilterSlsmn2) <> Cint(rsGroupM("SecondarySalesman")) Then ShowThisRecord = False
				End If
		
			End If
			

			Month3Sales_NoRent = TotalSalesByCustByMonthByYear_NoRentals(SelectedCustomerID,Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))	

			If ShowAllCusts <> 1 Then
				If Month3Sales_NoRent >= CustMonthlyContractedSalesDollars Then ShowThisRecord = False
			End If

			
			If ShowThisRecord <> False Then
			
				Month1Sales_NoRent = TotalSalesByCustByMonthByYear_NoRentals(SelectedCustomerID,Month(DateAdd("m",-3,ReportDate)),Year(DateAdd("m",-3,ReportDate)))
				Month2Sales_NoRent = TotalSalesByCustByMonthByYear_NoRentals(SelectedCustomerID,Month(DateAdd("m",-2,ReportDate)),Year(DateAdd("m",-2,ReportDate)))
				ThreePPSales = Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent
				
				Month3Cost_NoRent = TotalCostByCustByMonthByYear_NoRent(SelectedCustomerID,Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
				
				Month3GP = Month3Sales_NoRent - Month3Cost_NoRent
				If Not IsNumeric(Month3GP) Then Month3GP  = 0
			
				ThreePPAvgSales = (Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent) / 3
				
				ShortageHolder = ThreePPSales - (CustMonthlyContractedSalesDollars * 3)
				VarianceHolder = Month3Sales_NoRent - CustMonthlyContractedSalesDollars 
				LVFHolder = TotalPostedLVFByCustByMonthByYear(SelectedCustomerID,Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
				LVFHolderCurrent = TotalPostedLVFByCustByMonthByYear(SelectedCustomerID,Month(ReportDate),Year(ReportDate))
				
				If Month3Sales_NoRent <> 0 Then
					VariancePercentHolder = 100 - ((Month3Sales_NoRent/CustMonthlyContractedSalesDollars) * 100) 
				Else
					VariancePercentHolder = 100
				End If
				VariancePercentHolder  = VariancePercentHolder  * -1
				
				TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(SelectedCustomerID)
				
				TotalCustsReported = TotalCustsReported + 1

				Response.Write("<tr id=""CUST" & SelectedCustomerID & """>")
				'Response.Write("<input type='hidden' id='txtCustIDToPass' name='txtCustIDToPass' value='" & SelectedCustomerID & "'>")
			    Response.Write("<td class='smaller-detail-line'><a href='../CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>"& SelectedCustomerID  & "</a></td>")
			    Response.Write("<td class='smaller-detail-line'><a href='../CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>"& CustName & "</a></td>")	
			    PrimarySalesPerson = GetSalesmanNameBySlsmnSequence(PrimarySalesMan)
			    SecondarySalesPerson = GetSalesmanNameBySlsmnSequence(SecondarySalesman)
			    If Instr(PrimarySalesPerson ," ") <> 0 Then
					Response.Write("<td class='smaller-detail-line'>" & Left(PrimarySalesPerson,Instr(PrimarySalesPerson ," ")+1) & "</td>")
				Else
					Response.Write("<td class='smaller-detail-line'>" & PrimarySalesPerson & "</td>")
				End If
				If Instr(SecondarySalesPerson," ") <> 0 Then
					Response.Write("<td class='smaller-detail-line'>" & Left(SecondarySalesPerson,Instr(SecondarySalesPerson," ")+1) & "</td>")
				Else
					Response.Write("<td class='smaller-detail-line'>" & SecondarySalesPerson & "</td>")
				End If
				
				InstallDate = cDate(InstallDate) 
				iYear = Year(InstallDate)
				If Month(InstallDate) < 10 Then iMonth = "0" & Month(InstallDate) else iMonth = Month(InstallDate)
				If Day(InstallDate) < 10 Then iDay = "0" & Day(InstallDate) else iDay = Day(InstallDate)
				Response.Write("<td align='right' class='smaller-detail-line'><span class='hidden'>" & iYear & iMonth & iDay & "</span>" & InstallDate & "</td>")
				Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(Month1Sales_NoRent,0) & "</td>")
				Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(Month2Sales_NoRent,0) & "</td>")
				Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(Month3Sales_NoRent,0) & "</td>")

				
				'Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(Month3Sales_NoRent,0) & "</td>")
				Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ThreePPAvgSales,0) & "</td>")
				
			    If ShortageHolder < 0 Then
			    	Response.Write("<td align='right' class='negative-thin smaller-detail-line'>" & FormatCurrency(ShortageHolder ,0) & "</td>")
			    Else
			    	Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ShortageHolder ,0) & "</td>")				    
			    End If
			    
			    CurrentHolder = 0
			    Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(CurrentHolder,0) & "</td>")
			    
				Response.Write("<td align='right' class='not-as-small-detail-line' style='border-left: 2px solid #555 !important;'>" & FormatCurrency(CustMonthlyContractedSalesDollars,0) & "</td>")

				If VarianceHolder < 0 Then 
					Response.Write("<td align='right' class='negative-thin not-as-small-detail-line' style='border-right: 2px solid #555 !important;'>" & FormatCurrency(VarianceHolder,0,0,0) & "</td>")
				Else
					Response.Write("<td align='right' class='not-as-small-detail-line' style='border-right: 2px solid #555 !important;'>" & FormatCurrency(VarianceHolder,0,0,0) & "</td>")
				End If

				RentalHolder = TotalSalesByCustByMonthByYear_RentalsOnly(SelectedCustomerID,Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
				If RentalHolder < 0 Then
					Response.Write("<td align='right' class='negative-thin smaller-detail-line'>" & FormatCurrency(RentalHolder ,0) & "</td>")
				Else
					Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(RentalHolder ,0) & "</td>")				
				End If
				
				Response.Write("<td align='right' class='smaller-detail-line'>" &  FormatCurrency(LVFHolder,2)  & "</td>")					

				
				Response.Write("<td align='right' class='smaller-detail-line'>" &  FormatCurrency(Month3GP,0)  & "</td>")	
				
				If TotalEquipmentValue > 0 Then	
				
					LCPGP = 0 
					
				    If TotalEquipmentValue <> 0 Then %>
				    	<td align='right' class='smaller-detail-line'>
				    	<a data-toggle="modal" data-show="true" href="#" data-cust-id="<%= SelectedCustomerID %>" data-lcp-gp="<%= LCPGP %>" data-target="#modalEquipmentVPC" data-tooltip="true" data-title="View Customer Equipment"><%= FormatCurrency(TotalEquipmentValue,0) %></a>    
				    	</td>
					<% Else %>
						<%= FormatCurrency(TotalEquipmentValue,0) %>
					<% End If %>
					
				<%
				Else
					Response.Write("<td align='right' class='smaller-detail-line'>No Equipment</td>")
				End If

				'Action
				Response.Write("<td align='right' class='smaller-detail-line'>")
				'Response.Write("<select name='selActions' id='selActions'>")
				'Response.Write("<option value='None'>None</option>")
				'Response.Write("<option value='Follow Up'>Follow Up</option>")
				'Response.Write("<option value='Bill'>Bill</option>")
				'Response.Write("<option value='Remove'>Remove</option>")
				'Response.Write("</select>")
				btncolor = "btn-default"
				if GetMCSNotesStatus(SelectedCustomerID, MonthName(Month(ReportDate)-1)) Then 
					btncolor = "btn-success"
				End if 
				Response.Write "<button type=""button"" class=""" & btncolor & """ id=""btn" & SelectedCustomerID & """ data-toggle=""modal"" data-target=""#modalGeneralNotesGroupM"" data-cust-id=""" & SelectedCustomerID & """ data-cust-name=""" &CustName & """ data-mcs-variance=""" & VarianceHolder & """ data-mcs-salespersonid1=""" & PrimarySalesMan & """ data-mcs-salespersonid2=""" & SecondarySalesMan & """  data-mcs-salesperson1=""" & PrimarySalesPerson & """ data-mcs-salesperson2=""" & SecondarySalesPerson & """ data-mcs-month=""" & MonthName(Month(ReportDate)-1) & """ data-mcs-userno=""" & Session("userNo") & """ data-maxmcscharge=""" & MaxMCSCharge & """ >Action</button>"
				Response.Write("</td>")
				
				'Additional Info / Notes
				
				'Allow for a note here as a way to put in a note for the customer in general
				'Use -2 as the category number for MCS notes
				If CustHasMCSNotes(SelectedCustomerID) = True Then
					FirstName = GetUserFirstAndLastNameByUserNo(GetMostRecentMCSNoteUserNo(SelectedCustomerID))
					FirstName = Left(FirstName ,InStr(FirstName," "))
					FirstName = Trim(FirstName)
					MostRecentNoteText = FirstName & ": " & GetMostRecentMCSNote(SelectedCustomerID)
					If NoteNewCatAnalForUser(SelectedCustomerID ,-1) = True Then
						'Pulsing icon
						Response.Write("<td align='center' class='smaller-detail-line'>")
						Response.Write("<a data-toggle='modal' data-target='#modalEditCustomerNotes' data-category-id='-2' data-cust-id='" & SelectedCustomerID & "' class='ole' rel='tooltip' data-original-title='" & MostRecentNoteText  & "' style='cursor:pointer;'><i class='fa fa-file-text-o faa-pulse animated fa-2x' aria-hidden='true'></i></a>")																	
						Response.Write("</td>")
					Else
						'Regular icon
						Response.Write("<td align='center' class='smaller-detail-line'>")
						Response.Write("<a data-toggle='modal' data-target='#modalEditCustomerNotes' data-category-id='-2' data-cust-id='" & SelectedCustomerID  & "' class='ole' rel='tooltip' data-original-title='" & MostRecentNoteText  & "' style='cursor:pointer;'><i class='fa fa-file-text-o' aria-hidden='true'></i></a>")											
						Response.Write("</td>")
					End If
				Else
					'Pencil icon
					Response.Write("<td align='center' class='smaller-detail-line'>")
					Response.Write("<a data-toggle='modal' data-target='#modalEditCustomerNotes' data-category-id='-2' data-cust-id='" & SelectedCustomerID  & "' class='ole' rel='tooltip' data-original-title='Click to create a new note.' style='cursor:pointer;'><i class='fa fa-pencil' aria-hidden='true'></i></a>")																
					Response.Write("</td>")
				End If


			    Response.Write("</tr>")
			    
		    End If			
			
			rsGroupM.movenext
				
		Loop	
	end if
		rsGroupM.Close
		
		set rsGroupM = Nothing
		cnnGroupM.Close
		set cnnGroupM = Nothing	
	else 
		response.write "Error: No Customer ID passed"
	end if
	
End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Function getSelectUsers()

	SelectedUsers = Request("SelectedUsers")
	
	Set cnnUserList = Server.CreateObject("ADODB.Connection")
	cnnUserList.open Session("ClientCnnString")

	SQLUserList = "SELECT UserNo,userFirstName,userLastName FROM tblUsers"
	if SelectedUsers <> "" Then
		SQLUserList = SQLUserList & " WHERE tblUsers.UserNo NOT IN (" & SelectedUsers & ")"
	End If
	SQLUserList = SQLUserList & " ORDER BY userFirstName,userLastName"
	
	Set rsUserList = Server.CreateObject("ADODB.Recordset")
	rsUserList.CursorLocation = 3 
	Set rsUserList = cnnUserList.Execute(SQLUserList)

	response.write "["
	If Not rsUserList.EOF Then
		sep = ""
		Do While Not rsUserList.EOF
			Response.Write(sep)
			sep = ","
			FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName")
			Response.Write("{")
			Response.Write("""UserNo"":""" & EscapeQuotes(rsUserList("UserNo")) & """")
			Response.Write(",""FullName"":""" & EscapeQuotes(FullName) & """")			
			Response.Write("}")
			rsUserList.MoveNext
		Loop
	End If
	Response.Write("]")
	
	Set rsUserList = Nothing
	cnnUserList.Close
	Set cnnUserList = Nothing


End Function
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Function getCCUsers()

	SelectedUsers = Request("SelectedUsers")
	
	Set cnnUserList = Server.CreateObject("ADODB.Connection")
	cnnUserList.open Session("ClientCnnString")

	SQLCCList = "SELECT MCSUserNosToCC FROM Settings_BizIntel"
	
	Set rsUserList = Server.CreateObject("ADODB.Recordset")
	rsUserList.CursorLocation = 3 

	MCSUserNosToCC = ""
	Set rsUserList = cnnUserList.Execute(SQLCCList)	
	If Not rsUserList.EOF Then
		MCSUserNosToCC = rsUserList("MCSUserNosToCC")
	END If
	if SelectedUsers <> "" Then
		MCSUserNosToCC = MCSUserNosToCC & "," & SelectedUsers
	End If
	rsUserList.close()
	
	SQLUserList = "SELECT UserNo,userFirstName,userLastName FROM tblUsers"
	if MCSUserNosToCC <> "" Then
		SQLUserList = SQLUserList & " WHERE tblUsers.UserNo IN (" & MCSUserNosToCC & ")"
	End If
	SQLUserList = SQLUserList & " ORDER BY userFirstName,userLastName"	
	
	Set rsUserList = cnnUserList.Execute(SQLUserList)
	
	response.write "["
	If Not rsUserList.EOF Then
		sep = ""
		Do While Not rsUserList.EOF
			Response.Write(sep)
			sep = ","
			FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName")
			Response.Write("{")
			Response.Write("""UserNo"":""" & EscapeQuotes(rsUserList("UserNo")) & """")
			Response.Write(",""FullName"":""" & EscapeQuotes(FullName) & """")			
			Response.Write("}")
			rsUserList.MoveNext
		Loop
	End If
	Response.Write("]")
	
	Set rsUserList = Nothing
	cnnUserList.Close
	Set cnnUserList = Nothing

End Function
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Function EscapeQuotes(val)
	EscapeQuotes = Replace(val, """", "\""")
End Function

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Function getUsersBySalesperson()

	SalesPersons = Request("SalesPersons")
	
	If SalesPersons = "" Then
		Response.write "Error: SalesPerson No. is not provided."
		Exit Function
	End If
	
	Set cnnUserList = Server.CreateObject("ADODB.Connection")
	cnnUserList.open Session("ClientCnnString")

	Set rsUserList = Server.CreateObject("ADODB.Recordset")
	rsUserList.CursorLocation = 3 

	SQLUserList = "SELECT UserNo,userFirstName,userLastName FROM tblUsers"
	SQLUserList = SQLUserList & " WHERE tblUsers.userSalesPersonNumber IN (" & SalesPersons & ") OR tblUsers.userSalesPersonNumber2 IN (" & SalesPersons & ")"
	SQLUserList = SQLUserList & " ORDER BY userFirstName,userLastName"	
	Set rsUserList = cnnUserList.Execute(SQLUserList)
	

	response.write "["
	If Not rsUserList.EOF Then
		sep = ""
		Do While Not rsUserList.EOF
			Response.Write(sep)
			sep = ","
			FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName")
			Response.Write("{")
			Response.Write("""UserNo"":""" & EscapeQuotes(rsUserList("UserNo")) & """")
			Response.Write(",""FullName"":""" & EscapeQuotes(FullName) & """")			
			Response.Write("}")
			rsUserList.MoveNext
		Loop
	End If
	Response.Write("]")
	
	Set rsUserList = Nothing
	cnnUserList.Close
	Set cnnUserList = Nothing

End Function
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub DeleteMESClientbyCustID() 
	
	CustIDPassed = Request("CustID")

	if (CustIDPassed <> "") then
		Set rsMES = Server.CreateObject("ADODB.Recordset")
		rsMES.CursorLocation = 3 

		Set cnnMES = Server.CreateObject("ADODB.Connection")
		cnnMES.open (Session("ClientCnnString"))

		SQLMES = "SELECT Name FROM AR_Customer WHERE CustNum=" & CustIDPassed
	
		set rsMES = cnnMES.Execute(SQLMES)
		If NOT rsMES.EOF Then 
			rsMES.Close
			SQLRemoveCust = "UPDATE AR_Customer SET MonthlyExpectedSalesDollars=0 WHERE CustNum=" & CustIDPassed
			dummy = MUV_WRITE("MESFLAG","1")
			set rsMES = cnnMES.Execute(SQLRemoveCust)
			MES = "MES Client Removed"
			MSCMonth = ""
			ActionNote = "Client Removed from MES program."
			SQLMCSActions = "INSERT INTO BI_MESActions (RecordCreationDateTime, CustID, MESMonth, Action, ActionNotes) VALUES "
			SQLMCSActions = SQLMCSActions & "(GetDate(), '" & CustIDPassed & "','" & Monthname(Month(Now() -1)) & "','" & MES & "','" & trim(replace(ActionNote&"","'","''")) & "')"
			set rsMES = cnnMES.Execute(SQLMCSActions)				
			
			SQLMCSActions = "INSERT INTO AR_CustomerNotes (CustID, Category, EnteredByUserNo, Note, NoteType) VALUES "
			SQLMCSActions = SQLMCSActions & "('" & CustIDPassed & "','" & "-2" & "','" & GNGMuserno & "','" & trim(replace(ActionNote&"","'","''"))  & "','MES')"  
			set rsMES = cnnMES.Execute(SQLMCSActions)

			
			
			response.write "Client " & CustIDPassed & " removed"
		Else
			rsMES.Close
			response.write "Error: Client with Customer ID " & CustIDPassed & " is not found in database " & vbcrlf
		End If
		
		set rsMES = Nothing
		cnnMES.Close
		set cnnMES = Nothing	
	else 
		response.write "Error: No Customer ID passed"
	end if
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub AddMESClientbyCustID() 
	
	CustIDPassed = Request("CustID")
	MESDollars = Request("MESDollars")
	MaxMCSCharge = Request("MaxMCSCharge")
	MECMonth = ""

	if (CustIDPassed <> "") then
		Set rsAddMES = Server.CreateObject("ADODB.Recordset")
		rsAddMES.CursorLocation = 3 

		Set cnnAddMES = Server.CreateObject("ADODB.Connection")
		cnnAddMES.open (Session("ClientCnnString"))

		SQLAddMES = "SELECT Name, MonthlyExpectedSalesDollars FROM AR_Customer WHERE CustNum=" & CustIDPassed
	
		set rsAddMES = cnnAddMES.Execute(SQLAddMES)
		If NOT rsAddMES.EOF Then 
			if rsAddMES("MonthlyExpectedSalesDollars") > 0 then
				response.write "Error: Client " & CustIDPassed & " already exists in MES program"
			else 
				rsAddMES.Close
				If NOT IsNumeric(MaxMCSCharge) Then MaxMCSCharge = 0
				SQLRemoveCust = "UPDATE AR_Customer SET MonthlyExpectedSalesDollars =" & MESDollars & ", MESEnrollmentDate = getdate() WHERE CustNum=" & CustIDPassed
				dummy = MUV_WRITE("MESFLAG","1")
				set rsAddMES = cnnAddMES.Execute(SQLRemoveCust)
				AddMES = "MES Client Added"
				ActionNote = "Client added to MES Program. The MES amount was set to " & FormatCurrency(MESDollars,0) & "."
				SQLMESActions = "INSERT INTO BI_MESActions (RecordCreationDateTime, CustID, MESMonth, Action, ActionNotes) VALUES (GetDate(), '" & CustIDPassed & "','" & Monthname(Month(Now() - 1)) & "','" & AddMES & "','" & trim(replace(ActionNote&"","'","''")) & "')"
				set rsAddMES = cnnAddMES.Execute(SQLMESActions)		

				SQLMESActions = "INSERT INTO AR_CustomerNotes (RecordCreationDateTime, CustID, Category, EnteredByUserNo, Note, NoteType) VALUES "
				SQLMESActions = SQLMESActions & "(GetDate(), '" & CustIDPassed & "','" & "-2" & "','" & GNGMuserno & "','" & trim(replace(ActionNote&"","'","''"))  & "','MES')"  
				set rsAddMES = cnnAddMES.Execute(SQLMESActions)
					
		
				response.write "Client " & CustIDPassed & " is added to MES."
			end if
		Else
			rsAddMES.Close
			response.write "Error: Client with Customer ID " & CustIDPassed & " is not found in database " & vbcrlf
		End If 
		
		set rsAddMES = Nothing
		cnnAddMES.Close
		set cnnAddMES = Nothing	
	else 
		response.write "Error: No Customer ID passed"
	end if
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetAccountingPeriodDeleteInformationForModal()

	accountingPeriodsArray = Split(Request.Form("accountingPeriodsArray"),",")
	
	%>

	<input type="hidden" name="periodArray" id="periodArray" value="<%= Request.Form("accountingPeriodsArray") %>">


	<%
	For i = 0 to uBound(accountingPeriodsArray)

		IntRecID = cInt(accountingPeriodsArray(i))
		
		Set rsDelete = Server.CreateObject("ADODB.Recordset")
		rsDelete.CursorLocation = 3 
	
		SQLDelete = "SELECT * FROM Settings_AccountingPeriods WHERE InternalRecordIdentifier = " & IntRecID		
		
		Set cnnDelete = Server.CreateObject("ADODB.Connection")
		cnnDelete.open (Session("ClientCnnString"))
		Set rsDelete = cnnDelete.Execute(SQLDelete)
		
		If NOT rsDelete.EOF Then
			PeriodYear = rsDelete("PeriodYear")
			Period = rsDelete("Period")
			PeriodBeginDate = formatDateTime(rsDelete("BeginDate"),2)
			PeriodEndDate = formatDateTime(rsDelete("EndDate"),2)				
			%><strong><%= PeriodYear %></strong>,&nbsp;Period <%= Period %>,&nbsp;<%= PeriodBeginDate %> - <%= PeriodEndDate %><br><%
		End If
		
	Next
		
	Set rsDelete = Nothing
	cnnDelete.Close
	Set cnnDelete = Nothing

%>


	<label class="control-label" style="padding-left:0px; margin-top:20px;">Click the delete button below to PERMANENTLY DELETE accounting period(s). This cannot be undone.</label>


<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub WritePeriodsInUseDropdownForAccountingYearAdd()
	
	periodYear = cInt(Request.Form("periodYear"))
	periodNum = cInt(Request.Form("periodNum"))
	periodsInUseThisYear = ""
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 

	SQLCheckPeriodInUseForYear = "SELECT * FROM Settings_AccountingPeriods WHERE PeriodYear = " & periodYear

	Set rs = cnn8.Execute(SQLCheckPeriodInUseForYear)
		
	If NOT rs.EOF Then
		Do While Not rs.EOF
			periodsInUseThisYear = periodsInUseThisYear & "---" & rs("Period")
			rs.MoveNext
		Loop							
	End If
						
	cnn8.close
	Set rs = Nothing
	Set cnn8 = Nothing
	
	%>
	<label for="selPeriodNumAdd1">Period</label>
	<select class="form-control" id="selPeriodNumAdd1" name="selPeriodNumAdd1">				
		<%
	
		For i = 1 To 100
		
		  currentPeriod = cStr(i)
		  
		  If InStr(periodsInUseThisYear, currentPeriod) Then
		  
		  	If cInt(periodNum) = cInt(i) Then
		  		%><option value="<%= i %>" disabled selected="selected"><%= i %> (currently in use, please delete first)</option><%
		  	Else
		  		%><option value="<%= i %>" disabled><%= i %> (currently in use, please delete first)</option><%
		  	End If
		  	
		  Else
		  
		  	If cInt(periodNum) = cInt(i) Then
		  		%><option value="<%= i %>" selected="selected"><%= i %></option><%
		  	Else
		  		%><option value="<%= i %>"><%= i %></option><%
		  	End If
		  	
		  End If
		Next
		%>				
	</select>	
	<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub WriteStartDateForAccountingYearAdd()
	
	periodYear = cInt(Request.Form("periodYear"))
	periodNum = cInt(Request.Form("periodNum"))
	periodsInUseThisYear = ""

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 

	SQLCheckPeriodInUseForYear = "SELECT count(*) as TotalRecord FROM Settings_AccountingPeriods WHERE PeriodYear = " & periodYear

	Set rs = cnn8.Execute(SQLCheckPeriodInUseForYear)
		
	If rs.EOF = false Then
		'totalCount = rs.RecordCount
		totalCount = rs("TotalRecord")
	End If	

						
	cnn8.close
	Set rs = Nothing
	Set cnn8 = Nothing



	SQLCheckEndDateInUseForYear = "SELECT TOP 1 [EndDate] FROM Settings_AccountingPeriods WHERE PeriodYear = " & periodYear & " ORDER BY Period DESC"	

	Set cnnBuildDateDataSource = Server.CreateObject("ADODB.Connection")
	cnnBuildDateDataSource.open (Session("ClientCnnString"))
	Set rsBuildDateDataSource = Server.CreateObject("ADODB.Recordset")
	rsBuildDateDataSource.CursorLocation = 3 
	
	Set rsBuildDateDataSource = cnnBuildDateDataSource.Execute(SQLCheckEndDateInUseForYear)
	If rsBuildDateDataSource.EOF = false Then
		recCount = rsBuildDateDataSource.RecordCount
		DateValueNext = rsBuildDateDataSource("EndDate")
	Else
		DateValueNext = ""
	End If	
	
	
	rsBuildDateDataSource.close
	Set rsBuildDateDataSource = Nothing
	Set rsBuildDateDataSource = Nothing
	
	Response.Write DateValueNext & "+" & totalCount
	%>
	
	<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub EditStartDateForAccountingYearAdd()
	
	periodYear = cInt(Request.Form("periodYear"))
	periodNum = cInt(Request.Form("periodNum"))
	periodsInUseThisYear = ""

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 

	SQLCheckPeriodInUseForYear = "SELECT count(*) as TotalRecord FROM Settings_AccountingPeriods WHERE PeriodYear = " & periodYear

	Set rs = cnn8.Execute(SQLCheckPeriodInUseForYear)
		
	If rs.EOF = false Then
		totalCount = rs("TotalRecord")
	End If	

						
	cnn8.close
	Set rs = Nothing
	Set cnn8 = Nothing

	Response.Write totalCount
	%>
	
	<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub ValidateAndAddAccountingPeriod()

	periodYear = Request.Form("periodYear")
	periodNum = Request.Form("periodNum")
	periodStartDate = Request.Form("periodStartDate")
	periodEndDate = Request.Form("periodEndDate")
	
	Set rsValidatePeriodToAdd = Server.CreateObject("ADODB.Recordset")
	rsValidatePeriodToAdd.CursorLocation = 3 

	SQLValidatePeriodToAdd = "SELECT * FROM Settings_AccountingPeriods WHERE PeriodYear = " & periodYear & " AND Period = " & periodNum	
	
	Set cnnValidatePeriodToAdd = Server.CreateObject("ADODB.Connection")
	cnnValidatePeriodToAdd.open (Session("ClientCnnString"))
	Set rsValidatePeriodToAdd = cnnValidatePeriodToAdd.Execute(SQLValidatePeriodToAdd)
	
	If NOT rsValidatePeriodToAdd.EOF Then
	
		Response.write("Period " & periodNum & " already exists in " & periodYear & ".")
		
	Else
		SQLAddPeriod = "INSERT INTO Settings_AccountingPeriods (PeriodYear, Period, BeginDate, EndDate) "
		SQLAddPeriod = SQLAddPeriod & " VALUES (" & periodYear & "," & periodNum & ",'" & periodStartDate & "','" & periodEndDate & "') "
		
		Set cnnAddPeriod = Server.CreateObject("ADODB.Connection")
		cnnAddPeriod.open (Session("ClientCnnString"))
		Set rsAddPeriod = Server.CreateObject("ADODB.Recordset")
		rsAddPeriod.CursorLocation = 3 
		Set rsAddPeriod = cnnAddPeriod.Execute(SQLAddPeriod)
		
		set rsAddPeriod = Nothing
		cnnAddPeriod.close
		set cnnAddPeriod = Nothing
		
		Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added accounting " & periodNum & " in " & periodYear & " ranging from " & periodStartDate & " to " & periodEndDate & "."	 			
		CreateAuditLogEntry "Company Accounting Period Added", "Company Accounting Period Added", "Major", 1, Description		

		Response.write("Success")
			
	End If
	
	Set rsValidatePeriodToAdd = Nothing
	cnnValidatePeriodToAdd.Close
	Set cnnValidatePeriodToAdd = Nothing
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub UpdateAccountingPeriod()

	periodYear = Request.Form("periodYear")
	periodNum = Request.Form("periodNum")
	periodStartDate = Request.Form("periodStartDate")
	periodEndDate = Request.Form("periodEndDate")
	periodIntRecID = Request.Form("periodIntRecID")
	
	Set rsValidatePeriodToUpdate = Server.CreateObject("ADODB.Recordset")
	rsValidatePeriodToUpdate.CursorLocation = 3 

	SQLValidatePeriodToUpdate = "SELECT * FROM Settings_AccountingPeriods WHERE InternalRecordIdentifier = " & periodIntRecID	
	
	Set cnnValidatePeriodToUpdate = Server.CreateObject("ADODB.Connection")
	cnnValidatePeriodToUpdate.open (Session("ClientCnnString"))
	Set rsValidatePeriodToUpdate = cnnValidatePeriodToUpdate.Execute(SQLValidatePeriodToUpdate)
	
	If NOT rsValidatePeriodToUpdate.EOF Then
		orig_periodStartDate = rsValidatePeriodToUpdate("BeginDate")
		orig_periodEndDate = rsValidatePeriodToUpdate("EndDate")			
	End If
	
	SQLValidatePeriodToUpdate = "UPDATE Settings_AccountingPeriods SET BeginDate = '" & periodStartDate & "', EndDate = '" & periodEndDate & "' WHERE InternalRecordIdentifier = " & periodIntRecID
	Set rsValidatePeriodToUpdate = cnnValidatePeriodToUpdate.Execute(SQLValidatePeriodToUpdate)
		
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " edited period " & periodNum & " in " & periodYear & ", changing the date range from (" & orig_periodStartDate & " - " & orig_periodEndDate & ") to (" & periodStartDate & " - " & periodEndDate & ")."	 			
	CreateAuditLogEntry "Company Accounting Period Edited", "Company Accounting Perioding Edited", "Major", 1, Description		
	
	Set rsValidatePeriodToUpdate = Nothing
	cnnValidatePeriodToUpdate.Close
	Set cnnValidatePeriodToUpdate = Nothing
	
	Response.write("Success")
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GenerateMCSPendingChargesPDF()

	'baseURL should always have a trailing /slash, just in case, handle either way
	If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
	sURL = Request.ServerVariables("SERVER_NAME")
	
	DebugMessages = False ' Set to true to turn om Response.Writes
	
	'Generate a unique number to be used for all pdfs throughout this page
	Randomize
	UniqueNum = int((9999999-1111111+1)*rnd+1111111)
	
	Set Pdf = Server.CreateObject("Persits.Pdf")
	Set Doc = Pdf.CreateDocument
	
	ImpVar = baseURL & "bizintel/tools/MCS/MCS_Report1_Tab_Pending_Charges_PDFGen.asp?un=" & Session("UserNo") & "&cl=" & MUV_Read("ClientID") & "&u=" & MUV_Read("SQL_Owner")
	
	If DebugMessages = True Then Response.Write("<br><br><br><br>" & ImpVar & "<br>")
	
	Doc.ImportFromUrl ImpVar, "scale=0.75; hyperlinks=false; landscape=true; drawbackground=true"
	
	fn = "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\MCS_PendingCharges_" & Trim(UniqueNum) & "_" & Replace(FormatDateTime(d,2),"/","") & ".pdf"
	fn = Replace(fn,"/","-")
	fn = Replace(fn,":","-")

	fn2 = Left(baseURL,Len(baseURL)-1) & fn
	fn2 = Replace(fn2,"\","/")

	Filename = Doc.Save(Server.MapPath(fn), False)
	
	'Now wait until the file exists on the server before we try to mail it
	TimeoutSecs = 60
	TimeoutCounter=0
	FOundFile = False
	Do While TimeoutCounter < TimeoutSecs 
		If CheckRemoteURL(fn2) = True Then
			FoundFile = True
			Exit Do ' The file is there
		End If
		DelayResponse(1) ' wait 1 sec & try again
		TimeoutCounter = TimeoutCounter + 1
	Loop
	
	If FoundFile <> True Then 
		Response.End ' Could not fine the pdf, so just bail
	End If
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set fso = Nothing
	
	'*******************************************************************************************************************************
	'*******************************************************************************************************************************
	
	'*******************************************************************************************************************************
	'Return the path of the generated PDF to open in a new window	
	Response.Write(fn2)
	'*******************************************************************************************************************************

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'HELPER FUNCTIONS FOR GENERATE PDF SUB: GenerateMCSPendingChargesPDF()
'********************************************************************************************************************************************************

	Sub DelayResponse(numberOfseconds)
	 Dim WshShell
	 Set WshShell=Server.CreateObject("WScript.Shell")
	 WshShell.Run "waitfor /T " & numberOfSecond & "SignalThatWontHappen", , True
	End Sub
	
	Function CheckRemoteURL(fileURL)
	    ON ERROR RESUME NEXT
	    Dim xmlhttp
	
	    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
	
	    xmlhttp.open "GET", fileURL, False
	    xmlhttp.send
	    If(Err.Number<>0) then
	        Response.Write "Could not connect to remote server"
	    else
	        Select Case Cint(xmlhttp.status)
	            Case 200, 202, 302
	                Set xmlhttp = Nothing
	                CheckRemoteURL = True
	            Case Else
	                Set xmlhttp = Nothing
	                CheckRemoteURL = False
	        End Select
	    end if
	    ON ERROR GOTO 0
	End Function
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'END ALL AJAX MODAL SUBROUTINES AND FUNCTIONS

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

%>