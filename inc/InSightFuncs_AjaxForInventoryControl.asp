<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<!--#include file="InsightFuncs_InventoryControl.asp"-->
<!--#include file="mail.asp"-->
<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
 
'Sub ReturnUMInfoForProduct()
'Sub AssignUPCCodeToProductAndUM()
'Sub RemoveUPCCodeFromICProduct()
'Sub DisplaySKULookupInformation()
'Sub ReturnAvailabilityInfoForProduct()

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
	Case "ReturnUMInfoForProduct"
		ReturnUMInfoForProduct()
	Case "AssignUPCCodeToProductAndUM"
		AssignUPCCodeToProductAndUM()
	Case "RemoveUPCCodeFromICProduct"
		RemoveUPCCodeFromICProduct()
	Case "DisplaySKULookupInformation"
		DisplaySKULookupInformation()
	Case "ReturnAvailabilityInfoForProduct"
		ReturnAvailabilityInfoForProduct()
End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub DisplaySKULookupInformation()


	prodSKU = Request.Form("prodSKU")
	prodUM = ""
	prodDesc = ""
	prodBin = ""

	Set cnnprodLookup = Server.CreateObject("ADODB.Connection")
	cnnprodLookup.open (Session("ClientCnnString"))
	Set rsprodLookup = Server.CreateObject("ADODB.Recordset")
	rsprodLookup.CursorLocation = 3 
		
	SQL_prodLookup = "SELECT * FROM IC_Product WHERE prodSKU = '" & prodSKU & "'"
	
	Set rsprodLookup = cnnprodLookup.Execute(SQL_prodLookup)
	
	If Not rsprodLookup.EOF Then
	
		'First get the on hand qty
		QtyOnHand_Units = rsprodLookup("QtyOnHand_Units")
		prodCaseConversionFactor = rsprodLookup("prodCaseConversionFactor")			
		prodCasePricing = rsprodLookup("prodCasePricing")
		prodDesc = rsprodLookup("prodDescription")
		prodUnitBin = rsprodLookup("prodUnitBin")
		prodCaseBin = rsprodLookup("prodCaseBin")	
		prodImage = GetProdImage(prodSKU)	
		QtyOnHand_LastUpdated = rsprodLookup("QtyOnHand_LastUpdated")
		
	End If
				
	'Post to the backend to try to get an up-to-the-minute on hand
	'Construct xml fields based on record
	xmlData = "<DATASTREAM>"
	xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			
	xmlData = xmlData & "<MODE>" & GetPOSTParams("BackendInventoryPostsMode") & "</MODE>"
			
	xmlData = xmlData & "  <RECORD_TYPE>INVENTORY</RECORD_TYPE>"
	xmlData = xmlData & "  <RECORD_SUBTYPE>QUERY_ONHAND</RECORD_SUBTYPE>"
			
	xmlData = xmlData & "<SERNO>" & MUV_READ("SERNO") & "</SERNO>"
			
				
	xmlData = xmlData & " <QUERY_ONHAND>"

	xmlData = xmlData & "        <PROD_ID>" & prodSKU & "</PROD_ID>"
	xmlData = xmlData & "        <RETURN_VALUE_UM>U</RETURN_VALUE_UM>"
	
	xmlData = xmlData & " </QUERY_ONHAND>"
		 
	xmlData = xmlData & "</DATASTREAM>"

	xmlDataForDisp = Replace(xmlData,"<","[")
	xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
	xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
	xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
	xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")

	Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")

	'Response.Write(GetPOSTParams("BackendInventoryPostsURL"))
	
	httpRequest.Open "POST", GetPOSTParams("BackendInventoryPostsURL"), False
	httpRequest.SetRequestHeader "Content-Type", "text/xml"
	
	xmlData = Replace(xmlData,"&","&amp;")
	xmlData = Replace(xmlData,chr(34),"")			
	httpRequest.Send xmlData

	data = xmlData

	If (Err.Number <> 0 ) Then
		emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND"& "<br><br>"
		emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
		emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
		emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
		emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
		SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",SERNO & " Post Error Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
		Description = emailBody 
		Write_API_AuditLog_Entry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("BackendInventoryPostsMode"),MUV_READ("SERNO"),MUV_READ("SERNO"),"Inventory API"
	End If

	If httpRequest.status = 200 THEN 
	
		If IsNumeric(httpRequest.responseText) Then ' Success
	
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND<"& "<br><br>"
			emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
			SendMail "mailsender@" & maildomain ,"insight@ocsaccess.com", MUV_READ("SERNO") & " Good Post Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
			
			Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("BackendInventoryPostsMode"),"CountOnHand_submit.asp")
			
			QtyOnHand_Units = httpRequest.responseText
			QtyOnHand_UnitsStatus = "LIVE"
			QtyOnHand_UnitsStatus = "1 MIN AGO"
			
		Else
			'FAILURE
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND<"& "<br><br>"
			emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
			SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",MUV_READ("SERNO") & " Post Error Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
		
			Call Write_API_AuditLog_Entry(Identity ,emailBody ,GetPOSTParams("BackendInventoryPostsMode"),"CountOnHand_submit.asp")

	
			QtyOnHand_UnitsStatus = "CACHED"
			
			QtyOnHand_LastUpdatedTimeDiff = DateDiff("n",Now(),QtyOnHand_LastUpdated)
			QtyOnHand_LastUpdatedTimeDiff = Abs(QtyOnHand_LastUpdatedTimeDiff) 
			
			If cInt(QtyOnHand_LastUpdatedTimeDiff) >= cInt(1440) Then
			
				QtyOnHand_LastUpdatedDays = QtyOnHand_LastUpdatedTimeDiff \ 1440
				QtyOnHand_LastUpdatedHours = (QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedDays * 1440)) \ 60
				QtyOnHand_LastUpdatedMinutes = (QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedDays * 1440)) mod 60 
				QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedDays & " DAYS " & QtyOnHand_LastUpdatedHours & " HRS " & QtyOnHand_LastUpdatedMinutes & " MIN AGO "
				
			ElseIf cInt(QtyOnHand_LastUpdatedTimeDiff) >= cInt(60) AND cInt(QtyOnHand_LastUpdatedTimeDiff) < cInt(1440) Then
			
				QtyOnHand_LastUpdatedHours = QtyOnHand_LastUpdatedTimeDiff \ 60
				QtyOnHand_LastUpdatedMinutes = QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedHours * 60)
				QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedHours & " HRS " & QtyOnHand_LastUpdatedMinutes & " MIN AGO "
				
			ElseIf cInt(QtyOnHand_LastUpdatedTimeDiff) < cInt(60) Then
			
				QtyOnHand_LastUpdatedMinutes = QtyOnHand_LastUpdatedTimeDiff
				QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedMinutes & " MIN AGO "
				
			End If
					
		End If
		
	Else
	
			'FAILURE
			emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND<"& "<br><br>"
			emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
			SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",MUV_READ("SERNO") & " Post Error Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
			
			QtyOnHand_UnitsStatus = "CACHED"
			
			QtyOnHand_LastUpdatedTimeDiff = DateDiff("n",Now(),QtyOnHand_LastUpdated)
			QtyOnHand_LastUpdatedTimeDiff = Abs(QtyOnHand_LastUpdatedTimeDiff) 
			
			If cInt(QtyOnHand_LastUpdatedTimeDiff) >= cInt(1440) Then
			
				QtyOnHand_LastUpdatedDays = QtyOnHand_LastUpdatedTimeDiff \ 1440
				QtyOnHand_LastUpdatedHours = (QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedDays * 1440)) \ 60
				QtyOnHand_LastUpdatedMinutes = (QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedDays * 1440)) mod 60 
				QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedDays & " DAYS " & QtyOnHand_LastUpdatedHours & " HRS " & QtyOnHand_LastUpdatedMinutes & " MIN AGO "
				
			ElseIf cInt(QtyOnHand_LastUpdatedTimeDiff) >= cInt(60) AND cInt(QtyOnHand_LastUpdatedTimeDiff) < cInt(1440) Then
			
				QtyOnHand_LastUpdatedHours = QtyOnHand_LastUpdatedTimeDiff \ 60
				QtyOnHand_LastUpdatedMinutes = QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedHours * 60)
				QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedHours & " HRS " & QtyOnHand_LastUpdatedMinutes & " MIN AGO "
				
			ElseIf cInt(QtyOnHand_LastUpdatedTimeDiff) < cInt(60) Then
			
				QtyOnHand_LastUpdatedMinutes = QtyOnHand_LastUpdatedTimeDiff
				QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedMinutes & " MIN AGO "
				
			End If
			
		
			Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("BackendInventoryPostsMode"),"CountOnHand_submit.asp")

	End If



	Set rsprodLookup = Nothing
	cnnprodLookup.Close
	Set cnnprodLookup = Nothing
	%>

    <div class="row">
    
		<div class="col-lg-7 col-md-7 col-sm-7 col-xs-7">
			<div class="container-fluid" style="padding-left:0; padding-right:0">
				<div class="row">

					<div class="col-xs-12" style="padding-bottom:10px;"><strong>Product ID:</strong>&nbsp;<strong class="red"><%= prodSKU %></strong></div>
					
					<% If prodCasePricing = "N" Then %>
						<div class="col-xs-12" style="padding-bottom:10px;">
							<strong>U/M:</strong>&nbsp;<strong class="red">N</strong>
						</div>
					<% End If %>
					
					<% If prodCasePricing = "N" Then %>
						<div class="col-xs-12" style="padding-bottom:5px;"><strong>Bin:</strong>&nbsp;<strong class="red"><%= prodUnitBin %></strong></div>
					<% Else %>
						<div class="col-xs-12" style="padding-bottom:5px;"><strong>Unit Bin:</strong>&nbsp;<strong class="red"><%= prodUnitBin %></strong></div>
						<div class="col-xs-12" style="padding-bottom:5px;"><strong>Case Bin:</strong>&nbsp;<strong class="red"><%= prodCaseBin %></strong></div>
					<% End If %>
				</div>
			</div>
		</div>
		
		<div class="col-lg-5 col-md-5 col-sm-5 col-xs-5">
			<% If prodImage <> "" Then %>
				<img src="<%=GetProdImage(prodSKU )%>" class="general-image mobile-image img-thumbnail" style="width:100%;">
			<% End IF %>
		</div>
		
	</div>
    
	<div class="row row-line">
		<div class="col-xs-12" style="padding-bottom:10px;"><strong>Description:</strong>&nbsp;<strong class="red"><%= prodDesc %></strong></div>
	</div>
	
	<div class="row" style="margin-left:5px;margin-right:5px;">	
	
		<div class="col-xs-12" style="border-top:2px solid #000000;"></div>
		
		<% If prodCasePricing = "N" Then %>
		
			<div class="row" style="padding-bottom:5px;">
				<div class="col-xs-6 text-center"><strong>On Hand</strong></div>
				<div class="col-xs-6 text-center"><strong>Tot Units</strong></div>
			</div>
			
			<div class="row row-info">
				<div class="col-xs-6 text-center"><%= QtyOnHand_UnitsStatus %></div>
				<div class="col-xs-6 text-center"><strong class="red"><%= QtyOnHand_Units %></strong></div>				
			</div>
				
 		<% Else %>
 		
			<div class="row" style="padding-bottom:5px;">
				<div class="col-xs-3"><strong>On Hand</strong></div>
				<div class="col-xs-3 text-center"><strong>Tot Units</strong></div>	
				<div class="col-xs-3 text-center"><strong>Cases</strong></div>
				<div class="col-xs-3 text-center"><strong>Units</strong></div>
			</div>
			
			<div class="row row-info">
			
				<div class="col-xs-3 text-center"><%= QtyOnHand_UnitsStatus %></div>

				<div class="col-xs-3 text-center">
					<strong class="red"><%= QtyOnHand_Units %></strong>
				</div>	
				
				<% If prodCaseConversionFactor <> "" AND cint(prodCaseConversionFactor) <> 0 Then %>
					<div class="col-xs-3 text-center"><strong class="red"><%= Int(QtyOnHand_Units / cInt(prodCaseConversionFactor))  %></strong></div>
				<% Else %>
					<div class="col-xs-3 text-center">---</div>
				<% End If %>
				
				<% If cint(prodCaseConversionFactor) <> 0 Then %>
					<% If QtyOnHand_Units Mod cint(prodCaseConversionFactor) = 0 Then %>	
						<div class="col-xs-3 text-center"><strong class="red">---</strong></div>
					<% Else %>
						<div class="col-xs-3 text-center"><strong class="red"><%=  QtyOnHand_Units Mod cInt(prodCaseConversionFactor) %></strong></div>	
					<% End If %>	
				<% Else %>
					<div class="col-xs-3 text-center"><strong class="red">---</strong></div>
				<% End If %>
					
			</div>
				
 		<% End If %>
		
	</div>
		
	<form method="POST" action="CountOnHand_submit.asp" name="frmCountOnHand">

		<input type="hidden" name="txtprodSKU" id="txtprodSKU" value="<%= prodSKU %>">
		<input type="hidden" name="txtproductUPC" id="txtproductUPC" value="<%= UPCCode %>">    

 
 		<div class="row row-line" style="margin-left:-8px;margin-right:5px;">
		
		<% If prodCasePricing <> "N" Then %>
		
			<div class="col-xs-4"><strong>New Count:</strong></div>
			<div class="col-xs-3"><strong>Cases</strong></div>
			<div class="col-xs-3"><strong>Units</strong></div>
			<div class="col-xs-4">&nbsp;</div>
			<div class="col-xs-3"><input type="search" class="form-control" name="txtCasesCounted" id="txtCasesCounted" AUTOCOMPLETE="off"></div>
			<div class="col-xs-3"><input type="search" class="form-control" name="txtUnitsCounted" id="txtUnitsCounted" AUTOCOMPLETE="off"></div>

		<% Else %>
		
			<div class="col-xs-4"><strong>New Count:</strong></div>
			<div class="col-xs-3"><strong>Units</strong></div>
			<div class="col-xs-3">&nbsp;</div>
			<div class="col-xs-4">&nbsp;</div>
			<div class="col-xs-3"><input type="search" class="form-control" name="txtUnitsCounted" id="txtUnitsCounted" AUTOCOMPLETE="off"></div>
			<div class="col-xs-3">&nbsp;</div>
		
		<% End If %>
		
		
		</div>
		<div class="row row-line" style="margin-left:-8px;margin-right:5px;">
			
		<% If prodCasePricing = "N" Then %>
		
			<div class="col-xs-4"><strong>New Bin:</strong></div>
			<div class="col-xs-5"><input type="search" class="form-control" name="txtBinLocation" id="txtBinLocation" AUTOCOMPLETE="off"></div>
	
		<% Else %>
		
			<div class="col-xs-4"><strong>New Bin:</strong></div>
			<div class="col-xs-4"><strong>Case Bin</strong></div>
			<div class="col-xs-4"><strong>Unit Bin</strong></div>
			<div class="col-xs-4">&nbsp;</div>
			<div class="col-xs-4"><input type="text" class="form-control" name="txtCaseBinLocation" id="txtCaseBinLocation" AUTOCOMPLETE="off"></div>
			<div class="col-xs-4"><input type="text" class="form-control" name="txtUnitBinLocation" id="txtUnitBinLocation" AUTOCOMPLETE="off"></div>	
	
		<% End If %>
		
		</div>
			
		<div class="row row-line">
			<div class="col-xs-12"><button class="btn btn-primary btn-go btn-md">SUBMIT INVENTORY CHANGES</button></div>
		</div>
		
	</form>
	

<%
End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************






'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub ReturnUMInfoForProduct()

	productSKUSelected = Request.Form("prodSKU") 
	dummy = MUV_Write("savedProdSKUFromLookup",productSKUSelected)
	
	Set rsReturnUMInfo= Server.CreateObject("ADODB.Recordset")
	rsReturnUMInfo.CursorLocation = 3 
	
	If productSKUSelected <> "" Then	
	
		SQLReturnUMInfo = "SELECT * FROM IC_Product WHERE prodSKU = '" & productSKUSelected & "'"

		Set cnnReturnUMInfo = Server.CreateObject("ADODB.Connection")
		cnnReturnUMInfo.open (Session("ClientCnnString"))
		Set rsReturnUMInfo = cnnReturnUMInfo.Execute(SQLReturnUMInfo)
		
		If NOT rsReturnUMInfo.EOF Then

			Do While Not rsReturnUMInfo.EOF
											
				ProdCasePricingFlag = rsReturnUMInfo("ProdCasePricing")
				
				If ProdCasePricingFlag = "N" Then
					%>
						<div class="col-xs-12"><strong class="green">N PRODUCT, WILL BE AUTOMATICALLY ASSIGNED TO UNIT</strong></div>
			        	<div class="clearfix"></div><br>						
						<div class='col-xs-12'>
							<select class="form-control" name="selProdUM" id="selProdUM">
								<option value="UNIT" selected="selected">ASSIGN TO UNIT</option>
							</select>
						</div>
			        	<div class="clearfix"></div><br>				    				    				    	
					<%	
				Else
					%>
						<div class="col-xs-12"><strong class="green">Please Select U/M To Assign UPC Code To:</strong></div>
			        	<div class="clearfix"></div><br>						
						<div class='col-xs-12'>
							<select class="form-control" name="selProdUM" id="selProdUM">
								<option value="">SELECT UNIT</option>
								<option value="UNIT">ASSIGN TO UNIT</option>
								<option value="CASE">ASSIGN TO CASE</option>
							</select>
						</div>
			        	<div class="clearfix"></div><br>				    				    				    	
					<%
				End If
				
			rsReturnUMInfo.MoveNext
			Loop
		Else
			Response.Write("No Product SKU was found in the product file matching " & productSKUSelected)
		End If
		
		
		set rsReturnUMInfo = Nothing
		cnnReturnUMInfo.close
		set cnnReturnUMInfo = Nothing
		
	Else
		Response.Write("Cannot Lookup, Invalid Data")
		
	End If

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub AssignUPCCodeToProductAndUM()

	productSKUSelected = Request.Form("prodSKU") 
	productUMSelected = Request.Form("prodUM")
	productUPCEntered = Request.Form("prodUPC")
	dummy = MUV_Write("savedProdSKUFromLookup",productSKUSelected)
	
	If productSKUSelected <> "" AND productUMSelected <> "" AND productUPCEntered <> "" Then	
	
		If productUMSelected = "UNIT" Then
			SQLAssignUPCCode = "UPDATE IC_Product SET prodUnitUPC = '" & productUPCEntered & "' WHERE prodSKU = '" & productSKUSelected & "'"
		ElseIf productUMSelected = "CASE" Then
			SQLAssignUPCCode = "UPDATE IC_Product SET prodCaseUPC = '" & productUPCEntered & "' WHERE prodSKU = '" & productSKUSelected & "'"
		End If

		Set rsAssignUPCCode= Server.CreateObject("ADODB.Recordset")
		rsAssignUPCCode.CursorLocation = 3 	
		Set cnnAssignUPCCode = Server.CreateObject("ADODB.Connection")
		cnnAssignUPCCode.open (Session("ClientCnnString"))
		Set rsAssignUPCCode = cnnAssignUPCCode.Execute(SQLAssignUPCCode)
				
		set rsAssignUPCCode = Nothing
		cnnAssignUPCCode.close
		set cnnAssignUPCCode = Nothing
		
		Response.Write("SQLAssignUPCCode : " & SQLAssignUPCCode)
		
	Else
	
		Response.Write("Cannot Assign UPC Code, Invalid Data")
		
	End If

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub RemoveUPCCodeFromICProduct()

	productUPCEntered = Request.Form("prodUPC")
	
	If productUPCEntered <> "" Then	

		Set rsRemoveUPCCode= Server.CreateObject("ADODB.Recordset")
		rsRemoveUPCCode.CursorLocation = 3 	
		Set cnnRemoveUPCCode = Server.CreateObject("ADODB.Connection")
		cnnRemoveUPCCode.open (Session("ClientCnnString"))
	

		SQLRemoveUPCCode = "UPDATE IC_Product SET prodUnitUPC = '' WHERE prodUnitUPC = '" & productUPCEntered & "'"
		Set rsRemoveUPCCode = cnnRemoveUPCCode.Execute(SQLRemoveUPCCode)
		
		SQLRemoveUPCCode = "UPDATE IC_Product SET prodCaseUPC = '' WHERE prodCaseUPC = '" & productUPCEntered & "'"
		Set rsRemoveUPCCode = cnnRemoveUPCCode.Execute(SQLRemoveUPCCode)

		set rsRemoveUPCCode = Nothing
		cnnRemoveUPCCode.close
		set cnnRemoveUPCCode = Nothing
		
		'Response.Write("SQLRemoveUPCCode : " & SQLRemoveUPCCode)
	Else
	
		Response.Write("Cannot Remove UPC Code from IC Product, Invalid Data")
		
	End If

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************








'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub ReturnAvailabilityInfoForProduct()


	prodSKU = Request.Form("prodSKU")

	Set cnnprodLookup = Server.CreateObject("ADODB.Connection")
	cnnprodLookup.open (Session("ClientCnnString"))
	Set rsprodLookup = Server.CreateObject("ADODB.Recordset")
	rsprodLookup.CursorLocation = 3 
		
	SQL_prodLookup = "SELECT * FROM IC_Product WHERE prodSKU = '" & prodSKU & "'"
	
	Set rsprodLookup = cnnprodLookup.Execute(SQL_prodLookup)
	
	If Not rsprodLookup.EOF Then

		prodCaseConversionFactor = rsprodLookup("prodCaseConversionFactor")			
		prodCasePricing = rsprodLookup("prodCasePricing")
		prodDesc = rsprodLookup("prodDescription")
		prodUnitBin = rsprodLookup("prodUnitBin")
		prodCaseBin = rsprodLookup("prodCaseBin")	
		prodImage = GetProdImage(prodSKU)	
		QtyOnHand_LastUpdated = rsprodLookup("QtyOnHand_LastUpdated")
		
	End If
	
	QtyAvailable_Units = 0
	QtyAvailable_UnitsStatus = "NO BACKEND"
				
	'Post to the backend to try to get an up-to-the-minute on hand
	'Construct xml fields based on record
	xmlData = "<DATASTREAM>"
	xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			
	xmlData = xmlData & "<MODE>" & GetPOSTParams("BackendInventoryPostsMode") & "</MODE>"
			
	xmlData = xmlData & "  <RECORD_TYPE>INVENTORY</RECORD_TYPE>"
	xmlData = xmlData & "  <RECORD_SUBTYPE>QUERY_AVAILABLE</RECORD_SUBTYPE>"
			
	xmlData = xmlData & "<SERNO>" & MUV_READ("SERNO") & "</SERNO>"
			
				
	xmlData = xmlData & " <QUERY_ONHAND>"

	xmlData = xmlData & "        <PROD_ID>" & prodSKU & "</PROD_ID>"
	xmlData = xmlData & "        <RETURN_VALUE_UM>U</RETURN_VALUE_UM>"
	
	xmlData = xmlData & " </QUERY_ONHAND>"
		 
	xmlData = xmlData & "</DATASTREAM>"

	xmlDataForDisp = Replace(xmlData,"<","[")
	xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
	xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
	xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
	xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")

	Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")

	'Response.Write(GetPOSTParams("BackendInventoryPostsURL"))
	
	httpRequest.Open "POST", GetPOSTParams("BackendInventoryPostsURL"), False
	httpRequest.SetRequestHeader "Content-Type", "text/xml"
	
	xmlData = Replace(xmlData,"&","&amp;")
	xmlData = Replace(xmlData,chr(34),"")			
	httpRequest.Send xmlData

	data = xmlData

	If (Err.Number <> 0 ) Then
		emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND"& "<br><br>"
		emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
		emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
		emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
		emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
		SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",SERNO & " Post Error Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
		Description = emailBody 
		Write_API_AuditLog_Entry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("BackendInventoryPostsMode"),MUV_READ("SERNO"),MUV_READ("SERNO"),"Inventory API"
	End If

	If httpRequest.status = 200 THEN 
	
		If IsNumeric(httpRequest.responseText) Then ' Success
	
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND<"& "<br><br>"
			emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
			SendMail "mailsender@" & maildomain ,"insight@ocsaccess.com", MUV_READ("SERNO") & " Good Post Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
			
			Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("BackendInventoryPostsMode"),"CountOnHand_submit.asp")
			
			QtyAvailable_Units = httpRequest.responseText
			QtyAvailable_UnitsStatus = "LIVE"
			QtyAvailable_UnitsStatus = "1 MIN AGO"
			
		Else
			'FAILURE
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND<"& "<br><br>"
			emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
			SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",MUV_READ("SERNO") & " Post Error Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
			
			QtyAvailable_UnitsStatus = "CACHED"
			
			QtyAvailable_LastUpdatedTimeDiff = DateDiff("n",Now(),QtyAvailable_LastUpdated)
			QtyAvailable_LastUpdatedTimeDiff = Abs(QtyAvailable_LastUpdatedTimeDiff) 
			
			If cInt(QtyAvailable_LastUpdatedTimeDiff) >= cInt(1440) Then
			
				QtyAvailable_LastUpdatedDays = QtyAvailable_LastUpdatedTimeDiff \ 1440
				QtyAvailable_LastUpdatedHours = (QtyAvailable_LastUpdatedTimeDiff - (QtyAvailable_LastUpdatedDays * 1440)) \ 60
				QtyAvailable_LastUpdatedMinutes = (QtyAvailable_LastUpdatedTimeDiff - (QtyAvailable_LastUpdatedDays * 1440)) mod 60 
				QtyAvailable_UnitsStatus = QtyAvailable_LastUpdatedDays & " DAYS " & QtyAvailable_LastUpdatedHours & " HRS " & QtyAvailable_LastUpdatedMinutes & " MIN AGO "
				
			ElseIf cInt(QtyAvailable_LastUpdatedTimeDiff) >= cInt(60) AND cInt(QtyAvailable_LastUpdatedTimeDiff) < cInt(1440) Then
			
				QtyAvailable_LastUpdatedHours = QtyAvailable_LastUpdatedTimeDiff \ 60
				QtyAvailable_LastUpdatedMinutes = QtyAvailable_LastUpdatedTimeDiff - (QtyAvailable_LastUpdatedHours * 60)
				QtyAvailable_UnitsStatus = QtyAvailable_LastUpdatedHours & " HRS " & QtyAvailable_LastUpdatedMinutes & " MIN AGO "
				
			ElseIf cInt(QtyAvailable_LastUpdatedTimeDiff) < cInt(60) Then
			
				QtyAvailable_LastUpdatedMinutes = QtyAvailable_LastUpdatedTimeDiff
				QtyAvailable_UnitsStatus = QtyAvailable_LastUpdatedMinutes & " MIN AGO "
				
			End If
			
		
			Call Write_API_AuditLog_Entry(Identity ,emailBody ,GetPOSTParams("BackendInventoryPostsMode"),"CountOnHand_submit.asp")
			
		End If
		
	Else
	
		'FAILURE
		emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND<"& "<br><br>"
		emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
		emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
		emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
		emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
		SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",MUV_READ("SERNO") & " Post Error Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
	
		Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("BackendInventoryPostsMode"),"CountOnHand_submit.asp")

	End If


	Set rsprodLookup = Nothing
	cnnprodLookup.Close
	Set cnnprodLookup = Nothing
	%>

	<div class="row">
		<div class="col-xs-12 text-center" style="padding-bottom:10px;"><strong style="font-size:30px;">AVAILABLE UNITS:</strong></div>
	</div>	
			
	<div class="row row-line">
		<div class="col-xs-12 text-center" style="padding-bottom:10px;">
			<% If QtyAvailable_UnitsStatus = "NO BACKEND" Then %>
				<strong class="red" style="font-size:40px;">No Backend Connection</strong>
			<% Else %>
				<strong class="green" style="font-size:80px;"><%= QtyAvailable_Units %></strong>
			<% End If %>
		</div>
	</div>

    <div class="row">
    
		<div class="col-lg-7 col-md-7 col-sm-7 col-xs-7">
			<div class="container-fluid" style="padding-left:0; padding-right:0">
				<div class="row">

					<div class="col-xs-12" style="padding-bottom:10px;"><strong>Product ID:</strong>&nbsp;<strong class="red"><%= prodSKU %></strong></div>
					
					<% If prodCasePricing = "N" Then %>
						<div class="col-xs-12" style="padding-bottom:10px;">
							<strong>U/M:</strong>&nbsp;<strong class="red">N</strong>
						</div>
					<% End If %>
					
					<% If prodCasePricing = "N" Then %>
						<div class="col-xs-12" style="padding-bottom:5px;"><strong>Bin:</strong>&nbsp;<strong class="red"><%= prodUnitBin %></strong></div>
					<% Else %>
						<div class="col-xs-12" style="padding-bottom:5px;"><strong>Unit Bin:</strong>&nbsp;<strong class="red"><%= prodUnitBin %></strong></div>
						<div class="col-xs-12" style="padding-bottom:5px;"><strong>Case Bin:</strong>&nbsp;<strong class="red"><%= prodCaseBin %></strong></div>
					<% End If %>
				</div>
			</div>
		</div>
		
		<div class="col-lg-5 col-md-5 col-sm-5 col-xs-5">
			<% If prodImage <> "" Then %>
				<img src="<%=GetProdImage(prodSKU )%>" class="general-image mobile-image img-thumbnail" style="width:100%;">
			<% End IF %>
		</div>
		
	</div>
    
	<div class="row row-line">
		<div class="col-xs-12" style="padding-bottom:10px;"><strong>Description:</strong>&nbsp;<strong class="red"><%= prodDesc %></strong></div>
	</div>
	
	<div class="row" style="margin-left:5px;margin-right:5px;">	
	
		<div class="col-xs-12" style="border-top:2px solid #000000;"></div>
		
		<% If prodCasePricing = "N" Then %>
		
			<div class="row" style="padding-bottom:5px;">
				<div class="col-xs-6 text-center"><strong>On Hand</strong></div>
				<div class="col-xs-6 text-center"><strong>Tot Units</strong></div>
			</div>
			
			<div class="row row-info">
				<div class="col-xs-6 text-center"><%= QtyAvailable_UnitsStatus %></div>
				<div class="col-xs-6 text-center"><strong class="red"><%= QtyAvailable_Units %></strong></div>				
			</div>
				
 		<% Else %>
 		
			<div class="row" style="padding-bottom:5px;">
				<div class="col-xs-3"><strong>On Hand</strong></div>
				<div class="col-xs-3 text-center"><strong>Tot Units</strong></div>	
				<div class="col-xs-3 text-center"><strong>Cases</strong></div>
				<div class="col-xs-3 text-center"><strong>Units</strong></div>
			</div>
			
			<div class="row row-info">
			
				<div class="col-xs-3 text-center"><%= QtyAvailable_UnitsStatus %></div>

				<div class="col-xs-3 text-center">
					<strong class="red"><%= QtyAvailable_Units %></strong>
				</div>	
				
				<% If prodCaseConversionFactor <> "" AND cint(prodCaseConversionFactor) <> 0 Then %>
					<div class="col-xs-3 text-center"><strong class="red"><%= Int(QtyAvailable_Units / cInt(prodCaseConversionFactor))  %></strong></div>
				<% Else %>
					<div class="col-xs-3 text-center">---</div>
				<% End If %>
				
				<% If cint(prodCaseConversionFactor) <> 0 Then %>
					<% If QtyAvailable_Units Mod cint(prodCaseConversionFactor) = 0 Then %>	
						<div class="col-xs-3 text-center"><strong class="red">---</strong></div>
					<% Else %>
						<div class="col-xs-3 text-center"><strong class="red"><%=  QtyAvailable_Units Mod cInt(prodCaseConversionFactor) %></strong></div>	
					<% End If %>	
				<% Else %>
					<div class="col-xs-3 text-center"><strong class="red">---</strong></div>
				<% End If %>
					
			</div>
				
 		<% End If %>
		
	</div>
		
	

<%
End Sub


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