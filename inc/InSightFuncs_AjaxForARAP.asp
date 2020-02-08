<!--#include file="SubsAndFuncs.asp"-->
<!--#include file="InSightFuncs.asp"-->
<!--#include file="InsightFuncs_AR_AP.asp"-->
<!--#include file="mail.asp"-->
<!--#include file="InsightFuncs_API.asp"-->
<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
 
'Sub SaveEquivalentCustomerAccount()
'Sub GetContentForEmailConsolidatedInvoiceModalAccount()
'Sub EmailConsolidatedInvoiceModalAccount()
'Sub GetContentForEmailConsolidatedInvoiceModalChain()
'Sub EmailConsolidatedInvoiceModalChain()
'Sub ConvertCatchallAccountTransactions(passedOurCustAccount,passedEquivCustAccount,passedPartnerRecId)
'Sub SaveCustomerMES()
'Sub GetContentForSageInvoiceDetailExpansion()
'Sub GetContentForEditSageInvoiceModal()
'Sub SaveEditSageInvoice()
'Sub ExportSelectedSageInvoices()
'Sub ToggleShowHideExportedSageInvoices()
'Sub CheckIfDefaultBillingLocation()
'Sub GetCustomerAccountInformationForModal()
'Sub GetCustomerPricingInformationForModal()
'Sub CheckIfCustomerIDAlreadyExists()
'Sub GetCustomerNoteCountByNoteType()
'Sub GetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes()
'Sub MarkAllNotesForNoteTypeForUserAsRead()
'Sub GetContentForCustomerNotesModal()

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
	Case "SaveEquivalentCustomerAccount" 
		SaveEquivalentCustomerAccount()	
	Case "GetContentForEmailConsolidatedInvoiceModalAccount" 
		GetContentForEmailConsolidatedInvoiceModalAccount()
	Case "EmailConsolidatedInvoiceModalAccount" 
		EmailConsolidatedInvoiceModalAccount()	
	Case "GetContentForEmailConsolidatedInvoiceModalChain" 
		GetContentForEmailConsolidatedInvoiceModalChain()
	Case "EmailConsolidatedInvoiceModalChain" 
		EmailConsolidatedInvoiceModalChain()
	Case "SaveCustomerMES"
		SaveCustomerMES()	
	Case "GetContentForSageInvoiceDetailExpansion"
		GetContentForSageInvoiceDetailExpansion()
	Case "GetContentForEditSageInvoiceModal"
		GetContentForEditSageInvoiceModal()		
	Case "SaveEditSageInvoice"
		SaveEditSageInvoice()		
	Case "ExportSelectedSageInvoices"
		ExportSelectedSageInvoices()	
	Case "ToggleShowHideExportedSageInvoices"
		ToggleShowHideExportedSageInvoices()
	Case "CheckIfDefaultBillingLocation"
		CheckIfDefaultBillingLocation()	
	Case "GetCustomerAccountInformationForModal"
		GetCustomerAccountInformationForModal()
	Case "CheckIfCustomerIDAlreadyExists"
		CheckIfCustomerIDAlreadyExists()	
	Case "GetCustomerPricingInformationForModal"
		GetCustomerPricingInformationForModal()	
	Case "GetCustomerNoteCountByNoteType"
		GetCustomerNoteCountByNoteType()
	Case "GetContentForCustomerNotesModal"
		GetContentForCustomerNotesModal()
	Case "GetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes"
		GetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes()
	Case "MarkAllNotesForNoteTypeForUserAsRead"
		MarkAllNotesForNoteTypeForUserAsRead()
End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForEditSageInvoiceModal() 

	InvoiceID = Request.Form("InvoiceID")
	LineItemNo = Request.Form("LineItemNo")
	
	%>
	<input type="hidden" id="txtInvoiceID" name="txtInvoiceID" value="<%= InvoiceID %>">
	<input type="hidden" id="txtLineItemNo" name="txtLineItemNo" value="<%= LineItemNo %>">
	
	<%
	
	SQLEditInvoiceID = "SELECT * FROM IN_InvoiceHistDetail where InvoiceID = '" & InvoiceID & "' AND LineNumber = " & LineItemNo
		
	Set cnnEditInvoiceID = Server.CreateObject("ADODB.Connection")
	cnnEditInvoiceID.open(Session("ClientCnnString"))
	Set rsEditInvoiceID = Server.CreateObject("ADODB.Recordset")
	rsEditInvoiceID.CursorLocation = 3 
	
	Set rsEditInvoiceID = cnnEditInvoiceID.Execute(SQLEditInvoiceID)

	If not rsEditInvoiceID.EOF Then
		GLARAccount = rsEditInvoiceID("GL_AR_Account")
		GLAccount = rsEditInvoiceID("GL_Account")
	End If
		
	Set rsEditInvoiceID = Nothing
	cnnEditInvoiceID.Close
	Set cnnEditInvoiceID = Nothing
	
	%>
			
	<script type="text/javascript">
		
		$(document).ready(function() {
			
			$("#modalEditSageInvoice #btnEditSageInvoiceSave").bind("click",function(e){
			
				var GLARAccount = $("#modalEditSageInvoice #txtGLARAccount").val();
				var GLAccount = $("#modalEditSageInvoice #txtGLAccount").val();
				var InvoiceID = $("#modalEditSageInvoice #txtInvoiceID").val();
				var LineItemNo = $("#modalEditSageInvoice #txtLineItemNo").val();
				
		
				if (GLARAccount.length <=0) {
					swal({
						title: 'Error Saving Changes',
						text: 'Please specify a GL AR Account',
						type: 'error'
					});
					return false;
				}
		
				if (GLAccount.length <=0) {
					swal({
						title: 'Error Saving Changes',
						text: 'Please specify a GL Account',
						type: 'error'
					});
					return false;
				}
						
					
		    	$.ajax({
					type:"POST",
					url: "../../../../inc/InSightFuncs_AjaxForARAP.asp",
					cache: false,
					data: "action=SaveEditSageInvoice&LineItemNo=" + encodeURIComponent(LineItemNo) + "&InvoiceID=" + encodeURIComponent(InvoiceID) + "&GLARAccount=" + encodeURIComponent(GLARAccount) + "&GLAccount=" + encodeURIComponent(GLAccount),
					
					success: function(response)
					 {
						if (response.startsWith("Error:")) {				
							swal({
								title: 'Error Saving Invoice Changes',
								text: response,
								type: 'error'
							})
							return;
						} else {
							$("#frmSageInvoiceDateRange").submit();				
						}
					 },
					failure: function(response)
					 {
						swal({
							title: 'Error Saving Invoice Changes',
							text: response,
							type: 'error'
						})
		             }
				});


	    	});

	   });	
	   </script>		    
	
	<div class="modal-header">
		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		<h4 class="modal-title" id="modalEditSageInvoiceTitle"><i class="fa fa-pencil" aria-hidden="true"></i> Edit Invoice #<%= InvoiceID %> for Line Item <%= LineItemNo %></h4>
	</div>
	
	<div class="modal-body modalResponsiveTable">

     	<div class="row modalrow">
     	   	<div class="col-lg-4">GL AR Account</div>
         	<div class="col-lg-8">
				<input type="text" id="txtGLARAccount" name="txtGLARAccount" class="form-control" value="<%= GLARAccount %>">
			</div>
		</div>

     	<div class="row modalrow">
     	   	<div class="col-lg-4">GL Account</div>
         	<div class="col-lg-8">
				<input type="text" id="txtGLAccount" name="txtGLAccount" class="form-control" value="<%= GLAccount %>">
			</div>
		</div>
		
     	<div class="row" style="margin-top:20px">
     	   	<div class="col-lg-6">&nbsp;</div>
         	<div class="col-lg-6 pull-right">
				<button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
				<button type="button" class="btn btn-primary" id="btnEditSageInvoiceSave" name="btnEditSageInvoiceSave"><i class="fa fa-floppy-o" aria-hidden="true"></i>&nbsp;Save Invoice Changes</button>
			</div>
		</div>

	</div>
<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub SaveEditSageInvoice() 

	InvoiceID = Request.Form("InvoiceID")
	LineItemNo = Request.Form("LineItemNo")
	GLARAccount = Request.Form("GLARAccount")
	GLAccount = Request.Form("GLAccount")	
	
	Set cnnSaveEditSageInvoice = Server.CreateObject("ADODB.Connection")
	cnnSaveEditSageInvoice.open (Session("ClientCnnString"))
	Set rsSaveEditSageInvoice = Server.CreateObject("ADODB.Recordset")
	rsSaveEditSageInvoice.CursorLocation = 3 

	
	'**********************************************************************
	'Lookup the record as it exists now so we can fill in the audit trail
	'**********************************************************************
	
	SQL = "SELECT * FROM IN_InvoiceHistDetail where InvoiceID = '" & InvoiceID & "' AND LineNumber = " & LineItemNo
	
	'Response.Write(SQL & "<br><br>")
		
	Set rsSaveEditSageInvoice = cnnSaveEditSageInvoice.Execute(SQL)
		
	If not rsSaveEditSageInvoice.EOF Then
		IntRecID = rsSaveEditSageInvoice("InternalRecordIdentifier")
		ORIG_GLARAccount = rsSaveEditSageInvoice("GL_AR_Account")
		ORIG_GLAccount = rsSaveEditSageInvoice("GL_Account")
	End If
	
	'**********************************************************************
	'End Lookup the record as it exists now so we can fill in the audit trail
	'**********************************************************************
	
	
	
	'**********************************************************************
	'Now Update IN_InvoiceExportSage with edited values from modal
	'**********************************************************************
	SQL = "UPDATE IN_InvoiceHistDetail SET "
	SQL = SQL &  " GL_AR_Account= '" & GLARAccount & "' "
	SQL = SQL &  ", GL_Account = '" & GLAccount & "' "
	SQL = SQL &  " WHERE InternalRecordIdentifier = " & IntRecID
	
	'Response.Write(SQL & "<br><br>")
	
	Set rsSaveEditSageInvoice = cnnSaveEditSageInvoice.Execute(SQL)
	
	set rsSaveEditSageInvoice = Nothing
	
	
	'**********************************************************************
	'Create Audit Log Entries For Any Specific Changes Made
	'**********************************************************************
	
	Description = ""

	If ORIG_GLARAccount <> GLARAccount Then
		Description = "Sage Invoice GL AR Account changed from " & Orig_GLARAccount & " to " & GLARAccount
		CreateAuditLogEntry "Sage Invoice Edited","Sage Invoice Edited","Minor",0,Description
	End If

	If ORIG_GLAccount <> GLAccount Then
		Description = "Sage Invoice GL Account changed from " & Orig_GLARAccount & " to " & GLARAccount
		CreateAuditLogEntry "Sage Invoice Edited","Sage Invoice Edited","Minor",0,Description
	End If
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetContentForSageInvoiceDetailExpansion() 

	InvoiceID = Request.Form("InvoiceID")
	
	'***************************************************************************************
	'Get values for editing an existing web order
	'***************************************************************************************

	SQLSageInvoiceDetailExpansion = "SELECT * FROM IN_InvoiceHistHeader WHERE InvoiceID = '" & InvoiceID & "'"
	
	Set cnnSageInvoiceDetailExpansion = Server.CreateObject("ADODB.Connection")
	cnnSageInvoiceDetailExpansion.open (Session("ClientCnnString"))
	Set rsSageInvoiceDetailExpansion = Server.CreateObject("ADODB.Recordset")
	Set rsSageInvoiceDetailExpansion = cnnSageInvoiceDetailExpansion.Execute(SQLSageInvoiceDetailExpansion)

		
	If Not rsSageInvoiceDetailExpansion.EOF Then

		InternalRecordIdentifier = rsSageInvoiceDetailExpansion("InternalRecordIdentifier")
		InvoiceID = rsSageInvoiceDetailExpansion("InvoiceID")
		CustID = rsSageInvoiceDetailExpansion("CustID")
		AlternateCustID = rsSageInvoiceDetailExpansion("AlternateCustID")
		InvoiceType = rsSageInvoiceDetailExpansion("InvoiceType")
		InvoiceGrandTotal = rsSageInvoiceDetailExpansion("InvoiceGrandTotal")
		InvoiceCreationDate = rsSageInvoiceDetailExpansion("InvoiceCreationDate")
		InvoiceDueDate = rsSageInvoiceDetailExpansion("InvoiceDueDate")
		
		BillToName = rsSageInvoiceDetailExpansion("BillToName")
	    BillToAddr1 = rsSageInvoiceDetailExpansion("BillToAddr1")
	    BillToAddr2 = rsSageInvoiceDetailExpansion("BillToAddr2")
	    BillToCity = rsSageInvoiceDetailExpansion("BillToCity")
	    BillToState = rsSageInvoiceDetailExpansion("BillToState")
	    BillToPostalCode = rsSageInvoiceDetailExpansion("BillToPostalCode")
	    BillToContact = rsSageInvoiceDetailExpansion("BillToContact")
	    BillToDescription = rsSageInvoiceDetailExpansion("BillToDescription")
	    
		ShipToName = rsSageInvoiceDetailExpansion("ShipToName")
	    ShipToAddr1 = rsSageInvoiceDetailExpansion("ShipToAddr1")
	    ShipToAddr2 = rsSageInvoiceDetailExpansion("ShipToAddr2")
	    ShipToCity = rsSageInvoiceDetailExpansion("ShipToCity")
	    ShipToState = rsSageInvoiceDetailExpansion("ShipToState")
	    ShipToPostalCode = rsSageInvoiceDetailExpansion("ShipToPostalCode")
	    ShipToContact = rsSageInvoiceDetailExpansion("ShipToContact")
	    ShipToDescription = rsSageInvoiceDetailExpansion("ShipToDescription")


	End If
	set rsSageInvoiceDetailExpansion = Nothing
	cnnSageInvoiceDetailExpansion.close
	set cnnSageInvoiceDetailExpansion = Nothing
	
	'***************************************************************************************
	
	
%>
	
	<div class="container" style="background-color:#FFF;">
	    <div class="row">
	        <div class="col-xs-12">
	    		<div class="invoice-title">
	    			<h2>Invoice <%= InvoiceID %></h2><h2 class="pull-right"><%= InvoiceType %></h2>
	    		</div>
	    		<hr>
	    		<div class="row">
	    			<div class="col-xs-6">
	    				<address>
	    				<strong>Billed To:</strong><br>
	    					<%= BillToName %><br>
	    					<%= BillToAddr1 %><br>
	    					<%= BillToAddr2 %><br>
	    					<%= BillToCity %>,&nbsp;<%= BillToState %>&nbsp;<%= BillToPostalCode %><br>
	    					<%= BillToContact %><br>
	    					<%= BillToDescription  %>
	    				</address>
	    			</div>
	    			<div class="col-xs-6 text-right">
	    				<address>
	        			<strong>Shipped To:</strong><br>
	    					<%= ShipToName %><br>
	    					<%= ShipToAddr1 %><br>
	    					<%= ShipToAddr2 %><br>
	    					<%= ShipToCity %>,&nbsp;<%= ShipToState %>&nbsp;<%= ShipToPostalCode %><br>
	    					<%= ShipToContact %><br>
	    					<%= ShipToDescription %>
	    				</address>
	    			</div>
	    		</div>
	    		<div class="row">
	    			<div class="col-xs-6">
	    				<address>
	    					<strong>Customer Information:</strong><br>
	    					<%= GetCustNameByCustNum(CustID) %><br>
	    					<strong>CustID:</strong>&nbsp;<%= CustID %> <br>
	    					<strong>Alternate ID:</strong>&nbsp;<%= AlternateCustID %><br>
	    				</address>
	    			</div>
	    			<div class="col-xs-3 text-right">
	    				<address>
		    				<% 
		    				If IsDate(InvoiceCreationDate) Then RndDate = InvoiceCreationDate Else RndDate = ""
					
							If RndDate <> "" Then
								eYear = Year(RndDate)
								If Month(RndDate) < 10 Then eMonth = "0" & Month(RndDate) else eMonth = Month(RndDate)
								If Day(RndDate) < 10 Then eDay = "0" & Day(RndDate) else eDay = Day(RndDate)
								DispayableDate = eMonth & "/" & eDay  & "/" & eYear
								DispayableDate  = cDate(DispayableDate) 
							End If
							%>
						
	    					<strong>Invoice Date:</strong><br>
	    					<%= Left(DispayableDate,Len(DispayableDate)-4) %><%= Right(DispayableDate,4) %><br><br>
	    				</address>
	    			</div>
	    			<div class="col-xs-3 text-right">
	    				<address>
	    					<strong>Due Date:</strong><br>
	    					<% If IsDate(InvoiceDueDate) Then
	    						Response.Write(formatDateTime(InvoiceDueDate,2))
	    					End If %>
	    					<br><br>
	    				</address>
	    			</div>
	    		</div>
	    	</div>
	    </div>

	
		<%
		'***************************************************************************************
		'Get values for editing an existing web order
		'***************************************************************************************
	
		SQLSageInvoiceDetailExpansion = "SELECT * FROM IN_InvoiceHistDetail WHERE InvoiceID = '" & InvoiceID & "' ORDER BY InternalRecordIdentifier"
		
		Set cnnSageInvoiceDetailExpansion = Server.CreateObject("ADODB.Connection")
		cnnSageInvoiceDetailExpansion.open (Session("ClientCnnString"))
		Set rsSageInvoiceDetailExpansion = Server.CreateObject("ADODB.Recordset")
		Set rsSageInvoiceDetailExpansion = cnnSageInvoiceDetailExpansion.Execute(SQLSageInvoiceDetailExpansion)
	
			
		%>	    
	    <div class="row">
	    	<div class="col-md-12">
	    		<div class="panel panel-default">
	    			<div class="panel-heading">
	    				<h3 class="panel-title"><strong>Order Details</strong></h3>
	    			</div>
	    			<div class="panel-body">
	    				<div class="table-responsive">
	    					<table class="table table-condensed">
	    						<thead>
	                                <tr>
	        							<td><strong>GL_AR_Account</strong></td>
	        							<td><strong>GL_Account</strong></td>
	        							<td><strong>Prod ID</strong></td>
	        							<td><strong>Description</strong></td>
	        							<td class="text-center"><strong>Price</strong></td>
	        							<td class="text-center"><strong>Qty Shipped</strong></td>
	        							<td class="text-center"><strong>Taxable</strong></td>
	        							<td class="text-center"><strong>Tax</strong></td>
	        							<td class="text-right"><strong>Line Total</strong></td>
	                                </tr>
	    						</thead>
	    						<tbody>
	    							<!-- foreach ($order->lineItems as $line) or some such thing here -->
	    							<%
	    							If Not rsSageInvoiceDetailExpansion.EOF Then
	    							
	    								InvoiceSubtotal = 0
	    								TotalInvoiceTax = 0
	
										Do While Not rsSageInvoiceDetailExpansion.EOF
										
											LineNumber = rsSageInvoiceDetailExpansion("LineNumber")
											prodSKU = rsSageInvoiceDetailExpansion("prodSKU")
											prodDescription = rsSageInvoiceDetailExpansion("prodDescription")
											PricePerUnitSold = rsSageInvoiceDetailExpansion("PricePerUnitSold")
										    QtyShipped = rsSageInvoiceDetailExpansion("QtyShipped")
										    If rsSageInvoiceDetailExpansion("Taxable") Then Taxable = "Y" Else Taxable = "N"  
										    GL_AR_Account = rsSageInvoiceDetailExpansion("GL_AR_Account")
										    GL_Account = rsSageInvoiceDetailExpansion("GL_Account")
										    If IsNumeric(rsSageInvoiceDetailExpansion("TotalTaxForLine")) Then 
												TotalTaxForLine = rsSageInvoiceDetailExpansion("TotalTaxForLine")
											Else
												TotalTaxForLine = 0
											End If
											%>

			    							<tr>
												<td><a data-toggle="modal" data-target="#modalEditSageInvoice" data-line-item-no="<%= LineNumber %>" data-invoice-id="<%= InvoiceID %>" class="btn btn-xs btn-success" style="cursor:pointer;"><i class="fa fa-pencil" aria-hidden="true"></i></a>&nbsp;<%= GL_AR_Account %></td>
			    								<td><%= GL_Account %></td>
			    								<td><%= prodSKU %></td>
			    								<td><%= prodDescription %></td>
			    								<td class="text-center"><%= formatCurrency(PricePerUnitSold,2) %></td>
			    								<td class="text-center"><%= QtyShipped %></td>
			    								<td class="text-center"><%= Taxable %></td>
			    								<td class="text-center"><%= formatCurrency(TotalTaxForLine,2) %></td>
			    								<td class="text-right"><%= formatCurrency(QtyShipped * PricePerUnitSold,2) %></td>
			    							</tr>
	    									<%
	    									InvoiceSubtotal = InvoiceSubtotal + (QtyShipped * PricePerUnitSold)
	    									rsSageInvoiceDetailExpansion.MoveNext
	    									TotalInvoiceTax = TotalInvoiceTax + TotalTaxForLine
	    								Loop
	    								
									End If
									
									set rsSageInvoiceDetailExpansion = Nothing
									cnnSageInvoiceDetailExpansion.close
									set cnnSageInvoiceDetailExpansion = Nothing
									
									'***************************************************************************************
									%>

	    							<tr>
	    								<td class="thick-line"></td>
	    								<td class="thick-line"></td>
	    								<td class="thick-line"></td>
	    								<td class="thick-line"></td>
	    								<td class="thick-line"></td>
	    								<td class="thick-line"></td>
	    								<td class="thick-line"></td>	    								
	    								<td class="thick-line text-center"><strong>Subtotal</strong></td>
	    								<td class="thick-line text-right"><%= formatCurrency(InvoiceSubtotal,2) %></td>
	    							</tr>
	    							<tr>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>	    								
	    								<td class="no-line text-center"><strong>Tax</strong></td>
	    								<td class="no-line text-right"><%= formatCurrency(TotalInvoiceTax,2) %></td>
	    							</tr>

	    							<tr>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>	    								
	    								<td class="no-line text-center"><strong>Total</strong></td>
	    								<td class="no-line text-right"><%= formatCurrency(InvoiceSubtotal+TotalInvoiceTax,2) %></td>
	    							</tr>
	    						</tbody>
	    					</table>
	    				</div>
	    			</div>
	    		</div>
	    	</div>
	    </div>
	</div>
	


<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub SaveEquivalentCustomerAccount()

	equivAccountEnteredByUser = Request.Form("equivID") 
	equivAccountEnteredByUser = Replace(equivAccountEnteredByUser, "'", "''")
	
	ourCustAccountIdentifyingInfoForSQL = Request.Form("id")

	ourCustAccountIdentifyingInfoForSQLArray = Split(ourCustAccountIdentifyingInfoForSQL,"*")
	
	ourCustID = ourCustAccountIdentifyingInfoForSQLArray(1)
	partnerRecID = ourCustAccountIdentifyingInfoForSQLArray(2)

	
	Set rsSaveEquivCustID = Server.CreateObject("ADODB.Recordset")
	rsSaveEquivCustID.CursorLocation = 3 
	
	If equivAccountEnteredByUser <> "" AND partnerRecID <> "" AND ourCustID <> "" Then
	
		SQLSaveEquivCustID = "SELECT * FROM AR_CustomerMapping WHERE "
		SQLSaveEquivCustID = SQLSaveEquivCustID & "partnerRecID = " & partnerRecID & " AND "
		SQLSaveEquivCustID = SQLSaveEquivCustID & "ourCustID = '" & ourCustID & "'"		
		Set cnnSaveEquivCustID = Server.CreateObject("ADODB.Connection")
		cnnSaveEquivCustID.open (Session("ClientCnnString"))
		Set rsSaveEquivCustID = cnnSaveEquivCustID.Execute(SQLSaveEquivCustID)
		
		If NOT rsSaveEquivCustID.EOF Then
		

			SQLUpdate = "UPDATE AR_CustomerMapping SET partnerCustID = '" & equivAccountEnteredByUser & "', RecordSource='Insight' WHERE "
			SQLUpdate = SQLUpdate & "partnerRecID = " & partnerRecID & " AND "
			SQLUpdate = SQLUpdate & "ourCustID = '" & ourCustID & "'"
			
			Set cnnUpdate = Server.CreateObject("ADODB.Connection")
			cnnUpdate.open (Session("ClientCnnString"))
			Set rsUpdate = Server.CreateObject("ADODB.Recordset")
			rsUpdate.CursorLocation = 3 
			Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
			cnnUpdate.close
			
			'Response.Write(SQLUpdate)
			Response.Write("Success")
				
		Else


			SQLInsert = "INSERT INTO AR_CustomerMapping (partnerRecID, partnerCustID, ourCustID,RecordSource) VALUES "
			SQLInsert = SQLInsert & " (" & partnerRecID & ",'" & equivAccountEnteredByUser & "','" & ourCustID & "','Insight')"
	
			Set cnnInsert = Server.CreateObject("ADODB.Connection")
			cnnInsert.open (Session("ClientCnnString"))
			Set rsInsert = Server.CreateObject("ADODB.Recordset")
			rsInsert.CursorLocation = 3 
			Set rsInsert = cnnInsert.Execute(SQLInsert)
			cnnInsert.close
			
			'Response.Write(SQLSaveEquivCustID)
			'Response.Write(SQLInsert)
			Response.Write("Success")
			
		End If
		
		set rsSaveEquivCustID = Nothing
		cnnSaveEquivCustID.close
		set cnnSaveEquivCustID = Nothing	

		Call ConvertCatchallAccountTransactions(ourCustID,equivAccountEnteredByUser,partnerRecID)
				
	Else
		Response.Write("Cannot Save, Invalid Data")
		
	End If

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub ConvertCatchallAccountTransactions(passedOurCustAccount,passedEquivCustAccount,passedPartnerRecId)

	baseURL = Request.ServerVariables("SERVER_NAME")
	If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

	'1. Will run update the API Order & Invoice & RA tables changing any records from the catchall account to the new account number
	'2. Will post the Order & Invoice numbers and equiv account number info the the Backend so it can be processed there as wll
'	Dim OrderArray(1)
'	Dim InvoiceArray(1)

	ourCustID = passedOurCustAccount
	equivAccount = passedEquivCustAccount
	partnerRecID = passedPartnerRecId

	Set rsPartner  = Server.CreateObject("ADODB.Recordset")
	rsPartner.CursorLocation = 3 
	
	If equivAccount <> "" AND partnerRecID <> "" AND ourCustID <> "" Then
	
		SQLPartner = "SELECT * FROM IC_Partners WHERE "
		rsPartner= SQLPartner & "InternalRecordIdentifier = " & partnerRecID 

		Set cnnPartner = Server.CreateObject("ADODB.Connection")
		cnnPartner.open (Session("ClientCnnString"))
		Set rsPartner = cnnPartner.Execute(rsPartner)
		
		If NOT rsPartner.EOF Then
		
			CatchAllAccount = rsPartner("partnerUnmappedCustomerID")
			PartnerAPIKey = rsPartner("PartnerAPIKey")
			
			'First select all the appropriate unmapped orders to build the order array for
			'passing to the backend
			
			'************
			' O R D E R S 
			'************

			Set cnnUpdate = Server.CreateObject("ADODB.Connection")
			cnnUpdate.open (Session("ClientCnnString"))
			Set rsUpdate = Server.CreateObject("ADODB.Recordset")
			rsUpdate.CursorLocation = 3 
			Set rsTransaction = Server.CreateObject("ADODB.Recordset")
			rsTransaction.CursorLocation = 3 
			
			' This just to get the count
			SQLUpdate = "SELECT Count(*) AS OrderCount FROM API_OR_OrderHeader WHERE "
			SQLUpdate = SQLUpdate & "APIKey = '" & PartnerAPIKey & "' AND "
			SQLUpdate = SQLUpdate & "Orig_CustID = '" & equivAccount & "'"
			
			Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
			
			If Not rsUpdate.Eof Then OrderCount = cint(rsUpdate("OrderCount")-1)
			
			
			SQLUpdate = "SELECT * FROM API_OR_OrderHeader WHERE "
			SQLUpdate = "SELECT DISTINCT OrderID FROM API_OR_OrderHeader WHERE "
			SQLUpdate = SQLUpdate & "APIKey = '" & PartnerAPIKey & "' AND "
			SQLUpdate = SQLUpdate & "Orig_CustID = '" & equivAccount & "'"


			Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
			
			If Not rsUpdate.Eof Then
			
				ReDim OrderArray(OrderCount)
				OrdArrayElement = 0
				
				Do While Not rsUpdate.Eof
				
					OrderArray(OrdArrayElement) = rsUpdate("OrderID")
				
					SQLTransaction = "UPDATE API_OR_OrderHeader "
					SQLTransaction = SQLTransaction & " SET CustId = '" & ourCustID  & "' "
					SQLTransaction = SQLTransaction & " WHERE OrderID = '" & OrderArray(OrdArrayElement) & "'"
					
					Set rsTransaction = cnnUpdate.Execute(SQLTransaction)
									
					OrdArrayElement = OrdArrayElement +1
					
				 rsUpdate.MoveNext
				Loop
			
			End If
			
			Set rsTransaction = Nothing
			Set rsUpdate = Nothing
			cnnUpdate.Close
			Set cnnUpdate = Nothing

			'****************
			' I N V O I C E S 
			'****************

			Set cnnUpdate = Server.CreateObject("ADODB.Connection")
			cnnUpdate.open (Session("ClientCnnString"))
			Set rsUpdate = Server.CreateObject("ADODB.Recordset")
			rsUpdate.CursorLocation = 3 
			Set rsTransaction = Server.CreateObject("ADODB.Recordset")
			rsTransaction.CursorLocation = 3 
			
			'This just to get the count
			SQLUpdate = "SELECT Count(*) AS InvoiceCount FROM API_IN_InvoiceHeader WHERE "
			SQLUpdate = SQLUpdate & "APIKey = '" & PartnerAPIKey & "' AND "
			SQLUpdate = SQLUpdate & "Orig_CustID = '" & equivAccount & "'"

			Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
			
			If Not rsUpdate.Eof Then InvoiceCount = cint(rsUpdate("InvoiceCount")-1)

			SQLUpdate = "SELECT * FROM API_IN_InvoiceHeader WHERE "
			SQLUpdate = "SELECT DISTINCT InvoiceID FROM API_IN_InvoiceHeader WHERE "
			SQLUpdate = SQLUpdate & "APIKey = '" & PartnerAPIKey & "' AND "
			SQLUpdate = SQLUpdate & "Orig_CustID = '" & equivAccount & "'"


			Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
			
			If Not rsUpdate.Eof Then
			
				ReDim InvoiceArray(InvoiceCount)
				InvArrayElement = 0


				InvArrayElement = 0
				
				Do While Not rsUpdate.Eof
				
					InvoiceArray(InvArrayElement) = rsUpdate("InvoiceID")
				
					SQLTransaction = "UPDATE API_IN_InvoiceHeader "
					SQLTransaction = SQLTransaction & " SET CustId = '" & ourCustID  & "' "
					SQLTransaction = SQLTransaction & " WHERE InvoiceID = '" & InvoiceArray(InvArrayElement) & "'"
					
					Set rsTransaction = cnnUpdate.Execute(SQLTransaction)
									
					InvArrayElement = InvArrayElement +1
					
				 rsUpdate.MoveNext
				Loop
			
			End If
			
			Set rsTransaction = Nothing
			Set rsUpdate = Nothing
			cnnUpdate.Close
			Set cnnUpdate = Nothing

			
			'******
			' R A s
			'******

			Set cnnUpdate = Server.CreateObject("ADODB.Connection")
			cnnUpdate.open (Session("ClientCnnString"))
			Set rsUpdate = Server.CreateObject("ADODB.Recordset")
			rsUpdate.CursorLocation = 3 
			Set rsTransaction = Server.CreateObject("ADODB.Recordset")
			rsTransaction.CursorLocation = 3 
			
			' This just to get the count
			SQLUpdate = "SELECT Count(*) AS RACount FROM API_OR_RAHeader WHERE "
			SQLUpdate = SQLUpdate & "APIKey = '" & PartnerAPIKey & "' AND "
			SQLUpdate = SQLUpdate & "Orig_CustID = '" & equivAccount & "'"
			
			Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
			
			If Not rsUpdate.Eof Then RACount = cint(rsUpdate("RACount")-1)
			
			
			SQLUpdate = "SELECT * FROM API_OR_RAHeader WHERE "
			SQLUpdate = "SELECT DISTINCT RAID FROM API_OR_RAHeader WHERE "
			SQLUpdate = SQLUpdate & "APIKey = '" & PartnerAPIKey & "' AND "
			SQLUpdate = SQLUpdate & "Orig_CustID = '" & equivAccount & "'"


			Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
			
			If Not rsUpdate.Eof Then
			
				ReDim RAArray(RACount)
				RAArrayElement = 0
				
				Do While Not rsUpdate.Eof
				
					RAArray(RAArrayElement) = rsUpdate("RAID")
				
					SQLTransaction = "UPDATE API_OR_RAHeader "
					SQLTransaction = SQLTransaction & " SET CustId = '" & ourCustID  & "' "
					SQLTransaction = SQLTransaction & " WHERE RAID = '" & RAArray(RAArrayElement) & "'"
					
					Set rsTransaction = cnnUpdate.Execute(SQLTransaction)
									
					RAArrayElement = RAArrayElement+1
					
				 rsUpdate.MoveNext
				Loop
			
			End If
			
			Set rsTransaction = Nothing
			Set rsUpdate = Nothing
			cnnUpdate.Close
			Set cnnUpdate = Nothing
			Set rsPartner = Nothing
			cnnPartner.Close
			Set rsPartner = Nothing


			' NOW  POST  ALL  THIS  STUFF  TO  THE  BACKEND 
			
			'************
			' O R D E R S 
			'************
			If IsArray(OrderArray) Then
			
				'Construct xml fields based on record
				xmlData = "<DATASTREAM>"
				xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
				
				xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTORDERMODE") & "</MODE>"


				xmlData = xmlData & "<RECORD_TYPE>ORDER</RECORD_TYPE>"
			
				xmlData = xmlData & "<RECORD_SUBTYPE>CHANGE_CUSTOMER</RECORD_SUBTYPE>"
	
				xmlData = xmlData & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
				
				xmlData = xmlData & "<ORDERS>"
				
				For x = 0 to Ubound(OrderArray)
					If Len(OrderArray(x)) > 0 Then
						xmlData = xmlData & "<ORDER>"
						xmlData = xmlData & "<ORDER_ID>" & OrderArray(x) & "</ORDER_ID>"
						xmlData = xmlData & "<CUST_ID>" & ourCustID & "</CUST_ID>"	
						xmlData = xmlData & "</ORDER>"
					End If
				Next
					
				xmlData = xmlData & "</ORDERS>"
				
				xmlData = xmlData & "</DATASTREAM>"
				
				xmlDataForDisp = Replace(xmlData,"<","[")
				xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
				xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
				xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
				xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")
		
		
				Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
				'setTimeouts (long resolveTimeout, long connectTimeout, long sendTimeout, long receiveTimeout)
				httpRequest.SetTimeouts 1200000, 1200000, 1200000, 1200000
				httpRequest.Open "POST", GetAPIRepostURL(), True

				httpRequest.SetRequestHeader "Content-Type", "text/xml"
				'httpRequest.SetRequestHeader "accept-encoding", "gzip, deflate"
				
				xmlData = Replace(xmlData,"&","&amp;")
				xmlData = Replace(xmlData,chr(34),"")			
				httpRequest.Send xmlData

				While (httpRequest.readyState<>4)
                	httpRequest.waitForResponse(10000)      
                Wend
                
                                    
				data = xmlData
			
				If (httpRequest.readyState = 4) Then
					
					If (Err.Number <> 0 ) Then
						emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
						emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
						emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
						emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
						emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
						emailBody = emailBody & "SERNO: " & SERNO & "<br>"
						SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CHANGE_CUSTOMER",emailBody, "Order API", "Order API"
					
						Description = emailBody 
						'CreateSystemAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTORDERMODE"),"1071d","1071d","Order API"
					End If
		
					If httpRequest.status = 200 THEN 
					
						If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
					
							emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
							emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
							emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
							emailBody = emailBody & "SERNO: " & SERNO & "<br>"
							SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost CHANGE_CUSTOMER",emailBody, "Order API", "Order API"
							
							'Call CreateSystemAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
							
						Else
							'FAILURE
							emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
							emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
							emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
							emailBody = emailBody & "SERNO: " & SERNO & "<br>"
							SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CHANGE_CUSTOMER",emailBody, "Order API", "Order API"
						
							'Call CreateSystemAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
							
						End If
						
					Else
					
							'FAILURE
							emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
							emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
							emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
							emailBody = emailBody & "SERNO: " & SERNO & "<br>"
							SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CHANGE_CUSTOMER",emailBody, "Order API", "Order API"
						
							'Call CreateSystemAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
				
					End If
					
				End If
			
			End If
				

			'****************
			' I N V O I C E S  
			'****************
			If IsArray(InvoiceArray) Then
			
				'Construct xml fields based on record
				xmlData = "<DATASTREAM>"
				xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
				
				xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTINVOICEMODE") & "</MODE>"
				'xmlData = xmlData & "<MODE>TEST</MODE>"
				
				xmlData = xmlData & "<RECORD_TYPE>INVOICE</RECORD_TYPE>"
				xmlData = xmlData & "<RECORD_SUBTYPE>CHANGE_CUSTOMER</RECORD_SUBTYPE>"
				xmlData = xmlData & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
				
				xmlData = xmlData & "<INVOICES>"
				
				For x = 0 to Ubound(InvoiceArray)
					If Len(InvoiceArray(x)) > 0 Then
						xmlData = xmlData & "<INVOICE>"
						xmlData = xmlData & "<INVOICE_ID>" & InvoiceArray(x) & "</INVOICE_ID>"
						xmlData = xmlData & "<CUST_ID>" & ourCustID & "</CUST_ID>"	
						xmlData = xmlData & "</INVOICE>"
					End If
				Next
					
				xmlData = xmlData & "</INVOICES>"
				
				xmlData = xmlData & "</DATASTREAM>"
				
				xmlDataForDisp = Replace(xmlData,"<","[")
				xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
				xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
				xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
				xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")
		
		
				Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
				'setTimeouts (long resolveTimeout, long connectTimeout, long sendTimeout, long receiveTimeout)
				httpRequest.SetTimeouts 1200000, 1200000, 1200000, 1200000
				httpRequest.Open "POST", GetAPIRepostInvoicesURL(), True
			'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				httpRequest.SetRequestHeader "Content-Type", "text/xml"
				
				xmlData = Replace(xmlData,"&","&amp;")
				xmlData = Replace(xmlData,chr(34),"")			
				httpRequest.Send xmlData

				While (httpRequest.readyState<>4)
                	httpRequest.waitForResponse(10000)      
                Wend
			
				data = xmlData
				
				If (httpRequest.readyState = 4) Then
				
					If (Err.Number <> 0 ) Then
						emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
						emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
						emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
						emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
						emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
						emailBody = emailBody & "SERNO: " & SERNO & "<br>"
						SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CHANGE_CUSTOMER",emailBody, "Invoice API", "Invoice API"
					
						Description = emailBody 
						'CreateSystemAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTINVOICEMODE"),"1071d","1071d","Order API"
					End If
		
					If httpRequest.status = 200 THEN 
					
						If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
					
							emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
							emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
							emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
							emailBody = emailBody & "SERNO: " & SERNO & "<br>"
							SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost CHANGE_CUSTOMER",emailBody, "Invoice API", "Invoice API"
							
							'Call CreateSystemAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTINVOICEMODE"),"rePostings.asp")
							
						Else
							'FAILURE
							emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
							emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
							emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
							emailBody = emailBody & "SERNO: " & SERNO & "<br>"
							SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CHANGE_CUSTOMER",emailBody, "Invoice API", "Invoice API"
						
							'Call CreateSystemAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTINVOICEMODE"),"rePostings.asp")
							
						End If
						
					Else
					
							'FAILURE
							emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
							emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
							emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
							emailBody = emailBody & "SERNO: " & SERNO & "<br>"
							SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CHANGE_CUSTOMER",emailBody, "Invoice API", "Invoice API"
						
							'Call CreateSystemAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTINVOICEMODE"),"rePostings.asp")
				
					End If
					
				End If
				
			End If
			
			
			'************
			' R A s 
			'************
			If IsArray(RAArray) Then
			
				'Construct xml fields based on record
				xmlData = "<DATASTREAM>"
				xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
				
				xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTRAMODE") & "</MODE>"
				'xmlData = xmlData & "<MODE>TEST</MODE>"

				xmlData = xmlData & "<RECORD_TYPE>RETURN_AUTHORIZATION</RECORD_TYPE>"
			
				xmlData = xmlData & "<RECORD_SUBTYPE>CHANGE_CUSTOMER</RECORD_SUBTYPE>"
	
				xmlData = xmlData & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
				
				xmlData = xmlData & "<RETURN_AUTHORIZATIONS>"
				
				For x = 0 to Ubound(RAArray)
					If Len(RAArray(x)) > 0 Then
						xmlData = xmlData & "<RETURN_AUTHORIZATION>"
						xmlData = xmlData & "<RA_ID>" & RAArray(x) & "</RA_ID>"
						xmlData = xmlData & "<CUST_ID>" & ourCustID & "</CUST_ID>"	
						xmlData = xmlData & "</RETURN_AUTHORIZATION>"
					End If
				Next
					
				xmlData = xmlData & "</RETURN_AUTHORIZATIONS>"
				
				xmlData = xmlData & "</DATASTREAM>"
				
				xmlDataForDisp = Replace(xmlData,"<","[")
				xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
				xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
				xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
				xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")
		
		
				Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
				
				'setTimeouts (long resolveTimeout, long connectTimeout, long sendTimeout, long receiveTimeout)
				httpRequest.SetTimeouts 1200000, 1200000, 1200000, 1200000				
				httpRequest.Open "POST", GetAPIRepostRAURL(), True
				
			'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				httpRequest.SetRequestHeader "Content-Type", "text/xml"
				
				xmlData = Replace(xmlData,"&","&amp;")
				xmlData = Replace(xmlData,chr(34),"")			
				httpRequest.Send xmlData

				While (httpRequest.readyState<>4)
                	httpRequest.waitForResponse(10000)      
                Wend
			
				data = xmlData
			
				
				If (httpRequest.readyState = 4) Then
					
					If (Err.Number <> 0 ) Then
						emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN_AUTHORIZATION and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
						emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
						emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
						emailBody = emailBody & "Posted to " & GetAPIRepostRAURL() & "<br><br>"
						emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
						emailBody = emailBody & "SERNO: " & GetPOSTParams("SERNO") & "<br>"
						SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CHANGE_CUSTOMER",emailBody, "Order API", "Order API"
					
						Description = emailBody 
						'CreateSystemAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTORDERMODE"),"1071d","1071d","Order API"
					End If
		
					If httpRequest.status = 200 THEN 
					
						If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
					
							emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN_AUTHORIZATION and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
							emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & GetAPIRepostRAURL() & "<br><br>"
							emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
							emailBody = emailBody & "SERNO: " & GetPOSTParams("SERNO")& "<br>"
							SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost CHANGE_CUSTOMER",emailBody, "Order API", "Order API"
							
							'Call CreateSystemAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
							
						Else
							'FAILURE
							emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN_AUTHORIZATION and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
							emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
							emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
							emailBody = emailBody & "SERNO: " & GetPOSTParams("SERNO")& "<br>"
							SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CHANGE_CUSTOMER",emailBody, "Order API", "Order API"
						
							'Call CreateSystemAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
							
						End If
						
					Else
					
							'FAILURE
							emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN_AUTHORIZATION and <RECORD_SUBTYPE>CHANGE_CUSTOMER"& "<br><br>"
							emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
							emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
							emailBody = emailBody & "SERNO: " & GetPOSTParams("SERNO")& "<br>"
							SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CHANGE_CUSTOMER",emailBody, "Order API", "Order API"
						
							'Call CreateSystemAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
				
					End If
				End If
			
			End If
				
			
		End If
	End If

End Sub



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForEmailConsolidatedInvoiceModalAccount() 

	ConsInvoiceNumber = Request.Form("consInvoiceNumber")
	CustID = Request.Form("custID")
	EndDate = Request.Form("endDate")
	PaidOrUnpaid = Request.Form("paidOrUnpaid")
		
%>

	<script type="text/javascript">
	
		$(document).ready(function() {

			$('#btnEmailConsolidatedInvoice').on('click', function(e) {
			
			    //get data-id attribute of the clicked alert
			    var consInvoiceNum = $("#txtConsInvNumber").val();
			    var custID = $("#txtCustID").val();
			    var endDate = $("#txtEndDate").val();
			    var emailto = $("#selEmailto").val();
			    var addlemails = $("#txtAdditionalEmails").val();
			    var paidOrUnpaid = $("#txtPaidOrUnpaid").val();
			    							    		    		
		    	$.ajax({
					type:"POST",
					url:"../../../inc/InSightFuncs_AjaxForARAP.asp",
					data: "action=EmailConsolidatedInvoiceModalAccount&consInvoiceNum=" + encodeURIComponent(consInvoiceNum) + "&custID=" + encodeURIComponent(custID) + "&endDate=" + encodeURIComponent(endDate) + "&emailto=" + encodeURIComponent(emailto) + "&addlemails=" + encodeURIComponent(addlemails) + "&paidOrUnpaid=" + encodeURIComponent(paidOrUnpaid),
					success: function(response)
					 {
					 	swal("Consolidated Invoice Emailed Successfully.")
		             }
				});
	    	});	
	    	
	    	    		
		});
	</script>

		<!-- email alert line !-->
		<div class="row" style="margin-bottom:20px;">

			<!-- email alert !-->
			<div class="col-lg-4">
				<% If PaidOrUnpaid = "UNPAID" Then %>
					<label class="right">Email Consolidated Invoice (Unpaid) To:</label>
				<% Else %>
					<label class="right">Email Consolidated Invoice To:</label>
				<% End If %>
			</div>
			<!-- eof email alert !-->

			<!-- multi select !-->
			<div class="col-lg-8">
				<select class="form-control multiselect" id="selEmailto" name="selEmailto" multiple style="min-height:200px;">
					<option value="0">--- none from here ---</option>
					<option value="<%=Session("UserNo")%>"><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option> 								
					<% 
			      	
					Set cnnDeliveryModal = Server.CreateObject("ADODB.Connection")
					cnnDeliveryModal.open (Session("ClientCnnString"))
					Set rsDeliveryModal = Server.CreateObject("ADODB.Recordset")
							      	
			      	 
		      	  	SQLDeliveryModal = "SELECT UserNo, userFirstName, userLastName FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
		      	  	SQLDeliveryModal = SQLDeliveryModal & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo")
		      	  	SQLDeliveryModal = SQLDeliveryModal & " ORDER BY  userFirstName, userLastName"
	
					rsDeliveryModal.CursorLocation = 3 
					Set rsDeliveryModal = cnnDeliveryModal.Execute(SQLDeliveryModal)
				
					If not rsDeliveryModal.EOF Then
						Do
							FullName = rsDeliveryModal("userFirstName") & " " & rsDeliveryModal("userLastName")
							Response.Write("<option value='" & rsDeliveryModal("UserNo") & "'>" & FullName & "</option>")
							rsDeliveryModal.movenext
							
						Loop until rsDeliveryModal.eof
					End If
					set rsDeliveryModal = Nothing
					cnnDeliveryModal.close
					set cnnDeliveryModal = Nothing
			      	%>
				</select>
				<strong>Use CTRL and SHIFT to make multiple selections</strong>
            </div>
			<!-- eof multi select !-->
        </div>
        <!-- eof email alert line !-->
        
        
		<!-- email alert line !-->
		<div class="row" style="margin-bottom:20px;">

			<!-- email alert !-->
			<div class="col-lg-4">
				<label class="right">Additional Emails:</label>
			</div>
			<!-- eof email alert !-->
    
            <!-- text area !-->
            <div class="col-lg-8">
				<textarea class="form-control textarea" rows="4" id="txtAdditionalEmails" name="txtAdditionalEmails"></textarea>
	            <strong>Separate multiple email addresses with a semicolon</strong>
            </div>
            <!-- eof text area !-->
        </div>
        <!-- eof email alert line !-->
        

	<!-- eof modal body !-->
      
	<!-- modal footer !-->
    <div class="modal-footer">
		      	      
		<!-- close / save !-->
		<div class="col-lg-12" style="margin-right:-25px;">
			<button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
			<% If PaidOrUnpaid = "UNPAID" Then %>
				<button type="button" id="btnEmailConsolidatedInvoice" class="btn btn-primary">Email Consolidated Invoice (Unpaid)</button>
			<% Else %>
				<button type="button" id="btnEmailConsolidatedInvoice" class="btn btn-primary">Email Consolidated Invoice</button>			
			<% End If %>
		</div>
		<!-- eof close / save !-->

	</div>
	<!-- eof modal footer !-->

<%
End Sub




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub EmailConsolidatedInvoiceModalAccount() 

	
	consInvoiceNum = Request.Form("consInvoiceNum")
	Account = Request.Form("custID")
	EndDate = Request.Form("endDate")
	Emailto = Request.Form("emailto") 
	AdditionalEmails = Request.Form("addlemails")
	PaidOrUnpaid = Request.Form("paidOrUnpaid")
		
	
	If AdditionalEmails <> "" Then
		AdditionalEmails = Trim(AdditionalEmails)
		AdditionalEmails = Replace(AdditionalEmails,",",";") ' Common for the user to type , instead of ; So we fix it
		If Right(AdditionalEmails,1)=";" Then AdditionalEmails = Left(AdditionalEmails,Len(AdditionalEmails)-1)
	End If
	
	'**************************************************************************************************
	' FIRST SAVE CONSOLIDATED INVOICE TO CLIENTFILES\CLIENTID\CUSTOMER\ACCOUNTSRECEIVABLE - 
	' MUST BE ABLE TO EMAIL ATTACHMENT FROM A PHYSICAL LOCATION
	'**************************************************************************************************

		'baseURL should always have a trailing /slash, just in case, handle either way
		If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
		sURL = Request.ServerVariables("SERVER_NAME")
	
		'Now change the name of the file
		Orig_Name =  "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_Account_" & Trim(Account) & "_" & Trim(Account) & Trim(Replace(EndDate,"/","")) & ".pdf"
		New_Name =  "\clientfilesV\" & Left(MUV_Read("ClientID"),4) & "\accountsreceivable\consolidated\ConsolidatedStatement_Account_" & Trim(Account) & "_" & Trim(Account) & Trim(Replace(EndDate,"/","")) & ".pdf"
		
		'Response.Write(Orig_Name & "<br>")
		'Response.Write(New_Name & "<br>")
		
		Dim fso
		
		Set fso = CreateObject("Scripting.FileSystemObject")
				
		'Kill it first in case an old one is there
		On error resume next
		fso.DeleteFile Server.MapPath(New_Name)
		On error goto 0
		
		fso.CopyFile Server.MapPath(Orig_Name), Server.MapPath(New_Name)
		
		Set fso = Nothing
	
	'**************************************************************************************************
	'END CODE TO SAVE CONSOLIDATED INVOICE TO CLIENTFILES\CLIENTID\CUSTOMER\ACCOUNTSRECEIVABLE
	'**************************************************************************************************

	
	
	'**************************************************************************************************
	' NOW EMAIL PDF TO ALL THE USERS AND ADDITIONAL EMAIL ADDRESSES SPECIFIED
	'**************************************************************************************************

		fn = Server.MapPath(New_Name)
		Response.Write("File to attach: " & fn & "<br>")

		'Send user based emails
		If Emailto <> "" Then
			UserNoList = Split(Emailto,",")
			For x = 0 To UBound(UserNoList)
				Send_To = Send_To & GetUserEmailByUserNo(UserNoList(x)) & ";"
			Next
		End If
		
		'Got all the addresses so now break them up
		Send_To_Array = Split(Send_To,";")
		
		'Response.Write("<br>Send_To: " & Send_To & "<br>")
		
		'HERE WE ACTUALLY SEND THE EMAIL
		For x = 0 to Ubound(Send_To_Array) -1
			Send_To = Send_To_Array(x)
	
			If PaidOrUnpaid = "UNPAID" Then
				emailSubject = "Consolidated Invoice (Unpaid) For Account " & Account & " " & GetCustNameByCustNum(Account)
			Else
				emailSubject = "Consolidated Invoice For Account " & Account & " " & GetCustNameByCustNum(Account)
			End If

			emailBody = ""
			'Failsafe for dev
			sURL = Request.ServerVariables("SERVER_NAME")
			If Instr(ucase(sURL),"DEV.") <> 0 Then Send_To = "rsmith@ocsaccess.com"
			If Instr(ucase(sURL),"DEV2.") <> 0 Then Send_To = "cgrecco@ocsaccess.com"
			emailBody = "The consolidate invoice for the customer is attached. (" & MUV_Read("ClientID") & ")"
			
			Response.Write("<br>FN: " & fn & "<br>")

			SendMailWatt "mailsender@" & maildomain,Send_To,emailSubject,emailBody,fn,GetTerm("Accounts Receivable"),"Consolidated Invoice Email"
			
			If PaidOrUnpaid = "UNPAID" Then
				CreateAuditLogEntry "Consolidated (Unpaid) Invoice Email","Consolidated Invoice (Unpaid) Email","Minor",0,"Consolidated Invoice (Unpaid) Email Sent to " & Send_To 
			Else
				CreateAuditLogEntry "Consolidated Invoice Email","Consolidated Invoice Email","Minor",0,"Consolidated Invoice Email Sent to " & Send_To
			End If
			Response.Write("Sent the email to " & Send_To & "<br>")
			Response.Write("Sent the email, all done<br>")
		Next 


		'Send additional emails
		If AdditionalEmails <> "" Then
			AddlEmailAddressList = Split(AdditionalEmails,";")
			For x = 0 To UBound(AddlEmailAddressList)
				Send_To = Send_To & AddlEmailAddressList(x) & ";"
			Next
		End If
		
		'Got all the addresses so now break them up
		Send_To_Array_Additional = Split(Send_To,";")
		
		'Response.Write("<br>Send_To: " & Send_To & "<br>")
		
		'HERE WE ACTUALLY SEND THE EMAIL
		For x = 0 to Ubound(Send_To_Array_Additional) -1
			Send_To = Send_To_Array_Additional(x)
	
			If PaidOrUnpaid = "UNPAID" Then
				emailSubject = "Consolidated Invoice (Unpaid) For Account " & Account & " " & GetCustNameByCustNum(Account)
			Else
				emailSubject = "Consolidated Invoice For Account " & Account & " " & GetCustNameByCustNum(Account)
			End If
			
			emailBody = ""
			'Failsafe for dev
			sURL = Request.ServerVariables("SERVER_NAME")
			If Instr(ucase(sURL),"DEV.") <> 0 Then Send_To = "rsmith@ocsaccess.com"
			If Instr(ucase(sURL),"DEV2.") <> 0 Then Send_To = "cgrecco@ocsaccess.com"
			emailBody = "The consolidate invoice for the customer is attached. (" & MUV_Read("ClientID") & ")"
			
			Response.Write("<br>FN: " & fn & "<br>")

			SendMailWatt "mailsender@" & maildomain,Send_To,emailSubject,emailBody,fn,GetTerm("Accounts Receivable"),"Consolidated Invoice Email"
			
			If PaidOrUnpaid = "UNPAID" Then
				CreateAuditLogEntry "Consolidated (Unpaid) Invoice Email","Consolidated Invoice (Unpaid) Email","Minor",0,"Consolidated Invoice (Unpaid) Email Sent to " & Send_To 
			Else
				CreateAuditLogEntry "Consolidated Invoice Email","Consolidated Invoice Email","Minor",0,"Consolidated Invoice Email Sent to " & Send_To
			End If
			
			Response.Write("Sent the email to " & Send_To & "<br>")
			Response.Write("Sent the email, all done<br>")
		Next	

	'**************************************************************************************************
	' END SEND EMAIL
	'**************************************************************************************************

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************






'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForEmailConsolidatedInvoiceModalChain() 

	ConsInvoiceNumber = Request.Form("consInvoiceNumber")
	ChainID = Request.Form("ChainID")
	ChainName = GetChainDescByChainNum(ChainID)
	EndDate = Request.Form("endDate")
	PaidOrUnpaid = Request.Form("type")
		
%>

	<script type="text/javascript">
	
		$(document).ready(function() {

			$('#btnEmailConsolidatedInvoice').on('click', function(e) {
			
			    //get data-id attribute of the clicked alert
			    var consInvoiceNum = $("#txtConsInvNumber").val();
			    var chainID = $("#txtChainID").val();
			    var endDate = $("#txtEndDate").val();
			    var emailto = $("#selEmailto").val();
			    var addlemails = $("#txtAdditionalEmails").val();
			    var paidOrUnpaid = $("#txtPaidOrUnpaid").val();
			    							    		    		
		    	$.ajax({
					type:"POST",
					url:"../../../inc/InSightFuncs_AjaxForARAP.asp",
					data: "action=EmailConsolidatedInvoiceModalChain&consInvoiceNumber=" + encodeURIComponent(consInvoiceNum) + "&chainID=" + encodeURIComponent(chainID) + "&endDate=" + encodeURIComponent(endDate) + "&emailto=" + encodeURIComponent(emailto) + "&addlemails=" + encodeURIComponent(addlemails) + "&paidOrUnpaid=" + encodeURIComponent(paidOrUnpaid),
					success: function(response)
					 {
					 	swal("Consolidated Invoice Emailed Successfully.");
		             }
				});
	    	});	
	    	
	    	    		
		});
	</script>

		<!-- email alert line !-->
		<div class="row" style="margin-bottom:20px;">

			<!-- email alert !-->
			<div class="col-lg-4">
				<% If PaidOrUnpaid = "UNPAID" Then %>
					<label class="right">Email Consolidated Invoice (Unpaid) To:</label>
				<% Else %>
					<label class="right">Email Consolidated Invoice To:</label>
				<% End If %>
			</div>
			<!-- eof email alert !-->

			<!-- multi select !-->
			<div class="col-lg-8">
				<select class="form-control multiselect" id="selEmailto" name="selEmailto" multiple style="min-height:200px;">
					<option value="0">--- none from here ---</option>
					<option value="<%=Session("UserNo")%>"><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option> 								
					<% 
			      	
					Set cnnDeliveryModal = Server.CreateObject("ADODB.Connection")
					cnnDeliveryModal.open (Session("ClientCnnString"))
					Set rsDeliveryModal = Server.CreateObject("ADODB.Recordset")
							      	
			      	 
		      	  	SQLDeliveryModal = "SELECT UserNo, userFirstName, userLastName FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
		      	  	SQLDeliveryModal = SQLDeliveryModal & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo")
		      	  	SQLDeliveryModal = SQLDeliveryModal & " ORDER BY  userFirstName, userLastName"
	
					rsDeliveryModal.CursorLocation = 3 
					Set rsDeliveryModal = cnnDeliveryModal.Execute(SQLDeliveryModal)
				
					If not rsDeliveryModal.EOF Then
						Do
							FullName = rsDeliveryModal("userFirstName") & " " & rsDeliveryModal("userLastName")
							Response.Write("<option value='" & rsDeliveryModal("UserNo") & "'>" & FullName & "</option>")
							rsDeliveryModal.movenext
							
						Loop until rsDeliveryModal.eof
					End If
					set rsDeliveryModal = Nothing
					cnnDeliveryModal.close
					set cnnDeliveryModal = Nothing
			      	%>
				</select>
				<strong>Use CTRL and SHIFT to make multiple selections</strong>
            </div>
			<!-- eof multi select !-->
        </div>
        <!-- eof email alert line !-->
        
        
		<!-- email alert line !-->
		<div class="row" style="margin-bottom:20px;">

			<!-- email alert !-->
			<div class="col-lg-4">
				<label class="right">Additional Emails:</label>
			</div>
			<!-- eof email alert !-->
    
            <!-- text area !-->
            <div class="col-lg-8">
				<textarea class="form-control textarea" rows="4" id="txtAdditionalEmails" name="txtAdditionalEmails"></textarea>
	            <strong>Separate multiple email addresses with a semicolon</strong>
            </div>
            <!-- eof text area !-->
        </div>
        <!-- eof email alert line !-->
        

	<!-- eof modal body !-->
      
	<!-- modal footer !-->
    <div class="modal-footer">
		      	      
		<!-- close / save !-->
		<div class="col-lg-12" style="margin-right:-25px;">
			<button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
			<% If PaidOrUnpaid = "UNPAID" Then %>
				<button type="button" id="btnEmailConsolidatedInvoice" class="btn btn-primary">Email Consolidated Invoice (Unpaid)</button>
			<% Else %>
				<button type="button" id="btnEmailConsolidatedInvoice" class="btn btn-primary">Email Consolidated Invoice</button>			
			<% End If %>

		</div>
		<!-- eof close / save !-->

	</div>
	<!-- eof modal footer !-->

<%
End Sub




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub EmailConsolidatedInvoiceModalChain() 

	
	consInvoiceNum = Request.Form("consInvoiceNumber")
	ChainID = Request.Form("ChainID")
	ChainName = GetChainDescByChainNum(ChainID)
	EndDate = Request.Form("endDate") 
	Emailto = Request.Form("emailto") 
	AdditionalEmails = Request.Form("addlemails")
	PaidOrUnpaid = Request.Form("paidOrUnpaid")
		
	If AdditionalEmails <> "" Then
		AdditionalEmails = Trim(AdditionalEmails)
		AdditionalEmails = Replace(AdditionalEmails,",",";") ' Common for the user to type , instead of ; So we fix it
		If Right(AdditionalEmails,1)=";" Then AdditionalEmails = Left(AdditionalEmails,Len(AdditionalEmails)-1)
	End If
	
	'**************************************************************************************************
	' FIRST SAVE CONSOLIDATED INVOICE TO CLIENTFILES\CLIENTID\CUSTOMER\ACCOUNTSRECEIVABLE - 
	' MUST BE ABLE TO EMAIL ATTACHMENT FROM A PHYSICAL LOCATION
	'**************************************************************************************************

		'baseURL should always have a trailing /slash, just in case, handle either way
		If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
		sURL = Request.ServerVariables("SERVER_NAME")
	
		Orig_Name = "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_Chain_" & Trim(ChainID) & "_" & Trim(ChainID) & Trim(Replace(EndDate,"/","")) & ".pdf"
		New_Name =  "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\customer\accountsreceivable\ConsolidatedStatement_Chain_" & Trim(ChainID) & "_" & Trim(ChainID) & Trim(Replace(EndDate,"/","")) & ".pdf"
		
		Dim fso
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'Kill it first in case an old one is there
		On error resume next
		fso.DeleteFile Server.MapPath(New_Name)
		On error goto 0
		
		fso.CopyFile Server.MapPath(Orig_Name), Server.MapPath(New_Name)
		
		Set fso = Nothing
	
	'**************************************************************************************************
	'END CODE TO SAVE CONSOLIDATED INVOICE TO CLIENTFILES\CLIENTID\CUSTOMER\ACCOUNTSRECEIVABLE
	'**************************************************************************************************

	
	
	'**************************************************************************************************
	' NOW EMAIL PDF TO ALL THE USERS AND ADDITIONAL EMAIL ADDRESSES SPECIFIED
	'**************************************************************************************************

		fn = Server.MapPath(New_Name)
		Response.Write("File to attach: " & fn & "<br>")

		'Send user based emails
		If Emailto <> "" Then
			UserNoList = Split(Emailto,",")
			For x = 0 To UBound(UserNoList)
				Send_To = Send_To & GetUserEmailByUserNo(UserNoList(x)) & ";"
			Next
		End If
		
		'Got all the addresses so now break them up
		Send_To_Array = Split(Send_To,";")
		
		'Response.Write("<br>Send_To: " & Send_To & "<br>")
		
		'HERE WE ACTUALLY SEND THE EMAIL
		For x = 0 to Ubound(Send_To_Array) -1
			Send_To = Send_To_Array(x)
	
			If PaidOrUnpaid = "UNPAID" Then
				emailSubject = "Consolidated Invoice (Unpaid) For Chain " & ChainID & " (" & ChainName & ")"
			Else
				emailSubject = "Consolidated Invoice For Chain " & ChainID & " (" & ChainName & ")"
			End If

			emailBody = ""
			'Failsafe for dev
			sURL = Request.ServerVariables("SERVER_NAME")
			If Instr(ucase(sURL),"DEV.") <> 0 Then Send_To = "rsmith@ocsaccess.com"
			If Instr(ucase(sURL),"DEV2.") <> 0 Then Send_To = "cgrecco@ocsaccess.com"
			emailBody = "The consolidate invoice for the customer is attached. (" & MUV_Read("ClientID") & ")"
			
			Response.Write("<br>FN: " & fn & "<br>")

			SendMailWatt "mailsender@" & maildomain,Send_To,emailSubject,emailBody,fn,GetTerm("Accounts Receivable"),"Consolidated Invoice Email"
			
			If PaidOrUnpaid = "UNPAID" Then
				CreateAuditLogEntry "Consolidated (Unpaid) Invoice Email","Consolidated Invoice (Unpaid) Email","Minor",0,"Consolidated Invoice (Unpaid) Email Sent to " & Send_To 
			Else
				CreateAuditLogEntry "Consolidated Invoice Email","Consolidated Invoice Email","Minor",0,"Consolidated Invoice Email Sent to " & Send_To
			End If
 
			Response.Write("Sent the email to " & Send_To & "<br>")
			Response.Write("Sent the email, all done<br>")
		Next 


		'Send additional emails
		If AdditionalEmails <> "" Then
			AddlEmailAddressList = Split(AdditionalEmails,";")
			For x = 0 To UBound(AddlEmailAddressList)
				Send_To = Send_To & AddlEmailAddressList(x) & ";"
			Next
		End If
		
		'Got all the addresses so now break them up
		Send_To_Array_Additional = Split(Send_To,";")
		
		'Response.Write("<br>Send_To: " & Send_To & "<br>")
		
		'HERE WE ACTUALLY SEND THE EMAIL
		For x = 0 to Ubound(Send_To_Array_Additional) -1
			Send_To = Send_To_Array_Additional(x)
	
			If PaidOrUnpaid = "UNPAID" Then
				emailSubject = "Consolidated Invoice (Unpaid) For Chain " & ChainID & " (" & ChainName & ")"
			Else
				emailSubject = "Consolidated Invoice For Chain " & ChainID & " (" & ChainName & ")"
			End If

			emailBody = ""
			'Failsafe for dev
			sURL = Request.ServerVariables("SERVER_NAME")
			If Instr(ucase(sURL),"DEV.") <> 0 Then Send_To = "rsmith@ocsaccess.com"
			If Instr(ucase(sURL),"DEV2.") <> 0 Then Send_To = "cgrecco@ocsaccess.com"
			emailBody = "The consolidate invoice for the customer is attached. (" & MUV_Read("ClientID") & ")"
			
			Response.Write("<br>FN: " & fn & "<br>")

			SendMailWatt "mailsender@" & maildomain,Send_To,emailSubject,emailBody,fn,GetTerm("Accounts Receivable"),"Consolidated Invoice Email"
			
			If PaidOrUnpaid = "UNPAID" Then
				CreateAuditLogEntry "Consolidated (Unpaid) Invoice Email","Consolidated Invoice (Unpaid) Email","Minor",0,"Consolidated Invoice (Unpaid) Email Sent to " & Send_To 
			Else
				CreateAuditLogEntry "Consolidated Invoice Email","Consolidated Invoice Email","Minor",0,"Consolidated Invoice Email Sent to " & Send_To
			End If
 
			Response.Write("Sent the email to " & Send_To & "<br>")
			Response.Write("Sent the email, all done<br>")
		Next	

	'**************************************************************************************************
	' END SEND EMAIL
	'**************************************************************************************************

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub SaveCustomerMES()

	MESAmount = Request.Form("mes") 
	MESAmount = Replace(MESAmount , "'", "''")
	MESAmount = Round(MESAmount,2)

    MCSAmount = Request.Form("mcs") 
	MCSAmount = Replace(MCSAmount , "'", "''")
    IF LEN(MCSAmount)>0 THEN
	    MCSAmount = Round(MCSAmount,2)
        ELSE
        MCSAmount=0
    END IF

    maxmcsAmount = Request.Form("maxmcs") 
	maxmcsAmount = Replace(maxmcsAmount , "'", "''")
    IF LEN(maxmcsAmount)>0 THEN
	    maxmcsAmount = Round(maxmcsAmount,2)
        ELSE
        MCSAmount=0
    END IF
	

	CustAccountIdentifyingInfoForSQL = Request.Form("id")

	CustAccountIdentifyingInfoForSQLArray = Split(CustAccountIdentifyingInfoForSQL,"*")
	
	CustID = CustAccountIdentifyingInfoForSQLArray(1)
	
	Set rsSaveEquivCustID = Server.CreateObject("ADODB.Recordset")
	rsSaveEquivCustID.CursorLocation = 3 
	
	If MESAmount <> "" AND  CustID <> "" Then
	    SQLUpdate = "UPDATE AR_Customer SET "
		If cint(MESAmount) <> 0 Then 
			  SQLUpdate = SQLUpdate&"MonthlyExpectedSalesDollars = '" & MESAmount &"'"
		Else
			SQLUpdate = SQLUpdate&"MonthlyExpectedSalesDollars = NULL" 
		End If
		If cint(MESAmount) <> 0 Then 
			  SQLUpdate = SQLUpdate&",MonthlyContractedSalesDollars = '" & MCSAmount &"'"
		Else
			SQLUpdate = SQLUpdate&",MonthlyContractedSalesDollars = NULL" 
		End If
        If cint(MESAmount) <> 0 Then 
			  SQLUpdate = SQLUpdate&",maxMCSCharge = '" & maxmcsAmount &"'"
		Else
			SQLUpdate = SQLUpdate&",maxMCSCharge = NULL" 
		End If

        SQLUpdate = SQLUpdate&" WHERE CustNum = '" & CustID & "'"
		Set cnnUpdate = Server.CreateObject("ADODB.Connection")
		cnnUpdate.open (Session("ClientCnnString"))
		Set rsUpdate = Server.CreateObject("ADODB.Recordset")
		rsUpdate.CursorLocation = 3 
		Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
		cnnUpdate.close
		
		'Response.Write(SQLUpdate)
		Response.Write("Success")
			
	Else
		Response.Write("Cannot Save, Invalid Data" & CustID & MESAmount  )
		
	End If

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub ExportSelectedSageInvoices()

	DIM buffer
	buffer=array()


	If Request.Form("InvoicesToExport") <> "" Then
	
		InvArray = Split(Request.Form("InvoicesToExport"),",")
		
		For x = 0 to Ubound(InvArray)
		
			InvArray(x) = Trim(InvArray(x))
		
		Next


		'Build the export data
		Set cnnExportSage = Server.CreateObject("ADODB.Connection")
		cnnExportSage.open (Session("ClientCnnString"))
		Set rsExportSage = Server.CreateObject("ADODB.Recordset")
		
		SQLExportSage = "SELECT * FROM IN_InvoiceHistHeader WHERE InvoiceID IN ("
		
		For x = 0 to Ubound(InvArray)
		
			SQLExportSage = SQLExportSage & "'" & InvArray(x) & "',"
		
		Next

		If Right(SQLExportSage,1) = "," Then SQLExportSage = Left(SQLExportSage,Len(SQLExportSage)-1)
				
		SQLExportSage = SQLExportSage & ")"
		
		
		Set rsExportSage = cnnExportSage.Execute(SQLExportSage)

		If Not rsExportSage.EOF Then
		
			Set rsExportSageDetails = Server.CreateObject("ADODB.Recordset")
			Set rsExportSageDetailsForCounting = Server.CreateObject("ADODB.Recordset")
		
			Do While Not rsExportSage.EOF
	
				SQLExportSageDetails = "SELECT * FROM IN_InvoiceHistDetail WHERE InvoiceID = '" & rsExportSage("InvoiceID") & "' ORDER BY LineNumber"
			
				Set rsExportSageDetails = cnnExportSage.Execute(SQLExportSageDetails)
				
				SQLExportSageDetailsForCounting = "SELECT Count(*) As LineCnt FROM IN_InvoiceHistDetail WHERE InvoiceID = '" & rsExportSage("InvoiceID") & "'"
			
				Set rsExportSageDetailsForCounting = cnnExportSage.Execute(SQLExportSageDetailsForCounting)

				
				If Not rsExportSageDetails.EOF Then

					LineCount = 1
				
					Do While Not rsExportSageDetails.EOF

						'Now Build the output line
						
						
						bufferLine = ""
						
						
						
						bufferLine = bufferLine & """" & rsExportSage("AlternateCustID") & ""","
						bufferLine = bufferLine & """" &  GetCustNameByCustNum(rsExportSage("CustID")) & ""","
						bufferLine = bufferLine & """" &  rsExportSage("InvoiceID") & ""","
						
						bufferLine = bufferLine & ","	' Credit memo
						bufferLine = bufferLine & "0" & ","	' Progress billing invoice
						bufferLine = bufferLine & "0" & ","	' dunno
						
						InvoiceDate = cDate(rsExportSage("InvoiceCreationDate")) 
						eYear = Year(InvoiceDate)
						'eYear = Right(eYear,2)
						If Month(InvoiceDate) < 10 Then eMonth = "0" & Month(InvoiceDate) else eMonth = Month(InvoiceDate)
						If Day(InvoiceDate) < 10 Then eDay = "0" & Day(InvoiceDate) else eDay = Day(InvoiceDate)
						InvoiceDispayableDate = eMonth & "/" & eDay  & "/" & eYear
						'InvoiceDispayableDate = cDate(InvoiceDispayableDate ) 

						bufferLine = bufferLine & """" &  InvoiceDispayableDate & ""","

						bufferLine = bufferLine & ","	' G
						bufferLine = bufferLine & "0" & ","	' Ship by quote
						bufferLine = bufferLine & ","	' Quote number
						bufferLine = bufferLine & ","	' Quote	good thur date
						bufferLine = bufferLine & "0" & ","	' drop ship
						
						
						bufferLine = bufferLine & """" &  rsExportSage("ShipToName") & ""","
						bufferLine = bufferLine & """" &  rsExportSage("ShipToAddr1") & ""","
						bufferLine = bufferLine & """" &  rsExportSage("ShipToAddr2") & ""","
						bufferLine = bufferLine & """" &  rsExportSage("ShipToCity") & ""","
						bufferLine = bufferLine & """" &  rsExportSage("ShipToState") & ""","
						bufferLine = bufferLine & """" &  rsExportSage("ShipToPostalCode") & ""","
						bufferLine = bufferLine & ","	' ship to country
						
						bufferLine = bufferLine & """" & rsExportSage("PONumber") & ""","						
						bufferLine = bufferLine & """" &  "" & ""","	' ship via ship date
						bufferLine = bufferLine & """" &  "" & ""","	' V

						
						InvoiceDueDate = cDate(rsExportSage("InvoiceDueDate")) 
						eYear = Year(InvoiceDueDate)
						'eYear = Right(eYear,2)
						If Month(InvoiceDueDate) < 10 Then eMonth = "0" & Month(InvoiceDueDate) else eMonth = Month(InvoiceDueDate)
						If Day(InvoiceDueDate) < 10 Then eDay = "0" & Day(InvoiceDueDate) else eDay = Day(InvoiceDueDate)
						InvoiceDispayableDueDate = eMonth & "/" & eDay  & "/" & eYear
						'InvoiceDispayableDueDate = cDate(InvoiceDispayableDueDate ) 

						bufferLine = bufferLine & """" & InvoiceDispayableDueDate & ""","
			
						bufferLine = bufferLine & "0" & ","	' Discount Amount
						bufferLine = bufferLine & """" &  InvoiceDispayableDate & ""","	'Discount Date
						
						bufferLine = bufferLine & """" &  rsExportSage("Terms") & ""","
						
						bufferLine = bufferLine & """" &  "" & ""","	' sales rep

						bufferLine = bufferLine & """" &  rsExportSageDetails("GL_AR_Account") & ""","
						bufferLine = bufferLine & """" &  rsExportSage("InvoiceGrandTotal") & """," ' a/r amount
						
						bufferLine = bufferLine & """" &  rsExportSage("ShipToState") & ""","		' sales tax id
						
						bufferLine = bufferLine & """" &  "" & ""","	' invoice note
						bufferLine = bufferLine & "0" & ","	' invoice prints after line items
						bufferLine = bufferLine & """" &  "" & ""","	' stmt note
						bufferLine = bufferLine & "0" & ","	' stmt print
						bufferLine = bufferLine & """" &  "" & ""","	' internal note

						bufferLine = bufferLine & "0" & ","	' beginning blance
						bufferLine = bufferLine & """" &  "" & ""","	' ar dte cleared

						bufferLine = bufferLine & """" & rsExportSageDetailsForCounting("LineCnt") + 1 & ""","	' number of distributions
						
						bufferLine = bufferLine & """" & LineCount & ""","	' invoice cm dist  counter with tax being 0
						
						bufferLine = bufferLine & "0" & ","	' apply inv to dist
						
						bufferLine = bufferLine & "0" & ","	' apply to sales order
						bufferLine = bufferLine & "0" & ","	' apply to proposal
						
						bufferLine = bufferLine &  rsExportSageDetails("QtyShipped") & ","
						
						bufferLine = bufferLine & """" &  "" & ""","	' so proposal #
						bufferLine = bufferLine & """" &  rsExportSageDetails("prodSKU") & ""","
						bufferLine = bufferLine & """" &  "" & ""","	' serial #
						bufferLine = bufferLine & """" & "0" & ""","	'  proposal dist
						
						
						bufferLine = bufferLine & """" &  rsExportSageDetails("prodDescription") & ""","
						
						bufferLine = bufferLine & """" &  rsExportSageDetails("GL_Account") & ""","   'AAAAAAAAAAAAWWWWWWWWWWWWWWWWWWWWWWW user field 1
						
						
						
						bufferLine = bufferLine & """" &  "" & ""","	
						bufferLine = bufferLine & rsExportSageDetails("PricePerUnitSold") & ","
						
						
						bufferLine = bufferLine & """" &  rsExportSageDetails("GL_AR_Account") & ""","
						
						bufferLine = bufferLine & """" & rsExportSageDetailsForCounting("LineCnt") + 1  & ""","	' Number of line items  plus 1
						bufferLine = bufferLine & """" &  rsExportSageDetails("GL_Account") & ""","

						bufferLine = bufferLine & """" & "1" & ""","	' tax type always 1 unless tax line
						
						bufferLine = bufferLine & """" & "0" & ""","   'weight
						
						bufferLine = bufferLine & """" &  (rsExportSageDetails("QtyShipped") * rsExportSageDetails("PricePerUnitSold")) * -1  & ""","
						
						
						' Four blanks bd,be,
						bufferLine = bufferLine & """" & "" & ""","
						bufferLine = bufferLine & """" & "" & ""","
						bufferLine = bufferLine & """" & "" & ""","
						bufferLine = bufferLine & """" & "" & ""","

						bufferLine = bufferLine & """" & "<Each>" & ""","																		
						bufferLine = bufferLine & """" &  "1" & """," ' um no of stocking units	always a 1 unless its the tax line
						
						
						
						bufferLine = bufferLine & """" &  rsExportSageDetails("QtyShipped") & """," ' sotkcing qty
						bufferLine = bufferLine & """" &   rsExportSageDetails("PricePerUnitSold") & """," ' stocling un it proce
						
						bufferLine = bufferLine & """" & "0" & """" ' cost of sales job id
						bufferLine = bufferLine & """" & "" & ""","
						
						bufferLine = bufferLine & """" &  rsExportSage("ShipToState") & ""","' tax agencyid	
						
						
						
						bufferLine = bufferLine & """" &  (rsExportSageDetails("QtyShipped") * rsExportSageDetails("PricePerUnitSold")) + rsExportSageDetails("TotalTaxForLine") & ""","
						
						bufferLine = bufferLine & """" & "0" & """," ' Boolean credit memo flag
						
						bufferLine = bufferLine & """" &  "26"  & ""","' transaction period
						bufferLine = bufferLine & """" &  ""  & ""","' transaction number - leave it blank & see what happens

						' 3 blanks bd,be,
						bufferLine = bufferLine & """" & "" & ""","
						bufferLine = bufferLine & """" & "" & ""","
						bufferLine = bufferLine & """" & "" & ""","

						' 3 zeros
						bufferLine = bufferLine & "0" & ","
						bufferLine = bufferLine & "0" & ","
						bufferLine = bufferLine & "0" 
						
						buffer = AddItem(buffer,bufferLine)

						rsExportSageDetails.movenext

						LineCount = LineCount + 1
											
					Loop				

								
				End If
			
				rsExportSage.MoveNext
			Loop
		
		End If
		
			
		'Mark all of the invoices as having been exported
		Set cnnExportSage = Server.CreateObject("ADODB.Connection")
		cnnExportSage.open (Session("ClientCnnString"))
		Set rsExportSage = Server.CreateObject("ADODB.Recordset")
		
		For x = 0 to Ubound(InvArray)
		
			SQLExportSage = "INSERT INTO IN_InvoicesExportedSage (InvoiceID) VALUES ('" & InvArray(x) & "')"
		
			Set rsExportSage = cnnExportSage.Execute(SQLExportSage)
		Next
		
		Set rsExportSage = Nothing
		cnnExportSage.close
		Set	cnnExportSage = Nothing

		
		'Name & download the file
		'strFile = "VendmaxToSage_"&Right("0" & Day(Now),2)&+Right("0" & Month(Now),2)&YEAR(Now)&".txt"
		
		'Response.Clear
		' Download to user
		'Response.AddHeader "Content-Disposition", "attachment; filename=" & strFile
		'Response.AddHeader "Content-Length", LEN(JOIN(buffer,CHR(13)&CHR(10)))
		'Response.ContentType = "application/octet-stream"
		'Response.CharSet = "UTF-8"
		'-- send the stream in the response
		'Response.BinaryWrite(JOIN(buffer,CHR(13)&CHR(10)))
		
		Response.Write(JOIN(buffer,CHR(13)&CHR(10)))

	
	End IF

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Function AddItem(arr, val)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
    AddItem = arr
End Function

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub ToggleShowHideExportedSageInvoices()

    ShowHideValue = Request.Form("ShowHide")
	
	If ShowHideValue = "SHOW" Then
		dummy = MUV_Write("showExportedSageInvoices","SHOW")
	Else
		dummy = MUV_Write("showExportedSageInvoices","HIDE")
	End If
    
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub CheckIfDefaultBillingLocation() 

	IntRecID = Request.Form("i")
	
	SQLARCustBillTo = "SELECT * FROM AR_CustomerBillTo WHERE InternalRecordIdentifier = " & IntRecID 
		
	Set cnnARCustBillTo = Server.CreateObject("ADODB.Connection")
	cnnARCustBillTo.open(Session("ClientCnnString"))
	Set rsARCustBillTo = Server.CreateObject("ADODB.Recordset")
	rsARCustBillTo.CursorLocation = 3 
	
	Set rsARCustBillTo = cnnARCustBillTo.Execute(SQLARCustBillTo)

	If NOT rsARCustBillTo.EOF Then
		DefaultBillTo = rsARCustBillTo("DefaultBillTo")
	Else
		Response.Write("DONOTDELETE")
	End If
	
	If DefaultBillTo = 1 Then
		Response.Write("DONOTDELETE")
	Else
		Response.Write("OKTODELETE")
	End If
		
	Set rsARCustBillTo = Nothing
	cnnARCustBillTo.Close
	Set cnnARCustBillTo = Nothing
	
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetCustomerAccountInformationForModal()

	CustID = Request.Form("CustID")

	SQLCustomer = "SELECT * FROM AR_Customer WHERE CustNum = '" & CustID & "'"

	Set cnnCustomer = Server.CreateObject("ADODB.Connection")
	cnnCustomer.open (Session("ClientCnnString"))
	Set rsCustomer = Server.CreateObject("ADODB.Recordset")
	rsCustomer.CursorLocation = 3 
	Set rsCustomer = cnnCustomer.Execute(SQLCustomer)

	If not rsCustomer.EOF Then
		CompanyName = rsCustomer("Name")												
	End If
	
	set rsCustomer = Nothing
	cnnCustomer.close
	set cnnCustomer = Nothing
	
	%>
	<script language="JavaScript">
		<!--
		
		$(document).ready(function() {
			
			var focus = 0
			
			$("#txtAccountNumber").focusout(function() {
							
				var passedNewCustID = $("#txtAccountNumber").val();
				var passedCurrCustID = $("#txtCustID").val();
				
		    	$.ajax({
					type:"POST",
					url: "../../inc/InSightFuncs_AjaxForARAP.asp",
					cache: false,
					data: "action=CheckIfCustomerIDAlreadyExists&passedNewCustID=" + encodeURIComponent(passedNewCustID) + "&passedCurrCustID=" + encodeURIComponent(passedCurrCustID),
					success: function(response)
					 {
		               	 if (response == "CUSTIDALREADYEXISTS") {
		               	 	swal("That Account Number Already Exists for Another Customer.");
		               	 	$("#txtAccountNumber").val('');
		               	 }
					 }		
				});

				
			});
			
		
		        
		});
		
	
	   function validateEditCustomer()
	    {
	    
	       if (document.frmEditCustomerFromModal.txtAccountNumber.value == "") {
	            swal("Account number cannot be blank.");
	            return false;
	       }
	       if (document.frmEditCustomerFromModal.txtCompanyName.value == "") {
	            swal("Company name cannot be blank.");
	            return false;
	       }
	          
	       return true;
	
	    }
	// -->
	</script>  
	
	
	<style>
		label {
			margin-top:15px;
		}
		
		.input-group .form-control, .input-group-addon, .input-group-btn {
		    display: table-cell;
		    height: 38px;
		}	
		
		.input-group-addon {
		    padding: 6px 12px;
		    font-size: 14px;
		    font-weight: 400;
		    line-height: 1;
		    color: #555;
		    text-align: center;
		    background-color: #eee;
		    border: 1px solid #ccc;
		    border-radius: 4px;
		    height: 38px;
		}
	
	</style>

	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">

              <div class="form-group">         
	                <div class="col-sm-10">
	                  <label>Account Number</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtAccountNumberIcon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtAccountNumber" name="txtAccountNumber" value="<%= CustID %>">
	                   </div>
	                </div> 
               </div>
               
			  <br clear="all">
				
              <div class="form-group">         	                
	                <div class="col-sm-10">
	                  <label>Company Name</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtCompanyNameIcon"><i class="fa fa-briefcase"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtCompanyName" name="txtCompanyName" value="<%= CompanyName %>">
	                   </div>
	                </div> 
               </div>
                    
   		</div>
	</div>
</div>

<%

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetCustomerPricingInformationForModal()

	CustID = Request.Form("CustID")
	InternalRecordIdentifier = Request.Form("IntRecID")

	SQLCustomer = "SELECT * FROM AR_Customer WHERE CustNum = '" & CustID & "'"

	Set cnnCustomer = Server.CreateObject("ADODB.Connection")
	cnnCustomer.open (Session("ClientCnnString"))
	Set rsCustomer = Server.CreateObject("ADODB.Recordset")
	rsCustomer.CursorLocation = 3 
	Set rsCustomer = cnnCustomer.Execute(SQLCustomer)

	If not rsCustomer.EOF Then
		CompanyName = rsCustomer("Name")
		LastPriceChangeDate = rsCustomer("LastPriceChangeDate")												
	End If
	
	set rsCustomer = Nothing
	cnnCustomer.close
	set cnnCustomer = Nothing
	
	%>
	<script language="JavaScript">
		<!--
		$(document).ready(function() {
			
	        $('#datetimepickerLastPriceChangeDate').datetimepicker({
	        	useCurrent: false,
	        	format: 'MM/DD/YYYY',
	        	maxDate:moment(),
	        	ignoreReadonly: true,
	        	showClear: true,
			});   
			
			$("#resetLastPriceChangeDate").click(function(){
				$("#datetimepickerLastPriceChangeDate").data("DateTimePicker").date(null);
			});
		        
		});
	
	// -->
	</script>  
	
	
	<style>
		label {
			margin-top:15px;
		}
		
		.input-group .form-control, .input-group-addon, .input-group-btn {
		    display: table-cell;
		    height: 38px;
		}	
		
		.input-group-addon {
		    padding: 6px 12px;
		    font-size: 14px;
		    font-weight: 400;
		    line-height: 1;
		    color: #555;
		    text-align: center;
		    background-color: #eee;
		    border: 1px solid #ccc;
		    border-radius: 4px;
		    height: 38px;
		}
	
	</style>

	<div class="row">					
		<div class="col-lg-12">	
	        <input type="hidden" id="txtAccountNumber" name="txtAccountNumber" value="<%= CustID %>"> 
	        <input type="hidden" id="txtInternalRecordIdentifier" name="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">                
                    
			<div class="form-group">
	            <div class="col-sm-10">
	              <label>Last Price Change Date</label>
                  <div class="input-group" id="datetimepickerLastPriceChangeDate" style="width:250px;">
                    	<div class="input-group-addon" id="txtLastPriceChangeDateIcon"><span class="glyphicon glyphicon-calendar"></span></div>
                    	<% If IsNull(LastPriceChangeDate) OR LastPriceChangeDate ="1/1/1900" Then %>
                    		<input type="text" class="form-control" id="txtLastPriceChangeDate" name="txtLastPriceChangeDate" value="<%= Now() %>">
                    	<% Else %>
                    		<input type="text" class="form-control" id="txtLastPriceChangeDate" name="txtLastPriceChangeDate" value="<%= LastPriceChangeDate %>">
                    	<% End If %>
                   </div>
				</div>			  	
			</div>
		</div>
	</div>

	<div class="row" style="margin-top:20px;">					
		<div class="col-lg-12">			
			<div class="form-group">	
				<div class="col-sm-10">
					<button id="resetLastPriceChangeDate" type="button" class="btn btn-primary">Clear Last Price Change Date</button>
				</div>		  	
			</div>   
		</div>
	</div>

<%

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub CheckIfCustomerIDAlreadyExists()

	passedNewCustID = Request.Form("passedNewCustID")
	passedCurrCustID = Request.Form("passedCurrCustID")
	
	If passedCurrCustID <> passedNewCustID Then
		
		SQLCheckForDuplicateCustID = "SELECT * FROM AR_Customer WHERE custNum = '" & passedNewCustID & "'"
			
		Set cnnCheckForDuplicateCustID = Server.CreateObject("ADODB.Connection")
		cnnCheckForDuplicateCustID.open(Session("ClientCnnString"))
		Set rsCheckForDuplicateCustID = Server.CreateObject("ADODB.Recordset")
		rsCheckForDuplicateCustID.CursorLocation = 3 
		
		Set rsCheckForDuplicateCustID = cnnCheckForDuplicateCustID.Execute(SQLCheckForDuplicateCustID)
	
		If NOT rsCheckForDuplicateCustID.EOF Then
			Response.Write("CUSTIDALREADYEXISTS")
		Else
			Response.Write("CUSTIDNOTINUSE")
		End If
			
		Set rsCheckForDuplicateCustID = Nothing
		cnnCheckForDuplicateCustID.Close
		Set cnnCheckForDuplicateCustID = Nothing
		
	Else
	
		Response.Write("CUSTIDNOTCHANGED")
		
	End If
	

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetCustomerNoteCountByNoteType()

	passedNoteIntRecID = Request.Form("NoteIntRecID")
	passedCustID = Request.Form("CustID")
	passedShowOnlyMyNotes = Request.Form("ShowOnlyMyNotes")
	passesUserNo = Request.Form("userNo")

	resultGetCustomerNoteCountByNoteType = 0
	
	
	If passedNoteIntRecID <> "" AND passedCustID <> "" Then
	
		Set cnnGetCustomerNoteCountByNoteType = Server.CreateObject("ADODB.Connection")
		cnnGetCustomerNoteCountByNoteType.open Session("ClientCnnString")
		Set rsGetCustomerNoteCountByNoteType = Server.CreateObject("ADODB.Recordset")
		rsGetCustomerNoteCountByNoteType.CursorLocation = 3 
		
		If passedNoteIntRecID = 0 Then
			If passedShowOnlyMyNotes = 0 Then
				SQLGetCustomerNoteCountByNoteType = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID <> '' AND CustID = '" & passedCustID & "'"
			Else
				SQLGetCustomerNoteCountByNoteType = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID <> '' AND CustID = '" & passedCustID & "' AND EnteredByUserNo = " & passesUserNo
			End If
		Else
			If passedShowOnlyMyNotes = 0 Then
				SQLGetCustomerNoteCountByNoteType = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID = " & passedNoteIntRecID & " AND CustID = '" & passedCustID & "'"
			Else
				SQLGetCustomerNoteCountByNoteType = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID = " & passedNoteIntRecID & " AND CustID = '" & passedCustID & "' AND EnteredByUserNo = " & passesUserNo
			End If
		End If
	
		Set rsGetCustomerNoteCountByNoteType= cnnGetCustomerNoteCountByNoteType.Execute(SQLGetCustomerNoteCountByNoteType)
		
		If not rsGetCustomerNoteCountByNoteType.eof then resultGetCustomerNoteCountByNoteType = rsGetCustomerNoteCountByNoteType("NoteTypeCount")
		
		Set rsGetCustomerNoteCountByNoteType= Nothing
		cnnGetCustomerNoteCountByNoteType.Close
		Set cnnGetCustomerNoteCountByNoteType= Nothing
		
	End If
	
	Response.Write(resultGetCustomerNoteCountByNoteType)
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes()

	passedNoteIntRecID = Request.Form("NoteIntRecID")
	passedCustID = Request.Form("CustID")
	passedShowOnlyMyNotes = Request.Form("ShowOnlyMyNotes")
	passedUserNo = Request.Form("userNo")
	
	resultGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes = ""
	NoteTypeCount = 0
	
	If passedNoteIntRecID <> "" AND passedCustID <> "" Then
	
		Set cnnGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes = Server.CreateObject("ADODB.Connection")
		cnnGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes.open Session("ClientCnnString")
		Set rsGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes = Server.CreateObject("ADODB.Recordset")
		rsGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes.CursorLocation = 3 
		
		'***************************************************
		'THIS SECTION GETS THE NOTE COUNT FOR THE ALL NOTES TABS
		'***************************************************
		If passedNoteIntRecID = 0 Then
			If passedShowOnlyMyNotes = 0 Then
				SQLGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID <> '' AND CustID = '" & passedCustID & "'"
			Else
				SQLGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID <> '' AND CustID = '" & passedCustID & "' AND EnteredByUserNo = " & Session("UserNo")
			End If
		'***************************************************
		'THIS SECTION GETS THE NOTES COUNT FOR PARTICULAR NOTE TABS
		'***************************************************			
		Else
			If passedShowOnlyMyNotes = 0 Then
				SQLGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID = " & passedNoteIntRecID & " AND CustID = '" & passedCustID & "'"
			Else
				SQLGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID = " & passedNoteIntRecID & " AND CustID = '" & passedCustID & "' AND EnteredByUserNo = " & Session("UserNo")
			End If
		End If
	
		Set rsGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes= cnnGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes.Execute(SQLGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes)
		
		If not rsGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes.eof then NoteTypeCount = rsGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes("NoteTypeCount")
		
		Set rsGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes= Nothing
		cnnGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes.Close
		Set cnnGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes= Nothing
		
		'***************************************************
		'THIS SECTION GETS THE NAME OF THE TAB
		'***************************************************
		If passedNoteIntRecID = 0 Then
			NoteGetTermName = "All"
		Else
			NoteGetTermName = GetTerm(GetCustNoteTypeByNoteIntRecID(passedNoteIntRecID))	
		End If

		'***************************************************
		'THIS SECTION CHECKS TO SEE IF THERE ARE UNREAD NOTES
		'***************************************************
		
		NotesHaveBeenRead = "True"
		NotesHaveBeenRead = HasNoteTypeBeenViewedByUser(passedCustID,passedNoteIntRecID,passedShowOnlyMyNotes)
		
		resultGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes = NoteGetTermName & "*" & NoteTypeCount & "*" & NotesHaveBeenRead
		
	End If
	
	Response.Write(resultGetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes)
	
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub MarkAllNotesForNoteTypeForUserAsRead()

	passedNoteTypeIntRecID = Request.Form("NoteIntRecID")
	passedCustID = Request.Form("CustID")

	SQLMarkAllNotesForNoteTypeForUserAsRead = "SELECT * FROM AR_CustomerNotesUserViewed Where CustID ='" & passedCustID & "' AND UserNo = " & Session("Userno") & " AND NoteTypeIntRecID = " & passedNoteTypeIntRecID
	
	Set cnnMarkAllNotesForNoteTypeForUserAsRead = Server.CreateObject("ADODB.Connection")
	cnnMarkAllNotesForNoteTypeForUserAsRead.open (Session("ClientCnnString"))
	Set rMarkAllNotesForNoteTypeForUserAsRead = Server.CreateObject("ADODB.Recordset")
	rMarkAllNotesForNoteTypeForUserAsRead.CursorLocation = 3 
	Set rMarkAllNotesForNoteTypeForUserAsRead = cnnMarkAllNotesForNoteTypeForUserAsRead.Execute(SQLMarkAllNotesForNoteTypeForUserAsRead)

	If rMarkAllNotesForNoteTypeForUserAsRead.EOF Then ' Nothing there so we need to insert
		SQLMarkAllNotesForNoteTypeForUserAsRead = "INSERT INTO AR_CustomerNotesUserViewed (CustID, UserNo, Category, NoteTypeIntRecID) VALUES ('" & passedCustID & "'," & Session("UserNo") & ",-2," & passedNoteTypeIntRecID & ")"
	Else
		SQLMarkAllNotesForNoteTypeForUserAsRead = "UPDATE AR_CustomerNotesUserViewed Set DateLastViewed = getdate() Where CustID ='" & passedCustID & "' AND UserNo = " & Session("UserNo") & " AND NoteTypeIntRecID = " & passedNoteTypeIntRecID
	End If
	
	Set rMarkAllNotesForNoteTypeForUserAsRead = cnnMarkAllNotesForNoteTypeForUserAsRead.Execute(SQLMarkAllNotesForNoteTypeForUserAsRead)
		
	cnnMarkAllNotesForNoteTypeForUserAsRead.close
	set rMarkAllNotesForNoteTypeForUserAsRead = nothing
	set cnnMarkAllNotesForNoteTypeForUserAsRead= nothing	

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetContentForCustomerNotesModal() 

	CustIDPassed = Request.Form("CustID")
	CustNamePassed = GetCustNameByCustNum(CustIDPassed)
		
	'********************
	' **** Notes Tab ****
	'********************
	
	'**********************************************
	'Create Drop Down Selections For Note Type
	'**********************************************
	
	SQLNoteTypeDropdown = "SELECT * FROM SC_NoteType ORDER BY NoteType ASC"
	
	Set cnnNoteTypeDropdown = Server.CreateObject("ADODB.Connection")
	cnnNoteTypeDropdown.open (Session("ClientCnnString"))
	
	Set rsNoteTypeDropdown = Server.CreateObject("ADODB.Recordset")
	Set rsNoteTypeDropdown = cnnNoteTypeDropdown.Execute(SQLNoteTypeDropdown)
	
	NoteTypes = ("[{""id"":"""",""title"":""Select a Note Type""},")
	
	If not rsNoteTypeDropdown.EOF Then
		sep = ""
		Do While Not rsNoteTypeDropdown.EOF
			If rsNoteTypeDropdown("NoteTypeCanBeCreatedByUser") = 1 Then
				NoteTypes = NoteTypes & (sep)
				sep = ","
				NoteTypes = NoteTypes & ("{")
				NoteTypes = NoteTypes & ("""id"":""" & Replace(rsNoteTypeDropdown("InternalRecordIdentifier"), """", "\""") & """")
				NoteTypes = NoteTypes & (",""title"":""" & Replace(GetTerm(rsNoteTypeDropdown("NoteType")), """", "\""") & """")
				NoteTypes = NoteTypes & ("}")
			End If
			rsNoteTypeDropdown.MoveNext						
		Loop
	End If
	NoteTypes = NoteTypes & ("]")
	Set rsNoteTypeDropdown = Nothing
	
	%>
	<style type="text/css">
		.unread-note{
			background-color:yellow;
			font-weight:bold;
		}
		.read-note{
			background-color:white;
			font-weight:normal;
		}	
	</style>
	<!-- modal header !-->
	<div class="modal-header" style="min-height:35px !important;">
		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		<h4 class="modal-title" id="myWebOrdersLabel">Notes for <%= CustNamePassed %></h4>
	</div>
	<!-- eof modal header !-->
	
	<!-- modal body !-->
	<div class="modal-body" style="max-height:450px;overflow:scroll">
	
	<input type="hidden" name="txtNoteTypeIntRecIDToShow" id="txtNoteTypeIntRecIDToShow" value="0">
	<input type="hidden" name="txtCustIDToShow" id="txtCustIDToShow" value="<%= CustIDPassed %>">
	<input type="hidden" name="txtUserNoToShow" id="txtUserNoToShow" value="<%= Session("UserNo") %>">
	<input type="hidden" name="txtBaseURL" id="txtBaseURL" value="<%= BaseURL %>">
	
	<p style="margin-left:20px;margin-bottom:20px"><input type="checkbox" id="notesall" value="all"> View only notes I have entered</p>

    <!-- Nav tabs -->
    <ul class="nav nav-tabs" role="tablist" id="customerNoteTabs">
      <li class="active">
			<% If UserHasAnyUnviewedNotes(CustIDPassed) = "True" Then %>
				<a href="#All" role="tab" data-toggle="tab" id="0"><span class="unread-note" id="notebg0">All (<%= GetCustNoteCountByNoteType(0,CustIDPassed,0) %>)</span></a>
			<% Else %>
				<a href="#All" role="tab" data-toggle="tab" id="0"><span class="read-note" id="notebg0">All (<%= GetCustNoteCountByNoteType(0,CustIDPassed,0) %>)</span></a>
			<% End If %>
      </li>
	
		<%		
		
		Set rsNoteTypes = Server.CreateObject("ADODB.Recordset")
		rsNoteTypes.CursorLocation = 3 
	
		SQLNoteTypes = "SELECT * FROM SC_NoteType ORDER BY InternalRecordIdentifier"		
		
		Set cnnNoteTypes = Server.CreateObject("ADODB.Connection")
		cnnNoteTypes.open (Session("ClientCnnString"))
		Set rsNoteTypes = cnnNoteTypes.Execute(SQLNoteTypes)
		
		If NOT rsNoteTypes.EOF Then
			Do While NOT rsNoteTypes.EOF
				NoteTypeName = rsNoteTypes("NoteType")	
				'NoteTypeDivID = GetTerm(NoteTypeName)
				NoteTypeDivID = Replace(NoteTypeName, "/", "")	
				NoteTypeIntRecID = rsNoteTypes("InternalRecordIdentifier")
						
				If HasNoteTypeBeenViewedByUser(CustIDPassed,NoteTypeIntRecID,Session("UserNo")) = "True" Then	
					%><li><a href="#<%= NoteTypeDivID %>" role="tab" data-toggle="tab" id="<%= NoteTypeIntRecID %>"><span class="read-note" id="notebg<%= NoteTypeIntRecID %>"><%= GetTerm(NoteTypeName) %>&nbsp;(<%= GetCustNoteCountByNoteType(NoteTypeIntRecID,CustIDPassed,0) %>)</span></a></li><%
				Else
					%><li><a href="#<%= NoteTypeDivID %>" role="tab" data-toggle="tab" id="<%= NoteTypeIntRecID %>"><span class="unread-note" id="notebg<%= NoteTypeIntRecID %>"><%= GetTerm(NoteTypeName) %>&nbsp;(<%= GetCustNoteCountByNoteType(NoteTypeIntRecID,CustIDPassed,0) %>)</span></a></li><%
				End If
				
				rsNoteTypes.MoveNext
			Loop
		End If
		
		Set rsNoteTypes = Nothing
		cnnNoteTypes.Close
		Set cnnNoteTypes = Nothing
		
		%>

	    </ul>
	    
	    <!-- Tab panes -->
	    <div class="tab-content">
	    
	    
	      <div class="tab-pane fade active in" id="All" style="padding:20px;">
				<div id="notes-all">
					<p>
						<button type="button" class="btn btn-success" onclick="ajaxRowNewCustomerNotes('0');"><i class="fas fa-user-edit"></i>&nbsp;Create New Customer Note</button>
					</p>
				
					<div class="input-group narrow-results" style="margin-bottom:20px"> <span class="input-group-addon"><i class="fas fa-search"></i>&nbsp;Search Notes</span>
					    <input id="filter-notes-all" type="text" class="form-control filter-search-width" placeholder="Type here...">
					</div>
						
					<div class="table-responsive">
				            <table id="ajaxContainerCustomerNotesTable0" class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
				              <thead>
				                <tr>
			                      <th width="10%">Type</th>
				                  <th width="10%">Date</th>
								  <th width="10%">Time</th>
								  <th width="10%">Entered By</th>
								  <th>Details</th>
								  <th width="15%">Reason</th>
				                  <th class="sorttable_nosort text-center" style="width: 80px;">Actions</th>
				                </tr>
				              </thead>
				
							<tbody id="ajaxContainerCustomerNotes0" class='searchable-notes-all ajax-loading'></tbody>
						</table>
					</div>
				</div>
	      </div>
	      
	      
		<%		
		
		Set rsNoteTypes = Server.CreateObject("ADODB.Recordset")
		rsNoteTypes.CursorLocation = 3 
	
		SQLNoteTypes = "SELECT * FROM SC_NoteType ORDER BY InternalRecordIdentifier"		
		
		Set cnnNoteTypes = Server.CreateObject("ADODB.Connection")
		cnnNoteTypes.open (Session("ClientCnnString"))
		Set rsNoteTypes = cnnNoteTypes.Execute(SQLNoteTypes)
		
		If NOT rsNoteTypes.EOF Then
			Do While NOT rsNoteTypes.EOF
				NoteTypeName = rsNoteTypes("NoteType")
				NoteTypeDivID = Replace(NoteTypeName, "/", "")
				NoteTypeIntRecID = rsNoteTypes("InternalRecordIdentifier")	
				NoteTypeCanBeCreatedByUser = rsNoteTypes("NoteTypeCanBeCreatedByUser")	
				%>
			      <div class="tab-pane fade" id="<%= NoteTypeDivID %>" style="padding:20px;">

					<div id="notes-<%= NoteTypeDivID %>">
					
						<% If NoteTypeCanBeCreatedByUser = 1 Then %>
							<p>
								<button type="button" class="btn btn-success" onclick="ajaxRowNewCustomerNotes('<%= NoteTypeIntRecID %>');"><i class="fas fa-user-edit"></i>&nbsp;Create New <%= GetTerm(NoteTypeName) %> Note</button>
							</p>
						<% End If %>
					
						<div class="input-group narrow-results" style="margin-bottom:20px"> <span class="input-group-addon"><i class="fas fa-search"></i>&nbsp;Search <%= GetTerm(NoteTypeName) %> Notes</span>
						    <input id="filter-notes-with-<%= NoteTypeIntRecID %>" type="text" class="form-control filter-search-width" placeholder="Type here...">
						</div>
							
						<div class="table-responsive">
					            <table id="ajaxContainerCustomerNotesTable<%= NoteTypeIntRecID %>" class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
					              <thead>
					                <tr>
				                      <th width="10%">Type</th>
					                  <th width="10%">Date</th>
									  <th width="10%">Time</th>
									  <th width="10%">Entered By</th>
									  <th>Details</th>
									  <th width="15%">Reason</th>
					                  <th class="sorttable_nosort text-center" style="width: 80px;">Actions</th>
					                </tr>
					              </thead>
					
								<tbody id="ajaxContainerCustomerNotes<%= NoteTypeIntRecID %>" class='searchable-notes-<%= NoteTypeIntRecID %> ajax-loading'></tbody>
							</table>
						</div>
					</div>
			      </div>
				<%
				rsNoteTypes.MoveNext
			Loop
		End If
		
		Set rsNoteTypes = Nothing
		cnnNoteTypes.Close
		Set cnnNoteTypes = Nothing
		
		%>
	      
	</div>
	    
								
	<%'**********************
	' **** eof Notes Tab ****
	'************************
	%>
	
	<script>
		
		var NoteTypes = <%= NoteTypes %>;
		
		$(document).ready(function () { 
		
			//This special code removes the initial yellow unread note highlighting that you see when the modal first loads
			$("span.unread-note").each(function () {
		       $(this).removeClass("unread-note");
		       $(this).addClass("read-note");  
		    });
		
			//First load the default ALL NOTES tab
			ajaxLoadCustomerNotes(); 
			
			//when you show a new notes tab, update the hidden text field that
			//stores the int rec id of that note type, so we always know which
			//note type we are currently viewing
			
			//this value gets passed to customerNote.asp, so the proper notes
			//are loaded for each tab
			$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
			
				var target = $(e.target).attr("href") // activated tab
				var targetNoteIntRecID = $(e.target).attr("id") // activated tab
				
				if (targetNoteIntRecID == '') {
					targetNoteIntRecID = 0;
				}
				
				//The isnumeric check is to ensure that the tab's ID is numeric, which means it is a customer note tab
				//This prevents other tabs in the background calling page from being affected
				if ($.isNumeric(targetNoteIntRecID)) {
					
					$("#txtNoteTypeIntRecIDToShow").val(targetNoteIntRecID);
					
					var custID = $("#txtCustIDToShow").val();
					
					var BaseURL = $("#txtBaseURL").val();
					
					//When a user clicks on tab, mark all the notes in that category as having been read	
					//Remove any highlight styles from the tab for unread notes	
		
					$.ajax({
						type:"POST",
						url: BaseURL + "inc/InSightFuncs_AjaxForARAP.asp",
						cache: false,
						data: "action=MarkAllNotesForNoteTypeForUserAsRead&NoteIntRecID=" + encodeURIComponent(targetNoteIntRecID) + "&CustID=" + encodeURIComponent(custID),
						success: function(response)
						{
							//alert("action=MarkNoteTypeForUserAsRead&NoteIntRecID=" + encodeURIComponent(targetNoteIntRecID) + "&CustID=" + encodeURIComponent(custID));				
							$("#notebg"+targetNoteIntRecID).removeClass("unread-note");
							$("#notebg"+targetNoteIntRecID).addClass("read-note");	
						}
					});
	
					ajaxLoadCustomerNotes();
				 }
			});			
			
		});
		
		

		$("#notesall").change(function() {
			updateNoteTabCounts();
		  	ajaxLoadCustomerNotes();
		});
		
		
		
		function updateNoteTabCounts() {
		
				var targetNoteIntRecID = $("#txtNoteTypeIntRecIDToShow").val();
				if (targetNoteIntRecID == '') {
					targetNoteIntRecID = 0; //default to all notes tab if no value
				}
				
				//the customer ID is stored in a hidden text field in the modal
				var custID = $("#txtCustIDToShow").val();
				
				//the current user no is stored in a hidden text field in the modal
				var userNo = $("#txtUserNoToShow").val();
				
				//check to see if the checkbox  is checked				
				if ($('#notesall').is(":checked")) {
					var ShowOnlyMyNotes = 1 
				}
				else  {
					var ShowOnlyMyNotes = 0
				}
		
				//go through each tab and change the html to reflect the note count changes
				$('a[data-toggle="tab"]').each(function() {
				
					//alert("in data toggle tab each");
					
				
					//the value of the anchor tag is the note type name
					var href = $(this).attr('href');
					
					//strip the # sign off of the href attribute to get the note type name
					var noteTypeName = href.substring(1, href.length)
					
					//the id is the internal record identifier of the note type in SC_Notes
					var noteTypeIntRecID = $(this).attr('id');
					
					var BaseURL = $("#txtBaseURL").val();
					
					if ($.isNumeric(noteTypeIntRecID)) {
											
						//now post to ajax funcs to get the new note counts and update the counts on the tabs					
						$.ajax({
							type:"POST",
							url: BaseURL + "inc/InSightFuncs_AjaxForARAP.asp",
							cache: false,
							data: "action=GetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes&NoteIntRecID=" + encodeURIComponent(noteTypeIntRecID) + "&CustID=" + encodeURIComponent(custID) + "&ShowOnlyMyNotes=" + encodeURIComponent(ShowOnlyMyNotes) + "&userNo=" + encodeURIComponent(userNo),
							success: function(response)
							{
								values = response.split('*');
								NoteGetTermName = values[0];
								NoteTypeCount = values[1];	
								NotesHaveBeenRead = values[2];	
								
								//alert("action=GetCustomerNoteTypeCountAndGetTermNameAndUnreadNotes&NoteIntRecID=" + encodeURIComponent(noteTypeIntRecID) + "&CustID=" + encodeURIComponent(custID) + "&ShowOnlyMyNotes=" + encodeURIComponent(ShowOnlyMyNotes) + "&userNo=" + encodeURIComponent(userNo));
								//alert("have the notes been read for " + NoteGetTermName + " (id: " + noteTypeIntRecID + ")? " + NotesHaveBeenRead);		
										
								if (NotesHaveBeenRead == 'True') {	
									$("#notebg"+noteTypeIntRecID).removeClass("unread-note");
									$("#notebg"+noteTypeIntRecID).addClass("read-note");	
									$("#notebg"+noteTypeIntRecID).text(NoteGetTermName + " (" + NoteTypeCount + ")");
							 		//$("#"+noteTypeIntRecID).html("<span class='read-note'>" + NoteGetTermName + " (" + NoteTypeCount + ")</span>");								
							 	}
							 	else {
							 		
							 		$("#notebg"+noteTypeIntRecID).removeClass("read-note");
							 		$("#notebg"+noteTypeIntRecID).addClass("unread-note");
							 		$("#notebg"+noteTypeIntRecID).text(NoteGetTermName + " (" + NoteTypeCount + ")");
							 		//$("#"+noteTypeIntRecID).html("<span class='unread-note'>" + NoteGetTermName + " (" + NoteTypeCount + ")</span>");
							 	}
							}
						});
					}
				});

		}	
		
	
		
		function ajaxRowNewCustomerNotes(passedDefaultNoteIntRecID) {
			var value = {};
			value.id = 0;
			value.NoteTypeGetTerm = ""; 
			value.NoteTypeIntRecID = passedDefaultNoteIntRecID; 
			value.NoteTypeCanBeCreatedByUser = 1;
			value.Date = "-";
			value.Time = "-";
			value.User = "-";
			value.CustomerNote = "";
			$('#ajaxRowCustomerNotes-' + 0 + '').remove();		
			
			NoteTypeIntRecIDToShow = $("#txtNoteTypeIntRecIDToShow").val();
					
			$("#ajaxContainerCustomerNotes" + NoteTypeIntRecIDToShow).prepend(ajaxRowHtmlNotes(value));
		}
		
		
		
		function ajaxRowHtmlNotes(value) {
		
			var NoteTypesSelect = '<select class="form-control" data-type="NoteTypes" name="txtNoteTypes" id="txtNoteTypeTab' + value.id + '">';		
			$.each(NoteTypes, function (key, NoteType) {
				NoteTypesSelect +='<option value="'+NoteType.id+'" ' + (value.NoteTypeIntRecID+""==NoteType.id+""?'selected':'') + '>'+NoteType.title+'</option>';
			});		
			NoteTypesSelect +='</select>';
			
			NoteTypeIntRecIDToShow = $("#txtNoteTypeIntRecIDToShow").val();
			
			if (value.NoteTypeCanBeCreatedByUser == 1) {
			
				//The code below is meant to fix conflicts between the note ID on the ALL tab and the note ID on its tab by note type
				if (NoteTypeIntRecIDToShow == 0) {
					var rowID = value.id + "000000000";
					var btns = '\
								<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'CustomerNotes\', ' + rowID + ', \'Edit\');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxLoadCustomerNotes(\'delete\', ' + rowID + ');"><i class="fas fa-trash-alt"></i></a></div>\
								<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadCustomerNotes(\'save\', ' + rowID + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'CustomerNotes\', ' + rowID + ', \'View\');"><i class="fa fa-times"></i></a></div>\
							';
				}
				else {
					var btns = '\
								<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'CustomerNotes\', ' + value.id + ', \'Edit\');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxLoadCustomerNotes(\'delete\', ' + value.id + ');"><i class="fas fa-trash-alt"></i></a></div>\
								<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadCustomerNotes(\'save\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'CustomerNotes\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
							';
				}				
			}
			else {
				var btns = '\
							<div class="visibleRowView btn-group btn-group-sm">&nbsp;</div>\
							<div class="visibleRowEdit btn-group btn-group-sm">&nbsp;</div>\
						';
			}
						
			if(value.id==0) {
				//The code below is meant to fix conflicts between the note ID on the ALL tab and the note ID on its tab by note type
				if (NoteTypeIntRecIDToShow == 0) {
					var rowID = value.id + "000000000";
					btns = '\
							<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadCustomerNotes(\'insert\', ' + rowID + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'CustomerNotes\', ' + rowID + ', \'View\');"><i class="fa fa-times"></i></a></div>\
						';
				}
				else {
					btns = '\
							<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadCustomerNotes(\'insert\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'CustomerNotes\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
						';
				}				
			}	

			var reg = /(MCS shortage)([\d\.\,]+)/;

			var result = value.CustomerNote.match(reg);
			//alert(result);
			if (result != null) {
				if (typeof(result[1]) != "undefined") {				
					value.CustomerNote = value.CustomerNote.replace("MCS shortage", "MCS shortage $");
					if (typeof(result[2]) != undefined) {
						value.CustomerNote = value.CustomerNote.replace(result[2], result[2].toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","));
					}
				}
			}
			
			//The code below is meant to fix conflicts between the note ID on the ALL tab and the note ID on its tab by note type
			if (NoteTypeIntRecIDToShow == 0) {
				var rowID = value.id + "000000000";
				var html = '<tr id="ajaxRowCustomerNotes-' + rowID + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">'
			}
			else {
				var html = '<tr id="ajaxRowCustomerNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">'
			}
			

				html += '<td>\
					<div class="visibleRowView">' + value.NoteTypeGetTerm + '</div>\
					<div class="visibleRowEdit">' + NoteTypesSelect + '</div>\
				</td>\
				<td>' + value.Date + '</td>\
				<td>' + value.Time + '</td>\
				<td>' + value.User + '</td>';
				
				html += '<td>\
					<div class="visibleRowView">' + value.CustomerNote + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="CustomerNote" value="' + value.CustomerNote.replace(/"/g, '&quot;') + '" /></div>\
					</td>';


				if (value.hasReason == 1) {
					if (typeof(value.Reason) == "undefined") {
						html += '<td>&nbsp;</td>';
					} else {
						html += '<td>'+ value.Reason +'</td>';
					}
				}
				else {
					html += '<td>&nbsp;</td>';
				}												
				html += '<td class="text-center">'+btns+'</td>\
		   </tr>\
			';	
						
			return html;
		}
		
		
		
		
		function ajaxLoadCustomerNotes(updateAction, updateActionId) {
				
			updateCustID = $("#txtCustIDToPassToGenerateNotes").val();
			updateNoteTypeIntRecID = $("#txtNoteTypeIntRecIDToShow").val();			
			notesall = $('#notesall').is(":checked");
			if (updateAction == "delete" && !confirm("Are you sure you want to delete this customer note?")) return;
			var BaseURL = $("#txtBaseURL").val();
			var url = BaseURL + "inc/customerNotesLoadAjax.asp?custID=" + updateCustID + "&notesall=" + notesall + "&notetypeintrecid=" + updateNoteTypeIntRecID;
			
			//alert(url);
			$("#ajaxContainerCustomerNotes" + updateNoteTypeIntRecID).addClass("ajax-loading");
			
			var jsondata = {};
			jsondata.updateAction = updateAction;
			jsondata.updateActionId = updateActionId;
			jsondata.updateCustID = updateCustID;

			if (updateAction=="insert"){
				if (updateNoteTypeIntRecID == 0) {
					var rowID = "0000000000";
					jsondata.NoteType = $('#ajaxRowCustomerNotes-' + rowID + ' [data-type="NoteTypes"]').val();
					jsondata.CustomerNote= $('#ajaxRowCustomerNotes-' + rowID + ' [data-type="CustomerNote"]').val();
				}
				else {
					jsondata.NoteType = $('#ajaxRowCustomerNotes-' + updateActionId + ' [data-type="NoteTypes"]').val();
					jsondata.CustomerNote= $('#ajaxRowCustomerNotes-' + updateActionId + ' [data-type="CustomerNote"]').val();			
				}
			}

			if (updateAction=="save"){
				jsondata.NoteType = $('#ajaxRowCustomerNotes-' + updateActionId + ' [data-type="NoteTypes"]').val();
				jsondata.CustomerNote= $('#ajaxRowCustomerNotes-' + updateActionId + ' [data-type="CustomerNote"]').val();		
			}

			//alert("updateActionId/ACtion/NoteType/Note: " + updateActionId + "/" + updateAction + "/" + jsondata.NoteType + "/" + jsondata.CustomerNote);
			
			
			$.ajax({
				type: "POST",
				url: url,
				dataType: "json",
				data: jsondata,
				success: function (data) {	
				
					//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }		
					var html = "";
					
					updateNoteTypeIntRecID = $("#txtNoteTypeIntRecIDToShow").val();
					if (updateNoteTypeIntRecID == '') {
					  	updateNoteTypeIntRecID = 0;
					}
										
					$.each(data, function (key, value) {
						html += ajaxRowHtmlNotes(value);
					});
					
					$("#ajaxContainerCustomerNotes" + updateNoteTypeIntRecID).html(html);
					
					updateNoteTabCounts();
					
					var newTableObject = document.getElementById("ajaxContainerCustomerNotesTable" + updateNoteTypeIntRecID);
					sorttable.makeSortable(newTableObject);
					
					setTimeout(function(){
						$("#ajaxContainerCustomerNotes" + updateNoteTypeIntRecID).removeClass("ajax-loading");
					}, 0);
						
					
				}
			});
			
		}
		
		
	</script>	
	
	 <!-- custom table search !-->
	
	<script>
	
		$(document).ready(function () {
			
		    (function ($) {
		    
		    	//This filter notes function is for specific note type tabs
		    	//It looks for search boxes that have id's that start with "filter-notes-with-"
		    	//Then it grabs the IntRecID of the note type for the search box and appends it to the filter
		        
		        $('[id^=filter-notes-with-]').keyup(function () {
		
					updateNoteTypeIntRecID = $("#txtNoteTypeIntRecIDToShow").val();
					if (updateNoteTypeIntRecID == '') {
					  	updateNoteTypeIntRecID = 0;
					}
		
		            var rex = new RegExp($(this).val(), 'i');
		            $('.searchable-notes-' + updateNoteTypeIntRecID + ' tr').hide();
		            $('.searchable-notes-' + updateNoteTypeIntRecID + ' tr').filter(function () {
		                return rex.test($(this).text());
		            }).show();
		        })
		 
		        $('#filter-notes-all').keyup(function () {
		
		            var rex = new RegExp($(this).val(), 'i');
		            $('.searchable-notes-all tr').hide();
		            $('.searchable-notes-all tr').filter(function () {
		                return rex.test($(this).text());
		            }).show();
		        })
		
		    }(jQuery));
		
		});
	</script>
	<!-- eof custom table search !-->
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