<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/Insightfuncs.asp"-->

    <script type="text/javascript" language="javascript" src="<%= baseURL %>inc/jquery/jquery-1.7.2.min.js"></script>
    

    <link type="text/css" href="<%= baseURL %>inc/jquery/jquery-ui-1.7.2.custom.css" rel="stylesheet" />
	<script type="text/javascript" language="javascript" src="<%= baseURL %>/inc/jquery/jquery-1.3.2.js"></script>
    <script type="text/javascript" language="javascript" src="<%= baseURL %>inc/jquery/ui.core.js"></script>
 	<script type="text/javascript" language="javascript" src="<%= baseURL %>inc/jquery/ui.dialog.js"></script>
	<script type="text/javascript" language="javascript" src="<%= baseURL %>inc/jquery/jquery.bgiframe.min.js"></script> 
	<script type="text/javascript" language="javascript" src="<%= baseURL %>inc/func.js"></script>



<%

'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
sURL = Request.ServerVariables("SERVER_NAME")

If MUV_Inspect("ConStmt-StartDate") = False  Then 'Didnt come from Ajax, just picked any field


		StartDate = Request.Form("txtStartDate")
		EndDate = Request.Form("txtEndDate")
		SelectedPeriod = Request.Form("selPeriod")
		If Request.Form("chkZeroDollar") = "on" then
			SkipZeroDollar = True
		Else
			SkipZeroDollar = False
		End If
		If Request.Form("chkLessThanZero") = "on" then
			SkipLessThanZero = True
		Else
			SkipLessThanZero = False
		End If
		If Request.Form("chkLessThanZeroLines") = "on" then
			SkipLessThanZeroLines = True
		Else
			SkipLessThanZeroLines = False
		End If

		IncludedType = ""
		If Request.Form("chkBackOrder") = "on" Then IncludedType = IncludedType & "B"
		If Request.Form("chkCreditMemo") = "on" Then IncludedType = IncludedType & "C"
		If Request.Form("chkSimpleDebit") = "on" Then IncludedType = IncludedType & "E"
		If Request.Form("chkRental") = "on" Then IncludedType = IncludedType & "G"
		If Request.Form("chkRouteInvoice") = "on" Then IncludedType = IncludedType & "I"
		If Request.Form("chkInterest") = "on" Then IncludedType = IncludedType & "O"
		If Request.Form("chkTelsel") = "on" Then IncludedType = IncludedType & "T"
		
		Account = Request.Form("txtCustIDToPass")
		CustomOrPredefined =  Request.Form("optCustomOrPredefined")
		If CustomOrPredefined = "Predefined" Then
			'Set start & end date
			StartDate = Left(SelectedPeriod,Instr(SelectedPeriod,"~")-1)
			EndDate = Right(SelectedPeriod,len(SelectedPeriod)-Instr(SelectedPeriod,"~"))
		End If
		
		DuesDateDaysOrSingleDate =  Request.Form("radInvoiceDueDate")
		
		If DuesDateDaysOrSingleDate = "SINGLEDATE" Then
			DueDateDays = ""
			DueDateSingleDate = Request.Form("txtDueDate")
		Else
			DueDateDays = Request.Form("selDueDate")
			DueDateSingleDate = ""
		End If
		
		DoNotShowDueDate = Request.Form("chkDoNotShowDueDate")
		
		If DoNotShowDueDate = "on" OR DoNotShowDueDate = "1" OR DoNotShowDueDate = "true" Then
			DoNotShowDueDate = "CHECKED"
		End If
		
		
		dummy = MUV_Write("ConStmt-StartDate",StartDate) '0
		dummy = MUV_Write("ConStmt-EndDate",EndDate) '1
		dummy = MUV_Write("ConStmt-SelectedPeriod",SelectedPeriod) '2
		dummy = MUV_Write("ConStmt-SkipZeroDollar",SkipZeroDollar)	'3
		dummy = MUV_Write("ConStmt-SkipLessThanZero",SkipLessThanZero) '4
		dummy = MUV_Write("ConStmt-IncludedType",IncludedType) '5
		dummy = MUV_Write("ConStmt-CustomOrPredefined",CustomOrPredefined) '5
		dummy = MUV_Write("ConStmt-Account",Account) '6
		dummy = MUV_Write("ConStmt-IncludeIndividuals",IncludeIndividuals) '7
		dummy = MUV_Write("ConStmt-DueDateDays",DueDateDays) '8
		dummy = MUV_Write("ConStmt-DueDateSingleDate",DueDateSingleDate) '9
		dummy = MUV_Write("ConStmt-DoNotShowDueDate",DoNotShowDueDate) '10
		dummy = MUV_Write("ConStmt-SkipLessThanZeroLines",SkipLessThanZeroLines) '11
		
		
Else
		'We came from Ajax

		StartDate = MUV_Read("ConStmt-StartDate")
		EndDate = MUV_Read("ConStmt-EndDate")
		SelectedPeriod = MUV_Read("ConStmt-SelectedPeriod")
		If MUV_Read("ConStmt-SkipZeroDollar") = "True" Then SkipZeroDollar = True Else SkipZeroDollar = False
		If MUV_Read("ConStmt-SkipLessThanZero") = "True" Then SkipLessThanZero = True Else SkipLessThanZero = False
		If MUV_Read("ConStmt-IncludeIndividuals") = "True" Then IncludeIndividuals = True Else IncludeIndividuals = False
		If MUV_Read("ConStmt-SkipLessThanZeroLines") = "True" Then SkipLessThanZeroLines = True Else SkipLessThanZeroLines = False
		IncludedType = MUV_Read("ConStmt-IncludedType")
		CustomOrPredefined = MUV_Read("ConStmt-CustomOrPredefined")
		Account = MUV_Read("ConStmt-Account")
		DueDateDays = MUV_Read("ConStmt-DueDateDays")
		DueDateSingleDate = MUV_Read("ConStmt-DueDateSingleDate")
		DoNotShowDueDate = MUV_Read("ConStmt-DoNotShowDueDate")		
		

End If

Description = MUV_Read("DisplayName") & " ran the report: #4 - Consolidated Invoice By Location for " & GetTerm("account") & " # " & Account 
Description = Description & " - " & GetCustNameByCustNum(Account)
CreateAuditLogEntry "A/R Report","A/R Report","Minor",0 ,Description
%>

<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
		<title>Consolidated Statement By Location</title>
		
		<style type="text/css">
			body{
				padding:30px;
				max-width: 1170px;
				margin: 0 auto;
				font-family: arial;
				font-size: 14px;
				line-height: 1.4;
			}
			
			table{
				border-collapse: collapse;
				text-align: left;
 
			}
			
			table th,td{
				font-weight: normal;
				vertical-align: middle;
			}
			
			h1,h2,h3,h4,h5,h6{
				margin: 0px;
			}
			
			.consolidated-invoice-number-date-account{
				font-size: 16px;
				margin-top: 10px;
			}
			
			.sold-to{
				background: #ccc;
				text-transform: uppercase;
				padding: 10px;
 			}
			
			.sold-to-address{
				background: #f5f5f5;
				padding: 10px;
				text-transform: uppercase;
			}
			
			#sold-to-table{
				margin-top: 30px;
			}
			
			#general-table-margin{
				margin-top: 5px;
			}
			
			.reset-omitted{
				display: inline-block;
				padding: 10px 15px 10px 15px;
				background: #f0ad4e;
				color: #fff;
 				cursor: pointer;
				border: 0px;
				border-radius:5px;
				font-size: 14px;
 			}
 			
 			.reset-omitted:hover{
	 			opacity:0.8;
 			}
 			
 			.generate-pdf{
				display: inline-block;
				padding: 10px 15px 10px 15px;
				background: #5bc0de;
				color: #fff;
 				cursor: pointer;
				border: 0px;
				border-radius:5px;
				font-size: 14px;
				margin-left: 10px;
 			}
 			
 			.generate-pdf:hover{
	 			opacity:0.8;
 			}
 			
 			.invoice-main-titles{
	 			background: #f5f5f5;
	 			border: 1px solid #000;
	 			font-size: 16px;
   			}
		
			.invoice-main-body{
  				margin-top: 20px;
			}
			
			.invoice-date-customer-po-line{
  				font-size: 16px;
  				border-bottom:1px solid #ccc;
				margin-top: 20px;
			}
			
			.table-subtotal{
  				border-top: 3px solid #000;
  				border-bottom: 3px solid #000;
  				margin-top: 20px;
			}
			
			.table-total{
   				font-size: 16px;
			}
  			
  			.grand-total{
	  				border-top: 3px solid #000;
 	  				margin-top: 20px;
 	  				font-size: 16px;
 	  				background: yellow;
   			}
		
  			.grand-total-message{
	  				
	  				border-top: 0px;
	  				border-bottom: 3px solid #000;
 	  				margin-top: 0px;
 	  				font-size: 16px;
 	  				background: yellow;
   			}
	  
	   
	   .thead-titles{
 		   background: #eaeaea;
 		   border-bottom: 3px solid #000;
 	   }
 	   
 	   .tr-lines{
	 	   border-bottom:1px solid #999;
	 	   line-height: 40px;
	 	   vertical-align:middle;
 	   }
 	   
  	   .tr-lines-main-title{	
			background: #ADD8E6;
			border: 3px solid #000;
			font-size: 16px;
			line-height: 40px;
			vertical-align:middle;
			margin-top:20px;
 	   }

 	   
 	   .tr-lines:hover{
	 	   background: #f5f5f5;
 	   }
		</style>

	</head>
	
	<body>
	
	<!-- header starts here !-->
	<table width="100%" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<!-- logo / address !-->
				<th scope="col" width="70%">
					<table>
						<tbody>
							<tr>
								<th scope="col"><img src="../../../clientfiles/<%=MUV_Read("ClientID")%>/logos/logo.png"></th>
								<th scope="col"><%
									SQL = "SELECT * FROM Settings_CompanyID"
									
									Set cnn8 = Server.CreateObject("ADODB.Connection")
									cnn8.open (Session("ClientCnnString"))
									Set rs = Server.CreateObject("ADODB.Recordset")
									rs.CursorLocation = 3 
									Set rs = cnn8.Execute(SQL)
									
									If not rs.EOF Then
										Attention = rs("Stmt_Attn")
										CompanyName = rs("Stmt_CompanyName")
										Address1 = rs("Stmt_Address1")
										Address2 = rs("Stmt_Address2")
										City = rs("Stmt_City")
										State = rs("Stmt_State")
										Zip = rs("Stmt_Zip")
										Phone1 = rs("Stmt_Phone1")
										Phone2 = rs("Stmt_Phone2")
										Phone3 = rs("Stmt_Phone3")
										Fax = rs("Stmt_Fax")
										Email = rs("Stmt_Email")
										Attention = rs("Stmt_Attn")
										MessageToPrint = rs("Stmt_Message")
									End If
									
									set rs = Nothing
									cnn8.close
									set cnn8 = Nothing
									If Attention <> "" Then Response.Write("Attn: " & Attention & "<br>")
									If CompanyName <> "" Then Response.Write(CompanyName & "<br>")
									If Address1 <> "" Then Response.Write(Address1 & "<br>")
									If Address2 <> "" Then Response.Write(Address2 & "<br>")
									If City <> "" Then Response.Write(City & ", ")
									If State <> "" Then Response.Write(State & " ")
									If Zip <> "" Then Response.Write(Zip & "<br>")	
									If Phone1 <> "" Then Response.Write(Phone1 & "<br>")															
									If Phone2 <> "" Then Response.Write(Phone2 & "<br>")															
									If Phone3 <> "" Then Response.Write(Phone3 & "<br>")															
									If Fax <> "" Then Response.Write("Fax:" & Fax & "<br>")																						
									If Email <> "" Then Response.Write(Email & "<br>")																						
								%></th>
							</tr>
						</tbody>
					</table>
				</th>
				<!-- eof logo  / address !-->			
			 					
	 					
	 			<!-- consolidated invoice - number - date - account !-->
	 			<th scope="col" width="30%">
		 			<h2 align="center">Consolidated Invoice</h2>
	
		 			
		 			<table width="100%" class="consolidated-invoice-number-date-account">
			 			<tbody>
				 			<tr>
					 			<th scope="col" width="70%"><strong>Consolidated Invoice Number:</strong></th>
					 			<th scope="col" width="30%" align="right"><strong><% Response.Write(Trim(Account) & Trim(Replace(EndDate,"/",""))) %></strong></th>
				 			</tr>
			 			</tbody>
		 			</table>
		 			
		 			<table width="100%" class="consolidated-invoice-number-date-account">
			 			<tbody>
				 			<tr>
					 			<th scope="col" width="40%"><strong>Invoice Dates:</strong></th>
					 			<th scope="col" width="60%" align="right"><strong><%=StartDate%> - <%=EndDate%></strong></th>
				 			</tr>
			 			</tbody>
		 			</table>
		 			
		 			<table width="100%" class="consolidated-invoice-number-date-account">
			 			<tbody>				 			
				 			<tr>
					 			<th scope="col" width="40%"><strong>Account Number:</strong></th>
					 			<th scope="col" width="60%" align="right"><strong><%= Account %></strong></th>
				 			</tr>
			 			</tbody>
		 			</table>
		 			
	 			
		 			
	 			</th>
	 			<!-- eof consolidated invoice - number - date - account !-->			
	 					
					</tr>
			</tbody>
		</table>
		<!-- header ends here !-->
		
		
		<!-- sold to box !-->
		<table width="30%" cellpadding="0" cellspacing="0" id="sold-to-table">
			<tbody>
				
				<tr>
					<th scope="col" width="100%" class="sold-to"><h3>SOLD TO: <%=GetTerm("Account")%> # <%=Account%></h3></th>
				</tr>
				
				<tr>
					<th scope="col" width="100%" class="sold-to-address">
						<% 
							SQLBillTo = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".AR_CustomerBillto Where CustNum = '" & Account &"'"
							Set cnnBillTo = Server.CreateObject("ADODB.Connection")
							cnnBillTo.open (Session("ClientCnnString"))
							Set rsBillTo = Server.CreateObject("ADODB.Recordset")
							rsBillTo.CursorLocation = 3 
							Set rsBillTo = cnnBillTo.Execute(SQLBillTo)
							If NOT rsBillTo.EOF Then %>
								<strong><%= rsBillTo("BillName")%></strong><br>	
								<%= rsBillTo("Addr1")%><br>
								<%= rsBillTo("Addr2")%><br>		
								<%= rsBillTo("City")%>, <%= rsBillTo("State")%>&nbsp;<%= rsBillTo("Zip")%><br>
							<%
							Else
							
								SQLBillToSecondary = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".AR_Customer Where CustNum = '" & Account &"'"
								Set cnnBillToSecondary = Server.CreateObject("ADODB.Connection")
								cnnBillToSecondary.open (Session("ClientCnnString"))
								Set rsBillToSecondary = Server.CreateObject("ADODB.Recordset")
								rsBillToSecondary.CursorLocation = 3 
								Set rsBillToSecondary = cnnBillToSecondary.Execute(SQLBillToSecondary)
								
								If NOT rsBillToSecondary.EOF Then %>
									<strong><%= rsBillToSecondary("Name")%></strong><br>	
									<%= rsBillToSecondary("Addr1")%><br>
									<%= rsBillToSecondary("Addr2")%><br>		
									<%= rsBillToSecondary("CityStateZip")%><br>
								<%
								End If
								
								Set rsBillToSecondary = Nothing
								cnnBillToSecondary.Close
								Set cnnBillToSecondary = Nothing


							End If
							Set rsBillTo = Nothing
							cnnBillTo.Close
							Set cnnBillTo = Nothing
						%>
						</th>
				</tr>
			</tbody>
		</table>
		<!-- eof sold to box !-->
		
		<!-- reset / generate buttons !-->
		<table width="100%" cellpadding="0" cellspacing="0" id="general-table-margin">
			<tbody>
				<tr>
					<th scope="col" width="100%">

						<%
						
						linkVar="<a href='consolidated_stmt_frm_acct_PDFlaunch.asp?s=" & Replace(StartDate,"/","~") & "&e=" & Replace(EndDate,"/","~") & "&c=" & Account 
						
						If SkipZeroDollar = True Then linkVar = linkVar & "&z=T" Else linkVar = linkVar & "&z=F"
						If SkipLessThanZero = True Then linkVar = linkVar & "&lz=T" Else linkVar = linkVar & "&lz=F"
						If IncludeIndividuals = True Then linkVar = linkVar & "&ind=T" Else linkVar = linkVar & "&ind=F"
						If IncludedType <> "" Then linkVar = linkVar & "&ty=" & IncludedType 
						If SkipLessThanZeroLines = True Then linkVar = linkVar & "&lzl=T" Else linkVar = linkVar & "&lzl=F"
						linkVar = linkVar & "&ddd=" & DueDateDays
						linkVar = linkVar & "&dds=" & DueDateSingleDate
						linkVar = linkVar & "&dnsdd=" & DoNotShowDueDate
						linkVar = linkVar &  "'>"
						Response.Write(linkVar)
						%>
						<button type="button" class="generate-pdf">Generate PDF</button>
						</a>
					</th>
				</tr>
			</tbody>
		</table>
		<!-- eof reset / generate buttons !-->
		
	
			
		<!-- the actual invoice starts here !-->
		<table width="100%" cellpadding="0" cellspacing="0" id="general-table-margin">

				<% 'Now get the actual invoice data
				
					SQLInvoices = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistory WHERE CustNum = '" & Account &"' "
					SQLInvoices = SQLInvoices & "AND IvsDate >= '" & StartDate & "' AND IvsDate <= '" & EndDate & "' "
					
					If SkipZeroDollar = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt <> 0 "
					If SkipLessThanZero = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt > 0 "
					If IncludedType <> "" Then SQLInvoices = SQLInvoices & "AND CHARINDEX(IvsType,'" & IncludedType & "') <> 0 "

					SQLInvoices = SQLInvoices & " ORDER BY CustNum, IvsNum"
					
					'Response.Write(SQLInvoices & "<br><br>")
	
					Set cnnInvoices = Server.CreateObject("ADODB.Connection")
					cnnInvoices.open (Session("ClientCnnString"))
					Set rsInvoices = Server.CreateObject("ADODB.Recordset")
					rsInvoices.CursorLocation = 3 
					Set rsInvoices = cnnInvoices.Execute(SQLInvoices)
					If not rsInvoices.Eof Then
					
						Set cnnTmpTable = Server.CreateObject("ADODB.Connection")
						cnnTmpTable.open (Session("ClientCnnString"))
						Set rsTmpTable = Server.CreateObject("ADODB.Recordset")
						rsTmpTable.CursorLocation = 3 

						AccountSubtotal = 0
						TotalAmt = 0
						TotalTax = 0

						Do While NOT rsInvoices.EOF
								
								If HeldCust <> rsInvoices("CustNum") Then
								
									HeldCust = rsInvoices("CustNum")

									SQLBillTo = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".AR_CustomerBillto Where CustNum = '" & rsInvoices("CustNum") &"'"
									Set cnnBillTo = Server.CreateObject("ADODB.Connection")
									cnnBillTo.open (Session("ClientCnnString"))
									Set rsBillTo = Server.CreateObject("ADODB.Recordset")
									rsBillTo.CursorLocation = 3 
									Set rsBillTo = cnnBillTo.Execute(SQLBillTo)
									
									If NOT rsBillTo.EOF Then 
										CustName = rsBillTo("BillName")
										Address1 = rsBillTo("Addr1")
										Address2 = rsBillTo("Addr2")
										City = rsBillTo("City") & ", " & rsBillTo("State") & "&nbsp;" & rsBillTo("Zip")
									Else
									
										SQLBillToSecondary = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".AR_Customer Where CustNum = '" & rsInvoices("CustNum") &"'"
										Set cnnBillToSecondary = Server.CreateObject("ADODB.Connection")
										cnnBillToSecondary.open (Session("ClientCnnString"))
										Set rsBillToSecondary = Server.CreateObject("ADODB.Recordset")
										rsBillToSecondary.CursorLocation = 3 
										Set rsBillToSecondary = cnnBillToSecondary.Execute(SQLBillToSecondary)
										
										If NOT rsBillToSecondary.EOF Then 
											CustName = rsBillToSecondary("Name")
											Address1 = rsBillToSecondary("Addr1")
											Address2 = rsBillToSecondary("Addr2")
											City = rsBillToSecondary("CityStateZip")
										End If
										Set rsBillToSecondary = Nothing
										cnnBillToSecondary.Close
										Set cnnBillToSecondary = Nothing
									End If
									
									Set rsBillTo = Nothing
									cnnBillTo.Close
									Set cnnBillTo = Nothing
						
									
									If AccountSubtotal <> 0 Then ' not on the first one but print the subtotals%>
										<tr class="tr-lines" style="background-color:#FFFF99;border:1px gray solid;">
											<th scope="col" colspan="5" align="right" class="invoice-date"><strong>Subtotal <%=FormatCurrency(AccountSubtotal)%></strong></th>
										</tr>
										<tr><th colspan="5" style="height:50px;">&nbsp;</th><tr>
										<% AccountSubtotal = 0
									End If%>
										<tr class="tr-lines-main-title">
											<th scope="col" colspan="5" class="invoice-date" align="center"><strong>ACCOUNT #<%= rsInvoices("CustNum") %>&nbsp;&nbsp;<%= CustName %></strong></th>
										</tr>
										<tr class="tr-lines-main-title">
											<th scope="col" colspan="5" class="invoice-date" align="center"><strong><%= Address1 %>&nbsp;&nbsp;<%= Address2 %>&nbsp;&nbsp;<%= City %></strong></th>
										</tr>
										
										
										<tr><th colspan="5">							
											<!-- the actual invoice starts here !-->
											<table width="100%" cellpadding="0" cellspacing="0" id="general-table-margin">
												<!-- titles !-->
												<tr>
													<th scope="col" width="100%">
														<table  width="100%" cellpadding="10" cellspacing="0">
															
															<tbody class="invoice-main-titles">
																<tr>
																	<th scope="col" width="20%" valign="middle"><strong>SKU</strong></th>
																	<th scope="col" width="48%" valign="middle"><strong>DESCRIPTION</strong></th>
																	<th scope="col" width="15%" valign="middle"><strong>UNIT PRICE</strong></th>
																	<th scope="col" width="10%" valign="middle"><strong>QTY</strong></th>
																	<th scope="col" width="10%" valign="middle" align ="right"><strong>EXT PRICE</strong></th>
																</tr>
															</tbody>
									 					</table>
													</th>
												</tr>
												<!-- eof titles !-->
											</table>
										</th></tr>
										
								
								
								
									<%
									'*******************************************************************************************************
									'FIRST GET ALL THE INVOICE NUMBERS FROM THE INVOICE HISTORY TABLE FOR THE CURRENT CUSTOMER ACCOUNT NUMBER,
									'AND WITH THE CURRENT SELECTION CRITERIA MATCHED. STORE THIS INTO AN ARRAY OF INVOICE NUMBERS
									'FOR THAT ACCOUNT/LOCATION
									'*******************************************************************************************************
				
									SQLGetInvoicesInDateRange = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistory Where "
									SQLGetInvoicesInDateRange = SQLGetInvoicesInDateRange & "IvsDate >= '" & StartDate & "' AND IvsDate <= '" & EndDate & "' "
									SQLGetInvoicesInDateRange = SQLGetInvoicesInDateRange & "AND CustNum = " & HeldCust & " "
									If SkipZeroDollar = True Then SQLGetInvoicesInDateRange = SQLGetInvoicesInDateRange & "AND IvsTotalAmt <> 0 "
									If SkipLessThanZero = True Then SQLGetInvoicesInDateRange = SQLGetInvoicesInDateRange & "AND IvsTotalAmt > 0 "
									If IncludedType <> "" Then SQLGetInvoicesInDateRange = SQLGetInvoicesInDateRange & "AND CHARINDEX(IvsType,'" & IncludedType & "') <> 0 "
									SQLGetInvoicesInDateRange = SQLGetInvoicesInDateRange & " ORDER BY IvsNum"
									
									'Response.Write(SQLGetInvoicesInDateRange & "<br><br>")
					
									Set cnnGetInvoicesInDateRange = Server.CreateObject("ADODB.Connection")
									cnnGetInvoicesInDateRange.open (Session("ClientCnnString"))
									Set rsGetInvoicesInDateRange = Server.CreateObject("ADODB.Recordset")
									rsGetInvoicesInDateRange.CursorLocation = 3 
									Set rsGetInvoicesInDateRange = cnnGetInvoicesInDateRange.Execute(SQLGetInvoicesInDateRange)
									If not rsGetInvoicesInDateRange.Eof Then
									
										Set cnnTmpTable = Server.CreateObject("ADODB.Connection")
										cnnTmpTable.open (Session("ClientCnnString"))
										Set rsTmpTable = Server.CreateObject("ADODB.Recordset")
										rsTmpTable.CursorLocation = 3 
										
										invoiceNumArrayString = ""
																				
										DO WHILE NOT rsGetInvoicesInDateRange.EOF
											If rsGetInvoicesInDateRange("IvsNum") <> "" Then
												If invoiceNumArrayString = "" Then
													invoiceNumArrayString = rsGetInvoicesInDateRange("IvsNum")
												Else
													invoiceNumArrayString = invoiceNumArrayString & "," & rsGetInvoicesInDateRange("IvsNum")
												End If
											End If
											rsGetInvoicesInDateRange.MoveNext
										LOOP
									End If
									
									'*******************************************************************************************************
									'NEXT, GET THE TOTALS FOR EVERY PRODUCT SKU WITHIN THOSE INVOICES, BUT PRESENT IT AS ONE CONSOLIDATED
									'INVOICE WITH INDIVIDUAL SKUS
									'*******************************************************************************************************
									
									SQLDateRangeSKUCount = " SELECT CustNum, partnum, Description, Price, ItemQTY, Expr1 "
									SQLDateRangeSKUCount = SQLDateRangeSKUCount & " FROM (SELECT CustNum, MAX(partNum) AS partnum, MAX(prodDescription) AS Description, MAX(itemPrice) AS Price, "
									SQLDateRangeSKUCount = SQLDateRangeSKUCount & " SUM(itemQuantity) AS ItemQTY, SUM(itemPrice * itemQuantity) AS Expr1 FROM InvoiceHistoryDetail "
									
									
									WHERE_CLAUSE_INVOICES = ""
									
									locationInvoiceNumArray= ""
									locationInvoiceNumArray = Split(invoiceNumArrayString,",")
			
									
									For z = 0 to UBound(locationInvoiceNumArray)
										If z = 0 Then
											If z = UBound(locationInvoiceNumArray) Then
												WHERE_CLAUSE_INVOICES = " WHERE ((ivsNum = " & Trim(locationInvoiceNumArray(z)) & ")) "
											Else
												WHERE_CLAUSE_INVOICES = " WHERE ((ivsNum = " & Trim(locationInvoiceNumArray(z)) & ") "
											End If
										Else
											If z = UBound(locationInvoiceNumArray) Then
												WHERE_CLAUSE_INVOICES = WHERE_CLAUSE_INVOICES & " OR (ivsNum = " & Trim(locationInvoiceNumArray(z)) & ")) "
											Else
												WHERE_CLAUSE_INVOICES = WHERE_CLAUSE_INVOICES & " OR (ivsNum = " & Trim(locationInvoiceNumArray(z)) & ") "
											End If
										End If
									Next	
									
									If WHERE_CLAUSE_INVOICES <> "" Then
										SQLDateRangeSKUCount = SQLDateRangeSKUCount & WHERE_CLAUSE_INVOICES
									End If
									
							
									SQLDateRangeSKUCount = SQLDateRangeSKUCount & " GROUP BY partNum,CustNum) "
									SQLDateRangeSKUCount = SQLDateRangeSKUCount & " AS derivedtbl_1 "
									If SkipLessThanZeroLines = True Then SQLDateRangeSKUCount = SQLDateRangeSKUCount & " WHERE (Expr1 <> 0)  "
									SQLDateRangeSKUCount = SQLDateRangeSKUCount & "ORDER BY partnum,CustNum	"					
									
									'Response.Write(SQLDateRangeSKUCount & "<br>")
								
									Set cnnDateRangeSKUCount = Server.CreateObject("ADODB.Connection")
									cnnDateRangeSKUCount.open (Session("ClientCnnString"))
									Set rsDateRangeSKUCount = Server.CreateObject("ADODB.Recordset")
									rsDateRangeSKUCount.CursorLocation = 3 
									Set rsDateRangeSKUCount = cnnDateRangeSKUCount.Execute(SQLDateRangeSKUCount)
									
									
									If NOT rsDateRangeSKUCount.EOF Then
									
										Do While NOT rsDateRangeSKUCount.EOF
										
											%>
													<tr class="tr-lines">
														<th scope="col" class="invoice-date" width="20%"><%= rsDateRangeSKUCount("partnum") %></th>
														<th scope="col" class="invoice-nr" width="50%"><%= rsDateRangeSKUCount("Description") %></th>
														<th scope="col" width="15%"><%= formatCurrency(rsDateRangeSKUCount("Price"),2) %></th>
														<th scope="col" width="10%"><%= rsDateRangeSKUCount("ItemQTY") %></th>
														<th scope="col" align ="right" class="amount" width="10%"><%= formatCurrency(rsDateRangeSKUCount("Expr1"),2) %></th>
													</tr>			
												<%
			
											rsDateRangeSKUCount.MoveNext
										Loop
									End If
									

						End If 'HELD CUST
						
									
						
						TotalAmt = TotalAmt + rsInvoices("IvsTotalAmt")
						TotalTax = TotalTax + rsInvoices("IvsSalesTax")
						AccountSubtotal = TotalAmt - TotalTax
	
						rsInvoices.movenext
					Loop

				End If
				Set rsInvoices = Nothing
				cnnInvoices.Close
				Set cnnInvoices = Nothing
				%>
		</table>
		<!-- the table ends here !-->

			<!-- grand total !-->
			<table  width="100%" cellpadding="10" cellspacing="0" class="grand-total">

				<tr>
					<th scope="col" width="80%" align="right"><strong>SUBTOTAL: </strong></th>
					<th scope="col" width="20%" align="right"><strong><%= FormatCurrency(AccountSubtotal) %></strong></th>
				</tr>

				<tr>
					<th scope="col" width="80%" align="right"><strong>TAX: </strong></th>
					<th scope="col" width="20%" align="right"><strong><%= FormatCurrency(TotalTax) %></strong></th>
				</tr>
			
				<tr>
					<th scope="col" width="80%" align="right"><strong>TOTAL: </strong></th>
					<th scope="col" width="20%" align="right"><strong><%= FormatCurrency(TotalAmt) %></strong></th>
				</tr>

				<% If DoNotShowDueDate <> "CHECKED" Then %>
					<% If DueDateSingleDate <> "" Then %>
						<tr>
							<th scope="col" width="80%" align="right"><strong>INVOICE DUE DATE:  </strong></th>
							<th scope="col" width="20%" align="right"><strong><%= FormatDateTime(DueDateSingleDate,2) %></strong></th>
						</tr>
					<% Else %>
						<tr>
							<th scope="col" width="80%" align="right"><strong>INVOICE DUE DATE:  </strong></th>
							<th scope="col" width="20%" align="right"><strong><%= DateAdd("d",DueDateDays,EndDate) %></strong></th>
						</tr>
					<% End If %>
				<% End If %>	
									
			</table>

			<table  width="100%" cellpadding="10" cellspacing="0" class="grand-total-message">

				<tr>
					<th scope="col" width="100%" align="center"><strong><%= UCASE(MessageToPrint) %></strong></th>
				</tr>
			</table>
			<!-- eof grand total !-->
	</body>
	
</html>