<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/Insightfuncs.asp"-->
    <script type="text/javascript" language="javascript" src="<%= baseURL %>inc/jquery/jquery-1.7.2.min.js"></script>
    
    <link type="text/css" href="<%= baseURL %>inc/jquery/jquery-ui-1.7.2.custom.css" rel="stylesheet" />
	<script type="text/javascript" language="javascript" src="<%= baseURL %>/inc/jquery/jquery-1.3.2.js"></script>
    <script type="text/javascript" language="javascript" src="<%= baseURL %>inc/jquery/ui.core.js"></script>
 	<script type="text/javascript" language="javascript" src="<%= baseURL %>inc/jquery/ui.dialog.js"></script>
	<script type="text/javascript" language="javascript" src="<%= baseURL %>inc/jquery/jquery.bgiframe.min.js"></script> 
	<script type="text/javascript" language="javascript" src="<%= baseURL %>inc/func.js"></script>

<script>
  function myFunction(num)
	  {   

		  var  ivshistsequence=num;
          
		   if(num!='')
		   {
		    $.ajax({
		   type:'post',
		      url:'omitInvoice.asp',
		          data:{ivshistsequence: ivshistsequence},
					success: function(msg){
						window.location = "consolidated_InvDetail_frm_acct.asp";
					}
		 });
		  }
	}
</script>

<script>
  function myFunction2()
	  {   
		    $.ajax({
		   type:'post',
		      url:'omitReset.asp',
					success: function(msg){
						window.location = "consolidated_InvDetail_frm_acct.asp";
					}
		 });
	}
</script>
<%
Set cnnTmpTable = Server.CreateObject("ADODB.Connection")
cnnTmpTable.open (Session("ClientCnnString"))
Set rsTmpTable = Server.CreateObject("ADODB.Recordset")
rsTmpTable.CursorLocation = 3 
SQLTmpTable = "DELETE FROM zReportConsolidatedInvoiceInclude_" & Trim(Session("userNo")) 
Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)
set rsTmpTable = Nothing
cnnTmpTable.close
set cnnTmpTable = Nothing


'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
sURL = Request.ServerVariables("SERVER_NAME")

If Session("MultiUseVar") = "" Then 'Didnt come from Ajax
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
		
		'Response.Write("DoNotShowDueDate : " & DoNotShowDueDate & "<br>")
		
		Session("MultiUseVar") = ""
		Session("MultiUseVar") = Session("MultiUseVar") &  StartDate & "__"
		Session("MultiUseVar") = Session("MultiUseVar") &  EndDate & "__"
		Session("MultiUseVar") = Session("MultiUseVar") &  SelectedPeriod & "__"
		Session("MultiUseVar") = Session("MultiUseVar") &  SkipZeroDollar & "__"
		Session("MultiUseVar") = Session("MultiUseVar") &  SkipLessThanZero & "__"
		Session("MultiUseVar") = Session("MultiUseVar") &  IncludedType & "__"
		Session("MultiUseVar") = Session("MultiUseVar") &  CustomOrPredefined & "__"
		Session("MultiUseVar") = Session("MultiUseVar") &  Account & "__"	
		Session("MultiUseVar") = Session("MultiUseVar") &  SkipLessThanZeroLines & "__"
		Session("MultiUseVar") = Session("MultiUseVar") &  DueDateDays & "__"
		Session("MultiUseVar") = Session("MultiUseVar") &  DueDateSingleDate & "__"
		Session("MultiUseVar") = Session("MultiUseVar") &  DoNotShowDueDate & "__"
		'Response.Write("MultiUseVar1:" & Session("MultiUseVar") & ":MultiUseVar1<br><br>")
		
Else

		'We came from Ajax
		MultiUseVar = Split(Session("MultiUseVar"),"__")
		StartDate = MultiUseVar(0)
		EndDate = MultiUseVar(1)
		SelectedPeriod = MultiUseVar(2)
		If MultiUseVar(3) = "True" Then SkipZeroDollar = True Else SkipZeroDollar = False
		If MultiUseVar(4) = "True" Then SkipLessThanZero = True Else SkipLessThanZero = False
		If MultiUseVar(8) = "True" Then SkipLessThanZeroLines = True Else SkipLessThanZeroLines = False
		'Response.Write("SkipLessThanZero " & SkipLessThanZero & "<br>")
		'Response.Write("SkipLessThanZeroLines " & SkipLessThanZeroLines & "<br>")
		IncludedType = MultiUseVar(5)
		CustomOrPredefined = MultiUseVar(6)
		Account = MultiUseVar(7)
		DueDateDays = MultiUseVar(9)
		DueDateSingleDate = MultiUseVar(10) 
		DoNotShowDueDate = MultiUseVar(11) 
		
		'Response.Write("MultiUseVar2:" & Session("MultiUseVar") & ":MultiUseVar2<br><br>")
End If

Description = MUV_Read("DisplayName") & " ran the report: #3 - Detailed Consolidated Invoice for " & GetTerm("account") & " # " & Account 
Description = Description & " - " & GetCustNameByCustNum(Account)
CreateAuditLogEntry "A/R Report","A/R Report","Minor",0 ,Description

%>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
		<title>Consolidated Statement</title>
		
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
				vertical-align: top;
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
				margin-top: 30px;
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
	 			border: 3px solid #000;
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
							If Phone1 <> "" Then Response.Write(Phone1 & "    ")															
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
							 			<th scope="col" width="50%"><strong>Invoice Number:	</strong></th>
							 			<th scope="col" width="50%" align="right"><strong><%Response.Write(Trim(Account) & Trim(Replace(EndDate,"/","")))%></strong></th>
						 			</tr>
						 			
						 			<tr>
							 			<th scope="col" width="50%"><strong>Invoice Date:</strong></th>
							 			<th scope="col" width="50%" align="right"><strong><%=EndDate%></strong></th>
						 			</tr>
						 			
						 			<tr>
							 			<th scope="col" width="50%"><strong>Account Number:	</strong></th>
							 			<th scope="col" width="50%" align="right"><strong><%=Account%></strong></th>
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
					<th scope="col" width="100%" class="sold-to"><h3>Sold To</h3></th>
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
						<button type="button" class="reset-omitted" onclick='myFunction2();'>Reset Omitted</button>
				<%
				linkVar="<a href='consolidated_InvDetail_frm_acct_PDFlaunch.asp?s=" & Replace(StartDate,"/","~") & "&e=" & Replace(EndDate,"/","~") & "&c=" & Account 
				
				If SkipZeroDollar = True Then linkVar = linkVar & "&z=T" Else linkVar = linkVar & "&z=F"
				If SkipLessThanZero = True Then linkVar = linkVar & "&lz=T" Else linkVar = linkVar & "&lz=F"
				If SkipLessThanZeroLines = True Then linkVar = linkVar & "&lzl=T" Else linkVar = linkVar & "&lzl=F"
				If IncludedType <> "" Then linkVar = linkVar & "&ty=" & IncludedType 
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
			
			<!-- titles !-->
			<tr>
				<th scope="col" width="100%">
					<table  width="100%" cellpadding="10" cellspacing="0">
						
			<tbody class="invoice-main-titles" >
				<tr >
					<th scope="col" width="10%" valign="middle"><strong>Omit</strong></th>
					<th scope="col" width="10%"  valign="middle"><strong>Item Number</strong></th>
					<th scope="col" width="40%"  valign="middle"><strong>Description</strong></th>
					<th scope="col" width="10%"  valign="middle"><strong>UOM</strong></th>
					<th scope="col" width="10%"  valign="middle"><strong>QTY</strong></th>
					<th scope="col" width="10%"  valign="middle"><strong>Unit Price</strong></th>
					<th scope="col" width="10%"  valign="middle"><strong>Extended Total</strong></th>
				</tr>
			</tbody>
			
 		</table>
				</th>
		</tr>
			<!-- eof titles !-->
			
			<!-- invoice nr / date / po !-->
				<tr>
				<th scope="col" width="100%">
					
					<% 'Now get the actual invoice data
								SQLInvoices = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistory Where CustNum = '" & Account &"' "
								SQLInvoices = SQLInvoices & "AND IvsDate >= '" & StartDate & "' AND IvsDate <= '" & EndDate & "' "
								
								If SkipZeroDollar = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt <> 0 "
								If SkipLessThanZero = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt > 0 "
								
								If IncludedType <> "" Then SQLInvoices = SQLInvoices & "AND CHARINDEX(IvsType,'" & IncludedType & "') <> 0 "
			
								SQLInvoices = SQLInvoices & "AND IvsHistSequence NOT IN (Select IvsHistSequence from zReportConsolidatedInvoiceOmit_" & Trim(Session("userNo")) & ") "
								
								SQLInvoices = SQLInvoices & " order by IvsNum"
				
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

									TotalAmt = 0
									Do While not rsInvoices.Eof%>
					<table  width="100%" cellpadding="10" cellspacing="0" class="invoice-date-customer-po-line">
						
 				<tr>
	 				<%Response.Write("<th scope='col' width='8.5%'><a href='#'><input type='checkbox' name='chk'" & rsInvoices("IvsNum") & "' id='chk" & rsInvoices("IvsNum") & "' onclick='myFunction(" & rsInvoices("IvsHistSequence") & ")')></a></th>")%>

					 
					
					<th scope="col" width="15%">
						<strong>Invoice#:  <%=rsInvoices("IvsNum")%> </strong> 
					</th>
					
					<th scope="col" width="15%">
						<strong>Invoice Date:  <%=Month(rsInvoices("IvsDate")) & "/" & Day(rsInvoices("IvsDate")) & "/" & Year(rsInvoices("IvsDate"))%>
</strong>
					</th>
					
					<% If Trim(rsInvoices("PurchOrderNum")) = "" or IsNull(rsInvoices("PurchOrderNum")) Then %>
						
						<th scope="col" width="15%">
					<strong>Customer PO#: N/A</strong>
					</th>
					
					<% Else %>
					<th scope="col" width="15%">
					<strong>Customer PO#:  <%=rsInvoices("PurchOrderNum")%></strong>
					</th>
					<% End If%>
					
					<th scope="col" width="10%">
						&nbsp;
					</th>
					
					<th scope="col" width="10%">
						&nbsp;
					</th>
					
					<th scope="col" width="10%">
						&nbsp;
					</th>
					
				</tr>
				
				</table>
				</th>
		</tr>
 				<!--eof invoice nr / date / po !-->

 <tr>
				<th scope='col' width='100%' >
					<table  width='100%' cellpadding='10' cellspacing='0'>
						
			<tbody class='invoice-main-body'> 			 				 
							
									 								
											<%
											SQLTmpTable = "INSERT INTO zReportConsolidatedInvoiceInclude_" & Trim(Session("UserNo")) & " (IvsHistSequence) VALUES ('" & rsInvoices("IvsHistSequence") & "')"
											Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)
											
											'That did the header info, now we need to do the details
											SQLInvoiceDetails =  "Select * from InvoiceHistoryDetail WHERE "
											SQLInvoiceDetails = SQLInvoiceDetails & "InvoiceHistoryDetail.IvsHistSequence = " & rsInvoices("IvsHistSequence")
											
											If SkipLessThanZeroLines = True Then SQLInvoiceDetails = SQLInvoiceDetails & "AND InvoiceHistoryDetail.itemPrice <> 0 " 
											
											SQLInvoiceDetails = SQLInvoiceDetails & " order by IvsHistDetSequence"
											
											Set cnnInvoiceDetails = Server.CreateObject("ADODB.Connection")
											cnnInvoiceDetails.open (Session("ClientCnnString"))
											Set rsInvoiceDetails = Server.CreateObject("ADODB.Recordset")
											rsInvoiceDetails.CursorLocation = 3 
											Set rsInvoiceDetails = cnnInvoiceDetails.Execute(SQLInvoiceDetails)

											If not rsInvoiceDetails.Eof Then
												SubTot = 0
												Do While Not rsInvoiceDetails.eof
													Response.Write("<tr>")
													
													Response.Write("<th scope='col' width='10%'>&nbsp;</th>")

 
													Response.Write("<th scope='col' width='10%'> " & rsInvoiceDetails("partnum") & " </th>")
													Response.Write("<th scope='col' width='40%'>" & Replace(rsInvoiceDetails("prodDescription"),"<","") & "</th>")
													Response.Write("<th scope='col' width='10%'>" & rsInvoiceDetails("prodSalesUnit") & "</th>")
													Response.Write("<th scope='col' width='10%'>" & rsInvoiceDetails("itemQuantity") & "</th>")
													Response.Write("<th scope='col' width='10%'>" & FormatCurrency(rsInvoiceDetails("itemPrice")) & "</th>")
													Response.Write("<th scope='col' width='10%'>" & FormatCurrency(rsInvoiceDetails("itemQuantity") * rsInvoiceDetails("itemPrice")) & "</th>")
													Response.Write("</tr>")	
													SubTot = SubTot +		rsInvoiceDetails("itemQuantity") * rsInvoiceDetails("itemPrice")										
													rsInvoiceDetails.movenext
												Loop
											End If
											
											Response.Write("</tbody></table></th></tr>")
											
											'Now print the total info
											Response.Write("<tr><th scope='col' width='100%'><table  width='100%' cellpadding='0' cellspacing='0'  ><th scope='col' width='100%' align='right'><table  width='30%' cellpadding='10' cellspacing='0' class='table-subtotal'   >")
											Response.Write("<tr><th scope='col' width='70%' align='right'><strong>SubTotal: </strong></th><th scope='col' width='30%' align='right'><strong>" & FormatCurrency(rsInvoices("IvsTotalAmt") - (rsInvoices("IvsSalesTax")+rsInvoices("IvsDepositChg"))) & "</strong></th></tr>")
											Response.Write("<tr><th scope='col' width='70%' align='right'><strong>Sales Tax: </strong></th><th scope='col' width='30%' align='right'><strong>" & FormatCurrency(rsInvoices("IvsSalesTax")) & "</strong></th></tr>")
											Response.Write("<tr><th scope='col' width='70%' align='right'><strong>Deposits: </strong></th><th scope='col' width='30%' align='right'><strong>" & FormatCurrency(rsInvoices("IvsDepositChg")) & "</strong></th></tr>")
											
											Response.Write("</table></th>	</tr>")	
											
											Response.Write("<tr><th scope='col' width='100%' align='right'>	<table  width='30%' cellpadding='10' cellspacing='0' class='table-total'   >")	
											
											Response.Write("<tr><th scope='col' width='70%' align='right'><strong>Total For  Invoice " & rsInvoices("IvsNum") & "</strong></th><th scope='col' width='30%' align='right'><strong>   " & FormatCurrency(rsInvoices("IvsTotalAmt")) & " </strong></th></tr>")
 											
											
									Response.Write("</table></th>	</tr>")	
											
											set rsInvoiceDetails = Nothing
											cnnInvoiceDetails.close
											set cnnInvoiceDetails= Nothing

											TotalAmt = TotalAmt + rsInvoices("IvsTotalAmt")
											SQLTmpTable = "INSERT INTO zReportConsolidatedInvoiceInclude_" & Trim(Session("UserNo")) & " (IvsHistSequence) VALUES ('" & rsInvoices("IvsHistSequence") & "')"
											Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)

											rsInvoices.movenext
									Loop
									
									set rsTmpTable = Nothing
									cnnTmpTable.close
									set cnnTmpTable = Nothing

								End If
								Set rsInvoices = Nothing
								cnnInvoices.Close
								Set cnnInvoices = Nothing
							%>

						</table>
				</th>
			</tr>
			
			<!-- eof total  and subtotal !-->
			
			<!-- grand total !-->
			<table  width="100%" cellpadding="10" cellspacing="0" class="grand-total"   >
				
									<tr>
										<th scope="col" width="80%" align="right"><strong>Total amount due:  </strong></th>
										<th scope="col" width="20%" align="right"><strong><%=FormatCurrency(TotalAmt)%></strong></th>
									</tr>
									<% If DoNotShowDueDate <> "CHECKED" Then %>
										<% If DueDateSingleDate <> "" Then %>
											<tr>
												<th scope="col" width="80%" align="right"><strong>Invoice Due Date:  </strong></th>
												<th scope="col" width="20%" align="right"><strong><%= FormatDateTime(DueDateSingleDate,2) %></strong></th>
											</tr>
										<% Else %>
											<tr>
												<th scope="col" width="80%" align="right"><strong>Invoice Due Date:  </strong></th>
												<th scope="col" width="20%" align="right"><strong><%= DateAdd("d",DueDateDays,EndDate) %></strong></th>
											</tr>
										<% End If %>
									<% End If %>	
									
								</table>
			<!-- eof grand total !-->
			
		</table>
		<!-- the actual invoice ends here !-->
		
		<p align="center"><strong>*** End of Consolidated Invoice # <%=Trim(Account) & Trim(Replace(EndDate,"/",""))%>***</strong>
		
		
						 		
	</body>
	
</html>