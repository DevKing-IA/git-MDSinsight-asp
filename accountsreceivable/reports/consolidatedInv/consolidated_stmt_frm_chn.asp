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
						window.location = "consolidated_stmt_frm_chn.asp";
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
						window.location = "consolidated_stmt_frm_chn.asp";
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

If MUV_Inspect("ConStmt-StartDate") = False Then 'Didnt come from Ajax, just picked any field

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
		If Request.Form("chkIncludeIndividuals") = "on" then
			IncludeIndividuals = True
		Else
			IncludeIndividuals = False
		End If

		IncludedType = ""
		If Request.Form("chkBackOrder") = "on" Then IncludedType = IncludedType & "B"
		If Request.Form("chkCreditMemo") = "on" Then IncludedType = IncludedType & "C"
		If Request.Form("chkSimpleDebit") = "on" Then IncludedType = IncludedType & "E"
		If Request.Form("chkRental") = "on" Then IncludedType = IncludedType & "G"
		If Request.Form("chkRouteInvoice") = "on" Then IncludedType = IncludedType & "I"
		If Request.Form("chkInterest") = "on" Then IncludedType = IncludedType & "O"
		If Request.Form("chkTelsel") = "on" Then IncludedType = IncludedType & "T"
		
		Chain = Request.Form("txtChainIDToPass")
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
		dummy = MUV_Write("ConStmt-Chain",Chain) '6
		dummy = MUV_Write("ConStmt-IncludeIndividuals",IncludeIndividuals) '7
		dummy = MUV_Write("ConStmt-DueDateDays",DueDateDays) '8
		dummy = MUV_Write("ConStmt-DueDateSingleDate",DueDateSingleDate) '9
		dummy = MUV_Write("ConStmt-DoNotShowDueDate",DoNotShowDueDate) '10
		
		
Else

		'We came from Ajax
		StartDate = MUV_Read("ConStmt-StartDate")
		EndDate = MUV_Read("ConStmt-EndDate")
		SelectedPeriod = MUV_Read("ConStmt-SelectedPeriod")
		If MUV_Read("ConStmt-SkipZeroDollar") = "True" Then SkipZeroDollar = True Else SkipZeroDollar = False
		If MUV_Read("ConStmt-SkipLessThanZero") = "True" Then SkipLessThanZero = True Else SkipLessThanZero = False
		If MUV_Read("ConStmt-IncludeIndividuals") = "True" Then IncludeIndividuals = True Else IncludeIndividuals = False
		IncludedType = MUV_Read("ConStmt-IncludedType")
		CustomOrPredefined = MUV_Read("ConStmt-CustomOrPredefined")
		Chain = MUV_Read("ConStmt-Chain")
		DueDateDays = MUV_Read("ConStmt-DueDateDays")
		DueDateSingleDate = MUV_Read("ConStmt-DueDateSingleDate")
		DoNotShowDueDate = MUV_Read("ConStmt-DoNotShowDueDate")			
		
End If

Description = MUV_Read("DisplayName") & " ran the report: #1 - Consolidated Invoice for chain # " & Chain 
Description = Description & " - " & GetCustNameByCustNum(Chain)
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
					 			<th scope="col" width="50%"><strong>Consolidated Invoice Number:</strong></th>
					 			<th scope="col" width="50%" align="right"><strong><% Response.Write(Trim(Chain) & Trim(Replace(EndDate,"/",""))) %></strong></th>
				 			</tr>
				 			
				 			<tr>
					 			<th scope="col" width="50%"><strong>Invoice Dates:</strong></th>
					 			<th scope="col" width="50%" align="right"><strong><%=StartDate%> - <%=EndDate%></strong></th>
				 			</tr>
				 			
				 			<tr>
					 			<th scope="col" width="50%"><strong>Chain Number:</strong></th>
					 			<th scope="col" width="50%" align="right"><strong><%= Chain %></strong></th>
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
					<th scope="col" width="100%" class="sold-to"><h3>SOLD TO: Chain # <%= Chain %><br><%= GetChainDescByChainNum(Chain) %></h3></th>
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
							linkVar="<a href='consolidated_stmt_frm_chn_PDFlaunch.asp?s=" & Replace(StartDate,"/","~") & "&e=" & Replace(EndDate,"/","~") & "&c=" & Chain
							
							If IncludeIndividuals = True Then linkVar = linkVar & "&ind=T" Else linkVar = linkVar & "&ind=F"
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
						
						<tbody class="invoice-main-titles">
							<tr>
								<th scope="col" width="10%" valign="middle"><strong>Omit</strong></th>
								<th scope="col" width="20%" valign="middle"><strong>Invoice Date</strong></th>
								<th scope="col" width="40%" valign="middle"><strong>Invoice #</strong></th>
								<th scope="col" width="30%" valign="middle" align ="right"><strong>Amount</strong></th>
							</tr>
						</tbody>
 					</table>
				</th>
			</tr>
			<!-- eof titles !-->
		</table>
			
		<!-- the actual invoice starts here !-->
		<table width="100%" cellpadding="0" cellspacing="0" id="general-table-margin">

				<% 'Now get the actual invoice data
					SQLInvoices = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistory Where "
					SQLInvoices = SQLInvoices & "IvsDate >= '" & StartDate & "' AND IvsDate <= '" & EndDate & "' "
					SQLInvoices = SQLInvoices & "And CustNum In "
					SQLInvoices = SQLInvoices & "(Select CustNum From AR_Customer Where ChainNum = " & Chain & ") "
					
					If SkipZeroDollar = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt <> 0 "
					If SkipLessThanZero = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt > 0 "
					If IncludedType <> "" Then SQLInvoices = SQLInvoices & "AND CHARINDEX(IvsType,'" & IncludedType & "') <> 0 "

					SQLInvoices = SQLInvoices & "AND IvsHistSequence NOT IN (Select IvsHistSequence from zReportConsolidatedInvoiceOmit_" & Trim(Session("userNo")) & ") "
					
					SQLInvoices = SQLInvoices & " order by CustNum, IvsNum"
					'Response.Write(SQLInvoices)
	
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

						HeldCust = ""
						ChainSubtotal = 0
						TotalAmt = 0
						Do While not rsInvoices.Eof%>
								<% If HeldCust <> rsInvoices("CustNum") Then
									HeldCust = rsInvoices("CustNum")

									SQLBillTo = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".AR_CustomerBillto Where CustNum = '" & rsInvoices("CustNum") &"'"
									Set cnnBillTo = Server.CreateObject("ADODB.Connection")
									cnnBillTo.open (Session("ClientCnnString"))
									Set rsBillTo = Server.CreateObject("ADODB.Recordset")
									rsBillTo.CursorLocation = 3 
									Set rsBillTo = cnnBillTo.Execute(SQLBillTo)
									If NOT rsBillTo.EOF Then 
										Address1 = rsBillTo("Addr1")
										City = rsBillTo("City") & ", " & rsBillTo("State") & "&nbsp;" & rsBillTo("Zip")
									Else
									
										SQLBillToSecondary = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".AR_Customer Where CustNum = '" & rsInvoices("CustNum") &"'"
										Set cnnBillToSecondary = Server.CreateObject("ADODB.Connection")
										cnnBillToSecondary.open (Session("ClientCnnString"))
										Set rsBillToSecondary = Server.CreateObject("ADODB.Recordset")
										rsBillToSecondary.CursorLocation = 3 
										Set rsBillToSecondary = cnnBillToSecondary.Execute(SQLBillToSecondary)
										
										If NOT rsBillToSecondary.EOF Then 
											Address1 = rsBillToSecondary("Addr1")
											City = rsBillToSecondary("CityStateZip")
										End If
										Set rsBillToSecondary = Nothing
										cnnBillToSecondary.Close
										Set cnnBillToSecondary = Nothing
		
		
									End If
									Set rsBillTo = Nothing
									cnnBillTo.Close
									Set cnnBillTo = Nothing
						
									
									If ChainSubtotal <> 0 Then ' not on the first one but print the subtotals%>
										<tr class="tr-lines" style="background-color:#FFFF99;border:1px gray solid;">
										<th scope="col" colspan="4" align="right" class="invoice-date"><strong>Subtotal <%=FormatCurrency(ChainSubtotal)%></strong></th>
										</tr>
										<% ChainSubtotal = 0
									End If%>
										<tr class="tr-lines" style="background-color:#ccc;border:1px gray solid;">
										<th scope="col" colspan="4" class="invoice-date" align="center"><strong>Account# <%=rsInvoices("CustNum")%>&nbsp;&nbsp;-&nbsp;&nbsp;<%=Address1%>,<%=City%></strong></th>
										</tr>
								<% End If %>
								<tr class="tr-lines">
									<%Response.Write("<th scope='col' width='12%'><a href='#'><input type='checkbox' name='chk'" & rsInvoices("IvsNum") & "' id='chk" & rsInvoices("IvsNum") & "' onclick='myFunction(" & rsInvoices("IvsHistSequence") & ")')></a></th>")%>
									<th scope="col" class="invoice-date" width="20%"><%=Month(rsInvoices("IvsDate")) & "/" & Day(rsInvoices("IvsDate")) & "/" & Year(rsInvoices("IvsDate"))%></th>
									<th scope="col" class="invoice-nr" width="40%"><%=rsInvoices("IvsNum")%></th>
									<th scope="col" align ="right" class="amount" width="30%"><%=FormatCurrency(rsInvoices("IvsTotalAmt"))%></th>
									<% 
									ChainSubtotal = ChainSubtotal + rsInvoices("IvsTotalAmt")
									TotalAmt = TotalAmt + rsInvoices("IvsTotalAmt")%>
								</tr>
								<%
								SQLTmpTable = "INSERT INTO zReportConsolidatedInvoiceInclude_" & Trim(Session("UserNo")) & " (IvsHistSequence) VALUES ('" & rsInvoices("IvsHistSequence") & "')"
								Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)

							rsInvoices.movenext
							
							'********************************************
							'Code to print the last chain subtotal
							'********************************************
							If rsInvoices.EOF Then
							%>
								<tr class="tr-lines" style="background-color:#FFFF99;border:1px gray solid;">
								<th scope="col" colspan="4" align="right" class="invoice-date"><strong>Subtotal <%=FormatCurrency(ChainSubtotal)%></strong></th>
								</tr>
							<%					
							End If
							'********************************************
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