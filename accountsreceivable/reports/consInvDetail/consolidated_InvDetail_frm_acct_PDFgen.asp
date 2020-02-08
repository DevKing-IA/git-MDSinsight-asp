<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/Insightfuncs.asp"-->

<%
dummy = MUV_Write("ClientID","") = "" 'Need this here

'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
sURL = Request.ServerVariables("SERVER_NAME")

StartDate = Request.QueryString("s")
EndDate = Request.QueryString("e")
Account = Request.QueryString("c")
StartDate = Replace(StartDate, "~","/")
EndDate = Replace(EndDate, "~","/")
Username = Request.QueryString("u")
Password = Request.QueryString("p")
ClientKey = Request.QueryString("cl")
UserNo = Request.QueryString("un")
DueDateDays = Request.QueryString("ddd")
DueDateSingleDate = Request.QueryString("dds")
DoNotShowDueDate = Request.QueryString("dnsdd")

SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open InsightCnnString
'Response.Write("InsightCnnString:" & InsightCnnString & "<br>")

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3
'Response.Write("SQL:" & SQL& "<br>")

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Automatic service ticket threashold report<%
	Response.End
Else
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & Recordset.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & Recordset.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	Recordset.close
	Connection.close	
End If	

%>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
		<title>Consolidated Statement</title>
		
		<style type="text/css">
			body{
   				font-family: arial;
				font-size: 14px;
				line-height: 1.4;
				width:100%;
				float: left;
			}
			
			table{
				border-collapse: collapse;
 
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
	   				margin-bottom: 20px;
   				}
  			
  			.grand-total{
	  				border-top: 3px solid #000;
 	  				margin-top: 20px;
    				}
			</style>
			
					
	</head>
	
	<body>
	
	<!-- header starts here !-->
		<table width="1024" style="width:1024px;" cellpadding="0" cellspacing="0">
			<tbody>
					<tr >
			
						<!-- logo / address !-->
						<th scope="col"   >
							<table align="left"  >
								<tbody>
									<tr >
										<th scope="col" align="left"><img src="../../../clientfiles/<%=MUV_Read("ClientID")%>/logos/logo.png"></th>
										<th scope="col" align="left" width="200px"><%
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
										%>
										</th>
									</tr>
								</tbody>
							</table>
						</th>
						<!-- eof logo  / address !-->			
			 					
			 					
			 			<!-- consolidated invoice - number - date - account !-->
			 			<th scope="col" width="30%" align="right">
				 			<h2 align="center">Consolidated Invoice</h2>
				 			
				 			<table width="100%" class="consolidated-invoice-number-date-account">
					 			<tbody>
						 			
						 			<tr>
							 			<th scope="col" width="50%" align="left"><strong>Invoice Number:	</strong></th>
							 			<th scope="col" width="50%" align="right"><p align="right"><strong><%Response.Write(Trim(Account) & Trim(Replace(EndDate,"/","")))%></strong></p></th>
						 			</tr>
						 			
						 			<tr>
							 			<th scope="col" width="50%" align="left"><strong>Invoice Date:</strong></th>
							 			<th scope="col" width="50%" align="right"><p align="right"><strong><%=EndDate%></strong></p></th>
						 			</tr>
						 			
						 			<tr>
							 			<th scope="col" width="50%" align="left"><strong>Account Number:	</strong></th>
							 			<th scope="col" width="50%" align="right"><p align="right"><strong><%=Account%></strong></p></th>
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
					<th scope="col" width="100%" class="sold-to-address" align="left">
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
		
		
		<!-- the actual invoice starts here !-->
		<table width="1024" cellpadding="0" cellspacing="0" id="general-table-margin" style="width:1024px;" >
			
			<!-- titles !-->
			<tr >
				<th scope="col" width="1024"  style="width:1024px;" align="left">
					
					<table  width="1024"   style="width:1024px;" cellpadding="10" cellspacing="0" bgcolor="#f5f5f5" style="border:3px solid #222;">
						
			<tbody class="invoice-main-titles" >
				<tr >
 					<th scope="col"  valign="middle" align="left" width="100" style="width:100px;"><strong>Item Number</strong></th>
					<th scope="col"  valign="middle" align="left" width="400" style="width:400px;"><strong>Description</strong></th>
					<th scope="col" valign="middle" align="left"  width="100" style="width:100px;"><strong>UOM</strong></th>
					<th scope="col"   valign="middle" align="left" width="100" style="width:100px;"><strong>QTY</strong></th>
					<th scope="col"   valign="middle" align="left" width="100" style="width:100px;"><strong>Unit Price</strong></th>
					<th scope="col"   valign="middle" align="left" width="100" style="width:100px;"><strong>Extended Total</strong></th>
				</tr>
			</tbody>
			
 		</table>
 		
 		
				</th>
		</tr>
			<!-- eof titles !-->
			
			<!-- invoice nr / date / po !-->
				
					
					<% 'Now get the actual invoice data
								SQLInvoices = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistory Where CustNum = '" & Account &"' "
								SQLInvoices = SQLInvoices & "AND IvsDate >= '" & StartDate & "' AND IvsDate <= '" & EndDate & "' "
								
								If SkipZeroDollar = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt <> 0 "
								If SkipLessThanZero = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt > 0 "
								
								If IncludedType <> "" Then SQLInvoices = SQLInvoices & "AND CHARINDEX(IvsType,'" & IncludedType & "') <> 0 "
			
								SQLInvoices = SQLInvoices & "AND IvsHistSequence NOT IN (Select IvsHistSequence from zReportConsolidatedInvoiceOmit_" & Trim(userNo) & ") "
								
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
									<tr>
				<th scope="col" width="1024" style="width:1024px;" align="left">
					<table  width="1024" style="width:1024px;" cellpadding="10" cellspacing="0" class="invoice-date-customer-po-line">
						
 				<tr>
	 			 

					 
					
					<th scope="col"  width='200' style='width:200px;' align="left">
						<strong>Invoice#:  <%=rsInvoices("IvsNum")%> </strong> 
					</th>
					
					<th scope="col" width='200' style='width:200px;'  align="left">
						<strong>Invoice Date:  <%=Month(rsInvoices("IvsDate")) & "/" & Day(rsInvoices("IvsDate")) & "/" & Year(rsInvoices("IvsDate"))%>
</strong>
					</th>
					
					<% If Trim(rsInvoices("PurchOrderNum")) = "" or IsNull(rsInvoices("PurchOrderNum")) Then %>
						
						<th scope="col" width='200' style='width:200px;' align="left">
					<strong>Customer PO#: N/A</strong>
					</th>
					
					<% Else %>
					<th scope="col" width='200' style='width:200px;'  align="left">
					<strong>Customer PO#:  <%=rsInvoices("PurchOrderNum")%></strong>
					</th>
					<% End If%>
					
					 
					
				</tr>
				
				</table>
				</th>
		</tr>
 				<!--eof invoice nr / date / po !-->

 <tr>
				<th scope='col' width='100%' align="left" >
					<table  width='100%' cellpadding='10' cellspacing='0'>
						<tbody class='invoice-main-body'>
					 									 								
											<%
											SQLTmpTable = "INSERT INTO zReportConsolidatedInvoiceInclude_" & Trim(UserNo) & " (IvsHistSequence) VALUES ('" & rsInvoices("IvsHistSequence") & "')"
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
													
 
 
													Response.Write("<th scope='col'   align='left' width='100' style='width:100px;'> " & rsInvoiceDetails("partnum") & " </th>")
													Response.Write("<th scope='col' align='left' width='440' style='width:440px;'>" & Replace(rsInvoiceDetails("prodDescription"),"<","") & " </th>")
													Response.Write("<th scope='col'   align='left' width='120' style='width:120px;'>" & rsInvoiceDetails("prodSalesUnit") & "</th>")
													Response.Write("<th scope='col'   align='left' width='140' style='width:140px;'>" & rsInvoiceDetails("itemQuantity") & "</th>")
													Response.Write("<th scope='col'  align='left' width='100' style='width:100px;'>" & FormatCurrency(rsInvoiceDetails("itemPrice")) & "</th>")
													Response.Write("<th scope='col'  align='left' width='100' style='width:100px;'>" & FormatCurrency(rsInvoiceDetails("itemQuantity") * rsInvoiceDetails("itemPrice")) & "</th>")
													Response.Write("</tr>")	
													SubTot = SubTot +		rsInvoiceDetails("itemQuantity") * rsInvoiceDetails("itemPrice")										
													rsInvoiceDetails.movenext
												Loop
											End If
											
											Response.Write("</tbody></table></th></tr>")
											
											'Now print the total info
											Response.Write("<tr><th scope='col' width='970' style='width:970px;' align='right'><table  width='400' style='width:400px;' cellpadding='0' cellspacing='0' align='right' ><th scope='col' width='400' style='width:400px;' align='right'  ><table  width='400' style='width:400px;' cellpadding='10' cellspacing='0' class='table-subtotal' align='right'  >")
											Response.Write("<tr><th scope='col' width='70%' align='right'><strong>SubTotal: </strong></th><th scope='col' width='30%' align='right'><strong>" & FormatCurrency(rsInvoices("IvsTotalAmt") - (rsInvoices("IvsSalesTax")+rsInvoices("IvsDepositChg"))) & "</strong></th></tr>")
											Response.Write("<tr><th scope='col' width='70%' align='right'><strong>Sales Tax: </strong></th><th scope='col' width='30%' align='right'><strong>" & FormatCurrency(rsInvoices("IvsSalesTax")) & "</strong></th></tr>")
											Response.Write("<tr><th scope='col' width='70%' align='right'><strong>Deposits: </strong></th><th scope='col' width='30%' align='right'><strong>" & FormatCurrency(rsInvoices("IvsDepositChg")) & "</strong></th></tr>")
											
											Response.Write("</table></th></table></tr>")	
											
											Response.Write("<tr><th scope='col' width='970' style='width:970px;' align='right'  >	<table  width='400' style='width:400px;' cellpadding='10' cellspacing='0' class='table-total' align='right'   >")	
											
											Response.Write("<tr><th scope='col' width='70%' align='right'><strong>Total For  Invoice " & rsInvoices("IvsNum") & "</strong></th><th scope='col' width='30%' align='right'><strong>   " & FormatCurrency(rsInvoices("IvsTotalAmt")) & " </strong></th></tr>")
 											
											
									Response.Write("</table></th>	</tr>")	
											
											set rsInvoiceDetails = Nothing
											cnnInvoiceDetails.close
											set cnnInvoiceDetails= Nothing

											TotalAmt = TotalAmt + rsInvoices("IvsTotalAmt")
											SQLTmpTable = "INSERT INTO zReportConsolidatedInvoiceInclude_" & Trim(UserNo) & " (IvsHistSequence) VALUES ('" & rsInvoices("IvsHistSequence") & "')"
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
			<table  width="1050" style="width:1050px;" bgcolor="#f5f5f5" cellpadding="15" cellspacing="0"  >
				
									<tr>
										<th scope="col"  width="1050" style="width:1050px; margin-top:20px;" align="center" class="grand-total" ><strong style="font-size:21px;">Total amount due:  <%=FormatCurrency(TotalAmt)%></strong></th>
 									</tr>
 									
									<% If DoNotShowDueDate <> "CHECKED" Then %>
										<% If DueDateSingleDate <> "" Then %>
											<tr>
												<th scope="col"  width="1050" style="width:1050px;" align="center" ><strong style="font-size:21px;">Invoice Due Date:  <%= FormatDateTime(DueDateSingleDate,2) %></strong></th>
		 									</tr>
										<% Else %>
											<tr>
												<th scope="col"  width="1050" style="width:1050px;" align="center" ><strong style="font-size:21px;">Invoice Due Date:  <%= DateAdd("d",DueDateDays,EndDate) %></strong></th>
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