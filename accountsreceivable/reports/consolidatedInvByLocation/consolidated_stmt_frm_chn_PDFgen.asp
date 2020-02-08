<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/Insightfuncs.asp"-->
<%
dummy = MUV_Write("ClientID","") 'Need this here

'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
sURL = Request.ServerVariables("SERVER_NAME")

StartDate = Request.QueryString("s")
EndDate = Request.QueryString("e")
Chain = Request.QueryString("c")
StartDate = Replace(StartDate, "~","/")
EndDate = Replace(EndDate, "~","/")
Username = Request.QueryString("u")
ClientKey = Request.QueryString("cl")
UserNo = Request.QueryString("un")
DueDateDays = Request.QueryString("ddd")
DueDateSingleDate = Request.QueryString("dds")
DoNotShowDueDate = Request.QueryString("dnsdd")

If Request.QueryString("z") = "T" then
	SkipZeroDollar = True
Else
	SkipZeroDollar = False
End If
If Request.QueryString("lz") = "T" then
	SkipLessThenZero = True
Else
	SkipLessThanZero = False
End If
If Request.QueryString("lzl") = "T" then
	SkipLessThanZeroLines = True
Else
	SkipLessThanZeroLines = False
End If

IncludedType = Request.QueryString("ty")

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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Consolidated Invoice By Location from Chain Report<%
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
	CompanyIdentityColor1 = rs("CompanyIdentityColor1")
	CompanyIdentityColor2 = rs("CompanyIdentityColor2")
	
	If CompanyIdentityColor1 = "" Then CompanyIdentityColor1 = "#6c7271"
	If CompanyIdentityColor2 = "" Then CompanyIdentityColor2 = "#6c7271"

End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
%>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
		<title>Consolidated Statement By Location</title>
		
		<style type="text/css">
		body{
			margin:30px;
 			font-family: Arial;
			font-size: 13px;
			overflow-x: hidden;
			text-align: left;
		}
		
		.line{
			width: 100%;
			float: left;
		}
		
		  table{
	 	   border-collapse: collapse;
	  
 	   }
		
		/* first statement */
		.first-statement{
			width: 100%;
			float: left;
 		}
		
	  .first-statement .logo,.address{
 		  float: left;
 		  margin-right: 20px;
	  }
	  
	   .first-statement .sold-to{
		   float: left;
		   padding: 15px;
		   background: #eaeaea;
		   margin: 20px 0px 20px 0px;
		   width: 280px;
	   }
	   
	   .first-statement .sold-to h3{
		   background: #ccc;
		   float: left;
		   width: 100%;
		   margin:-15px 0px 10px -15px;
		   padding: 15px;
	   }
	  
	  
	   .first-statement .main-heading{
		   width: 100%;
		   float: left;
		   background: #ccc;
		   text-align: center;
		   border-bottom: 3px solid #000;
	   }
	   
	   .first-statement .main-heading h1{
		   display: block;
		   text-transform: uppercase;
		   padding: 10px;
	   }
	   
	   .thead-titles{
 		   background: #eaeaea;
 		   border-bottom: 3px solid #000;
 	   }
 	   
 	   .tr-lines{
	 	   border-bottom:1px solid #999;
 	   }
 	   
 	   .tr-lines:hover{
	 	   background: #f5f5f5;
 	   }
 	   
 	   .first-statement .invoice-date,.invoice-nr,.amount{
	 	   width: 10%;
	 	   font-weight: normal;
 	   }
 	   
 	   .first-statement .blank-col{
	 	   width: 70%;
 	   }
 	 
 	 .first-statement .total{
	 	 background: yellow;
	 	 text-align: right;
	 	 width: 100%;
	 	 float: left;
 	 	 
	 	
 	 }
 	 
 	  .first-statement .total span{
	 	  text-transform: uppercase;
 	 	  display: block;
	 	  padding: 10px;
	 	   font-size: 19px;
	 	   font-weight: bold;
 	  }
 
		.sold-to-strong{
			<% Response.Write("color:" & CompanyIdentityColor1 & " !important;") %>
		}
 		
		</style>
		
	</head>
	
	<body>
		
		<!-- main table starts here !-->
 		<table width="650" align="center">
			<tbody >
				<tr>
					<td width="100%">
		
		<!-- logo / address / account starts here !-->
		<table width="850" style="margin-bottom:20px;">
			<tbody>
				<tr>
					
					<!-- logo !-->
					<th scope="col" align="left">
							<img src="../../../clientfiles/<%=MUV_Read("ClientID")%>/logos/logo.png">
					</th>
					<!-- eof logo !-->
					
  				</tr>
			</tbody>
		</table>
		
		<table width="850" border="1" bordercolor="#111111" cellpadding="4" style="margin-bottom:-1px;">
			<tbody>
				<tr>
					
					<!-- address !-->
					<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
						<%
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
						%>
					</th>
					<!-- eof address !-->
					
					<!-- sold to !-->
					<th scope="col" style="font-size:12px; font-weight:normal;" align="right">
						<!-- title !-->
						 <strong class="sold-to-strong">SOLD TO:  # <%=Chain%></strong><br>
						 <%=GetChainDescByChainNum(Chain)%>
						<!-- eof title !-->
						

											</th>
					<!-- eof sold to !-->
					
					</tr>
			</tbody>
		</table>
		<!-- logo / address / account ends here !-->
		
	
		<!-- monthly consolidated invoice title !-->
		<table width="850" cellpadding="4" style="margin-bottom:-1px;">
			<tbody>
				<tr>
					<th scope="col" >
						<h3 style="line-height:1; margin-top:10px; margin-bottom:10px;" align="center">Consolidated Invoice #<%Response.Write(Trim(Account) & Trim(Replace(EndDate,"/","")))%></h3>
					</th>
				</tr>
				<tr>
					<th scope="col" >
						<h3 style="line-height:1; margin-top:10px; margin-bottom:10px;" align="center">Invoice Dates: <small><%=StartDate%> - <%=EndDate%></small></h3>
					</th>
				</tr>
				
			</tbody>
		</table>
		<!-- eof monthly consolidated invoice title !-->
			
		<!-- the actual invoice starts here !-->
		<table width="850" cellpadding="0" cellspacing="0" style="margin-top: 5px;" border="1" bordercolor="#111111">

				<% 'Now get the actual invoice data
				
					SQLInvoices = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistory WHERE "
					SQLInvoices = SQLInvoices & "IvsDate >= '" & StartDate & "' AND IvsDate <= '" & EndDate & "' "
					SQLInvoices = SQLInvoices & "AND CustNum IN "
					SQLInvoices = SQLInvoices & "(SELECT CustNum FROM AR_Customer WHERE ChainNum = " & Chain & ") "
					
					If SkipZeroDollar = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt <> 0 "
					If SkipLessThanZero = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt > 0 "
					If IncludedType <> "" Then SQLInvoices = SQLInvoices & "AND CHARINDEX(IvsType,'" & IncludedType & "') <> 0 "
					
					SQLInvoices = SQLInvoices & " ORDER BY CustNum, IvsNum"
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
						TotalTax = 0
						TotaSubtotal = 0

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
						
									
									If ChainSubtotal <> 0 Then ' not on the first one but print the subtotals
									
									%>
										<tr style="font-size:16px;line-height:40px;" bgcolor="#FFFF99">
											<th scope="col" colspan="5" align="right" style="font-weight:bold;" valign="middle" height="40"><strong>Subtotal <%=FormatCurrency(ChainSubtotal)%></strong></th>
										</tr>
										<tr><th colspan="5" style="height:50px;">&nbsp;</th><tr>
										<% ChainSubtotal = 0
									End If%>
										<tr style="font-size:16px;line-height:40px;" bgcolor="#ADD8E6">
											<th scope="col" colspan="5" align="center" valign="middle" height="40"><strong>ACCOUNT #<%= rsInvoices("CustNum") %>&nbsp;&nbsp;<%= CustName %></strong></th>
										</tr>
										<tr style="font-size:16px;line-height:40px;" bgcolor="#ADD8E6">
											<th scope="col" colspan="5" align="center" valign="middle" height="40"><strong><%= Address1 %>&nbsp;&nbsp;<%= Address2 %>&nbsp;&nbsp;<%= City %></strong></th>
										</tr>
										
										
										<tr><th colspan="5">							
											<!-- the actual invoice starts here !-->
											<table width="850" cellpadding="0" cellspacing="0" bgcolor="#f5f5f5" border="1" bordercolor="#000" style="font-size: 16px;">
												<!-- titles !-->
												<tr>
													<th scope="col" width="850">
														<table width="850" cellpadding="10" cellspacing="0">
															<tbody>
																<tr>
																	<th scope="col" width="20%" valign="middle" align="center"><strong>SKU</strong></th>
																	<th scope="col" width="48%" valign="middle" align="center"><strong>DESCRIPTION</strong></th>
																	<th scope="col" width="15%" valign="middle" align="center"><strong>UNIT PRICE</strong></th>
																	<th scope="col" width="10%" valign="middle" align="center"><strong>QTY</strong></th>
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
													<tr height="40">
														<th scope="col" width="20%" align="left" style="font-weight:normal;"><%= rsDateRangeSKUCount("partnum") %></th>
														<th scope="col"width="50%" align="left" style="font-weight:normal;"><%= rsDateRangeSKUCount("Description") %></th>
														<th scope="col" width="15%" align="right" style="font-weight:normal;"><%= formatCurrency(rsDateRangeSKUCount("Price"),2) %></th>
														<th scope="col" width="10%" align="center" style="font-weight:normal;"><%= rsDateRangeSKUCount("ItemQTY") %></th>
														<th scope="col" width="10%" align ="right" style="font-weight:normal;"><%= formatCurrency(rsDateRangeSKUCount("Expr1"),2) %></th>
													</tr>			
												<%
											rsDateRangeSKUCount.MoveNext
										Loop
									End If

						End If 'HELD CUST

						ChainSubtotal = ChainSubtotal + rsInvoices("IvsTotalAmt") - rsInvoices("IvsSalesTax")
						TotalAmt = TotalAmt + rsInvoices("IvsTotalAmt")
						TotalTax = TotalTax + rsInvoices("IvsSalesTax")
						totalSubtotal = TotalAmt - TotalTax 

						rsInvoices.movenext
					Loop
				%>
					<tr style="font-size:16px;" bgcolor="#FFFF99">
						<th scope="col" colspan="5" align="right" style="font-weight:bold;" valign="middle" height="40"><strong>Subtotal <%=FormatCurrency(ChainSubtotal)%></strong></th>
					</tr>
					<tr>
						<th colspan="5" height="40">&nbsp;</th>
					<tr>
					<%
				

				End If
				Set rsInvoices = Nothing
				cnnInvoices.Close
				Set cnnInvoices = Nothing
				%>
		</table>
		<!-- the table ends here !-->

			<!-- grand total !-->
			<table  width="850" cellpadding="10" cellspacing="0" border="3" bgcolor="#DDDDDD" style="margin-top:20px;font-size: 16px;">

				<tr>
					<th scope="col" width="80%" align="right"><strong>SUBTOTAL: </strong></th>
					<th scope="col" width="20%" align="right"><strong><%= FormatCurrency(totalSubtotal) %></strong></th>
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

			<table  width="850" cellpadding="10" cellspacing="0" border="3" bgcolor="#FFFF00" style="margin-top:20px;font-size: 16px;">

				<tr>
					<th scope="col" width="100%" align="center"><strong><%= UCASE(MessageToPrint) %></strong></th>
				</tr>
			</table>
			<!-- eof grand total !-->

			</td>						
		</tr>
	</tbody>
</table>
					 		
	</body>
	
</html>