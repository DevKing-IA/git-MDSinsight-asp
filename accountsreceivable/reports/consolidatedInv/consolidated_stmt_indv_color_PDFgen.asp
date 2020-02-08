<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/Insightfuncs.asp"-->

<%  

'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
sURL = Request.ServerVariables("SERVER_NAME")

IvsSeq = Request.QueryString("i")
Username = Request.QueryString("u")
ClientKey = Request.QueryString("cl")
UserNo = Request.QueryString("un")

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
	 	   font-size: 13px;
	  
 	   }
 	   
 	   h3{
	 	   margin-bottom: 5px;
 	   }
 	   
 	   .invoicenr{
	 	   margin-top: 5px;
	 	   color: red;
	 	   display: block;
	 	   width: 100%;
 	   }
 	   
 	   .accountnr{
	 	   margin-top: 5px;
	 	   color: green;
	 	   display: block;
	 	   width: 100%;
	 	   float: left;
 	   }
 	   
 	   .maintitle{
	 	   margin-bottom: 0px;
	 	   text-align: center;
 	   }
		
		
		.batch-invoice-titles th{
			font-size: 11px;
			font-weight: normal;
			text-transform: uppercase;
		}
		
		
			.the-form th{
			font-size: 11px;
			font-weight: normal;
			text-transform: uppercase;
		}
		
		.subtotal-table th:first-child{
			/* border: 0px; */
		}
		
		.subtotal-table th{
			border: 1px solid #111;
		}
		
		.the-form thead{
 			font-weight: bold;
			text-transform: uppercase;
 		}
		
		.deliver-to{
			width: 848px;
			padding: 10px 0px 10px 0px;
			border-left: 1px solid #222;
			border-right: 1px solid #222;
			text-align: center;
			font-size: 15px;
			font-weight: bold;
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
 
 	  	.tr-fonts{
	 	  	font-size: 13px;
 	  	}
			  
		/* eof first statement */
		
		</style>
		
	</head>
	
	<body>
		
		<!-- main table starts here !-->
 		<table width="650" align="center">
			<tbody >
				<tr>
					<td width="100%">
		
		<!-- logo / address / invoice starts here !-->
		<table width="850" style="margin-bottom:20px;" cellpadding="5" cellspacing="5">
			<tbody>
				<tr>
					
					<!-- logo !-->
					<th scope="col" align="left">
							<img src="../../../clientfiles/<%=MUV_Read("ClientID")%>/logos/logo.png">
					</th>
					<!-- eof logo !-->
					
					<!-- address !-->
					
						<%
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
					<th scope="col" align="left" style="color:<%= CompanyIdentityColor1 %> !important;">
					<%
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
					<!-- eof address !-->
					
					<!-- invoice !-->
					<th scope="col" align="center">
						<h3>INVOICE</h3>
						PLEASE REFER TO<br> OUR INVOICE<br> NUMBER WHEN REMITTING<br>
						<strong class="invoicenr">INVOICE NO.: <%=GetInvoiceNumberByIvsSeq(IvsSeq) %></strong>
						
					</th>
					<!-- eof invoice !-->
					
  				</tr>
			</tbody>
		</table>
		<!-- logo / address / invoice ends here !-->
		
		
		<!-- history reprint !-->
		<table width="850" style="margin-bottom:20px;" cellpadding="5" cellspacing="5" bgcolor="#f5f5f5">
			<tbody>
				<tr>
					
					<!-- title !-->
					<th scope="col" align="left">
							<h2 class="maintitle">HISTORY REPRINT</h2>
					</th>
					<!-- eof title !-->
 
		 		
		</tr>
			</tbody>
		</table>
		<!-- eof history reprint !-->
		
		
		<!-- sold / ship to -->
		<table width="850" style="margin-bottom:20px;" cellpadding="5" cellspacing="5" >
			<tbody>
				<tr>
					
					<!-- sold to !-->
					<th scope="col" align="left" width="50%">
						<h3>SOLD TO</h3>
							<strong class="accountnr">ACCOUNT NO.: <%=GetCustNumberByInvSeq(IvsSeq)%></strong>
							
							<div style="width:425px;float:left;">
							<%
							
							SQLBillTo = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".AR_CustomerBillto Where CustNum = '" & GetCustNumberByInvSeq(IvsSeq) &"'"
							Set cnnBillTo = Server.CreateObject("ADODB.Connection")
							cnnBillTo.open (Session("ClientCnnString"))
							Set rsBillTo = Server.CreateObject("ADODB.Recordset")
							rsBillTo.CursorLocation = 3 
							Set rsBillTo = cnnBillTo.Execute(SQLBillTo)
							If NOT rsBillTo.EOF Then 
							
								tmpCustName = rsBillTo("BillName")
								tmpAddr1 = rsBillTo("Addr1")
								tmpAddr2 = rsBillTo("Addr2")
								tmpCSZ = rsBillTo("City") & ", " & rsBillTo("State") & "&nbsp;&nbsp;" & rsBillTo("Zip")

								Response.Write(tmpCustName & "<br>")
								Response.Write(tmpAddr1 & "<br>")
								If tmpAddr2  <> "" Then
									Response.Write(tmpAddr2 & "<br>")
								End If
								Response.Write(tmpCSZ & "<br>") 

							Else
							
								SQLBillToSecondary = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".AR_Customer Where CustNum = '" & GetCustNumberByInvSeq(IvsSeq) &"'"
								Set cnnBillToSecondary = Server.CreateObject("ADODB.Connection")
								cnnBillToSecondary.open (Session("ClientCnnString"))
								Set rsBillToSecondary = Server.CreateObject("ADODB.Recordset")
								rsBillToSecondary.CursorLocation = 3 
								Set rsBillToSecondary = cnnBillToSecondary.Execute(SQLBillToSecondary)
								
								If NOT rsBillToSecondary.EOF Then 
								
									tmpCustName = rsBillToSecondary("Name")
									tmpAddr1 = rsBillToSecondary("Addr1")
									tmpAddr2 = rsBillToSecondary("Addr2")
									tmpCSZ = rsBillToSecondary("CityStateZip")
	
									Response.Write(tmpCustName & "<br>")
									Response.Write(tmpAddr1 & "<br>")
									If tmpAddr2  <> "" Then
										Response.Write(tmpAddr2 & "<br>")
									End If
									Response.Write(tmpCSZ & "<br>") 
								End If
								
								Set rsBillToSecondary = Nothing
								cnnBillToSecondary.Close
								Set cnnBillToSecondary = Nothing


							End If
							Set rsBillTo = Nothing
							cnnBillTo.Close
							Set cnnBillTo = Nothing
							
							%>
							</div>
					</th>
					<!-- eof sold to !-->
					
					<!-- ship to !-->
					<th scope="col" align="left" width="50%">
					<h3>SHIP TO</h3>
							<%
							
								SQLBillToSecondary = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".AR_Customer Where CustNum = '" & GetCustNumberByInvSeq(IvsSeq) &"'"
								Set cnnBillToSecondary = Server.CreateObject("ADODB.Connection")
								cnnBillToSecondary.open (Session("ClientCnnString"))
								Set rsBillToSecondary = Server.CreateObject("ADODB.Recordset")
								rsBillToSecondary.CursorLocation = 3 
								Set rsBillToSecondary = cnnBillToSecondary.Execute(SQLBillToSecondary)
								
								If NOT rsBillToSecondary.EOF Then 
									tmpCustName = rsBillToSecondary("Name")
									tmpAddr1 = rsBillToSecondary("Addr1")
									tmpAddr2 = rsBillToSecondary("Addr2")
									tmpCSZ = rsBillToSecondary("CityStateZip")
									Response.Write(tmpCustName & "<br>")
									Response.Write(tmpAddr1 & "<br>")
									If tmpAddr2  <> "" Then
										Response.Write(tmpAddr2 & "<br>")
									End If
									Response.Write(tmpCSZ & "<br>") 
								End If	
															
								Set rsBillToSecondary = Nothing
								cnnBillToSecondary.Close
								Set cnnBillToSecondary = Nothing
						%>
							

					</th>
					<!-- eof ship to !-->
 
		 		
		</tr>
			</tbody>
		</table>
		<!-- eof sold / ship to !-->
		
		<!-- batch invoice titles -->
		<table width="850"  cellpadding="5" cellspacing="5" border="1" bordercolor="#111111" style="margin-bottom:-1px;" class="batch-invoice-titles">
			<tbody>
				<tr>
					
					<!-- order no. !-->
					<th scope="col" align="center" valign="top">
						Your Order No.<br>
						<strong>
						<%If GetPONumberByInvSeq(IvsSeq) <> "" Then
							Response.Write(GetPONumberByInvSeq(IvsSeq))
						Else
							Response.Write("&nbsp;")
						End If%>
						</strong>
					</th>
					<!-- eof order no. !-->
					
					<!-- total cases !-->
					<th scope="col" align="center" valign="top">
						Total Cases
					</th>
					<!-- eof total cases !-->
					
					<!-- packer !-->
					<th scope="col" align="center" valign="top">
						Packer
					</th>
					<!-- eof packer !-->
					
					<!-- route !-->
					<th scope="col" align="center" valign="top">
						Route<br>
						<strong>
						<%If GetRouteNumByInvSeq(IvsSeq) <> "" Then
							Response.Write(GetRouteNameByRouteNum(GetRouteNumByInvSeq(IvsSeq)))
						Else
							Response.Write("&nbsp;")
						End If%>
						</strong>
					</th>
					<!-- eof route !-->
					
					<!-- batch invoice !-->
					<th scope="col" align="center" valign="middle">
						Terms<br>
						<strong>
						<%If GetTermsNumByInvSeq(IvsSeq) <> "" Then
							Response.Write(GetTermsNameByTermsNum(GetTermsNumByInvSeq(IvsSeq)))
						Else
							Response.Write("&nbsp;")
						End If%>
						</strong>
					</th>
					<!-- eof batch invoice !-->
					
						<!-- cust. rep.  !-->
					<th scope="col" align="center" valign="middle" >
						Cust. Rep.<br>
						<strong>
						<%If GetPrimarySalesmanByInvSeq(IvsSeq) <> "" Then
							Response.Write(GetSalesmanNameBySalesmanNum(GetPrimarySalesmanByInvSeq(IvsSeq)))
						Else
							Response.Write("&nbsp;")
						End If%>
						</strong>

					</th>
					<!-- eof cust. rep. !-->
					
					<!-- date  !-->
					<th scope="col" align="center" valign="middle" width="107">
					Date<br>
					<strong>
					<%If GetInvoiceDateByInvSeq(IvsSeq) <> "" Then
						Response.Write(FormatDateTime(GetInvoiceDateByInvSeq(IvsSeq)))
					Else
						Response.Write("&nbsp;")
					End If%>
					</strong>
					</th>
					<!-- eof date !-->
					  
		 		
		</tr>
			</tbody>
		</table>
		<!-- eof batch invoice titles !-->
		
		
		
		<!-- the form starts here -->
		<table width="850"  cellpadding="5" cellspacing="5" border="1" bordercolor="#111111" style="margin-bottom:-1px;" class="the-form">
			
			<!-- form titles !-->
			<thead>
				<tr bgcolor="#<%= CompanyIdentityColor1 %>" style="color:#fff;">
					
				<!-- code !-->
					<th scope="col" align="center" width="121">Code</th>
				<!-- eof code !-->
				
				<!-- quantity !-->
				<th scope="col" align="center" width="79">U/M</th>
				<!-- eof quantity !-->

				<!-- quantity !-->
				<th scope="col" align="center" width="64">QTY</th>
				<!-- eof quantity !-->
				
				<!-- indicator !-->
				<th scope="col" align="center" width="15">#*</th>
				<!-- eof indicator !-->

				<!-- description !-->
					<th scope="col" align="center" width="333">Description</th>
				<!-- eof description !-->
				
				<!-- price !-->
					<th scope="col" align="center" width="105">Price</th>
				<!-- eof price !-->
				
				<!-- amount !-->
					<th scope="col" align="center" width="133">Amount</th>
				<!-- eof amount !-->
				
				</tr>
			</thead>
			<!-- eof form titles !-->
			
			
			<!-- form entries !-->
			<%
			'Invoice line items start here%>
			<tbody>
			
			<% If GetSpecialCommentByCustNum(GetCustNumberByInvSeq(IvsSeq)) <> "" Then
				Response.Write("<div class='deliver-to'>")
				Response.Write(GetSpecialCommentByCustNum(GetCustNumberByInvSeq(IvsSeq)))
				Response.Write("</div>")
			End If %>

			
			
			<%			
			SQLInvDets = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail WHERE IvsHistSequence = '" & IvsSeq & "' AND partNum <> 'CYCLE' "
			SQLInvDets = SQLInvDets  & " Order By CASE PartNext WHEN 0 THEN 99999999 ELSE PartNext END"
			
			'Response.Write(SQLInvDets  & "<br>")
			
			Set cnnInvDets = Server.CreateObject("ADODB.Connection")
			cnnInvDets.open (Session("ClientCnnString"))
			Set rsInvDets = Server.CreateObject("ADODB.Recordset")
			rsInvDets.CursorLocation = 3 
			Set rsInvDets = cnnInvDets.Execute(SQLInvDets)
			If not rsInvDets.Eof Then

				Do While Not rsInvDets.Eof 
				
					Response.Write("<tr>")
					Response.Write("<th scope='col' align='left' width='121' style='font-weight:normal; font-size:13px;'>")
					Response.Write(rsInvDets("partNum"))
					Response.Write("</th>")
		
					Response.Write("<th scope='col' align='center' width='79' style='font-weight:normal; font-size:13px;'>")
					Select Case ucase(rsInvDets("prodSalesUnit"))
						Case "U"
							Response.Write("EACH")
						Case "N"
							Response.Write("EACH")							
						Case Else
							Response.Write("CASE")
					End Select
					Response.Write("</th>")
					
					Response.Write("<th scope='col' align='center' width='64' style='font-weight:normal; font-size:13px;'>")
					Response.Write(rsInvDets("itemQuantity"))
					Response.Write("</th>")
					
					Indicator = ""
					Select Case GetTaxableFlagByIvsHistDetSequence(rsInvDets("IvsHistDetSequence"))
						Case "Y" ' taxable
							Indicator = "*"
						Case "N" ' nothing
							Indicator = "&nbsp;"
						Case "B" 'both
							Indicator = "#*"
						Case "D" ' deposit
							Indicator = "#"
					End Select	
					Response.Write("<th scope='col' align='center' width='15' style='font-weight:normal; font-size:13px;'>")
					Response.Write(Indicator)
					Response.Write("</th>")

					
					Response.Write("<th scope='col' align='left' width='333' style='font-weight:normal; font-size:13px;'>")
					Response.Write(rsInvDets("prodDescription"))
					Response.Write("</th>")
					
					Response.Write("<th scope='col' align='right' width='105' style='font-weight:normal; font-size:13px;'>")
					Response.Write(FormatCurrency(rsInvDets("itemPrice"),2,,0))
					Response.Write("</th>")

					Response.Write("<th scope='col' align='right' width='133' style='font-weight:normal; font-size:13px;'>")
					Response.Write(FormatCurrency(rsInvDets("itemPrice")*rsInvDets("itemQuantity"),2,,0))
					Response.Write("</th>")
			
					Response.Write("</tr>")
					Response.Write("</tbody>")
				
					rsInvDets.Movenext
				Loop
			End IF
			rsInvDets.close
			Set rsInvDets = Nothing
			cnnInvDets.close
			Set cnnInvDets = Nothing
			%>

			<!-- eof form entries !-->
			
		</table>
		<!-- the form ends here !-->
		
		
		<%
		'*******************************************************
		'*******************************************************
		'*******************************************************
		'Now do all the subtotaling at the bottom of the invoice
		'*******************************************************
		'*******************************************************		
		'*******************************************************
		%>
		
	 
					
		<!-- the total-->
		<table width="237"  cellpadding="5" cellspacing="5" border="1" bordercolor="#111111"  class="the-form" align="right">
			
			
 				
				<!-- subtotal line !-->
				<tr>
					
					 
				
				 				
				<!-- subtotal !-->
					<th scope="col" align="right" width="44.9%"><strong>SUBTOTAL</strong></th>
				<!-- eof subtotal !-->
				
				<!-- amount !-->
					<th scope="col" align="right"   ><strong><% Response.Write(FormatCurrency(GetInvoiceSubTotsByInvSeq(IvsSeq,"MERCH"),2,,0)) %></strong></th>
				<!-- eof amount !-->
				
  </tr>
 <!-- eof subtotal line !-->
 
 
 <!-- subtotal line !-->
				<tr>
					
					 			
				<!-- subtotal !-->
					<th scope="col" align="right" width="44.9%"><strong>Recycling Charge</strong></th>
				<!-- eof subtotal !-->
				
				<!-- amount !-->
					<th scope="col" align="right"   ><strong><% Response.Write(FormatCurrency(GetInvoiceSubTotsByInvSeq(IvsSeq,"RECYCLE"),2,,0)) %></strong></th>
				<!-- eof amount !-->
				
  </tr>
 <!-- eof subtotal line !-->
 
 
 <!-- subtotal line !-->
				<tr>
					
					 				
				<!-- subtotal !-->
					<th scope="col" align="right" width="44.9%"><strong>* Sales Tax</strong></th>
				<!-- eof subtotal !-->
				
				<!-- amount !-->
					<th scope="col" align="right"   ><strong><% Response.Write(FormatCurrency(GetInvoiceSubTotsByInvSeq(IvsSeq,"TAX"),2,,0)) %></strong></th>
				<!-- eof amount !-->
				
  </tr>
 <!-- eof subtotal line !-->
 
 <!-- subtotal line !-->
				<tr>
					
					 
				
		 
				
				<!-- subtotal !-->
					<th scope="col" align="right" width="44.9%"><strong># Deposit</strong></th>
				<!-- eof subtotal !-->
				
				<!-- amount !-->
					<th scope="col" align="right"   ><strong><% Response.Write(FormatCurrency(GetInvoiceSubTotsByInvSeq(IvsSeq,"DEPOSIT"),2,,0)) %></strong></th>
				<!-- eof amount !-->
				
  </tr>
 <!-- eof subtotal line !-->
 
 <!-- subtotal line !-->
				<tr>
					
					 
								
				<!-- subtotal !-->
					<th scope="col" align="right" width="44.9%"><strong>Total Due</strong></th>
				<!-- eof subtotal !-->
				
				<!-- amount !-->
					<th scope="col" align="right"  ><strong><% Response.Write(FormatCurrency(GetInvoiceSubTotsByInvSeq(IvsSeq,"GRAND"),2,,0)) %></strong></th>
				<!-- eof amount !-->
				
  </tr>
 <!-- eof subtotal line !-->
 
 			
		</table>
				 
		 
		
		</td>
		</tr>
		</tbody>
		</table>
		<!-- main table ends here !-->
		
	</body>
	
</html>