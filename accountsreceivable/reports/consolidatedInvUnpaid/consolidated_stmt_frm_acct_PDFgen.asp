<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/Insightfuncs.asp"-->
<%
dummy = MUV_Write("ClientID","") ' Need this here

'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
sURL = Request.ServerVariables("SERVER_NAME")

StartDate = Request.QueryString("s")
EndDate = Request.QueryString("e")
Account = Request.QueryString("c")
StartDate = Replace(StartDate, "~","/")
EndDate = Replace(EndDate, "~","/")
Username = Request.QueryString("u")
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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Consolidated Invoice From Account report<%
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
						 <strong class="sold-to-strong">SOLD TO: <%=GetTerm("Account")%> # <%=Account%></strong><br>
						<!-- eof title !-->
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
					<!-- eof sold to !-->
					
					</tr>
			</tbody>
		</table>
		<!-- logo / address / account ends here !-->
		
		<!-- monthly consolidated invoice title !-->
		<table width="850" border="1" bordercolor="#111111"  cellpadding="4" style="margin-bottom:-1px;">
			<tbody>
				<tr>
					
					<!-- address !-->
					<th scope="col" >
						
						<h3 style="line-height:1; margin-top:10px; margin-bottom:10px;" align="center">Consolidated Invoice#
						<%Response.Write(Trim(Account) & Trim(Replace(EndDate,"/","")))%>
						<br><br>
						<small><%=StartDate%> - <%=EndDate%></small>
						</h3>

						
					</th>
				</tr>
			</tbody>
		</table>
		<!-- eof monthly consolidated invoice title !-->
		
		<!-- the table with statements starts here !-->
		<table width="850" border="1" bordercolor="#111111" cellpadding="4" style="margin-bottom:-1px;">
			<thead>
				<tr bgcolor="<%= CompanyIdentityColor1 %>" style="color:#fff;">
				<th scope="col">Invoice Date</th>
				<th scope="col">Invoice #</th>
				<th scope="col">Amount</th>
				</tr>
			</thead>
			
			<tbody>
				
				
							<% 'Now get the actual invoice data
								'Only need this simple query here because the previous page built the include file for us
								
								'*************************************
								'ACCOUNT
								'*************************************
						
								SQLInvoices = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistory WHERE "								
								SQLInvoices = SQLInvoices & " IvsTotalAmt - PaidAmt > 0 AND IvsHistSequence IN (Select IvsHistSequence from zReportConsolidatedInvoiceInclude_" & Trim(UserNo) & ") "
								SQLInvoices = SQLInvoices & " ORDER BY IvsNum"
															

								'Response.Write(SQLInvoices & "<br>")
				
								Set cnnInvoices = Server.CreateObject("ADODB.Connection")
								cnnInvoices.open (Session("ClientCnnString"))
								Set rsInvoices = Server.CreateObject("ADODB.Recordset")
								rsInvoices.CursorLocation = 3 
								
								Set rsInvoices = cnnInvoices.Execute(SQLInvoices)

								If not rsInvoices.Eof Then
									TotalAmt = 0
									Do While not rsInvoices.Eof

									%>
									
										<!-- line !-->
										<tr>
											
											<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
												<%=Month(rsInvoices("IvsDate")) & "/" & Day(rsInvoices("IvsDate")) & "/" & Year(rsInvoices("IvsDate"))%>
											</th>
											
											<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
												<%=rsInvoices("IvsNum")%>
											</th>
											
											<th scope="col" style="font-size:12px; font-weight:normal;" align="right">
												<%=FormatCurrency(rsInvoices("IvsTotalAmt"))%>
											</th>
											<% TotalAmt = TotalAmt + rsInvoices("IvsTotalAmt")%>
										</tr>
										<%rsInvoices.movenext
									Loop
								End If
								Set rsInvoices = Nothing
								cnnInvoices.Close
								Set cnnInvoices = Nothing
							%>
				<!-- eof line !-->
			</tbody>
		</table>
		<!-- the table with statements ends here !-->
		
		<!-- total !-->
		<table width="850" border="1" bordercolor="#111111" cellpadding="4" style="margin-bottom:20px;">
			<tbody>
				<tr>
	
					<% If DoNotShowDueDate <> "CHECKED" Then %>
						<% If DueDateSingleDate <> "" Then %>
							<th scope="col" align="right">
								<h5 style="line-height:1;  margin-top:10px; margin-bottom:10px;"><%=MessageToPrint & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"%> TOTAL: <%=FormatCurrency(TotalAmt) %>
								<br><br>INVOICE DUE DATE:  <%= FormatDateTime(DueDateSingleDate,2) %></h5>
							</th>
						<% Else %>
							<th scope="col" align="right">
								<h5 style="line-height:1;  margin-top:10px; margin-bottom:10px;"><%=MessageToPrint & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"%> TOTAL: <%=FormatCurrency(TotalAmt) %>
								<br><br>INVOICE DUE DATE:  <%= DateAdd("d",DueDateDays,EndDate) %></h5>
							</th>
						<% End If %>
					<% Else %>
						<th scope="col" align="right">
							<h5 style="line-height:1;  margin-top:10px; margin-bottom:10px;"><%=MessageToPrint & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"%> TOTAL: <%=FormatCurrency(TotalAmt) %></h5>
						</th>
					<% End If %>	
					
				</tr>
			</tbody>
		</table>
		<!-- eof total !-->
		
		
		 
		
		</td>
		</tr>
		</tbody>
		</table>
		<!-- main table ends here !-->
		
	</body>
	
</html>
