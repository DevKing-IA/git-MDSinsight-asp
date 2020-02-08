<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/Insightfuncs.asp"-->
<%
dummy = MUV_Write("ClientID","") ' Need this here

'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
sURL = Request.ServerVariables("SERVER_NAME")

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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Generate MCS Pending Charges PDF<%
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
		<title>MCS Pending Charges</title>
		
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
 		<table width="1150" align="center">
			<tbody >
				<tr>
					<td width="100%">
		
		<!-- logo / address / account starts here !-->
		<table width="1150" style="margin-bottom:20px;">
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
		
		<table width="1150" border="1" bordercolor="#111111" cellpadding="4" style="margin-bottom:-1px;">
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
						%>
					</th>
					<!-- eof address !-->
					
					<!-- sold to !-->
					<th scope="col" style="font-size:12px; font-weight:normal;" align="right">
						<!-- title !-->
						<%
						
							If Phone1 <> "" Then Response.Write(Phone1 & "<br>")															
							If Phone2 <> "" Then Response.Write(Phone2 & "<br>")															
							If Phone3 <> "" Then Response.Write(Phone3 & "<br>")															
							If Fax <> "" Then Response.Write("Fax:" & Fax & "<br>")																						
							If Email <> "" Then Response.Write(Email & "<br>")	
						
						%>
						<!-- eof title !-->
					</th>
					<!-- eof sold to !-->
					
					</tr>
			</tbody>
		</table>
		<!-- logo / address / account ends here !-->
		
		<!-- monthly consolidated invoice title !-->
		<table width="1150" border="1" bordercolor="#111111"  cellpadding="4" style="margin-bottom:-1px;">
			<tbody>
				<tr>
					<%
			
						'***************************************************************
						'Get Month/Year of Pending Charges
						'***************************************************************

						Set cnnPendingDates = Server.CreateObject("ADODB.Connection")
						cnnPendingDates.open (Session("ClientCnnString"))
						
						Set rsPendingDates = Server.CreateObject("ADODB.Recordset")
						rsPendingDates.CursorLocation = 3
						
						
						SQLPendingDates = "SELECT TOP (1) MCSMonth AS PendingChargesMonth "
						SQLPendingDates = SQLPendingDates & " FROM  BI_MCSActions "
						SQLPendingDates = SQLPendingDates & " WHERE (Action LIKE '%invoice%') "
						SQLPendingDates = SQLPendingDates & " ORDER BY InternalRecordIdentifier DESC "

						Set rsPendingDates = cnnPendingDates.Execute(SQLPendingDates)
					
						If Not rsPendingDates.EOF Then
							PendingChargesMonth = rsPendingDates("PendingChargesMonth")
						End If


						SQLPendingDates = "SELECT TOP (1) YEAR(RecordCreationDateTime) AS Expr1 "
						SQLPendingDates = SQLPendingDates & " FROM  BI_MCSActions "
						SQLPendingDates = SQLPendingDates & " WHERE (Action LIKE '%invoice%') "
						SQLPendingDates = SQLPendingDates & " ORDER BY InternalRecordIdentifier DESC "

						Set rsPendingDates = cnnPendingDates.Execute(SQLPendingDates)
					
						If Not rsPendingDates.EOF Then
							PendingChargesYear = rsPendingDates("Expr1")
							
							If PendingChargesMonth = "December" Then
								PendingChargesYear = PendingChargesYear - 1
							End If
						End If	
					
						Set rsPendingDates = Nothing
						cnnPendingDates.Close
						Set cnnPendingDates = Nothing
						
				
					%>			
					<!-- address !-->
					<th scope="col" >
						<h3 style="line-height:1; margin-top:10px; margin-bottom:10px;" align="center">MCS Pending Charges for <%= PendingChargesMonth %>&nbsp;<%= PendingChargesYear %></h3>
					</th>
				</tr>
			</tbody>
		</table>
		<!-- eof monthly consolidated invoice title !-->
		
		<!-- the table with statements starts here !-->
		<table width="1150" border="1" bordercolor="#111111" cellpadding="4" style="margin-bottom:-1px;">
			<thead>
				<tr bgcolor="<%= CompanyIdentityColor1 %>" style="color:#fff;">
					<th scope="col" align="center">Account</th>
					<th scope="col">Client</th>
					<th scope="col">Primary Salesman</th>
					<th scope="col">Action Notes</th>
					<th scope="col">Pending Charges</th>
				</tr>
			</thead>
			
			<tbody>
		
	
				<%
				 'Now get the actual invoice data
					'Only need this simple query here because the previous page built the include file for us
					
					'*************************************
					'ACCOUNTS WITH PENDING CHARGES
					'*************************************
			
					SQLPendingCharges = "SELECT BI_MCSActions.CustID, BI_MCSActions.ActionNotes, AR_Customer.Salesman "
					SQLPendingCharges = SQLPendingCharges & " FROM  BI_MCSActions INNER JOIN "
					SQLPendingCharges = SQLPendingCharges & "  AR_Customer ON BI_MCSActions.CustID = AR_Customer.CustNum "
					SQLPendingCharges = SQLPendingCharges & " WHERE      (BI_MCSActions.MCSMonth = "
					SQLPendingCharges = SQLPendingCharges & " (SELECT      TOP (1) MCSMonth "
					SQLPendingCharges = SQLPendingCharges & "  FROM         BI_MCSActions AS BI_MCSActions_1 "
					SQLPendingCharges = SQLPendingCharges & "  ORDER BY InternalRecordIdentifier DESC)) AND (BI_MCSActions.Action LIKE '%invoice%') AND (YEAR(BI_MCSActions.RecordCreationDateTime) = "
					SQLPendingCharges = SQLPendingCharges & " (SELECT      TOP (1) YEAR(RecordCreationDateTime) AS Expr1 "
					SQLPendingCharges = SQLPendingCharges & "  FROM         BI_MCSActions AS BI_MCSActions_1 "
					SQLPendingCharges = SQLPendingCharges & "  WHERE      (Action LIKE '%invoice%') "
					SQLPendingCharges = SQLPendingCharges & "  ORDER BY InternalRecordIdentifier DESC)) "
					SQLPendingCharges = SQLPendingCharges & " ORDER BY AR_Customer.Salesman "
				
					'Response.Write(SQLPendingCharges & "<br>")
	
					Set cnnPendingCharges = Server.CreateObject("ADODB.Connection")
					cnnPendingCharges.open (Session("ClientCnnString"))
					Set rsPendingCharges = Server.CreateObject("ADODB.Recordset")
					rsPendingCharges.CursorLocation = 3 
					
					Set rsPendingCharges = cnnPendingCharges.Execute(SQLPendingCharges)

					If not rsPendingCharges.Eof Then
					
						TotalPendingCharges = 0

						Do While not rsPendingCharges.Eof
						
							PrimarySalesMan =  ""
							SelectedCustomerID = rsPendingCharges("CustID")
							CustName = GetCustNameByCustNum(rsPendingCharges("CustID"))
							
							PrimarySalesMan = rsPendingCharges("Salesman")
							
							If PrimarySalesMan <> "" Then
								PrimarySalesPerson = GetSalesmanNameBySlsmnSequence(PrimarySalesMan)
							Else
								PrimarySalesPerson = ""
							End If
							
							ActionNotes = rsPendingCharges("ActionNotes")
		
			
							'***************************************************************
							'Get Current Pending Charges For This Account
							'***************************************************************
							
							SQLPendingLVF = "SELECT * FROM AR_Customer INNER JOIN BI_MCSData ON BI_MCSData.CustID = AR_Customer.CustNum WHERE AR_Customer.CustNum = " & SelectedCustomerID 
						
							Set cnnPendingLVF = Server.CreateObject("ADODB.Connection")
							cnnPendingLVF.open (Session("ClientCnnString"))
							
							Set rsPendingLVF = Server.CreateObject("ADODB.Recordset")
							rsPendingLVF.CursorLocation = 3
							Set rsPendingLVF = cnnPendingLVF.Execute(SQLPendingLVF)
						
							If Not rsPendingLVF.EOF Then
								PendingLVFHolder = rsPendingLVF("PendingLVF")
							Else
								PendingLVFHolder = 0
							End If
							
							TotalPendingCharges = TotalPendingCharges + PendingLVFHolder
							
							PendingLVFHolder = FormatCurrency(PendingLVFHolder,2)
							
							Set rsPendingLVF = Nothing
							cnnPendingLVF.Close
							Set cnnPendingLVF = Nothing
							
							'***************************************************************

							%>
						
							<!-- line !-->
							<tr>
								
								<th scope="col" style="font-size:12px; font-weight:normal;" align="center">
									<%= SelectedCustomerID %>
								</th>
								
								<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
									<%= CustName %>
								</th>
								
								<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
								<%
								
							    If Instr(PrimarySalesPerson ," ") <> 0 Then
									Response.Write(Left(PrimarySalesPerson,Instr(PrimarySalesPerson ," ")))
								Else
									Response.Write(PrimarySalesPerson)
									
								End If
								%>
								</th>

								<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
									<%= ActionNotes %>
								</th>
								
								<th scope="col" style="font-size:12px; font-weight:normal;" align="right">
									<%= PendingLVFHolder %>
								</th>
								
							</tr>
							<%
							
							rsPendingCharges.movenext
						Loop
					End If
					Set rsPendingCharges = Nothing
					cnnPendingCharges.Close
					Set cnnPendingCharges = Nothing
					%>
				<!-- eof line !-->
			</tbody>
		</table>
		<!-- the table with statements ends here !-->


		<!-- total !-->
		<table width="1150" border="1" bordercolor="#111111" cellpadding="4" style="margin-bottom:20px;">
			<tbody>
				<tr>
					<th scope="col" align="right">
						<h5 style="line-height:1;  margin-top:10px; margin-bottom:10px;">TOTAL PENDING CHARGES TO BE COLLECTED&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=FormatCurrency(TotalPendingCharges,2) %></h5>
					</th>
					
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
