<% @ Language = VBScript %>
<%
Response.Buffer = True
Response.Expires = 0
Response.Clear
%>
<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_API.asp"-->

<%
Dim PageNo, LineCount


dummy=MUV_Write("ClientID","") 'Need this here

ClientKey = Request.QueryString("c")

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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - API Daily Activity Summary By Partner Report<%
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



'This is here so we only open it once for the whole page
Set cnn_Settings_Global = Server.CreateObject("ADODB.Connection")
cnn_Settings_Global.open (Session("ClientCnnString"))
Set rs_Settings_Global = Server.CreateObject("ADODB.Recordset")
rs_Settings_Global.CursorLocation = 3 
SQL_Settings_Global = "SELECT * FROM Settings_Global"
Set rs_Settings_Global = cnn_Settings_Global.Execute(SQL_Settings_Global)
If not rs_Settings_Global.EOF Then
	APIDailyActivityReportOnOff = rs_Settings_Global("APIDailyActivityReportOnOff")
	APIDailyActivityReportUserNos = rs_Settings_Global("APIDailyActivityReportUserNos")
	APIDailyActivityReportAdditionalEmails = rs_Settings_Global("APIDailyActivityReportAdditionalEmails")
	APIDailyActivityReportEmailSubject = rs_Settings_Global("APIDailyActivityReportEmailSubject")
	Order_OffSetFromToday = rs_Settings_Global("OrderAPIOffsetDays")
	Invoice_OffSetFromToday = rs_Settings_Global("InvoiceAPIOffsetDays")
	RA_OffSetFromToday = rs_Settings_Global("RAAPIOffsetDays")
	CM_OffSetFromToday = rs_Settings_Global("CMAPIOffsetDays")
	SumInv_OffSetFromToday = rs_Settings_Global("SumInvAPIOffsetDays")
	OrderCutoffTime = rs_Settings_Global("OrderCutoffTime")
	If Len(OrderCutoffTime) = 3 Then
		OrderCutoffTime = Left(OrderCutoffTime,1) & ":" & Right(OrderCutoffTime,2)
	Else
		OrderCutoffTime = Left(OrderCutoffTime,2) & ":" & Right(OrderCutoffTime,2)
	End IF
	InvoiceCutoffTime = rs_Settings_Global("InvoiceCutoffTime")
	If Len(InvoiceCutoffTime) = 3 Then
		InvoiceCutoffTime = Left(InvoiceCutoffTime,1) & ":" & Right(InvoiceCutoffTime,2)
	Else
		InvoiceCutoffTime = Left(InvoiceCutoffTime,2) & ":" & Right(InvoiceCutoffTime,2)
	End IF

Else
	APIDailyActivityReportOnOff = vbFalse
End If
Set rs_Settings_Global = Nothing
cnn_Settings_Global.Close
Set cnn_Settings_Global = Nothing

Order_currentDay = day(date()) - Order_OffSetFromToday 
Order_currentMonth = month(date())
Order_currentYear = year(date())

Invoice_currentDay = day(date()) - Invoice_OffSetFromToday
Invoice_currentMonth = month(date())
Invoice_currentYear = year(date())

RA_currentDay = day(date()) - RA_OffSetFromToday
RA_currentMonth = month(date())
RA_currentYear = year(date())

CM_currentDay = day(date()) - CM_OffSetFromToday
CM_currentMonth = month(date())
CM_currentYear = year(date())

SumInv_currentDay = day(date()) - SumInv_OffSetFromToday
SumInv_currentMonth = month(date())
SumInv_currentYear = year(date())

%>
<!DOCTYPE html>
<!--[if lt IE 7 ]> <html class="no-js ie6 oldie" lang="en"> <![endif]-->
<!--[if IE 7 ]>    <html class="no-js ie7 oldie" lang="en"> <![endif]-->
<!--[if IE 8 ]>    <html class="no-js ie8 oldie" lang="en"> <![endif]-->
<!--[if IE 9 ]>    <html class="no-js ie9" lang="en"> <![endif]-->
<!--[if (gte IE 9)|!(IE)]><![endif]--><!-->
<html class="no-js" lang="en">
<!--<![endif]-->
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta name="description" content="">
    <meta name="author" content="">

    <title>API Daily Activity Summary By Partner Report</title>

<%
    
Response.Write("<script src='https://use.fontawesome.com/3382135cdc.js'></script>")


Response.Write("<style type='text/css'>")
Response.Write("mark {")
Response.Write("    background-color: yellow;")
Response.Write("    color: black;")
Response.Write("}")
Response.Write("</style>")

Response.Write("<style type='text/css'>")
	
Response.Write("	body{font-family: arial, helvetica, sans-serif;}")
	
Response.Write("	div.table-title {")
Response.Write("	  display: block;")
Response.Write("	  margin: auto;")
Response.Write("	  max-width: 800px;")
Response.Write("	  padding:5px;")
Response.Write("	  width: 100%;")
Response.Write("	}")

	
Response.Write("	div.table-data {")
Response.Write("	  display: block;")
Response.Write("	  margin: auto;")
Response.Write("	  max-width: 1200px;")
Response.Write("	  padding:5px;")
Response.Write("	  width: 100%;")
Response.Write("	}")
	
Response.Write("	p, h1, h2 {")
Response.Write("	  display: block;")
Response.Write("	  margin: auto;")
Response.Write("	  max-width: 800px;")
Response.Write("	  padding:5px;")
Response.Write("	  width: 100%;")
Response.Write("	}")

Response.Write("	h1 {")
Response.Write("		color: #193048;")
Response.Write("	    font-size: 30px;")
Response.Write("	    font-weight: 400;")
Response.Write("	    font-style: normal;")
Response.Write("	    font-family: arial, helvetica, sans-serif;")
Response.Write("	    text-transform: uppercase;")
Response.Write("	    text-align:center;")
Response.Write("	}")
	
Response.Write("	.generated {")
Response.Write("		color: #3e94ec;")
Response.Write("	    font-size: 20px;")
Response.Write("	    font-weight: 400;")
Response.Write("	    font-style: normal;")
Response.Write("	    font-family: arial, helvetica, sans-serif;")
Response.Write("	    text-transform: uppercase;")
Response.Write("	    text-align:center;")
Response.Write("	}")
	
Response.Write("	h2 {")
Response.Write("		color: #3e94ec;")
Response.Write("	    font-size: 20px;")
Response.Write("	    font-weight: 400;")
Response.Write("	    font-style: normal;")
Response.Write("	    font-family: arial, helvetica, sans-serif;")
Response.Write("	    text-transform: uppercase;")
Response.Write("	    text-align:center;")
Response.Write("	}")
	
Response.Write("	hr{")
Response.Write("	   /* margin-top: 40px;")
Response.Write("	    margin-bottom: 40px;*/")
Response.Write("	}")
	
Response.Write("	.table-title h3 {")
Response.Write("	   color: #193048;")
Response.Write("	   font-size: 22px;")
Response.Write("	   font-weight: 400;")
Response.Write("	   font-style:normal;")
Response.Write("	   font-family: arial, helvetica, sans-serif;")
Response.Write("	   text-transform:uppercase;")
Response.Write("	   font-weight:bold;")
Response.Write("	}")
	
	
Response.Write("	/*** Table Styles **/")

Response.Write("	.table-fill {")
Response.Write("	  background: white;")
Response.Write("	  border-collapse: collapse;")
Response.Write("	  margin: auto;")
Response.Write("	  max-width: 800px;")
Response.Write(" 	  padding:5px;")
Response.Write("	  width: 100%;")
Response.Write("	  font-family: arial, helvetica, sans-serif;")
Response.Write("	}")
	 
Response.Write("	th {")
Response.Write("	   color:#483D8B;")
Response.Write("	  /*font-size:23px;*/")
Response.Write("	  font-size: 18px;")
Response.Write("	  font-weight: 100;")
Response.Write("	  padding:13px !important;")
Response.Write("	  text-align:left;")
Response.Write("	  vertical-align:middle;")
Response.Write("	  border: 1px solid #C1C3D1;")
Response.Write("	  /* width: 12.5% !important; */")
Response.Write("  	}")

Response.Write("	tr {")
Response.Write("	  color:#666B85;")
Response.Write("	  font-size:16px;")
Response.Write("	  font-weight:normal;")
Response.Write("	}")
	 	 
Response.Write("	tr:nth-child(odd) td {")
Response.Write("	  background:#EBEBEB;")
Response.Write("	}")
	 	
	 
Response.Write("	td {")
Response.Write("	  background:#FFFFFF;")
Response.Write("	  padding:9px 13px 8px 20px !important;")
Response.Write("	  text-align:left;")
Response.Write("	  vertical-align:middle;")
Response.Write("	  font-weight:300;")
Response.Write("	  font-size:18px;")
Response.Write("	  border: 1px solid #C1C3D1;")
Response.Write("	}")
	
Response.Write("	/* custom table */")
	
	 
	
Response.Write("	.custom-table th{")
Response.Write("		padding:5px;")
Response.Write("	}")
	
Response.Write("	.custom-table td{")
Response.Write("		padding:5px;")
Response.Write("	}")
	
Response.Write("	#leftcol{")
Response.Write("		width:65%;")
Response.Write("	}")
	
Response.Write("	#rightcol{")
Response.Write("		width:35%;")
Response.Write("	}")
	
Response.Write("	#table-fill-short{")
Response.Write("		max-width: 500px;")
Response.Write("	}")
Response.Write("	/* eof custom table */")
	
Response.Write("	.cust-logo{")
Response.Write("		position: absolute;")
Response.Write("		margin-left: -280px;")
Response.Write("	}")


Response.Write("	</style>")
     
Response.Write("</head>")



Response.Write("<body bgcolor='#FFFFFF' text='#000000' link='#000080' topmargin='0' leftmargin='0' rightmargin='0' bottommargin='0' marginwidth='0' marginheight='0'>")
	 

Response.Write("<div class='table-title'>")

PageNo = 0
Call PageHeader 

Response.Write("<br>")
Response.Write("</div>")

Response.Write("<div class='table-data'>")


SQLDailyAPIPartnersLoop = "SELECT DISTINCT(partnerAPIKey) FROM IC_PARTNERS	"	

Set cnnDailyAPIPartnersLoop = Server.CreateObject("ADODB.Connection")
cnnDailyAPIPartnersLoop.open(Session("ClientCnnString"))
Set rsDailyAPIPartnersLoop = Server.CreateObject("ADODB.Recordset")
rsDailyAPIPartnersLoop.CursorLocation = 3 
Set rsDailyAPIPartnersLoop = cnnDailyAPIPartnersLoop.Execute(SQLDailyAPIPartnersLoop)

If NOT rsDailyAPIPartnersLoop.EOF Then

	rowCount = 1

	Do While Not rsDailyAPIPartnersLoop.EOF
	
		currentPartnerAPIKey = rsDailyAPIPartnersLoop("partnerAPIKey")
	

			Response.Write("<hr>")
			Response.Write("<h2>Partner: " & GetPartnerNameByAPIKey(currentPartnerAPIKey) & "</h2>")
			Response.Write("<hr>")
				    	
		   	Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Order posts received " & FormatDateTime(dateadd("d",-1,Date())) & "&nbsp;" &  FormatDateTime(OrderCutoffTime,3) & " through ")
			If Time() > cDate(OrderCutoffTime) Then
			   	Response.Write(FormatDateTime(Date()) &  "&nbsp;" & FormatDateTime(OrderCutoffTime,3) & " </h4><br>")
			Else
			   	Response.Write(FormatDateTime(Date()) &  "&nbsp;" & FormatDateTime(Now(),3) & " </h4><br>")			
			End If
    
		    
			Response.Write("<table style='margin-left:20px;width:1300px;'>")	
		        Response.Write("<thead>")
		            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
			            Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='2%'>Count</th>")
   		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='28%'>Received</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='4%'>Order ID</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='13%'>Result</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='53%'>Message (if any)</th>")
		            Response.Write("</tr>")
		        Response.Write("</thead>")
		        Response.Write("<tbody>")

				StartRange = Year(dateadd("d",-1,Now())) & "-" & Month(dateadd("d",-1,Now())) & "-" &  Day(dateadd("d",-1,Now())) & " "
				StartRange = StartRange & OrderCutoffTime & ":00"
				EndRange = Year(Now()) & "-" & Month(Now()) & "-" & Day(Now()) & " "
				If Time() > cDate(OrderCutoffTime) Then
					EndRange = EndRange & OrderCutoffTime & ":00"
				Else
					EndRange = EndRange & Hour(Now()) & ":" & Minute(Now()) & ":00"
				End If

				SQLDailyAPIOrders = "SELECT * FROM API_OR_PostResults "
				SQLDailyAPIOrders = SQLDailyAPIOrders & " WHERE RecordCreationDateTime BETWEEN '" & StartRange & "' AND '" & EndRange & "' "
				SQLDailyAPIOrders = SQLDailyAPIOrders & " AND APIKey = '" & currentPartnerAPIKey & "' "
				SQLDailyAPIOrders = SQLDailyAPIOrders & " ORDER BY CASE PostResults WHEN 'notification!' THEN 1  WHEN 'success' THEN 2 WHEN 'failure!' THEN 3 ELSE 4 END, OrderID"
	'Response.Write(SQLDailyAPIOrders)

                Dim currentStatus: currentStatus = "notification!"
				
				Set cnnDailyAPIOrders = Server.CreateObject("ADODB.Connection")
				cnnDailyAPIOrders.open(Session("ClientCnnString"))
				Set rsDailyAPIOrders = Server.CreateObject("ADODB.Recordset")
				rsDailyAPIOrders.CursorLocation = 3 
				Set rsDailyAPIOrders = cnnDailyAPIOrders.Execute(SQLDailyAPIOrders)

				rowCount = 1
				DailyCount = 0
				
				If NOT rsDailyAPIOrders.EOF Then
					
					Do While Not rsDailyAPIOrders.EOF
						
						OrderID = "" : PostResults = "" : PostStatusMessage = ""
						
						DailyCount = DailyCount + 1
						
						Received = rsDailyAPIOrders("RecordCreationDateTime")
						OrderID = rsDailyAPIOrders("OrderID")
                        If isNull(OrderID) Then OrderID = "&nbsp;&nbsp;"
                        If OrderID = "" Then OrderID = "&nbsp;&nbsp;"
						PostResults = rsDailyAPIOrders("PostResults")
						PostStatusMessage = rsDailyAPIOrders("PostStatusMessage")
                        If isNull(PostStatusMessage) Then PostStatusMessage = "&nbsp;&nbsp;"
                        If PostStatusMessage = "" Then PostStatusMessage = "&nbsp;&nbsp;"

                        If PostResults<>currentStatus Then                            
                            If rowCount <> 1 Then
                                Response.Write("<tr><td style='padding-top: 8px; text-align: right;' align='right' colspan='5'>Total " & Replace(currentStatus,"!","") & " #:" & DailyCount-1 & "</td></tr>")
                                LineCount = LineCount + 1
                            End If 
                            DailyCount = 1
                            currentStatus = PostResults
                        End If
						
						Response.Write("<tr>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & DailyCount & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" &  Received & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & OrderID & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & PostResults & "</td>")
			             	Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & PostStatusMessage & "</td>")
			            Response.Write("</tr>")

						rowCount = rowCount + 1
						
						LineCount = LineCount + 1
						
						
						rsDailyAPIOrders.MoveNext
						
						If LineCount >= 20 Then 
					        Response.Write("</tbody>")
						    Response.Write("</table>")

							Call PageHeader

							Call SubHeader("Orders","Body")
							
							Response.Write("<table style='margin-left:20px;width:1300px;'>")	
						        Response.Write("<thead>")
						            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
							            Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' width='2%'>Count</th>")
				   		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='28%'>Received</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' width='4%'>Order ID</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='13%'>Result</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='53%'>Message (if any)</th>")
						            Response.Write("</tr>")
						        Response.Write("</thead>")
					        Response.Write("<tbody>")

						End if
						
					Loop
                    Response.Write("<tr><td style='padding-top: 8px; text-align: right;' align='right' colspan='5'>Total " & Replace(currentStatus,"!","") & " #:" & DailyCount & "</td></tr>")
                    LineCount = LineCount + 1
				Else
					Response.Write("<tr ><td colspan='8'>No Order API Data</td></tr>")
				End If

				
				Response.Write("<tr style='border-top: 2px solid #ddd;'>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'><strong>Count:&nbsp;&nbsp;" & rowCount-1 & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'>&nbsp;&nbsp;</td>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'>&nbsp;&nbsp;</td>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'>&nbsp;&nbsp;</td>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'>&nbsp;&nbsp;</td>")
	            Response.Write("</tr>")
		        Response.Write("</tbody>")
		    Response.Write("</table>")

		    
		    
		    '*********************************
			'		    I N V O I C E S 
		    '*********************************
		    If LineCount + 5 >= 20 Then
				Call PageHeader
				Call SubHeader("Invoices","Top")
			End If

		    LineCount = LineCount + 5
		    	    
		   	Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Invoice posts received " & FormatDateTime(dateadd("d",-1,Date())) & "&nbsp;" &  FormatDateTime(InvoiceCutoffTime,3) & " through ")
			If Time() > cDate(InvoiceCutoffTime) Then
			   	Response.Write(FormatDateTime(Date()) &  "&nbsp;" & FormatDateTime(InvoiceCutoffTime,3) & " </h4><br>")
			Else
			   	Response.Write(FormatDateTime(Date()) &  "&nbsp;" & FormatDateTime(Now(),3) & " </h4><br>")			
			End If

						
			
			Response.Write("<table style='margin-left:20px;width:1300px'>")	
		        Response.Write("<thead>")
		            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
			            Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' width='2%'>Count</th>")
   		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='28%'>Received</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' width='4%'>Invoice ID</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='13%'>Result</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='53%'>Message (if any)</th>")
		            Response.Write("</tr>")
		        Response.Write("</thead>")
		        Response.Write("<tbody>")
		        

                currentStatus = "notification!"
				SQLDailyAPIInvoices = "SELECT * FROM API_IN_PostResults "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " WHERE RecordCreationDateTime BETWEEN '" & StartRange & "' AND '" & EndRange & "' "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " AND APIKey = '" & currentPartnerAPIKey & "' "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " ORDER BY CASE PostResults WHEN 'notification!' THEN 1  WHEN 'success' THEN 2 WHEN 'failure!' THEN 3 ELSE 4 END, InvoiceID"

				Set cnnDailyAPIInvoices = Server.CreateObject("ADODB.Connection")
				cnnDailyAPIInvoices.open(Session("ClientCnnString"))
				Set rsDailyAPIInvoices = Server.CreateObject("ADODB.Recordset")
				rsDailyAPIInvoices.CursorLocation = 3 
				'Response.Write(SQLDailyAPIInvoices)
				Set rsDailyAPIInvoices = cnnDailyAPIInvoices.Execute(SQLDailyAPIInvoices)

				rowCount = 1
				DailyCount = 0
				
				If NOT rsDailyAPIInvoices.EOF Then

					Do While Not rsDailyAPIInvoices.EOF
						
						DailyCount = DailyCount + 1
						Received = rsDailyAPIInvoices("RecordCreationDateTime")
						InvoiceID = rsDailyAPIInvoices("InvoiceID")
						PostResults = rsDailyAPIInvoices("PostResults")
						PostStatusMessage = rsDailyAPIInvoices("PostStatusMessage")
                        If isNull(PostStatusMessage) Then PostStatusMessage = "&nbsp;&nbsp;"
                        If PostStatusMessage = "" Then PostStatusMessage = "&nbsp;&nbsp;"

                         If PostResults<>currentStatus Then                            
                            If rowCount <> 1 Then
                                Response.Write("<tr><td style='padding-top: 8px; text-align: right;' align='right' colspan='5'>Total " & Replace(currentStatus,"!","") & " #:" & DailyCount-1 & "</td></tr>")
                                LineCount = LineCount + 1
                            End If 
                            DailyCount = 1
                            currentStatus = PostResults
                        End If
						
						
						Response.Write("<tr>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & DailyCount & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & Received & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & InvoiceID & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & PostResults & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & PostStatusMessage & "</td>")
			            Response.Write("</tr>")


						rowCount = rowCount + 1
						
						LineCount = LineCount + 1
						
						
						rsDailyAPIInvoices.MoveNext
						
						If LineCount >= 20 Then 
					        Response.Write("</tbody>")
						    Response.Write("</table>")

							Call PageHeader

							Call SubHeader("Invoices","Body")
							
							Response.Write("<table style='margin-left:20px;width:1300px;'>")	
						        Response.Write("<thead>")
						            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
						            Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' width='2%'>Count</th>")
			   		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  width='28%'>Received</th>")
					                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' width='4%'>Invoice ID</th>")
					                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='13%'>Result</th>")
					                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' width='53%'>Message (if any)</th>")
						            Response.Write("</tr>")
						        Response.Write("</thead>")
						        Response.Write("<tbody>")

						End if

					Loop
                    Response.Write("<tr><td style='padding-top: 8px; text-align: right;' align='right' colspan='5'>Total " & Replace(currentStatus,"!","") & " #:" & DailyCount & "</td></tr>")
                    LineCount = LineCount + 1
				Else
					response.Write("<tr><td colspan='8'>No Invoice API Data</td></tr>")
				End If
				


				Response.Write("<tr style='border-top: 2px solid #ddd;'>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'><strong>Count:&nbsp;&nbsp;" & rowCount-1 & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'>&nbsp;&nbsp;</td>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'>&nbsp;&nbsp;</td>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'>&nbsp;&nbsp;</td>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'>&nbsp;&nbsp;</td>")
	            Response.Write("</tr>")
		        Response.Write("</tbody>")
		    Response.Write("</table>")


	rsDailyAPIPartnersLoop.MoveNext
Loop
End If

Sub PageHeader


	LineCount = 0	
 	PageNo = PageNo + 1

	If PageNo > 1 Then Response.Write("<div style='page-break-before: always'>")

 	Response.Write("<div style='width:100%;'>")

 	Response.Write("<img src='/clientfiles/" & ClientKey & "/logos/logo.png' style='float:left; margin-top:30px;'><center><h1 >DAILY API LOG <Br>BY PARTNER SORTED BY RESULT"  & "</h1><h2 class='generated' >Generated " & WeekDayName(WeekDay(DateValue(Now()))) & "&nbsp;" &  Now() & "</h2></center>")

 	Response.Write("</div><BR><BR>")

 	If PageNo > 1 Then Response.Write("</div>") 	
End Sub

Sub SubHeader(passedSection, passedTopOrBody)

	If passedSection="Orders" Then
		Response.Write("<hr>")
		Response.Write("<h2>Partner: " & GetPartnerNameByAPIKey(currentPartnerAPIKey) & "</h2>")
		Response.Write("<hr>")
				    	
	   	Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Order posts received " & FormatDateTime(dateadd("d",-1,Date())) & "&nbsp;" &  FormatDateTime(OrderCutoffTime,3) & " through ")
		If Time() > cDate(OrderCutoffTime) Then
		   	Response.Write(FormatDateTime(Date()) &  "&nbsp;" & FormatDateTime(OrderCutoffTime,3) & " </h4><br>")
		Else
		   	Response.Write(FormatDateTime(Date()) &  "&nbsp;" & FormatDateTime(Now(),3) & " </h4><br>")			
		End If

	End If
	
	If passedSection="Invoices" Then
		Response.Write("<hr>")
		Response.Write("<h2>Partner: " & GetPartnerNameByAPIKey(currentPartnerAPIKey) & "</h2>")
		Response.Write("<hr>")
				    	
	    	    
	   	Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Invoice posts received " & FormatDateTime(dateadd("d",-1,Date())) & "&nbsp;" &  FormatDateTime(InvoiceCutoffTime,3) & " through ")
		If Time() > cDate(InvoiceCutoffTime) Then
		   	Response.Write(FormatDateTime(Date()) &  "&nbsp;" & FormatDateTime(InvoiceCutoffTime,3) & " </h4><br>")
		Else
		   	Response.Write(FormatDateTime(Date()) &  "&nbsp;" & FormatDateTime(Now(),3) & " </h4><br>")			
		End If

	End If


End Sub

%></div></body></html>