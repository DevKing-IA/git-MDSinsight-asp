<% @ Language = VBScript %>
<%
Response.Buffer = True
Response.Expires = 0
Response.Clear
%>
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/Settings.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.asp"-->
<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way

ServiceTicketNumber = Request.QueryString("t")
UserNumber = Request.QueryString("u")
CustNum = Request.QueryString("c")
ClientKey =  Request.QueryString("cl")

SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
'Response.Write("InsightCnnString:" & InsightCnnString & "<br>")
Connection.Open InsightCnnString


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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Please contact your administrator.<%
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


'Response.Write("ServiceTicketNumber :" & ServiceTicketNumber & "<br>")
'Response.Write("UserNumber :" & UserNumber & "<br>")
'Response.Write("CustNum :" & CustNum & "<br>")
'Response.End

If ServiceTicketNumber = "" Then
	%>MDS Insight is unable to acknowledge this dispatch due to a blank service ticket id. Please contact your administrator.<% 
	Response.End
End If
If UserNumber= "" Then
	%>MDS Insight is unable to acknowledge this dispatch due to a blank user number. Please contact your administrator.<% 
	Response.End
End If
If CustNum = "" Then
	%>MDS Insight is unable to acknowledge this dispatch due to a blank customer id. Please contact your administrator.<% 
	Response.End
End If


'Now the code to show the ticket info
%>

<!--#include file="header-moreinfo_dispatch_from_email_or_text.asp"-->

<style type="text/css">
	.fieldservice-heading{
		line-height: 1;
		padding-top: 5px;
		padding-bottom: 5px;
	}
	
	.btn-home{
		color: #fff;
		margin-top: -2px;
		margin-left: 5px;
		float: left;
 	}
	
 ul{
	 color: #666;
	 font-size: 13px;
	 text-transform: uppercase;
	 list-style-type: none;
	     -webkit-margin-before: 0px;
    -webkit-margin-after: 0px;
    -webkit-margin-start: 0px;
    -webkit-margin-end: 0px;
    -webkit-padding-start: 0px;
 }
 
 .enroute{
	 color: green;
 }
 
 .btn-spacing{
	 margin-bottom: 40px;
 }
 
 
 
.btn-block {
    width: auto;
    display: inline-block;
}
 
 .row{
 	/* flex-wrap: nowrap !important; */
 }

 
 @media (max-width: 767px) {
 	.mob-col{
 		/* width: auto !important;  */
 	}
 }

	</style>       


 
<h1 class="fieldservice-heading" >Ticket <%= ServiceTicketNumber %></h1>

<div class="container-fluid">

<%
'Now lookup the other info
SQL = "SELECT * From FS_ServiceMemos WHERE MemoNumber = '" & ServiceTicketNumber  & "'"
 
'Response.Write(SQL)

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rsCust = Server.CreateObject("ADODB.Recordset")
rsCust.CursorLocation = 3 


Set rs = cnn8.Execute(SQL)
			
	If not rs.EOF Then
						
			%>

			
			<!-- client info !-->
			<div class="col-lg-12 mob-col">
			
				<strong>
				
				<% If TicketIsUrgent(rs("MemoNumber")) Then %>
					<font color='red'>*THIS TICKET IS URGENT*</font><br>
				<% Else %>
					<font color='green'>*NO LONGER URGENT*</font><br>
				<% End If %>
				
				<%=GetTerm("Account")%>:&nbsp;<%=rs("AccountNumber")%><br>
				</strong><br>
				
				<%'Lookup cust info
				SQL = "Select Name,Addr1,Addr2,CityStateZip,Phone,Contact from AR_Customer where CustNum = '" & rs("AccountNumber") & "'"
				Set rsCust = cnn8.Execute(SQL)
				If not rsCust.Eof Then
				%>
				<ul><%
					Response.Write("<li>" & rsCust("Name") & "</li>")
					Response.Write("<li>" & rsCust("Addr1") & "</li>")
					Response.Write("<li>" & rsCust("Addr2") & "</li>")
					Response.Write("<li>" & rsCust("CityStateZip") & "</li>")
					Response.Write("<li>" & rsCust("Phone") & "</li>")
					Response.Write("<li>" & rsCust("Contact") & "</li>")
				%></ul>
				<%End If%> 
				
				<strong>
				<br><br><%=GetServiceTicketProblemByTicketNumber(ServiceTicketNumber)%><br>
				</strong><br>


			</div>
		
		<!-- eof client box !-->	
		
		<hr />

		<%
	Else
		%>No Service calls for you!<%
	End IF

cnn8.close
Set rsCust = Nothing
Set rs = Nothing
Set cnn8 = Nothing				
%></div><!--#include file="../../inc/footer-field-service-noTimeout.asp"--><%

%>