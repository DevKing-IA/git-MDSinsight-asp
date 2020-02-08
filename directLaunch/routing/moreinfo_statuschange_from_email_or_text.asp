<% @ Language = VBScript %>
<%
Response.Buffer = True
Response.Expires = 0
Response.Clear
%>
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/Settings.asp"-->
<!--#include file="../../inc/InsightFuncs_Routing.asp"-->
<!--#include file="header-moreinfo_statuschange_from_email_or_text.asp"-->
<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way

CustID = Request.QueryString("c")
ClientKey =  Request.QueryString("cl")
UserNo = Request.QueryString("u")
InvoiceNumber = Request.QueryString("i")

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
	cnnVar = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	cnnVar = cnnVar & ";Database=" & Recordset.Fields("dbCatalog")
	cnnVar = cnnVar & ";Uid=" & Recordset.Fields("dbLogin")
	cnnVar = cnnVar & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	dummy = MUV_Write("ClientCnnString",cnnVar)
	Recordset.close
	Connection.close	
	Session("ClientCnnString") = MUV_READ("ClientCnnString") 'bacause some functions use the session var
End If	



If CustID = "" Then
	%>MDS Insight is unable to show more info due to a blank customer id. Please contact your administrator.<% 
	Response.End
End If
If UserNo = "" Then
	%>MDS Insight is unable to show more info due to a blank usserno. Please contact your administrator.<% 
	Response.End
End If


'Now the code to show the ticket info
%>


<style type="text/css">
.input-lg::-webkit-input-placeholder, textarea::-webkit-input-placeholder {
  color: #666;
}
.input-lg:-moz-placeholder, textarea:-moz-placeholder {
  color: #666;
}
.checkboxes label{
	font-weight: normal;
	margin-right: 20px;
}
.close-service-client-output{
	text-align: left;
}
.ticket-details{
	margin-bottom: 15px;
}

.alert{
	padding: 5px 5px 20px 5px;
}

.back-arrow{
	color: #fff;
	text-decoration: none;
} 

.back-arrow:hover{
	color:#ccc;
}

 
.ac-container{
	width: 100%;
 	text-align: left;
  }
.ac-container label{
	margin-top: 20px;
	float: left;
	width: 100%;
	font-family:  Arial, sans-serif;
	padding: 5px 20px;
	position: relative;
	z-index: 20;
	display: block;
 	cursor: pointer;
	color: #777;
	text-shadow: 1px 1px 1px rgba(255,255,255,0.8);
	line-height: 33px;
	font-size: 14px;
	background: #ffffff;
	background: -moz-linear-gradient(top, #ffffff 1%, #eaeaea 100%);
	background: -webkit-gradient(linear, left top, left bottom, color-stop(1%,#ffffff), color-stop(100%,#eaeaea));
	background: -webkit-linear-gradient(top, #ffffff 1%,#eaeaea 100%);
	background: -o-linear-gradient(top, #ffffff 1%,#eaeaea 100%);
	background: -ms-linear-gradient(top, #ffffff 1%,#eaeaea 100%);
	background: linear-gradient(top, #ffffff 1%,#eaeaea 100%);
	filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#ffffff', endColorstr='#eaeaea',GradientType=0 );
	box-shadow: 
		0px 0px 0px 1px rgba(155,155,155,0.3), 
		1px 0px 0px 0px rgba(255,255,255,0.9) inset, 
		0px 2px 2px rgba(0,0,0,0.1);
}
.ac-container label:hover{
	background: #fff;
}
.ac-container input:checked + label,
.ac-container input:checked + label:hover{
	background: #c6e1ec;
	color: #3d7489;
	text-shadow: 0px 1px 1px rgba(255,255,255, 0.6);
	box-shadow: 
		0px 0px 0px 1px rgba(155,155,155,0.3), 
		0px 2px 2px rgba(0,0,0,0.1);
}
.ac-container label:hover:after,

<%
Response.Write(".ac-container input:checked + label:hover:after{")
Response.Write("	content: '';")
Response.Write("	position: absolute;")
Response.Write("	width: 24px;")
Response.Write("	height: 24px;")
Response.Write("	right: 13px;")
Response.Write("	top: 7px;")
Response.Write("	background: url('../../../../" & BaseURL & "/img/accordion/arrow_down.png') no-repeat center;")
Response.Write("}"}

Response.Write(".ac-container input:checked + label:hover:after{"}
Response.Write("	background-image: url('../../../../" & BaseURL & "img/accordion/arrow_up.png');"}
Response.Write("}"}

%>
.ac-container input{
	display: none;
}
.ac-container article{
	background: rgba(255, 255, 255, 0.5);
	margin-top: -3px;
	overflow: hidden;
 	position: relative;
	z-index: 10;
	-webkit-transition: height 0.3s ease-in-out, box-shadow 0.6s linear;
	-moz-transition: height 0.3s ease-in-out, box-shadow 0.6s linear;
	-o-transition: height 0.3s ease-in-out, box-shadow 0.6s linear;
	-ms-transition: height 0.3s ease-in-out, box-shadow 0.6s linear;
	transition: height 0.3s ease-in-out, box-shadow 0.6s linear;
 }
.ac-container article p{
	font-style: italic;
	color: #777;
	line-height: 23px;
	font-size: 14px;
	padding: 20px;
	text-shadow: 1px 1px 1px rgba(255,255,255,0.8);
}
.ac-container input:checked ~ article{
	-webkit-transition: height 0.5s ease-in-out, box-shadow 0.1s linear;
	-moz-transition: height 0.5s ease-in-out, box-shadow 0.1s linear;
	-o-transition: height 0.5s ease-in-out, box-shadow 0.1s linear;
	-ms-transition: height 0.5s ease-in-out, box-shadow 0.1s linear;
	transition: height 0.5s ease-in-out, box-shadow 0.1s linear;
	box-shadow: 0px 0px 0px 1px rgba(155,155,155,0.3);
}
.ac-container input:checked ~ article.ac-small{
padding:10px 20px 30px 20px;
width: 100%;
float: left;
background: #fff;
display: block;
 }

article.ac-small{
	display: none;
}

h3{
	margin-top: 0px;
}

.btn{
	white-space: normal !important;
}
 
</style>


<h1 class="fieldservice-heading">Delivery Details</h1>

<div class="container-fluid">
 <%

 
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (MUV_READ("ClientCnnString"))
'Now get all the invoices for that customer - include truck in case another truck has this customer alsp
SQL = "Select * from RT_DeliveryBoard Where CustNum = '" & CustID  & "'"
SQL = SQL & " AND TruckNumber = '" & GetTruckNumberByUser(UserNo)  & "'"
SQL = SQL & " AND IvsNum = " & InvoiceNUmber
SQL = SQL & " Order by DeliveryStatus, SequenceNumber"
Set rsInvoices = Server.CreateObject("ADODB.Recordset")
rsInvoices.CursorLocation = 3 
Set rsInvoices = cnn8.Execute(SQL)

If not rsInvoices.EOF Then

	Do While Not rsInvoices.EOF
		
		Response.Write("<div class='row alert alert-warning'>")
		Response.Write("<div class='col-lg-6 col-md-6 col-sm-6 col-xs-6'>")
		Response.Write("<br><h3>Invoice #: <strong>" & rsInvoices("IvsNum") & "</strong></h3><br>") 
		If GetPONumberByInvoiceNum(rsInvoices("IvsNum"))  <> "" Then Response.Write("Cust PO #: " & GetPONumberByInvoiceNum(rsInvoices("IvsNum")) & "<br>")
		Response.Write("</div>")
		%>					
		
		<!-- accordion starts here !-->
		<div class="col-lg-6 col-md-6 col-sm-12 col-xs-12">
		
			<% If DelBoardDontShowDeliveryLineItems() = False Then %>
			
				<section class="ac-container">
					<div>
						<input id="ac-<%=rsInvoices("IvsNum")%>" name="accordion-<%=rsInvoices("IvsNum")%>" type="checkbox" />
						<label for="ac-<%=rsInvoices("IvsNum")%>">
						<%
						If GetNumberOfLinesByInvoiceNumber(rsInvoices("IvsNum")) < 2 Then
							Response.Write(GetNumberOfLinesByInvoiceNumber(rsInvoices("IvsNum")) & " line - tap to expand / collapse")
						Else
							Response.Write(GetNumberOfLinesByInvoiceNumber(rsInvoices("IvsNum")) & " lines - tap to expand / collapse")
						End If
						%>
						</label>
						<article class="ac-small">
							<table>
							<%
							Response.Write("<td>Qty</td>")
							Response.Write("<td>&nbsp;&nbsp;</td>")
							Response.Write("<td>Item#</td>")
							Response.Write("<td>&nbsp;</td>")									
							Response.Write("<td>Description</td>")
	
							SQL = "Select * from InvoiceHistoryDetail Where IvsNum = " & rsInvoices("IvsNum") & " Order By IvsHistDetSequence"
							Set rsInvoiceDetail = Server.CreateObject("ADODB.Recordset")
							rsInvoiceDetail.CursorLocation = 3 
							Set rsInvoiceDetail = cnn8.Execute(SQL)
							If not rsInvoiceDetail.Eof Then
								Do While Not rsInvoiceDetail.EOF
									Response.Write("<tr>")
										Response.Write("<td>" & rsInvoiceDetail("itemQuantity") & "</td>")
										Response.Write("<td>&nbsp;&nbsp;</td>")
										Response.Write("<td>" & rsInvoiceDetail("partnum") & "</td>")
										Response.Write("<td>&nbsp;</td>")									
										If Len(Description) > 20 Then Description = Left(Description,20)
										Description = rsInvoiceDetail("prodDescription")
										Response.Write("<td>" & Description  & "</td>")
									Response.Write("<tr>")						
									rsInvoiceDetail.MoveNext
								Loop
							End If
							Set rsInvoiceDetail = Nothing
							%>
							</table>
						</article>
					</div>
				</section>
				
			<% End If %>
		 
		
		</div>
		<!-- accordion ends here !-->
		
</div>
		<%
		
		rsInvoices.MoveNext
	Loop				
End IF

Set rsInvoices = Nothing
cnn8.Close
Set cnn8 = Nothing

SelectedCustomer = CustID
%>
<!--#include file="../../mobile/drivers/deliveryboard/commonCustomerDisplaypanel.asp"-->  
</div>

 	
 
<!--#include file="../../inc/footer-field-service-noTimeout.asp"-->