<!--#include file="../../../inc/header-deliveryboard-drivers-mobile.asp"-->

<!--#include file="../../../inc/InSightFuncs_Routing.asp"-->
<%If Request.ServerVariables("REQUEST_METHOD") = "POST" Then 
	CustNum = Request.Form("txtCustNum")
Else
	CustNum = Request.QueryString("c")
End If%>

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
.ac-container input:checked + label:hover:after{
	content: '';
	position: absolute;
	width: 24px;
	height: 24px;
	right: 13px;
	top: 7px;
	background: transparent url(../../../img/accordion/arrow_down.png) no-repeat center center;	
}
.ac-container input:checked + label:hover:after{
	background-image: url(../../../img/accordion/arrow_up.png);
}
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

<h1 class="fieldservice-heading" >
<a class="back-arrow pull-left" href="main.asp" role="button"><i class="fa fa-arrow-left" aria-hidden="true"></i> </a>Delivery Details</h1>

<div class="container-fluid">
 <%

 
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
'Now get all the invoices for that customer - include truck in case another truck has this customer alsp
SQL = "Select * from RT_DeliveryBoard Where CustNum = '" & CustNum & "'"
SQL = SQL & " AND TruckNumber = '" & GetTruckNumberByUser(Session("UserNo"))  & "'"
SQL = SQL & " Order by DeliveryStatus, SequenceNumber"
Set rsInvoices = Server.CreateObject("ADODB.Recordset")
rsInvoices.CursorLocation = 3 

Set rsInvoices = cnn8.Execute(SQL)

If not rsInvoices.EOF Then

	Do While Not rsInvoices.EOF
		
		Response.Write("<div class='row alert alert-warning'>")
		Response.Write("<div class='col-lg-6 col-md-6 col-sm-6 col-xs-6'>")
		Response.Write("<br><h3>Invoice #: <strong>" & rsInvoices("IvsNum") & "</strong></h3><br>") 
		
		If rsInvoices("Priority") = 1 Then
			Response.Write("<h3>PRIORITY DELIVERY</h3>")
		End If
		
		If rsInvoices("AMorPM") = "AM" Then
			Response.Write("<h3>AM DELIVERY</h3>")
		End If
		
		If GetPONumberByInvoiceNum(rsInvoices("IvsNum"))  <> "" Then Response.Write("Cust PO #: " & GetPONumberByInvoiceNum(rsInvoices("IvsNum")) & "<br>")
		Response.Write("</div>")
		%>					
		<!-- buttons !-->
		<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6" align="right">
			<%
			If GetDeliveryStatusByInvoice(rsInvoices("IvsNum")) <> "" Then %>
				<br>
				<a href="tap_reset_invoice.asp?i=<%=rsInvoices("IvsNum")%>&c=<%=rsInvoices("CustNum")%>" class="btn btn-danger" role="button">Reset</a>
				<br><br><%=GetDeliveryStatusByInvoice(rsInvoices("IvsNum"))%>
			<% Else %>
				<br>
				<a href="driver_comments.asp?i=<%=rsInvoices("IvsNum")%>&c=<%=rsInvoices("CustNum")%>&s=d" class="btn btn-success" role="button">Delivered</a><br><br>
				<a href="driver_comments.asp?i=<%=rsInvoices("IvsNum")%>&c=<%=rsInvoices("CustNum")%>&s=n" class="btn btn-warning" role="button">Not Delivered</a><br><br>
				<% If DeliveryInProgress(rsInvoices("IvsNum")) <> True Then %>
					<a href="mark_in_progress.asp?i=<%=rsInvoices("IvsNum")%>&c=<%=rsInvoices("CustNum")%>" class="btn btn-primary" role="button">Mark as In Progress</a><br>
				<% Else %>
					<a href="mark_not_in_progress.asp?i=<%=rsInvoices("IvsNum")%>&c=<%=rsInvoices("CustNum")%>" class="btn btn-primary" role="button">Mark as NOT In Progress</a><br>
				<%End If %>
				<!-- <a href="tap_delivered_invoice.asp?i=<%=rsInvoices("IvsNum")%>&c=<%=rsInvoices("CustNum")%>" class="btn btn-success" role="button">Delivered</a><br><br>!-->
				<!-- <a href="tap_no_delivery_invoice.asp?i=<%=rsInvoices("IvsNum")%>&c=<%=rsInvoices("CustNum")%>" class="btn btn-danger" role="button">Not Delivered</a><br><br>!-->
			<% End If %>
		</div>
		<!-- eof buttons !-->
		
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

SelectedCustomer = CustNum
%>
<!--#include file="commonCustomerDisplaypanel.asp"-->  
</div>

 	
 
<!--#include file="../../../inc/footer-field-service-noTimeout.asp"-->