<!--#include file="../../../inc/header-deliveryboard-drivers-mobile.asp"-->

<!--#include file="../../../inc/InSightFuncs_Routing.asp"-->
<style type="text/css">
	.fieldservice-heading{
		line-height: 1;
		padding-top: 5px;
		padding-bottom: 5px;
	}
	
	.btn-home{
		font-size: 11px;
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
 
 .pull-left{
	 margin-left: 5px;
 }
 
.btn-block {
    width: auto;
    display: inline-block;
}

p{
	margin: 0;
}

.fieldservice-container{
	margin: 0px;
}

.alert{
	padding: 5px 0px 0px 0px;
	margin-bottom: 0px;
}

.btn-spacing{
	margin-bottom: 10px;
}

hr{
	margin-top: 10px;
	margin-bottom: 10px;
}

.btn{
	margin-bottom: 15px;
}

.left-arrow{
	color: #fff;
}

.left-arrow:hover{
	color: #ccc;
}

.red-text{
	color: red;
}
 
	</style>       


 
<h1 class="fieldservice-heading" ><a href="main_menu.asp" class="left-arrow"><i class="fa fa-arrow-left pull-left" aria-hidden="true"></i></a> Deliveries
<a href="viewstops.asp" class="left-arrow"><i class="fa fa fa-list pull-right" aria-hidden="true"></i></a>
</h1>

<div class="container-fluid fieldservice-container">
 
<%
'Lookup customers because there can be more than 1 invoice

'SQL = "SELECT CustNum, Priority AS Expr5, MIN(SequenceNumber) AS Expr1, Count(CustNum) AS Expr2, Max(Len(DeliveryStatus)) as Expr3, Max(ManualNextStop) as Expr4 FROM RT_DeliveryBoard "
'SQL = SQL & "WHERE (CustNum IN "
'SQL = SQL & "(SELECT CustNum FROM RT_DeliveryBoard AS RT_DeliveryBoard_1 "
'SQL = SQL & "WHERE (TruckNumber = '" & GetTruckNumberByUser(Session("UserNo")) & "') GROUP BY CustNum)) GROUP BY CustNum ORDER BY Expr5 DESC, Expr4 Desc,Expr3, Expr1"


SQL = "SELECT CustNum, Priority, AMorPM, MIN(SequenceNumber) AS Expr1, COUNT(CustNum) AS Expr2, MAX(LEN(DeliveryStatus)) AS Expr3, MAX(ManualNextStop) AS Expr4 "
SQL = SQL & "FROM RT_DeliveryBoard "
SQL = SQL & "WHERE  (CustNum IN "
             SQL = SQL & "(SELECT CustNum "
             SQL = SQL & "FROM RT_DeliveryBoard AS RT_DeliveryBoard_1 "
             SQL = SQL & "WHERE (TruckNumber = '" & GetTruckNumberByUser(Session("UserNo")) & "') "
             SQL = SQL & "GROUP BY CustNum)) "
SQL = SQL & "GROUP BY CustNum, Priority, AMorPM "
SQL = SQL & "ORDER BY Priority DESC, AMorPM DESC, Expr4 DESC, Expr3, Expr1 "

'Response.write(SQL)

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rsCust = Server.CreateObject("ADODB.Recordset")
rsCust.CursorLocation = 3 
Set rsCust = cnn8.Execute(SQL)
			
	If not rsCust.EOF Then

		Do While Not rsCust.EOF

			If CustHasANYPriorityDelivery(rsCust("CustNum"),Session("UserNo")) = True Then
			
				Response.Write("<div class='row alert alert-danger'>")
				
			ElseIf CustHasANYAMDelivery(rsCust("CustNum"),Session("UserNo")) = True Then
			
				Response.Write("<div class='row alert alert-danger'>")
				
			Else
				Response.Write("<div class='row alert alert-warning'>")
			End If
			%>
			 
			<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6">
				<%
				Response.Write("<strong>") 
				If rsCust("Expr2") = 1 Then Response.Write("1 Invoice") Else Response.Write(rsCust("Expr2") & " Invoices")
				Response.Write("</strong>") 

				Response.Write("<p align='left' style='color:red;font-weight:bold;'>") 
				If CustHasANYPriorityDelivery(rsCust("CustNum"),Session("UserNo")) = True Then
					If rsCust("Expr2") = 1 Then
						Response.Write("Priority Delivery") 
					Else
						'Response.Write("Has Priority Invoices")
						Response.Write("Priority Delivery")
					End If
				End If
				Response.Write("</p>") 

				Response.Write("<p align='left' style='color:red;font-weight:bold;'>")
				If CustHasANYAMDelivery(rsCust("CustNum"),Session("UserNo")) = True Then
					If rsCust("Expr2") = 1 Then
						Response.Write("AM Delivery") 
					Else
						'Response.Write("Has AM Invoices")
						Response.Write("AM Delivery")
					End If
				End If
				Response.Write("</p>") 
										
				'Lookup cust info
				SQL = "Select CustNum,Name,Addr1,Addr2,CityStateZip,Phone,Contact from AR_Customer where CustNum='" & rsCust("CustNum") & "'"
				Set rsCust2 = cnn8.Execute(SQL)
				If not rsCust2.Eof Then
					%> 
					<ul><%
						Response.Write("<li><strong>" & rsCust2("Name") & "</strong></li>")
						Response.Write("<li>" & rsCust2("Addr1") & "</li>")
						Response.Write("<li>" & rsCust2("Addr2") & "</li>")
					%></ul><br>
				<% End If %>
								
				<form method="post" action="viewInvoices.asp" name="frmViewInvocies" id="frmInvoices">
					<input type='hidden' id='txtCustNum' name='txtCustNum' value='<%=rsCust("CustNum")%>'>
					<button type="submit" class="btn btn-primary btn-block btn-spacing" >Details / Partial</button>
				</form>
				
			</div>
			
			<!-- delivered / no delivery / reset !-->
			<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6">
				<p align="right">

					<%If CustHasANYDelivery(rsCust("CustNum"),Session("UserNo")) = True Then %>
						<a href="tap_reset_customer.asp?c=<%=rsCust("CustNum")%>&r=main" class="btn btn-danger" role="button">Reset</a><br>
						<br><%=GetDeliveryStatusByCust(rsCust("CustNum"))%>
					<% Else %>
						<a href="driver_comments.asp?c=<%=rsCust("CustNum")%>&s=d" class="btn btn-success" role="button">Delivered</a><br>
						<a href="driver_comments.asp?c=<%=rsCust("CustNum")%>&s=n" class="btn btn-warning" role="button">Not Delivered</a><br>
						<!-- <a href="tap_delivered_customer.asp?c=<%=rsCust("CustNum")%>" class="btn btn-success" role="button">Delivered</a><br>
						<a href="tap_no_delivery_customer.asp?c=<%=rsCust("CustNum")%>" class="btn btn-warning" role="button">Not Delivered</a><br>!-->										
					<% End If
					Set rsCheckCust = Nothing
					%>
				 

				<% If DeliveryInProgressByCust(rsCust("CustNum")) = True Then
					Response.Write("<span class='red-text'>*Delivery In Progress*</span>")				
				End If %>
				</p>
			</div>
			<!-- eof delivered / no delivery / reset !-->
	
		</div>
		<hr>
		
		<%
		rsCust.movenext
	loop
Else
	%>No deliveries for you!<%
End IF

cnn8.close
Set rsCust = Nothing
Set cnn8 = Nothing				
%><!--#include file="../../../inc/footer-field-service-noTimeout.asp"-->