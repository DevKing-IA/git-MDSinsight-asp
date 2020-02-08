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
 
 
	</style>       


 
<h1 class="fieldservice-heading" ><a href="main_menu.asp" class="left-arrow"><i class="fa fa-arrow-left pull-left" aria-hidden="true"></i></a> Stops
<a href="main.asp" class="left-arrow"><i class="fa fa fa-list pull-right" aria-hidden="true"></i></a>
</h1>

<div class="container-fluid fieldservice-container">

<% If showForceNextStopMsg = "True" Then %>
	<h1 class="fieldservice-heading" style="background-color:#d9534f;">PLEASE SELECT A NEXT STOP TO CONTINUE PAST THIS SCREEN</h1>
<% End If %>



 
			 

<%
'Lookup customers because there can be more than 1 invoice

SQL = "SELECT CustNum, Priority, AMorPM, MIN(SequenceNumber) AS Expr1, Count(CustNum) AS Expr2, Max(Len(DeliveryStatus)) as Expr3, Max(ManualNextStop) as Expr4  FROM RT_DeliveryBoard "
SQL = SQL & "WHERE (CustNum IN "
SQL = SQL & "(SELECT CustNum FROM RT_DeliveryBoard AS RT_DeliveryBoard_1 "
SQL = SQL & "WHERE (TruckNumber = '" & GetTruckNumberByUser(Session("UserNo")) & "') GROUP BY CustNum)) GROUP BY CustNum, Priority, AMorPM ORDER BY Priority DESC, AMorPM DESC, Expr4 Desc, Expr3, Expr1"

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
			<div class="col-lg-7 col-md-7 col-sm-7 col-xs-7">
				<%


				Response.Write("<p align='left' style='color:red;font-weight:bold;'>") 
				If CustHasANYPriorityDelivery(rsCust("CustNum"),Session("UserNo")) = True Then
					If rsCust("Expr2") = 1 Then
						Response.Write("Priority Delivery") 
					End If
				End If
				Response.Write("</p>") 
				
				Response.Write("<p align='left' style='color:red;font-weight:bold;'>")
				If CustHasANYAMDelivery(rsCust("CustNum"),Session("UserNo")) = True Then
					If rsCust("Expr2") = 1 Then
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
						Response.Write("<li>" & rsCust2("Name") & "</li>")
						Response.Write("<li>" & rsCust2("Addr1") & "</li>")
						Response.Write("<li>" & rsCust2("Addr2") & "</li>")
					%></ul><br>
				<%End If
				Set rsCust2 = Nothing
				%>
								
			</div>
			
			<!-- delivered / no delivery / reset !-->
			<div class="col-lg-5 col-md-5 col-sm-5 col-xs-5">
				<p align="right">


					<%If CustHasANYDelivery(rsCust("CustNum"),Session("UserNo")) = True Then %>
						<a href="tap_reset_customer.asp?c=<%=rsCust("CustNum")%>&r=stop" class="btn btn-danger" role="button">Reset</a>
						<br><%=GetDeliveryStatusByCust (rsCust("CustNum"))%>
					<% Else
						If rsCust("Expr4") <> 1 Then %>
							<a href="tap_nextStop.asp?c=<%=rsCust("CustNum")%>" class="btn btn-primary" role="button">Set As Next</a>
						<% Else %>
							<a href="tap_nextStop_undo.asp?c=<%=rsCust("CustNum")%>" class="btn btn-danger" role="button">Undo Next</a>
						<% End If							
					 End If%>
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