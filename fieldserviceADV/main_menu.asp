<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/insightfuncs_service.asp"-->
<% ' Init Session Vars
Session("MemoNumber") = ""
Session("ServiceCustID") = ""
Session("ServiceCustName") = ""

Set cnn10 = Server.CreateObject("ADODB.Connection")
cnn10.open (Session("ClientCnnString"))
Set rs10 = Server.CreateObject("ADODB.Recordset")
rs10.CursorLocation = 3 
SQL10 = "SELECT * FROM Settings_FieldService"
Set rs10 = cnn10.Execute(SQL10 )
ShowPartsButtonValue = rs10.fields("ShowPartsButton")
%>

<style type="text/css">
	.fieldservice-heading{
		line-height: 1;
		padding-top: 5px;
		padding-bottom: 5px;
	}
	
 
 .fieldservice-container{
 	margin:0px !important;
 }
 
	</style>       

<% Session("MemoNumber") = "" 'Init to empty %>
 
<h1 class="fieldservice-heading" >Field Service Menu</h1>

  <!-- field service menu starts here !-->
  <div class="container-fluid fieldservice-container">
  <div class="row">

 <!-- tickets start here !-->
 <div class="tickets">

		<% 'Check for urgents
			If NumberOfUrgentServiceTicketsByTech(Session("UserNo")) <> 0 Then %>
				<!-- urgent tickets !-->
				<div class="col-lg-12">
					<p align="right"><strong><font color="red">URGENT: <%=NumberOfUrgentServiceTicketsByTech(Session("UserNo"))%></font></strong></p>
				</div>
				<!-- eof urgent tickets !-->
		<% End If%>
				
		<!-- pending tickets !-->
		<div class="col-lg-12">
			
			<p align="right"><strong>Pending Tkts: <%=NumberOfServiceTicketsDispatchedToTech(Session("UserNo"))%></strong></p>
			
			 
		</div>
		<!-- eof pending tickets !-->
		
		
		<!-- awaiting ack  !-->
		<div class="col-lg-12">
 			
			<p align="right"><strong>Awaiting Ack: <%=NumberOfServiceTicketsAwaitingACKFromTech(Session("UserNo"))%></strong></p>
			
 
 		</div>
			
			
		</div>
		<!-- eof awaiting ack  !-->
		
 </div>
 <!-- tickets end here !-->
 
<!-- buttons start here -->
<div class="row">

 <!-- button !-->	
 <div class="col-lg-12">
 <a href="main_OpenTickets.asp"> 
 	<button type="button" class="btn btn-success btn-block fieldservice-btn">Your Service Calls</button>
 </a>
</div>
 <!-- eof button !-->
 
 
  <!-- button !-->	
   <div class="col-lg-12">
 <a href="addServiceMemo_PassThru.asp"> 
 <button type="button" class="btn btn-success   btn-block fieldservice-btn">New Service Call</button>
 </a>
</div>
 <!-- eof button !-->
 
 
 
   <!-- button 
 <a href="../fieldservice/menu_assets.asp"> 
 <button type="button" class="btn btn-success   btn-block fieldservice-btn">Assets</button>
 </a>
 eof button !-->
 


<% If ShowPartsButtonValue = 1 Then %>
  <!-- button !-->	
   <div class="col-lg-12">
 <a href="requestParts.asp"> 
 <button type="button" class="btn btn-success   btn-block fieldservice-btn">Request Parts</button>
 </a>
</div>
 <!-- eof button !-->
<%End If %>
 
 <!-- button !-->
  <div class="col-lg-12">	
 <a href="../logout.asp"> 
 <button type="button" class="btn btn-success   btn-block fieldservice-btn">Logout</button>
 </a>
</div>
 <!-- eof button !-->
 
 
	  </div>
  </div>
  <!-- buttons end here !-->            

<!--#include file="../inc/footer-field-service-noTimeout.asp"-->