<!--#include file="inc/header-tech-and-driver.asp"-->
<!--#include file="../../inc/InsightFuncs_Routing.asp"-->
<% ' Init Session Vars
Session("MemoNumber") = ""
Session("ServiceCustID") = ""
Session("ServiceCustName") = ""

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

<script type="text/javascript">

	if (navigator.geolocation) {  
		navigator.geolocation.getCurrentPosition(showGPS, gpsError);
	} else {  
		document.cookie="gps=" + "0,0";  
	}
	

function gpsError(error) {
		document.cookie="gps=" + "0,0";  
}

function showGPS(position) {

	document.cookie="gps=" + position.coords.latitude+","+position.coords.longitude;
	
	
}

</script>

<% Session("MemoNumber") = "" 'Init to empty %>
 
<h1 class="fieldservice-heading" >Tech / Driver Main</h1>

  <!-- field service menu starts here !-->
  <div class="container-fluid fieldservice-container">
  <div class="row">

<!-- tickets start here !-->
<div class="tickets">
		
	<!-- service tickets !-->
	<div class="col-lg-12">
		<p align="left"><strong>Service: <%=NumberOfServiceTicketsDispatchedToTech(Session("UserNo"))%></strong></p>
	</div>
	<!-- eof service tickets !-->
		
	<!-- awaiting ack  !-->
	<div class="col-lg-12">
		<p align="left"><strong>Awaiting Ack: <%=NumberOfServiceTicketsAwaitingACKFromTech(Session("UserNo"))%></strong></p>
	</div>

	<!-- delivery tickets !-->
	<div class="col-lg-12">
		<p align="left"><strong>Delivery: <%=GetRemainingStopsByUserNo(Session("UserNo"))%></strong></p>
	</div>
	<!-- eof delivery tickets !-->
			
</div>
<!-- eof awaiting ack  !-->


		
 </div>
 <!-- tickets end here !-->
 
<!-- buttons start here -->
<div class="row">

 <!-- button !-->	
 <div class="col-lg-12">
 <a href="main.asp"> 
 <button type="button" class="btn btn-success  btn-block   fieldservice-btn">Your Stops</button>
 </a>
</div>
 <!-- eof button !-->
 

<!-- button !-->	
 <div class="col-lg-12">
&nbsp;
</div>
 <!-- eof button !-->

 
  <!-- button !-->	
   <div class="col-lg-12">
 <a href="addServiceMemo_PassThru.asp"> 
 <button type="button" class="btn btn-success   btn-block fieldservice-btn">New Service Ticket</button>
 </a>
</div>
 <!-- eof button !-->
 
 
 
   <!-- button 
 <a href="../../fieldservice/menu_assets.asp"> 
 <button type="button" class="btn btn-success   btn-block fieldservice-btn">Assets</button>
 </a>
 eof button !-->
 
<%If filterChangeModuleOn() = True and prevMaintModuleOn() <> True Then%>
	<!-- button !-->
	 <div class="col-lg-12">	
	<a href="filterchanges.asp"> 
		<button type="button" class="btn btn-warning  btn-block fieldservice-btn">Filter Changes</button>
	</a>
</div>
	<!-- eof button !-->
<% ElseIf filterChangeModuleOn() <> True and prevMaintModuleOn() = True Then %>
	<!-- button !-->
	 <div class="col-lg-12">	
	<a href="filterchanges.asp"> 
		<button type="button" class="btn btn-warning  btn-block fieldservice-btn">PM Calls</button>
	</a>
</div>
	<!-- eof button !-->
<% ElseIf filterChangeModuleOn() = True and prevMaintModuleOn() = True Then %>
	<!-- button !-->	
	 <div class="col-lg-12">
	<a href="filterchanges.asp"> 
		<button type="button" class="btn btn-warning  btn-block fieldservice-btn">Filter Changes / PM Calls</button>
	</a>
</div>
	<!-- eof button !-->
<%End If%>
   
 <!-- button !-->
  <div class="col-lg-12">	
 <a href="../../logout.asp"> 
 <button type="button" class="btn btn-success   btn-block fieldservice-btn">Logout</button>
 </a>
</div>
 <!-- eof button !-->
 
 
	  </div>
  </div>
  <!-- buttons end here !-->            

<!--#include file="inc/footer-tech-and-driver.asp"-->