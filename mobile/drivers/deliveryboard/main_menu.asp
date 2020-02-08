<!--#include file="../../../inc/header-deliveryboard-drivers-mobile.asp"-->

<!--#include file="../../../inc/InSightFuncs_Routing.asp"-->
<style type="text/css">
	.fieldservice-heading{
		line-height: 1;
		padding-top: 5px;
		padding-bottom: 5px;
	}
</style>       

<!--
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
-->

<h1 class="fieldservice-heading" >Main Menu</h1>

<!-- driver menu starts here !-->
<div class="container-fluid fieldservice-container">
	<div class="row">

		 <!-- summary stats start here !-->
		 <div class="tickets">
		
			<!-- completed  !-->
			<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6">
				<p align="left"><strong>Total Stops: <%= GetTotalStopsByUserNo(Session("UserNo"))%></strong></p>
			</div>
			<!-- eof completed  !-->

			<!-- remaining deliveries !-->
			<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6">
				<p align="right"><strong>Remaining: <%=GetRemainingStopsByUserNo(Session("UserNo"))%></strong></p>
			</div>
			<!-- eof remaining deliveries !-->

		
		</div>
		<!-- eof summary stats !-->
		
		 <!-- summary stats start here !-->
		 <div class="tickets">
			<!-- priority  !-->
			<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6">
				<p align="left" style="color:red;font-weight:bold;">PRIORITY: <%= GetTotalPriorityStopsByUserNo(Session("UserNo"))%></p>
			</div>
			<!-- eof priority  !-->
			<!-- priority remaining  !-->
			<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6">
				<p align="right" style="color:red;font-weight:bold;">Remaining: <%= GetRemainingPriorityStopsByUserNo(Session("UserNo"))%></p>
			</div>
			<!-- eof priority remaining !-->
			
		</div>
		<!-- eof priority  stats !-->
		
		
		 <!-- summary stats start here !-->
		 <div class="tickets">
			<!-- am deliveries  !-->
			<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6">
				<p align="left" style="color:red;font-weight:bold;">AM: <%= GetTotalAMStopsByUserNo(Session("UserNo"))%></p>
			</div>
			<!-- eof am deliveries !-->
			<!-- am deliveries remaining  !-->
			<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6">
				<p align="right" style="color:red;font-weight:bold;">Remaining: <%= GetRemainingAMStopsByUserNo(Session("UserNo"))%></p>
			</div>
			<!-- eof am deliveries remaining !-->
			
		</div>
		<!-- eof priority  stats !-->
		
		
 
		<!-- your deliveries !-->	
		<a href="main.asp"> 
			<button type="button" class="btn btn-success  btn-block  col-lg-12 fieldservice-btn">Your Deliveries</button>
		</a>
		<!-- eof your deliveries !-->
 
		<!-- logout !-->	
		<a href="../../../logout.asp"> 
			<button type="button" class="btn btn-success   btn-block  col-lg-12 fieldservice-btn">Logout</button>
		</a>
		<!-- eof logout !-->
	</div>
</div>

<!--#include file="../../../inc/footer-field-service-noTimeout.asp"-->