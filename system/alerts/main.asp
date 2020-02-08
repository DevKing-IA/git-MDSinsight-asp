<!--#include file="../../inc/header.asp"-->


<!-- function that gets the value of the tab when it is clicked and then
updates the value of a hidden form field so when the page posts, it returns
back to the tab that was previously opened -->

<script type="text/javascript">
	$(function () {
		$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
		var target = $(e.target).attr("href");
		$('input[name="txtTab"]').val(target);
		//alert(target);
		});
	})
</script>

 
<style type="text/css">
 	.email-table{
		width:46%;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
	}

	.nav-tabs>li>a{
		background: #f5f5f5;
		border: 1px solid #ccc;
		color: #000;
	}

	.nav-tabs>li>a:hover{
		border: 1px solid #ccc;
	}

	.nav-tabs>li.active>a, .nav-tabs>li.active>a:focus, .nav-tabs>li.active>a:hover{
		color: #000;
		border: 1px solid #ccc;
	}
 </style>



<h1 class="page-header"><i class="fa fa-fw fa-exclamation"></i> Alerts & Notifications</h1>

	

<!-- tabs start here !-->
<div class="row">
	<div class="col-lg-12">

		<!-- tabs navigation !-->
		<ul class="nav nav-tabs" role="tablist" id="tablist">
			<% If userIsAdmin(Session("UserNo"))  Then %> 
				<li role="presentation" class="active"><a href="#System" aria-controls="manage" role="tab" data-toggle="tab">System<br>&nbsp;</a></li>
			    <li role="presentation"><a href="#ServiceNumTicks" aria-controls="manage" role="tab" data-toggle="tab">Service<br>(# tickets)</a></li>
			<% Else %>
				<!--<li role="presentation"><a href="#ServiceNumTicks" aria-controls="manage" role="tab" data-toggle="tab">Service<br>(# tickets)</a></li>-->
			<% End If %>
   		    <li role="presentation"><a href="#ServiceElapsed" aria-controls="manage" role="tab" data-toggle="tab">Service<br>(time based)</a></li>
   		    <li role="presentation"><a href="#ServiceOtherConditions" aria-controls="manage" role="tab" data-toggle="tab">Service<br>(other)</a></li>
		    <% If MUV_READ("nightBatchModuleOn") = "Enabled" Then %>
		    	<li role="presentation"><a href="#NightBatchAlerts" aria-controls="manage" role="tab" data-toggle="tab">Night<br>Batch</a></li>
		    <% End If %>
		    <% If MUV_READ("routingModuleOn") = "Enabled" Then %>
		    	<li role="presentation"><a href="#DeliveryBoardAlerts" aria-controls="manage" role="tab" data-toggle="tab">Delivery<br>Board</a></li>
		    <% End If %>
		    <% If MUV_READ("OrderAPIModuleOn") = "Enabled" Then %>
		    	<li role="presentation"><a href="#OrderAPIAlerts" aria-controls="manage" role="tab" data-toggle="tab">Order<br>API</a></li>
		    <% End If %>
		    
		</ul>
		<!-- eof tabs navigation !-->
			
		<!-- tabs content !-->
		<div class="tab-content">
		
			<!-- SYSTEM Tab !-->
			<!--#include file="tabs/tabSystem.asp"-->
			<!-- eof SYSTEM Tab !-->	
        
	        <!-- Service (# tickets) Tab !-->
			<!--#include file="tabs/tabServiceNumTicks.asp"-->
			<!-- eof Service (# tickets) Tab !-->
		
			<!-- Service Elapsed Time !-->
			<!--#include file="tabs/tabServiceElapsed.asp"-->
			<!-- Service Elapsed Time !-->

			<!-- Other Criteria, like Service Technician Notes Tab !-->
			<!--#include file="tabs/tabServiceOtherConditions.asp"-->
			<!-- eof Other Criteria tab !-->

			<!-- Night Batch Alert Tab !-->
			<!--#include file="tabs/tabNightbatch.asp"-->
			<!-- eof Night Batch Alert Tab !-->

			<!-- Delivery Board Alert Tab !-->
            <!--#include file="tabs/tabDeliveryboard.asp"-->
			<!-- eof Delivery Board Alert Tab !-->

			<!-- Order API Alert Tab !-->
            <!--#include file="tabs/tabOrderAPI.asp"-->
			<!-- eof Order API Alert Tab !-->
			
			
		</div>
	</div>
</div>
<!-- tabs end here !-->
    

<!--#include file="../../inc/footer-main.asp"-->