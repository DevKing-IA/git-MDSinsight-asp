<!--#include file="../../../../inc/header.asp"-->

<!-- bootstrap timepicker !-->
<script type="text/javascript" src="http://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>	
<link href="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.css" rel="stylesheet" type="text/css">
<script src="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.min.js" type="text/javascript"></script>
<!-- eof bootstrap timepicker !-->

<!-- bootstrap multiselect !-->
<link rel="stylesheet" href="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.css" type="text/css">
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.js"></script>
<!-- eof bootstrap multiselect !-->

<script>

	function showSavingChangesDiv() {
	  document.getElementById('PleaseWaitPanel').style.display = "block";
	  setTimeout(function() {
	    document.getElementById('PleaseWaitPanel').style.display = "none";
	  },1500);
	   
	}

	$(document).ready(function() {
					
		$('#modalDailyAPIActivityByPartnerReportScheduler').on('show.bs.modal', function(e) {
		    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
				cache: false,
				data: "action=GetContentForDailyAPIActivityByPartnerReportScheduler",
				success: function(response)
				 {
	               	 $modal.find('#modalDailyAPIActivityByPartnerReportSchedulerContent').html(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalDailyAPIActivityByPartnerReportSchedulerContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
			
		});	
		
		
		
        $('#APIDailyActivityReportAdditionalEmails').on('show.bs.modal', function (e) {
            var $modal = $(this);
        });	
		
		
		
		$('.panel .panel-body').css('display','none');
		$('.panel-heading span.clickable').addClass('panel-collapsed');
		$('.panel-heading span.clickable').find('i').removeClass('glyphicon-chevron-up').addClass('glyphicon-chevron-down');

		$(document).on('click', '.panel-heading span.clickable', function(e){
		    var $this = $(this);
			if(!$this.hasClass('panel-collapsed')) {
				$this.parents('.panel').find('.panel-body').slideUp();
				$this.addClass('panel-collapsed');
				$this.find('i').removeClass('glyphicon-chevron-up').addClass('glyphicon-chevron-down');
			} else {
				$this.parents('.panel').find('.panel-body').slideDown();
				$this.removeClass('panel-collapsed');
				$this.find('i').removeClass('glyphicon-chevron-down').addClass('glyphicon-chevron-up');
			}
		});
		
		
 		$("#toggle").click(function(){
 		
            if(!$('.panel-heading span.clickable').hasClass('panel-collapsed')) {
				$('.panel .panel-body').css('display','none');
				$('.panel-heading span.clickable').addClass('panel-collapsed');
				$('.panel-heading span.clickable').find('i').removeClass('glyphicon-chevron-up').addClass('glyphicon-chevron-down');
            }
            else {
				$('.panel .panel-body').css('display','block');
				$('.panel-heading span.clickable').removeClass('panel-collapsed');
				$('.panel-heading span.clickable').find('i').removeClass('glyphicon-chevron-down').addClass('glyphicon-chevron-up');
            }
        });	
        
		
		$('#lstExistingAPIDailyActivityReportUserNos').multiselect({
		   buttonTitle: function(options, select) {
			    var selected = '';
			    options.each(function () {
			      selected += $(this).text() + ', ';
			    });
			    return selected.substr(0, selected.length - 2);
			  },
			buttonClass: 'btn btn-primary',
			buttonWidth: '425px',
			maxHeight: 400,
			dropRight:true,
			enableFiltering:true,
			filterPlaceholder:'Search',
			enableCaseInsensitiveFiltering:true,
			// possible options: 'text', 'value', 'both'
			filterBehavior:'text',
			includeFilterClearBtn:true,
			nonSelectedText:'No Users Selected For Daily Activity Report',
			numberDisplayed: 20,
		    onChange: function() {
		        var selected = this.$select.val();
		        $("#lstSelectedAPIDailyActivityReportUserNos").val(selected);
		        console.log(selected);
		        // ...
		    }
    			
	    });	
	    
		//*************************************************************************************************
		//Load the bootstrap multiselect box with the current daily activity report users preselected
		//*************************************************************************************************
		var data= $("#lstSelectedAPIDailyActivityReportUserNos").val();
		
		if (data) {
			//Make an array
			var dataarray=data.split(",");
			// Set the value
			$("#lstExistingAPIDailyActivityReportUserNos").val(dataarray);
			// Then refresh
			$("#lstExistingAPIDailyActivityReportUserNos").multiselect("refresh");
		}
		//*************************************************************************************************
        
        
	
		
	});
</script>

<style type="text/css">


	.content-element{
	  margin:50px 0 0 50px;
	}
	.circles-list ol {
	  list-style-type: none;
	  margin-left: 1.25em;
	  padding-left: 2.5em;
	  counter-reset: li-counter;
	  border-left: 1px solid #3c763d;
	  position: relative; }
	
	.circles-list ol > li {
	  position: relative;
	  margin-bottom: 3.125em;
	  clear: both; }
	
	.circles-list ol > li:before {
	  position: absolute;
	  top: -0.5em;
	  font-family: "Open Sans", sans-serif;
	  font-weight: 600;
	  font-size: 1em;
	  left: -3.75em;
	  width: 2.25em;
	  height: 2.25em;
	  line-height: 2.25em;
	  text-align: center;
	  z-index: 9;
	  color: #3c763d;
	  border: 2px solid #3c763d;
	  border-radius: 50%;
	  content: counter(li-counter);
	  background-color: #DFF0D8;
	  counter-increment: li-counter; }
	  	
	.row .panel-row{
	    margin-top:40px;
	    padding: 0 10px;
	}
	
	.clickable{
	    cursor: pointer;   
	}
	
	.panel-heading span {
		margin-top: -20px;
		font-size: 15px;
	}

	.container {
		margin-bottom: 20px;
		margin-top: 20px;
		margin-left:0px;
		width: 100%;
	}

	.container .row {
		margin-bottom: 20px;
		/*margin-top: 20px;*/
	}

	.full-spectrum .sp-palette {
		max-width: 200px;
	}
	
	.tab-colors-box{
		padding:15px;
		border:2px solid #000;
		margin:0px 0px 15px 0px;
		width:100%;
		display:block;
		float:left;
	}
	
	.tab-colors-title strong{
		width:100%;
		text-align:center;
		display:block;
	}
	
	.tab-colors-title .row{
		margin-bottom:0px;
	}
	
	.line-full{
	 	margin-bottom:20px;
	}
	
	.multi-select{
		min-height:200px;
		min-width:180px;
	}
	
	.custom-select{
		width: auto !important;
		display:inline-block;
	}

	.multi-select-dispatch{
	  min-height: 160px;
	  margin-top: 20px;
	 }

	
	.select-large{
		min-width:40% !important;
	}
	
	.ui-timepicker-table td a{
		padding: 3px;
		width:auto;
		text-align: left;
		font-size: 11px;
	}	
	
	.ui-timepicker-table .ui-timepicker-title{
		font-size: 13px;
	}
	
	.ui-timepicker-table th.periods{
		font-size: 13px;
	}
	
	.ui-widget-header{
		background: #193048;
		border: 1px solid #193048;
	}
	
	#PleaseWaitPanel{
		position: fixed;
		left: 470px;
		top: 275px;
		width: 975px;
		height: 300px;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	}   

	.btn-huge{
	    padding: 18px 28px;
	    font-size: 22px;	    
	}
</style>


<%
	SQL = "SELECT * FROM Settings_Global"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		OrderAPIRepostMode = rs("OrderAPIRepostMode")	
		OrderAPIOffsetDays = rs("OrderAPIOffsetDays")	
		InvoiceAPIRepostMode = rs("InvoiceAPIRepostMode")	
		InvoiceAPIOffsetDays = rs("InvoiceAPIOffsetDays")	
		SendInvoiceType = rs("SendInvoiceType")	
		RAAPIRepostMode = rs("RAAPIRepostMode")	
		RAAPIOffsetDays = rs("RAAPIOffsetDays")	
		CMAPIRepostMode = rs("CMAPIRepostMode")	
		CMAPIOffsetDays = rs("CMAPIOffsetDays")	
		SumInvAPIRepostMode = rs("SumInvAPIRepostMode")	
		SumInvAPIOffsetDays = rs("SumInvAPIOffsetDays")	
		OrderAPIRepostURL = rs("OrderAPIRepostURL")	
		OrderAPIRepostONOFF = rs("OrderAPIRepostONOFF")	
		InvoiceAPIRepostURL = rs("InvoiceAPIRepostURL")	
		InvoiceAPIRepostONOFF = rs("InvoiceAPIRepostONOFF")	
		RAAPIRepostURL = rs("RAAPIRepostURL")	
		RAAPIRepostONOFF = rs("RAAPIRepostONOFF")	
		CMAPIRepostURL = rs("CMAPIRepostURL")	
		CMAPIRepostONOFF = rs("CMAPIRepostONOFF")	
		SumInvAPIRepostURL = rs("SumInvAPIRepostURL")	
		SumInvAPIRepostONOFF = rs("SumInvAPIRepostONOFF")	
		APIDailyActivityReportOnOff	= rs("APIDailyActivityReportOnOff")			
		APIDailyActivityReportAdditionalEmails = rs("APIDailyActivityReportAdditionalEmails")	
		APIDailyActivityReportEmailSubject = rs("APIDailyActivityReportEmailSubject")	
		APIDailyActivityReportUserNos = rs("APIDailyActivityReportUserNos")	
		OrderCutoffTime = rs("OrderCutoffTime")
		InvoiceCutoffTime = rs("InvoiceCutoffTime")
		RACutoffTime = rs("RACutoffTime")
		CMCutoffTime = rs("CMCutoffTime")	
		OrderAPISwapAddressLines = rs("OrderAPISwapAddressLines")
						
	End If
	
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
%>


<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;Order API 
	<button id="toggle" class="btn btn-small btn-success"><i class="fas fa-arrows-v"></i>&nbsp;EXPAND/COLLAPSE ALL SETTINGS</button>
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
	<a href="<%= BaseURL %>admin/global/tiles/api/main.asp"><button class="btn btn-small btn-secondary pull-right"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fa fa-external-link"></i>&nbsp;API MAIN</button></a>
</h1>

<% If MUV_Read("orderAPIModuleOn")  = "Disabled" Then %>
	<div class="col-lg-6">
		<br><br>
		Please contact support if you would like to activate the Order API module.
	</div>
<% ElseIf MUV_Read("orderAPIModuleOn")  = "Enabled" Then  %>

	<form method="post" action="order-api-submit.asp" name="frmOrderAPI" id="frmOrderAPI">
	
	<div class="container">
		
		<%
			Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
			Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
			Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
			Response.Write("</div>")
			Response.Flush()
		%>
	
		<div class="row">
			<h3><i class="fad fa-sliders-h"></i>&nbsp;Order API Master Settings</h3>
		</div>
	
	    <div class="row">
	    
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">ORDERS RePost Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
					
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
	                	
	               			<div class="col-lg-6">
	               			<%
								If OrderAPIRepostONOFF = 0 Then
									Response.Write("<input type='checkbox' class='check' id='chkOrderAPIRepostONOFF' name='chkOrderAPIRepostONOFF'")
								Else
									Response.Write("<input type='checkbox' class='check' id='chkOrderAPIRepostONOFF' name='chkOrderAPIRepostONOFF' checked")
								End If
								Response.Write("> Check to turn on")
							%>
	               			</div>
	               			
							<div class="col-lg-6">
								Mode:<br>
								<select class="form-control pull-left" name="selOrderAPIRepostMode">
									<option value="TEST" <% If OrderAPIRepostMode = "TEST" Then Response.Write("selected") %>>TEST</option>
									<option value="LIVE" <% If OrderAPIRepostMode = "LIVE" Then Response.Write("selected") %>>LIVE</option>
								</select>	
							</div>
	               			
						</div>
					
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:10px;">
	               			<div class="col-lg-10">Re-post received orders to URL</div>
						</div>
	                 	<!-- eof line !-->
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
	               			<div class="col-lg-12"><input type="text"class="form-control" style="width:100%;" name="txtOrderAPIRepostURL" id="txtOrderAPIRepostURL" value="<%= OrderAPIRepostURL %>"></div>
						</div>


						<!-- order cutoff time -->
						<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
							<!-- select time -->
							<div class="col-lg-6">
								Cutoff Time:
								<select class="form-control" name="selOrderCutoffTime" id="selOrderCutoffTime">
									<option value="0000"<%If OrderCutoffTime = "0000" Then Response.Write(" selected ")%>>-Midnight-</option>
									<option value="0015"<%If OrderCutoffTime = "0015" Then Response.Write(" selected ")%>>12:15 AM</option>
									<option value="0030"<%If OrderCutoffTime = "0030" Then Response.Write(" selected ")%>>12:30 AM</option>
									<option value="0045"<%If OrderCutoffTime = "0045" Then Response.Write(" selected ")%>>12:45 AM</option>
									<option value="100"<%If OrderCutoffTime = "100" Then Response.Write(" selected ")%>>1:00 AM</option>
									<option value="115"<%If OrderCutoffTime = "115" Then Response.Write(" selected ")%>>1:15 AM</option>
									<option value="130"<%If OrderCutoffTime = "130" Then Response.Write(" selected ")%>>1:30 AM</option>
									<option value="145"<%If OrderCutoffTime = "145" Then Response.Write(" selected ")%>>1:45 AM</option>
									<option value="200"<%If OrderCutoffTime = "200" Then Response.Write(" selected ")%>>2:00 AM</option>
									<option value="215"<%If OrderCutoffTime = "215" Then Response.Write(" selected ")%>>2:15 AM</option>
									<option value="230"<%If OrderCutoffTime = "230" Then Response.Write(" selected ")%>>2:30 AM</option>
									<option value="245"<%If OrderCutoffTime = "245" Then Response.Write(" selected ")%>>2:45 AM</option>
									<option value="300"<%If OrderCutoffTime = "300" Then Response.Write(" selected ")%>>3:00 AM</option>
									<option value="315"<%If OrderCutoffTime = "315" Then Response.Write(" selected ")%>>3:15 AM</option>
									<option value="330"<%If OrderCutoffTime = "330" Then Response.Write(" selected ")%>>3:30 AM</option>
									<option value="345"<%If OrderCutoffTime = "345" Then Response.Write(" selected ")%>>3:45 AM</option>
									<option value="400"<%If OrderCutoffTime = "400" Then Response.Write(" selected ")%>>4:00 AM</option>
									<option value="415"<%If OrderCutoffTime = "415" Then Response.Write(" selected ")%>>4:15 AM</option>
									<option value="430"<%If OrderCutoffTime = "430" Then Response.Write(" selected ")%>>4:30 AM</option>
									<option value="445"<%If OrderCutoffTime = "445" Then Response.Write(" selected ")%>>4:45 AM</option>
									<option value="500"<%If OrderCutoffTime = "500" Then Response.Write(" selected ")%>>5:00 AM</option>
									<option value="515"<%If OrderCutoffTime = "515" Then Response.Write(" selected ")%>>5:15 AM</option>
									<option value="530"<%If OrderCutoffTime = "530" Then Response.Write(" selected ")%>>5:30 AM</option>
									<option value="545"<%If OrderCutoffTime = "545" Then Response.Write(" selected ")%>>5:45 AM</option>
									<option value="600"<%If OrderCutoffTime = "600" Then Response.Write(" selected ")%>>6:00 AM</option>
									<option value="615"<%If OrderCutoffTime = "615" Then Response.Write(" selected ")%>>6:15 AM</option>
									<option value="630"<%If OrderCutoffTime = "630" Then Response.Write(" selected ")%>>6:30 AM</option>
									<option value="645"<%If OrderCutoffTime = "645" Then Response.Write(" selected ")%>>6:45 AM</option>
									<option value="700"<%If OrderCutoffTime = "700" Then Response.Write(" selected ")%>>7:00 AM</option>
									<option value="715"<%If OrderCutoffTime = "715" Then Response.Write(" selected ")%>>7:15 AM</option>
									<option value="730"<%If OrderCutoffTime = "730" Then Response.Write(" selected ")%>>7:30 AM</option>
									<option value="745"<%If OrderCutoffTime = "745" Then Response.Write(" selected ")%>>7:45 AM</option>
									<option value="800"<%If OrderCutoffTime = "800" Then Response.Write(" selected ")%>>8:00 AM</option>
									<option value="815"<%If OrderCutoffTime = "815" Then Response.Write(" selected ")%>>8:15 AM</option>
									<option value="830"<%If OrderCutoffTime = "830" Then Response.Write(" selected ")%>>8:30 AM</option>
									<option value="845"<%If OrderCutoffTime = "845" Then Response.Write(" selected ")%>>8:45 AM</option>
									<option value="900"<%If OrderCutoffTime = "900" Then Response.Write(" selected ")%>>9:00 AM</option>
									<option value="915"<%If OrderCutoffTime = "915" Then Response.Write(" selected ")%>>9:15 AM</option>
									<option value="930"<%If OrderCutoffTime = "930" Then Response.Write(" selected ")%>>9:30 AM</option>
									<option value="945"<%If OrderCutoffTime = "945" Then Response.Write(" selected ")%>>9:45 AM</option>
									<option value="1000"<%If OrderCutoffTime = "1000" Then Response.Write(" selected ")%>>10:00 AM</option>
									<option value="1015"<%If OrderCutoffTime = "1015" Then Response.Write(" selected ")%>>10:15 AM</option>
									<option value="1030"<%If OrderCutoffTime = "1030" Then Response.Write(" selected ")%>>10:30 AM</option>
									<option value="1045"<%If OrderCutoffTime = "1045" Then Response.Write(" selected ")%>>10:45 AM</option>
									<option value="1100"<%If OrderCutoffTime = "1100" Then Response.Write(" selected ")%>>11:00 AM</option>
									<option value="1115"<%If OrderCutoffTime = "1115" Then Response.Write(" selected ")%>>11:15 AM</option>
									<option value="1130"<%If OrderCutoffTime = "1130" Then Response.Write(" selected ")%>>11:30 AM</option>
									<option value="1145"<%If OrderCutoffTime = "1145" Then Response.Write(" selected ")%>>11:45 AM</option>
									<option value="1200"<%If OrderCutoffTime = "1200" Then Response.Write(" selected ")%>>-Noon-</option>
									<option value="1215"<%If OrderCutoffTime = "1215" Then Response.Write(" selected ")%>>12:15 PM</option>
									<option value="1230"<%If OrderCutoffTime = "1230" Then Response.Write(" selected ")%>>12:30 PM</option>
									<option value="1245"<%If OrderCutoffTime = "1245" Then Response.Write(" selected ")%>>12:45 PM</option>
									<option value="1300"<%If OrderCutoffTime = "1300" Then Response.Write(" selected ")%>>1:00 PM</option>
									<option value="1315"<%If OrderCutoffTime = "1315" Then Response.Write(" selected ")%>>1:15 PM</option>
									<option value="1330"<%If OrderCutoffTime = "1330" Then Response.Write(" selected ")%>>1:30 PM</option>
									<option value="1345"<%If OrderCutoffTime = "1345" Then Response.Write(" selected ")%>>1:45 PM</option>
									<option value="1400"<%If OrderCutoffTime = "1400" Then Response.Write(" selected ")%>>2:00 PM</option>
									<option value="1415"<%If OrderCutoffTime = "1415" Then Response.Write(" selected ")%>>2:15 PM</option>
									<option value="1430"<%If OrderCutoffTime = "1430" Then Response.Write(" selected ")%>>2:30 PM</option>
									<option value="1445"<%If OrderCutoffTime = "1445" Then Response.Write(" selected ")%>>2:45 PM</option>
									<option value="1500"<%If OrderCutoffTime = "1500" Then Response.Write(" selected ")%>>3:00 PM</option>
									<option value="1515"<%If OrderCutoffTime = "1515" Then Response.Write(" selected ")%>>3:15 PM</option>
									<option value="1530"<%If OrderCutoffTime = "1530" Then Response.Write(" selected ")%>>3:30 PM</option>
									<option value="1545"<%If OrderCutoffTime = "1545" Then Response.Write(" selected ")%>>3:45 PM</option>
									<option value="1600"<%If OrderCutoffTime = "1600" Then Response.Write(" selected ")%>>4:00 PM</option>
									<option value="1615"<%If OrderCutoffTime = "1615" Then Response.Write(" selected ")%>>4:15 PM</option>
									<option value="1630"<%If OrderCutoffTime = "1630" Then Response.Write(" selected ")%>>4:30 PM</option>
									<option value="1645"<%If OrderCutoffTime = "1645" Then Response.Write(" selected ")%>>4:45 PM</option>
									<option value="1700"<%If OrderCutoffTime = "1700" Then Response.Write(" selected ")%>>5:00 PM</option>
									<option value="1715"<%If OrderCutoffTime = "1715" Then Response.Write(" selected ")%>>5:15 PM</option>
									<option value="1730"<%If OrderCutoffTime = "1730" Then Response.Write(" selected ")%>>5:30 PM</option>
									<option value="1745"<%If OrderCutoffTime = "1745" Then Response.Write(" selected ")%>>5:45 PM</option>
									<option value="1800"<%If OrderCutoffTime = "1800" Then Response.Write(" selected ")%>>6:00 PM</option>
									<option value="1815"<%If OrderCutoffTime = "1815" Then Response.Write(" selected ")%>>6:15 PM</option>
									<option value="1830"<%If OrderCutoffTime = "1830" Then Response.Write(" selected ")%>>6:30 PM</option>
									<option value="1845"<%If OrderCutoffTime = "1845" Then Response.Write(" selected ")%>>6:45 PM</option>
									<option value="1900"<%If OrderCutoffTime = "1900" Then Response.Write(" selected ")%>>7:00 PM</option>
									<option value="1915"<%If OrderCutoffTime = "1915" Then Response.Write(" selected ")%>>7:15 PM</option>
									<option value="1930"<%If OrderCutoffTime = "1930" Then Response.Write(" selected ")%>>7:30 PM</option>
									<option value="1945"<%If OrderCutoffTime = "1945" Then Response.Write(" selected ")%>>7:45 PM</option>
									<option value="2000"<%If OrderCutoffTime = "2000" Then Response.Write(" selected ")%>>8:00 PM</option>
									<option value="2015"<%If OrderCutoffTime = "2015" Then Response.Write(" selected ")%>>8:15 PM</option>
									<option value="2030"<%If OrderCutoffTime = "2030" Then Response.Write(" selected ")%>>8:30 PM</option>
									<option value="2045"<%If OrderCutoffTime = "2045" Then Response.Write(" selected ")%>>8:45 PM</option>
									<option value="2100"<%If OrderCutoffTime = "2100" Then Response.Write(" selected ")%>>9:00 PM</option>
									<option value="2115"<%If OrderCutoffTime = "2115" Then Response.Write(" selected ")%>>9:15 PM</option>
									<option value="2130"<%If OrderCutoffTime = "2130" Then Response.Write(" selected ")%>>9:30 PM</option>
									<option value="2145"<%If OrderCutoffTime = "2145" Then Response.Write(" selected ")%>>9:45 PM</option>
									<option value="2200"<%If OrderCutoffTime = "2200" Then Response.Write(" selected ")%>>10:00 PM</option>
									<option value="2215"<%If OrderCutoffTime = "2215" Then Response.Write(" selected ")%>>10:15 PM</option>
									<option value="2230"<%If OrderCutoffTime = "2230" Then Response.Write(" selected ")%>>10:30 PM</option>
									<option value="2245"<%If OrderCutoffTime = "2245" Then Response.Write(" selected ")%>>10:45 PM</option>
									<option value="2300"<%If OrderCutoffTime = "2300" Then Response.Write(" selected ")%>>11:00 PM</option>
									<option value="2315"<%If OrderCutoffTime = "2315" Then Response.Write(" selected ")%>>11:15 PM</option>
									<option value="2330"<%If OrderCutoffTime = "2330" Then Response.Write(" selected ")%>>11:30 PM</option>
									<option value="2345"<%If OrderCutoffTime = "2345" Then Response.Write(" selected ")%>>11:45 PM</option>	
								</select>
							</div>
							<!-- eof order cutoff time -->
							
							<div class="col-lg-6 pull-left">
								<br><p><em>Posts received after the cutoff time will have their delivery date increased by one business day.</em></p>
							</div>
							<!-- eof excerpt -->
						</div>
						<!-- eof cutoff time -->
							
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
	                			<div class="col-lg-12">
								<%
									If OrderAPISwapAddressLines = 0 Then
										Response.Write("<input type='checkbox' class='check' id='chkOrderAPISwapAddressLines' name='chkOrderAPISwapAddressLines'")
									Else
										Response.Write("<input type='checkbox' class='check' id='chkOrderAPISwapAddressLines' name='chkOrderAPISwapAddressLines' checked")
									End If
									Response.Write("> Swap address line 1 & 2")
								%>
								</div>
						</div>
						<!-- eof address swap lines -->
							
											
					</div>
				</div>
			</div>
			
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">INVOICE RePost Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					

					
		              		<!-- line !-->
		                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
		                	
		               			<div class="col-lg-6">
		               			<%
									If InvoiceAPIRepostONOFF = 0 Then
										Response.Write("<input type='checkbox' class='check' id='chkInvoiceAPIRepostONOFF' name='chkInvoiceAPIRepostONOFF'")
									Else
										Response.Write("<input type='checkbox' class='check' id='chkInvoiceAPIRepostONOFF' name='chkInvoiceAPIRepostONOFF' checked")
									End If
									Response.Write("> Check to turn on")
								%>
		               			</div>
		               			
								<div class="col-lg-6">
									Mode:<br>
									<select class="form-control pull-left" name="selInvoiceAPIRepostMode">
										<option value="TEST" <% If InvoiceAPIRepostMode = "TEST" Then Response.Write("selected") %>>TEST</option>
										<option value="LIVE" <% If InvoiceAPIRepostMode = "LIVE" Then Response.Write("selected") %>>LIVE</option>
									</select>	
								</div>
		               			
							</div>

							
							<div class="row schedule-info" style="margin-top:0px;margin-bottom:10px;">
								<div class="col-lg-10"><label for="txtInvoicesAPIRepostURL" class="post-labels">Re-post received invoices to URL</label></div>
							</div>
							

							<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
								<div class="col-lg-12"><input type="text" class="form-control" name="txtInvoiceAPIRepostURL" id="txtInvoiceAPIRepostURL" value="<%= InvoiceAPIRepostURL %>"></div>
							</div>
							

							<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
	                			<div class="col-lg-12">
	                				Send Invoices As:
									<select class="form-control" name="selSendInvoiceType" style="max-width:185px;">
										<option value="POSTED" <% If SendInvoiceType = "POSTED" Then Response.Write("selected") %>>POSTED</option>
										<option value="UNPOSTED" <% If SendInvoiceType = "UNPOSTED" Then Response.Write("selected") %>>UNPOSTED</option>
									</select>
									
								</div>
							</div>

	
	
							<!-- order cutoff time -->
							<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
								<!-- select time -->
								<div class="col-lg-6">
									Cutoff Time:
									<select class="form-control" name="selInvoiceCutoffTime" id="selInvoiceCutoffTime">
										<option value="0000"<%If InvoiceCutoffTime = "0000" Then Response.Write(" selected ")%>>-Midnight-</option>
										<option value="0015"<%If InvoiceCutoffTime = "0015" Then Response.Write(" selected ")%>>12:15 AM</option>
										<option value="0030"<%If InvoiceCutoffTime = "0030" Then Response.Write(" selected ")%>>12:30 AM</option>
										<option value="0045"<%If InvoiceCutoffTime = "0045" Then Response.Write(" selected ")%>>12:45 AM</option>
										<option value="100"<%If InvoiceCutoffTime = "100" Then Response.Write(" selected ")%>>1:00 AM</option>
										<option value="115"<%If InvoiceCutoffTime = "115" Then Response.Write(" selected ")%>>1:15 AM</option>
										<option value="130"<%If InvoiceCutoffTime = "130" Then Response.Write(" selected ")%>>1:30 AM</option>
										<option value="145"<%If InvoiceCutoffTime = "145" Then Response.Write(" selected ")%>>1:45 AM</option>
										<option value="200"<%If InvoiceCutoffTime = "200" Then Response.Write(" selected ")%>>2:00 AM</option>
										<option value="215"<%If InvoiceCutoffTime = "215" Then Response.Write(" selected ")%>>2:15 AM</option>
										<option value="230"<%If InvoiceCutoffTime = "230" Then Response.Write(" selected ")%>>2:30 AM</option>
										<option value="245"<%If InvoiceCutoffTime = "245" Then Response.Write(" selected ")%>>2:45 AM</option>
										<option value="300"<%If InvoiceCutoffTime = "300" Then Response.Write(" selected ")%>>3:00 AM</option>
										<option value="315"<%If InvoiceCutoffTime = "315" Then Response.Write(" selected ")%>>3:15 AM</option>
										<option value="330"<%If InvoiceCutoffTime = "330" Then Response.Write(" selected ")%>>3:30 AM</option>
										<option value="345"<%If InvoiceCutoffTime = "345" Then Response.Write(" selected ")%>>3:45 AM</option>
										<option value="400"<%If InvoiceCutoffTime = "400" Then Response.Write(" selected ")%>>4:00 AM</option>
										<option value="415"<%If InvoiceCutoffTime = "415" Then Response.Write(" selected ")%>>4:15 AM</option>
										<option value="430"<%If InvoiceCutoffTime = "430" Then Response.Write(" selected ")%>>4:30 AM</option>
										<option value="445"<%If InvoiceCutoffTime = "445" Then Response.Write(" selected ")%>>4:45 AM</option>
										<option value="500"<%If InvoiceCutoffTime = "500" Then Response.Write(" selected ")%>>5:00 AM</option>
										<option value="515"<%If InvoiceCutoffTime = "515" Then Response.Write(" selected ")%>>5:15 AM</option>
										<option value="530"<%If InvoiceCutoffTime = "530" Then Response.Write(" selected ")%>>5:30 AM</option>
										<option value="545"<%If InvoiceCutoffTime = "545" Then Response.Write(" selected ")%>>5:45 AM</option>
										<option value="600"<%If InvoiceCutoffTime = "600" Then Response.Write(" selected ")%>>6:00 AM</option>
										<option value="615"<%If InvoiceCutoffTime = "615" Then Response.Write(" selected ")%>>6:15 AM</option>
										<option value="630"<%If InvoiceCutoffTime = "630" Then Response.Write(" selected ")%>>6:30 AM</option>
										<option value="645"<%If InvoiceCutoffTime = "645" Then Response.Write(" selected ")%>>6:45 AM</option>
										<option value="700"<%If InvoiceCutoffTime = "700" Then Response.Write(" selected ")%>>7:00 AM</option>
										<option value="715"<%If InvoiceCutoffTime = "715" Then Response.Write(" selected ")%>>7:15 AM</option>
										<option value="730"<%If InvoiceCutoffTime = "730" Then Response.Write(" selected ")%>>7:30 AM</option>
										<option value="745"<%If InvoiceCutoffTime = "745" Then Response.Write(" selected ")%>>7:45 AM</option>
										<option value="800"<%If InvoiceCutoffTime = "800" Then Response.Write(" selected ")%>>8:00 AM</option>
										<option value="815"<%If InvoiceCutoffTime = "815" Then Response.Write(" selected ")%>>8:15 AM</option>
										<option value="830"<%If InvoiceCutoffTime = "830" Then Response.Write(" selected ")%>>8:30 AM</option>
										<option value="845"<%If InvoiceCutoffTime = "845" Then Response.Write(" selected ")%>>8:45 AM</option>
										<option value="900"<%If InvoiceCutoffTime = "900" Then Response.Write(" selected ")%>>9:00 AM</option>
										<option value="915"<%If InvoiceCutoffTime = "915" Then Response.Write(" selected ")%>>9:15 AM</option>
										<option value="930"<%If InvoiceCutoffTime = "930" Then Response.Write(" selected ")%>>9:30 AM</option>
										<option value="945"<%If InvoiceCutoffTime = "945" Then Response.Write(" selected ")%>>9:45 AM</option>
										<option value="1000"<%If InvoiceCutoffTime = "1000" Then Response.Write(" selected ")%>>10:00 AM</option>
										<option value="1015"<%If InvoiceCutoffTime = "1015" Then Response.Write(" selected ")%>>10:15 AM</option>
										<option value="1030"<%If InvoiceCutoffTime = "1030" Then Response.Write(" selected ")%>>10:30 AM</option>
										<option value="1045"<%If InvoiceCutoffTime = "1045" Then Response.Write(" selected ")%>>10:45 AM</option>
										<option value="1100"<%If InvoiceCutoffTime = "1100" Then Response.Write(" selected ")%>>11:00 AM</option>
										<option value="1115"<%If InvoiceCutoffTime = "1115" Then Response.Write(" selected ")%>>11:15 AM</option>
										<option value="1130"<%If InvoiceCutoffTime = "1130" Then Response.Write(" selected ")%>>11:30 AM</option>
										<option value="1145"<%If InvoiceCutoffTime = "1145" Then Response.Write(" selected ")%>>11:45 AM</option>
										<option value="1200"<%If InvoiceCutoffTime = "1200" Then Response.Write(" selected ")%>>-Noon-</option>
										<option value="1215"<%If InvoiceCutoffTime = "1215" Then Response.Write(" selected ")%>>12:15 PM</option>
										<option value="1230"<%If InvoiceCutoffTime = "1230" Then Response.Write(" selected ")%>>12:30 PM</option>
										<option value="1245"<%If InvoiceCutoffTime = "1245" Then Response.Write(" selected ")%>>12:45 PM</option>
										<option value="1300"<%If InvoiceCutoffTime = "1300" Then Response.Write(" selected ")%>>1:00 PM</option>
										<option value="1315"<%If InvoiceCutoffTime = "1315" Then Response.Write(" selected ")%>>1:15 PM</option>
										<option value="1330"<%If InvoiceCutoffTime = "1330" Then Response.Write(" selected ")%>>1:30 PM</option>
										<option value="1345"<%If InvoiceCutoffTime = "1345" Then Response.Write(" selected ")%>>1:45 PM</option>
										<option value="1400"<%If InvoiceCutoffTime = "1400" Then Response.Write(" selected ")%>>2:00 PM</option>
										<option value="1415"<%If InvoiceCutoffTime = "1415" Then Response.Write(" selected ")%>>2:15 PM</option>
										<option value="1430"<%If InvoiceCutoffTime = "1430" Then Response.Write(" selected ")%>>2:30 PM</option>
										<option value="1445"<%If InvoiceCutoffTime = "1445" Then Response.Write(" selected ")%>>2:45 PM</option>
										<option value="1500"<%If InvoiceCutoffTime = "1500" Then Response.Write(" selected ")%>>3:00 PM</option>
										<option value="1515"<%If InvoiceCutoffTime = "1515" Then Response.Write(" selected ")%>>3:15 PM</option>
										<option value="1530"<%If InvoiceCutoffTime = "1530" Then Response.Write(" selected ")%>>3:30 PM</option>
										<option value="1545"<%If InvoiceCutoffTime = "1545" Then Response.Write(" selected ")%>>3:45 PM</option>
										<option value="1600"<%If InvoiceCutoffTime = "1600" Then Response.Write(" selected ")%>>4:00 PM</option>
										<option value="1615"<%If InvoiceCutoffTime = "1615" Then Response.Write(" selected ")%>>4:15 PM</option>
										<option value="1630"<%If InvoiceCutoffTime = "1630" Then Response.Write(" selected ")%>>4:30 PM</option>
										<option value="1645"<%If InvoiceCutoffTime = "1645" Then Response.Write(" selected ")%>>4:45 PM</option>
										<option value="1700"<%If InvoiceCutoffTime = "1700" Then Response.Write(" selected ")%>>5:00 PM</option>
										<option value="1715"<%If InvoiceCutoffTime = "1715" Then Response.Write(" selected ")%>>5:15 PM</option>
										<option value="1730"<%If InvoiceCutoffTime = "1730" Then Response.Write(" selected ")%>>5:30 PM</option>
										<option value="1745"<%If InvoiceCutoffTime = "1745" Then Response.Write(" selected ")%>>5:45 PM</option>
										<option value="1800"<%If InvoiceCutoffTime = "1800" Then Response.Write(" selected ")%>>6:00 PM</option>
										<option value="1815"<%If InvoiceCutoffTime = "1815" Then Response.Write(" selected ")%>>6:15 PM</option>
										<option value="1830"<%If InvoiceCutoffTime = "1830" Then Response.Write(" selected ")%>>6:30 PM</option>
										<option value="1845"<%If InvoiceCutoffTime = "1845" Then Response.Write(" selected ")%>>6:45 PM</option>
										<option value="1900"<%If InvoiceCutoffTime = "1900" Then Response.Write(" selected ")%>>7:00 PM</option>
										<option value="1915"<%If InvoiceCutoffTime = "1915" Then Response.Write(" selected ")%>>7:15 PM</option>
										<option value="1930"<%If InvoiceCutoffTime = "1930" Then Response.Write(" selected ")%>>7:30 PM</option>
										<option value="1945"<%If InvoiceCutoffTime = "1945" Then Response.Write(" selected ")%>>7:45 PM</option>
										<option value="2000"<%If InvoiceCutoffTime = "2000" Then Response.Write(" selected ")%>>8:00 PM</option>
										<option value="2015"<%If InvoiceCutoffTime = "2015" Then Response.Write(" selected ")%>>8:15 PM</option>
										<option value="2030"<%If InvoiceCutoffTime = "2030" Then Response.Write(" selected ")%>>8:30 PM</option>
										<option value="2045"<%If InvoiceCutoffTime = "2045" Then Response.Write(" selected ")%>>8:45 PM</option>
										<option value="2100"<%If InvoiceCutoffTime = "2100" Then Response.Write(" selected ")%>>9:00 PM</option>
										<option value="2115"<%If InvoiceCutoffTime = "2115" Then Response.Write(" selected ")%>>9:15 PM</option>
										<option value="2130"<%If InvoiceCutoffTime = "2130" Then Response.Write(" selected ")%>>9:30 PM</option>
										<option value="2145"<%If InvoiceCutoffTime = "2145" Then Response.Write(" selected ")%>>9:45 PM</option>
										<option value="2200"<%If InvoiceCutoffTime = "2200" Then Response.Write(" selected ")%>>10:00 PM</option>
										<option value="2215"<%If InvoiceCutoffTime = "2215" Then Response.Write(" selected ")%>>10:15 PM</option>
										<option value="2230"<%If InvoiceCutoffTime = "2230" Then Response.Write(" selected ")%>>10:30 PM</option>
										<option value="2245"<%If InvoiceCutoffTime = "2245" Then Response.Write(" selected ")%>>10:45 PM</option>
										<option value="2300"<%If InvoiceCutoffTime = "2300" Then Response.Write(" selected ")%>>11:00 PM</option>
										<option value="2315"<%If InvoiceCutoffTime = "2315" Then Response.Write(" selected ")%>>11:15 PM</option>
										<option value="2330"<%If InvoiceCutoffTime = "2330" Then Response.Write(" selected ")%>>11:30 PM</option>
										<option value="2345"<%If InvoiceCutoffTime = "2345" Then Response.Write(" selected ")%>>11:45 PM</option>	
									</select>
								</div>
								<!-- eof order cutoff time -->
								
								<div class="col-lg-6 pull-left">
									<br><p><em>Posts received after the cutoff time will have their delivery date increased by one business day.</em></p>
								</div>
								<!-- eof excerpt -->
							</div>
							<!-- eof cutoff time -->
	
		
											
					</div>
				</div>
			</div>
	
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">RETURN AUTHORIZATION RePost Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
					
		              		<!-- line !-->
		                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
		                	
		               			<div class="col-lg-6">
		               			<%
									If RAAPIRepostONOFF = 0 Then
										Response.Write("<input type='checkbox' class='check' id='chkRAAPIRepostONOFF' name='chkRAAPIRepostONOFF'")
									Else
										Response.Write("<input type='checkbox' class='check' id='chkRAAPIRepostONOFF' name='chkRAAPIRepostONOFF' checked")
									End If
									Response.Write("> Check to turn on")
								%>
		               			</div>
		               			
								<div class="col-lg-6">
									Mode:<br>
									<select class="form-control pull-left" name="selRAAPIRepostMode">
										<option value="TEST" <% If RAAPIRepostMode = "TEST" Then Response.Write("selected") %>>TEST</option>
										<option value="LIVE" <% If RAAPIRepostMode = "LIVE" Then Response.Write("selected") %>>LIVE</option>
									</select>	
								</div>
		               			
							</div>
							
				         	<!-- POST Received invoices To URL-->
							<div class="row schedule-info" style="margin-top:0px;margin-bottom:10px;">
								<div class="col-lg-10"><label  for="txtRAsAPIRepostURL" class="post-labels">Re-post received return authorizations to URL</label></div>
							</div>
								
							<div class="row schedule-info">
								<div class="col-lg-12"><input type="text" class="form-control" name="txtRAAPIRepostURL" id="txtRAAPIRepostURL" value="<%= RAAPIRepostURL %>"></div>
							</div>
							
	
	
							<!-- order cutoff time -->
							<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
								<!-- select time -->
								<div class="col-lg-6">
									Cutoff Time:
									<select class="form-control" name="selRACutoffTime" id="selRACutoffTime">
										<option value="0000"<%If RACutoffTime = "0000" Then Response.Write(" selected ")%>>-Midnight-</option>
										<option value="0015"<%If RACutoffTime = "0015" Then Response.Write(" selected ")%>>12:15 AM</option>
										<option value="0030"<%If RACutoffTime = "0030" Then Response.Write(" selected ")%>>12:30 AM</option>
										<option value="0045"<%If RACutoffTime = "0045" Then Response.Write(" selected ")%>>12:45 AM</option>
										<option value="100"<%If RACutoffTime = "100" Then Response.Write(" selected ")%>>1:00 AM</option>
										<option value="115"<%If RACutoffTime = "115" Then Response.Write(" selected ")%>>1:15 AM</option>
										<option value="130"<%If RACutoffTime = "130" Then Response.Write(" selected ")%>>1:30 AM</option>
										<option value="145"<%If RACutoffTime = "145" Then Response.Write(" selected ")%>>1:45 AM</option>
										<option value="200"<%If RACutoffTime = "200" Then Response.Write(" selected ")%>>2:00 AM</option>
										<option value="215"<%If RACutoffTime = "215" Then Response.Write(" selected ")%>>2:15 AM</option>
										<option value="230"<%If RACutoffTime = "230" Then Response.Write(" selected ")%>>2:30 AM</option>
										<option value="245"<%If RACutoffTime = "245" Then Response.Write(" selected ")%>>2:45 AM</option>
										<option value="300"<%If RACutoffTime = "300" Then Response.Write(" selected ")%>>3:00 AM</option>
										<option value="315"<%If RACutoffTime = "315" Then Response.Write(" selected ")%>>3:15 AM</option>
										<option value="330"<%If RACutoffTime = "330" Then Response.Write(" selected ")%>>3:30 AM</option>
										<option value="345"<%If RACutoffTime = "345" Then Response.Write(" selected ")%>>3:45 AM</option>
										<option value="400"<%If RACutoffTime = "400" Then Response.Write(" selected ")%>>4:00 AM</option>
										<option value="415"<%If RACutoffTime = "415" Then Response.Write(" selected ")%>>4:15 AM</option>
										<option value="430"<%If RACutoffTime = "430" Then Response.Write(" selected ")%>>4:30 AM</option>
										<option value="445"<%If RACutoffTime = "445" Then Response.Write(" selected ")%>>4:45 AM</option>
										<option value="500"<%If RACutoffTime = "500" Then Response.Write(" selected ")%>>5:00 AM</option>
										<option value="515"<%If RACutoffTime = "515" Then Response.Write(" selected ")%>>5:15 AM</option>
										<option value="530"<%If RACutoffTime = "530" Then Response.Write(" selected ")%>>5:30 AM</option>
										<option value="545"<%If RACutoffTime = "545" Then Response.Write(" selected ")%>>5:45 AM</option>
										<option value="600"<%If RACutoffTime = "600" Then Response.Write(" selected ")%>>6:00 AM</option>
										<option value="615"<%If RACutoffTime = "615" Then Response.Write(" selected ")%>>6:15 AM</option>
										<option value="630"<%If RACutoffTime = "630" Then Response.Write(" selected ")%>>6:30 AM</option>
										<option value="645"<%If RACutoffTime = "645" Then Response.Write(" selected ")%>>6:45 AM</option>
										<option value="700"<%If RACutoffTime = "700" Then Response.Write(" selected ")%>>7:00 AM</option>
										<option value="715"<%If RACutoffTime = "715" Then Response.Write(" selected ")%>>7:15 AM</option>
										<option value="730"<%If RACutoffTime = "730" Then Response.Write(" selected ")%>>7:30 AM</option>
										<option value="745"<%If RACutoffTime = "745" Then Response.Write(" selected ")%>>7:45 AM</option>
										<option value="800"<%If RACutoffTime = "800" Then Response.Write(" selected ")%>>8:00 AM</option>
										<option value="815"<%If RACutoffTime = "815" Then Response.Write(" selected ")%>>8:15 AM</option>
										<option value="830"<%If RACutoffTime = "830" Then Response.Write(" selected ")%>>8:30 AM</option>
										<option value="845"<%If RACutoffTime = "845" Then Response.Write(" selected ")%>>8:45 AM</option>
										<option value="900"<%If RACutoffTime = "900" Then Response.Write(" selected ")%>>9:00 AM</option>
										<option value="915"<%If RACutoffTime = "915" Then Response.Write(" selected ")%>>9:15 AM</option>
										<option value="930"<%If RACutoffTime = "930" Then Response.Write(" selected ")%>>9:30 AM</option>
										<option value="945"<%If RACutoffTime = "945" Then Response.Write(" selected ")%>>9:45 AM</option>
										<option value="1000"<%If RACutoffTime = "1000" Then Response.Write(" selected ")%>>10:00 AM</option>
										<option value="1015"<%If RACutoffTime = "1015" Then Response.Write(" selected ")%>>10:15 AM</option>
										<option value="1030"<%If RACutoffTime = "1030" Then Response.Write(" selected ")%>>10:30 AM</option>
										<option value="1045"<%If RACutoffTime = "1045" Then Response.Write(" selected ")%>>10:45 AM</option>
										<option value="1100"<%If RACutoffTime = "1100" Then Response.Write(" selected ")%>>11:00 AM</option>
										<option value="1115"<%If RACutoffTime = "1115" Then Response.Write(" selected ")%>>11:15 AM</option>
										<option value="1130"<%If RACutoffTime = "1130" Then Response.Write(" selected ")%>>11:30 AM</option>
										<option value="1145"<%If RACutoffTime = "1145" Then Response.Write(" selected ")%>>11:45 AM</option>
										<option value="1200"<%If RACutoffTime = "1200" Then Response.Write(" selected ")%>>-Noon-</option>
										<option value="1215"<%If RACutoffTime = "1215" Then Response.Write(" selected ")%>>12:15 PM</option>
										<option value="1230"<%If RACutoffTime = "1230" Then Response.Write(" selected ")%>>12:30 PM</option>
										<option value="1245"<%If RACutoffTime = "1245" Then Response.Write(" selected ")%>>12:45 PM</option>
										<option value="1300"<%If RACutoffTime = "1300" Then Response.Write(" selected ")%>>1:00 PM</option>
										<option value="1315"<%If RACutoffTime = "1315" Then Response.Write(" selected ")%>>1:15 PM</option>
										<option value="1330"<%If RACutoffTime = "1330" Then Response.Write(" selected ")%>>1:30 PM</option>
										<option value="1345"<%If RACutoffTime = "1345" Then Response.Write(" selected ")%>>1:45 PM</option>
										<option value="1400"<%If RACutoffTime = "1400" Then Response.Write(" selected ")%>>2:00 PM</option>
										<option value="1415"<%If RACutoffTime = "1415" Then Response.Write(" selected ")%>>2:15 PM</option>
										<option value="1430"<%If RACutoffTime = "1430" Then Response.Write(" selected ")%>>2:30 PM</option>
										<option value="1445"<%If RACutoffTime = "1445" Then Response.Write(" selected ")%>>2:45 PM</option>
										<option value="1500"<%If RACutoffTime = "1500" Then Response.Write(" selected ")%>>3:00 PM</option>
										<option value="1515"<%If RACutoffTime = "1515" Then Response.Write(" selected ")%>>3:15 PM</option>
										<option value="1530"<%If RACutoffTime = "1530" Then Response.Write(" selected ")%>>3:30 PM</option>
										<option value="1545"<%If RACutoffTime = "1545" Then Response.Write(" selected ")%>>3:45 PM</option>
										<option value="1600"<%If RACutoffTime = "1600" Then Response.Write(" selected ")%>>4:00 PM</option>
										<option value="1615"<%If RACutoffTime = "1615" Then Response.Write(" selected ")%>>4:15 PM</option>
										<option value="1630"<%If RACutoffTime = "1630" Then Response.Write(" selected ")%>>4:30 PM</option>
										<option value="1645"<%If RACutoffTime = "1645" Then Response.Write(" selected ")%>>4:45 PM</option>
										<option value="1700"<%If RACutoffTime = "1700" Then Response.Write(" selected ")%>>5:00 PM</option>
										<option value="1715"<%If RACutoffTime = "1715" Then Response.Write(" selected ")%>>5:15 PM</option>
										<option value="1730"<%If RACutoffTime = "1730" Then Response.Write(" selected ")%>>5:30 PM</option>
										<option value="1745"<%If RACutoffTime = "1745" Then Response.Write(" selected ")%>>5:45 PM</option>
										<option value="1800"<%If RACutoffTime = "1800" Then Response.Write(" selected ")%>>6:00 PM</option>
										<option value="1815"<%If RACutoffTime = "1815" Then Response.Write(" selected ")%>>6:15 PM</option>
										<option value="1830"<%If RACutoffTime = "1830" Then Response.Write(" selected ")%>>6:30 PM</option>
										<option value="1845"<%If RACutoffTime = "1845" Then Response.Write(" selected ")%>>6:45 PM</option>
										<option value="1900"<%If RACutoffTime = "1900" Then Response.Write(" selected ")%>>7:00 PM</option>
										<option value="1915"<%If RACutoffTime = "1915" Then Response.Write(" selected ")%>>7:15 PM</option>
										<option value="1930"<%If RACutoffTime = "1930" Then Response.Write(" selected ")%>>7:30 PM</option>
										<option value="1945"<%If RACutoffTime = "1945" Then Response.Write(" selected ")%>>7:45 PM</option>
										<option value="2000"<%If RACutoffTime = "2000" Then Response.Write(" selected ")%>>8:00 PM</option>
										<option value="2015"<%If RACutoffTime = "2015" Then Response.Write(" selected ")%>>8:15 PM</option>
										<option value="2030"<%If RACutoffTime = "2030" Then Response.Write(" selected ")%>>8:30 PM</option>
										<option value="2045"<%If RACutoffTime = "2045" Then Response.Write(" selected ")%>>8:45 PM</option>
										<option value="2100"<%If RACutoffTime = "2100" Then Response.Write(" selected ")%>>9:00 PM</option>
										<option value="2115"<%If RACutoffTime = "2115" Then Response.Write(" selected ")%>>9:15 PM</option>
										<option value="2130"<%If RACutoffTime = "2130" Then Response.Write(" selected ")%>>9:30 PM</option>
										<option value="2145"<%If RACutoffTime = "2145" Then Response.Write(" selected ")%>>9:45 PM</option>
										<option value="2200"<%If RACutoffTime = "2200" Then Response.Write(" selected ")%>>10:00 PM</option>
										<option value="2215"<%If RACutoffTime = "2215" Then Response.Write(" selected ")%>>10:15 PM</option>
										<option value="2230"<%If RACutoffTime = "2230" Then Response.Write(" selected ")%>>10:30 PM</option>
										<option value="2245"<%If RACutoffTime = "2245" Then Response.Write(" selected ")%>>10:45 PM</option>
										<option value="2300"<%If RACutoffTime = "2300" Then Response.Write(" selected ")%>>11:00 PM</option>
										<option value="2315"<%If RACutoffTime = "2315" Then Response.Write(" selected ")%>>11:15 PM</option>
										<option value="2330"<%If RACutoffTime = "2330" Then Response.Write(" selected ")%>>11:30 PM</option>
										<option value="2345"<%If RACutoffTime = "2345" Then Response.Write(" selected ")%>>11:45 PM</option>	
									</select>
								</div>
								<!-- eof order cutoff time -->
								
								<div class="col-lg-6 pull-left">
									<br><p><em>Posts received after the cutoff time will have their delivery date increased by one business day.</em></p>
								</div>
								<!-- eof excerpt -->
							</div>
							<!-- eof cutoff time -->
	
		
											
					</div>
				</div>
			</div>

		</div>
		
		
	    <div class="row">
	    
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">CREDIT MEMO RePost Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
					
		              		<!-- line !-->
		                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
		                	
		               			<div class="col-lg-6">
		               			<%
									If CMAPIRepostONOFF = 0 Then
										Response.Write("<input type='checkbox' class='check' id='chkCMAPIRepostONOFF' name='chkCMAPIRepostONOFF'")
									Else
										Response.Write("<input type='checkbox' class='check' id='chkCMAPIRepostONOFF' name='chkCMAPIRepostONOFF' checked")
									End If
									Response.Write("> Check to turn on")
								%>
		               			</div>
		               			
								<div class="col-lg-6">
									Mode:<br>
									<select class="form-control pull-left" name="selCMAPIRepostMode">
										<option value="TEST" <% If CMAPIRepostMode = "TEST" Then Response.Write("selected") %>>TEST</option>
										<option value="LIVE" <% If CMAPIRepostMode = "LIVE" Then Response.Write("selected") %>>LIVE</option>
									</select>	
								</div>
		               			
							</div>
							
				         	<!-- POST Received invoices To URL-->
							<div class="row schedule-info" style="margin-top:0px;margin-bottom:10px;">
								<div class="col-lg-10"><label  for="txtCMsAPIRepostURL" class="post-labels">Re-post received credit memos to URL</label></div>
							</div>
								
							<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
								<div class="col-lg-12"><input type="text" class="form-control" name="txtCMAPIRepostURL" id="txtCMAPIRepostURL" value="<%= CMAPIRepostURL %>"></div>
							</div>
							
	
	
							<!-- order cutoff time -->
							<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
								<!-- select time -->
								<div class="col-lg-6">
									Cutoff Time:
									<select class="form-control" name="selCMCutoffTime" id="selCMCutoffTime">
										<option value="0000"<%If CMCutoffTime = "0000" Then Response.Write(" selected ")%>>-Midnight-</option>
										<option value="0015"<%If CMCutoffTime = "0015" Then Response.Write(" selected ")%>>12:15 AM</option>
										<option value="0030"<%If CMCutoffTime = "0030" Then Response.Write(" selected ")%>>12:30 AM</option>
										<option value="0045"<%If CMCutoffTime = "0045" Then Response.Write(" selected ")%>>12:45 AM</option>
										<option value="100"<%If CMCutoffTime = "100" Then Response.Write(" selected ")%>>1:00 AM</option>
										<option value="115"<%If CMCutoffTime = "115" Then Response.Write(" selected ")%>>1:15 AM</option>
										<option value="130"<%If CMCutoffTime = "130" Then Response.Write(" selected ")%>>1:30 AM</option>
										<option value="145"<%If CMCutoffTime = "145" Then Response.Write(" selected ")%>>1:45 AM</option>
										<option value="200"<%If CMCutoffTime = "200" Then Response.Write(" selected ")%>>2:00 AM</option>
										<option value="215"<%If CMCutoffTime = "215" Then Response.Write(" selected ")%>>2:15 AM</option>
										<option value="230"<%If CMCutoffTime = "230" Then Response.Write(" selected ")%>>2:30 AM</option>
										<option value="245"<%If CMCutoffTime = "245" Then Response.Write(" selected ")%>>2:45 AM</option>
										<option value="300"<%If CMCutoffTime = "300" Then Response.Write(" selected ")%>>3:00 AM</option>
										<option value="315"<%If CMCutoffTime = "315" Then Response.Write(" selected ")%>>3:15 AM</option>
										<option value="330"<%If CMCutoffTime = "330" Then Response.Write(" selected ")%>>3:30 AM</option>
										<option value="345"<%If CMCutoffTime = "345" Then Response.Write(" selected ")%>>3:45 AM</option>
										<option value="400"<%If CMCutoffTime = "400" Then Response.Write(" selected ")%>>4:00 AM</option>
										<option value="415"<%If CMCutoffTime = "415" Then Response.Write(" selected ")%>>4:15 AM</option>
										<option value="430"<%If CMCutoffTime = "430" Then Response.Write(" selected ")%>>4:30 AM</option>
										<option value="445"<%If CMCutoffTime = "445" Then Response.Write(" selected ")%>>4:45 AM</option>
										<option value="500"<%If CMCutoffTime = "500" Then Response.Write(" selected ")%>>5:00 AM</option>
										<option value="515"<%If CMCutoffTime = "515" Then Response.Write(" selected ")%>>5:15 AM</option>
										<option value="530"<%If CMCutoffTime = "530" Then Response.Write(" selected ")%>>5:30 AM</option>
										<option value="545"<%If CMCutoffTime = "545" Then Response.Write(" selected ")%>>5:45 AM</option>
										<option value="600"<%If CMCutoffTime = "600" Then Response.Write(" selected ")%>>6:00 AM</option>
										<option value="615"<%If CMCutoffTime = "615" Then Response.Write(" selected ")%>>6:15 AM</option>
										<option value="630"<%If CMCutoffTime = "630" Then Response.Write(" selected ")%>>6:30 AM</option>
										<option value="645"<%If CMCutoffTime = "645" Then Response.Write(" selected ")%>>6:45 AM</option>
										<option value="700"<%If CMCutoffTime = "700" Then Response.Write(" selected ")%>>7:00 AM</option>
										<option value="715"<%If CMCutoffTime = "715" Then Response.Write(" selected ")%>>7:15 AM</option>
										<option value="730"<%If CMCutoffTime = "730" Then Response.Write(" selected ")%>>7:30 AM</option>
										<option value="745"<%If CMCutoffTime = "745" Then Response.Write(" selected ")%>>7:45 AM</option>
										<option value="800"<%If CMCutoffTime = "800" Then Response.Write(" selected ")%>>8:00 AM</option>
										<option value="815"<%If CMCutoffTime = "815" Then Response.Write(" selected ")%>>8:15 AM</option>
										<option value="830"<%If CMCutoffTime = "830" Then Response.Write(" selected ")%>>8:30 AM</option>
										<option value="845"<%If CMCutoffTime = "845" Then Response.Write(" selected ")%>>8:45 AM</option>
										<option value="900"<%If CMCutoffTime = "900" Then Response.Write(" selected ")%>>9:00 AM</option>
										<option value="915"<%If CMCutoffTime = "915" Then Response.Write(" selected ")%>>9:15 AM</option>
										<option value="930"<%If CMCutoffTime = "930" Then Response.Write(" selected ")%>>9:30 AM</option>
										<option value="945"<%If CMCutoffTime = "945" Then Response.Write(" selected ")%>>9:45 AM</option>
										<option value="1000"<%If CMCutoffTime = "1000" Then Response.Write(" selected ")%>>10:00 AM</option>
										<option value="1015"<%If CMCutoffTime = "1015" Then Response.Write(" selected ")%>>10:15 AM</option>
										<option value="1030"<%If CMCutoffTime = "1030" Then Response.Write(" selected ")%>>10:30 AM</option>
										<option value="1045"<%If CMCutoffTime = "1045" Then Response.Write(" selected ")%>>10:45 AM</option>
										<option value="1100"<%If CMCutoffTime = "1100" Then Response.Write(" selected ")%>>11:00 AM</option>
										<option value="1115"<%If CMCutoffTime = "1115" Then Response.Write(" selected ")%>>11:15 AM</option>
										<option value="1130"<%If CMCutoffTime = "1130" Then Response.Write(" selected ")%>>11:30 AM</option>
										<option value="1145"<%If CMCutoffTime = "1145" Then Response.Write(" selected ")%>>11:45 AM</option>
										<option value="1200"<%If CMCutoffTime = "1200" Then Response.Write(" selected ")%>>-Noon-</option>
										<option value="1215"<%If CMCutoffTime = "1215" Then Response.Write(" selected ")%>>12:15 PM</option>
										<option value="1230"<%If CMCutoffTime = "1230" Then Response.Write(" selected ")%>>12:30 PM</option>
										<option value="1245"<%If CMCutoffTime = "1245" Then Response.Write(" selected ")%>>12:45 PM</option>
										<option value="1300"<%If CMCutoffTime = "1300" Then Response.Write(" selected ")%>>1:00 PM</option>
										<option value="1315"<%If CMCutoffTime = "1315" Then Response.Write(" selected ")%>>1:15 PM</option>
										<option value="1330"<%If CMCutoffTime = "1330" Then Response.Write(" selected ")%>>1:30 PM</option>
										<option value="1345"<%If CMCutoffTime = "1345" Then Response.Write(" selected ")%>>1:45 PM</option>
										<option value="1400"<%If CMCutoffTime = "1400" Then Response.Write(" selected ")%>>2:00 PM</option>
										<option value="1415"<%If CMCutoffTime = "1415" Then Response.Write(" selected ")%>>2:15 PM</option>
										<option value="1430"<%If CMCutoffTime = "1430" Then Response.Write(" selected ")%>>2:30 PM</option>
										<option value="1445"<%If CMCutoffTime = "1445" Then Response.Write(" selected ")%>>2:45 PM</option>
										<option value="1500"<%If CMCutoffTime = "1500" Then Response.Write(" selected ")%>>3:00 PM</option>
										<option value="1515"<%If CMCutoffTime = "1515" Then Response.Write(" selected ")%>>3:15 PM</option>
										<option value="1530"<%If CMCutoffTime = "1530" Then Response.Write(" selected ")%>>3:30 PM</option>
										<option value="1545"<%If CMCutoffTime = "1545" Then Response.Write(" selected ")%>>3:45 PM</option>
										<option value="1600"<%If CMCutoffTime = "1600" Then Response.Write(" selected ")%>>4:00 PM</option>
										<option value="1615"<%If CMCutoffTime = "1615" Then Response.Write(" selected ")%>>4:15 PM</option>
										<option value="1630"<%If CMCutoffTime = "1630" Then Response.Write(" selected ")%>>4:30 PM</option>
										<option value="1645"<%If CMCutoffTime = "1645" Then Response.Write(" selected ")%>>4:45 PM</option>
										<option value="1700"<%If CMCutoffTime = "1700" Then Response.Write(" selected ")%>>5:00 PM</option>
										<option value="1715"<%If CMCutoffTime = "1715" Then Response.Write(" selected ")%>>5:15 PM</option>
										<option value="1730"<%If CMCutoffTime = "1730" Then Response.Write(" selected ")%>>5:30 PM</option>
										<option value="1745"<%If CMCutoffTime = "1745" Then Response.Write(" selected ")%>>5:45 PM</option>
										<option value="1800"<%If CMCutoffTime = "1800" Then Response.Write(" selected ")%>>6:00 PM</option>
										<option value="1815"<%If CMCutoffTime = "1815" Then Response.Write(" selected ")%>>6:15 PM</option>
										<option value="1830"<%If CMCutoffTime = "1830" Then Response.Write(" selected ")%>>6:30 PM</option>
										<option value="1845"<%If CMCutoffTime = "1845" Then Response.Write(" selected ")%>>6:45 PM</option>
										<option value="1900"<%If CMCutoffTime = "1900" Then Response.Write(" selected ")%>>7:00 PM</option>
										<option value="1915"<%If CMCutoffTime = "1915" Then Response.Write(" selected ")%>>7:15 PM</option>
										<option value="1930"<%If CMCutoffTime = "1930" Then Response.Write(" selected ")%>>7:30 PM</option>
										<option value="1945"<%If CMCutoffTime = "1945" Then Response.Write(" selected ")%>>7:45 PM</option>
										<option value="2000"<%If CMCutoffTime = "2000" Then Response.Write(" selected ")%>>8:00 PM</option>
										<option value="2015"<%If CMCutoffTime = "2015" Then Response.Write(" selected ")%>>8:15 PM</option>
										<option value="2030"<%If CMCutoffTime = "2030" Then Response.Write(" selected ")%>>8:30 PM</option>
										<option value="2045"<%If CMCutoffTime = "2045" Then Response.Write(" selected ")%>>8:45 PM</option>
										<option value="2100"<%If CMCutoffTime = "2100" Then Response.Write(" selected ")%>>9:00 PM</option>
										<option value="2115"<%If CMCutoffTime = "2115" Then Response.Write(" selected ")%>>9:15 PM</option>
										<option value="2130"<%If CMCutoffTime = "2130" Then Response.Write(" selected ")%>>9:30 PM</option>
										<option value="2145"<%If CMCutoffTime = "2145" Then Response.Write(" selected ")%>>9:45 PM</option>
										<option value="2200"<%If CMCutoffTime = "2200" Then Response.Write(" selected ")%>>10:00 PM</option>
										<option value="2215"<%If CMCutoffTime = "2215" Then Response.Write(" selected ")%>>10:15 PM</option>
										<option value="2230"<%If CMCutoffTime = "2230" Then Response.Write(" selected ")%>>10:30 PM</option>
										<option value="2245"<%If CMCutoffTime = "2245" Then Response.Write(" selected ")%>>10:45 PM</option>
										<option value="2300"<%If CMCutoffTime = "2300" Then Response.Write(" selected ")%>>11:00 PM</option>
										<option value="2315"<%If CMCutoffTime = "2315" Then Response.Write(" selected ")%>>11:15 PM</option>
										<option value="2330"<%If CMCutoffTime = "2330" Then Response.Write(" selected ")%>>11:30 PM</option>
										<option value="2345"<%If CMCutoffTime = "2345" Then Response.Write(" selected ")%>>11:45 PM</option>	
									</select>
								</div>
								<!-- eof order cutoff time -->
								
								<div class="col-lg-6 pull-left">
									<br><p><em>Posts received after the cutoff time will have their delivery date increased by one business day.</em></p>
								</div>
								<!-- eof excerpt -->
							</div>
							<!-- eof cutoff time -->
	
		
				
					</div>
				</div>
			</div>
			
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">SUMMARY INVOICE RePost Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
					
		              		<!-- line !-->
		                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
		                	
		               			<div class="col-lg-6">
								<%
									If SumInvAPIRepostONOFF = 0 Then
										Response.Write("<input type='checkbox' class='check' id='chkSumInvAPIRepostONOFF' name='chkSumInvAPIRepostONOFF'")
									Else
										Response.Write("<input type='checkbox' class='check' id='chkSumInvAPIRepostONOFF' name='chkSumInvAPIRepostONOFF' checked")
									End If
									Response.Write("> Check to turn on")
								%>
		               			</div>
		               			
								<div class="col-lg-6">
									Mode:<br>
									<select class="form-control pull-left" name="selSumInvAPIRepostMode">
										<option value="TEST" <% If SumInvAPIRepostMode = "TEST" Then Response.Write("selected") %>>TEST</option>
										<option value="LIVE" <% If SumInvAPIRepostMode = "LIVE" Then Response.Write("selected") %>>LIVE</option>
									</select>	
								</div>
		               			
							</div>
					
							<!-- POST Received invoices To URL-->
							<div class="row schedule-info" style="margin-top:0px;margin-bottom:10px;">
								<div class="col-lg-10"><label  for="txtSumInvsAPIRepostURL" class="post-labels">Re-post received summary invoices to URL</label></div>
							</div>
								
							<div class="row schedule-info" style="margin-top:0px;margin-bottom:2px;">
								<div class="col-lg-12"><input type="text" class="form-control" name="txtSumInvAPIRepostURL" id="txtSumInvAPIRepostURL" value="<%= SumInvAPIRepostURL %>"></div>
							</div>
							
											
					</div>
				</div>
			</div>
	
			<div class="col-md-4">
				&nbsp;
			</div>
			
		</div>
		
	
	
	
		<div class="row">
			<h3><i class="fad fa-file-pdf"></i>&nbsp;Order API Report Settings</h3>
		</div>
	
		<div class="row">
		
			<div class="col-md-6">
				<% If APIDailyActivityReportOnOff = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title">Daily API Activity Summary By Partner Report (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title">Daily API Activity Summary By Partner Report (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
					<div class="panel-body">

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON 
					      		<%
					      		If APIDailyActivityReportOnOff = 0 Then
									Response.Write("<input type='checkbox' id='chkAPIDailyActivityReportOnOff' name='chkAPIDailyActivityReportOnOff'")
								Else
									Response.Write("<input type='checkbox' id='chkAPIDailyActivityReportOnOff' name='chkAPIDailyActivityReportOnOff' checked")
								End If
								Response.Write(">")
								%>
				            </div>
				            <!-- eof line -->
				         </div>  
				         					
					
					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
								
								<div class="text-element circles-list">
									<ol>
										<li>
											<p>Set the report send schedule:</p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalDailyAPIActivityByPartnerReportScheduler" data-tooltip="true" data-title="Daily API Activity Summary By Partner Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> Daily API Activity Summary By Partner Report Scheduler</button>
										</li>
										<li>								
											<p>Specify the subject line to be used for the email:</p>
											<input type="text"class="form-control" style="width:100%;" name="txtAPIDailyActivityReportEmailSubject" id="txtAPIDailyActivityReportEmailSubject" value="<%= APIDailyActivityReportEmailSubject %>">
										</li>
										<li>
											<p>Select users <i class="fad fa-user-friends"></i> to send the report to:</p>
											<input type="hidden" name="lstSelectedAPIDailyActivityReportUserNos" id="lstSelectedAPIDailyActivityReportUserNos" value="<%= APIDailyActivityReportUserNos %>">
											<select id="lstExistingAPIDailyActivityReportUserNos" multiple="multiple" name="lstExistingAPIDailyActivityReportUserNos">
												<%	'Get list of all users not currently archived or disabled
													
												Set cnnUserList = Server.CreateObject("ADODB.Connection")
												cnnUserList.open Session("ClientCnnString")
								
												SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 and userEnabled <> 0 ORDER BY userFirstName,userLastName"
												
												Set rsUserList = Server.CreateObject("ADODB.Recordset")
												rsUserList.CursorLocation = 3 
												Set rsUserList = cnnUserList.Execute(SQLUserList)
												
												If Not rsUserList.EOF Then
													Do While Not rsUserList.EOF
													
														FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName") & " (" & rsUserList("userDisplayName") & ")"
														Response.Write("<option value='" & rsUserList("UserNo") & "'>" & FullName & "</option>")
												
														rsUserList.MoveNext
													Loop
												End If
									
												Set rsUserList = Nothing
												cnnUserList.Close
												Set cnnUserList = Nothing
													
												%>
											</select>				
										</li>
										<li>
											<p>Select additional email addresses to send the report to:</p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalAPIDailyActivityReportAdditionalEmails" data-tooltip="true" data-title="Additional emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails</button>						
				             				<% If APIDailyActivityReportAdditionalEmails <> "" Then %>
				             					<p style="margin-top:20px;"><strong>Current Additional Emails:</strong> <%= APIDailyActivityReportAdditionalEmails %></p>
				             				<% End If %>
										</li>
										<li>
											<p>Include order offset by:</p>
											<select class="form-control pull-left" name="selOrderAPIOffsetDays" style="width:25%; margin-bottom:20px;">
												<option value="-2" <% If OrderAPIOffsetDays = -2 Then Response.Write("selected") %>>-2 DAYS</option>
												<option value="-1" <% If OrderAPIOffsetDays = -1 Then Response.Write("selected") %>>-1 DAYS</option>
												<option value="0" <% If OrderAPIOffsetDays = 0 Then Response.Write("selected") %>>0 DAYS</option>
												<option value="1" <% If OrderAPIOffsetDays = 1 Then Response.Write("selected") %>>1 DAY</option>
												<option value="2" <% If OrderAPIOffsetDays = 2 Then Response.Write("selected") %>>2 DAYS</option>
												<option value="3" <% If OrderAPIOffsetDays = 3 Then Response.Write("selected") %>>3 DAYS</option>
												<option value="4" <% If OrderAPIOffsetDays = 4 Then Response.Write("selected") %>>4 DAYS</option>
												<option value="5" <% If OrderAPIOffsetDays = 5 Then Response.Write("selected") %>>5 DAYS</option>
											</select>
										</li>										
										<li>
											<p>Include invoices offset by:</p>
											<select class="form-control pull-left" name="selInvoiceAPIOffsetDays" style="width:25%; margin-bottom:20px;">
												<option value="0" <% If InvoiceAPIOffsetDays = 0 Then Response.Write("selected") %>>0 DAYS</option>
												<option value="1" <% If InvoiceAPIOffsetDays = 1 Then Response.Write("selected") %>>1 DAY</option>
												<option value="2" <% If InvoiceAPIOffsetDays = 2 Then Response.Write("selected") %>>2 DAYS</option>
												<option value="3" <% If InvoiceAPIOffsetDays = 3 Then Response.Write("selected") %>>3 DAYS</option>
												<option value="4" <% If InvoiceAPIOffsetDays = 4 Then Response.Write("selected") %>>4 DAYS</option>
												<option value="5" <% If InvoiceAPIOffsetDays = 5 Then Response.Write("selected") %>>5 DAYS</option>
											</select>
										</li>
										<li>
											<p>Include RA's offset by:<p>
											<select class="form-control pull-left" name="selRAAPIOffsetDays" style="width:25%; margin-bottom:20px;">
												<option value="0" <% If RAAPIOffsetDays = 0 Then Response.Write("selected") %>>0 DAYS</option>
												<option value="1" <% If RAAPIOffsetDays = 1 Then Response.Write("selected") %>>1 DAY</option>
												<option value="2" <% If RAAPIOffsetDays = 2 Then Response.Write("selected") %>>2 DAYS</option>
												<option value="3" <% If RAAPIOffsetDays = 3 Then Response.Write("selected") %>>3 DAYS</option>
												<option value="4" <% If RAAPIOffsetDays = 4 Then Response.Write("selected") %>>4 DAYS</option>
												<option value="5" <% If RAAPIOffsetDays = 5 Then Response.Write("selected") %>>5 DAYS</option>
											</select>
										</li>
										<li>
											<p>Include CM's offset by:</p>
											<select class="form-control pull-left" name="selCMAPIOffsetDays" style="width:25%; margin-bottom:20px;">
												<option value="0" <% If CMAPIOffsetDays = 0 Then Response.Write("selected") %>>0 DAYS</option>
												<option value="1" <% If CMAPIOffsetDays = 1 Then Response.Write("selected") %>>1 DAY</option>
												<option value="2" <% If CMAPIOffsetDays = 2 Then Response.Write("selected") %>>2 DAYS</option>
												<option value="3" <% If CMAPIOffsetDays = 3 Then Response.Write("selected") %>>3 DAYS</option>
												<option value="4" <% If CMAPIOffsetDays = 4 Then Response.Write("selected") %>>4 DAYS</option>
												<option value="5" <% If CMAPIOffsetDays = 5 Then Response.Write("selected") %>>5 DAYS</option>
											</select>
										</li>
										<li>
											<p>Include sum. inv. offset by:</p>
											<select class="form-control pull-left" name="selSumInvAPIOffsetDays" style="width:25%; margin-bottom:20px;">
												<option value="0" <% If SumInvAPIOffsetDays = 0 Then Response.Write("selected") %>>0 DAYS</option>
												<option value="1" <% If SumInvAPIOffsetDays = 1 Then Response.Write("selected") %>>1 DAY</option>
												<option value="2" <% If SumInvAPIOffsetDays = 2 Then Response.Write("selected") %>>2 DAYS</option>
												<option value="3" <% If SumInvAPIOffsetDays = 3 Then Response.Write("selected") %>>3 DAYS</option>
												<option value="4" <% If SumInvAPIOffsetDays = 4 Then Response.Write("selected") %>>4 DAYS</option>
												<option value="5" <% If SumInvAPIOffsetDays = 5 Then Response.Write("selected") %>>5 DAYS</option>
											</select>
										</li>
										
									</ol>
								</div>
					
							</div>
						</div>
					
					
					</div>
				</div>
				
				<div class="col-md-6">
					&nbsp;
				</div>
				
			</div>
			
		</div>	

         
		<!-- cancel / save !-->
		<div class="row pull-right">
			<div class="col-lg-12">
				<a href="<%= BaseURL %>admin/global/tiles/api/main.asp"><button type="button" class="btn btn-default btn-lg btn-huge"><i class="far fa-times-circle"></i> Cancel</button></a> 
				<button type="submit" class="btn btn-primary btn-lg btn-huge" onclick="showSavingChangesDiv()"><i class="far fa-save"></i> Save Changes</button>
			</div>
		</div>
	
	<% End If %>
	</div><!-- row -->
</div><!-- container -->

</form>


<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR REPORT SCHEDULERS START HERE !-->
<!-- **************************************************************************************************************************** -->

<!-- pencil Modal -->
<div class="modal fade" id="modalDailyAPIActivityByPartnerReportScheduler" tabindex="-1" role="dialog" aria-labelledby="modalDailyAPIActivityByPartnerReportSchedulerLabel">
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="titleDailyAPIActivityByPartnerReportSchedulerLabel">Daily API Activity Summary By Partner Report Generation Scheduler</h4>
		    </div>

			<form name="frmEditDailyAPIActivityByPartnerReportSchedulerModal" id="frmEditDailyAPIActivityByPartnerReportSchedulerModal" action="order-api-daily-activity-by-partner-report-scheduler-submit.asp" method="POST">

				<div class="modal-body">
				    
					<div id="modalDailyAPIActivityByPartnerReportSchedulerContent">
						<!-- Content for the modal will be generated and written here -->
						<!-- Content generated by Sub GetContentForDailyAPIActivityByPartnerReportScheduler() in InSightFuncs_AjaxForAdminTimepickerModals.asp -->
					</div>
						
				</div>
				<!-- eof modal body !-->
				
				 <div class="clearfix"></div>
			      
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
					<button type="submit" id="btnDailyAPIActivityByPartnerReportScheduleSave" class="btn btn-primary">Save Schedule Changes</button>
				</div>
				
			</form>

		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->



<div class="modal fade" id="modalAPIDailyActivityReportAdditionalEmails" tabindex="-1" role="dialog" aria-labelledby="modalAPIDailyActivityReportAdditionalEmailsLabel">
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="H6">Send report to the following additional email addresses</h4>
		    </div>

			<form name="frmEditAPIDailyActivityReportAdditionalEmails" id="frmEditAPIDailyActivityReportAdditionalEmails" action="users-list-update-api.asp" method="POST">
                <input type="hidden" name="userListName" value="APIDailyActivityReportAdditionalEmails" />
				<div class="modal-body">
				    
					<div id="Div6">
						<textarea class="form-control email-alert-line" rows="5" id="txtAPIDailyActivityReportAdditionalEmails" name="txtAPIDailyActivityReportAdditionalEmails"><%= APIDailyActivityReportAdditionalEmails %></textarea>
						<strong>Separate multiple email addresses with a semicolon</strong>
					</div>
						
				</div>
				<!-- eof modal body !-->
				
				 <div class="clearfix"></div>
			      
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
					<button type="submit" id="Button5" class="btn btn-primary">Save Changes</button>
				</div>
				
			</form>

		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>


<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR REPORT SCHEDULERS END HERE !-->
<!-- **************************************************************************************************************************** -->



<!--#include file="../../../../inc/footer-main.asp"-->
